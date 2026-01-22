# -*- coding: utf-8 -*-
"""
Subcon – 复制两份最新文件到本周命名并自动处理：
1) 复制 China / Non China 最新文件到本周命名（按 Wxx'yy）
2) 每个新文件：
   - 刷新连接 "Query - VW_VendorPerformance"（优先走 QueryTable，同步刷新）
   - 在 Sheet1 的 BH 列批量填入“合法的 Reason Code”（从 BH2 的数据验证读取；找不到则用 "Reason Code"）
"""

import os
import re
import time
import shutil
import datetime as dt
from glob import glob
import win32com.client as win32

# ================== 配置 ==================
BASE_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\Supplier SUBCON Performance\SUBCON"
PATTERN_CHINA    = "China SUBCON - KPIs Review (PO GR) - W*'*(First AB).xlsx"
PATTERN_NONCHINA = "Non China SUBCON - KPIs Review (PO GR) - W*'*(First AB).xlsx"

SHEET_NAME = "Sheet1"
CONN_NAME  = "Query - VW_VendorPerformance"
HEADER_BH  = None  # 若 BH 在表(ListObject)里且你知道表头，可填字符串；不知道就留 None

WEEK_OFFSET = -1  # 业务周 = ISO 周 - 1
# ==========================================

# Excel 常量
XL_UP = -4162

# --------- 撇号兼容 & 正则 ---------
APOS_VARIANTS = "'’`′"
APOS_CLASS = "[" + re.escape(APOS_VARIANTS) + "]"

# China / Non-China 文件名正则（兼容四种撇号）
RE_CHINA = re.compile(
    rf"^China SUBCON - KPIs Review \(PO GR\) - W(\d{{1,2}}){APOS_CLASS}(\d{{2}})\(First AB\)\.xlsx$",
    re.IGNORECASE
)
RE_NONCHINA = re.compile(
    rf"^Non China SUBCON - KPIs Review \(PO GR\) - W(\d{{1,2}}){APOS_CLASS}(\d{{2}})\(First AB\)\.xlsx$",
    re.IGNORECASE
)

def _glob_variants(base_dir: str, pattern: str) -> list[str]:
    """尝试四种撇号变体的 glob 通配符，汇总结果（兜底用）"""
    res = []
    for a in APOS_VARIANTS:
        pat = pattern.replace("'*'", f"{a}*{a}")
        res += [f for f in glob(os.path.join(base_dir, pat)) if os.path.isfile(f)]
    # 去重
    return list({os.path.abspath(p) for p in res})

# ---------------- 周标 & 文件复制 ----------------
def compute_week_token(today: dt.date | None = None) -> str:
    """业务周：W = ISO 周 - 1（<=0 跨到上一年末）"""
    d = today or dt.date.today()
    y, w, _ = d.isocalendar()
    w += WEEK_OFFSET
    if w <= 0:
        last_dec_28 = dt.date(y - 1, 12, 28)
        _, w_last, _ = last_dec_28.isocalendar()
        w = w_last + w
        y -= 1
    return f"W{w:02d}'{str(y)[-2:]}"

def _find_latest_by_regex(base_dir: str, regex) -> str | None:
    """用 scandir + 正则匹配（按 mtime 降序取最新）"""
    items: list[tuple[str, float]] = []
    with os.scandir(base_dir) as it:
        for de in it:
            if not de.is_file():
                continue
            if regex.match(de.name):
                try:
                    items.append((de.path, de.stat().st_mtime))
                except OSError:
                    pass
    if not items:
        return None
    items.sort(key=lambda x: x[1], reverse=True)
    return items[0][0]

def find_latest(base_dir: str, pattern: str, which: str) -> str:
    """
    寻找最新文件：优先正则（兼容撇号），找不到再回退到 glob 变体。
    which: "china" 或 "nonchina" 用于选择对应正则
    """
    regex = RE_CHINA if which.lower() == "china" else RE_NONCHINA
    p = _find_latest_by_regex(base_dir, regex)
    if p:
        return p

    # 回退：glob 四种撇号变体
    cands = _glob_variants(base_dir, pattern)
    if not cands:
        raise FileNotFoundError(f"未在 {base_dir} 找到匹配文件：{pattern} / No matching file in {base_dir}: {pattern}")
    cands.sort(key=os.path.getmtime, reverse=True)
    return cands[0]

def make_this_week_name(from_name: str, wyy: str) -> str:
    """把文件名中的 Wxx'yy 替换为本周；若无周标则追加在 (PO GR) - 后面前"""
    base, ext = os.path.splitext(from_name)
    # 先尝试替换（兼容四种撇号）
    new_base = re.sub(rf"W(\d{{1,2}}){APOS_CLASS}(\d{{2}})", wyy, base, flags=re.IGNORECASE)
    if new_base != base:
        return new_base + ext
    # 若文件名里原本没有周标，则在固定位置插入
    marker = " - W"
    if "(PO GR) - " in base:
        ins_at = base.find("(PO GR) - ") + len("(PO GR) - ")
        new_base = base[:ins_at] + wyy + base[ins_at:]
    else:
        new_base = f"{base} {wyy}"
    return new_base + ext

def copy_to_this_week(base_dir: str, latest_path: str, wyy: str) -> str:
    dst = os.path.join(base_dir, make_this_week_name(os.path.basename(latest_path), wyy))
    if os.path.abspath(dst) == os.path.abspath(latest_path):
        print("⚠ 已经是本周命名，无需复制：", os.path.basename(dst),
              "/ Already this week's name; no copy needed:", os.path.basename(dst))
        return latest_path
    if os.path.exists(dst):
        print("ℹ 本周文件已存在：", os.path.basename(dst),
              "/ This week's file already exists:", os.path.basename(dst))
        return dst
    shutil.copy2(latest_path, dst)
    print("✔ 已复制为本周文件：", os.path.basename(dst),
          "/ Copied as this week's file:", os.path.basename(dst))
    return dst

# ---------------- Excel 操作工具 ----------------
def open_excel_silent():
    ex = win32.DispatchEx("Excel.Application")
    ex.Visible = False
    ex.DisplayAlerts = False
    ex.AskToUpdateLinks = False
    ex.ScreenUpdating = False
    ex.AutomationSecurity = 3  # 禁用宏
    return ex

def open_wb_with_retry(path, tries=6, delay=1.0):
    last_err = None

    # 1) 本地先校验一下路径是否真存在（避免无谓重试）
    if not os.path.exists(path):
        raise FileNotFoundError(f"路径不存在：{path} / Path does not exist: {path}")

    for i in range(1, tries + 1):
        try:
            ex = open_excel_silent()
            # ⚠ 不要对路径做任何替换，Excel COM 能正确处理文件名里的单引号
            wb = ex.Workbooks.Open(
                Filename=path,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True
            )
            return ex, wb
        except Exception as e:
            last_err = e
            print(f"⏳ Open retry {i}/{tries} failed: {e}")
            # 若误传了带双引号路径，自动纠正一次
            path = path.replace("''", "'")
            time.sleep(delay)

    raise RuntimeError(f"无法打开文件：{path}\n最后错误：{last_err} / "
                       f"Unable to open file: {path}\nLast error: {last_err}")

def first_table(ws):
    try:
        return ws.ListObjects(1) if ws.ListObjects.Count > 0 else None
    except Exception:
        return None

def table_col_by_header(lo, header_text):
    for i in range(1, lo.ListColumns.Count + 1):
        if str(lo.ListColumns(i).Name).strip().lower() == str(header_text).strip().lower():
            return lo.ListColumns(i)
    return None

def col_to_index(ws, col):  # 'BH' -> 60
    if isinstance(col, int): return col
    return int(ws.Range(f"{col}1").Column)

def last_row_in_col(ws, col):
    col_idx = col_to_index(ws, col)
    return int(ws.Cells(ws.Rows.Count, col_idx).End(XL_UP).Row)

def get_validation_allowed_value(app, ws, addr, prefer_contains="reason code"):
    """从某单元格的数据验证列表里拿‘合法值’，优先包含 prefer_contains 的项。"""
    try:
        dv = ws.Range(addr).Validation
    except Exception:
        return None
    if getattr(dv, "Type", None) != 3:
        return None
    src = dv.Formula1
    if not src:
        return None

    values = []
    try:
        if src.startswith("="):
            res = app.Evaluate(src)
            try:
                vals = res.Value
            except Exception:
                vals = res
            if isinstance(vals, tuple):
                for row in vals:
                    if isinstance(row, tuple):
                        for v in row:
                            if v not in (None, ""):
                                values.append(str(v))
                    else:
                        if row not in (None, ""):
                            values.append(str(row))
        else:  # ="A,B,C"
            s = src[1:-1] if len(src) >= 2 and src[0] == '"' and src[-1] == '"' else src
            values = [x.strip() for x in s.split(",") if x.strip()]
    except Exception:
        pass

    if not values:
        return None
    pref = prefer_contains.lower().strip()
    for v in values:
        if pref in v.lower().strip():
            return v
    return values[0]

def refresh_target_connection_or_qt(app, wb, ws, conn_name) -> bool:
    """
    优先使用 ListObject.QueryTable 同步刷新（qt.BackgroundQuery=False）；
    找不到再回退 wb.Connections(conn_name).Refresh + CalculateUntilAsyncQueriesDone()
    """
    try:
        for lo in ws.ListObjects:
            qt = getattr(lo, "QueryTable", None)
            if qt is not None:
                wbc = getattr(qt, "WorkbookConnection", None)
                if wbc and getattr(wbc, "Name", "") == conn_name:
                    try:
                        qt.BackgroundQuery = False
                    except Exception:
                        pass
                    qt.Refresh(False)   # 同步刷新
                    return True
    except Exception:
        pass
    try:
        wb.Connections(conn_name).Refresh()
        app.CalculateUntilAsyncQueriesDone()
        return True
    except Exception as e:
        print(f"⚠️ 无法刷新连接 {conn_name}: {e} / Failed to refresh connection {conn_name}: {e}")
        return False

# ===== 新增：展开并移除所有筛选 =====
def expand_and_clear_filters(ws):
    """展开所有分组并清除工作表与表格的筛选"""
    # 1) 展开大纲分组（行/列）
    try:
        ws.Outline.ShowLevels(RowLevels=8, ColumnLevels=8)
    except Exception:
        try:
            ws.Outline.ShowLevels(8, 8)
        except Exception:
            pass

    # 2) 工作表级筛选清除
    try:
        if getattr(ws, "FilterMode", False):
            ws.ShowAllData()
    except Exception:
        pass
    try:
        if getattr(ws, "AutoFilterMode", False):
            ws.AutoFilterMode = False
    except Exception:
        pass

    # 3) 表格(ListObject)上的筛选清除
    try:
        if ws.ListObjects.Count > 0:
            for i in range(1, ws.ListObjects.Count + 1):
                lo = ws.ListObjects(i)
                try:
                    af = lo.AutoFilter
                    if getattr(af, "FilterMode", False):
                        af.ShowAllData()
                except Exception:
                    try:
                        lo.Range.AutoFilter(Field=1)  # 触发一次无条件AutoFilter
                        lo.AutoFilter.ShowAllData()
                    except Exception:
                        pass
    except Exception:
        pass
# ===== 新增结束 =====

# ---------------- 单文件处理 ----------------
def process_file(path: str):
    print(f"\n=== 处理文件 === {os.path.basename(path)} / Processing file: {os.path.basename(path)}")

    # Phase 1: 刷新
    print("🔄 Refreshing ...")
    app1, wb1 = open_wb_with_retry(path, tries=6, delay=1.0)

    # --- 新增：打开后先“展开并移除所有filter”（对所有工作表，确保刷新不受影响） ---
    try:
        for _ws in wb1.Worksheets:
            expand_and_clear_filters(_ws)
    except Exception:
        pass
    # --- 新增结束 ---

    ws1 = wb1.Sheets(SHEET_NAME)
    _ = refresh_target_connection_or_qt(app1, wb1, ws1, CONN_NAME)
    wb1.Save(); wb1.Close(SaveChanges=True); app1.Quit()
    print("✅ Refresh done.")
    time.sleep(0.3)

    # Phase 2: 填满 BH
    print("✏️ Fill BH with 'Reason Code' ...")
    app2, wb2 = open_wb_with_retry(path, tries=6, delay=1.0)
    ws2 = wb2.Sheets(SHEET_NAME)

    # --- 新增：再次确保目标工作表已“展开并移除所有filter” ---
    try:
        expand_and_clear_filters(ws2)
    except Exception:
        pass
    # --- 新增结束 ---

    app2.EnableEvents = False; app2.ScreenUpdating = False
    try:
        ws2.Unprotect()
    except Exception:
        pass

    allowed = get_validation_allowed_value(app2, ws2, "BH2", prefer_contains="reason code") or "Reason Code"
    did_bh = False
    lo = first_table(ws2)
    if lo is not None and HEADER_BH:
        lc = table_col_by_header(lo, HEADER_BH)
        if lc is not None and lc.DataBodyRange is not None:
            lc.DataBodyRange.Value = allowed
            did_bh = True

    if not did_bh:
        last_bh = last_row_in_col(ws2, "BH")
        if last_bh >= 2:
            ws2.Range(f"BH2:BH{last_bh}").Value = allowed
        # else：只有表头，无需写

    wb2.Save(); wb2.Close(SaveChanges=True)
    app2.EnableEvents = True; app2.ScreenUpdating = True
    app2.Quit()
    print("🎉 BH 填充完成 / BH fill completed")

# ---------------- 主流程 ----------------
def main():
    print("==== Subcon – 复制到本周并自动处理 ==== / Subcon – copy to this week and auto process ====")
    wyy = compute_week_token()
    print("本周标识:", wyy, "/ Week token:", wyy)

    latest_ch  = find_latest(BASE_DIR, PATTERN_CHINA, which="china")
    latest_nc  = find_latest(BASE_DIR, PATTERN_NONCHINA, which="nonchina")
    print("源(China):", os.path.basename(latest_ch), "/ Source (China):", os.path.basename(latest_ch))
    print("源(NonChina):", os.path.basename(latest_nc), "/ Source (NonChina):", os.path.basename(latest_nc))

    out_ch = copy_to_this_week(BASE_DIR, latest_ch, wyy)
    out_nc = copy_to_this_week(BASE_DIR, latest_nc, wyy)
    print("本周(China):", os.path.basename(out_ch), "/ This week (China):", os.path.basename(out_ch))
    print("本周(NonChina):", os.path.basename(out_nc), "/ This week (NonChina):", os.path.basename(out_nc))

    # 依次处理两个新文件
    process_file(out_ch)
    process_file(out_nc)

    print("\n✅ 全部完成。 / All done.")

if __name__ == "__main__":
    main()
