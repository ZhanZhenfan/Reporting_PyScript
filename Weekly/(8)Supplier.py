# -*- coding: utf-8 -*-
"""
Supplier – Weekly automation (完整版、简洁日志)
流程：
1) 寻找“上周”文件（优先按周标精确匹配，找不到则按修改时间最新）→ 复制为“本周”命名
2) 刷新指定连接
3) 在 'Details' 中批量写 BN='Reason Code'（按验证列表拿合法值）、清空 BT
4) 以 BN 列最后一个 'Reason Code' 的行为末行：表内安全 Resize / 非表删除尾行
5) 硬缩 UsedRange（删末行以下/末列以右），保存关闭，重开一次触发生效
"""

import os
import re
import shutil
import time
import datetime as dt
from glob import glob
import win32com.client as win32

# ================== CONFIG ==================
BASE_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\Supplier SUBCON Performance\Supplier"
# 你原来的 GLOB_PATTERN 仅用于极端兜底；主逻辑改为正则 + scandir
GLOB_PATTERN = "Supplier - KPIs Review (PO GR) W*'*.xlsx"

# 手动覆盖本周周号（两位或不带前导0均可），用于紧急场景，例如：set SUPPLIER_WEEK=41
ENV_WEEK_OVERRIDE = "SUPPLIER_WEEK"

SHEET_NAME = "Details"
CONNECTIONS = ["Query - VW_VendorPerformance"]
HEADER_BN = "Final KPI Reason Code"

# Excel 常量
XL_UP = -4162          # xlUp
XL_TOLEFT = -4159      # xlToLeft

# 允许的撇号字符集合：' / ’ / ` / ′
APOS_VARIANTS = "'’`′"
APOS_CLASS = "[" + re.escape(APOS_VARIANTS) + "]"

# 匹配规范文件名：Supplier - KPIs Review (PO GR) Wxx'yy.xlsx
WEEK_FILE_RE = re.compile(
    rf"^Supplier - KPIs Review \(PO GR\) W(\d{{1,2}}){APOS_CLASS}(\d{{2}})\.xlsx$",
    re.IGNORECASE,
)
# ======================================================


# ---------- 周标计算 ----------
def normalize_week_token(token: str) -> str:
    """把 W40’25 / W40`25 / W40′25 统一为 W40'25"""
    token = token.upper().strip()
    for a in APOS_VARIANTS:
        token = token.replace(a, "'")
    return token

def compute_week_tokens(today: dt.date | None = None) -> tuple[str, str]:
    """
    返回 (last_week_token, this_week_token)，按你的口径：W = ISO周号 - 1
    规则：
      - 默认用今天的 ISO 周；显示周号 = iso_week - 1（<=0 则为 W0）
      - 若设置了 SUPPLIER_WEEK=N，则 N 代表“本周的显示周号 WN”（不是 ISO 周）
        上周显示周号 = (N - 1)，若 <0 则回滚到上一年最后一周的显示周号
    """
    d = today or dt.date.today()
    manual = os.getenv(ENV_WEEK_OVERRIDE)

    def last_iso_week_of_year(y: int) -> int:
        return dt.date(y, 12, 28).isocalendar().week  # ISO 约定：当年最大周

    if manual and manual.isdigit():
        # 手动：N 表示“本周显示周号”
        tw_disp = int(manual)             # 本周显示周号
        y = d.year                         # 先按今年
        # 找到“今年的最大显示周号”= (ISO最大周 - 1)，可能为 51 或 52
        max_disp = last_iso_week_of_year(y) - 1
        if max_disp < 0:  # 理论不会发生，但防御
            max_disp = 0

        lw_disp = tw_disp - 1
        ly = y
        if lw_disp < 0:
            # 上周显示周号 < 0，则回滚到上一年的最后显示周
            ly = y - 1
            max_disp_prev = last_iso_week_of_year(ly) - 1
            lw_disp = max_disp_prev

        last_token = f"W{lw_disp:02d}'{str(ly)[-2:]}"
        this_token = f"W{tw_disp:02d}'{str(y)[-2:]}"
        return last_token, this_token

    # 自动：根据今天的 ISO 周计算显示周号 = ISO周 - 1
    y, iso_w, _ = d.isocalendar()
    tw_disp = iso_w - 1                   # 本周显示周号
    # 上周显示周号
    lw_disp = tw_disp - 1
    ly = y

    if lw_disp < 0:
        # 回滚上一年
        ly = y - 1
        lw_disp = last_iso_week_of_year(ly) - 1  # 上一年最后显示周号

    last_token = f"W{lw_disp:02d}'{str(ly)[-2:]}"
    this_token = f"W{tw_disp:02d}'{str(y)[-2:]}"
    return last_token, this_token


# ---------- 文件查找与复制 ----------
def find_file_for_week(base_dir: str, target_wyy: str) -> str | None:
    """精确查找指定周标（如 'W40\\'25'），多命中取修改时间最新；找不到返回 None"""
    target = normalize_week_token(target_wyy)
    hits: list[tuple[str, float]] = []
    with os.scandir(base_dir) as it:
        for de in it:
            if not de.is_file():
                continue
            m = WEEK_FILE_RE.match(de.name)
            if not m:
                continue
            w = int(m.group(1)); yy = m.group(2)
            token = normalize_week_token(f"W{w:02d}'{yy}")
            if token == target:
                try:
                    hits.append((de.path, de.stat().st_mtime))
                except OSError:
                    pass
    hits.sort(key=lambda x: x[1], reverse=True)
    return hits[0][0] if hits else None

def latest_match_by_mtime(base_dir: str) -> str | None:
    """按修改时间降序取最新的规范文件；没有则返回 None"""
    items: list[tuple[str, float]] = []
    with os.scandir(base_dir) as it:
        for de in it:
            if not de.is_file():
                continue
            if WEEK_FILE_RE.match(de.name):
                try:
                    items.append((de.path, de.stat().st_mtime))
                except OSError:
                    pass
    if not items:
        return None
    items.sort(key=lambda x: x[1], reverse=True)
    return items[0][0]

def make_this_week_name(from_name: str, wyy: str) -> str:
    """把文件名中的 Wxx'yy 替换为本周 wyy；不存在则追加"""
    base, ext = os.path.splitext(from_name)
    new_base = re.sub(rf"W(\d{{1,2}}){APOS_CLASS}(\d{{2}})", normalize_week_token(wyy), base, flags=re.IGNORECASE)
    if new_base == base:
        new_base = f"{base} {normalize_week_token(wyy)}"
    return new_base + ext

def copy_to_this_week(latest_path: str, this_week_token: str) -> str:
    """复制最近文件为本周命名；若源已是本周命名或目标已存在，则不重复复制"""
    src_name = os.path.basename(latest_path)
    dst_name = make_this_week_name(src_name, this_week_token)
    dst_path = os.path.join(BASE_DIR, dst_name)

    if os.path.abspath(dst_path) == os.path.abspath(latest_path):
        print(f"  本周文件已就绪：{dst_name} / This week's file is ready: {dst_name}")
        return latest_path
    if os.path.exists(dst_path):
        print(f"  本周文件已存在：{dst_name} / This week's file already exists: {dst_name}")
        return dst_path

    shutil.copy2(latest_path, dst_path)
    print(f"  复制为本周文件：{dst_name} / Copied as this week's file: {dst_name}")
    return dst_path


# ---------- Excel 自动化 ----------
def open_excel_silent():
    import os, sys, shutil, tempfile
    import pythoncom
    from win32com.client import DispatchEx, gencache

    pythoncom.CoInitialize()  # 稳妥起见

    try:
        # 优先确保 Excel 的类型库已生成（Excel TypeLib GUID）
        gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 9)
    except Exception:
        # 如果类型库确保失败，先清 gen_py 再重建
        try:
            shutil.rmtree(os.path.join(tempfile.gettempdir(), 'gen_py'), ignore_errors=True)
            # 某些环境 gen_py 会落在 site-packages 下，也顺手清掉
            for p in sys.path:
                gp = os.path.join(p, 'win32com', 'gen_py')
                shutil.rmtree(gp, ignore_errors=True)
        except Exception:
            pass
        # 再试一次
        gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 9)

    # 现在再起一个新的 Excel 实例（保持你原本的“独立实例”语义）
    ex = DispatchEx("Excel.Application")
    ex.Visible = False
    ex.DisplayAlerts = False
    ex.AskToUpdateLinks = False
    ex.ScreenUpdating = False
    ex.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    return ex

def open_wb_with_retry(path, tries=6, delay=1.0):
    last_err = None
    for _ in range(1, tries + 1):
        try:
            ex = open_excel_silent()
            wb = ex.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
            return ex, wb
        except Exception as e:
            last_err = e
            time.sleep(delay)
    raise RuntimeError(f"无法打开文件：{path}\n最后错误：{last_err} / "
                       f"Unable to open file: {path}\nLast error: {last_err}")

def col_to_index(ws, col):
    if isinstance(col, int): return col
    return int(ws.Range(f"{col}1").Column)

def last_row_in_col(ws, col):
    col_idx = col_to_index(ws, col)
    return int(ws.Cells(ws.Rows.Count, col_idx).End(XL_UP).Row)

def first_table(ws):
    try:
        return ws.ListObjects(1) if ws.ListObjects.Count > 0 else None
    except Exception:
        return None

def table_col_by_header(lo, header_text):
    for i in range(1, lo.ListColumns.Count + 1):
        if str(lo.ListColumns(i).Name).strip().lower() == header_text.strip().lower():
            return lo.ListColumns(i)
    return None

def get_validation_allowed_value(app, ws, addr, prefer_contains="reason code"):
    try:
        dv = ws.Range(addr).Validation
    except Exception:
        return None
    if getattr(dv, "Type", None) != 3:
        return None
    src = dv.Formula1
    if not src: return None

    values = []
    try:
        if src.startswith("="):
            res = app.Evaluate(src)
            try: vals = res.Value
            except Exception: vals = res
            if isinstance(vals, tuple):
                for row in vals:
                    if isinstance(row, tuple):
                        for v in row:
                            if v not in (None, ""): values.append(str(v))
                    else:
                        if row not in (None, ""): values.append(str(row))
        else:
            s = src[1:-1] if len(src) >= 2 and src[0] == '"' and src[-1] == '"' else src
            values = [x.strip() for x in s.split(",") if x.strip()]
    except Exception:
        pass

    if not values: return None
    pref = prefer_contains.lower().strip()
    for v in values:
        if pref in v.lower().strip():
            return v
    return values[0]

def find_last_reason_row(ws, allowed_bn: str | None):
    if not allowed_bn: return None
    last_bn = last_row_in_col(ws, "BN")
    if last_bn < 2: return None
    rng = ws.Range(f"BN2:BN{last_bn}")
    vals = rng.Value
    if isinstance(vals, tuple):
        col_vals = [(row[0] if isinstance(row, tuple) else row) for row in vals]
    else:
        col_vals = [vals]
    target = allowed_bn.strip().lower()
    for i in range(len(col_vals) - 1, -1, -1):
        v = "" if col_vals[i] is None else str(col_vals[i])
        if v.strip().lower() == target:
            return 2 + i
    return None

def hard_shrink_sheet(ws, last_keep_row: int, last_keep_col: int):
    total_rows = int(ws.Rows.Count)
    total_cols = int(ws.Columns.Count)
    if last_keep_row < total_rows:
        ws.Range(f"{last_keep_row+1}:{total_rows}").EntireRow.Delete()
    if last_keep_col < total_cols:
        ws.Range(ws.Cells(1, last_keep_col + 1), ws.Cells(1, total_cols)).EntireColumn.Delete()
    _ = ws.UsedRange  # 触发 UsedRange 更新

# ===== 新增：展开并移除所有筛选 =====
def expand_and_clear_filters(ws):
    """展开所有分组并清除工作表与表格的筛选"""
    # 1) 展开大纲分组（行/列）
    try:
        # Excel 的行/列分组最大只会到 8 级，这里给个充足的级别
        ws.Outline.ShowLevels(RowLevels=8, ColumnLevels=8)
    except Exception:
        try:
            # 少数版本需要位置参数
            ws.Outline.ShowLevels(8, 8)
        except Exception:
            pass

    # 2) 工作表级筛选清除
    try:
        # 若有手动 AutoFilter，先显示全部
        if getattr(ws, "FilterMode", False):
            ws.ShowAllData()
    except Exception:
        pass
    try:
        if getattr(ws, "AutoFilterMode", False):
            # 关闭工作表级 AutoFilter（不影响表格 ListObject 的标题筛选按钮）
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
                    # 兼容某些版本：用 Range.AutoFilter 显示全部
                    try:
                        lo.Range.AutoFilter(Field=1)  # 触发一次无条件的 AutoFilter
                        lo.AutoFilter.ShowAllData()
                    except Exception:
                        pass
    except Exception:
        pass
# ===== 新增结束 =====

def refresh_and_format(file_path: str):
    # Phase 1: 刷新
    excel = open_excel_silent()
    wb = excel.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)

    # --- 新增：打开后先“展开并移除所有filter”（对所有工作表更保险） ---
    try:
        for ws in wb.Worksheets:
            expand_and_clear_filters(ws)
    except Exception:
        pass
    # --- 新增结束 ---

    existing = {wb.Connections(i).Name for i in range(1, wb.Connections.Count + 1)} if wb.Connections.Count else set()
    for name in CONNECTIONS:
        if name in existing:
            try:
                wb.Connections(name).Refresh()
            except Exception:
                pass
    excel.CalculateUntilAsyncQueriesDone()
    wb.Save(); wb.Close(SaveChanges=True); excel.Quit()
    print("  刷新完成 / Refresh complete")

    # Phase 2: 写入 + 裁剪 + 硬缩
    app2, wb2 = open_wb_with_retry(file_path, tries=6, delay=1.2)
    ws = wb2.Sheets(SHEET_NAME)
    app2.EnableEvents = False; app2.ScreenUpdating = False

    # --- 新增：再次确保目标工作表已“展开并移除所有filter” ---
    try:
        expand_and_clear_filters(ws)
    except Exception:
        pass
    # --- 新增结束 ---

    try: ws.Unprotect()
    except Exception: pass

    lo = first_table(ws)

    # BN 合法值（来自数据验证）
    allowed_bn = get_validation_allowed_value(app2, ws, "BN2", prefer_contains="reason code") or "Reason Code"

    # 批量写入 BN；清空 BT（标题不动）
    did_bn = False
    if lo is not None and HEADER_BN:
        lc_bn = table_col_by_header(lo, HEADER_BN)
        if lc_bn is not None and lc_bn.DataBodyRange is not None:
            lc_bn.DataBodyRange.Value = allowed_bn
            did_bn = True
    if not did_bn:
        last_bn_write = last_row_in_col(ws, "BN")
        if last_bn_write >= 2:
            ws.Range(f"BN2:BN{last_bn_write}").Value = allowed_bn

    last_bt = last_row_in_col(ws, "BT")
    if last_bt >= 2:
        ws.Range(f"BT2:BT{last_bt}").ClearContents()

    # 以 BN 的“最后一个 Reason Code 行”为末行
    last_reason_row = find_last_reason_row(ws, allowed_bn)
    if last_reason_row is not None:
        if lo is not None and lo.Range is not None:
            header_row = lo.HeaderRowRange.Row
            first_col = lo.Range.Column
            last_col  = first_col + lo.Range.Columns.Count - 1
            new_last  = max(header_row, last_reason_row)
            lo.Resize(ws.Range(ws.Cells(header_row, first_col), ws.Cells(new_last, last_col)))
            hard_shrink_sheet(ws, last_keep_row=new_last, last_keep_col=last_col)
        else:
            last_col = int(ws.Cells(last_reason_row, ws.Columns.Count).End(XL_TOLEFT).Column)
            hard_shrink_sheet(ws, last_keep_row=last_reason_row, last_keep_col=last_col)

    wb2.Save(); wb2.Close(SaveChanges=True)
    app2.EnableEvents = True; app2.ScreenUpdating = True; app2.Quit()

    # 再开一次触发 UsedRange 生效（滚动条缩短）
    app3, wb3 = open_wb_with_retry(file_path, tries=3, delay=0.8)
    ws3 = wb3.Sheets(SHEET_NAME)
    _ = ws3.UsedRange
    wb3.Save(); wb3.Close(SaveChanges=True); app3.Quit()
    print("  表内整理完成 / In-sheet cleanup complete")


# ---------- 主流程 ----------
def main():
    print("==== Supplier – Weekly ====")
    last_token, this_token = compute_week_tokens()  # 例：("W40'25", "W41'25")
    print(f"周标：上周 {last_token} | 本周 {this_token} / Week tokens: last {last_token} | this {this_token}")

    # 1) 先按“上周周标”精确查找
    last_path = find_file_for_week(BASE_DIR, last_token)

    # 2) 精确找不到则退回“按修改时间最新”
    if not last_path:
        last_path = latest_match_by_mtime(BASE_DIR)

    if not last_path:
        # 最后兜底：用 glob（与旧逻辑一致）
        cands = [f for f in glob(os.path.join(BASE_DIR, GLOB_PATTERN)) if os.path.isfile(f)]
        if not cands:
            raise FileNotFoundError(f"未找到任何周文件：{BASE_DIR} / No weekly files found in {BASE_DIR}")
        cands.sort(key=os.path.getmtime, reverse=True)
        last_path = cands[0]

    print(f"上周文件：{os.path.basename(last_path)} / Last week's file: {os.path.basename(last_path)}")

    # 3) 复制成本周命名（若已存在则直接使用）
    cur_path = copy_to_this_week(last_path, this_token)
    print(f"本周文件：{os.path.basename(cur_path)} / This week's file: {os.path.basename(cur_path)}")

    # 4) 刷新 & 表内处理
    refresh_and_format(cur_path)

    print("✅ 完成 / Done")

if __name__ == "__main__":
    main()
