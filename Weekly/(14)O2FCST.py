# -*- coding: utf-8 -*-
"""
流程（稳定版，跨盘发布修复）：
1) 抓取最新 FCST 邮件附件（加密）
2) 总是从 Archived 里选“最新有效”的 .xlsx 模板（跳过 ~$ 锁文件；无效就找下一个）
3) 复制到根目录并命名为当周周一：Lumileds FCST (TBL) - <YYYYMMDD>.xlsx（去只读）
4) 从源附件 Weekly 表抽取数据 → 粘贴到目标工作簿第 2 个工作表（C2 起），把 #N/A/错误/非数值→0.00，格式 0.00
5) 只刷新 Workbook.Connection 'Query - Table1' 一次（写死）
6) 保存：一律 SaveAs 到本地临时 → 关闭工作簿 → 复制到目标目录的同目录临时(__swap__) → 同目录原子替换为正式文件（带重试）
7) 触发 SQL Agent Job
"""

import os
import re
import stat
import time
import shutil
import zipfile
import tempfile
import datetime as dt
from pathlib import Path
from typing import Optional, Tuple, List, Sequence, Any

import pythoncom  # type: ignore
import win32com.client as win32  # type: ignore

from Utils.graph_mail_attachment_tool import GraphMailAttachmentTool
from Utils.sql_agent_tool import SqlAgentTool

# ============== 可配置 ==============
TENANT_ID = "5c2be51b-4109-461d-a0e7-521be6237ce2"
CLIENT_ID = "09004044-1c60-48e5-b1eb-bb42b3892006"

MAIL_CONTAINS  = "Lumileds_FCST"
MAIL_EQUALS    = None
MAIL_EXT       = ".xlsx"
MAIL_DAYS_BACK = 30
DOWNLOAD_DIR   = r"C:\WeeklyReport\Download_From_Email"

# O2 FCST 目录
O2_BASE     = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\External\O2 FCST"
O2_ARCHIVED = os.path.join(O2_BASE, "Archived")

# 源 FCST 打开密码
SOURCE_FILE_PASSWORD = "greenapple"

# 源/目标工作表与粘贴定位
SRC_SHEET_NAME = "Weekly"   # 源：取这个表
SRC_START_COL  = 8          # 从 H 列（1-based）起找数据块
DST_SHEET_INDEX = 2         # ⚠️ 目标工作簿“第 2 个工作表”
DST_START_ROW   = 2         # 目标从第 2 行起（保留标题）
DST_START_COL   = 3         # 目标从第 3 列（C 列）起（保留 A、B）

# 目标文件命名
TARGET_PREFIX = "Lumileds FCST (TBL) - "
CANON_NAME_RE = re.compile(r"^(Lumileds FCST \(TBL\) - )(\d{8})(?:_.+)?\.xlsx$", re.IGNORECASE)

# ============== 小工具 ==============
def ensure_dir(p: str | os.PathLike) -> Path:
    Path(p).mkdir(parents=True, exist_ok=True)
    return Path(p)

def timestamp_name() -> str:
    now = dt.datetime.now()
    return now.strftime("%Y%m%d_%Y_%m_%d_%H_%M_%S")

def monday_str(today: Optional[dt.date] = None) -> str:
    """返回当前周的周一日期（YYYYMMDD）。周一=0 ... 周日=6"""
    if today is None:
        today = dt.date.today()
    monday = today - dt.timedelta(days=today.weekday())
    return monday.strftime("%Y%m%d")

def _make_writable(path: Path) -> None:
    """确保文件非只读（Windows/UNC）。"""
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        try:
            # 仅在 Windows 上有效；静默处理
            os.system(f'attrib -R "{str(path)}"')
        except Exception:
            pass
    except Exception:
        pass

def is_valid_xlsx(path: Path, min_kb: int = 50) -> bool:
    """快速校验 xlsx：大小阈值 + ZIP 魔数 + 能读 [Content_Types].xml"""
    try:
        if not path.exists() or not path.is_file():
            return False
        if path.stat().st_size < min_kb * 1024:
            return False
        with open(path, "rb") as f:
            if f.read(2) != b"PK":
                return False
        with zipfile.ZipFile(path, "r") as zf:
            zf.getinfo("[Content_Types].xml")
        return True
    except Exception:
        return False

def select_archived_template(archived_dir: str) -> Path:
    """
    在 Archived 中按修改时间倒序选择第一个“非锁文件且通过有效性校验”的 .xlsx。
    若都不合格则抛错。
    """
    p = Path(archived_dir)
    candidates = [x for x in p.glob("*.xlsx") if x.is_file() and not x.name.startswith("~$")]
    candidates.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    if not candidates:
        raise RuntimeError(f"未在 Archived 目录找到 xlsx：{archived_dir} / No xlsx found in Archived: {archived_dir}")
    for c in candidates:
        if is_valid_xlsx(c):
            print(f"[TEMPLATE] 采用 Archived 模板：{c.name}  ({round(c.stat().st_size/1024)} KB) / "
                  f"Using Archived template: {c.name} ({round(c.stat().st_size/1024)} KB)")
            return c
        else:
            print(f"[SKIP] 无效/损坏，跳过：{c.name} / Invalid/corrupt, skipped: {c.name}")
    raise RuntimeError("Archived 中没有可用的有效 .xlsx 模板（均无效或为锁文件）。 / "
                       "No valid .xlsx templates in Archived (all invalid or lock files).")

def copy_archived_to_base_with_monday_name(archived_file: Path, base_dir: str, monday: str) -> Path:
    """复制模板到根目录并命名为当周周一（覆盖），采用临时名防半成品，并去除只读属性。"""
    ensure_dir(base_dir)
    target_name = f"{TARGET_PREFIX}{monday}.xlsx"
    target_path = Path(base_dir) / target_name
    tmp_path = Path(base_dir) / f"{TARGET_PREFIX}{timestamp_name()}.tmp.xlsx"
    shutil.copy2(archived_file, tmp_path)
    _make_writable(tmp_path)
    if not is_valid_xlsx(tmp_path):
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise RuntimeError(f"从 Archived 复制出的文件无效：{archived_file} / Invalid file copied from Archived: {archived_file}")
    try:
        if target_path.exists():
            _make_writable(target_path)
            target_path.unlink()
    except Exception:
        pass
    os.replace(tmp_path, target_path)
    _make_writable(target_path)
    print(f"[COPY] Archived → {target_path.name}")
    return target_path

def canonicalize_o2_filename(path: Path) -> Path:
    """当前流程已直接按周一定名；此函数留作兼容（通常不会触发重命名）。"""
    m = CANON_NAME_RE.match(path.name)
    if not m:
        return path
    prefix, yyyymmdd = m.group(1), m.group(2)
    new_name = f"{prefix}{yyyymmdd}.xlsx"
    new_path = path.with_name(new_name)
    if new_path != path:
        try:
            if new_path.exists():
                _make_writable(new_path)
                new_path.unlink()
        except Exception:
            pass
        os.replace(path, new_path)
        _make_writable(new_path)
        print(f"[RENAME] {path.name} → {new_path.name}")
    return new_path

def safe_stage_to_dst_and_swap(src_local: Path, dst_remote: Path, retries: int = 6, sleep_sec: float = 1.0) -> None:
    """
    把本地临时文件 src_local 发布到远端目标 dst_remote：
      - 先复制到目标目录的“同目录临时文件” (__swap__.xlsx)
      - 再用 os.replace 在同目录内原子替换为最终文件
    避免 WinError 17（跨盘 replace 不允许），并对共享冲突重试。
    """
    dst_dir = dst_remote.parent
    ensure_dir(dst_dir)

    for i in range(retries):
        tmp_in_dst = dst_dir / f"{dst_remote.stem}.__swap__.xlsx"
        try:
            # 清理可能存在的上次残留
            try:
                if tmp_in_dst.exists():
                    _make_writable(tmp_in_dst)
                    tmp_in_dst.unlink()
            except Exception:
                pass

            # 复制到目标目录的同级临时
            shutil.copy2(src_local, tmp_in_dst)
            _make_writable(tmp_in_dst)

            # 同目录原子替换到最终名
            _make_writable(dst_remote)
            if dst_remote.exists():
                try:
                    dst_remote.unlink()
                except Exception:
                    pass
            os.replace(tmp_in_dst, dst_remote)
            _make_writable(dst_remote)
            return
        except Exception:
            # 清理临时并重试
            try:
                if tmp_in_dst.exists():
                    _make_writable(tmp_in_dst)
                    tmp_in_dst.unlink()
            except Exception:
                pass
            if i == retries - 1:
                raise
            time.sleep(sleep_sec)

# ============== Graph：抓附件 ==============
def fetch_latest_fcst_attachment() -> Path:
    ensure_dir(DOWNLOAD_DIR)
    tool = GraphMailAttachmentTool(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        scopes="Mail.Read offline_access",
        token_cache="graph_token_cache.json",
        request_timeout=60,
    )
    paths: List[Path] = tool.download_latest_attachments(
        contains=MAIL_CONTAINS,
        equals=MAIL_EQUALS,
        ext=MAIL_EXT,
        need_count=1,
        days_back=MAIL_DAYS_BACK,
        save_dir=DOWNLOAD_DIR,
    )
    if not paths:
        raise RuntimeError("没有在邮箱里找到匹配的 FCST 附件，请检查关键词/时间范围。 / "
                           "No matching FCST attachment found in mailbox; check keywords/time range.")
    print(f"[OK] 已下载最新附件：{paths[0]} / Downloaded latest attachment: {paths[0]}")
    return paths[0]

# ============== Excel 辅助 ==============
def read_used_bounds(ws) -> Tuple[int, int]:
    ur = ws.UsedRange
    first_row = ur.Row
    first_col = ur.Column
    nrows = ur.Rows.Count
    ncols = ur.Columns.Count
    last_row = first_row + nrows - 1
    last_col = first_col + ncols - 1
    return last_row, last_col

def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, tuple):
        return [[values]]
    rows = []
    for r in values:
        if isinstance(r, tuple):
            rows.append(list(r))
        else:
            rows.append([r])
    return rows

def is_blank_row(row_vals: Sequence[Any]) -> bool:
    return all((v is None or (isinstance(v, str) and v.strip() == "")) for v in row_vals)

def numeric_ratio(row_vals: Sequence[Any]) -> float:
    total = len(row_vals)
    if total == 0:
        return 0.0
    numerics = sum(1 for v in row_vals if isinstance(v, (int, float)) and v is not None)
    return numerics / total

def is_date_like(v: Any) -> bool:
    import datetime as _dt
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return False
    if isinstance(v, (_dt.date, _dt.datetime)):
        return True
    if isinstance(v, (int, float)):
        return 20000 <= float(v) <= 60000  # 粗识别 Excel 序列日期
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%Y.%m.%d"):
            try:
                _dt.datetime.strptime(s, fmt)
                return True
            except Exception:
                pass
    return False

def date_ratio(row_vals: Sequence[Any]) -> float:
    total = len(row_vals)
    if total == 0:
        return 0.0
    cnt = sum(1 for v in row_vals if is_date_like(v))
    return cnt / total

def expand_all_outlines(ws) -> None:
    try:
        ws.Outline.ShowLevels(RowLevels=8, ColumnLevels=8)
    except Exception:
        pass
    try:
        for lo in ws.ListObjects:
            try:
                lo.AutoFilter.ShowAllData()
            except Exception:
                pass
    except Exception:
        pass

def detect_first_numeric_block(ws, start_col: int) -> Tuple[int, int, int]:
    last_row, last_col = read_used_bounds(ws)
    if last_col < start_col:
        raise RuntimeError("源表最后一列在 H 之前，无法复制。 / Source sheet last column is before H; cannot copy.")
    rng = ws.Range(ws.Cells(1, start_col), ws.Cells(last_row, last_col))
    vals = to_2d(rng.Value)
    row_start = None
    for idx, row in enumerate(vals, start=1):
        if is_blank_row(row):
            continue
        if numeric_ratio(row) >= 0.5:
            row_start = idx
            break
    if row_start is None:
        raise RuntimeError("没有检测到数值数据块的起始行，可能页面仍是折叠或表为空。 / "
                           "Start row of numeric block not detected; sheet may be collapsed or empty.")
    r = row_start
    while r <= len(vals) and not is_blank_row(vals[r - 1]):
        r += 1
    row_end = r - 1
    if row_end < row_start:
        raise RuntimeError("数据块范围检测异常。 / Data block range detection failed.")
    return row_start, row_end, last_col

def find_date_header_last_col(ws, start_col: int, data_row_start: int) -> Tuple[Optional[int], int]:
    last_row, last_col = read_used_bounds(ws)
    if data_row_start <= 1:
        return None, last_col
    scan_rng = ws.Range(ws.Cells(1, start_col), ws.Cells(data_row_start - 1, last_col))
    rows = to_2d(scan_rng.Value)
    for idx in range(len(rows), 0, -1):
        row = rows[idx - 1]
        if not row or is_blank_row(row):
            continue
        if date_ratio(row) >= 0.5:
            last_date_col_offset = None
            for j, v in enumerate(row):
                if is_date_like(v):
                    last_date_col_offset = j
            if last_date_col_offset is not None:
                return idx, start_col + last_date_col_offset
    return None, last_col

# --- 把单元格值强制变成数值（#N/A / 错误 / 空白 / 非数字 → 0.00） ---
EXCEL_ERROR_CODES = {
    -2146826281,  # #DIV/0!
    -2146826259,  # #VALUE!
    -2146826265,  # #REF!
    -2146826246,  # #N/A
    -2146826252,  # #NUM!
    -2146826269,  # #NAME?
    -2146826273,  # #NULL!
}
def coerce_to_number(v: Any) -> float:
    try:
        if isinstance(v, (int, float)) and int(float(v)) in EXCEL_ERROR_CODES:
            return 0.0
    except Exception:
        pass
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().upper()
        if s in {"#N/A", "#NA", "N/A", "NA", "#DIV/0!", "#REF!", "#VALUE!", "#NUM!", "#NULL!", "#NAME?"}:
            return 0.0
        try:
            return float(s.replace(",", ""))
        except Exception:
            return 0.0
    return 0.0

def sanitize_numeric_block(values_2d: List[List[Any]]) -> List[List[float]]:
    if not values_2d:
        return values_2d
    out: List[List[float]] = []
    for row in values_2d:
        out.append([coerce_to_number(v) for v in row])
    return out

def paste_values(ws_dst, top_row: int, left_col: int, values_2d: List[List[Any]]):
    if not values_2d:
        return
    nrows = len(values_2d)
    ncols = len(values_2d[0])
    tgt = ws_dst.Range(ws_dst.Cells(top_row, left_col), ws_dst.Cells(top_row + nrows - 1, left_col + ncols - 1))
    tgt.Value = values_2d
    tgt.NumberFormat = "0.00"  # 两位小数

def clear_target_region(ws_dst, start_row: int = 2, start_col: int = 3):
    max_row = 1048576
    max_col = 16384   # XFD
    ws_dst.Range(ws_dst.Cells(start_row, start_col), ws_dst.Cells(max_row, max_col)).ClearContents()

def list_all_queries_and_connections(wb):
    print("\n[INFO] ===== 连接 / 查询清单（供核对） ===== / Connections & Queries (for review) =====")
    try:
        cons = wb.Connections
        n = cons.Count
        if n:
            print(f"Workbook.Connections ({n}):")
            for i in range(1, n + 1):
                c = cons.Item(i)
                print(f"  - {i}. Name='{getattr(c, 'Name', None)}' | Type='{getattr(c, 'Type', None)}'")
        else:
            print("Workbook.Connections: <none>")
    except Exception as e:
        print(f"Workbook.Connections 读取异常: {e} / Workbook.Connections read error: {e}")
    try:
        qs = wb.Queries
        n = qs.Count
        if n:
            print(f"Workbook.Queries ({n}):")
            for i in range(1, n + 1):
                q = qs.Item(i)
                print(f"  - {i}. QueryName='{getattr(q, 'Name', None)}'")
        else:
            print("Workbook.Queries: <none>")
    except Exception as e:
        print(f"Workbook.Queries 读取异常: {e} / Workbook.Queries read error: {e}")
    try:
        for sh in wb.Worksheets:
            qts = getattr(sh, "QueryTables", None)
            if qts and qts.Count:
                print(f"{sh.Name}.QueryTables ({qts.Count}):")
                for i in range(1, qts.Count + 1):
                    qt = qts.Item(i)
                    print(f"  - {i}. Name='{getattr(qt, 'Name', None)}'")
            los = getattr(sh, "ListObjects", None)
            if los and los.Count:
                print(f"{sh.Name}.ListObjects ({los.Count}):")
                for i in range(1, los.Count + 1):
                    lo = los.Item(i)
                    print(f"  - {i}. Name='{getattr(lo, 'Name', None)}'")
    except Exception as e:
        print(f"Worksheets 查询枚举异常: {e} / Worksheets enumeration error: {e}")
    print("[INFO] ===== 以上为本文件的连接/查询清单 ===== / End of connections/queries list =====\n")

# ============== 写死：只刷新 Query - Table1 ==============
def refresh_query_table1_only(wb):
    """
    不再“找第一个成功目标”，只尝试刷新 Workbook.Connections("Query - Table1") 一次。
    刷新失败就打印 warning，继续后续流程，不抛异常、不等待。
    """
    try:
        conn = wb.Connections("Query - Table1")
        print("[REFRESH] Workbook.Connection: Query - Table1 ...")
        conn.Refresh()
        print("[REFRESH] Done via Connection 'Query - Table1'.")
    except Exception as e:
        print(f"[WARN] 刷新 Connection 'Query - Table1' 失败：{e}（不中断流程） / "
              f"Failed to refresh 'Query - Table1': {e} (continuing)")

# ============== 主流程 ==============
def run_pipeline():
    # 0) 预检查目录
    ensure_dir(O2_BASE)
    ensure_dir(O2_ARCHIVED)

    # 1) 抓取最新邮件附件（源）
    src_attachment = fetch_latest_fcst_attachment()

    # 2) 选择 Archived 中“最新有效模板”（跳过 ~$；无效就找下一个）
    archived_template = select_archived_template(O2_ARCHIVED)

    # 3) 复制到根目录并命名为“当周周一”，并移除只读属性
    monday = monday_str()  # 当周周一 YYYYMMDD
    dst_file = copy_archived_to_base_with_monday_name(archived_template, O2_BASE, monday)
    dst_file = canonicalize_o2_filename(Path(dst_file))  # 一般不会触发
    _make_writable(dst_file)

    # 4) 打开 Excel（COM）并处理
    pythoncom.CoInitialize()
    excel = None
    wb_src = None
    wb_dst = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        print("[EXCEL] 打开源附件（带密码）... / Opening source attachment (with password)...")
        # 使用参数顺序传递密码，避免某些环境命名参数问题
        wb_src = excel.Workbooks.Open(
            str(src_attachment),
            0,          # UpdateLinks
            False,      # ReadOnly
            None,       # Format
            SOURCE_FILE_PASSWORD  # Password
        )
        ws_src = wb_src.Worksheets(SRC_SHEET_NAME)

        print("[EXCEL] 展开源表分组/大纲/筛选... / Expanding source sheet groups/outlines/filters...")
        expand_all_outlines(ws_src)

        print("[EXCEL] 打开目标文件（当周周一）... / Opening target file (this Monday)...")
        wb_dst = excel.Workbooks.Open(
            Filename=str(dst_file),
            CorruptLoad=1,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )
        ws_dst = wb_dst.Worksheets(DST_SHEET_INDEX)  # ⚠️ 第二个工作表

        # —— 定位数据块并按“日期行”确定列终点 ——
        row_start, row_end, used_last_col = detect_first_numeric_block(ws_src, start_col=SRC_START_COL)
        _, last_date_col = find_date_header_last_col(ws_src, start_col=SRC_START_COL, data_row_start=row_start)
        copy_col_end = last_date_col if last_date_col and last_date_col >= SRC_START_COL else used_last_col

        # 源数据（不含日期行）→ 数值化
        src_rng = ws_src.Range(ws_src.Cells(row_start, SRC_START_COL), ws_src.Cells(row_end, copy_col_end))
        values_raw = to_2d(src_rng.Value)
        values_2d = sanitize_numeric_block(values_raw)
        print(f"[COPY] 数据行 {row_start}-{row_end}，列 H..{copy_col_end}（已把 #N/A/错误/非数字→0.00） / "
              f"Rows {row_start}-{row_end}, cols H..{copy_col_end} (#N/A/errors/non-numeric -> 0.00)")

        # —— 清空目标区域（C2:XFD1048576） ——
        print("[CLEAR] 清空目标工作表（二号表） C2:XFD1048576 ... / Clearing target sheet (sheet 2) C2:XFD1048576 ...")
        clear_target_region(ws_dst, start_row=DST_START_ROW, start_col=DST_START_COL)

        # 5) 粘贴到目标 C2 起（并设置两位小数格式）
        print("[PASTE] 粘贴到目标（二号表） C2 起 ... / Pasting to target (sheet 2) from C2 ...")
        paste_values(ws_dst, DST_START_ROW, DST_START_COL, values_2d)

        # 6) 列出连接/查询清单（便于核对）
        list_all_queries_and_connections(wb_dst)

        # 6b) 写死刷新：只刷 'Query - Table1'，不等待
        refresh_query_table1_only(wb_dst)

        # ======= 统一安全保存：SaveAs 本地临时 → 关闭 → 发布到 UNC（同目录替换） =======
        local_tmp_dir = Path(tempfile.gettempdir())
        local_tmp_path = local_tmp_dir / f"{Path(dst_file).stem}.__tmp__.xlsx"
        try:
            if local_tmp_path.exists():
                _make_writable(local_tmp_path)
                local_tmp_path.unlink()
        except Exception:
            pass

        print(f"[SAVE] 安全保存到本地临时：{local_tmp_path} / Safe save to local temp: {local_tmp_path}")
        wb_dst.SaveAs(Filename=str(local_tmp_path), FileFormat=51)

        # 关闭 workbook，释放句柄（尤其是 Power Query/连接可能持有）
        wb_dst.Close(SaveChanges=False)
        wb_dst = None

        # 发布到目标目录：先复制到 __swap__ 再同目录 os.replace
        print(f"[SAVE] 发布到目标（同目录替换）：{dst_file} / Publish to target (same-folder replace): {dst_file}")
        safe_stage_to_dst_and_swap(local_tmp_path, Path(dst_file))
        print(f"[SAVE] 完成替换：{dst_file} / Replace completed: {dst_file}")

        # 清理本地临时
        try:
            if local_tmp_path.exists():
                _make_writable(local_tmp_path)
                local_tmp_path.unlink()
        except Exception:
            pass

        print(f"[DONE] 已完成：从 Archived 复制当周模板 → 粘贴到【第 2 个工作表】并把 #N/A→0.00；刷新 'Query - Table1'；最终文件：{dst_file} / "
              f"Done: copied weekly template from Archived -> pasted to sheet 2 and #N/A->0.00; refreshed 'Query - Table1'; final file: {dst_file}")

    except Exception as e:
        print(f"[ERROR] 运行失败：{e} / Run failed: {e}")
        raise
    finally:
        # 关书不保存（目标已保存），源不保存
        try:
            if wb_src is not None:
                wb_src.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if wb_dst is not None:
                wb_dst.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.DisplayAlerts = False
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # # 7) 触发 SQL Agent Job（可选）
    # tool = SqlAgentTool(server="tcp:10.80.127.71,1433")
    # result = tool.run_job(
    #     job_name="Lumileds BI - SC O2 Forecast",
    #     archive_dir=O2_ARCHIVED,
    #     timeout=1800,
    #     poll_interval=3,
    #     fuzzy=False,
    # )
    # print(result)

if __name__ == "__main__":
    run_pipeline()
