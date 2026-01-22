# -*- coding: utf-8 -*-
import os
import re
import time
import shutil
import win32com.client as win32
from contextlib import suppress

# === 源/目标文件夹 ===
SRC_FOLDER = r"\\Mp1do4ce0373ndz\C\MonthlyRawFile"
DST_FOLDER = r"\\Mp1do4ce0373ndz\d\Reporting\Raw\Inventory"

FILES = ["CN MB52 Raw.xls", "MY MB52 Raw.xls", "US MB52 Raw.xls", "SG MB52 Raw.xls"]
COPY_ONLY = ["MB5TD Raw.xls"]

# Excel 常量
xlCellTypeConstants    = 2
xlCellTypeFormulas     = -4123
# 不再改动 Application.Calculation，避免“Unable to set the Calculation property...”异常

# =============== 公共基础函数 ===============
def copy_from_weekly_to_inventory() -> list[str]:
    os.makedirs(DST_FOLDER, exist_ok=True)
    copied = []
    for fname in FILES + COPY_ONLY:
        src = os.path.join(SRC_FOLDER, fname)
        dst = os.path.join(DST_FOLDER, fname)
        if not os.path.exists(src):
            print(f"⚠ 源文件不存在（跳过）: {src} / Source file missing (skipped): {src}")
            continue
        try:
            shutil.copy2(src, dst)
            print(f"📥 已复制: {fname} / Copied: {fname}")
            copied.append(dst)
        except Exception as e:
            print(f"❌ 复制失败: {fname} -> {e} / Copy failed: {fname} -> {e}")
    return copied

def to_text_full_digits(ws, col_letter: str) -> None:
    """
    两步保真：把长数字转成文本且不丢位
      1) 设为 "0" 让 .Text 展示完整数字（避免科学计数法）
      2) 把 .Text 写回文本，并设为 "@"
    """
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 1:
        return
    rng = ws.Range(f"{col_letter}1:{col_letter}{last_row}")
    rng.NumberFormat = "0"
    with suppress(Exception):
        for c in rng.SpecialCells(xlCellTypeConstants):
            t = c.Text
            if t:
                c.NumberFormat = "@"
                c.Value = "'" + t
    with suppress(Exception):
        for c in rng.SpecialCells(xlCellTypeFormulas):
            t = c.Text
            if t:
                c.NumberFormat = "@"
                c.Value = "'" + t
    rng.NumberFormat = "@"

def set_col_text(ws, col_letter: str) -> None:
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 1:
        return
    ws.Range(f"{col_letter}1:{col_letter}{last_row}").NumberFormat = "@"

# =============== 新逻辑（用于 A/L） ===============
_DMY_PATTERN = re.compile(r"^\s*(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\s*$")
_PREFIX_DMY = re.compile(r"^\s*(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})(\s+.*)?$")

def _year4(y):
    y = int(y)
    return 1900 + y if y < 100 and y >= 50 else (2000 + y if y < 100 else y)

def _collect_preview(ws, col, maxn=3, need_prefix_date=False):
    res = []
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 2:
        return res
    base = 2
    rng = ws.Range(f"{col}{base}:{col}{last_row}")
    try:
        vals = rng.Value
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        for i, row in enumerate(vals):
            val = row[0] if isinstance(row, tuple) else row
            if val in (None, ""):
                continue
            text = None
            with suppress(Exception):
                text = ws.Cells(base + i, col).Text
            if text is None:
                text = str(val)
            if need_prefix_date and not _PREFIX_DMY.match(str(text)):
                continue
            res.append(text)
            if len(res) >= maxn:
                break
    except Exception:
        pass
    return res

def force_date_reformat_new(app, ws, cols):
    """新逻辑：用于 L 列（将 dd.mm.yyyy 等文本转成真正日期，并显示 MM/DD/YYYY）"""
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 2:
        return
    for col in cols:
        before = _collect_preview(ws, col, 3)
        print(f"🗓 即将格式化列 {col} 为 MM/DD/YYYY；样本(前)：{before} / "
              f"Formatting column {col} to MM/DD/YYYY; sample (before): {before}")
        rng = ws.Range(f"{col}2:{col}{last_row}")
        with suppress(Exception):
            for c in rng.SpecialCells(xlCellTypeConstants):
                v = c.Value
                if isinstance(v, str):
                    m = _DMY_PATTERN.match(v.strip())
                    if m:
                        d = int(m.group(1)); mth = int(m.group(2)); y = _year4(m.group(3))
                        with suppress(Exception):
                            c.Value = app.WorksheetFunction.Date(y, mth, d)
        with suppress(Exception):
            rng.NumberFormat = "mm/dd/yyyy"
        after = _collect_preview(ws, col, 3)
        print(f"✅ 列 {col} 已设为 MM/DD/YYYY；样本(后)：{after} / "
              f"Column {col} set to MM/DD/YYYY; sample (after): {after}")

def reformat_prefix_date_in_text(ws, col):
    """A 列：仅替换文本前缀的日期，保留后续说明文本"""
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 2:
        return
    before = _collect_preview(ws, col, 3, need_prefix_date=True)
    print(f"🗓 即将替换列 {col} 文本前缀日期；样本(前)：{before} / "
          f"Replacing date prefix text in column {col}; sample (before): {before}")
    rng = ws.Range(f"{col}2:{col}{last_row}")
    with suppress(Exception):
        for cell in rng:
            v = cell.Value
            if not isinstance(v, str):
                continue
            m = _PREFIX_DMY.match(v.strip())
            if not m:
                continue
            d, mo, y = int(m.group(1)), int(m.group(2)), _year4(m.group(3))
            tail = m.group(4) or ""
            cell.NumberFormat = "@"
            cell.Value = f"{mo:02d}/{d:02d}/{y:04d}{tail}"
    after = _collect_preview(ws, col, 3, need_prefix_date=True)
    print(f"✅ 列 {col} 文本前缀日期已替换；样本(后)：{after} / "
          f"Column {col} date prefix replaced; sample (after): {after}")

# =============== 旧逻辑（用于 AQ/AR 以及 MB5TD 的 R/S） ===============
_DMY_PATTERN_OLD = re.compile(r"^\s*(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\s*$")

def _parse_dmy_token(token: str):
    if not isinstance(token, str):
        return None
    m = _DMY_PATTERN_OLD.match(token)
    if not m:
        return None
    d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y = 1900 + y if y >= 50 else 2000 + y
    return y, mth, d

def force_date_reformat_legacy(app, ws, cols):
    """旧逻辑：笨但稳定（用于 AQ/AR、MB5TD 的 R/S）"""
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 2:
        return
    for col in cols:
        rng = ws.Range(f"{col}2:{col}{last_row}")
        before = []
        try:
            for i in range(2, min(5, last_row)):
                before.append(ws.Cells(i, col).Text)
        except Exception:
            pass
        print(f"🗓 即将格式化列 {col}（旧逻辑）；样本(前)：{before} / "
              f"Formatting column {col} (legacy); sample (before): {before}")
        with suppress(Exception):
            for c in rng.SpecialCells(xlCellTypeConstants):
                v = c.Value
                if not isinstance(v, str):
                    continue
                parsed = _parse_dmy_token(v.strip())
                if parsed:
                    y, mth, d = parsed
                    try:
                        c.Value = app.WorksheetFunction.Date(y, mth, d)
                    except Exception:
                        c.Value = f"{mth:02d}/{d:02d}/{y:04d}"
        with suppress(Exception):
            rng.NumberFormat = "mm/dd/yyyy"
        after = []
        try:
            for i in range(2, min(5, last_row)):
                after.append(ws.Cells(i, col).Text)
        except Exception:
            pass
        print(f"✅ 列 {col} 已设为 MM/DD/YYYY（旧逻辑）；样本(后)：{after} / "
              f"Column {col} set to MM/DD/YYYY (legacy); sample (after): {after}")

# =============== 额外工具：确保 A 列存在（MB5TD 用） ===============
def ensure_leading_blank_column(ws):
    a1 = str(ws.Cells(1, 1).Value or "").strip().lower()
    if a1 in {"material"}:
        ws.Columns("A").Insert()  # 原 A 整列右移
    for r in (1, 2):
        with suppress(Exception):
            c = ws.Cells(r, 1)
            c.NumberFormat = "@"
            c.Value = chr(160)  # NBSP
    _ = ws.UsedRange

# =============== 主流程（MB52） ===============
def process_one_excel(excel_app, path, is_sg, open_retries=3, open_sleep=2.0):
    wb = None
    for attempt in range(1, open_retries + 1):
        try:
            wb = excel_app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False, Notify=False)
            break
        except Exception as e:
            print(f"⏳ 打开失败: {e}，重试({attempt}) / Open failed: {e}, retry ({attempt})"); time.sleep(open_sleep)
    if not wb:
        return
    try:
        ws = wb.ActiveSheet
        # 原始逻辑
        with suppress(Exception):
            if is_sg:
                ws.Columns("N").Delete()
        to_text_full_digits(ws, "C")
        set_col_text(ws, "N")

        # A + L（新逻辑）
        reformat_prefix_date_in_text(ws, "A")
        force_date_reformat_new(excel_app, ws, ["L"])
        # AQ + AR（旧逻辑）
        force_date_reformat_legacy(excel_app, ws, ["AQ", "AR"])

        wb.Save()
        print(f"✅ 已格式化: {os.path.basename(path)} / Formatted: {os.path.basename(path)}")
    finally:
        with suppress(Exception): wb.Close(SaveChanges=True)

# =============== 主流程（MB5TD） ===============
def process_mb5td(excel_app, path, open_retries=3, open_sleep=2.0):
    """
    MB5TD 的完整处理：
      - A 列保留（空列不被裁掉）
      - B、U 两列：两步保真为文本（防科学计数法/丢位）
      - A 列前缀日期 -> MM/DD/YYYY（仅改前缀）
      - L 列（新逻辑）-> MM/DD/YYYY
      - R / S 列（旧逻辑）-> MM/DD/YYYY
    """
    wb = None
    for attempt in range(1, open_retries + 1):
        try:
            wb = excel_app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False, Notify=False)
            break
        except Exception as e:
            print(f"⏳ 打开失败(MB5TD): {e}，重试({attempt}) / Open failed (MB5TD): {e}, retry ({attempt})"); time.sleep(open_sleep)
    if not wb:
        return

    try:
        ws = wb.ActiveSheet

        # 先保住 A 列（避免保存时被 Excel 裁掉）
        ensure_leading_blank_column(ws)

        # B / U 列：两步保真为文本（避免科学计数法、尾部 00000）
        to_text_full_digits(ws, "B")
        to_text_full_digits(ws, "U")

        # A 列（文本前缀日期 → MM/DD/YYYY，仅改前缀部分）
        reformat_prefix_date_in_text(ws, "A")

        # L 列：新逻辑转换为真正日期显示
        force_date_reformat_new(excel_app, ws, ["L"])

        # R / S 列：旧逻辑转换为真正日期显示
        force_date_reformat_legacy(excel_app, ws, ["R", "S"])

        wb.Save()
        print(f"✅ 已处理 MB5TD：A列保留、B/U 文本保真、A前缀日期、L日期、R/S日期。 / "
              f"MB5TD processed: A kept, B/U text, A prefix date, L date, R/S date.")
    finally:
        with suppress(Exception): wb.Close(SaveChanges=True)

# =============== Runner ===============
def main():
    print("==== Step 1: 文件复制 / File copy ====")
    copied = copy_from_weekly_to_inventory()
    if not copied:
        print("⚠ 无可处理文件。 / No files to process."); return

    print("\n==== Step 2: 格式化处理 / Formatting ====")
    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    # 不必切换 Application.Calculation；部分环境会抛异常
    with suppress(Exception): excel.ScreenUpdating = False
    with suppress(Exception): excel.EnableEvents = False

    try:
        # 处理 4 个 MB52
        for f in FILES:
            dst = os.path.join(DST_FOLDER, f)
            if os.path.exists(dst):
                process_one_excel(excel, dst, is_sg=("SG MB52" in f))

        # 处理 MB5TD
        for f in COPY_ONLY:
            if f == "MB5TD Raw.xls":
                dst = os.path.join(DST_FOLDER, f)
                if os.path.exists(dst):
                    print(f"\n—— 处理 MB5TD: {f} / Processing MB5TD: {f}")
                    process_mb5td(excel, dst)

    finally:
        with suppress(Exception): excel.EnableEvents = True
        with suppress(Exception): excel.ScreenUpdating = True
        with suppress(Exception): excel.Quit()

    print("\n🎉 全部完成。 / All done.")

if __name__ == "__main__":
    main()
