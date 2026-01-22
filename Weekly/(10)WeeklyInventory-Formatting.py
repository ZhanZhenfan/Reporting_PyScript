# -*- coding: utf-8 -*-
import os
import time
import shutil
import win32com.client as win32
from contextlib import suppress

# === 源/目标文件夹 ===
SRC_FOLDER = r"\\Mp1do4ce0373ndz\C\WeeklyRawFile"
DST_FOLDER = r"\\Mp1do4ce0373ndz\d\Reporting\Raw\Inventory"

# 要“格式化处理”的文件名
FILES = [
    "CN MB52 Raw.xls",
    "MY MB52 Raw.xls",
    "US MB52 Raw.xls",
    "SG MB52 Raw.xls",
]

# 仅复制、不做任何处理的文件
COPY_ONLY = [
    "MB5TD Raw.xls",
]

# Excel 常量
xlCellTypeConstants = 2
xlCellTypeFormulas = -4123
xlCalculationManual = -4135
xlCalculationAutomatic = -4105

# ---------- 文件复制 ----------
def copy_from_weekly_to_inventory() -> list[str]:
    os.makedirs(DST_FOLDER, exist_ok=True)
    copied = []
    for fname in (FILES + COPY_ONLY):
        src = os.path.join(SRC_FOLDER, fname)
        dst = os.path.join(DST_FOLDER, fname)
        if not os.path.exists(src):
            print(f"⚠ 源文件不存在（跳过）: {src} / Source file missing (skipped): {src}")
            continue
        try:
            shutil.copy2(src, dst)  # 覆盖
            print(f"📥 已复制: {fname} / Copied: {fname}")
            copied.append(dst)
        except Exception as e:
            print(f"❌ 复制失败: {fname} -> {e} / Copy failed: {fname} -> {e}")
    return copied

# ---------- 列格式化工具 ----------
def to_text_full_digits(ws, col_letter: str):
    """
    将指定列现有的数值转成 '纯文本完整数字'：
      - 先设为 "0" 获取完整数字 .Text（避免 3.25E+11）
      - 再设为 "@"，把 .Text 写回（带前置 '）
    """
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 1:
        return
    rng = ws.Range(f"{col_letter}1:{col_letter}{last_row}")
    rng.NumberFormat = "0"

    # 常量
    with suppress(Exception):
        for c in rng.SpecialCells(xlCellTypeConstants):
            t = c.Text
            if t:
                c.NumberFormat = "@"
                c.Value = "'" + t

    # 公式
    with suppress(Exception):
        for c in rng.SpecialCells(xlCellTypeFormulas):
            t = c.Text
            if t:
                c.NumberFormat = "@"
                c.Value = "'" + t

    rng.NumberFormat = "@"

def set_col_text(ws, col_letter: str):
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    if last_row < 1:
        return
    ws.Range(f"{col_letter}1:{col_letter}{last_row}").NumberFormat = "@"

# ---------- 新增：确保最左侧空列（A 列）不会被 Excel 吃掉 ----------
def ensure_leading_blank_column(ws):
    """
    保证左侧有一列“空列”并在保存/重开后不消失：
    - 如果 A1 看起来是标题（如 Material），说明空列已被吃掉 -> 插入一列 A
    - 在 A1/A2 写入不可见占位符（NBSP，chr(160)）并设为文本，让 Excel 认为该列“已用”
    - 不隐藏列，也不改变列宽（保持你当前视觉效果）
    """
    a1 = str(ws.Cells(1, 1).Value or "").strip().lower()
    # 这里用最常见的标题判断；如果你的文件标题不是 Material，可按需扩展集合
    if a1 in {"material"}:
        ws.Columns("A").Insert()  # 原A整体右移

    # 写入占位符，防止保存时被裁掉
    for r in (1, 2):
        with suppress(Exception):
            cell = ws.Cells(r, 1)
            cell.NumberFormat = "@"
            cell.Value = chr(160)  # NBSP 不间断空格

    # 触发 UsedRange 更新（可选）
    _ = ws.UsedRange

# ---------- 单文件处理（含重试，防止占用） ----------
def process_one_excel(excel_app, path: str, is_sg: bool, open_retries: int = 3, open_sleep: float = 2.0):
    # 打开参数：不更新外链、不提示
    for attempt in range(1, open_retries + 1):
        try:
            wb = excel_app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False, Notify=False)
            break
        except Exception as e:
            if attempt >= open_retries:
                raise
            print(f"⏳ 打开失败，可能被占用：{os.path.basename(path)} -> {e}，{open_sleep}s 后重试（{attempt}/{open_retries-1}） / "
                  f"Open failed (maybe in use): {os.path.basename(path)} -> {e}, retry in {open_sleep}s ({attempt}/{open_retries-1})")
            time.sleep(open_sleep)

    try:
        ws = wb.ActiveSheet  # 如需特定表，可改为 wb.Worksheets("Sheet1")

        if is_sg:
            # 1) 删除整列 N（空列）：若不存在则忽略
            with suppress(Exception):
                ws.Columns("N").Delete()

            # 2) C 列完整数字转文本
            to_text_full_digits(ws, "C")

            # 3) N 列（原 O → Lot ID）设为文本
            set_col_text(ws, "N")
        else:
            # 非 SG：C/N
            to_text_full_digits(ws, "C")
            set_col_text(ws, "N")

        wb.Save()  # 保存一次即可
        print(f"✅ 已格式化: {os.path.basename(path)} / Formatted: {os.path.basename(path)}")
    finally:
        # 确保关闭以释放文件句柄
        with suppress(Exception):
            wb.Close(SaveChanges=True)

# ---------- 主流程 ----------
def main():
    print("==== Step 1: 复制文件到 Inventory 目录（覆盖） / Copy files to Inventory folder (overwrite) ====")
    copied_paths = copy_from_weekly_to_inventory()
    if not copied_paths:
        print("⚠ 没有可复制/可处理的文件，结束。 / No files to copy/process. Exiting.")
        return

    print("\n==== Step 2: Excel 后台格式化（仅对 MB52 文件） / Excel background formatting (MB52 only) ====")
    excel = win32.Dispatch("Excel.Application")
    # 完全后台
    excel.DisplayAlerts = False
    # 提速与稳定性
    with suppress(Exception):
        excel.ScreenUpdating = False
    with suppress(Exception):
        excel.EnableEvents = False
    try:
        prev_calc = None
        with suppress(Exception):
            prev_calc = excel.Calculation
            excel.Calculation = xlCalculationManual

        for fname in FILES:
            dst_path = os.path.join(DST_FOLDER, fname)
            if not os.path.exists(dst_path):
                print(f"⚠ 目标缺失（跳过格式化）: {fname} / Target missing (skip formatting): {fname}")
                continue
            process_one_excel(excel, dst_path, is_sg=("SG MB52" in fname))

        # ---------- 新增：MB5TD 的 A/B/U 处理 ----------
        mb5td = os.path.join(DST_FOLDER, "MB5TD Raw.xls")
        if os.path.exists(mb5td):
            print("\n—— 处理 MB5TD：保留 A 列空列，并将 B/U 列设为文本 —— / "
                  "Process MB5TD: keep A blank column and set B/U to text ——")
            for attempt in range(1, 4):
                try:
                    wb2 = excel.Workbooks.Open(mb5td, UpdateLinks=0, ReadOnly=False, Notify=False)
                    break
                except Exception as e:
                    if attempt >= 3:
                        raise
                    print(f"⏳ 打开失败（MB5TD）：{e}，2s 后重试（{attempt}/3） / "
                          f"Open failed (MB5TD): {e}, retry in 2s ({attempt}/3)")
                    time.sleep(2.0)
            try:
                ws2 = wb2.ActiveSheet
                # 1) 确保 A 列（最左侧空列）不会被 Excel 自动裁掉
                ensure_leading_blank_column(ws2)
                # 2) 将 B 与 U 列设为文本格式（按你的要求）
                set_col_text(ws2, "B")
                set_col_text(ws2, "U")
                wb2.Save()
                print("✅ MB5TD：A 列已保留；B/U 列已设为文本。 / MB5TD: A kept; B/U set to text.")
            finally:
                with suppress(Exception):
                    wb2.Close(SaveChanges=True)
        else:
            print("ℹ 未找到 MB5TD Raw.xls，跳过该文件的 A/B/U 处理。 / "
                  "MB5TD Raw.xls not found; skipping A/B/U handling.")
        # ---------- 新增结束 ----------

    finally:
        # 还原环境
        with suppress(Exception):
            excel.Calculation = prev_calc if prev_calc is not None else xlCalculationAutomatic
        with suppress(Exception):
            excel.EnableEvents = True
        with suppress(Exception):
            excel.ScreenUpdating = True
        with suppress(Exception):
            excel.Quit()

    print("\n🎉 全部完成。 / All done.")

if __name__ == "__main__":
    main()
