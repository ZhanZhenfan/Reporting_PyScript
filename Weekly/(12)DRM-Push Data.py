# -*- coding: utf-8 -*-
r"""
DRM 月报处理（保留Excel验证/外部连接/表）
1) 在 \\mygbynbyn1msis2\SCM_Excellence\Weekly Report\DRM\ 下选择最新的 “DRM Report*.xlsx”
2) 复制到 \\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\DRM\ ，命名为 Monthly DRM File.xlsx
3) 用 Excel COM 修改：将 'details' 重命名为 'Sheet1'；在 O 列插入 'Delivery Num'（如有表则在表内新增列）
"""

import os
import shutil
from datetime import datetime
from typing import Optional

from Utils.sql_agent_tool import SqlAgentTool

# ===== 可配 =====
SRC_DIR  = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\DRM"
DEST_DIR = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\DRM"
DEST_FN  = "Monthly DRM File.xlsx"

TARGET_SHEET_NAME = "Sheet1"
DETAILS_NAME_CANDIDATES = {"details", "(details)", "detail", "details ", " details", "DETAILS"}
COL_INDEX_O = 15
COL_HEADER  = "Delivery Num"
# =================


def find_latest_drm_report(src_dir: str) -> Optional[str]:
    if not os.path.isdir(src_dir):
        print(f"⚠ 目录不存在：{src_dir} / Directory not found: {src_dir}")
        return None
    cands = []
    for name in os.listdir(src_dir):
        lower = name.lower()
        if lower.endswith(".xlsx") and lower.startswith("drm report") and not name.startswith("~$"):
            full = os.path.join(src_dir, name)
            if os.path.isfile(full):
                cands.append(full)
    if not cands:
        return None
    cands.sort(key=os.path.getmtime, reverse=True)
    return cands[0]


def process_with_excel_com(dest_path: str) -> None:
    import win32com.client as win32  # pip install pywin32

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        # UpdateLinks=0 避免弹窗与外部连接更新；只做结构性修改不会破坏连接/验证/表
        wb = excel.Workbooks.Open(dest_path, UpdateLinks=0, ReadOnly=False)
        try:
            # —— 1) 找到需要改名的 sheet —— #
            ws_target = None
            # 先精确匹配（不区分大小写/空格）
            names_norm = {s.Name.strip().lower(): s for s in wb.Worksheets}
            for key in list(names_norm.keys()):
                if key in {n.strip().lower() for n in DETAILS_NAME_CANDIDATES}:
                    ws_target = names_norm[key]
                    break
            # 若没找到，降级：包含 'details' 的也算
            if ws_target is None:
                for s in wb.Worksheets:
                    if "details" in s.Name.strip().lower():
                        ws_target = s
                        break
            # 再不行就用第一个
            if ws_target is None:
                ws_target = wb.Worksheets(1)

            # 若已存在 Sheet1 且不是同一个 sheet，则先把现有 Sheet1 改个名
            try:
                ws_existing_sheet1 = wb.Worksheets(TARGET_SHEET_NAME)
                if ws_existing_sheet1.Name != ws_target.Name:
                    ws_existing_sheet1.Name = TARGET_SHEET_NAME + "_old"
            except Exception:
                pass  # 没有 Sheet1 就跳过

            if ws_target.Name != TARGET_SHEET_NAME:
                ws_target.Name = TARGET_SHEET_NAME

            ws = wb.Worksheets(TARGET_SHEET_NAME)

            # —— 2) 在 O 列插入新列，列名 'Delivery Num' —— #
            # 若工作表包含表（ListObject），且 O 列位于表范围内部或紧随其后，则在表中新增列
            def _add_column_in_table_or_sheet(ws):
                try:
                    if ws.ListObjects.Count > 0:
                        tbl = ws.ListObjects(1)
                        start_col = tbl.Range.Column
                        end_col = start_col + tbl.ListColumns.Count - 1
                        # 若 O 列落在表范围内（或刚好在表的右侧一列），按表位置新增列，保持表结构/验证
                        if COL_INDEX_O >= start_col and COL_INDEX_O <= end_col + 1:
                            pos = COL_INDEX_O - start_col + 1
                            if pos < 1:
                                pos = 1
                            if pos > tbl.ListColumns.Count + 1:
                                pos = tbl.ListColumns.Count + 1
                            new_col = tbl.ListColumns.Add(Position=pos)
                            new_col.Name = COL_HEADER
                            return
                except Exception:
                    # 某些版本/保护状态下可能取表属性失败，退化为整列插入
                    pass

                # 不在表范围内或无表：按整列插入
                ws.Columns(COL_INDEX_O).Insert()
                ws.Cells(1, COL_INDEX_O).Value = COL_HEADER

            _add_column_in_table_or_sheet(ws)

            wb.Save()  # 用 Save 保留外部连接/验证/表
        finally:
            wb.Close(SaveChanges=True)
    finally:
        excel.DisplayAlerts = True
        excel.Quit()


def main():
    src_file = find_latest_drm_report(SRC_DIR)
    if not src_file:
        print(f"❌ 在 {SRC_DIR} 未找到 'DRM Report*.xlsx' / 'DRM Report*.xlsx' not found in {SRC_DIR}")
        return
    print(f"✅ 选定源文件：{src_file} / Selected source file: {src_file}")

    os.makedirs(DEST_DIR, exist_ok=True)
    dest_path = os.path.join(DEST_DIR, DEST_FN)

    shutil.copy2(src_file, dest_path)
    print(f"📤 已复制到：{dest_path} / Copied to: {dest_path}")

    process_with_excel_com(dest_path)

    print("\n🎉 完成： / Completed:")
    print("  源文件：", src_file, "/ Source:", src_file)
    print("  目标：  ", dest_path, "/ Destination:", dest_path)

    tool = SqlAgentTool(server="tcp:10.80.127.71,1433")

    result = tool.run_job(
        job_name="Lumileds BI - SC Shipped Backlog Report",  # 用完整精确名最稳妥
        archive_dir=r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\DRM\Archive",
        timeout=1800,
        poll_interval=3,
        fuzzy=False,  # 若你 later 拿到读 sysjobs 的权限，可改 True
        start_step="DRMMeasurement"
    )
    print(result)

if __name__ == "__main__":
    main()
