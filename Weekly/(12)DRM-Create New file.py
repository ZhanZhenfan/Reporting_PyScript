# -*- coding: utf-8 -*-
"""
DRM 周报自动化（修正版）- 带日志：
1) 在目录中查找最新周报文件，根据周次+1复制为新文件；
2) 打开后先清除所有筛选；
3) 刷新匹配 TARGET_CONN_KEYS 的连接；
4) 在最后一个工作表（应名为 Details）：
   - 读取 AD2 的周次信息（仅用于日志）；
   - 用 R1C1 公式写入 BL 列（必要时逐行写入）；
   - 填充 BM、BO 列为其下拉列表首项。
5) 保存并退出。
"""

import os
import re
import shutil
import datetime as dt
import time
import pythoncom
import win32com.client as win32
from win32com.client import constants

# =============== 日志辅助 ===============
def log(level, msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")

def info(msg): log("INFO", msg)
def ok(msg):   log("OK", msg)
def warn(msg): log("WARN", msg)
def err(msg):  log("ERROR", msg)

# 配置
ROOT_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\DRM"
FILENAME_REGEX = re.compile(r"^DRM Report W(\d{1,2})'(\d{2})\.xlsx$", re.IGNORECASE)
TARGET_CONN_KEYS = ["VW_DRMMeasurement_CY"]  # 只刷新名称/连接串包含这些关键词的连接
READ_AD_CELL = "AD2"  # 日志用

def safe_excel_app():
    """创建静默 Excel 实例并优化设置"""
    info("创建 Excel 实例（静默/禁用提示/优化性能）…… / Creating Excel instance (silent/disable prompts/optimize performance)...")
    excel = win32.DispatchEx("Excel.Application")
    time.sleep(1)
    for attr, val in [("Visible", False), ("DisplayAlerts", False)]:
        try:
            setattr(excel, attr, val)
        except Exception:
            warn(f"设置 Excel.{attr} 失败，继续。 / Failed to set Excel.{attr}, continuing.")
    try:
        excel.Calculation = -4135  # Manual
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        ok("Excel 实例创建完成。 / Excel instance created.")
    except Exception:
        warn("部分 Excel 优化属性设置失败。 / Some Excel optimization settings failed.")
    return excel

def find_latest_matching_file(folder: str):
    info(f"在目录中查找最新周报：{folder} / Searching latest report in: {folder}")
    cands = [os.path.join(folder, f) for f in os.listdir(folder) if FILENAME_REGEX.match(f)]
    if not cands:
        raise FileNotFoundError("未找到任何 DRM Report W##'YY.xlsx 文件。 / No DRM Report W##'YY.xlsx files found.")
    cands.sort(key=os.path.getmtime, reverse=True)
    ok(f"找到最新文件：{os.path.basename(cands[0])} / Latest file found: {os.path.basename(cands[0])}")
    return cands[0]

def parse_week_from_filename(path: str):
    name = os.path.basename(path)
    m = FILENAME_REGEX.match(name)
    week = int(m.group(1))
    yy = int(m.group(2))
    info(f"解析文件名周次/年份：W{week}'{yy:02d} / Parsed week/year: W{week}'{yy:02d}")
    return week, 2000 + yy

def next_week_token_from_filename(latest_path: str) -> str:
    info("计算下一周周次标记（基于 ISO 周）…… / Calculating next week token (ISO)...")
    week, year_full = parse_week_from_filename(latest_path)
    base_monday = dt.date.fromisocalendar(year_full, week, 1)
    next_monday = base_monday + dt.timedelta(days=7)
    iso = next_monday.isocalendar()
    token = f"W{iso.week:02d}'{iso.year % 100:02d}"
    ok(f"下一周标记为：{token} / Next week token: {token}")
    return token

def get_details_sheet(wb):
    """锁定最后一个工作表（其名称应为 Details），否则搜索名为 Details"""
    info("定位 Details 工作表（若最后一张非 Details 则遍历查找）…… / Locate Details sheet (search if last isn't Details)...")
    last = wb.Sheets(wb.Sheets.Count)
    last_name = getattr(last, "Name", "")
    if str(last_name).strip().lower() == "details":
        ok("已定位最后一张为 Details。 / Last sheet is Details.")
        return last
    for s in wb.Sheets:
        if str(getattr(s, "Name", "")).strip().lower() == "details":
            ok("已找到名为 Details 的工作表。 / Found sheet named Details.")
            return s
    warn("未找到名为 Details 的工作表，使用最后一张。 / Details sheet not found; using last sheet.")
    return last

def clear_all_filters(wb):
    """清除所有工作表的筛选"""
    info("清除所有工作表的筛选…… / Clearing filters on all sheets...")
    cleared = 0
    for sh in wb.Sheets:
        try:
            if sh.AutoFilterMode:
                try:
                    sh.ShowAllData()
                except Exception:
                    pass
                sh.AutoFilterMode = False
                cleared += 1
        except Exception:
            continue
    ok(f"筛选清除完成：共处理 {cleared} 张工作表。 / Filters cleared: processed {cleared} sheets.")

def refresh_target_connections(wb, app, keys, timeout_sec=300):
    """刷新匹配 keys 的连接"""
    info(f"筛选并刷新匹配连接（关键词：{keys}）…… / Filter and refresh connections (keys: {keys})...")
    refreshed, targets = [], []
    for conn in wb.Connections:
        name = getattr(conn, "Name", "") or ""
        hay = name.upper()
        conn_str = ""
        try:
            if conn.Type == 1:
                conn_str = getattr(conn.OLEDBConnection, "Connection", "") or ""
            elif conn.Type == 2:
                conn_str = getattr(conn.ODBCConnection, "Connection", "") or ""
        except Exception:
            pass
        hay += " " + conn_str.upper()
        if any(k.upper() in hay for k in keys):
            targets.append((name, conn))
    info(f"目标连接数量：{len(targets)} / Target connections: {len(targets)}")
    for nm, conn in targets:
        info(f"刷新连接：{nm} / Refreshing connection: {nm}")
        try:
            try:
                if conn.Type == 1:
                    conn.OLEDBConnection.BackgroundQuery = False
                elif conn.Type == 2:
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass
            start = time.time()
            conn.Refresh()
            while True:
                pythoncom.PumpWaitingMessages()
                try:
                    app.CalculateUntilAsyncQueriesDone()
                except Exception:
                    pass
                time.sleep(0.5)
                if time.time() - start > timeout_sec:
                    raise TimeoutError(f"刷新连接超时（>{timeout_sec}s）：{nm} / Refresh timeout (>{timeout_sec}s): {nm}")
                try:
                    app.CalculateFullRebuild()
                except Exception:
                    pass
                time.sleep(1.0)
                break
            ok(f"刷新完成：{nm} / Refresh complete: {nm}")
            refreshed.append(nm)
        except Exception as e:
            err(f"刷新失败：{nm} | {e} / Refresh failed: {nm} | {e}")
    if not refreshed:
        warn("无连接被刷新（可能未匹配关键词或全部失败）。 / No connections refreshed (no match or all failed).")
    return refreshed

def first_dropdown_value(cell):
    """获取数据验证的首项（下拉列表）"""
    try:
        val = cell.Validation.Formula1
        if not val:
            return "<DEFAULT>"
        if val.startswith("="):
            rng = cell.Application.Evaluate(val)
            return rng.Cells(1, 1).Value
        else:
            return val.split(",")[0]
    except Exception:
        return "<DEFAULT>"

def set_column_formula_robust(ws, col_letter: str, start_row: int, end_row: int):
    """
    使用 R1C1 形式为某一列写入公式，并对复杂情况下逐行写入。
    公式逻辑等价于：
      IFERROR(VLOOKUP(G,Exclusion!A:C,2,0),
        IFERROR(VLOOKUP(L,Exclusion!G:H,2,0),
          IFERROR(VLOOKUP(F,Exclusion!K:L,2,0),
            IF(ISERROR(SEARCH(\"Malaysia\",E,1)), \"\", \"X\"))))
    """
    if end_row < start_row:
        warn(f"{col_letter} 列写公式：数据区为空（{start_row}>{end_row}），跳过。 / "
             f"Column {col_letter} formula: no data ({start_row}>{end_row}), skipping.")
        return

    info(f"为 {col_letter} 列写入公式（R1C1），行区间：{start_row}-{end_row} …… / "
         f"Writing R1C1 formulas to {col_letter}, rows {start_row}-{end_row}...")

    # 解除工作表保护（如果无密码）
    try:
        if ws.ProtectContents:
            try:
                ws.Unprotect("")
            except Exception:
                ws.Unprotect()
            info("已尝试解除工作表保护。 / Attempted to unprotect worksheet.")
    except Exception:
        pass

    # 获取列号
    try:
        col_num = ws.Range(col_letter + "1").Column
    except Exception:
        col_num = 0
        for c in col_letter.upper():
            col_num = col_num * 26 + (ord(c) - ord('A') + 1)

    # R1C1 版本公式
    formula_r1c1 = (
        '=IFERROR(VLOOKUP(RC7,Exclusion!C1:C3,2,0),'
        'IFERROR(VLOOKUP(RC12,Exclusion!C7:C8,2,0),'
        'IFERROR(VLOOKUP(RC6,Exclusion!C11:C12,2,0),'
        'IF(ISERROR(SEARCH("Malaysia",RC5,1)),"","X"))))'
    )

    top_cell = ws.Cells(start_row, col_num)
    dest_range = ws.Range(top_cell, ws.Cells(end_row, col_num))
    # 尝试批量写入 FormulaR1C1
    try:
        top_cell.FormulaR1C1 = formula_r1c1
        top_cell.AutoFill(Destination=dest_range, Type=constants.xlFillDefault)
        ok(f"{col_letter} 列公式批量填充完成。 / Bulk formula fill complete for {col_letter}.")
        return
    except Exception:
        warn(f"{col_letter} 列批量填充失败，改为逐行写入。 / "
             f"Bulk fill failed for {col_letter}, switching to row-by-row.")

    # 批量失败则逐行写入 FormulaR1C1
    try:
        for r in range(start_row, end_row + 1):
            ws.Cells(r, col_num).FormulaR1C1 = formula_r1c1
        ok(f"{col_letter} 列逐行写入完成。 / Row-by-row fill complete for {col_letter}.")
        return
    except Exception:
        warn(f"{col_letter} 列逐行写入失败，改用 A1 公式降级尝试。 / "
             f"Row-by-row failed for {col_letter}, fallback to A1 formulas.")

    # 最后尝试普通 Formula（逗号和分号两种），如仍失败则抛出异常
    formula_en = (
        '=IFERROR(VLOOKUP(G{row},Exclusion!A:C,2,0),'
        'IFERROR(VLOOKUP(L{row},Exclusion!G:H,2,0),'
        'IFERROR(VLOOKUP(F{row},Exclusion!K:L,2,0),'
        'IF(ISERROR(SEARCH("Malaysia",E{row},1)),"","X"))))'
    )
    separators = [",", ";"]
    for sep in separators:
        try:
            top_formula = formula_en.format(row=start_row).replace(",", sep)
            top_cell.Formula = top_formula
            for r in range(start_row, end_row + 1):
                formula_row = formula_en.format(row=r).replace(",", sep)
                ws.Cells(r, col_num).Formula = formula_row
            ok(f"{col_letter} 列使用普通公式写入完成（分隔符：{sep}）。 / "
               f"Formula fill complete for {col_letter} (sep: {sep}).")
            return
        except Exception:
            try:
                top_formula = formula_en.format(row=start_row).replace(",", sep)
                top_cell.FormulaLocal = top_formula
                for r in range(start_row, end_row + 1):
                    formula_row = formula_en.format(row=r).replace(",", sep)
                    ws.Cells(r, col_num).FormulaLocal = formula_row
                ok(f"{col_letter} 列使用 FormulaLocal 写入完成（分隔符：{sep}）。 / "
                   f"FormulaLocal fill complete for {col_letter} (sep: {sep}).")
                return
            except Exception:
                warn(f"{col_letter} 列普通公式尝试失败（分隔符：{sep}）。 / "
                     f"Formula attempt failed for {col_letter} (sep: {sep}).")

    raise RuntimeError(f"无法将公式写入 {col_letter} 列 / Failed to write formulas to column {col_letter}")

def refresh_and_update(file_path):
    info(f"打开文件：{file_path} / Opening file: {file_path}")
    excel = safe_excel_app()
    wb = excel.Workbooks.Open(file_path)
    try:
        clear_all_filters(wb)

        ws = get_details_sheet(wb)

        refreshed = refresh_target_connections(wb, excel, TARGET_CONN_KEYS)
        info(f"本次已刷新连接：{refreshed if refreshed else '（无）'} / "
             f"Refreshed connections this run: {refreshed if refreshed else '(none)'}")

        try:
            ad_value = ws.Range(READ_AD_CELL).Value
            info(f"读取 {READ_AD_CELL}：{ad_value} / Read {READ_AD_CELL}: {ad_value}")
        except Exception as e:
            warn(f"读取 {READ_AD_CELL} 失败：{e} / Failed to read {READ_AD_CELL}: {e}")

        # 末行探测
        last_row = ws.Cells(ws.Rows.Count, "G").End(-4162).Row
        if last_row < 2:
            last_row = 2
        info(f"数据区末行：{last_row} / Last data row: {last_row}")

        # 读取 BM/BO 下拉首项
        bm_first = first_dropdown_value(ws.Range("BM2"))
        bo_first = first_dropdown_value(ws.Range("BO2"))
        info(f"BM 下拉首项：{bm_first} | BO 下拉首项：{bo_first} / "
             f"BM first dropdown: {bm_first} | BO first dropdown: {bo_first}")

        # 写 BL 公式（必要时逐行）
        set_column_formula_robust(ws, "BL", 2, last_row)

        # 填充 BM、BO 列
        info("填充 BM、BO 列为下拉首项…… / Filling BM/BO with first dropdown values...")
        ws.Range(f"BM2:BM{last_row}").Value = bm_first
        ws.Range(f"BO2:BO{last_row}").Value = bo_first
        ok("BM、BO 填充完成。 / BM/BO fill complete.")

        info("保存工作簿…… / Saving workbook...")
        wb.Save()
        ok("保存完成。 / Save complete.")
    finally:
        info("关闭工作簿并退出 Excel…… / Closing workbook and exiting Excel...")
        try:
            wb.Close(SaveChanges=True)
        except Exception as e:
            warn(f"关闭工作簿异常：{e} / Error closing workbook: {e}")
        try:
            excel.Quit()
        except Exception as e:
            warn(f"退出 Excel 异常：{e} / Error quitting Excel: {e}")
        ok("Excel 已退出。 / Excel exited.")

def main():
    info(f"工作目录：{ROOT_DIR} / Working directory: {ROOT_DIR}")
    latest_src = find_latest_matching_file(ROOT_DIR)
    next_week_token = next_week_token_from_filename(latest_src)
    target_name = f"DRM Report {next_week_token}.xlsx"
    target_path = os.path.join(ROOT_DIR, target_name)

    if not os.path.exists(target_path):
        info(f"复制最新文件为新周报：{os.path.basename(latest_src)} -> {target_name} / "
             f"Copy latest file as new report: {os.path.basename(latest_src)} -> {target_name}")
        shutil.copy2(latest_src, target_path)
        ok(f"已复制：{target_path} / Copied: {target_path}")
    else:
        info(f"目标文件已存在，直接使用：{target_path} / Target file exists, using: {target_path}")

    refresh_and_update(target_path)
    ok("流程结束。 / Process finished.")

if __name__ == "__main__":
    main()
