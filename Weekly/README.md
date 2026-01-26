# Weekly 说明

本目录为周度/定期任务脚本，按业务场景拆分。

## (1)MRP_Weekly_Waterfall.py
流程：
1) 下载两份 ZMRP_WATERFALL_Run 附件到本地。
2) 选最新两份文件，按大小区分大/小。
3) 拼接数据（小文件去首行）。
4) 生成周一命名文件并保存。
5) 复制到共享盘。
6) 触发 SQL Agent Job。
7) 打开宏文件。

## (2)MRP_Waterfall_Monthly.py
流程：
1) 根据 INPUT_MODE 从邮箱或文件夹取最新源文件。
2) 清洗数据（物料号文本化、数值列处理）。
3) 等待共享盘无占位文件。
4) 复制清洗结果到共享盘。
5) 触发 SQL Agent Job（可选）。

## (3)Rel_SNOP.py
流程：
1) 创建锁文件，避免多实例并发。
2) 按源文件 mtime 备份固定文件（幂等）。
3) 打开 Excel 刷新并保存。
4) 触发 SQL Agent Job。

## (4&5)BW Static+GIT-Move file only.py
流程：
1) 在 Archived 里找最新文件（优先包含关键字）。
2) 复制并按日期重命名到目标目录。

## (7)InfoRecord.py
流程：
1) 找到最新子目录及最新 PR1 Info Record 文件。
2) 检查目标目录阻塞文件并等待。
3) L 列去前导 0 并设为文本。
4) 在 Q 列插入空列并写表头。
5) 备份旧文件后覆盖保存。
6) 触发 SQL Agent Job。

## (8)Supplier.py
流程：
1) 计算上周/本周周标。
2) 找到上周文件并复制为本周命名（若不存在则兜底）。
3) 打开并刷新连接。
4) 批量写 Reason Code、清空 BT。
5) 调整表范围与 UsedRange，保存关闭。

## (9)Subcon.py
流程：
1) 计算本周周标。
2) 分别复制 China/NonChina 最新文件为本周命名。
3) 打开并刷新连接。
4) 填充 BH 列 Reason Code。
5) 保存并完成。

## (10)WeeklyInventory-ChangetoMon.py
流程：
1) 在 Inventory 目录匹配目标文件。
2) 将最后一列统一改为本周周一日期并保留原格式。

## (10)WeeklyInventory-Formatting.py
流程：
1) 从 Weekly 目录复制文件到 Inventory 目录。
2) 后台打开 Excel 格式化列（C/N 等）。
3) MB5TD 额外处理 A/B/U 列。

## (10)WeeklyInventory-Formatting Monthend.py
流程：
1) 复制月末所需文件到 Inventory 目录。
2) 按列执行日期格式化与文本前缀替换。
3) MB5TD 进行 A/B/U/L/R/S 列处理。

## (10)WeeklyInventory-KKAQ.py
流程：
1) 通过 Graph 下载 KKAQ 附件。
2) 选择最新文件并复制到 Inventory 目录。

## (10)WeeklyInventory-Recover files.py
流程：
1) 扫描 Archive 中带时间戳的文件。
2) 每个前缀取最新版本。
3) 复制并重命名到目标目录。

## (12)DRM-Create New file.py
流程：
1) 找最新周报并复制为新周文件名。
2) 清除全部筛选。
3) 刷新匹配连接。
4) 读取 Details 表周次信息。
5) 写入 BL 公式并填充 BM/BO。
6) 保存退出。

## (12)DRM-Push Data.py
流程：
1) 找最新 DRM Report 并复制到目标目录。
2) 重命名工作表为 Sheet1。
3) 在 O 列插入 Delivery Num 列（表内/整列）。
4) 触发 SQL Agent Job。

## (14)O2FCST.py
流程：
1) 下载最新 FCST 附件（加密）。
2) 在 Archived 选最新有效模板。
3) 复制模板并命名为当周周一。
4) 从源表提取数据并写入目标第 2 个表。
5) 刷新指定连接（Query - Table1）。
6) 安全保存并发布到目标目录。
7) 触发 SQL Agent Job（可选）。

## (15)ExportInventoryReport.py
流程：
1) 自动找到最新 CSV。
2) 转换日期列、保留文本列。
3) 写出 Excel 到共享目录。
4) 复制到第二个共享目录。

## (16)SeleneRefined.py
流程：
1) 备份现有固定文件（按日期重命名）。
2) 复制最新源文件到固定名。
3) 运行 BAT。
4) 触发 SQL Agent Job。

## (17)REL SNOP updates.py
流程：
1) 找最新 REL Custom 文件。
2) 复制到目标目录覆盖。
3) 可选触发 SQL Agent Job。

---

# Weekly Notes (EN)

This folder contains weekly/periodic scripts organized by scenario.

## (1)MRP_Weekly_Waterfall.py
Steps:
1) Download two ZMRP_WATERFALL_Run attachments locally.
2) Pick the latest two, split by file size (big/small).
3) Merge data (drop first row of the small file).
4) Save as Monday-named file.
5) Copy to shared drive.
6) Trigger SQL Agent Job.
7) Open macro workbook.

## (2)MRP_Waterfall_Monthly.py
Steps:
1) Get latest source file from email or folder (INPUT_MODE).
2) Clean data (material as text, numeric columns normalized).
3) Wait for shared folder to clear blocking files.
4) Copy cleaned output to share.
5) Trigger SQL Agent Job (optional).

## (3)Rel_SNOP.py
Steps:
1) Create lock file to prevent concurrent runs.
2) Back up fixed file by mtime (idempotent).
3) Open Excel, refresh, save.
4) Trigger SQL Agent Job.

## (4&5)BW Static+GIT-Move file only.py
Steps:
1) Find latest file in Archived (prefer keyword match).
2) Copy and rename by date to target folder.

## (7)InfoRecord.py
Steps:
1) Find latest subfolder and latest PR1 Info Record file.
2) Wait for destination folder to clear blocking files.
3) Remove leading zeros in column L and set text format.
4) Insert blank column Q with header.
5) Back up old file, then overwrite.
6) Trigger SQL Agent Job.

## (8)Supplier.py
Steps:
1) Compute last/this week tokens.
2) Find last-week file and copy to this-week name (fallback if missing).
3) Open and refresh connections.
4) Fill Reason Code and clear BT.
5) Resize table/UsedRange, save and close.

## (9)Subcon.py
Steps:
1) Compute current week token.
2) Copy latest China/NonChina files to this-week names.
3) Open and refresh connections.
4) Fill BH column Reason Code.
5) Save and finish.

## (10)WeeklyInventory-ChangetoMon.py
Steps:
1) Match target files in Inventory folder.
2) Set last column to this Monday and keep original date format.

## (10)WeeklyInventory-Formatting.py
Steps:
1) Copy files from Weekly folder to Inventory folder.
2) Open Excel in background and format columns (C/N etc.).
3) Extra MB5TD handling for A/B/U.

## (10)WeeklyInventory-Formatting Monthend.py
Steps:
1) Copy month-end files to Inventory folder.
2) Apply date formatting and text prefix replacement by column.
3) MB5TD handling for A/B/U/L/R/S.

## (10)WeeklyInventory-KKAQ.py
Steps:
1) Download KKAQ attachments via Graph.
2) Pick latest and copy to Inventory folder.

## (10)WeeklyInventory-Recover files.py
Steps:
1) Scan Archive for timestamped files.
2) Keep latest per prefix.
3) Copy and rename into destination.

## (12)DRM-Create New file.py
Steps:
1) Find latest weekly report and copy to new week name.
2) Clear all filters.
3) Refresh matching connections.
4) Read week info from Details sheet.
5) Write BL formulas and fill BM/BO.
6) Save and exit.

## (12)DRM-Push Data.py
Steps:
1) Find latest DRM Report and copy to destination.
2) Rename worksheet to Sheet1.
3) Insert Delivery Num column at O (table or sheet).
4) Trigger SQL Agent Job.

## (14)O2FCST.py
Steps:
1) Download latest FCST attachment (encrypted).
2) Pick latest valid template from Archived.
3) Copy template and name by this Monday.
4) Extract source data and paste to target sheet 2.
5) Refresh Query - Table1.
6) Safe save and publish to target directory.
7) Trigger SQL Agent Job (optional).

## (15)ExportInventoryReport.py
Steps:
1) Find latest CSV automatically.
2) Convert date columns and keep text columns.
3) Write Excel to shared folder.
4) Copy to secondary shared folder.

## (16)SeleneRefined.py
Steps:
1) Back up existing fixed file (date-stamped name).
2) Copy latest source to fixed name.
3) Run BAT.
4) Trigger SQL Agent Job.

## (17)REL SNOP updates.py
Steps:
1) Find latest REL Custom file.
2) Copy to destination (overwrite).
3) Optionally trigger SQL Agent Job.
