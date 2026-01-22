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
