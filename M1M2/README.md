# M1M2 说明

本目录为 M1/M2 相关自动化脚本。每个脚本对应一段固定流程。

## Step1.py
重建 Excel 文件，降低 gencache/typelib 异常影响。

流程：
1) 定位固定文件清单（6 个文件）。
2) 按月备份（同月同版本跳过）。
3) 清理损坏的 typelibs 缓存。
4) 用 dynamic.Dispatch 启动 Excel。
5) 打开 → SaveAs 临时 → 关闭 → 原子替换；失败则降级 Save()。
6) 单文件失败不影响其他文件。

## Step2.py
CSV 转 XLSX 并拷贝到网络盘，最后归档原 CSV。

流程：
1) 读取 CSV（多编码兜底）。
2) 生成本地 XLSX（文本列整列设为文本）。
3) 复制到网络盘目标路径。
4) 归档原 CSV 到 _processed 目录。

## Step4.py
从 VL06O 源文件复制数据到模板。

流程：
1) 在 Archive 找最新 VL06O*.xlsx。
2) 打开源/目标工作簿。
3) 清空目标数据区。
4) 按列映射写入目标。
5) L 列填充公式 =TRIM(A2)。
6) 保存目标文件。

## Step5.py
更新 Product_list New.xlsx 的表头与数据。

流程：
1) 找 PR1 reports 下最新 Product_List*.xlsx。
2) 读取旧文件表头（保留重复列名）。
3) 读取新文件数据（跳过首行）。
4) 校验列数一致后套用旧表头。
5) 保存为 Product_list New.xlsx。

## Step8/Step8_1.py
处理 Scrap/Machinery 模板的空值清理，并串联后续步骤。

流程：
1) 在 Archive 找最新 Scrap/Machinery 文件。
2) 复制到 Customs 目录并按固定名保存。
3) 清空数据区（保留表头与第 1 列）。
4) 保存并依次调用 Step8_2、Step8_3。

## Step8/Step8_2.py
将 Raw 文件的数据复制到模板，并补充月份标识。

流程：
1) 找最新含关键字的 Raw 文件。
2) 找对应模板文件（.xlsm/.xlsx）。
3) 按规则列位移复制数据。
4) 对有数据的行填充第 1 列为上月 YYYYMM。
5) 保存模板。

## Step8/Step8_3.py
准备 M1/M2 文件并打开用于手工对拷。

流程：
1) 找 Raw 中最新 M1 文件，复制到 Customs 固定名。
2) 找 Archive 中最新 M2 文件，复制到 Customs 固定名并打开源+目标。
3) 找 Raw 中最新 M2 文件并打开（供人工对比/对拷）。

## check_excel_blank_rows.py
修复 Excel 的空白行/UsedRange 异常。

流程：
1) 在目录中模糊匹配最新文件（按基础名）。
2) 统一备份到时间戳目录。
3) 用 openpyxl 计算真实数据范围。
4) 使用 Excel COM 重建工作表并保留类型/列宽。
5) 验证 Ctrl+End 最后单元格位置并输出日志。

---

# M1M2 Notes (EN)

This folder contains M1/M2 automation scripts. Each script corresponds to a fixed workflow.

## Step1.py
Rebuilds Excel files to avoid gencache/typelib errors.

Steps:
1) Locate the fixed file list (6 files).
2) Monthly backup (skip if same version already backed up this month).
3) Purge broken typelibs cache.
4) Launch Excel via dynamic.Dispatch.
5) Open → SaveAs temp → close → atomic replace; fallback to Save() on failure.
6) Per-file failures do not stop others.

## Step2.py
Convert CSV to XLSX, copy to network share, then archive the CSV.

Steps:
1) Read CSV (multi-encoding fallback).
2) Write local XLSX (text columns formatted as text).
3) Copy to network target.
4) Move original CSV to _processed archive.

## Step4.py
Copy data from the latest VL06O source into a template.

Steps:
1) Find latest VL06O*.xlsx in Archive.
2) Open source and target workbooks.
3) Clear target data region.
4) Write mapped columns to target.
5) Fill column L with =TRIM(A2).
6) Save target file.

## Step5.py
Update Product_list New.xlsx headers and data.

Steps:
1) Find latest Product_List*.xlsx in PR1 reports.
2) Read old headers (preserve duplicate names).
3) Read new data (skip first row).
4) Validate column count and apply old headers.
5) Save Product_list New.xlsx.

## Step8/Step8_1.py
Clear Scrap/Machinery templates and chain subsequent steps.

Steps:
1) Find latest Scrap/Machinery files in Archive.
2) Copy to Customs with fixed names.
3) Clear data area (keep headers and column 1).
4) Save and call Step8_2, Step8_3.

## Step8/Step8_2.py
Copy data from Raw files into templates and add month tags.

Steps:
1) Find latest Raw file containing keyword.
2) Find matching template (.xlsm/.xlsx).
3) Copy with column shift rules.
4) Fill column 1 with previous month YYYYMM on rows with data.
5) Save template.

## Step8/Step8_3.py
Prepare M1/M2 files and open for manual copy/compare.

Steps:
1) Find latest M1 in Raw and copy to Customs fixed name.
2) Find latest M2 in Archive, copy to Customs fixed name, open source + target.
3) Find latest M2 in Raw and open for manual compare.

## check_excel_blank_rows.py
Fix Excel blank rows/UsedRange issues.

Steps:
1) Fuzzy-match latest files by base name.
2) Backup to timestamped directory.
3) Compute true ranges with openpyxl.
4) Rebuild sheets via Excel COM preserving types/widths.
5) Verify Ctrl+End last cell and log results.
