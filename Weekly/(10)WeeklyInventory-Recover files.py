import os
import shutil
import re
from datetime import datetime

# 📁 源/目标路径
src_folder = r"\\Mp1do4ce0373ndz\d\Reporting\Raw\Inventory\Archive"
dst_folder = r"\\Mp1do4ce0373ndz\d\Reporting\Raw\Inventory"

# 这些文件保留 .xlsx，其余改为 .xls
keep_xlsx_prefixes = {"KKAQ_1", "KKAQ_2", "KKAQ_3"}

# ✅ 正则：提取前缀和时间戳（年份4位：20xx）
# 例：'SG MB52 Raw_2025-09-17-092020.xlsx' → prefix='SG MB52 Raw', ts='2025-09-17-092020'
pattern = re.compile(r"^(.*?)(_20\d{2}-\d{2}-\d{2}-\d{6})\.xlsx$", re.IGNORECASE)

# 用于存储每组 prefix 下最新的文件 (prefix → (datetime, filename))
latest_files = {}

# 确保目标目录存在
os.makedirs(dst_folder, exist_ok=True)

print(f"[INFO] Scan: {src_folder}")
for filename in os.listdir(src_folder):
    if not filename.lower().endswith(".xlsx"):
        # 如需查看被跳过的非xlsx：取消下一行注释
        # print("  skip ext:", filename)
        continue

    m = pattern.match(filename)
    if not m:
        # 如需查看未匹配命名：取消下一行注释
        # print("  no match :", filename)
        continue

    prefix, ts_str = m.groups()        # e.g. ('SG MB52 Raw', '_2025-09-17-092020')
    ts_clean = ts_str.lstrip("_")      # '2025-09-17-092020'

    try:
        ts_dt = datetime.strptime(ts_clean, "%Y-%m-%d-%H%M%S")
    except ValueError:
        # print("  bad ts  :", filename)
        continue

    if (prefix not in latest_files) or (ts_dt > latest_files[prefix][0]):
        latest_files[prefix] = (ts_dt, filename)

if not latest_files:
    print("[WARN] 没有匹配到任何带时间戳的 .xlsx 文件。请检查文件命名是否为 *_YYYY-MM-DD-HHMMSS.xlsx / "
          "No timestamped .xlsx files matched. Check naming like *_YYYY-MM-DD-HHMMSS.xlsx")
else:
    print(f"[INFO] Groups found: {len(latest_files)}")

# 复制并重命名（覆盖旧文件）
for prefix, (ts_dt, filename) in sorted(latest_files.items()):
    src_path = os.path.join(src_folder, filename)

    # 决定目标扩展名
    new_ext = ".xlsx" if prefix in keep_xlsx_prefixes else ".xls"
    new_filename = prefix + new_ext
    dst_path = os.path.join(dst_folder, new_filename)

    try:
        shutil.copy2(src_path, dst_path)  # ✅ 改为复制，保留时间戳元数据
        print(f"✅ Copy: {filename}  →  {new_filename}")
    except Exception as e:
        print(f"❌ Copy failed: {filename}  -> {e}")

print("[DONE] 最新文件已复制到目标目录并按规则重命名。 / Latest files copied and renamed in destination.")
