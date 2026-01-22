# -*- coding: utf-8 -*-
import os, shutil, datetime as dt
from glob import glob

# ================== 配置 ==================
ROOT = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\SAP\BW Helios\GP1 and Static"
ARCHIVED = os.path.join(ROOT, "Archived")

OUTPUT_PREFIXES = {
    "static": "LAMPs and OTH Static",
    "git":    "LAMPs and OTH GIT",
}
# =========================================

def latest_file_by_keyword(keyword: str):
    """
    在 Archived 下找到最新文件：
      - 优先包含 keyword 的文件（忽略大小写）
      - 没有则用最新的任意 .xlsx
    """
    candidates = [f for f in glob(os.path.join(ARCHIVED, "*.xlsx")) if os.path.isfile(f)]
    if not candidates:
        raise FileNotFoundError("Archived 中未找到任何 .xlsx。 / No .xlsx found in Archived.")

    if keyword:
        matches = [f for f in candidates if keyword.lower() in os.path.basename(f).lower()]
        if matches:
            matches.sort(key=os.path.getmtime, reverse=True)
            return matches[0]

    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]

def copy_and_rename(keyword: str, prefix: str):
    src = latest_file_by_keyword(keyword)
    today_str = dt.date.today().strftime("%Y%m%d")
    dest_name = f"{prefix} - {today_str}.xlsx"
    dest_path = os.path.join(ROOT, dest_name)

    shutil.copy2(src, dest_path)
    print(f"✅ 已复制并重命名：\n  来源: {src}\n  目标: {dest_path}\n  / Copied and renamed:")

def main():
    for kw, prefix in OUTPUT_PREFIXES.items():
        try:
            copy_and_rename(kw, prefix)
        except Exception as e:
            print(f"⚠ {prefix} 未能处理: {e} / Failed to process {prefix}: {e}")

if __name__ == "__main__":
    main()
