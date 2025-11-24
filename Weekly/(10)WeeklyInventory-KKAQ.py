# -*- coding: utf-8 -*-
import os
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from Utils.graph_mail_attachment_tool import GraphMailAttachmentTool

# ========= 配置 =========
TENANT_ID = "5c2be51b-4109-461d-a0e7-521be6237ce2"
CLIENT_ID = "09004044-1c60-48e5-b1eb-bb42b3892006"

# token 缓存放 Weekly 目录里（用绝对路径，避免每次都重新认证）
PROJECT_DIR = Path(__file__).resolve().parents[0]   # Weekly/
TOKEN_CACHE = (PROJECT_DIR / "graph_token_cache.json").as_posix()

# 附件关键词 -> 目标文件名 的映射
JOBS = {
    "US0X": ("KKAQ_1.xlsx",),
    "MY0X": ("KKAQ_2.xlsx",),
    "SG04": ("KKAQ_3.xlsx",),
}

# 从邮箱下载到的临时目录
TMP_DIR = r"\\mp1do4ce0373ndz\C\WeeklyRawFile\Download_From_Eamil"

# 复制/重命名到这个目录
DEST_DIR = r"\\Mp1do4ce0373ndz\d\Reporting\Raw\Inventory"

# 搜索窗口（天）
DAYS_BACK = 14
MAIL_FOLDER = "inbox"  # 不限制可设为 None
# =======================

# 匹配我们工具类保存的时间戳：..._YYYYMMDDThhmmss.xlsx
TS_RE = re.compile(r"_(\d{8}T\d{6})\.xlsx$", re.IGNORECASE)

def ensure_dir(p: str):
    Path(p).mkdir(parents=True, exist_ok=True)

def newest(paths):
    paths = [p for p in paths if p and Path(p).is_file()]
    return max(paths, key=lambda p: Path(p).stat().st_mtime) if paths else None

def extract_received_utc_from_name(path: str):
    """从保存的文件名末尾解析出接收时间（UTC），格式 YYYYMMDDThhmmss。"""
    m = TS_RE.search(os.path.basename(path))
    if not m:
        return None
    s = m.group(1)  # e.g. 20250921T221054
    try:
        return datetime.strptime(s, "%Y%m%dT%H%M%S")
    except Exception:
        return None

def main():
    ensure_dir(TMP_DIR)
    ensure_dir(DEST_DIR)

    tool = GraphMailAttachmentTool(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        token_cache=TOKEN_CACHE,  # 绝对路径，避免重复认证
    )

    for keyword, (target_name,) in JOBS.items():
        contains = f"KKAQ_{keyword}_"
        print(f"\n=== 拉取 {contains}* 最新一份 ===")

        paths = tool.download_latest_attachments(
            contains=contains,
            ext=".xlsx",
            need_count=1,
            days_back=DAYS_BACK,
            save_dir=TMP_DIR,
            mail_folder=MAIL_FOLDER,
        )
        latest = newest([str(p) for p in paths])
        if not latest:
            print(f"⚠ 没找到附件：{contains}*.xlsx（请检查邮箱/关键字/时间窗口）")
            continue

        # 打印接收时间与大小
        recv_utc = extract_received_utc_from_name(latest)
        sz = os.path.getsize(latest)
        if recv_utc:
            print(f"  • 选用附件：{os.path.basename(latest)}")
            print(f"  • 接收时间(UTC)：{recv_utc.strftime('%Y-%m-%d %H:%M:%S')}  | 大小：{sz:,} bytes")
        else:
            print(f"  • 选用附件：{os.path.basename(latest)}（未解析到时间戳） | 大小：{sz:,} bytes")

        dest_path = os.path.join(DEST_DIR, target_name)
        shutil.copy2(latest, dest_path)  # 覆盖
        print(f"✅ 已复制并重命名：{latest}  →  {dest_path}")

    print("\n🎉 全部完成。")

if __name__ == "__main__":
    main()
