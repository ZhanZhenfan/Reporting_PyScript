# -*- coding: utf-8 -*-
"""
流程：
1) 在 MRP 目录中找到现有的 ReportRefinedSeleneSupplyDemand.csv
   → 读取其 “最后修改时间” 生成 YYYYMMDD
   → 重命名为 ReportRefinedSeleneSupplyDemand_YYYYMMDD.csv（若重名自动 _v2/_v3…）
2) 在 Archive 目录中找到最新的 ReportRefinedSeleneSupplyDemand*.csv
   → 复制到 MRP 并命名为 ReportRefinedSeleneSupplyDemand.csv
3) 执行 BAT：\\mygbynbyn1msis1\Supply-Chain-Analytics\Temp Report\CopyPasteSelene and PlateletGrouping.BAT
"""

import os
import glob
import shutil
import subprocess
from datetime import datetime

from Utils.sql_agent_tool import SqlAgentTool

# ---------------- 配置区 ----------------
SRC_DIR = r"\\sggsintsysvw068\data\SCPS\Interfaces\ReportRefinedSeleneSupplyDemand\Archive"
DST_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\DataFile\MRP"
DST_FIXED_NAME = "ReportRefinedSeleneSupplyDemand.csv"
BAT_FILE = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Temp Report\CopyPasteSelene and PlateletGrouping.BAT"
PATTERN = "ReportRefinedSeleneSupplyDemand*.csv"   # 在 Archive 中匹配的文件模式
# --------------------------------------

def ensure_dir(path: str):
    if not os.path.isdir(path):
        raise FileNotFoundError(f"目录不存在：{path} / Directory not found: {path}")

def latest_file(folder: str, pattern: str) -> str:
    files = glob.glob(os.path.join(folder, pattern))
    if not files:
        raise FileNotFoundError(f"未在 {folder} 找到匹配文件：{pattern} / No matching file in {folder}: {pattern}")
    return max(files, key=os.path.getmtime)

def uniquify(path: str) -> str:
    """若 path 已存在，则在扩展名前追加 _v2/_v3... 返回不重名的路径"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while True:
        candidate = f"{base}_v{i}{ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1

def backup_existing_dst(dst_dir: str, fixed_name: str) -> str | None:
    """将 MRP 目录中现有的固定文件按其 mtime 备份为 _YYYYMMDD.csv，返回备份路径；不存在则返回 None"""
    fixed_path = os.path.join(dst_dir, fixed_name)
    if not os.path.exists(fixed_path):
        print(f"ℹ️ 目标目录中不存在 {fixed_name}，跳过备份。 / {fixed_name} not found in destination; skip backup.")
        return None

    mtime = os.path.getmtime(fixed_path)
    ymd = datetime.fromtimestamp(mtime).strftime("%Y%m%d")
    bak_name = f"ReportRefinedSeleneSupplyDemand_{ymd}.csv"
    bak_path = os.path.join(dst_dir, bak_name)
    bak_path = uniquify(bak_path)  # 若同日已有备份，追加 _v2/_v3…

    # 用 move 更快也保留原文件时间戳
    shutil.move(fixed_path, bak_path)
    print(f"✅ 已备份：{fixed_path}  →  {bak_path} / Backed up: {fixed_path} -> {bak_path}")
    return bak_path

def copy_latest_from_src(src_dir: str, pattern: str, dst_dir: str, fixed_name: str) -> str:
    src_latest = latest_file(src_dir, pattern)
    dst_fixed = os.path.join(dst_dir, fixed_name)
    shutil.copy2(src_latest, dst_fixed)
    print(f"✅ 已复制最新源文件：\n    {src_latest}\n  → {dst_fixed}\n  / Copied latest source file.")
    return dst_fixed

def run_bat(bat_path: str):
    print(f"▶️ 执行批处理：{bat_path} / Running batch: {bat_path}")
    # 用 cmd /c 处理带空格的 UNC 路径；check=True 失败会抛异常
    subprocess.run(["cmd", "/c", bat_path], check=True)
    print("✅ 批处理执行完成。 / Batch completed.")

def main():
    ensure_dir(SRC_DIR)
    ensure_dir(DST_DIR)

    # Step 1: 备份 MRP 里现有固定文件
    backup_existing_dst(DST_DIR, DST_FIXED_NAME)

    # Step 2: 从 Archive 复制最新的文件到 MRP（固定名）
    copy_latest_from_src(SRC_DIR, PATTERN, DST_DIR, DST_FIXED_NAME)

    # Step 3: 执行 BAT
    run_bat(BAT_FILE)

    tool = SqlAgentTool(server="tcp:10.80.127.71,1433")

    result = tool.run_job(
        job_name="Lumileds BI - SC MRP Waterfall",  # 用完整精确名最稳妥
        archive_dir=r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\Solver\Transactional Data\Platelet Waterfall\Archive",
        timeout=1800,
        poll_interval=3,
        fuzzy=False,  # 若你 later 拿到读 sysjobs 的权限，可改 True
    )
    print(result)

if __name__ == "__main__":
    try:
        main()
        print("🎉 全流程完成。 / Workflow completed.")
    except Exception as e:
        print(f"❌ 出错：{e} / Error: {e}")
        raise
