# -*- coding: utf-8 -*-
"""
Copy latest 'REL Custom.xlsx' then (optionally) trigger a SQL Agent job via your SqlAgentTool.
"""

import os
import shutil
import time
from glob import glob

# 如果 SqlAgentTool 在另一个文件/包，请改成你的导入方式：
# from your_module import SqlAgentTool
# 这里直接从同文件作用域使用（假设你已把 SqlAgentTool 类放到同一工程里）。
from Utils.sql_agent_tool import SqlAgentTool

# ---------------------- 路径配置 ----------------------
SRC_DIR  = r"\\mygbynbyn1msis2\SCM_Excellence\REL Demand"
DEST_DIR = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\SNOP Reports\RnD REL SNOP"
DEST_NAME = "REL Custom.xlsx"  # 目标文件名（固定为这个）

# 匹配“最新”的文件；如果目录里只有固定文件名，也可以改成 ['REL Custom.xlsx']
SRC_PATTERN = "REL Custom*.xlsx"

# ---------------------- SQL Agent 配置（可选） ----------------------
RUN_SQL_JOB = True                 # 不跑 SQL Job 就改为 False
SQL_SERVER  = r"tcp:10.80.127.71,1433"     # ← 改成你的 SQL Server 名（如 'myssql01\prod'）
JOB_NAME    = "Lumileds BI - SC RelSNOP"  # ← 要触发的 Job 名
ARCHIVE_DIR = "\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\SNOP Reports\RnD REL SNOP\Archive"
TIMEOUT_SEC = 1800                 # 等待最多 30 分钟
POLL_SEC    = 5                    # 轮询间隔 5 秒
START_STEP  = None                 # 可设成 int(step_id) 或 str(step_name)，默认从 Step 1 开始
# ----------------------------------------------------

def _latest_rel_custom(src_dir: str, pattern: str) -> str:
    """在 src_dir 找到最新的 REL Custom 文件并返回绝对路径。"""
    candidates = [p for p in glob(os.path.join(src_dir, pattern)) if os.path.isfile(p)]
    if not candidates:
        raise FileNotFoundError(f"未在 {src_dir} 找到匹配文件：{pattern} / No matching file in {src_dir}: {pattern}")
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]

def _copy_with_retry(src: str, dst: str, tries: int = 5, delay: float = 1.0) -> None:
    """带重试的覆盖复制，避免网络共享临时锁住时报错。"""
    last_err = None
    for i in range(1, tries + 1):
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            return
        except Exception as e:
            last_err = e
            print(f"⏳ 复制重试 {i}/{tries} 失败：{e} / Copy retry {i}/{tries} failed: {e}")
            time.sleep(delay)
    raise RuntimeError(f"复制失败：{src} -> {dst}\n最后错误：{last_err} / Copy failed: {src} -> {dst}\nLast error: {last_err}")

def main():
    print("==== REL Custom | 复制最新并（可选）跑 SQL Job ===="
          " / REL Custom | Copy latest and (optional) run SQL Job ====")
    # 1) 找最新
    src_path = _latest_rel_custom(SRC_DIR, SRC_PATTERN)
    print(f"📄 最新源文件：{os.path.basename(src_path)} / Latest source file: {os.path.basename(src_path)}")

    # 2) 复制到目标（同名覆盖）
    dest_path = os.path.join(DEST_DIR, DEST_NAME)
    print(f"📥 复制到：{dest_path} / Copy to: {dest_path}")
    _copy_with_retry(src_path, dest_path)
    print("✅ 复制完成。 / Copy completed.")

    # 3) 可选：触发 SQL Agent Job
    if RUN_SQL_JOB:
        print(f"▶ 触发 SQL Agent Job：{JOB_NAME} @ {SQL_SERVER} / Triggering SQL Agent Job: {JOB_NAME} @ {SQL_SERVER}")
        agent = SqlAgentTool(server=SQL_SERVER)
        res = agent.run_job(
            job_name=JOB_NAME,
            archive_dir=ARCHIVE_DIR,
            timeout=TIMEOUT_SEC,
            poll_interval=POLL_SEC,
            start_step=START_STEP,
        )
        print(f"✅ SQL Job 完成：{res} / SQL Job completed: {res}")
    else:
        print("ℹ️ 已关闭 SQL Job 触发（RUN_SQL_JOB=False）。 / SQL Job trigger disabled (RUN_SQL_JOB=False).")

if __name__ == "__main__":
    main()
