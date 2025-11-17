# -*- coding: utf-8 -*-
import os
import hashlib
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
import win32com.client as win32

from Utils.sql_agent_tool import SqlAgentTool

# ============ 配置 ============
ROOT = Path(r"\\mygbynbyn1msis2\SCM_Excellence\REL Demand")
BASE_NAME = "REL Custom.xlsx"
LOCK_NAME = ".rel_custom_backup.lock"  # 简易互斥锁文件名
ONLY_REFRESH = False  # 设 True 时仅刷新不复制


# ============ 工具函数 ============
def sha256(p: Path, buf_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        while True:
            b = f.read(buf_size)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def atomic_copy(src: Path, dst: Path):
    """原子复制：写到临时文件后 os.replace 到目标，避免中间态文件"""
    with tempfile.NamedTemporaryFile(dir=str(dst.parent), delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        shutil.copy2(src, tmp_path)  # 保留元数据
        os.replace(tmp_path, dst)  # 原子替换
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass


# ============ 主流程 ============
root = ROOT
base_path = root / BASE_NAME
if not base_path.exists():
    raise FileNotFoundError(f"未找到 {base_path}")

# 简易互斥：若锁文件已存在，直接退出（避免多实例并发）
lock_path = root / LOCK_NAME
try:
    lock_fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    os.write(lock_fd, b"lock")
    os.close(lock_fd)
except FileExistsError:
    print("发现锁文件，可能已有实例在运行。本次退出以保证幂等。")
    raise SystemExit(0)

try:
    # 1) 计算备份名（按源文件的最后修改“日期”）
    mtime = datetime.fromtimestamp(base_path.stat().st_mtime)
    date_str = mtime.strftime("%Y%m%d")
    dated_path = root / f"REL Custom - {date_str}.xlsx"

    if not ONLY_REFRESH:
        if dated_path.exists():
            # 校验内容是否一致
            src_hash = sha256(base_path)
            dst_hash = sha256(dated_path)
            if src_hash == dst_hash:
                print(f"[幂等] 备份已存在且内容一致：{dated_path.name}，跳过复制。")
            else:
                # 仅覆盖，不新建额外文件
                atomic_copy(base_path, dated_path)
                print(f"[更新] 已覆盖备份：{dated_path.name}")
        else:
            # 首次生成今日备份
            atomic_copy(base_path, dated_path)
            print(f"[创建] 已生成备份：{dated_path.name}")

    # 2) 刷新原始文件的数据源并保存
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(base_path.as_posix(), ReadOnly=False, UpdateLinks=False)
    try:
        wb.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        wb.Save()
        print(f"[刷新] 已刷新并保存：{BASE_NAME}")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

finally:
    # 释放锁
    try:
        lock_path.unlink()
    except Exception:
        pass

tool = SqlAgentTool(server="tcp:10.80.127.71,1433")

result = tool.run_job(
    job_name="Lumileds BI - SC RelSNOP",  # 用完整精确名最稳妥
    timeout=100,
    poll_interval=3,
    fuzzy=False,  # 若你 later 拿到读 sysjobs 的权限，可改 True
)
print(result)
