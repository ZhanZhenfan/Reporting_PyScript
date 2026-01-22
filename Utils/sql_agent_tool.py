# -*- coding: utf-8 -*-
"""
SqlAgentTool (Final + step_id support via step_name)
---------------------------------------------------
启动 SQL Server Agent Job，并等待完成（优先文件检测模式）。
- 支持 start_step：int(通过 sysjobsteps 解析为 step_name) 或 str(直接作为 step_name)。
- 若指定的 step 不存在/不可解析：立即报错退出，不再继续执行。
- 默认启用“文件检测模式”（只要传了 archive_dir），检测新文件/mtime 变化即判定成功。
"""

from __future__ import annotations
import os
import time
from glob import glob
from dataclasses import dataclass
from typing import Any, Dict, Optional, Union
import pyodbc

try:
    import winsound
except Exception:
    winsound = None


# -------------------- Connection Config --------------------
@dataclass
class SqlConn:
    server: str
    database: str = "msdb"
    driver: str = "ODBC Driver 17 for SQL Server"
    trusted: bool = True
    encrypt: bool = True
    trust_cert: bool = True
    autocommit: bool = True

    def conn_str(self) -> str:
        parts = [
            f"DRIVER={{{self.driver}}}",
            f"SERVER={self.server}",
            f"DATABASE={self.database}",
        ]
        if self.trusted:
            parts.append("Trusted_Connection=yes")
        if self.encrypt:
            parts.append("Encrypt=yes")
        if self.trust_cert:
            parts.append("TrustServerCertificate=yes")
        return ";".join(parts) + ";"


# -------------------- Main Tool --------------------
class SqlAgentTool:
    def __init__(self, *, server: str, database: str = "msdb"):
        self.cfg = SqlConn(server=server, database=database)

    # ----------- Internal Utilities -----------
    def _connect(self) -> pyodbc.Connection:
        return pyodbc.connect(self.cfg.conn_str(), autocommit=self.cfg.autocommit)

    @staticmethod
    def _beep_ok():
        if winsound:
            for _ in range(2):
                winsound.Beep(1000, 250)
                time.sleep(0.05)
        else:
            print("\a\a")

    @staticmethod
    def _beep_fail():
        if winsound:
            for f in (400, 350, 300):
                winsound.Beep(f, 300)
                time.sleep(0.05)
        else:
            print("\a\a\a")

    @staticmethod
    def _open_folder(path: str):
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                os.system(f'xdg-open "{path}" 2>/dev/null || open "{path}"')
        except Exception:
            pass

    # ----------- Job/Step Resolve -----------
    def _resolve_job_name(self, cur: pyodbc.Cursor, job_name: str, fuzzy: bool) -> str:
        if not fuzzy:
            return job_name
        like = job_name if any(c in job_name for c in "%_") else f"%{job_name}%"
        try:
            rows = cur.execute(
                "SELECT name FROM msdb.dbo.sysjobs WITH (NOLOCK) WHERE name LIKE ? ORDER BY name",
                like,
            ).fetchall()
            if not rows:
                raise ValueError(f"未找到匹配 Job: {job_name} / No matching job found: {job_name}")
            if len(rows) > 1:
                names = ", ".join(r[0] for r in rows)
                raise ValueError(
                    f"匹配到多个 Job（请改更精确或 fuzzy=False）: {names} / "
                    f"Multiple jobs matched (use a more specific name or fuzzy=False): {names}"
                )
            return rows[0][0]
        except pyodbc.ProgrammingError:
            return job_name

    def _resolve_step_name_from_id(self, cur: pyodbc.Cursor, job_name: str, step_id: int) -> Optional[str]:
        """
        将 step_id 解析为 step_name（按 job_name 精确匹配）。
        无权限或未找到时返回 None。
        """
        try:
            row = cur.execute(
                """
                SELECT s.step_name
                FROM msdb.dbo.sysjobsteps AS s WITH (NOLOCK)
                JOIN msdb.dbo.sysjobs AS j WITH (NOLOCK) ON s.job_id = j.job_id
                WHERE j.name = ? AND s.step_id = ?
                """,
                job_name, step_id
            ).fetchone()
            return (row.step_name if row else None)
        except Exception:
            return None

    def _step_exists_by_name(self, cur: pyodbc.Cursor, job_name: str, step_name: str) -> bool:
        try:
            row = cur.execute(
                """
                SELECT 1
                FROM msdb.dbo.sysjobsteps AS s WITH (NOLOCK)
                JOIN msdb.dbo.sysjobs   AS j WITH (NOLOCK) ON s.job_id = j.job_id
                WHERE j.name = ? AND s.step_name = ?
                """,
                job_name, step_name
            ).fetchone()
            return bool(row)
        except Exception:
            return False

    # ----------- File Watch Logic -----------
    @staticmethod
    def _latest_file_state(folder: str, pattern: str) -> Dict[str, Any]:
        files = glob(os.path.join(folder, pattern))
        if not files:
            return {"count": 0, "latest_mtime": 0.0, "latest_file": None}
        latest_file = max(files, key=os.path.getmtime)
        return {
            "count": len(files),
            "latest_mtime": os.path.getmtime(latest_file),
            "latest_file": latest_file,
        }

    def _poll_until_file_appears(
        self,
        archive_dir: str,
        pattern: str,
        timeout: int,
        poll_interval: int,
        requires_new_file: bool,
        baseline_state: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        if not os.path.isdir(archive_dir):
            raise FileNotFoundError(f"Archive 目录不存在：{archive_dir} / Archive directory not found: {archive_dir}")

        if baseline_state is None:
            baseline_state = self._latest_file_state(archive_dir, pattern)

        base_cnt = baseline_state["count"]
        base_mtm = baseline_state["latest_mtime"]
        base_name = baseline_state["latest_file"]

        print(f"⏳ 监控目录：{archive_dir} | 模式：{pattern} / Monitoring folder: {archive_dir} | Pattern: {pattern}")
        t0 = time.time()
        while True:
            cur = self._latest_file_state(archive_dir, pattern)
            if requires_new_file:
                if (cur["count"] > base_cnt) or (cur["latest_file"] and cur["latest_file"] != base_name):
                    return {"ok": True, "detected": cur}
            else:
                if cur["latest_mtime"] > base_mtm:
                    return {"ok": True, "detected": cur}

            if time.time() - t0 > timeout:
                raise TimeoutError(f"等待归档新文件超时（{timeout}s） / Timeout waiting for new archive file ({timeout}s)")
            time.sleep(poll_interval)

    # ----------- Main Run Job API -----------
    def run_job(
        self,
        job_name: str,
        archive_dir: Optional[str] = None,
        timeout: int = 1800,
        poll_interval: int = 5,
        fuzzy: bool = False,
        start_step: Optional[Union[int, str]] = None,
        # ⬇️ 可选的文件检测参数（有默认）
        use_file_watch: Optional[bool] = None,
        archive_pattern: Optional[str] = None,
        file_watch_requires_new_file: Optional[bool] = None,
    ) -> Dict[str, Any]:
        """
        启动并等待 SQL Agent Job 完成。
        - 若提供 archive_dir，将自动启用“文件检测模式”作为成功判定；
        - start_step:
            * int  -> 按 step_id 解析为 step_name 并从该步启动；解析失败则报错退出；
            * str  -> 视为 step_name，先校验存在性；不存在则报错退出；
        """

        # ------- 默认逻辑（文件检测）-------
        if use_file_watch is None:
            use_file_watch = bool(archive_dir)
        if not archive_pattern:
            archive_pattern = "*.xlsx"
        if file_watch_requires_new_file is None:
            file_watch_requires_new_file = False
        # -----------------------------------

        with self._connect() as conn:
            cur = conn.cursor()

            # 解析 Job 名
            target_name = self._resolve_job_name(cur, job_name, fuzzy)

            # 解析/校验 start_step
            step_name_to_start: Optional[str] = None
            if isinstance(start_step, int):
                # 无论 1 或更大，均尝试解析为 step_name；失败直接退出
                step_name_to_start = self._resolve_step_name_from_id(cur, target_name, start_step)
                if not step_name_to_start:
                    self._beep_fail()
                    raise ValueError(
                        f"指定的 step_id={start_step} 在 Job '{target_name}' 中不存在或不可访问。已停止执行。 / "
                        f"step_id={start_step} not found or inaccessible in job '{target_name}'. Stopping."
                    )
                print(
                    f"▶ 启动 SQL Job: {target_name}（按 step_id={start_step} → step_name='{step_name_to_start}'） / "
                    f"Starting SQL Job: {target_name} (step_id={start_step} -> step_name='{step_name_to_start}')"
                )

            elif isinstance(start_step, str) and start_step.strip():
                step_name_to_start = start_step.strip()
                if not self._step_exists_by_name(cur, target_name, step_name_to_start):
                    self._beep_fail()
                    raise ValueError(
                        f"指定的 step_name='{step_name_to_start}' 在 Job '{target_name}' 中不存在。已停止执行。 / "
                        f"step_name='{step_name_to_start}' not found in job '{target_name}'. Stopping."
                    )
                print(
                    f"▶ 启动 SQL Job: {target_name}（按 step_name='{step_name_to_start}'） / "
                    f"Starting SQL Job: {target_name} (step_name='{step_name_to_start}')"
                )

            else:
                print(f"▶ 启动 SQL Job: {target_name}（从 Step 1 开始） / Starting SQL Job: {target_name} (from Step 1)")

            # 文件检测基线（若启用）
            baseline_state = None
            if use_file_watch and archive_dir:
                baseline_state = self._latest_file_state(archive_dir, archive_pattern)

            # 启动 Job —— 注意：sp_start_job 只支持 @step_name，不支持 @step_id
            if step_name_to_start:
                cur.execute("EXEC msdb.dbo.sp_start_job @job_name = ?, @step_name = ?", target_name, step_name_to_start)
            else:
                cur.execute("EXEC msdb.dbo.sp_start_job @job_name = ?", target_name)

            # 文件检测模式（优先）
            if use_file_watch and archive_dir:
                print("⏳ 等待归档目录出现新文件（文件监控判定成功）... / Waiting for new files in archive folder...")
                ok = self._poll_until_file_appears(
                    archive_dir=archive_dir,
                    pattern=archive_pattern,
                    timeout=timeout,
                    poll_interval=poll_interval,
                    requires_new_file=file_watch_requires_new_file,
                    baseline_state=baseline_state,
                )
                if ok:
                    self._beep_ok()
                    if os.path.isdir(archive_dir):
                        self._open_folder(archive_dir)
                    return {"ok": True, "job": target_name, "mode": "file_watch", **ok}

            # 兜底：未启用文件检测则简单等待 timeout 秒后返回
            print("ℹ️ 未启用文件检测模式，将等待数据库任务状态返回（或超时）… / "
                  "File watch disabled; waiting for job status (or timeout)...")
            time.sleep(timeout)
            self._beep_ok()
            return {"ok": True, "job": target_name, "mode": "timeout-fallback"}
