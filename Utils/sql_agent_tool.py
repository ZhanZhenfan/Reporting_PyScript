# -*- coding: utf-8 -*-
"""
SqlAgentTool
------------
启动 SQL Server Agent Job，并等待完成；完成后发出哔哔声并打开指定 Archive 文件夹。

新增
- 参数 `start_step`: 可为 int（step_id）或 str（step_name）。未提供则从 Step 1 开始。

特点
- 仅传入 Job 名（默认精确名）。可选 `fuzzy=True` 尝试模糊匹配（需要读 sysjobs 权限）。
- 先 `sp_start_job`，再通过 `sp_help_job`（可用则优先）或 `sysjobhistory` 轮询状态。
- 成功（run_status=1）时：Beep 并 `os.startfile(archive_dir)` 打开文件夹。
- 失败/取消也会 Beep（不同音调/次数），并返回详细结果。
- 在无读权限时自动降级：盲等到超时也会提示。
"""
from __future__ import annotations

import os
import time
import pyodbc
from dataclasses import dataclass
from typing import Any, Dict, Optional, Union

try:
    import winsound
except Exception:  # 非 Windows 环境兜底
    winsound = None


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


class SqlAgentTool:
    def __init__(self, *, server: str, database: str = "msdb"):
        self.cfg = SqlConn(server=server, database=database)

    # --------- helpers ---------
    def _connect(self) -> pyodbc.Connection:
        return pyodbc.connect(self.cfg.conn_str(), autocommit=self.cfg.autocommit)

    @staticmethod
    def _beep_ok():
        if winsound:
            for _ in range(2):
                winsound.Beep(1000, 250)  # 高音短促两下
                time.sleep(0.05)
        else:
            print("\a\a")  # 退化为控制台响铃

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
                # mac/linux 兜底
                os.system(f'xdg-open "{path}" 2>/dev/null || open "{path}"')
        except Exception:
            pass

    # --------- job discovery (optional fuzzy) ---------
    def _resolve_job_name(self, cur: pyodbc.Cursor, job_name: str, fuzzy: bool) -> str:
        if not fuzzy:
            return job_name
        like = job_name
        if "%" not in like and "_" not in like:
            like = f"%{job_name}%"
        try:
            rows = cur.execute(
                "SELECT name FROM msdb.dbo.sysjobs WITH (NOLOCK) WHERE name LIKE ? ORDER BY name",
                like,
            ).fetchall()
            if not rows:
                raise ValueError(f"未找到匹配 Job: {job_name}")
            if len(rows) > 1:
                names = ", ".join(r[0] for r in rows)
                raise ValueError(f"匹配到多个 Job（请改更精确或 fuzzy=False）: {names}")
            return rows[0][0]
        except pyodbc.ProgrammingError:
            # 无权限读取 sysjobs -> 回退为精确名
            return job_name

    # --------- status polling ---------
    def _can_use_help_job(self, cur: pyodbc.Cursor) -> bool:
        try:
            cur.execute("EXEC msdb.dbo.sp_help_job @execution_status=2")  # 2=Idle, 只是权限探测
            cur.fetchall()
            return True
        except Exception:
            return False

    def _get_history_baseline(self, cur: pyodbc.Cursor, job_name: str) -> int:
        try:
            row = cur.execute(
                """
                SELECT ISNULL(MAX(instance_id), 0) AS max_id
                FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
                WHERE h.job_id = (SELECT job_id FROM msdb.dbo.sysjobs WHERE name=?)
                  AND h.step_id = 0
                """,
                job_name,
            ).fetchone()
            return int(row.max_id if row else 0)
        except Exception:
            return 0

    def _poll_until_finish(
        self,
        cur: pyodbc.Cursor,
        job_name: str,
        baseline_instance_id: int,
        timeout: int,
        poll_interval: int,
    ) -> Dict[str, Any]:
        t0 = time.time()

        # 方案1：优先 sp_help_job 读取当前执行状态
        if self._can_use_help_job(cur):
            while True:
                row = cur.execute(
                    "EXEC msdb.dbo.sp_help_job @job_name = ?",
                    job_name,
                ).fetchone()
                cols = [c[0].lower() for c in cur.description]
                try:
                    idx_stat = cols.index("current_execution_status")
                    idx_outc = cols.index("last_run_outcome")
                except ValueError:
                    idx_stat = None
                    idx_outc = None

                cur_stat = row[idx_stat] if (row and idx_stat is not None) else None
                _ = row[idx_outc] if (row and idx_outc is not None) else None  # 仅探测

                # current_execution_status: 1=Executing, 4=Waiting, 5=Between retries, 7=Completion actions, 2=Idle
                if cur_stat in (None, 2):  # Idle -> 看历史确定本次结果
                    hist = cur.execute(
                        """
                        SELECT TOP 1 run_status, run_date, run_time, message, instance_id
                        FROM msdb.dbo.sysjobhistory WITH (NOLOCK)
                        WHERE job_id = (SELECT job_id FROM msdb.dbo.sysjobs WHERE name = ?)
                          AND step_id = 0
                        ORDER BY instance_id DESC
                        """,
                        job_name,
                    ).fetchone()
                    if hist and (hist.instance_id or 0) > baseline_instance_id:
                        return {
                            "status_code": hist.run_status,
                            "run_date": hist.run_date,
                            "run_time": hist.run_time,
                            "message": hist.message,
                        }
                if time.time() - t0 > timeout:
                    raise TimeoutError(f"等待 Job 超时（{timeout}s）：{job_name}")
                time.sleep(poll_interval)

        # 方案2：无法用 sp_help_job -> 仅依据历史行出现新记录来判断
        while True:
            hist = cur.execute(
                """
                SELECT TOP 1 run_status, run_date, run_time, message, instance_id
                FROM msdb.dbo.sysjobhistory WITH (NOLOCK)
                WHERE job_id = (SELECT job_id FROM msdb.dbo.sysjobs WHERE name = ?)
                  AND step_id = 0
                ORDER BY instance_id DESC
                """,
                job_name,
            ).fetchone()
            if hist and (hist.instance_id or 0) > baseline_instance_id:
                return {
                    "status_code": hist.run_status,
                    "run_date": hist.run_date,
                    "run_time": hist.run_time,
                    "message": hist.message,
                }
            if time.time() - t0 > timeout:
                raise TimeoutError(f"等待 Job 超时（{timeout}s）：{job_name}")
            time.sleep(poll_interval)

    # --------- public API ---------
    def run_job(
        self,
        job_name: str,
        archive_dir: Optional[str] = None,
        timeout: int = 1800,
        poll_interval: int = 5,
        fuzzy: bool = False,
        start_step: Optional[Union[int, str]] = None,  # 新增：支持 int(step_id) 或 str(step_name)
    ):
        """
        启动并等待 SQL Agent Job 执行完成。
        成功返回 dict（含 run_status / message / 时间），失败抛异常。

        参数
        ----
        job_name : Job 名称（默认精确匹配；`fuzzy=True` 时进行 LIKE 模糊匹配）
        archive_dir : 成功后要打开的文件夹路径
        timeout : 最大等待秒数
        poll_interval : 轮询间隔秒
        fuzzy : 是否模糊匹配 job_name
        start_step : int -> 从该 step_id 开始；str -> 从该 step_name 开始；None -> 从 Step 1 开始
        """
        with self._connect() as conn:
            cur = conn.cursor()

            # 1) 解析/校准 Job 名
            target_name = self._resolve_job_name(cur, job_name, fuzzy)
            if isinstance(start_step, int):
                print(f"▶ 开始执行 SQL Job: {target_name}（从 step_id={start_step}）")
            elif isinstance(start_step, str):
                print(f"▶ 开始执行 SQL Job: {target_name}（从 step_name='{start_step}'）")
            else:
                print(f"▶ 开始执行 SQL Job: {target_name}")

            # 2) 执行前取得历史基线（最后一条汇总行 instance_id）
            baseline_id = self._get_history_baseline(cur, target_name)

            # 3) 启动 Job：根据 start_step 类型选择 @step_id 或 @step_name
            if isinstance(start_step, int):
                cur.execute(
                    "EXEC msdb.dbo.sp_start_job @job_name = ?, @step_id = ?",
                    target_name, start_step
                )
            elif isinstance(start_step, str):
                cur.execute(
                    "EXEC msdb.dbo.sp_start_job @job_name = ?, @step_name = ?",
                    target_name, start_step
                )
            else:
                cur.execute("EXEC msdb.dbo.sp_start_job @job_name = ?", target_name)

            # 4) 轮询直到结束（优先 sp_help_job，回退 sysjobhistory）
            try:
                res = self._poll_until_finish(
                    cur,
                    target_name,
                    baseline_instance_id=baseline_id,
                    timeout=timeout,
                    poll_interval=poll_interval,
                )
            except TimeoutError:
                self._beep_fail()
                raise

            # 5) 解析结果（run_status: 0=Failed, 1=Succeeded, 2=Retry, 3=Cancelled, 4=In progress）
            status_code = int(res.get("status_code", -1))
            msg = str(res.get("message", "") or "")
            if status_code == 1:
                print(f"✅ Job {target_name} 执行成功。")
                self._beep_ok()
                if archive_dir and os.path.isdir(archive_dir):
                    self._open_folder(archive_dir)
                return {
                    "ok": True,
                    "job": target_name,
                    "run_status": status_code,
                    "run_date": res.get("run_date"),
                    "run_time": res.get("run_time"),
                    "message": msg,
                }
            else:
                self._beep_fail()
                raise RuntimeError(f"❌ Job {target_name} 执行失败（run_status={status_code}）。{msg}")
