# -*- coding: utf-8 -*-
"""
Email notification helper (SMTP).

Usage:
  from Utils.email_notify_tool import EmailNotifier
  notifier = EmailNotifier.from_env()
  notifier.send(
      subject="Job done",
      body="Task succeeded",
      to=["a@b.com"],
  )
"""

from __future__ import annotations

import os
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from typing import Iterable, List, Optional


def _to_list(v: Optional[Iterable[str]]) -> List[str]:
    if not v:
        return []
    if isinstance(v, str):
        return [v]
    return [s for s in v if s]


DEFAULT_TO = [
    "Haibo.Zhang@lumileds.com",
    "Zhenfan.Zhan@lumileds.com",
    "lalitha.namburi@lumileds.com",
]


@dataclass
class SmtpConfig:
    host: str
    port: int = 587
    user: Optional[str] = None
    password: Optional[str] = None
    use_tls: bool = True
    from_addr: Optional[str] = None


class EmailNotifier:
    def __init__(self, cfg: SmtpConfig):
        self.cfg = cfg

    @staticmethod
    def _default_config_path() -> str:
        here = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(here, "email_notify_config.json")

    @staticmethod
    def _load_json(path: str) -> dict:
        try:
            import json
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    @classmethod
    def from_env(cls) -> "EmailNotifier":
        """
        Load SMTP config from environment:
          SMTP_HOST (required)
          SMTP_PORT (default 587)
          SMTP_USER
          SMTP_PASS
          SMTP_USE_TLS (default true)
          SMTP_FROM (default SMTP_USER)
        """
        host = os.getenv("SMTP_HOST", "").strip()
        if not host:
            raise ValueError("SMTP_HOST is required")
        port = int(os.getenv("SMTP_PORT", "587"))
        user = os.getenv("SMTP_USER")
        password = os.getenv("SMTP_PASS")
        use_tls = os.getenv("SMTP_USE_TLS", "true").strip().lower() not in {"0", "false", "no"}
        from_addr = os.getenv("SMTP_FROM") or user
        return cls(SmtpConfig(host=host, port=port, user=user, password=password, use_tls=use_tls, from_addr=from_addr))

    @classmethod
    def from_config(cls, config_path: str | None = None) -> "EmailNotifier":
        """
        Load SMTP config from JSON config:
          smtp.host (required)
          smtp.port (default 25)
          smtp.user
          smtp.password
          smtp.use_tls (default false)
          smtp.from_addr (default smtp.user or empty)
        """
        cfg_path = config_path or os.getenv("EMAIL_NOTIFY_CONFIG") or cls._default_config_path()
        cfg = cls._load_json(cfg_path)
        smtp = cfg.get("smtp", {}) if isinstance(cfg, dict) else {}
        host = (smtp.get("host") or "").strip()
        if not host:
            raise ValueError("smtp.host is required in email_notify_config.json")
        port = int(smtp.get("port") or 25)
        user = smtp.get("user")
        password = smtp.get("password")
        use_tls = bool(smtp.get("use_tls")) if "use_tls" in smtp else False
        from_addr = smtp.get("from_addr") or user
        return cls(SmtpConfig(host=host, port=port, user=user, password=password, use_tls=use_tls, from_addr=from_addr))

    def _resolve_recipients(self, job_key: str | None, config_path: str | None = None) -> tuple[list[str], list[str], list[str]]:
        cfg_path = config_path or os.getenv("EMAIL_NOTIFY_CONFIG") or self._default_config_path()
        cfg = self._load_json(cfg_path)
        default_cfg = cfg.get("default", {}) if isinstance(cfg, dict) else {}
        jobs_cfg = cfg.get("jobs", {}) if isinstance(cfg, dict) else {}
        job_cfg = jobs_cfg.get(job_key, {}) if job_key and isinstance(jobs_cfg, dict) else {}

        def _get_list(key: str, fallback: list[str]) -> list[str]:
            val = job_cfg.get(key) if isinstance(job_cfg, dict) else None
            if val is None:
                val = default_cfg.get(key) if isinstance(default_cfg, dict) else None
            if val is None:
                return fallback
            return _to_list(val)

        to_list = _get_list("to", DEFAULT_TO)
        cc_list = _get_list("cc", [])
        bcc_list = _get_list("bcc", [])
        return to_list, cc_list, bcc_list

    def send(
        self,
        *,
        subject: str,
        body: str,
        to: Iterable[str],
        cc: Optional[Iterable[str]] = None,
        bcc: Optional[Iterable[str]] = None,
        subtype: str = "plain",
    ) -> None:
        to_list = _to_list(to)
        cc_list = _to_list(cc)
        bcc_list = _to_list(bcc)
        if not to_list and not cc_list and not bcc_list:
            raise ValueError("No recipients provided")

        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = self.cfg.from_addr or ""
        msg["To"] = ", ".join(to_list)
        if cc_list:
            msg["Cc"] = ", ".join(cc_list)
        msg.set_content(body, subtype=subtype)

        all_rcpt = to_list + cc_list + bcc_list

        with smtplib.SMTP(self.cfg.host, self.cfg.port, timeout=60) as s:
            if self.cfg.use_tls:
                s.starttls()
            if self.cfg.user and self.cfg.password:
                s.login(self.cfg.user, self.cfg.password)
            s.send_message(msg, from_addr=self.cfg.from_addr, to_addrs=all_rcpt)

    def send_with_config(
        self,
        *,
        job_key: str,
        subject: str,
        body: str,
        config_path: str | None = None,
        subtype: str = "plain",
    ) -> None:
        to_list, cc_list, bcc_list = self._resolve_recipients(job_key, config_path=config_path)
        self.send(subject=subject, body=body, to=to_list, cc=cc_list, bcc=bcc_list, subtype=subtype)
