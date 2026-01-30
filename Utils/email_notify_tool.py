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
