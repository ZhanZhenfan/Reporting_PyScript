"""
Graph Mail Attachment Tool
---------------------------------
将 Microsoft Graph 邮件附件下载逻辑封装为可复用工具：
- 支持 Device Code 首次登录 + refresh_token 缓存
- 按名称关键词/精确名/扩展名筛选、只取最新 N 个
- 限定近 N 天、分页扫描上限
- 保存到指定目录并返回已保存文件路径列表（按接收时间新→旧）
- 既可作为库被 import，也可命令行调用

示例（库用法）
-----------------
from graph_mail_attachment_tool import GraphMailAttachmentTool

tool = GraphMailAttachmentTool(
    tenant_id="5c2be51b-4109-461d-a0e7-521be6237ce2",
    client_id="09004044-1c60-48e5-b1eb-bb42b3892006",
)
paths = tool.download_latest_attachments(
    contains="ZMRP_WATERFALL_Run",
    ext=".xlsx",
    need_count=2,
    days_back=90,
    save_dir=r"C:\\WeeklyReport\\Download_From_Eamil",
)
print(paths)

命令行：
python graph_mail_attachment_tool.py \
  --tenant-id 5c2be51b-4109-461d-a0e7-521be6237ce2 \
  --client-id 09004044-1c60-48e5-b1eb-bb42b3892006 \
  --contains ZMRP_WATERFALL_Run --ext .xlsx --need-count 2 --days-back 90 \
  --save-dir "C:\\WeeklyReport\\Download_From_Eamil"
"""

from __future__ import annotations

import os
import time
import json
import base64
import datetime as dt
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

import requests

try:
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except Exception:  # pragma: no cover - 环境没有 urllib3 Retry 也能运行
    HTTPAdapter = None
    Retry = None


# =========================
# 工具类封装
# =========================
@dataclass
class AuthConfig:
    tenant_id: str
    client_id: str
    scopes: str = "Mail.Read offline_access"
    token_cache: str = "graph_token_cache.json"


class GraphMailAttachmentTool:
    """Microsoft Graph 邮件附件下载器（带 Token 缓存）。"""

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant_id: str, client_id: str, *,
                 scopes: str = "Mail.Read offline_access",
                 token_cache: str = "graph_token_cache.json",
                 session: Optional[requests.Session] = None,
                 request_timeout: int = 60):
        self.auth = AuthConfig(tenant_id=tenant_id, client_id=client_id,
                               scopes=scopes, token_cache=token_cache)
        self.request_timeout = int(request_timeout)

        # AAD endpoints
        self.auth_base = f"https://login.microsoftonline.com/{self.auth.tenant_id}/oauth2/v2.0"
        self.device_code_url = f"{self.auth_base}/devicecode"
        self.token_url = f"{self.auth_base}/token"

        # HTTP session（带重试）
        self.session = session or self._build_session()

    # ---------- Session & helpers ----------
    def _build_session(self) -> requests.Session:
        s = requests.Session()
        s.headers.update({
            "User-Agent": "GraphMailAttachmentTool/1.0 (+https://microsoft.com/graph)",
        })
        if HTTPAdapter and Retry:
            retry = Retry(
                total=5,
                backoff_factor=0.6,
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["GET", "POST"],
            )
            adapter = HTTPAdapter(max_retries=retry)
            s.mount("https://", adapter)
            s.mount("http://", adapter)
        return s

    @staticmethod
    def _now() -> int:
        return int(time.time())

    def _load_tok(self) -> Optional[dict]:
        p = Path(self.auth.token_cache)
        if not p.exists():
            return None
        try:
            return json.loads(p.read_text("utf-8"))
        except Exception:
            return None

    def _save_tok(self, tok: dict) -> None:
        Path(self.auth.token_cache).write_text(
            json.dumps(tok, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    @staticmethod
    def _valid(tok: Optional[dict], skew: int = 60) -> bool:
        return bool(tok and "access_token" in tok and tok.get("expires_at", 0) - GraphMailAttachmentTool._now() > skew)

    def _refresh(self, tok: Optional[dict]) -> Optional[dict]:
        if not tok or "refresh_token" not in tok:
            return None
        r = self.session.post(self.token_url, data={
            "grant_type": "refresh_token",
            "client_id": self.auth.client_id,
            "refresh_token": tok["refresh_token"],
            "scope": self.auth.scopes,
        }, timeout=self.request_timeout)
        if r.status_code != 200:
            return None
        d = r.json()
        d["expires_at"] = self._now() + int(d.get("expires_in", 3600))
        # 有些情况下返回不会携带新的 refresh_token，需保留旧的
        d.setdefault("refresh_token", tok.get("refresh_token"))
        self._save_tok(d)
        return d

    def _device_login(self) -> dict:
        dc = self.session.post(self.device_code_url, data={
            "client_id": self.auth.client_id,
            "scope": self.auth.scopes,
        }, timeout=self.request_timeout).json()
        print("[LOGIN] 打开网址并输入验证码完成授权 / Open the URL and enter the code to authorize:")
        print("         URL:", dc.get("verification_uri"))
        print("         CODE:", dc.get("user_code"))
        print("         等待你完成登录... / Waiting for you to finish sign-in...")
        start = time.time()
        while True:
            if time.time() - start > dc["expires_in"]:
                raise RuntimeError("Device code 已过期，请重试运行 / Device code expired, please rerun.")
            r = self.session.post(self.token_url, data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": self.auth.client_id,
                "device_code": dc["device_code"],
            }, timeout=self.request_timeout)
            d = r.json()
            if "access_token" in d:
                d["expires_at"] = self._now() + int(d.get("expires_in", 3600))
                self._save_tok(d)
                print("[LOGIN] ✅ 首次授权成功，已获取 Access Token / First authorization succeeded; access token obtained")
                return d
            if d.get("error") in ("authorization_pending", "slow_down"):
                time.sleep(dc.get("interval", 5))
                continue
            raise RuntimeError(f"Token error: {d}")

    def get_access_token(self) -> str:
        tok = self._load_tok()
        if self._valid(tok):
            return tok["access_token"]
        tok2 = self._refresh(tok)
        if tok2 and "access_token" in tok2:
            print("[LOGIN] 🔄 已刷新 Access Token / Access token refreshed")
            return tok2["access_token"]
        tok3 = self._device_login()
        return tok3["access_token"]

    # ---------- HTTP wrappers ----------
    def _gget(self, url: str, token: str, params: Optional[dict] = None) -> dict:
        r = self.session.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=self.request_timeout)
        if r.status_code >= 400:
            msg = f"{r.status_code} GET {url}\n{r.text[:500]}"
            raise requests.HTTPError(msg)
        return r.json()

    def _download_value(self, url: str, token: str) -> bytes:
        r = self.session.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=max(180, self.request_timeout))
        r.raise_for_status()
        return r.content

    # ---------- Utilities ----------
    @staticmethod
    def _iso_utc_minus_days(days: int) -> str:
        t = dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=max(0, int(days)))
        return t.isoformat(timespec="seconds").replace("+00:00", "Z")

    @staticmethod
    def _parse_graph_dt(s: str) -> dt.datetime:
        if not s:
            return dt.datetime.min.replace(tzinfo=dt.timezone.utc)
        if s.endswith("Z"):
            s = s.replace("Z", "+00:00")
        return dt.datetime.fromisoformat(s)

    @staticmethod
    def _safe_name(n: str) -> str:
        return "".join(ch for ch in n if ch not in '\\/:*?"<>|')

    # ---------- Public API ----------
    def download_latest_attachments(
        self,
        *,
        contains: Optional[str] = None,
        equals: Optional[str] = None,
        ext: Optional[str] = None,
        need_count: int = 2,
        days_back: int = 90,
        page_size: int = 50,
        max_scan: int = 800,
        save_dir: str | os.PathLike = ".",
        mail_folder: Optional[str] = None,  # e.g. "inbox"；默认所有文件夹
    ) -> List[Path]:
        """
        下载符合筛选条件的最新 N 个附件，并返回保存路径列表（按接收时间新→旧）。

        :param contains: 文件名包含关键字（不区分大小写）。
        :param equals:   文件名精确等于（优先级高于 contains）。
        :param ext:      扩展名（如 ".xlsx"），若提供则按后缀过滤（不区分大小写）。
        :param need_count: 需要下载的附件个数（默认 2）。
        :param days_back:  仅考虑最近 N 天邮件（默认 90）。
        :param page_size:  每页抓取邮件数（默认 50）。
        :param max_scan:   最多扫描的邮件数（默认 800）。
        :param save_dir:   保存目录（不存在将自动创建）。
        :param mail_folder:指定邮件夹（如 "inbox"），不传则扫描所有文件夹。
        :return:           List[Path] 已保存文件路径，按邮件接收时间降序。
        """
        save_path = Path(save_dir)
        save_path.mkdir(parents=True, exist_ok=True)

        token = self.get_access_token()
        since_iso = self._iso_utc_minus_days(days_back)
        since_dt = self._parse_graph_dt(since_iso)

        # Messages endpoint（可选限定到某个文件夹）
        if mail_folder:
            messages_url = f"{self.GRAPH_BASE}/me/mailFolders/{mail_folder}/messages"
        else:
            messages_url = f"{self.GRAPH_BASE}/me/messages"

        params = {
            "$select": "id,subject,receivedDateTime,hasAttachments",
            "$orderby": "receivedDateTime desc",
            "$top": str(int(page_size)),
        }

        def _name_match(name: str) -> bool:
            low = (name or "").lower()
            if equals:
                return name == equals
            ok = True
            if contains:
                ok = (contains.lower() in low)
            if ok and ext:
                ok = low.endswith(ext.lower())
            return ok

        saved: List[Tuple[dt.datetime, Path]] = []
        found = scanned = 0
        url = messages_url
        local_params = params
        while url and scanned < max_scan and found < need_count:
            data = self._gget(url, token, params=local_params)
            local_params = None  # nextLink 已经包含分页参数
            msgs = data.get("value", [])
            if not msgs:
                break

            for m in msgs:
                scanned += 1
                rdt_str = m.get("receivedDateTime") or ""
                if not rdt_str:
                    continue
                rdt = self._parse_graph_dt(rdt_str)
                if rdt < since_dt:
                    url = None  # 时间更老，无需继续分页
                    break
                if not m.get("hasAttachments"):
                    continue

                mid = m["id"]
                # 读取附件元数据
                atts = self._gget(f"{self.GRAPH_BASE}/me/messages/{mid}/attachments", token)
                for a in atts.get("value", []):
                    if a.get("isInline"):
                        continue
                    otype = a.get("@odata.type", "")
                    if otype and not otype.endswith("fileAttachment"):
                        continue

                    name = a.get("name") or "attachment.bin"
                    if not _name_match(name):
                        continue

                    # 以邮件接收时间戳重命名
                    ts = rdt_str.replace(":", "").replace("-", "")[:15]  # e.g. 20250921T103000
                    safe = self._safe_name(name)
                    base, extname = os.path.splitext(safe)
                    filepath = save_path / f"{base}_{ts}{extname}"

                    # 下载内容：优先 contentBytes，否则 $value
                    content_b64 = a.get("contentBytes")
                    if content_b64:
                        content = base64.b64decode(content_b64)
                    else:
                        att_id = a["id"]
                        content = self._download_value(
                            f"{self.GRAPH_BASE}/me/messages/{mid}/attachments/{att_id}/$value",
                            token,
                        )

                    filepath.write_bytes(content)
                    print(f"[SAVE] {filepath}")
                    saved.append((rdt, filepath))
                    found += 1
                    if found >= need_count:
                        break
                if found >= need_count or scanned >= max_scan:
                    break
            url = data.get("@odata.nextLink")

        # 按时间排序（新→旧），仅返回路径
        saved.sort(key=lambda t: t[0], reverse=True)
        out_paths = [p for _, p in saved][:need_count]

        if not out_paths:
            print("[WARN] 未找到匹配附件。请检查关键词/扩展名或增大 days_back。"
                  " / No matching attachments found. Check keywords/extensions or increase days_back.")
        else:
            print(f"[OK] 下载完成，共 {len(out_paths)} 个。目录: {save_path.resolve()} / "
                  f"Download complete, total {len(out_paths)}. Folder: {save_path.resolve()}")
        return out_paths


# =========================
# CLI 入口
# =========================
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Download latest matching attachments via Microsoft Graph (with token cache)."
    )
    parser.add_argument("--tenant-id", required=True)
    parser.add_argument("--client-id", required=True)
    parser.add_argument("--contains", default=None, help="文件名包含关键字（不区分大小写）")
    parser.add_argument("--equals", default=None, help="文件名精确等于（优先级更高）")
    parser.add_argument("--ext", default=None, help="扩展名过滤，如 .xlsx")
    parser.add_argument("--need-count", type=int, default=2)
    parser.add_argument("--days-back", type=int, default=90)
    parser.add_argument("--page-size", type=int, default=50)
    parser.add_argument("--max-scan", type=int, default=800)
    parser.add_argument("--save-dir", default=".")
    parser.add_argument("--token-cache", default="graph_token_cache.json")
    parser.add_argument("--mail-folder", default=None, help="指定文件夹（如 inbox），默认为所有文件夹")
    parser.add_argument("--timeout", type=int, default=60, help="请求超时时间（秒）")

    args = parser.parse_args()

    tool = GraphMailAttachmentTool(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        scopes="Mail.Read offline_access",
        token_cache=args.token_cache,
        request_timeout=args.timeout,
    )

    tool.download_latest_attachments(
        contains=args.contains,
        equals=args.equals,
        ext=args.ext,
        need_count=args.need_count,
        days_back=args.days_back,
        page_size=args.page_size,
        max_scan=args.max_scan,
        save_dir=args.save_dir,
        mail_folder=args.mail_folder,
    )
