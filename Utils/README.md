# Utils 说明

本目录包含可复用工具，供 Weekly/M1M2 脚本调用。

## graph_mail_attachment_tool.py
用于从 Microsoft Graph 下载邮件附件并缓存 Token。

流程：
1) 读取本地 token 缓存；有效则直接使用。
2) 无有效 token 则尝试 refresh_token 刷新。
3) 刷新失败时走 Device Code 登录流程获取 token。
4) 拉取邮件列表（可限定邮箱文件夹、天数、分页/扫描上限）。
5) 读取附件元数据，按包含关键字/精确名/扩展名过滤。
6) 下载附件内容（contentBytes 或 $value），按接收时间重命名保存。
7) 返回已保存文件路径列表（新→旧）。

## sql_agent_tool.py
用于触发 SQL Server Agent Job 并等待完成（优先文件检测模式）。

流程：
1) 解析 Job 名（支持模糊匹配）。
2) 解析/校验 start_step（step_id → step_name，或直接 step_name）。
3) 启动 Job（sp_start_job，仅支持 step_name）。
4) 若提供 archive_dir，使用文件检测模式监控输出目录。
5) 文件检测成功即返回；否则超时回退等待模式。
6) 支持蜂鸣提示与打开输出目录。

## email_notify_tool.py
用于通过 SMTP 发送通知邮件（成功/失败等场景）。

流程：
1) 从 email_notify_config.json 读取 SMTP 配置、启用开关与默认/按 job 配置的收件人。
2) 由调用方传入 subject/body（收件人走配置或默认）。
3) 建立 SMTP 连接（可选 TLS + 登录）。
4) 发送邮件。

---

# Utils Notes (EN)

This folder contains reusable helpers used by Weekly/M1M2 scripts.

## graph_mail_attachment_tool.py
Downloads mail attachments via Microsoft Graph with token caching.

Steps:
1) Load local token cache; use if valid.
2) If invalid, try refresh_token.
3) If refresh fails, use Device Code flow to obtain token.
4) Fetch message list (optional folder filter, day range, paging/scan limits).
5) Read attachment metadata and filter by keyword/exact name/extension.
6) Download content (contentBytes or $value) and rename by received time.
7) Return saved file paths (newest → oldest).

## sql_agent_tool.py
Triggers a SQL Server Agent Job and waits for completion (file-watch first).

Steps:
1) Resolve job name (supports fuzzy matching).
2) Resolve/validate start_step (step_id → step_name, or step_name directly).
3) Start job (sp_start_job supports step_name only).
4) If archive_dir is provided, use file-watch mode on output folder.
5) Return on file detection; otherwise fall back to timeout wait.
6) Optional beep + open output folder.

## email_notify_tool.py
Sends notification emails via SMTP (success/failure, etc.).

Steps:
1) Load SMTP config, enable flag, and recipients from email_notify_config.json (default + per job overrides).
2) Caller provides subject/body (recipients resolved from config or defaults).
3) Connect to SMTP server (optional TLS + login).
4) Send email.
