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
