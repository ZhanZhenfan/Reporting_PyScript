import os
import sys
import shutil
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from Utils.graph_mail_attachment_tool import GraphMailAttachmentTool
from Utils.sql_agent_tool import SqlAgentTool

SRC_DIR = r"\\mp1do4ce0373ndz\C\WeeklyRawFile\Download_From_Eamil"   # 你截图里的目录名我按“Eamil”写的
SHARE_DIR = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\SAP\Transactional Data\MRP Waterfall"

# 要在作业执行完后打开的 Excel 文件
EXCEL_MACRO_PATH = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Temp Report\03 - MY0X MRP_NEW_WATERFALL_Master - button.xlsm"

# 1) 先下载邮件附件到本地目录
down = GraphMailAttachmentTool(
    tenant_id="5c2be51b-4109-461d-a0e7-521be6237ce2",
    client_id="09004044-1c60-48e5-b1eb-bb42b3892006"
)
downloaded_paths = down.download_latest_attachments(
    contains="ZMRP_WATERFALL_Run",
    ext=".xlsx",
    need_count=2,
    days_back=5,
    save_dir=SRC_DIR,           # 直接用你后续脚本的源目录
    mail_folder="inbox",        # 可不填；想限定收件箱就留着
)

print("[INFO] 下载到：", [p.name for p in downloaded_paths])


# ------------ 工具函数 ------------

def most_recent_monday(today=None):
    """返回本周一（如果今天是周一就取今天）的日期对象"""
    today = today or datetime.today()
    # Python: Monday=0 ... Sunday=6
    return today - timedelta(days=today.weekday())


def pick_big_small(files):
    """按文件大小挑出(大, 小)两个文件"""
    sizes = [(f, os.path.getsize(f)) for f in files]
    sizes.sort(key=lambda x: x[1], reverse=True)
    return sizes[0][0], sizes[1][0]


# ------------ 主逻辑 ------------

def main():
    src = Path(SRC_DIR)
    if not src.is_dir():
        print(f"[ERROR] 源目录不存在: {SRC_DIR}")
        sys.exit(1)

    # 1) 找到两份 ZMRP_WATERFALL_Run*.xlsx
    candidates = list(src.glob("ZMRP_WATERFALL_Run*.xlsx"))
    if len(candidates) < 2:
        print(f"[ERROR] 没找到两份文件，当前匹配到 {len(candidates)}: {[f.name for f in candidates]}")
        sys.exit(1)

    # 只取最新的两份（按修改时间降序）
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    files_to_use = candidates[:2]

    big_file, small_file = pick_big_small(files_to_use)
    print(f"[INFO] 大文件: {Path(big_file).name}  ({os.path.getsize(big_file):,} bytes)")
    print(f"[INFO] 小文件: {Path(small_file).name}  ({os.path.getsize(small_file):,} bytes)")

    # 2) 读取并上下拼接：小文件去掉第一行
    #   - 默认取第一个工作表；保留大文件的列顺序
    df_big = pd.read_excel(big_file, sheet_name=0, dtype=object, engine="openpyxl")

    # 小文件：把第一行当普通数据读进来，然后再去掉第一行
    df_small_raw = pd.read_excel(small_file, sheet_name=0, dtype=object, header=None, engine="openpyxl")

    # 去掉小文件首行（标题行），保留剩余数据
    df_small_no_header = df_small_raw.iloc[1:].copy()

    # 给小文件加上大文件的列名
    df_small_no_header.columns = df_big.columns

    # 拼接
    merged = pd.concat([df_big, df_small_no_header], ignore_index=True)

    # 3) 生成目标文件名（用本周一）
    monday = most_recent_monday()
    out_name = f"New MY0X ZMRP_WATERFALL_Run{monday.strftime('%Y%m%d')}.xlsx"
    out_path = src / out_name

    # 写出到本地源目录
    merged.to_excel(out_path, index=False)
    print(f"[OK] 已保存合并文件: {out_path}")

    # 复制到共享盘（若无权限或网络不可达会抛错）
    share_target = Path(SHARE_DIR) / out_name
    # 确保共享目录存在
    if not Path(SHARE_DIR).exists():
        print(f"[ERROR] 共享路径不可访问: {SHARE_DIR}")
        sys.exit(1)

    shutil.copy2(out_path, share_target)
    print(f"[OK] 已复制到共享盘: {share_target}")

    # 4) 触发 SQL Agent Job
    tool = SqlAgentTool(server="tcp:10.80.127.71,1433")
    result = tool.run_job(
        job_name="Lumileds BI - SC MRP Waterfall",  # 用完整精确名最稳妥
        archive_dir=r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\SAP\Transactional Data\MRP Waterfall\Archive",
        timeout=1800,
        poll_interval=3,
        fuzzy=False,  # 若你 later 拿到读 sysjobs 的权限，可改 True
    )
    print(result)

    # 5) SQL 作业完成后，打开 Excel 宏文件
    try:
        if os.path.exists(EXCEL_MACRO_PATH):
            print(f"[INFO] 正在打开 Excel 文件: {EXCEL_MACRO_PATH}")
            os.startfile(EXCEL_MACRO_PATH)
        else:
            print(f"[WARN] 找不到要打开的 Excel 文件: {EXCEL_MACRO_PATH}")
    except Exception as e:
        print(f"[ERROR] 打开 Excel 文件失败: {e}")


if __name__ == "__main__":
    main()
