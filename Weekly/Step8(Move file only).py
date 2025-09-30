# -*- coding: utf-8 -*-
"""
Supplier – Weekly automation (仅第一步：寻找并复制文件)
"""

import os
import re
import shutil
import datetime as dt
from glob import glob

# ================== CONFIG ==================
BASE_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\Supplier SUBCON Performance\Supplier"
GLOB_PATTERN = "Supplier - KPIs Review (PO GR) W*'*.xlsx"  # 文件命名格式

# 若想手动指定周号：设置环境变量 SUPPLIER_WEEK=38
ENV_WEEK_OVERRIDE = "SUPPLIER_WEEK"

# 业务口径：比 ISO 周慢 1 周
WEEK_OFFSET = -1
# ======================================================


# -------------------- week helpers --------------------
def compute_week_token(today: dt.date | None = None) -> str:
    """返回当周 token：WXX'YY（支持 ISO 偏移与手动覆盖）"""
    manual = os.getenv(ENV_WEEK_OVERRIDE)
    d = today or dt.date.today()
    if manual and manual.isdigit():
        w = int(manual)
        y = d.year
    else:
        y, w, _ = d.isocalendar()
        w = w + WEEK_OFFSET
        # 越界处理
        if w <= 0:
            last_dec_28 = dt.date(y - 1, 12, 28)
            _, w_last, _ = last_dec_28.isocalendar()
            w = w_last + w
            y = y - 1
        else:
            last_dec_28 = dt.date(y, 12, 28)
            _, w_last, _ = last_dec_28.isocalendar()
            if w > w_last:
                w = w - w_last
                y = y + 1
    return f"W{w:02d}'{str(y)[-2:]}"


# -------------------- file ops --------------------
def find_latest_lastweek_file() -> str:
    cands = [f for f in glob(os.path.join(BASE_DIR, GLOB_PATTERN)) if os.path.isfile(f)]
    if not cands:
        raise FileNotFoundError(f"在 {BASE_DIR} 未找到匹配文件：{GLOB_PATTERN}")
    cands.sort(key=os.path.getmtime, reverse=True)
    latest = cands[0]
    print(f"  ✔ 找到最近文件：{os.path.basename(latest)}")
    return latest


def make_this_week_name(from_name: str, wyy: str) -> str:
    base, ext = os.path.splitext(from_name)
    if re.search(r"W\d{1,2}'\d{2}", base, flags=re.I):
        new = re.sub(r"W\d{1,2}'\d{2}", wyy, base, flags=re.I)
    else:
        new = f"{base} {wyy}"
    return new + ext


def copy_to_this_week(latest_path: str, wyy: str) -> str:
    src_name = os.path.basename(latest_path)
    dst_name = make_this_week_name(src_name, wyy)
    dst_path = os.path.join(BASE_DIR, dst_name)
    if os.path.abspath(dst_path) == os.path.abspath(latest_path):
        print("  ⚠ 最近文件已经是本周命名，无需复制。")
        return latest_path
    shutil.copy2(latest_path, dst_path)
    print(f"  ✔ 已复制为本周文件：{dst_name}")
    return dst_path


# -------------------- main --------------------
def main():
    print("==== Supplier – Weekly (仅复制) ====")
    wyy = compute_week_token()
    print(f"本周周标：{wyy}")

    print("Step 1 | 寻找并复制上周文件 …")
    last_path = find_latest_lastweek_file()
    cur_path = copy_to_this_week(last_path, wyy)

    print(f"✅ 完成：生成文件 {cur_path}")


if __name__ == "__main__":
    main()
