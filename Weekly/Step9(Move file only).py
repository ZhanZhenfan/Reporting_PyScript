# -*- coding: utf-8 -*-
"""
Subcon – 仅复制重命名并用 Excel 打开 (无需 win32)
"""

import os
import re
import shutil
import datetime as dt
from glob import glob

# ======== 配置 ========
BASE_DIR = r"\\mygbynbyn1msis2\SCM_Excellence\Weekly Report\Supplier SUBCON Performance\SUBCON"
PATTERN_CHINA    = "China SUBCON - KPIs Review (PO GR) - W*'*(First AB).xlsx"
PATTERN_NONCHINA = "Non China SUBCON - KPIs Review (PO GR) - W*'*(First AB).xlsx"

WEEK_OFFSET   = -1   # 业务周 = ISO 周 - 1
# ======================

def compute_week_token(today: dt.date | None = None) -> str:
    """返回 'W38'25' 这种格式"""
    d = today or dt.date.today()
    y, w, _ = d.isocalendar()
    w += WEEK_OFFSET
    if w <= 0:
        last_dec_28 = dt.date(y - 1, 12, 28)
        _, w_last, _ = last_dec_28.isocalendar()
        w = w_last + w
        y -= 1
    return f"W{w:02d}'{str(y)[-2:]}"

def find_latest(pattern: str) -> str:
    cands = [f for f in glob(os.path.join(BASE_DIR, pattern)) if os.path.isfile(f)]
    if not cands:
        raise FileNotFoundError(f"未找到匹配文件：{pattern}")
    cands.sort(key=os.path.getmtime, reverse=True)
    return cands[0]

def make_this_week_name(from_name: str, wyy: str) -> str:
    base, ext = os.path.splitext(from_name)
    if re.search(r"W\d{1,2}'\d{2}", base, flags=re.I):
        base = re.sub(r"W\d{1,2}'\d{2}", wyy, base, flags=re.I)
    else:
        base = f"{base} {wyy}"
    return base + ext

def copy_to_this_week(latest_path: str, wyy: str) -> str:
    dst = os.path.join(BASE_DIR, make_this_week_name(os.path.basename(latest_path), wyy))
    if os.path.abspath(dst) == os.path.abspath(latest_path):
        print("⚠ 已经是本周命名，无需复制：", os.path.basename(dst))
        return latest_path
    shutil.copy2(latest_path, dst)
    print(f"✔ 已复制: {os.path.basename(dst)}")
    return dst

def open_with_excel(path: str):
    print(f"📂 打开文件: {path}")
    os.startfile(path)  # 会调用系统默认 Excel 打开

def main():
    print("==== Subcon – 仅复制重命名并打开 (无 win32) ====")
    wyy = compute_week_token()
    print("本周标识:", wyy)

    p_ch = copy_to_this_week(find_latest(PATTERN_CHINA), wyy)
    p_nc = copy_to_this_week(find_latest(PATTERN_NONCHINA), wyy)

    open_with_excel(p_ch)
    open_with_excel(p_nc)

    print("\n✅ 文件已复制并在 Excel 打开，后续请手动操作。")

if __name__ == "__main__":
    main()
