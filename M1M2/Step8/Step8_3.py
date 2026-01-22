# -*- coding: utf-8 -*-
import os
import glob
import shutil
import subprocess

RAW_DIR = r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\External\M1M2\Original Raw"
DST_DIR = r"\\mp1do4ce0373ndz\Customs"
ARCHIVE_DIR = r"\\mp1do4ce0373ndz\Customs\Archive"

KEY_M1 = "M1 Final Monthly Raw File"
KEY_M2_ARCHIVE = "M2"                      # Archive 中匹配关键字
KEY_M2_RAW = "M2 Final Monthly Raw File"   # Raw 中匹配关键字

EXTS = ("xlsx", "xlsm", "xls")

def latest_match(folder, keyword, exts=EXTS):
    files = []
    for ext in exts:
        files.extend(glob.glob(os.path.join(folder, f"*{keyword}*.{ext}")))
    if not files:
        raise FileNotFoundError(f"未在 {folder} 找到包含“{keyword}”的文件 / No file containing '{keyword}' in {folder}")
    return max(files, key=os.path.getmtime)

def safe_replace(target_path):
    """若目标文件存在且可能被占用，尽力删除；删不了则先挪成 .bak。"""
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except Exception:
            bak = target_path + ".bak"
            try:
                if os.path.exists(bak):
                    os.remove(bak)
            except Exception:
                pass
            try:
                os.replace(target_path, bak)
            except Exception:
                pass

def copy_as_name(src_path, dst_dir, base_name):
    os.makedirs(dst_dir, exist_ok=True)
    ext = os.path.splitext(src_path)[1].lower()
    dst_path = os.path.join(dst_dir, base_name + ext)
    safe_replace(dst_path)
    shutil.copy2(src_path, dst_path)
    print(f"[COPY] {os.path.basename(src_path)} → {dst_path}")
    return dst_path

def move_to_dir(src_path, dst_dir):
    """把 src 移动到目标目录，文件名保持原名。返回移动后的新路径。"""
    os.makedirs(dst_dir, exist_ok=True)
    dst_path = os.path.join(dst_dir, os.path.basename(src_path))
    safe_replace(dst_path)
    shutil.move(src_path, dst_path)  # 跨盘会自动走 copy+delete
    print(f"[MOVE] {src_path} → {dst_path}")
    return dst_path

def rename_in_place_as_name(src_path, base_name):
    """在原目录内改名为 base_name + 原扩展名；返回新路径。"""
    folder = os.path.dirname(src_path)
    ext = os.path.splitext(src_path)[1].lower()
    new_path = os.path.join(folder, base_name + ext)
    if os.path.abspath(src_path).lower() != os.path.abspath(new_path).lower():
        safe_replace(new_path)
        os.replace(src_path, new_path)
        print(f"[RENAME] {src_path} → {new_path}")
    else:
        print(f"[RENAME] 已是目标名：{new_path} / Already target name")
    return new_path

def open_in_excel(path):
    try:
        os.startfile(path)  # Windows
    except AttributeError:
        subprocess.run(["cmd", "/c", "start", "", path], shell=True)
    except OSError:
        subprocess.run(["cmd", "/c", "start", "", path], shell=True)
    print(f"[OPEN] {path}")

def main():
    # --- 1) 处理 M1（复制并命名为 M1） ---
    try:
        m1_src = latest_match(RAW_DIR, KEY_M1, EXTS)
        print(f"[M1] 发现最新：{os.path.basename(m1_src)} / Latest found")
        m1_dst = copy_as_name(m1_src, DST_DIR, "M1")
    except Exception as e:
        print(f"[M1][错误] {e} / Error: {e}")

    # --- 2) 处理 M2（从 Archive 复制到 Customs，命名为 M2.*；并打开“源+目标”两个文件） ---
    try:
        m2_arch = latest_match(ARCHIVE_DIR, KEY_M2_ARCHIVE, EXTS)
        print(f"[M2-Archive] 发现最新：{os.path.basename(m2_arch)} / Latest found")
        # 复制为固定名 M2.*（不再移动/改名）
        m2_dst = copy_as_name(m2_arch, DST_DIR, "M2")
        # 打开目标（Customs\M2.*）与源（Archive\原文件名）
        open_in_excel(m2_dst)
        open_in_excel(m2_arch)
    except Exception as e:
        print(f"[M2-Archive][警告] {e} / Warning: {e}")

    # --- 3) 打开 Raw 里的 M2 供手动对拷 ---
    try:
        m2_raw = latest_match(RAW_DIR, KEY_M2_RAW, EXTS)
        print(f"[M2-Raw] 发现最新：{os.path.basename(m2_raw)} / Latest found")
        open_in_excel(m2_raw)
    except Exception as e:
        print(f"[M2-Raw][警告] {e} / Warning: {e}")

    # 完成提示
    try:
        import winsound
        winsound.Beep(1200, 180); winsound.Beep(1600, 180); winsound.Beep(2000, 220)
    except Exception:
        print("\a")
    print("✅ 全部完成。 / All done.")

if __name__ == "__main__":
    main()
