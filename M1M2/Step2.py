# -*- coding: utf-8 -*-
import csv, shutil, datetime as dt
from pathlib import Path
import xlsxwriter  # pip install XlsxWriter (离线whl)
import winsound

# 固定路径
CSV_PATH = Path(r"C:\Users\70731224\Downloads\Sheet 123_完整数据_data.csv")
NET_XLSX = Path(r"\\mygbynbyn1msis1\Supply-Chain-Analytics\Data Warehouse\Data Source\External\M1M2\Original Raw\Inv_Tracker.xlsx")
LOCAL_XLSX = Path(r".\out\Inv_Tracker.xlsx")
ARCHIVE_DIR = Path(r"C:\Users\70731224\Downloads\_processed")

TEXT_COLS = {"lot id", "material", "material 1", "material num"}
norm = lambda s: (s or "").strip().lower().replace("_", " ")

def open_csv_safely(p: Path):
    for enc in ("utf-8-sig", "utf-8", "cp936", "cp1252"):
        try:
            f = open(p, "r", encoding=enc, newline="")
            _ = f.readline(); f.seek(0)
            return f
        except UnicodeDecodeError:
            continue
    return open(p, "r", encoding="utf-8", errors="replace", newline="")

def main():
    if not CSV_PATH.exists():
        raise SystemExit(f"[ERR] 未找到文件：{CSV_PATH} / File not found: {CSV_PATH}")

    # 1) 先写到本地 SSD（减少网络盘写入开销）
    LOCAL_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(LOCAL_XLSX), {'constant_memory': True})
    ws = wb.add_worksheet('Inv_Tracker')

    # 列文本格式（一次性设置，避免逐格设格式的巨大开销）
    text_fmt = wb.add_format({'num_format': '@'})

    f = open_csv_safely(CSV_PATH)
    rdr = csv.reader(f)

    try:
        header = next(rdr)
    except StopIteration:
        f.close(); wb.close()
        raise SystemExit("[ERR] CSV 为空 / CSV is empty")

    # 写表头
    ws.write_row(0, 0, header)

    # 找到需要按文本写入的列索引，并给整列套上文本格式
    name2idx = {norm(c): i for i, c in enumerate(header)}
    text_idx = sorted(i for n, i in name2idx.items() if n in TEXT_COLS)
    for i in text_idx:
        ws.set_column(i, i, None, text_fmt)  # 整列应用文本格式

    # 逐行写数据（csv.reader 给的都是字符串，默认不会被 XlsxWriter 自动转数字）
    r = 1
    for row in rdr:
        ws.write_row(r, 0, row)  # 不传格式 -> 继承列格式（文本列会是'@'）
        r += 1
        if r % 200000 == 0:
            print(f"[INFO] 已写入 {r-1:,} 行... / {r-1:,} rows written...")

    f.close()
    wb.close()
    print(f"[OK] 本地生成：{LOCAL_XLSX} / Local file created: {LOCAL_XLSX}")

    # 2) 复制到网络盘（一次性复制通常比边写边传快）
    try:
        NET_XLSX.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(LOCAL_XLSX), str(NET_XLSX))
        print(f"[OK] 已复制到网络盘：{NET_XLSX} / Copied to network share: {NET_XLSX}")
    except Exception as e:
        print(f"[WARN] 无法复制到网络盘（{e}），请手动从 {LOCAL_XLSX} 复制。 / "
              f"Failed to copy to network share ({e}); please copy from {LOCAL_XLSX} manually.")

    # 3) 归档原 CSV
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = ARCHIVE_DIR / f"{CSV_PATH.stem}_{ts}{CSV_PATH.suffix}"
    try:
        CSV_PATH.replace(archived)
        print(f"[OK] 原 CSV 已移动到：{archived} / Original CSV moved to: {archived}")
    except Exception as e:
        print(f"[WARN] 移动 CSV 失败（{e}） / Failed to move CSV ({e})")

if __name__ == "__main__":
    main()
    winsound.Beep(1000, 500)  # 1000Hz，响0.5秒
    winsound.Beep(1500, 500)
    winsound.Beep(2000, 500)
