# -*- coding: utf-8 -*-
"""
نقطهٔ ورود برای ساخت فایل اجرایی (exe).
- اگر فایل‌ها را روی برنامه بکشید (drag & drop) یا کنار آن باشند، همان‌ها پردازش می‌شوند.
- بدون آرگومان: همه فایل‌های .xlsx کنار برنامه پردازش می‌شوند.
- گزارش در همان پوشه با نام pos_monitoring_report.xlsx ساخته می‌شود.
"""
import glob
import os
import sys
import traceback
from pathlib import Path

import pos_analysis as pa

REPORT_NAME = "pos_monitoring_report.xlsx"


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def pick_files(args):
    if args:
        return [a for a in args if a.lower().endswith(".xlsx")]
    return sorted(glob.glob(str(app_dir() / "*.xlsx")))


def main():
    folder = app_dir()
    files = [f for f in pick_files(sys.argv[1:]) if Path(f).name != REPORT_NAME]
    try:
        if not files:
            print("هیچ فایل .xlsx پیدا نشد. فایل‌ها را کنار برنامه بگذارید یا روی آن بکشید.")
        else:
            print(f"{len(files)} فایل پیدا شد:")
            for f in files:
                print("  -", os.path.basename(f))
            pa.run(files, str(folder / REPORT_NAME))
            print("\nگزارش در این پوشه ذخیره شد:", folder / REPORT_NAME)
    except Exception:
        print("\nخطا رخ داد:")
        traceback.print_exc()
    input("\nبرای خروج Enter را بزنید...")


if __name__ == "__main__":
    main()
