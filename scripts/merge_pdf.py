#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 合併腳本 — 將施工記錄PDF + 天氣報告PDF + 潮汐預報PDF合併為完整版
=========================================================================
用途：獨立合併工具，適用於手動生成PDF後的合併場景
依賴：pip install pypdf

使用方法：
    python merge_pdf.py --date 2026-08-02
"""

import os
import sys
import argparse
from pypdf import PdfWriter, PdfReader

BASE_DIR = r"<PROJECTS_ROOT>\[項目名稱]\每日施工記錄"


def merge(date_str):
    """合併施工記錄 + 天氣 + 潮汐為一個PDF"""
    construction_pdf = os.path.join(BASE_DIR, f"{date_str}每日施工記錄表.pdf")
    weather_pdf = os.path.join(BASE_DIR, "附件", "天氣報告", f"天氣報告_{date_str}.pdf")
    tide_pdf = os.path.join(BASE_DIR, "附件", "潮汐預報", f"潮汐預報_{date_str}.pdf")
    output_pdf = os.path.join(BASE_DIR, f"{date_str}每日施工記錄表_完整版.pdf")

    # 檢查文件存在
    files = [
        ("施工記錄PDF", construction_pdf),
        ("天氣報告PDF", weather_pdf),
        ("潮汐預報PDF", tide_pdf),
    ]

    writer = PdfWriter()
    for label, path in files:
        if os.path.exists(path):
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
            print(f"✅ 已添加: {label} ({len(reader.pages)} 頁)")
        else:
            print(f"⚠️  跳過不存在的文件: {label} ({path})")

    with open(output_pdf, 'wb') as f:
        writer.write(f)

    print(f"\n📄 合併完成: {output_pdf}")

    # 可選：替換原始施工記錄PDF
    if os.path.exists(construction_pdf):
        os.remove(construction_pdf)
        os.rename(output_pdf, construction_pdf)
        print(f"🔄 已替換: {construction_pdf}")

    return construction_pdf


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF合併工具")
    parser.add_argument("--date", required=True, help="日期 YYYY-MM-DD")
    args = parser.parse_args()

    merge(args.date)
