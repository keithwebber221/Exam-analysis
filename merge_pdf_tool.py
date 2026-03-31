#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF 合併獨立工具
直接掃描個人報告資料夾並合併所有 PDF
"""

import os
import glob

print("=" * 60)
print("  PDF 合併工具")
print("=" * 60)

# 自動尋找個人報告資料夾
pdf_folders = glob.glob("*_個人報告/") + glob.glob("*_個人報告\\")

# 若找不到，列出當前資料夾內容讓用戶選擇
if not pdf_folders:
    print("\n⚠️  未自動找到個人報告資料夾")
    print("當前目錄內容：")
    for item in os.listdir("."):
        if os.path.isdir(item):
            print(f"  [資料夾] {item}")
    folder_name = input("\n請輸入資料夾名稱（例：2526_T1E_個人報告）：").strip()
    pdf_folders = [folder_name]
else:
    print(f"\n✅ 找到 {len(pdf_folders)} 個個人報告資料夾：")
    for i, f in enumerate(pdf_folders, 1):
        pdf_count = len(glob.glob(os.path.join(f, "*_個人報告.pdf")))
        print(f"  {i}. {f}  （{pdf_count} 個 PDF）")

# 若有多個資料夾，讓用戶選擇
if len(pdf_folders) > 1:
    choice = input("\n請輸入要合併的資料夾編號（預設 1）：").strip()
    try:
        idx = int(choice) - 1
    except:
        idx = 0
    pdf_folder = pdf_folders[idx]
else:
    pdf_folder = pdf_folders[0]

print(f"\n📂 使用資料夾：{pdf_folder}")

# 掃描 PDF
pdf_files = sorted(glob.glob(os.path.join(pdf_folder, "*_個人報告.pdf")))

if not pdf_files:
    print(f"❌ 在 {pdf_folder} 中找不到 *_個人報告.pdf 檔案")
    print("\n資料夾內容：")
    for f in os.listdir(pdf_folder):
        print(f"  {f}")
    input("\n按 Enter 退出...")
    exit()

print(f"\n✅ 找到 {len(pdf_files)} 個 PDF：")
for i, f in enumerate(pdf_files, 1):
    print(f"  {i:02d}. {os.path.basename(f)}")

# 確認合併
confirm = input(f"\n確認合併以上 {len(pdf_files)} 個 PDF？(Y/N)：").strip().upper()
if confirm != "Y":
    print("已取消合併。")
    input("按 Enter 退出...")
    exit()

# 匯入 PyPDF2
try:
    from PyPDF2 import PdfMerger
    print("\n▶ 使用 PyPDF2 合併中...")
except ImportError:
    try:
        from pypdf import PdfMerger
        print("\n▶ 使用 pypdf 合併中...")
    except ImportError:
        print("❌ 未安裝 PyPDF2，請執行：pip install PyPDF2")
        input("按 Enter 退出...")
        exit()

# 合併
merger = PdfMerger()
failed = []

for pdf_path in pdf_files:
    try:
        merger.append(pdf_path)
        print(f"  ✅ 已加入：{os.path.basename(pdf_path)}")
    except Exception as e:
        print(f"  ❌ 無法加入：{os.path.basename(pdf_path)} → {e}")
        failed.append(pdf_path)

# 決定輸出路徑（與資料夾同層）
folder_name = os.path.basename(pdf_folder.rstrip("/\\"))
prefix = folder_name.replace("_個人報告", "")
output_path = f"{prefix}_個人報告合併.pdf"

try:
    merger.write(output_path)
    merger.close()

    size_kb = os.path.getsize(output_path) / 1024
    print(f"""
╔══════════════════════════════════════════════╗
║  ✅ 合併完成！                               ║
╠══════════════════════════════════════════════╣
║  📄 檔案：{output_path:<35}║
║  📦 大小：{size_kb:.1f} KB{" "*(35 - len(f"{size_kb:.1f} KB"))}║
║  📋 包含：{len(pdf_files) - len(failed)} 份報告{" "*(31 - len(str(len(pdf_files) - len(failed))))}║
╚══════════════════════════════════════════════╝
""")
    if failed:
        print(f"⚠️  {len(failed)} 份 PDF 未能加入：")
        for f in failed:
            print(f"   - {os.path.basename(f)}")

except Exception as e:
    print(f"\n❌ 寫出合併 PDF 時失敗：{e}")
    import traceback
    traceback.print_exc()

input("\n按 Enter 退出...")
