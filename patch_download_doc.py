import os
import pandas as pd
import requests
from urllib.parse import urlparse
import uuid

# =========================
# 1. 路径配置
# =========================
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
excel_path = os.path.join(desktop_path, "SSC-COP文件备份.xlsx")
download_dir = os.path.join(desktop_path, "SSC-COP_DOC文件")

os.makedirs(download_dir, exist_ok=True)

# =========================
# 2. 读取 Excel
# =========================
df = pd.read_excel(excel_path)

if "url" not in df.columns:
    raise ValueError("Excel 中未找到 url 字段")

urls = df["url"].dropna().tolist()

# =========================
# 3. 下载函数
# =========================
def download_doc(url, save_dir):
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()

        content = r.content
        size_kb = len(content) / 1024

        # ✅ 跳过 0KB 文件
        if size_kb <= 1:
            print(f"⚠️ 跳过空文件: {url}")
            return

        # ✅ 处理文件名
        parsed = urlparse(url)
        filename = os.path.basename(parsed.path)

        # 没有后缀 or 后缀不是 doc/docx → 强制 docx
        if not filename.lower().endswith((".doc", ".docx")):
            filename = f"{uuid.uuid4()}.docx"

        file_path = os.path.join(save_dir, filename)

        with open(file_path, "wb") as f:
            f.write(content)

        print(f"✅ 下载成功: {filename} ({int(size_kb)} KB)")

    except Exception as e:
        print(f"❌ 下载失败: {url}")
        print(f"   错误原因: {e}")

# =========================
# 4. 批量下载
# =========================
for url in urls:
    download_doc(url, download_dir)

print("\n🎉 下载完成，0KB 文件已全部跳过")
