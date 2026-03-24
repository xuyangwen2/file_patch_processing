import os
import uuid
import pandas as pd
import requests
from urllib.parse import urlparse

# =========================
# 1. 路径配置
# =========================
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
excel_path = os.path.join(desktop, "SSC-COP文件备份.xlsx")
download_dir = os.path.join(desktop, "SSC-COP_原始格式文件")

os.makedirs(download_dir, exist_ok=True)

# =========================
# 2. 读取 Excel
# =========================
df = pd.read_excel(excel_path)

if "url" not in df.columns:
    raise ValueError("Excel 中必须包含 url 列")

urls = df["url"].dropna().tolist()

# =========================
# 3. Content-Type 映射
# =========================
CONTENT_TYPE_MAP = {
    "application/pdf": ".pdf",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
}

# =========================
# 4. 文件头识别（magic number）
# =========================
def detect_ext_by_magic(content: bytes) -> str:
    if content.startswith(b"%PDF"):
        return ".pdf"
    if content.startswith(b"\xD0\xCF\x11\xE0"):
        return ".doc"   # Word 97-2003
    if content.startswith(b"PK"):
        return ".docx"  # docx / xlsx / zip
    return ".bin"

# =========================
# 5. 下载函数
# =========================
def download_file_keep_format(url: str, save_dir: str):
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()

        content = resp.content
        size_kb = len(content) / 1024

        # ✅ 跳过 0KB 或异常小文件
        if size_kb <= 1:
            print(f"⚠️ 跳过空文件: {url}")
            return

        # ---------- 文件名与后缀 ----------
        parsed = urlparse(url)
        filename = os.path.basename(parsed.path)
        name, ext = os.path.splitext(filename)

        # 1️⃣ URL 无后缀 → Content-Type
        if not ext:
            content_type = resp.headers.get("Content-Type", "").split(";")[0]
            ext = CONTENT_TYPE_MAP.get(content_type, "")

        # 2️⃣ Content-Type 不可靠 → magic number
        if not ext:
            ext = detect_ext_by_magic(content)

        # 3️⃣ 仍无文件名 → UUID
        if not name:
            name = str(uuid.uuid4())

        final_name = f"{name}{ext}"
        file_path = os.path.join(save_dir, final_name)

        # ---------- 保存 ----------
        with open(file_path, "wb") as f:
            f.write(content)

        print(f"✅ 下载成功: {final_name} ({int(size_kb)} KB)")

    except Exception as e:
        print(f"❌ 下载失败: {url}")
        print(f"   错误原因: {e}")

# =========================
# 6. 批量下载
# =========================
for url in urls:
    download_file_keep_format(url, download_dir)

print("\n🎉 所有文件已按【原始格式】导出完成")
