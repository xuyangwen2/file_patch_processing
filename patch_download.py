import os
import re
import uuid
import base64
import requests
from urllib.parse import urlparse
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# =========================
# 1. 路径配置
# =========================
desktop     = os.path.join(os.path.expanduser("~"), "Desktop")
titles_path = os.path.join(desktop, "titles.txt")
download_dir = os.path.join(desktop, "SSC-COP_原始格式文件")

os.makedirs(download_dir, exist_ok=True)

# =========================
# 2. AES-GCM 解密（密钥仅通过环境变量提供，勿写入代码库）
#    环境变量：SSC_TITLE_ENCRYPT_AES_KEY_B64 = Base64 编码的 16/24/32 字节密钥
# =========================
def _aes_key_from_env(var_name: str) -> bytes:
    raw = os.environ.get(var_name, "").strip()
    if not raw:
        return b""
    try:
        return base64.b64decode(raw)
    except Exception:
        print(f"⚠️  环境变量 {var_name} 不是有效的 Base64")
        return b""


AES_KEY = _aes_key_from_env("SSC_TITLE_ENCRYPT_AES_KEY_B64")


def decrypt_title(hex_str: str) -> str:
    """从 HEX 字符串解密 AES-GCM 密文，返回明文标题。
    数据格式: [IV 前12字节] + [AES-GCM 密文 + 认证标签]
    """
    if len(AES_KEY) not in (16, 24, 32):
        return ""
    try:
        raw = bytes.fromhex(hex_str)
        nonce = raw[:12]
        ciphertext_tag = raw[12:]
        plaintext = AESGCM(AES_KEY).decrypt(nonce, ciphertext_tag, None)
        return plaintext.decode("utf-8")
    except Exception as e:
        print(f"⚠️  解密失败: {e}")
        return ""


def sanitize_filename(name: str) -> str:
    """移除文件名中不合法的字符"""
    return re.sub(r'[\\/:*?"<>|]', '_', name).strip()


def normalize_plain_title(title: str) -> str:
    """规范化明文标题：仅保留文件名主体（去前缀和后缀）。"""
    if not title:
        return ""

    normalized = title.strip()

    # 先取“看起来像文件名”的最后一段，兼容 / 和 \ 路径分隔
    normalized = normalized.replace("\\", "/").split("/")[-1].strip()

    # 去掉冒号及其前缀（兼容英文/中文冒号）
    colon_index = max(normalized.rfind(":"), normalized.rfind("："))
    if colon_index >= 0 and colon_index + 1 < len(normalized):
        normalized = normalized[colon_index + 1:].strip()

    # 去掉文件后缀（仅命中常见后缀时才移除，避免误删标题中的点号内容）
    removable_exts = {
        ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
        ".txt", ".md", ".rtf", ".csv", ".zip"
    }
    root, ext = os.path.splitext(normalized)
    if ext and ext.lower() in removable_exts:
        normalized = root.strip()

    return normalized


# =========================
# 3. 读取 titles.txt
#    格式（DMS 复制结果）：
#    行1: 空行
#    行2-4: 列头（id / url / title_hex，每行一个）
#    行5+: id\turl\ttitle_hex
# =========================
rows = []   # list of (url, title_hex)

with open(titles_path, encoding="utf-8") as f:
    lines = f.readlines()

for line in lines:
    parts = line.rstrip("\n").split("\t")
    if len(parts) != 3:
        continue            # 跳过空行、列头行等
    _, url, title_hex = parts
    # 过滤掉列头行本身
    if url.strip().lower() == "url":
        continue
    if not url.startswith("http"):
        continue
    rows.append((url.strip(), title_hex.strip()))

print(f"共读取 {len(rows)} 条记录")

# =========================
# 4. Content-Type 映射
# =========================
CONTENT_TYPE_MAP = {
    "application/pdf": ".pdf",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
}

# =========================
# 5. 文件头识别（magic number）
# =========================
def detect_ext_by_magic(content: bytes) -> str:
    if content.startswith(b"%PDF"):
        return ".pdf"
    if content.startswith(b"\xD0\xCF\x11\xE0"):
        return ".doc"   # Word 97-2003
    if content.startswith(b"PK"):
        return ".docx"  # docx / xlsx / zip
    return ".bin"


def resolve_unique_path(save_dir: str, file_name: str) -> tuple[str, bool]:
    """若重名则自动追加 (2)/(3) ...，避免覆盖已存在文件。"""
    base, ext = os.path.splitext(file_name)
    candidate = os.path.join(save_dir, file_name)
    if not os.path.exists(candidate):
        return candidate, False

    index = 2
    while True:
        new_name = f"{base} ({index}){ext}"
        candidate = os.path.join(save_dir, new_name)
        if not os.path.exists(candidate):
            return candidate, True
        index += 1

# =========================
# 6. 下载函数
# =========================
def download_file_keep_format(url: str, save_dir: str, title: str = ""):
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()

        content = resp.content
        size_kb = len(content) / 1024

        # 仅跳过真正的空文件（0 字节）
        if len(content) == 0:
            print(f"⚠️ 跳过空文件: {url}")
            return "skipped_empty"

        # ---------- 后缀判断（逻辑不变）----------
        parsed = urlparse(url)
        url_filename = os.path.basename(parsed.path)
        _, ext = os.path.splitext(url_filename)

        # 1️⃣ URL 无后缀 → Content-Type
        if not ext:
            content_type = resp.headers.get("Content-Type", "").split(";")[0]
            ext = CONTENT_TYPE_MAP.get(content_type, "")

        # 2️⃣ Content-Type 不可靠 → magic number
        if not ext:
            ext = detect_ext_by_magic(content)

        # 3️⃣ 文件名：优先用解密后的标题，否则回退到 UUID
        if title:
            name = sanitize_filename(title)
            if not name:
                name = str(uuid.uuid4())
        else:
            name = str(uuid.uuid4())

        final_name = f"{name}{ext}"
        file_path, renamed_due_to_conflict = resolve_unique_path(save_dir, final_name)
        saved_name = os.path.basename(file_path)

        # ---------- 保存 ----------
        with open(file_path, "wb") as f:
            f.write(content)

        if renamed_due_to_conflict:
            print(f"ℹ️ 文件名冲突，已重命名保存: {saved_name}")
        print(f"✅ 下载成功: {saved_name} ({int(size_kb)} KB)")
        return "renamed" if renamed_due_to_conflict else "success"

    except Exception as e:
        print(f"❌ 下载失败: {url}")
        print(f"   错误原因: {e}")
        return "failed"

# =========================
# 7. 批量下载
# =========================
success_count = 0
renamed_count = 0
skipped_empty_count = 0
failed_count = 0

for url, title_hex in rows:
    title = decrypt_title(title_hex) if title_hex else ""
    title = normalize_plain_title(title)
    status = download_file_keep_format(url, download_dir, title)
    if status == "success":
        success_count += 1
    elif status == "renamed":
        success_count += 1
        renamed_count += 1
    elif status == "skipped_empty":
        skipped_empty_count += 1
    else:
        failed_count += 1

print("\n🎉 所有文件已按【原始格式】导出完成")
print(
    f"统计: 成功 {success_count}（其中重名改名 {renamed_count}），"
    f"空文件跳过 {skipped_empty_count}，失败 {failed_count}"
)
