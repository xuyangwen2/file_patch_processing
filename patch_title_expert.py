import os
import re
import base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# =========================
# 1. 路径配置
# =========================
desktop     = os.path.join(os.path.expanduser("~"), "Desktop")
titles_path = os.path.join(desktop, "bitmain-COP.txt")
output_names_path = os.path.join(desktop, "bitmain-COP_文件名列表.txt")

# =========================
# 2. AES-GCM 解密（密钥仅通过环境变量 BITMAIN_TITLE_ENCRYPT_AES_KEY_B64 提供）
# =========================
def _aes_key_from_env(var_name: str) -> bytes:
    raw = os.environ.get(var_name, "").strip()
    if not raw:
        return b""
    try:
        return base64.b64decode(raw)
    except Exception:
        print(f"解密失败: 环境变量 {var_name} 不是有效的 Base64")
        return b""


AES_KEY = _aes_key_from_env("BITMAIN_TITLE_ENCRYPT_AES_KEY_B64")


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
        print(f"解密失败: {e}")
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
# 3. 读取 bitmain-COP.txt
#    兼容两种格式：
#    A) id\turl\ttitle_hex
#    B) id\ttitle_hex
# =========================
rows = []   # list of (url, title_hex)

with open(titles_path, encoding="utf-8") as f:
    lines = f.readlines()

for line in lines:
    parts = line.rstrip("\n").split("\t")
    if len(parts) == 3:
        _, url, title_hex = parts
        # 过滤掉列头行本身
        if url.strip().lower() == "url":
            continue
        if not url.startswith("http"):
            continue
        rows.append((url.strip(), title_hex.strip()))
    elif len(parts) == 2:
        _, title_hex = parts
        # 过滤掉列头行本身
        if title_hex.strip().lower() in {"hex(title_encrypt)", "title_hex"}:
            continue
        rows.append(("", title_hex.strip()))
    else:
        continue            # 跳过空行、列头行等

print(f"共读取 {len(rows)} 条记录")

# =========================
# 4. 仅生成文件名（不下载、不保存实际文件）
# =========================
def build_filename(url: str, title: str = "") -> str:
    # 优先使用解密标题；为空时回退到 URL 文件名主体
    name = sanitize_filename(title) if title else ""
    url_filename = os.path.basename(url.split("?", 1)[0])
    url_root, ext = os.path.splitext(url_filename)
    if not name:
        name = sanitize_filename(url_root) if url_root else "untitled"
    return f"{name}{ext}" if ext else name


# =========================
# 5. 批量生成文件名并写入 txt
# =========================
all_names = []

for url, title_hex in rows:
    title = decrypt_title(title_hex) if title_hex else ""
    title = normalize_plain_title(title)
    if not title and not url and title_hex:
        # 无 URL 且解密失败时，使用密文前缀兜底，确保可导出
        title = f"title_{title_hex[:24]}"
    final_name = build_filename(url, title)
    all_names.append(final_name)
    print(f"文件名: {final_name}")

with open(output_names_path, "w", encoding="utf-8") as f:
    for name in all_names:
        f.write(f"{name}\n")

print(f"\n文件名列表已生成: {output_names_path}")
