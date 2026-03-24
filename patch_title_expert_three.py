import os
import re
import base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# =========================
# 1. 路径配置
# =========================
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
txt_paths = {
    "SSC": os.path.join(desktop, "ssc-知识集合.txt"),
    "BITMAIN": os.path.join(desktop, "bitmain-知识集合.txt"),
    "SOPHGO": os.path.join(desktop, "sophgo-知识集合.txt"),
}
output_names_path = os.path.join(desktop, "三方文件名列表.txt")

# =========================
# 2. AES-GCM 解密（密钥仅通过环境变量提供，勿写入代码库）
#    SSC_TITLE_ENCRYPT_AES_KEY_B64 / BITMAIN_TITLE_ENCRYPT_AES_KEY_B64 / SOPHGO_TITLE_ENCRYPT_AES_KEY_B64
# =========================
def _aes_key_from_env(var_name: str) -> bytes:
    raw = os.environ.get(var_name, "").strip()
    if not raw:
        return b""
    try:
        return base64.b64decode(raw)
    except Exception:
        print(f"⚠️ 环境变量 {var_name} 不是有效的 Base64")
        return b""


AES_KEYS = {
    "SSC": _aes_key_from_env("SSC_TITLE_ENCRYPT_AES_KEY_B64"),
    "BITMAIN": _aes_key_from_env("BITMAIN_TITLE_ENCRYPT_AES_KEY_B64"),
    "SOPHGO": _aes_key_from_env("SOPHGO_TITLE_ENCRYPT_AES_KEY_B64"),
}


def decrypt_title(hex_str: str, aes_key: bytes) -> str:
    """从 HEX 字符串解密 AES-GCM 密文，返回明文标题。
    数据格式: [IV 前12字节] + [AES-GCM 密文 + 认证标签]
    """
    if len(aes_key) not in (16, 24, 32):
        return ""
    try:
        if not isinstance(hex_str, str) or not hex_str.strip():
            return ""
        raw = bytes.fromhex(hex_str.strip())
        nonce = raw[:12]
        ciphertext_tag = raw[12:]
        plaintext = AESGCM(aes_key).decrypt(nonce, ciphertext_tag, None)
        return plaintext.decode("utf-8")
    except Exception as e:
        print(f"⚠️ 解密失败: {e}")
        return ""


def sanitize_filename(name: str) -> str:
    """移除文件名中不合法的字符"""
    return re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip()


def normalize_plain_title(title: str) -> str:
    """规范化明文标题：仅保留文件名主体（去前缀和后缀）。"""
    if not title:
        return ""

    normalized = title.strip()
    normalized = normalized.replace("\\", "/").split("/")[-1].strip()

    colon_index = max(normalized.rfind(":"), normalized.rfind("："))
    if colon_index >= 0 and colon_index + 1 < len(normalized):
        normalized = normalized[colon_index + 1:].strip()

    removable_exts = {
        ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx",
        ".txt", ".md", ".rtf", ".csv", ".zip",
    }
    root, ext = os.path.splitext(normalized)
    if ext and ext.lower() in removable_exts:
        normalized = root.strip()

    return normalized


def build_filename(title_plain: str) -> str:
    """组装最终文件名。"""
    name = sanitize_filename(title_plain)
    if not name:
        return ""
    return name


def _parse_line_to_id_and_hex(line: str) -> tuple[str, str]:
    """
    将一行解析为 (id, hex_cipher)。
    兼容分隔符：tab / 逗号 / 连续空白。
    """
    raw = line.strip()
    if not raw:
        return "", ""

    if "\t" in raw:
        parts = raw.split("\t", 1)
    elif "," in raw:
        parts = raw.split(",", 1)
    else:
        parts = raw.split(None, 1)

    if len(parts) != 2:
        return "", ""

    row_id = parts[0].strip()
    hex_cipher = parts[1].strip()
    return row_id, hex_cipher


def _is_header_line(row_id: str, hex_cipher: str) -> bool:
    row_id_low = row_id.lower()
    hex_low = hex_cipher.lower()
    return (
        row_id_low in {"id", "doc_id"}
        or "title_encrypt" in hex_low
        or "title_encrpt" in hex_low
        or hex_low == "hex"
    )


def load_names_from_txt(txt_path: str, source_name: str) -> list[str]:
    """从 txt 读取两列数据（id + hex(title_encrypt)）并生成文件名列表。"""
    if not os.path.exists(txt_path):
        print(f"⚠️ 文件不存在: {txt_path}")
        return []

    aes_key = AES_KEYS.get(source_name)
    if len(aes_key) not in (16, 24, 32):
        print(f"⚠️ 未配置有效密钥（环境变量）: {source_name}")
        return []

    result = []
    with open(txt_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for index, line in enumerate(lines):
        row_id, title_hex = _parse_line_to_id_and_hex(line)
        if not row_id and not title_hex:
            continue
        if index == 0 and _is_header_line(row_id, title_hex):
            continue

        title_plain = normalize_plain_title(decrypt_title(title_hex, aes_key))
        if not title_plain:
            print(f"⚠️ 跳过（解密或标题为空）: id={row_id}")
            continue

        final_name = build_filename(title_plain)
        if final_name:
            result.append(final_name)

    return result


def write_sections_txt(output_path: str, sections: dict[str, list[str]]) -> None:
    """将三组文件名写入同一个 txt，分三段展示。"""
    with open(output_path, "w", encoding="utf-8") as f:
        total_sections = len(sections)
        for index, (section_name, names) in enumerate(sections.items(), start=1):
            f.write(f"【{section_name} 文件名】\n")
            if names:
                for name in names:
                    f.write(f"{name}\n")
            else:
                f.write("(无数据)\n")

            if index < total_sections:
                f.write("\n")


def main() -> None:
    ssc_names = load_names_from_txt(txt_paths["SSC"], "SSC")
    bitmain_names = load_names_from_txt(txt_paths["BITMAIN"], "BITMAIN")
    sophgo_names = load_names_from_txt(txt_paths["SOPHGO"], "SOPHGO")

    sections = {
        "SSC": ssc_names,
        "BITMAIN": bitmain_names,
        "SOPHGO": sophgo_names,
    }
    write_sections_txt(output_names_path, sections)

    print(f"\n🎉 文件名列表已生成: {output_names_path}")
    print(f"SSC: {len(ssc_names)} 条")
    print(f"BITMAIN: {len(bitmain_names)} 条")
    print(f"SOPHGO: {len(sophgo_names)} 条")


if __name__ == "__main__":
    main()
