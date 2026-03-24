"""
OCR 提取单元测试
测试目标：PDF 修订记录提取（pdfplumber 文字层 + pymupdf OCR）
"""
import os
import re
import sys

DESKTOP       = os.path.join(os.path.expanduser("~"), "Desktop")
PDF_DIR       = os.path.join(DESKTOP, "PR_原始文件")

_VERSION_KEYWORDS = ['版本', 'Version', 'version']
_AUTHOR_KEYWORDS  = ['作者', '编写', '拟制', '修订人', '编制', 'Author', 'author']
_COVER_PAGE_CHAR_THRESHOLD = 800

# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────
def _parse_table_data(table_data):
    if not table_data or len(table_data) < 2:
        return "", ""
    header = [str(c or "").strip() for c in table_data[0]]
    version_col = author_col = -1
    for j, h in enumerate(header):
        if version_col < 0 and any(k in h for k in _VERSION_KEYWORDS):
            version_col = j
        if author_col < 0 and any(k in h for k in _AUTHOR_KEYWORDS):
            author_col = j
    last_data = None
    for row in reversed(table_data[1:]):
        cells = [str(c or "").strip() for c in row]
        if any(cells):
            last_data = cells
            break
    if last_data is None:
        return "", ""
    version = last_data[version_col] if 0 <= version_col < len(last_data) else ""
    author  = last_data[author_col]  if 0 <= author_col  < len(last_data) else ""
    return author, version

# ──────────────────────────────────────────────
# 步骤 1：pdfplumber 文字层
# ──────────────────────────────────────────────
def test_pdfplumber(pdf_path):
    print("\n[pdfplumber]", os.path.basename(pdf_path))
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        print(f"  页数: {len(pdf.pages)}")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            chars = re.sub(r'\s', '', text)
            print(f"  Page {i+1}: {len(chars)} chars, 含'修订记录'={('修订记录' in text)}")
            if '修订记录' in text:
                print(f"  --- 修订记录页文本 (前500字) ---\n{text[:500]}\n  ---")
                tables = page.extract_tables() or []
                print(f"  提取到 {len(tables)} 个表格")
                for ti, tbl in enumerate(tables):
                    print(f"  表格[{ti}] 行数={len(tbl)}, 内容={tbl}")
                    a, v = _parse_table_data(tbl)
                    print(f"  → 解析结果: author={repr(a)}, version={repr(v)}")

# ──────────────────────────────────────────────
# 步骤 2：pymupdf 渲染 + pytesseract OCR
# ──────────────────────────────────────────────
def test_ocr(pdf_path):
    print("\n[OCR - pymupdf+pytesseract]", os.path.basename(pdf_path))
    import fitz
    import pytesseract
    from PIL import Image

    doc = fitz.open(pdf_path)
    print(f"  页数: {len(doc)}")
    mat = fitz.Matrix(2, 2)

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')
        chars = re.sub(r'\s', '', text)
        print(f"  Page {i+1}: {len(chars)} chars, 含'修订记录'={('修订记录' in text)}")

        if '修订记录' in text:
            print(f"  --- OCR 原始文本 (前800字) ---\n{text[:800]}\n  ---")

            lines = [ln for ln in text.splitlines() if ln.strip()]
            header_idx = next(
                (idx for idx, ln in enumerate(lines)
                 if any(k in ln for k in _VERSION_KEYWORDS + _AUTHOR_KEYWORDS)),
                -1
            )
            print(f"  表头行索引: {header_idx}")
            if header_idx >= 0:
                print(f"  表头行内容: {repr(lines[header_idx])}")
                header    = re.split(r'\s{2,}|\t', lines[header_idx].strip())
                print(f"  表头分列: {header}")
                data_lines = [ln for ln in lines[header_idx + 1:] if ln.strip()]
                print(f"  数据行数: {len(data_lines)}")
                if data_lines:
                    print(f"  最后数据行: {repr(data_lines[-1])}")
                    last_parts = re.split(r'\s{2,}|\t', data_lines[-1].strip())
                    print(f"  最后行分列: {last_parts}")

                    version_col = author_col = -1
                    for j, h in enumerate(header):
                        if version_col < 0 and any(k in h for k in _VERSION_KEYWORDS):
                            version_col = j
                        if author_col < 0 and any(k in h for k in _AUTHOR_KEYWORDS):
                            author_col = j
                    print(f"  version_col={version_col}, author_col={author_col}")
                    version = last_parts[version_col] if 0 <= version_col < len(last_parts) else ""
                    author  = last_parts[author_col]  if 0 <= author_col  < len(last_parts) else ""
                    print(f"  → 解析结果: author={repr(author)}, version={repr(version)}")

# ──────────────────────────────────────────────
# 主测试
# ──────────────────────────────────────────────
# 优先测试已知失败的文件，再抽样测其他 PDF
TARGET_FILES = [
    "KEY-POTA-v1.5.4.pdf",
    "年会服务规范-v3.0.13.pdf",
]

# 从目录中取前3个 PDF 追加作为参照
all_pdfs = [f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")]
for f in all_pdfs[:3]:
    if f not in TARGET_FILES:
        TARGET_FILES.append(f)

print("=" * 60)
print("PDF 修订记录提取单元测试（修复后）")
print("=" * 60)

for fname in TARGET_FILES:
    path = os.path.join(PDF_DIR, fname)
    if not os.path.exists(path):
        print(f"\n⚠️ 文件不存在: {fname}")
        continue
    print(f"\n{'='*60}\n文件: {fname}  ({os.path.getsize(path)//1024} KB)")
    try:
        test_pdfplumber(path)
    except Exception as e:
        print(f"  pdfplumber 异常: {e}")
    # OCR 仅在 tesseract 可用时测试
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        try:
            test_ocr(path)
        except Exception as e:
            print(f"  OCR 异常: {e}")
    except Exception:
        print("  [OCR 跳过：tesseract 未安装]")

print("\n" + "=" * 60)
print("测试结束")
