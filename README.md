# 文件批处理脚本说明

本目录为围绕 **DMS / 知识库导出数据** 的批处理工具：从制表符文本或 Excel 读取链接与加密标题，下载文件、生成文件名列表、从文档「修订记录」中提取作者与版本，以及在 Excel 中按规则高亮重复或匹配的条目。

---

## 环境要求

- Python 3.10+（脚本中使用了 `tuple[str, ...]` 等类型注解）
- 按需安装依赖（未提供统一 `requirements.txt` 时，可按下方「脚本索引」中的库自行安装）

常用库：

- `requests`、`cryptography`：下载与 AES-GCM 标题解密
- `pandas`、`openpyxl`：Excel 读写
- `python-docx`、`pdfplumber`、`pymupdf`（`fitz`）、`pytesseract`、`Pillow`：文档解析与扫描版 PDF 的 OCR（用于专家流水线脚本）
- `pytest`：API 测试（`test_API.py`）

**OCR**：使用 `pytesseract` 时需本机安装 [Tesseract](https://github.com/tesseract-ocr/tesseract)，并配置中文语言包（脚本中多为 `chi_sim+eng`）。

### 标题解密（环境变量）

解密标题的脚本**不再**在代码中存放任何密钥。运行前在终端或系统中设置环境变量，值为 **Base64 编码的 16/24/32 字节 AES 密钥**（勿写入仓库或 README）。

| 环境变量 | 对应脚本 |
|----------|----------|
| `SSC_TITLE_ENCRYPT_AES_KEY_B64` | `patch_download.py`；`patch_title_expert_three.py`（SSC 段） |
| `BITMAIN_TITLE_ENCRYPT_AES_KEY_B64` | `patch_title_expert.py`、`patch_title_BIT_expert.py`；`patch_title_expert_three.py`（BITMAIN 段） |
| `SOPHGO_TITLE_ENCRYPT_AES_KEY_B64` | `patch_title_expert_three.py`（SOPHGO 段） |
| `PR_TITLE_ENCRYPT_AES_KEY_B64` | `patch_title_PR_expert.py` |

PowerShell 示例（仅说明用法，请在本机会话中设置，不要把真实值写进脚本文件）：

```powershell
$env:SSC_TITLE_ENCRYPT_AES_KEY_B64 = "<运行时填入>"
python patch_download.py
```

---

## 脚本索引

### 下载类

| 脚本 | 作用简述 | 主要输入（默认在「桌面」） | 主要输出 |
|------|----------|---------------------------|----------|
| `patch_download.py` | 读 DMS 导出的 `titles.txt`（`id\turl\ttitle_hex`），解密标题后按标题命名下载，保留原始格式（PDF/DOC/DOCX 等） | `titles.txt` | `SSC-COP_原始格式文件\` |
| `patch_download_doc.py` | 从 Excel 读 `url` 列批量下载；无合法后缀时强制 `.docx` | `SSC-COP文件备份.xlsx` | `SSC-COP_DOC文件\` |
| `patch_download_备份.py` | Excel + `url` 列，按 Content-Type / 魔数判断扩展名（与主下载逻辑类似的备份版） | `SSC-COP文件备份.xlsx` | `SSC-COP_原始格式文件\` |

### 标题解密与文件名列表（不下载文件）

| 脚本 | 作用简述 | 主要输入 | 主要输出 |
|------|----------|----------|----------|
| `patch_title_expert.py` | 读 Bitmain 侧导出文本，支持 `id\turl\ttitle_hex` 或 `id\ttitle_hex`，解密后生成文件名列表 | `bitmain-COP.txt` | `bitmain-COP_文件名列表.txt` |
| `patch_title_expert_three.py` | 分别处理 SSC / BITMAIN / SOPHGO 三个 `知识集合` 文本（两列：id + hex），合并写入一个分段列表 | `ssc-知识集合.txt`、`bitmain-知识集合.txt`、`sophgo-知识集合.txt` | `三方文件名列表.txt` |

### 下载 + 修订记录解析 + Excel 报表

| 脚本 | 作用简述 | 主要输入 | 主要输出 |
|------|----------|----------|----------|
| `patch_title_BIT_expert.py` | 读 `title_BIT.txt`（前 4 行为表头，数据行为 `id\turl\ttitle_hex\towner`），**下载**到本地后从 docx/pdf 的「修订记录」表提取作者与当前版本；PDF 在文字层失败时用 OCR 兜底 | `title_BIT.txt` | `BIT_原始文件\`、`title_BIT_output.xlsx` |
| `patch_title_PR_expert.py` | 逻辑与 BIT 版类似，但针对 PR：**不重新下载**，在 `PR_原始文件\` 中按解密标题匹配已下载文件，再提取作者/版本并写 Excel | `title_PR.txt` | `title_PR_output.xlsx`（依赖已有 `PR_原始文件\`） |

两脚本均使用 **AES-GCM**：密文为十六进制字符串，结构为 12 字节 IV + 密文与认证标签。输出 Excel 列一般为：**标题**、**负责人**、**作者（当前版本）**。保存前请关闭已打开的同名 xlsx，否则会 `PermissionError`。

### Excel 高亮

| 脚本 | 作用简述 |
|------|----------|
| `patch_title_same_highlight.py` | 对固定目录下两个工作簿：用 **第 2 个 sheet（COP）** 第 A 列文件名（去末尾版本号）作为集合，在 **第 1 个 sheet** 中整行黄色高亮匹配行。默认处理 `桌面\三主体知识集合-COP\` 下指定 xlsx。 |
| `highlight_cop_filenames.py` | 根据 **TXT 中的文件名列表**，在工作簿所有 sheet 的单元格中：完全匹配标黄，仅版本号不同标绿。支持命令行参数 `--txt`、`--xlsx`、`--out`（有桌面默认路径）。 |

### 测试与调试

| 脚本 | 作用简述 |
|------|----------|
| `test_ocr.py` | 针对 `桌面\PR_原始文件\` 中指定/抽样 PDF，打印 pdfplumber 与 pymupdf+Tesseract 的修订记录提取过程（调试 OCR 与表格解析）。 |
| `test_API.py` | `pytest` 用例：对搜索/文档/RAG/聊天等 HTTP API 做结构与分页断言。接口地址、超时与可选认证等见该文件顶部的环境变量说明。 |

---

## 数据格式备忘

- **`titles.txt`（SSC 下载）**：前几行为列头或空行，有效行为三列制表符分隔：`id`、`url`、`title_hex`（`url` 以 `http` 开头）。
- **`title_BIT.txt` / `title_PR.txt`**：`utf-8-sig`，跳过前 4 行表头后，每行至少四列：`id`、`url`、`title_hex`、`owner`。
- **加密标题**：Hex 编码的 AES-GCM 包（IV + ciphertext+tag），解密后为 UTF-8 明文；脚本中会再做路径段、冒号前缀、常见扩展名等规范化以得到安全文件名。

---

## 运行示例

在项目目录下（请按本机 Python 命令调整）：

```bash
python patch_download.py
python patch_title_expert.py
python highlight_cop_filenames.py --txt "C:\Users\你的用户名\Desktop\bitmain-COP_文件名列表.txt" --xlsx "...\input.xlsx" --out "...\output.xlsx"
pytest test_API.py -v
```

---

## 文件一览

| 文件名 | 备注 |
|--------|------|
| `patch_download.py` | SSC：`titles.txt` → 原始格式批量下载 |
| `patch_download_doc.py` | Excel URL → DOC 目录 |
| `patch_download_备份.py` | Excel URL 下载（备份实现） |
| `patch_title_expert.py` | Bitmain 文件名列表 |
| `patch_title_expert_three.py` | 三方知识集合 → 合并文件名列表 |
| `patch_title_BIT_expert.py` | BIT：下载 + 修订记录提取 + xlsx |
| `patch_title_PR_expert.py` | PR：本地文件 + 修订记录提取 + xlsx |
| `patch_title_same_highlight.py` | 双 sheet COP 重复高亮 |
| `highlight_cop_filenames.py` | TXT 与 Excel 文件名交叉高亮 |
| `test_ocr.py` | PDF 提取调试 |
| `test_API.py` | 后端 API 契约测试 |
