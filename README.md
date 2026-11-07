# ChsToChtDocumentConverter — High-Fidelity Simplified-to-Traditional Chinese Document Converter

> 🌐 **Language / 语言**: [**English (Default)**](./README.md) | [**简体中文**](./README.zh-CN.md)  


---

## 1. Purpose

**ChsToChtDocumentConverter** is an offline desktop application for high-fidelity conversion and interactive proofreading of documents from Simplified Chinese to Traditional Chinese (Taiwan Standard / Traditional Orthography). Its core purposes are:
- **In-Place Layout-Preserving Conversion**: Batch-convert office documents, fixed-layout PDFs, and scanned images from Simplified Chinese to Traditional Chinese while **faithfully preserving the original typography, font styles, table borders, SmartArt diagrams, chart labels, and image layouts**.
- **Context-Aware Disambiguation & Interactive Proofreading**: Automatically identify one-to-many Simplified-to-Traditional ambiguous characters (such as `干 → 乾/幹`, `面 → 面/麵`, `脏 → 臟/髒`, `发 → 發/髮`, `系 → 係/繫`) and list **"Suspect Words" (待检字词)** in document reading order inside a dual-pane interactive spell-checker, allowing one-click candidate selection or custom edits that synchronize instantaneously to both the preview and final exported files.

---

## 2. Tech Stack

- **Core Language & Runtime**: Python 3.13
- **Desktop GUI Framework**: PySide6 (Qt 6) + PySide6-Fluent-Widgets (`qfluentwidgets`, Fluent Design component library)
- **Simplified-to-Traditional Conversion & Disambiguation Engine**:
  - `opencc-python-reimplemented` (OpenCC Forward Maximum Matching `mmseg` segmentation and conversion)
  - Custom merged high-coverage dictionary (`my_s2tw.json`, `my_STPhrases.txt`, `my_STCharacters.txt`) built by combining MediaWiki's `ZhConversion.php` (`ZH_TO_TW` / `ZH_TO_HANT`) with OpenCC's `STPhrases.txt`, `TWPhrases.txt`, and `STCharacters.txt`
- **Document & Layout Processing Engine**:
  - `python-docx` + `lxml` + Python standard library `zipfile` / `xml.etree.ElementTree` (in-place XML node parsing and rewriting for OpenXML and ODF containers)
  - `PyMuPDF` (`fitz`, span-level transparent redaction, baseline-accurate CJK font embedding, and high-DPI page rasterization)
  - `pywin32` (`win32com.client`, local Microsoft Word COM automation) + LibreOffice Headless CLI pipeline
- **Image Processing & OCR Engine**:
  - `Pillow` (`PIL`, adaptive scaling, grid-line filtering, median background/foreground color sampling, character bounding-box erasure, and centered glyph rendering)
  - Windows native `Windows.Media.Ocr` (WinRT `OcrEngine`) + `pytesseract` (Tesseract OCR) dual-channel text recognition and bounding-box extraction
- **Build, Packaging & Testing Tools**:
  - `Nuitka` (compiles Python source and runtime dependencies into a standalone Windows executable `app_ui.exe`)
  - `pytest` + `pytest-mock` (automated unit and integration test suite)

---

## 3. Supported Runtime Environments

- **Operating System**: Windows 10 / Windows 11 (64-bit, x64)
- **Built-in OS Dependencies**:
  - Windows 10/11 Simplified Chinese language pack and OCR capability (used by `Windows.Media.Ocr`, pre-installed on standard Chinese Windows systems)
  - Windows built-in CJK fonts (`msyh.ttc` Microsoft YaHei, `simhei.ttf` SimHei, `simsun.ttc` SimSun)
- **Optional External Software (On-Demand)**:
  - Either **Microsoft Word** or **LibreOffice** installed locally is required when converting legacy binary `.doc` files or generating on-demand true-layout page renders for `.docx` / `.odt` documents.

---

## 4. Supported Document Formats

- **Modern Word Processing Documents**: `.docx` (Word OpenXML), `.odt` (OpenDocument Text)
- **Legacy Word Processing Documents**: `.doc` (Word 97–2003 binary format)
- **Fixed-Layout & Scanned PDFs**: `.pdf` (native vector text PDFs, mixed text/image PDFs, and image-only scanned PDFs)
- **Images & Scanned Documents**: `.jpg` / `.jpeg`, `.png`, `.bmp`, `.tiff` / `.tif`

---

## 5. Current Features

1. **Merged Dictionary & Phrase-Boundary Disambiguation**:
   - Uses Forward Maximum Matching (`mmseg`) to lock multi-character phrase boundaries first, flagging only uncovered ambiguous single characters or ambiguous phrases as "Suspect Words" sorted strictly in natural reading order.
2. **Interactive Red-Box Highlighting & Proofreading**:
   - Dual-pane preview layout with document preview on the left and Suspect Words list on the right;
   - Hovering over any suspect entry on the right immediately draws a **red bounding box** around the target word in the left pane (across paragraphs, table cells, cards, or rendered pages) and scrolls it into view;
   - Clicking a suspect entry opens a high-contrast correction dialog for one-click candidate switching or custom text replacement, updating both the preview and export buffer in real time.
3. **Dual-Mode Preview Architecture (Instant Rich-Text + On-Demand True-Layout Page Render)**:
   - **📝 Instant Rich-Text Mode (Default)**: In-memory structured HTML rendering (`< 50ms` response) displaying body paragraphs, bordered tables, `[TextBox]` cards, `[SmartArt]` node chains, and `[Chart]` data preview tables;
   - **🖼️ On-Demand True-Layout Page Render Mode**: One-click high-DPI page rasterization preserving 100% of original page geometry, supporting 50%–200% zoom, character-level red-box overlays, and a solid-background loading modal when rendering exceeds 500ms.
4. **Full-Node High-Fidelity In-Place Replacement**:
   - Covers body text, tables, headers/footers, text boxes, SmartArt diagrams, and chart titles/axis labels;
   - Preserves vector borders, colored backgrounds, and embedded images in PDFs;
   - Image/JPEG pipeline features adaptive resolution scaling, smart table grid-line removal, second-pass horizontal strip stitching for unaligned/indented leading characters, wrapped tail-line phrase merging, `TextAngle` center-unrotation coordinate recovery, and automatic background/foreground color sampling (unchanged characters keep 100% of their original crisp pixels).
5. **Concurrent Batch Queue & Export Management**:
   - Supports batch file import, multi-threaded asynchronous conversion, individual file export, or one-click "Save All".

---

## 6. Unsupported Features

1. **File Drag-and-Drop into UI**:
   - Dragging files from Windows Explorer directly into the task queue is currently unsupported; please use the **"Add Files" (添加文件)** button at the top of the window.
2. **Password-Protected or DRM-Encrypted Documents**:
   - Encrypted documents must be unlocked prior to import.
3. **Outlined Vector Path Text in PDFs**:
   - Text converted into pure geometric curves/outlines without font encoding or bitmap layers cannot be extracted or replaced.
4. **Embedded Third-Party Binary OLE Objects**:
   - Legacy Equation 3.0 binary streams, embedded Visio binary objects, and rasterized WordArt bitmaps.
5. **Deep-Learning Background Inpainting & Handwritten OCR**:
   - Image character erasure currently uses median perimeter solid-color fills rather than neural texture inpainting over complex photographic gradients, and cursive handwriting OCR is not supported.
6. **Automatic Horizontal-to-Classical Vertical Right-to-Left Reflow**:
   - Original reading direction and layout flow are preserved without automatic reflow into classical vertical typography.

---

## 7. Known Issues

1. **Drag-and-Drop File Import into UI Does Not Work (Known Issue)**:
   - **Symptom**: Dragging files from Windows Explorer into the main queue table does not enqueue files (due to Windows UIPI/privilege isolation and Qt6/Fluent scroll-viewport event interception).
   - **Status**: Drag-and-drop prompt text has been removed from the UI and documented as a known issue. Please use the **"添加文件" (Add Files)** button at the top-left of the interface.
2. **`.doc` Conversion & True-Layout Office Rendering Fallback Without Word / LibreOffice**:
   - On machines without Microsoft Word or LibreOffice installed, binary `.doc` files cannot be converted, and `.docx` / `.odt` true-layout page rendering falls back to structured page-card rendering.
3. **OCR Accuracy on Ultra-Low-Resolution or Decorative Fonts**:
   - Very small glyphs (`< 14px`), heavy JPEG compression artifacts, or stylized decorative fonts may occasionally cause near-shape character misrecognition in Windows native OCR.

---

## 8. How to Build Locally

### 8.1 Prerequisites & Virtual Environment
1. Install **Python 3.13** (64-bit) and create a virtual environment:
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   ```
2. Install required dependencies:
   ```powershell
   pip install PySide6 PySide6-Fluent-Widgets opencc-python-reimplemented python-docx lxml pymupdf Pillow pytesseract pywin32 nuitka pytest pytest-mock
   ```

### 8.2 Build the Custom Merged Dictionary (First-Time Setup or Dictionary Update)
```powershell
.venv\Scripts\python build_custom_dict.py
```
> This generates `my_s2tw.json`, `my_STPhrases.txt`, and `my_STCharacters.txt` in the project root.

### 8.3 Run from Source & Execute Tests
```powershell
# Launch the desktop GUI
.venv\Scripts\python app_ui.py

# Run the automated test suite
.venv\Scripts\pytest -v
```

### 8.4 Compile Standalone Executable with Nuitka
Run the packaging script from the project root:
```cmd
cmd /c packaging.cmd
```
- The standalone distribution is output to `dist\app_ui.dist\` with the main executable at `dist\app_ui.dist\app_ui.exe` (bundling `my_s2tw.json`, `my_STPhrases.txt`, `my_STCharacters.txt`, and all required Qt / OpenCC / PyMuPDF runtime assets).

---

<details>
<summary>🇨🇳 <b>点击展开中文说明 / Click to expand Simplified Chinese README (简体中文)</b></summary>

### 1. 用途
**ChsToChtDocumentConverter** 是一款面向桌面端的本地离线简繁文档高保真转换与交互式校对工具，支持在保留原文档排版、表格、SmartArt、图表与扫描图布局的前提下完成简转繁及「待检字词」交互式校对。

### 2. 技术栈
- **语言与 GUI**：Python 3.13 + PySide6 (Qt 6) + PySide6-Fluent-Widgets (`qfluentwidgets`)
- **转换与消歧引擎**：`opencc-python-reimplemented` + MediaWiki `ZhConversion.php` 融合词典（`my_s2tw.json`、`my_STPhrases.txt`、`my_STCharacters.txt`）
- **文档与图像引擎**：`python-docx`、`lxml`、`PyMuPDF` (`fitz`)、`pywin32`、LibreOffice Headless、`Pillow`、`Windows.Media.Ocr`、`pytesseract`
- **打包与测试**：`Nuitka`、`pytest`、`pytest-mock`

### 3. 支持的运行环境
- Windows 10 / Windows 11 (x64)，内置简体中文 OCR 组件与系统字体；可选安装 Microsoft Word 或 LibreOffice。

### 4. 支持文档类型
- `.docx`、`.odt`、`.doc`、`.pdf`、`.jpg` / `.jpeg`、`.png`、`.bmp`、`.tiff` / `.tif`

### 5. 目前的特性
- 融合词库与 `mmseg` 词组边界保护消歧；
- 双栏交互式「待检字词」悬停红框定位与校对弹窗；
- 默认极速富文本预览（`< 50ms`）+ 按需原貌页高清渲染双模式；
- 全节点高保真原位替换（含 JPEG 非对齐缩进首字二次补扫与 `TextAngle` 倾斜角中心逆旋转坐标还原）；
- 多任务并发队列与一键保存。

### 6. 暂不支持的功能
- 文件拖拽（Drag & Drop）入 UI、加密文档、PDF 转曲路径艺术字、二进制 OLE 对象、复杂自然背景深度学习修补、古籍竖排自动重排。

### 7. 目前问题
- **文件无法通过鼠标拖拽进入 UI 界面**：已移除界面拖拽提示文案并标记为已知问题，请使用左上角「添加文件」按钮；
- 无 Word/LibreOffice 环境下的 `.doc` 与原貌页回退限制；
- 极端低分辨率或艺术字体的图片 OCR 误差。

### 8. 如何本地构建项目
完整中文说明请参见 [README.zh-CN.md](./README.zh-CN.md)。

</details>
