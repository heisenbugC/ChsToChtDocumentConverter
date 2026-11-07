# 简繁易 - 简繁文档专业转换工具

> 🌐 **Language / 语言**: [**English (Default)**](./README.md) | [**简体中文**](./README.zh-CN.md)

---

## 1. 用途

**简繁易** 是一款面向桌面端的本地离线简繁文档高保真转换与交互式校对工具。主要用途包括：
- 将各类办公文档、版式文件与扫描图片从简体中文批量转换为繁体中文（台湾正体 / 传统繁体标准），在转换文字的同时**原位保留源文档的排版、字体样式、表格边框、SmartArt 流程图、图表数据及图像布局**；
- 自动识别简体一简对多繁的歧义字（如“干/乾/幹”、“面/麵”、“脏/臟/髒”、“发/發/髮”、“系/係/繫”等），并在内置的双栏交互式校对器中按文档出现顺序高亮列出**「待检字词」**，支持人工一键切换候选字或自定义修改，实时同步至预览与最终导出文件。

---

## 2. 技术栈

- **核心语言与运行时**：Python 3.13
- **图形界面框架**：PySide6 (Qt 6) + PySide6-Fluent-Widgets (`qfluentwidgets`，Fluent Design 风格桌面组件库)
- **简繁转换与词库消歧引擎**：
  - `opencc-python-reimplemented`（OpenCC 正向最大匹配 `mmseg` 分词转换）
  - 融合 MediaWiki `ZhConversion.php`（`ZH_TO_TW` / `ZH_TO_HANT`）与 OpenCC `STPhrases.txt` / `TWPhrases.txt` / `STCharacters.txt` 构建的定制高覆盖率词典（`my_s2tw.json`、`my_STPhrases.txt`、`my_STCharacters.txt`）
- **文档与版式处理引擎**：
  - `python-docx` + `lxml` + 标准库 `zipfile` / `xml.etree.ElementTree`（OpenXML / ODF 底层 XML 节点解析与原位回写）
  - `PyMuPDF` (`fitz`，PDF Span 级透明擦除与原基线字体嵌入重绘、高清页面光栅化渲染)
  - `pywin32` (`win32com.client`，本机 Microsoft Word COM 自动化转换通道)与 LibreOffice Headless 命令行通道
- **图像处理与 OCR 识别引擎**：
  - `Pillow` (`PIL`，图像预处理、色彩中位数采样、字框背景抹除与居中繁体字绘制)
  - Windows 原生 `Windows.Media.Ocr` (WinRT `OcrEngine`) + `pytesseract` (Tesseract OCR) 双通道文字检测与坐标提取
- **编译打包与测试工具**：
  - `Nuitka`（将 Python 代码与依赖编译为独立免安装 Windows 可执行程序 `app_ui.exe`）
  - `pytest` + `pytest-mock`（自动化单元测试与集成测试套件）

---

## 3. 支持的运行环境

- **操作系统**：Windows 10 / Windows 11（64 位，x64）
- **系统内置组件依赖**：
  - Windows 10/11 中文（简体）语言包及 OCR 功能组件（供 `Windows.Media.Ocr` 调用，系统默认内置）
  - Windows 系统内置中文字体（微软雅黑 `msyh.ttc`、黑体 `simhei.ttf`、宋体 `simsun.ttc`）
- **可选外部软件（按需调用）**：
  - 处理 `.doc` 旧版二进制文档或为 `.docx`/`.odt` 生成「按需原貌页渲染」时，需本机已安装 **Microsoft Word** 或 **LibreOffice**（任选其一即可）。

---

## 4. 支持文档类型

- **现代文字处理文档**：`.docx`（Word OpenXML 文档）、`.odt`（OpenDocument 文本）
- **旧版文字处理文档**：`.doc`（Word 97-2003 二进制文档）
- **固定版式与扫描版 PDF**：`.pdf`（包含原生矢量文本 PDF、图文混排 PDF 及纯扫描图像型 PDF）
- **图像与扫描件格式**：`.jpg` / `.jpeg`、`.png`、`.bmp`、`.tiff` / `.tif`

---

## 5. 目前的特性

1. **融合词库与词组边界保护消歧**：
   - 采用正向最大匹配（`mmseg`）优先锁定多字词组边界，仅对未被多字词组覆盖的单字或含一简多繁歧义字的词组生成「待检字词」，按文档自然阅读顺序排列。
2. **交互式「待检字词」红框定位与校对**：
   - 预览界面采用左（文档内容预览）、右（待检字词列表）双栏布局；
   - 鼠标悬停右侧任一待检字词时，左侧预览区（含富文本段落、表格单元格、功能卡片或原貌渲染页）会实时绘制**红色高亮定位框**并自动滚动至对应位置；
   - 点击待检字词可弹出高对比度校对弹窗，支持一键选择推荐候选词或手动输入修正词，修改结果立即同步至左侧预览与导出缓存。
3. **双模式预览架构（默认极速富文本 + 按需原貌页渲染）**：
   - **📝 极速富文本模式（默认）**：纯内存解析生成结构化 HTML，将正文段落、网格表格、`【文字方块 / TextBox】`、`【SmartArt 流程图 / 视觉节点】`、`【图表 / Chart 数据预览】` 瞬时渲染展现（响应时间 `< 50ms`）；
   - **🖼️ 按需原貌页渲染模式**：支持一键将文档按真实页面尺寸光栅化为高清页面画布，支持 50%–200% 无级缩放及图片字符框原位红框高亮定位；耗时超过 500ms 时提供实色背景加载提示弹窗。
4. **全节点高保真原位替换**：
   - 完整覆盖正文、表格、页眉页脚、文字方块、SmartArt 节点、图表标题与数据轴标签；
   - PDF 转换保留底层矢量线条、彩色背景框与内嵌图像；
   - 扫描图片转换支持自适应宽高缩放、表格细线智能滤除、非对齐缩进首字二次补扫、折行短尾词上下文合并、`TextAngle` 倾斜角中心逆旋转坐标还原，以及原字背景色/前景色自动采样重绘（未发生简繁变化的字符 100% 保留原图清晰字迹）。
5. **多任务并发队列与导出管理**：
   - 支持批量添加文件、并发线程池异步转换、单文件独立保存或一键「全部保存」。

---

## 6. 暂不支持的功能

1. **文件拖拽（Drag & Drop）导入 UI**：
   - 目前暂不支持从 Windows 资源管理器直接拖拽文件进入任务列表，请使用界面顶部的**「添加文件」**按钮导入文档。
2. **受密码保护或权限锁定的加密文档**：
   - 带有打开密码或 DRM 权限加密的文档需先解密后再导入转换。
3. **PDF 纯矢量转曲路径（Outlines）艺术字**：
   - 已转曲为纯几何路径且既无字体编码层、亦非内嵌位图的文字无法提取与替换。
4. **内嵌第三方二进制 OLE 对象**：
   - 如旧版 Equation 3.0 二进制公式流、内嵌 Visio 二进制对象及纯位图贴图艺术字。
5. **复杂自然场景纹理背景修补（Inpainting）与手写体识别**：
   - 图片文字擦除目前采用周边像素中位数纯色矩形填充，暂不支持跨渐变/复杂自然照片背景的深度学习纹理修补及潦草手写体识别。
6. **古籍竖排右起版式自动重排**：
   - 转换过程保持源文档原有横竖排版流向不变，暂不支持横排自动重构为古籍竖排右起版式。

---

## 7. 目前问题

1. **文件无法通过鼠标拖拽进入 UI 界面（Known Issue）**：
   - **现象**：从 Windows 资源管理器拖拽文档至主界面任务列表区域时，无法触发文件入队（在 Windows UAC 权限隔离、进程完整性级别差异或 Qt6/Fluent 组件视口事件拦截下拖拽失效）。
   - **当前状态**：已移除界面上的拖拽提示文案，统一标记为已知问题；请通过点击界面左上角的**「添加文件」**按钮选择文档。
2. **无 Word / LibreOffice 环境下的 `.doc` 与 Office 原貌页渲染回退限制**：
   - 若运行机器既未安装 Microsoft Word 也未安装 LibreOffice，`.doc` 二进制格式无法转换，且 `.docx` / `.odt` 在切换至「🖼️ 按需原貌页渲染」时会回退为结构化页面卡片渲染。
3. **极端低分辨率或艺术字体的图片 OCR 误差**：
   - 对于字高过小（`< 14px`）、重度 JPEG 压缩伪影或特殊装饰性字体的图片，Windows 原生 OCR 仍可能出现个别近形字误识（可通过右侧校对列表或导出后复核）。

---

## 8. 如何本地构建项目

### 8.1 环境准备
1. 安装 **Python 3.13**（64 位）并创建虚拟环境：
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   ```
2. 安装项目依赖：
   ```powershell
   pip install PySide6 PySide6-Fluent-Widgets opencc-python-reimplemented python-docx lxml pymupdf Pillow pytesseract pywin32 nuitka pytest pytest-mock
   ```

### 8.2 生成定制融合词典（首次构建或更新词库时执行）
```powershell
.venv\Scripts\python build_custom_dict.py
```
> 执行后会在项目根目录生成 `my_s2tw.json`、`my_STPhrases.txt` 与 `my_STCharacters.txt`。

### 8.3 本地源码运行与测试
```powershell
# 启动桌面图形界面
.venv\Scripts\python app_ui.py

# 运行完整自动化测试套件
.venv\Scripts\pytest -v
```

### 8.4 使用 Nuitka 编译独立可执行文件
在项目根目录下执行打包脚本：
```cmd
cmd /c packaging.cmd
```
- 编译完成后，独立发行目录位于 `dist\app_ui.dist\`，主程序入口为 `dist\app_ui.dist\app_ui.exe`（已内嵌 `my_s2tw.json`、`my_STPhrases.txt`、`my_STCharacters.txt` 及所需的 Qt / OpenCC / PyMuPDF 运行时资源）。
