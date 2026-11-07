import os
import pytest
from unittest.mock import patch, MagicMock
from pipeline import DisambiguationEngine, DocumentProcessor
import docx
import zipfile
import io

def test_engine_disambiguation():
    engine = DisambiguationEngine()
    assert engine.convert_text("干燥") == "乾燥"
    assert engine.convert_text("干什么") == "幹什麼"

def test_process_docx(tmp_path):
    test_docx_path = tmp_path / "test.docx"
    doc = docx.Document()
    doc.add_paragraph("干燥的饼干")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0,0).paragraphs[0].add_run("发现干活")
    doc.save(str(test_docx_path))
    
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    output_path = tmp_path / "test_繁体.docx"
    
    processor._process_docx(str(test_docx_path), str(output_path), None)
    
    assert os.path.exists(output_path)
    res_doc = docx.Document(str(output_path))
    assert "乾" in res_doc.paragraphs[0].text
    assert "發" in res_doc.tables[0].cell(0,0).paragraphs[0].text
    assert "幹" in res_doc.tables[0].cell(0,0).paragraphs[0].text

@patch('pymupdf.open')
def test_process_pdf_mocked(mock_fitz_open, tmp_path):
    mock_doc = MagicMock()
    mock_page = MagicMock()
    mock_doc.__iter__.return_value = [mock_page]
    mock_doc.__len__.return_value = 1
    
    mock_page.get_text.return_value = [
        (0, 0, 10, 10, "头发", 0, 0)
    ]
    mock_fitz_open.return_value = mock_doc
    
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    
    test_pdf_path = str(tmp_path / "test.pdf")
    output_path = str(tmp_path / "test_繁体.pdf")
    
    processor._process_pdf(test_pdf_path, output_path, None)
    
    mock_page.add_redact_annot.assert_called()
    mock_page.apply_redactions.assert_called()
    mock_page.insert_textbox.assert_called()
    assert "頭髮" in mock_page.insert_textbox.call_args[0][1]

@patch('subprocess.run')
def test_process_via_libreoffice(mock_run, tmp_path):
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    
    test_doc_path = str(tmp_path / "test.doc")
    with open(test_doc_path, "w") as f:
        f.write("dummy")
        
    output_path = str(tmp_path / "test_繁体.doc")
    
    def side_effect(*args, **kwargs):
        # Determine if it's doc to docx or docx to doc
        outdir_idx = args[0].index("--outdir")
        outdir = args[0][outdir_idx+1]
        input_file = args[0][4]
        
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        if "doc" in args[0][3] and args[0][3] == "docx":
            # output is _converted.docx
            pass
        elif args[0][3] == "doc":
            # Mock the creation of the final .doc file
            final_tmp = os.path.join(outdir, f"{base_name.replace('_converted', '')}_converted.doc")
            with open(final_tmp, "w") as f:
                f.write("mocked doc")

    mock_run.side_effect = side_effect

    with patch.object(processor, '_process_docx') as mock_docx:
        processor._process_via_libreoffice(test_doc_path, output_path, None)
        
    assert mock_run.call_count == 2
    mock_docx.assert_called_once()
    assert os.path.exists(output_path)

@patch('pytesseract.image_to_data')
@patch('PIL.Image.open')
def test_process_jpg(mock_img_open, mock_tesseract, tmp_path):
    mock_img = MagicMock()
    mock_img.size = (100, 100)
    mock_img_open.return_value = mock_img.convert.return_value = mock_img
    
    mock_tesseract.return_value = {
        'level': [1],
        'text': ['干燥'],
        'left': [0],
        'top': [0],
        'width': [10],
        'height': [10]
    }
    
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    
    test_jpg_path = str(tmp_path / "test.jpg")
    output_path = str(tmp_path / "test_繁体.jpg")
    
    with patch('PIL.ImageDraw.Draw'):
        with patch('PIL.ImageFont.truetype'):
            processor._process_jpg(test_jpg_path, output_path, None)
            
    mock_img.save.assert_called_once_with(output_path)

def test_find_suspects_order():
    """测试 DisambiguationEngine.find_suspects 返回的待检字词列表是否严格按照字词在原文 text 中首次出现的先后顺序排列"""
    engine = DisambiguationEngine()
    
    # 提供一个包含多个可疑词汇的示例文本
    text = "头发很干净，干什么呢"
    
    # 获取待检字词列表
    suspects = engine.find_suspects(text)
    
    # 提取所有简化字（忽略其他字段）
    simplified_words = [s["simplified"] for s in suspects]
    
    # 验证顺序：如果两个或多个可疑词汇都出现在文本中，
    # 它们应该按照在文本中首次出现的顺序排列
    # 对于这个特定文本，如果 "头发" 和 "干什么" 都是可疑词汇，
    # 那么 "头发" 应该排在 "干什么" 前面，因为 "头发" 更早出现
    
    # 获取文本中所有可疑词汇的首次出现位置
    positions = {}
    for word in simplified_words:
        position = text.find(word)
        if position != -1:
            positions[word] = position
    
    # 验证词汇按照首次出现位置排序
    if len(positions) > 1:
        sorted_by_position = sorted(positions.items(), key=lambda x: x[1])
        expected_order = [word for word, pos in sorted_by_position]
        
        assert simplified_words == expected_order, \
            f"词汇顺序不正确。预期: {expected_order}, 实际: {simplified_words}"
    
    # 对于这个示例，验证具体情况（如果 "头发" 和 "干什么" 是可疑词汇）
    # 这种情况取决于 my_STPhrases.txt 和 my_STCharacters.txt 文件内容
    # 我们只验证方法返回的顺序是否符合文本顺序的一般规则

def test_process_and_read_odt_native(tmp_path):
    """测试使用 Python 标准库 zip 模块和 XML 解析来直接读取 ODT 文件的功能"""
    import zipfile
    from xml.etree import ElementTree as ET

    # 创建一个真实的最小 ODT 文件，其中包含 content.xml
    sample_odt_path = tmp_path / "sample.odt"
    
    # ODT 是一个 zip 压缩文件，包含 content.xml 等文件
    with zipfile.ZipFile(sample_odt_path, 'w') as odt_zip:
        # 创建 content.xml 文件内容
        content_xml = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text><text:p>头发很干净，干什么呢</text:p></office:text></office:body>
</office:document-content>"""
        
        # 将 content.xml 添加到 ODT 文件中
        odt_zip.writestr('content.xml', content_xml)
    
    # 实例化引擎和处理器
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    
    # 使用侧边 run 模拟无 LibreOffice 环境
    with patch('subprocess.run', side_effect=FileNotFoundError("LibreOffice not found")):
        res = processor.process_to_temp(str(sample_odt_path))
    
    # 断言生成的 temp_output_path 存在
    assert os.path.exists(res["temp_output_path"])
    
    # 断言 "頭髮" 和 "乾淨" 出现在 res["preview_data"] 中
    # 注意：这是繁体中文的预期结果
    assert "頭髮" in res["preview_data"]
    assert "乾淨" in res["preview_data"]
    
    # 断言 len(res["suspects"]) > 0
    assert len(res["suspects"]) > 0

def test_process_pdf_to_temp_preview_and_suspects(tmp_path):
    """测试使用 pymupdf 来直接读取 PDF 文件的功能"""
    # 由于 pymupdf 可能未安装，我们需要进行适当的模拟
    try:
        import pymupdf
    except ImportError:
        # 如果 pymupdf 未安装，跳过测试
        pytest.skip("pymupdf 未安装，跳过此测试")
    
    # 创建一个包含一页中文文本的 PDF
    test_pdf_path = tmp_path / "test_chinese.pdf"
    
    # 使用 pymupdf.open 来创建一个新的 PDF 文档
    doc = pymupdf.open()
    page = doc.new_page()
    
    # 设置页面内容
    page.insert_textbox((100, 100), "干燥的头发很干净，干什么呢")
    
    # 保存 PDF
    doc.save(str(test_pdf_path))
    doc.close()
    
    # 实例化引擎和处理器
    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    
    # 调用 processor.process_to_temp
    res = processor.process_to_temp(str(test_pdf_path))
    
    # 断言生成的结果包含 temp_output_path
    assert "temp_output_path" in res
    assert os.path.exists(res["temp_output_path"])
    
    # 断言 "頭髮" 和 "乾淨" 出现在 res["preview_data"] 中
    assert "頭髮" in res["preview_data"]
    assert "乾淨" in res["preview_data"]
    assert "preview_html" in res and "頭髮" in res["preview_html"]
    
    # 断言 len(res["suspects"]) > 0
    assert len(res["suspects"]) > 0


def test_rich_preview_html_and_render_pages(tmp_path):
    """测试方案 B 极速富文本预览 HTML (含表格) 与按需页面渲染接口"""
    import docx as py_docx
    test_docx = tmp_path / "rich_test.docx"
    d = py_docx.Document()
    d.add_paragraph("测试正文段落：头发很干净")
    tbl = d.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "老北京炸酱面"
    tbl.cell(0, 1).text = "温州猪脏粉"
    tbl.cell(1, 0).text = "广味腊肠煲仔饭"
    tbl.cell(1, 1).text = "湖南腊肉炒香干"
    d.save(str(test_docx))

    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)
    res = processor.process_to_temp(str(test_docx))

    assert "preview_html" in res
    html_str = res["preview_html"]
    assert "<table" in html_str
    assert "老北京炸醬麪" in html_str
    assert "溫州豬髒粉" in html_str
    assert "湖南臘肉炒香乾" in html_str

def test_image_process_sidecar_and_rich_preview(tmp_path):
    """测试图片处理时 sidecar .ocr.txt 的生成以及富文本预览不漏字（如发、胡）"""
    from PIL import Image, ImageDraw, ImageFont
    test_img_path = str(tmp_path / "test_img.png")
    output_path = str(tmp_path / "test_img_繁体.png")

    img = Image.new("RGB", (600, 200), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("msyh.ttc", 28)
    except Exception:
        font = ImageFont.load_default()
    draw.text((30, 50), "老胡对着镜子理头发", fill=(0, 0, 0), font=font)
    img.save(test_img_path)

    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)

    # 1. process_to_temp
    res = processor.process_to_temp(test_img_path)
    assert res["status"] == "success"
    temp_out = res["temp_output_path"]
    assert os.path.exists(temp_out)
    assert os.path.exists(temp_out + ".ocr.txt")

    # 2. 验证 preview_data 与 preview_html 中关键字符 100% 保留
    preview_data = res["preview_data"]
    preview_html = res["preview_html"]
    assert "老胡" in preview_data
    assert "老胡" in preview_html
    assert "頭髮" in preview_data or "头发" in preview_data
    assert "頭髮" in preview_html or "头发" in preview_html

    # 3. 验证 generate_rich_preview_html 直接命中 sidecar
    rich_html = processor.generate_rich_preview_html(temp_out)
    assert "老胡" in rich_html

def test_scanned_pdf_ocr_and_processing(tmp_path):
    """测试纯扫描件/无矢量文本块 PDF 页面的 OCR 提取与原地繁体替换"""
    import pymupdf
    from PIL import Image, ImageDraw, ImageFont

    test_png = str(tmp_path / "scan_page.png")
    img = Image.new("RGB", (600, 200), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("msyh.ttc", 28)
    except Exception:
        font = ImageFont.load_default()
    draw.text((30, 50), "老胡在单身宿舍理头发", fill=(0, 0, 0), font=font)
    img.save(test_png)

    # 创建无矢量文字的扫描 PDF
    scanned_pdf = str(tmp_path / "scanned.pdf")
    doc = pymupdf.open()
    page = doc.new_page(width=600, height=200)
    page.insert_image(page.rect, filename=test_png)
    doc.save(scanned_pdf)
    doc.close()

    engine = DisambiguationEngine()
    processor = DocumentProcessor(engine)

    # 验证扫描 PDF 文本提取
    extracted = processor._extract_pdf_text(scanned_pdf)
    assert "老胡" in extracted

    # 验证扫描 PDF 处理
    out_pdf = str(tmp_path / "scanned_繁体.pdf")
    processor._process_pdf(scanned_pdf, out_pdf, None)
    assert os.path.exists(out_pdf)

