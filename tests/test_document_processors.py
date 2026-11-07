import pytest
import docx
from unittest.mock import patch, MagicMock, Mock
from pipeline import DocumentProcessor, DisambiguationEngine
from docx_handler import convert_paragraph_runs_safely


class TestDisambiguationEngine:
    """测试简繁字转换歧义引擎"""

    def test_convert_text_basic(self):
        """测试基础的简繁转换"""
        engine = DisambiguationEngine()
        # 基础转换测试
        assert engine.convert_text("干燥") == "乾燥"
        assert engine.convert_text("头发") == "頭髮"
        assert engine.convert_text("干部") == "幹部"

    def test_convert_text_empty_string(self):
        """测试空字符串输入"""
        engine = DisambiguationEngine()
        assert engine.convert_text("") == ""
        assert engine.convert_text(None) == ""

    def test_convert_text_no_change(self):
        """测试无需转换的文本"""
        engine = DisambiguationEngine()
        # 已经是繁体字
        assert engine.convert_text("頭髮") == "頭髮"
        # 只有英文字母
        assert engine.convert_text("hello") == "hello"

    def test_convert_text_ambiguity_resolution(self):
        """测试歧义词的智能识别"""
        engine = DisambiguationEngine()
        # 测试可能的歧义词 "干"
        # 应该根据上下文推断
        result = engine.convert_text("干什么")
        assert result == "幹什麼"  # "干”在语境中应该转换为"幹"

    def test_convert_text_with_onnx_model(self, mock_session):
        """测试启用 ONNX 模型的歧义识别"""
        engine = DisambiguationEngine()
        engine.session = mock_session
        
        # 测试 ONNX 模型的影响
        result = engine.convert_text("干什么")
        assert isinstance(result, str)

    def test_convert_text_without_onnx_model(self):
        """测试没有 ONNX 模型的情况"""
        engine = DisambiguationEngine()
        result = engine.convert_text("干燥")
        assert result == "乾燥"


class TestConvertParagraphRunsSafely:
    """测试安全的段落run转换功能"""

    def test_convert_paragraph_runs_safe(self):
        """测试安全的run转换逻辑"""
        mock_paragraph = Mock()
        mock_run1 = Mock()
        mock_run1.text = "干燥"
        mock_run2 = Mock()
        mock_run2.text = "的饼干"
        mock_paragraph.runs = [mock_run1, mock_run2]

        def mock_converter(text):
            return text.replace("干", "乾").replace("饼", "餅")

        convert_paragraph_runs_safely(mock_paragraph, mock_converter)
        
        # 验证run文本已被正确转换
        assert mock_run1.text == "乾燥"
        assert mock_run2.text == "的餅乾"

    def test_convert_paragraph_runs_empty(self):
        """测试空run的情况"""
        mock_paragraph = Mock()
        mock_paragraph.runs = []

        convert_paragraph_runs_safely(mock_paragraph, lambda x: x)
        # 应该直接返回，不抛出异常

    def test_convert_paragraph_runs_safe_length_mismatch(self):
        """测试字符长度不匹配的异常情况"""
        mock_paragraph = Mock()
        mock_run1 = Mock()
        mock_run1.text = "干燥"  # 2个字符
        mock_paragraph.runs = [mock_run1]

        def converter_with_mismatch(text):
            return "乾乾乾乾"  # 4个字符，不匹配

        convert_paragraph_runs_safely(mock_paragraph, converter_with_mismatch)
        
        # 应该使用异常兜底逻辑
        assert mock_run1.text == "乾乾乾乾"


class TestDocumentProcessorIntegration:
    """测试 DocumentProcessor 的集成功能"""

    @patch('docx.Document')
    @patch('pymupdf.open')
    def test_process_docx_integration(self, mock_pymupdf_open, mock_docx_document):
        """测试 DOCX 文件处理的集成"""
        # 模拟 docx
        mock_doc = Mock()
        mock_doc.paragraphs = []
        mock_doc.tables = []
        mock_doc.save = Mock()
        mock_docx_document.return_value = mock_doc

        # 模拟 pymupdf
        mock_doc_obj = MagicMock()
        mock_pymupdf_open.return_value = mock_doc_obj
        mock_doc_obj.__len__.return_value = 0

        engine = DisambiguationEngine()
        processor = DocumentProcessor(engine)

        with patch('os.path.exists', return_value=True):
            processor._process_docx("/tmp/test.docx", "/tmp/test_繁体.docx", None)

        # 验证 docx 被处理
        mock_doc.save.assert_called_once_with("/tmp/test_繁体.docx")

    @patch('pymupdf.open')
    def test_process_pdf_integration(self, mock_pymupdf_open):
        """测试 PDF 文件处理的集成"""
        # 模拟 pymupdf 打开的 PDF
        mock_doc = MagicMock()
        mock_page = MagicMock()
        mock_doc.__len__.return_value = 1
        mock_doc.__iter__.return_value = [mock_page]
        mock_doc.save = Mock()
        mock_doc.close = Mock()
        mock_pymupdf_open.return_value = mock_doc
        
        mock_page.get_text.return_value = [
            (0, 0, 100, 20, "测试文本", 0, 0)  # (x0, y0, x1, y1, text, block_no, block_type)
        ]

        engine = DisambiguationEngine()
        processor = DocumentProcessor(engine)

        with patch('os.path.exists', return_value=True):
            processor._process_pdf("/tmp/test.pdf", "/tmp/test_繁体.pdf", None)

        # 验证 PDF 被处理并保存
        mock_doc.save.assert_called_once_with("/tmp/test_繁体.pdf")

    @patch.object(DocumentProcessor, '_process_docx')
    @patch('subprocess.run')
    @patch('tempfile.TemporaryDirectory')
    @patch('shutil.copy2')
    def test_process_via_libreoffice_integration(self, mock_copy2, mock_temp_dir, mock_subprocess, mock_process_docx):
        """测试通过 LibreOffice 处理文档的集成"""
        mock_temp_instance = "/tmp/fake_dir"
        mock_temp_dir.return_value.__enter__.return_value = mock_temp_instance
        mock_subprocess.run.return_value = None
        mock_copy2.return_value = None

        engine = DisambiguationEngine()
        processor = DocumentProcessor(engine)

        with patch('os.path.exists', return_value=True):
            processor._process_via_libreoffice("/tmp/test.doc", "/tmp/test_繁体.doc", None)

        # 验证子过程被调用
        assert mock_subprocess.call_count >= 2

    @patch('pytesseract.image_to_data')
    @patch('PIL.Image.open')
    def test_process_jpg_integration(self, mock_pil_open, mock_tesseract):
        """测试 JPG 文件处理的集成"""
        mock_img = Mock()
        mock_img.size = (100, 100)
        mock_pil_open.return_value = mock_img
        mock_img.convert.return_value = mock_img

        mock_tesseract.return_value = {
            'level': [1, 2],
            'text': ['干燥', '的饼干'],
            'left': [0, 50],
            'top': [0, 0],
            'width': [10, 30],
            'height': [20, 20]
        }

        engine = DisambiguationEngine()
        processor = DocumentProcessor(engine)

        with patch('os.path.exists', return_value=True):
            processor._process_jpg("/tmp/test.jpg", "/tmp/test_繁体.jpg", None)

        # 验证图像保存
        mock_img.save.assert_called_once_with("/tmp/test_繁体.jpg")





# 辅助函数和模拟
@pytest.fixture
def mock_session():
    """创建模拟的 ONNX 会话"""
    mock_session = Mock()
    mock_session.run.return_value = [[0.9]]  # 模拟置信度分数
    return mock_session
