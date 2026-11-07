import os
import pytest
from PySide6.QtCore import QObject
from app_ui import ConvertWorker, TaskSignals

class SignalCatcher(QObject):
    def __init__(self):
        super().__init__()
        self.progress_records = []
        self.status_records = []

    def on_progress(self, row, val):
        self.progress_records.append((row, val))

    def on_status(self, row, text, err):
        self.status_records.append((row, text, err))


def test_convert_worker_success():
    def mock_converter(file_path, progress_cb):
        progress_cb(50)
        progress_cb(100)
    
    worker = ConvertWorker(row_idx=0, file_path="fake.docx", converter_func=mock_converter)
    catcher = SignalCatcher()
    
    worker.signals.progress.connect(catcher.on_progress)
    worker.signals.status.connect(catcher.on_status)
    
    worker.run()
    
    assert catcher.progress_records == [(0, 50), (0, 100), (0, 100)] # 50, 100 from mock, 100 from run end
    assert catcher.status_records == [(0, "处理中...", ""), (0, "已完成", "")]

def test_convert_worker_failure():
    def mock_converter(file_path, progress_cb):
        raise ValueError("mock error")
    
    worker = ConvertWorker(row_idx=1, file_path="fake.docx", converter_func=mock_converter)
    catcher = SignalCatcher()
    
    worker.signals.progress.connect(catcher.on_progress)
    worker.signals.status.connect(catcher.on_status)
    
    worker.run()
    
    assert len(catcher.progress_records) == 0
    assert catcher.status_records == [(1, "处理中...", ""), (1, "失败", "mock error")]


def test_queue_table_drag_and_drop(tmp_path):
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import QMimeData, QUrl, QPoint, Qt
    from PySide6.QtGui import QDragEnterEvent, QDropEvent
    from app_ui import TaskPage

    app = QApplication.instance() or QApplication([])
    page = TaskPage()

    # 确认不再有独立的 drop_area，队列区表格关闭自身拦截向上冒泡给 TaskPage，且重写了 event() 双保险
    assert not hasattr(page, "drop_area")
    assert page.acceptDrops() is True
    assert page.table.viewport().acceptDrops() is False

    test_file = tmp_path / "drag_test.docx"
    test_file.write_bytes(b"dummy")

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(test_file))])

    # 模拟真实 Qt6 QWidgetWindow 拖拽（初始状态为 ignore()，通过 QApplication.sendEvent 投递）
    drag_enter = QDragEnterEvent(
        QPoint(50, 50),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    drag_enter.ignore()
    QApplication.sendEvent(page.table, drag_enter)
    assert drag_enter.isAccepted() is True

    # 模拟释放文件
    drop_ev = QDropEvent(
        QPoint(50, 50),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    drop_ev.ignore()
    QApplication.sendEvent(page.table, drop_ev)
    assert drop_ev.isAccepted() is True
    assert page.table.rowCount() == 1
    assert len(page.tasks) == 1
    assert page.tasks[0]["path"] == str(test_file).replace("/", "\\") or os.path.normpath(page.tasks[0]["path"]) == os.path.normpath(str(test_file))


def test_preview_page_500ms_loading_modal():
    from PySide6.QtWidgets import QApplication
    from app_ui import PreviewPage

    app = QApplication.instance() or QApplication([])
    page = PreviewPage()

    assert page._render_timer.interval() == 500
    assert page._render_timer.isSingleShot() is True

    # 模拟后台渲染超过 500ms 触发 _show_loading_modal
    page._is_rendering_pages = True
    page._show_loading_modal()
    assert page._loading_modal is not None
    assert page._loading_modal.isVisible() is True

    # 模拟后台渲染完成回调自动关闭 Modal
    page._on_page_render_finished(None, [])
    assert page._is_rendering_pages is False
    assert page._loading_modal.isVisible() is False


def test_preview_page_image_rich_preview_preserves_characters(tmp_path):
    """测试预览页在展示图片富文本时，优先读取 sidecar 或 preview_data，不漏发、胡等字"""
    from PySide6.QtWidgets import QApplication
    from app_ui import PreviewPage

    app = QApplication.instance() or QApplication([])
    page = PreviewPage()

    test_img = tmp_path / "sample.png"
    test_img.write_bytes(b"dummy")
    sidecar_txt = tmp_path / "sample.png.ocr.txt"
    sidecar_txt.write_text("老胡打理頭髮", encoding="utf-8")

    page.current_ext = ".png"
    page.current_temp_path = str(test_img)
    page.current_original_path = str(test_img)
    page.preview_data = "老胡打理頭髮"

    page._refresh_rich_preview()
    rendered_text = page.preview_text.toPlainText()
    assert "老胡" in rendered_text
    assert "頭髮" in rendered_text



