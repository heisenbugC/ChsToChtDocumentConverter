import os
import sys
import shutil
from PySide6.QtCore import Qt, QObject, Signal, QRunnable, QThreadPool, QRect, QPoint, QEvent, QTimer, qInstallMessageHandler
from PySide6.QtGui import QIcon, QDragEnterEvent, QDropEvent, QPixmap, QFont, QPainter, QPen, QColor, QCursor
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidgetItem, QHeaderView, QFileDialog, QTextEdit,
    QLabel, QScrollArea, QSizePolicy, QListWidgetItem, QAbstractItemView, QDialog
)
try:
    from qfluentwidgets import TextEdit as QFTextEdit
except Exception:
    QFTextEdit = None
try:
    from qfluentwidgets import IndeterminateProgressBar
except Exception:
    IndeterminateProgressBar = None
from qfluentwidgets import (
    FluentWindow, NavigationInterface, NavigationItemPosition,
    PrimaryPushButton, PushButton, TableWidget, ProgressBar,
    InfoBar, InfoBarPosition, TitleLabel, SubtitleLabel,
    CardWidget, FluentIcon as FIF, PlainTextEdit, BodyLabel,
    ListWidget, MessageBoxBase, ComboBox, LineEdit, setCustomStyleSheet
)
import pipeline

# 尝试加载当前目录下的 disambiguation.onnx 作为后备模型
onnx_path = os.path.join(os.path.dirname(__file__), "disambiguation.onnx")
if not os.path.exists(onnx_path):
    onnx_path = None

engine = pipeline.DisambiguationEngine(onnx_model_path=onnx_path)
processor = pipeline.DocumentProcessor(engine)


# 1. 跨线程信号定义
class TaskSignals(QObject):
    progress = Signal(int, int)  # row_idx, percentage
    status = Signal(int, str, str)  # row_idx, status_text, error_msg
    result = Signal(int, dict)  # row_idx, result_dict from converter


# 2. 后台任务 Worker
class ConvertWorker(QRunnable):
    def __init__(self, row_idx: int, file_path: str, converter_func, return_result: bool = False):
        super().__init__()
        self.row_idx = row_idx
        self.file_path = file_path
        self.converter_func = converter_func
        self.signals = TaskSignals()
        self.return_result = return_result

    def run(self):
        try:
            self.signals.status.emit(self.row_idx, "处理中...", "")

            def progress_callback(percent: int):
                self.signals.progress.emit(self.row_idx, percent)

            # 调用真正的转换管道
            result = self.converter_func(self.file_path, progress_callback)

            self.signals.progress.emit(self.row_idx, 100)
            self.signals.status.emit(self.row_idx, "已完成", "")
            if self.return_result and isinstance(result, dict):
                self.signals.result.emit(self.row_idx, result)
            elif self.return_result:
                # 尝试传递空结果
                self.signals.result.emit(self.row_idx, {"original_path": self.file_path})
        except Exception as e:
            self.signals.status.emit(self.row_idx, "失败", str(e))


# 3. 拖拽卡片控件（保留兼容）与直接支持拖拽 + 背景说明文字的队列区表格控件
class DropArea(CardWidget):
    file_dropped = Signal(list)

    SUPPORTED_EXTS = ('.docx', '.doc', '.odt', '.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title = SubtitleLabel("请点击上方“添加文件”按钮导入文档（支持 DOCX/DOC/ODT/PDF/JPG/PNG）", self)
        self.title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.title.setAcceptDrops(False)
        layout.addWidget(self.title, 0, Qt.AlignmentFlag.AlignCenter)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = []
        unsupported = []
        for u in urls:
            if not u.isLocalFile():
                continue
            local_path = u.toLocalFile()
            if local_path.endswith('\\') or local_path.endswith('/'):
                local_path = local_path.rstrip('\\/')
            lower_path = local_path.lower()
            if lower_path.endswith(self.SUPPORTED_EXTS):
                files.append(local_path)
            else:
                unsupported.append(os.path.basename(local_path) or local_path)

        if files:
            event.acceptProposedAction()
            self.file_dropped.emit(files)
        else:
            event.ignore()

        if unsupported:
            try:
                parent = self.parent()
                InfoBar.warning(
                    title="不支持的文件格式",
                    content="仅支持 DOCX/DOC/ODT/PDF/JPG/JPEG/PNG/BMP/TIFF/TIF 格式",
                    position=InfoBarPosition.TOP,
                    parent=parent if parent is not None else self
                )
            except Exception:
                pass


class QueueTableWidget(TableWidget):
    """直接作为拖拽目的地的任务队列表格，支持全区域拖拽入队与居中背景说明文字"""
    file_dropped = Signal(list)

    SUPPORTED_EXTS = ('.docx', '.doc', '.odt', '.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
        # 策略1（最彻底）：将表格本体、viewport、表头全部设为 acceptDrops(False)，
        # 使 Windows OLE 拖拽事件跳过 QAbstractScrollArea 100% 向上冒泡至父级 TaskPage(QWidget)！
        self.setAcceptDrops(False)
        self.viewport().setAcceptDrops(False)
        self.horizontalHeader().setAcceptDrops(False)
        self.horizontalHeader().viewport().setAcceptDrops(False)
        self.verticalHeader().setAcceptDrops(False)
        self.verticalHeader().viewport().setAcceptDrops(False)
        self.viewport().installEventFilter(self)

    def _check_drag_event(self, event) -> bool:
        if event.mimeData() and event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
            return True
        event.ignore()
        return False

    def _handle_drop_event(self, event) -> bool:
        if not (event.mimeData() and event.mimeData().hasUrls()):
            event.ignore()
            return False
        urls = event.mimeData().urls()
        files = []
        unsupported = []
        for u in urls:
            if not u.isLocalFile():
                continue
            local_path = u.toLocalFile()
            if local_path.endswith('\\') or local_path.endswith('/'):
                local_path = local_path.rstrip('\\/')
            lower_path = local_path.lower()
            if lower_path.endswith(self.SUPPORTED_EXTS):
                files.append(local_path)
            else:
                unsupported.append(os.path.basename(local_path) or local_path)

        if files:
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
            self.file_dropped.emit(files)
        else:
            event.ignore()

        if unsupported:
            try:
                parent = self.parent()
                InfoBar.warning(
                    title="不支持的文件格式",
                    content="仅支持 DOCX/DOC/ODT/PDF/JPG/JPEG/PNG/BMP/TIFF/TIF 格式",
                    position=InfoBarPosition.TOP,
                    parent=parent if parent is not None else self
                )
            except Exception:
                pass
        return True

    # 策略2（双保险）：重写 event() 绕过 QAbstractScrollArea::event() 对 Drag 事件直接 return false 的硬编码
    def event(self, e):
        etype = e.type()
        if etype in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            return self._check_drag_event(e)
        if etype == QEvent.Type.DragLeave:
            e.accept()
            return True
        if etype == QEvent.Type.Drop:
            return self._handle_drop_event(e)
        return super().event(e)

    def viewportEvent(self, event):
        etype = event.type()
        if etype in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            return self._check_drag_event(event)
        if etype == QEvent.Type.DragLeave:
            event.accept()
            return True
        if etype == QEvent.Type.Drop:
            return self._handle_drop_event(event)
        return super().viewportEvent(event)

    def eventFilter(self, obj, event):
        if obj is self.viewport():
            etype = event.type()
            if etype in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
                return self._check_drag_event(event)
            if etype == QEvent.Type.DragLeave:
                event.accept()
                return True
            if etype == QEvent.Type.Drop:
                return self._handle_drop_event(event)
        return super().eventFilter(obj, event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        self._check_drag_event(event)

    def dragMoveEvent(self, event):
        self._check_drag_event(event)

    def dropEvent(self, event: QDropEvent):
        self._handle_drop_event(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        try:
            painter = QPainter(self.viewport())
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setPen(QColor(148, 163, 184))  # #94a3b8 柔和背景说明文字色
            font = QFont("Microsoft YaHei UI", 12)
            painter.setFont(font)
            bg_text = "请点击上方“添加文件”按钮添加待转换文档\n（支持 DOCX / DOC / ODT / PDF / JPG / PNG / BMP / TIFF）"
            vp_rect = self.viewport().rect()
            if self.rowCount() == 0:
                painter.drawText(vp_rect, Qt.AlignmentFlag.AlignCenter, bg_text)
            else:
                used_h = sum(self.rowHeight(r) for r in range(self.rowCount()))
                if vp_rect.height() - used_h > 80:
                    bottom_rect = vp_rect.adjusted(0, used_h, 0, 0)
                    painter.drawText(bottom_rect, Qt.AlignmentFlag.AlignCenter, bg_text)
            painter.end()
        except Exception:
            pass


# 4. 预览页面 (Preview Session)
if QFTextEdit is not None:
    _HighlightBase = QFTextEdit
else:
    try:
        from PySide6.QtWidgets import QTextEdit as _QTextEditBase

        _HighlightBase = _QTextEditBase
    except Exception:
        _HighlightBase = PlainTextEdit


class HighlightPlainTextEdit(_HighlightBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        try:
            self.setFont(QFont("Microsoft YaHei UI", 11))
        except Exception:
            pass
        self.hover_red_boxes = []
        self.hover_word = []

    def setRichHtml(self, html_str: str):
        try:
            if html_str is None:
                html_str = ""
            self.setHtml(html_str)
        except Exception:
            try:
                self.setPlainText(str(html_str))
            except Exception:
                pass

    def paintEvent(self, event):
        try:
            super().paintEvent(event)
        except Exception:
            pass
        if self.hover_red_boxes and self.hover_word:
            try:
                painter = QPainter(self.viewport())
                pen = QPen(QColor(255, 0, 0), 2)
                painter.setPen(pen)
                text_edit = self
                for pos in self.hover_red_boxes:
                    try:
                        c_start = text_edit.textCursor()
                        c_start.setPosition(pos)
                        r_start = text_edit.cursorRect(c_start)

                        c_end = text_edit.textCursor()
                        c_end.setPosition(pos + len(self.hover_word))
                        r_end = text_edit.cursorRect(c_end)

                        if abs(r_start.top() - r_end.top()) <= max(4, r_start.height() // 2):
                            rect = QRect(
                                r_start.left(),
                                min(r_start.top(), r_end.top()),
                                max(6, r_end.left() - r_start.left()),
                                max(r_start.height(), r_end.height())
                            ).adjusted(-2, -1, 2, 1)
                            painter.drawRect(rect)
                        else:
                            for ch_i in range(len(self.hover_word)):
                                try:
                                    c1 = text_edit.textCursor()
                                    c1.setPosition(pos + ch_i)
                                    r1 = text_edit.cursorRect(c1)
                                    c2 = text_edit.textCursor()
                                    c2.setPosition(pos + ch_i + 1)
                                    r2 = text_edit.cursorRect(c2)
                                    w = r2.left() - r1.left() if abs(r1.top() - r2.top()) <= 4 and r2.left() > r1.left() else r1.height()
                                    rect = QRect(r1.left(), r1.top(), max(6, w), r1.height()).adjusted(-1, -1, 1, 1)
                                    painter.drawRect(rect)
                                except Exception:
                                    continue
                    except Exception:
                        continue
                painter.end()
            except Exception:
                pass


class PageRenderSignals(QObject):
    finished = Signal(str, list)  # temp_path, rendered_page_paths


class PageRenderWorker(QRunnable):
    """后台异步渲染原貌页 PNG 的工作器，避免阻塞 UI 主线程造成假死"""
    def __init__(self, temp_path: str):
        super().__init__()
        self.temp_path = temp_path
        self.signals = PageRenderSignals()

    def run(self):
        paths = []
        try:
            if self.temp_path and hasattr(processor, 'render_page_images'):
                res = processor.render_page_images(self.temp_path)
                if res:
                    paths = list(res)
        except Exception:
            paths = []
        self.signals.finished.emit(self.temp_path, paths)


class PageLoadingModal(QDialog):
    """当切换原貌页渲染耗时超过 500ms 时弹出的模态加载提示窗（采用实色背景与高对比度边框）"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("正在生成原貌页预览")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        self.setFixedSize(420, 160)
        dlg_qss = (
            "PageLoadingModal, QDialog { background-color: #ffffff; border: 2px solid #0284c7; border-radius: 8px; }"
        )
        setCustomStyleSheet(self, dlg_qss, dlg_qss)
        self.setStyleSheet(dlg_qss)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        card = QWidget(self)
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card.setAutoFillBackground(True)
        card_qss = (
            "QWidget { background-color: #f8fafc; border: 2px solid #0284c7; border-radius: 8px; }"
        )
        setCustomStyleSheet(card, card_qss, card_qss)
        card.setStyleSheet(card_qss)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(28, 22, 28, 22)
        card_layout.setSpacing(12)

        self.title_lbl = SubtitleLabel("⏳ 正在生成原貌页预览...", card)
        self.title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_qss = "QLabel { background-color: transparent; color: #0f172a; font-weight: bold; border: none; }"
        setCustomStyleSheet(self.title_lbl, title_qss, title_qss)
        self.title_lbl.setStyleSheet(title_qss)
        card_layout.addWidget(self.title_lbl)

        self.desc_lbl = BodyLabel("正在按源文档页面尺寸与排版渲染高清页面，请稍候...", card)
        self.desc_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_qss = "QLabel { background-color: transparent; color: #334155; border: none; }"
        setCustomStyleSheet(self.desc_lbl, desc_qss, desc_qss)
        self.desc_lbl.setStyleSheet(desc_qss)
        card_layout.addWidget(self.desc_lbl)

        if IndeterminateProgressBar is not None:
            self.pbar = IndeterminateProgressBar(card)
            try:
                self.pbar.start()
            except Exception:
                pass
        else:
            self.pbar = ProgressBar(card)
            self.pbar.setRange(0, 0)
        self.pbar.setMinimumWidth(320)
        card_layout.addWidget(self.pbar, 0, Qt.AlignmentFlag.AlignCenter)

        main_layout.addWidget(card)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            event.ignore()
        else:
            super().keyPressEvent(event)


class PreviewPage(QWidget):
    redo_requested = Signal(str)  # original_path
    save_requested = Signal(str, str)  # original_path, target_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_original_path = None
        self.current_temp_path = None
        self.current_ext = None
        self.preview_data = ""
        self.suspects = []
        self.preview_mode = "rich"  # "rich" or "page"
        self.rendered_page_paths = []
        self._is_rendering_pages = False
        self._loading_modal = None
        self._render_timer = QTimer(self)
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(500)
        self._render_timer.timeout.connect(self._show_loading_modal)
        self.init_ui()
        try:
            self.preview_text.viewport().setMouseTracking(True)
        except Exception:
            pass

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(30, 20, 30, 20)
        self.layout.setSpacing(15)

        self.title = TitleLabel("预览转换结果", self)
        self.layout.addWidget(self.title)

        self.subtitle = SubtitleLabel("文档转换预览页（Interactive Spell Checker）", self)
        self.layout.addWidget(self.subtitle)

        # 预览卡片 (分为左右两部分)
        preview_card = CardWidget(self)
        preview_layout = QHBoxLayout(preview_card)
        preview_layout.setContentsMargins(10, 10, 10, 10)
        preview_layout.setSpacing(10)

        # 左侧：主窗格 (文本编辑器)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        self.file_info_label = BodyLabel("未加载预览", left_widget)
        left_layout.addWidget(self.file_info_label)
        # 模式切换栏
        mode_bar_layout = QHBoxLayout()
        try:
            self.rich_mode_btn = PrimaryPushButton("📝 极速富文本预览 (表格/卡片)", left_widget)
        except Exception:
            self.rich_mode_btn = PushButton("📝 极速富文本预览 (表格/卡片)", left_widget)
        self.page_mode_btn = PushButton("🖼️ 按需原貌页渲染", left_widget)
        self.rich_mode_btn.clicked.connect(lambda: self.switch_preview_mode("rich"))
        self.page_mode_btn.clicked.connect(lambda: self.switch_preview_mode("page"))
        mode_bar_layout.addWidget(self.rich_mode_btn)
        mode_bar_layout.addWidget(self.page_mode_btn)
        mode_bar_layout.addStretch(1)
        left_layout.addLayout(mode_bar_layout)
        self.preview_text = HighlightPlainTextEdit(left_widget)
        try:
            self.preview_text.setReadOnly(True)
        except Exception:
            pass
        try:
            self.preview_text.setPlaceholderText("转换后的文本预览将在此显示...\n对于 PDF/JPG 等格式，将显示临时文件路径及查看提示。")
        except Exception:
            pass
        self.preview_text.setMinimumHeight(260)
        left_layout.addWidget(self.preview_text)
        # 原貌页渲染滚动区域
        self.page_scroll_area = QScrollArea(left_widget)
        self.page_scroll_area.setWidgetResizable(True)
        self.page_scroll_container = QWidget()
        self.page_scroll_layout = QVBoxLayout(self.page_scroll_container)
        self.page_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.page_scroll_layout.setSpacing(10)
        self.page_scroll_area.setWidget(self.page_scroll_container)
        self.page_scroll_area.setVisible(False)
        left_layout.addWidget(self.page_scroll_area)
        self.image_label = QLabel(left_widget)
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.image_label.setVisible(False)
        left_layout.addWidget(self.image_label)
        preview_layout.addWidget(left_widget, stretch=3)

        # 右侧：辅窗格 (嫌疑字词清单)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_title = BodyLabel("待检字词 (可能转换错误)", right_widget)
        right_layout.addWidget(right_title)
        self.suspect_list = ListWidget(right_widget)
        self.suspect_list.setMouseTracking(True)
        self.suspect_list.viewport().setMouseTracking(True)
        self.suspect_list.itemClicked.connect(self.on_suspect_clicked)
        self.suspect_list.itemEntered.connect(self.on_suspect_hover)
        self.suspect_list.setAttribute(Qt.WidgetAttribute.WA_Hover)
        # 鼠标离开列表时清除红框
        try:
            self.suspect_list.leaveEvent = self._on_suspect_list_leave
        except Exception:
            pass
        right_layout.addWidget(self.suspect_list)
        preview_layout.addWidget(right_widget, stretch=1)

        self.layout.addWidget(preview_card)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.save_btn = PrimaryPushButton(FIF.SAVE, "保存输出到指定目录", self)
        self.redo_btn = PushButton(FIF.SYNC, "重做转换", self)
        self.back_btn = PushButton(FIF.RETURN, "返回转换列表", self)

        self.save_btn.clicked.connect(self.on_save_clicked)
        self.redo_btn.clicked.connect(self.on_redo_clicked)
        self.back_btn.clicked.connect(self.on_back_clicked)

        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.redo_btn)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.back_btn)
        self.layout.addLayout(btn_layout)

    def _refresh_rich_preview(self):
        try:
            if not self.current_temp_path and not self.current_original_path:
                return
            rich_html = None
            is_image = bool(self.current_ext and self.current_ext.lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'])

            if is_image and self.current_temp_path and os.path.exists(self.current_temp_path + ".ocr.txt"):
                if hasattr(processor, 'generate_rich_preview_html'):
                    try:
                        rich_html = processor.generate_rich_preview_html(self.current_temp_path)
                    except Exception:
                        rich_html = None
            elif is_image and self.current_original_path and os.path.exists(self.current_original_path):
                if hasattr(processor, 'generate_rich_preview_html'):
                    try:
                        rich_html = processor.generate_rich_preview_html(self.current_original_path)
                    except Exception:
                        rich_html = None
            else:
                target_path = self.current_temp_path or self.current_original_path
                if hasattr(processor, 'generate_rich_preview_html'):
                    try:
                        rich_html = processor.generate_rich_preview_html(target_path)
                    except Exception:
                        rich_html = None
                if not rich_html and hasattr(processor, '_generate_rich_preview_html'):
                    try:
                        rich_html = processor._generate_rich_preview_html(target_path)
                    except Exception:
                        rich_html = None

            if not rich_html and self.preview_data:
                import html as _html
                rich_html = "<div style='font-family:\"Microsoft YaHei UI\",\"Microsoft YaHei\",sans-serif; font-size:11pt; color:#1e293b;'>" + "".join(
                    f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(line)}</p>'
                    for line in self.preview_data.splitlines() if line.strip()
                ) + "</div>"

            if rich_html:
                try:
                    self.preview_text.setRichHtml(rich_html)
                    return
                except Exception:
                    try:
                        self.preview_text.setHtml(rich_html)
                        return
                    except Exception:
                        pass
            # fallback
            if self.preview_data:
                try:
                    self.preview_text.setPlainText(self.preview_data)
                except Exception:
                    pass
        except Exception:
            pass

    def switch_preview_mode(self, mode: str):
        try:
            self.preview_mode = mode
            is_img = bool(self.current_ext and self.current_ext.lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'])
            if mode == "rich":
                try:
                    self.rich_mode_btn.setEnabled(False)
                    self.page_mode_btn.setEnabled(True)
                except Exception:
                    pass
                self.image_label.setVisible(False)
                self.page_scroll_area.setVisible(False)
                self.preview_text.setMaximumHeight(16777215)
                self.preview_text.setMinimumHeight(260)
                self.preview_text.setVisible(True)
                self._refresh_rich_preview()
                return
            if mode == "page":
                try:
                    self.rich_mode_btn.setEnabled(True)
                    self.page_mode_btn.setEnabled(False)
                except Exception:
                    pass
                self.image_label.setVisible(False)
                self.preview_text.setVisible(False)
                self.page_scroll_area.setVisible(True)
                self._load_page_images()
                return
        except Exception:
            pass

    def _clear_page_scroll(self):
        try:
            while self.page_scroll_layout.count():
                item = self.page_scroll_layout.takeAt(0)
                w = item.widget()
                if w:
                    w.deleteLater()
        except Exception:
            pass

    def _show_loading_modal(self):
        try:
            if not self._is_rendering_pages:
                return
            if self._loading_modal is None:
                self._loading_modal = PageLoadingModal(self.window())
            win = self.window()
            if win is not None and win.isVisible():
                geo = win.geometry()
                self._loading_modal.move(
                    geo.center().x() - self._loading_modal.width() // 2,
                    geo.center().y() - self._loading_modal.height() // 2
                )
            self._loading_modal.show()
            self._loading_modal.raise_()
        except Exception:
            pass

    def _hide_loading_modal(self):
        try:
            self._render_timer.stop()
            if self._loading_modal is not None:
                self._loading_modal.hide()
                self._loading_modal.close()
        except Exception:
            pass

    def _populate_page_images_ui(self, paths: list[str]):
        try:
            self._clear_page_scroll()
            if not paths:
                empty_lbl = BodyLabel(
                    "未能生成原貌页预览，请确认系统已安装 Microsoft Word 或 LibreOffice。",
                    self.page_scroll_container
                )
                empty_lbl.setStyleSheet("color: #64748b; padding: 24px;")
                self.page_scroll_layout.addWidget(empty_lbl, 0, Qt.AlignmentFlag.AlignCenter)
                self.page_scroll_layout.addStretch(1)
                return

            # 根据滚动区域可视宽度与源页面真实宽高比（横版/竖版自适应）计算每页显示尺寸，绝不挤压截断高度
            vp_w = self.page_scroll_area.viewport().width() if hasattr(self, "page_scroll_area") else 760
            base_target_w = max(560, min(920, vp_w - 36)) if vp_w > 100 else 720
            for page_idx, p in enumerate(paths):
                try:
                    if not p or not os.path.exists(p):
                        continue
                    pixmap = QPixmap(p)
                    if pixmap.isNull():
                        continue
                    src_w = pixmap.width()
                    src_h = pixmap.height()
                    is_landscape = src_w > src_h
                    orient_tag = "横向页面" if is_landscape else "纵向页面"
                    # 对横向大图（特别是多页并排横版图）赋予充足显示宽度，支持横向/纵向平滑滚动查看清晰细节
                    if is_landscape and src_w > src_h * 1.35:
                        page_target_w = max(base_target_w, min(src_w, int(base_target_w * min(2.4, src_w / max(1, src_h)))))
                    else:
                        page_target_w = base_target_w
                    # 页码与页面方向提示条
                    hdr = BodyLabel(f"📄 第 {page_idx + 1} / {len(paths)} 页 （{orient_tag} {src_w}×{src_h} px）", self.page_scroll_container)
                    hdr.setStyleSheet("color: #475569; font-weight: bold; padding: 6px 4px 2px 4px;")
                    self.page_scroll_layout.addWidget(hdr)

                    lbl = QLabel(self.page_scroll_container)
                    lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    scaled = pixmap.scaledToWidth(page_target_w, Qt.TransformationMode.SmoothTransformation)
                    lbl.setPixmap(scaled)
                    # 关键修复：将 QLabel 的固定高度设为按源页比例缩放后的完整高度，防止 QVBoxLayout 将其压缩到 400px 导致上下内容被截断！
                    lbl.setFixedSize(scaled.width() + 8, scaled.height() + 8)
                    lbl.setStyleSheet("background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; padding: 4px;")
                    self.page_scroll_layout.addWidget(lbl, 0, Qt.AlignmentFlag.AlignHCenter)
                except Exception:
                    continue
            self.page_scroll_layout.addStretch(1)
        except Exception:
            pass

    def _on_page_render_finished(self, temp_path: str, paths: list[str]):
        try:
            self._is_rendering_pages = False
            self._hide_loading_modal()
            if temp_path != self.current_temp_path:
                return
            self.rendered_page_paths = list(paths or [])
            if self.preview_mode == "page":
                self._populate_page_images_ui(self.rendered_page_paths)
        except Exception:
            self._is_rendering_pages = False
            self._hide_loading_modal()

    def _load_page_images(self):
        try:
            self._clear_page_scroll()
            if not self.current_temp_path:
                return
            # 若已有缓存的原貌页图像，直接极速渲染（不走后台任务和 500ms 弹窗）
            if self.rendered_page_paths:
                self._populate_page_images_ui(self.rendered_page_paths)
                return

            # 若当前窗口可见，启动后台异步渲染 + 500ms 延迟模态加载弹窗（避免 UI 假死）
            if self.isVisible():
                if self._is_rendering_pages:
                    return
                self._is_rendering_pages = True
                self._render_timer.start(500)
                worker = PageRenderWorker(self.current_temp_path)
                worker.signals.finished.connect(self._on_page_render_finished)
                QThreadPool.globalInstance().start(worker)
            else:
                # 无头测试或未显示窗口时的同步回退
                if hasattr(processor, 'render_page_images'):
                    try:
                        paths = processor.render_page_images(self.current_temp_path)
                        if paths:
                            self.rendered_page_paths = list(paths)
                    except Exception:
                        self.rendered_page_paths = []
                self._populate_page_images_ui(self.rendered_page_paths)
        except Exception:
            self._is_rendering_pages = False
            self._hide_loading_modal()

    def on_suspect_clicked(self, item):
        if item is None:
            return
        suspect = item.data(Qt.ItemDataRole.UserRole)
        if suspect is None:
            # 尝试兼容旧数据
            suspect = item.data(Qt.UserRole)
        if not isinstance(suspect, dict):
            return

        class CorrectionDialog(MessageBoxBase):
            def __init__(self, suspect_data, parent=None):
                super().__init__(parent)
                # qfluentwidgets 的 MessageBoxBase 并没有预置 self.titleLabel，需自行创建
                self.titleLabel = SubtitleLabel("拼写检查与修正", self)
                self.viewLayout.addWidget(self.titleLabel)

                self.orig_box = QLabel(f"原文词汇： {suspect_data['simplified']}", self)
                orig_qss = (
                    "QLabel { background-color: #f1f5f9; color: #000000; font-weight: bold; "
                    "font-size: 15px; border: 1px solid #94a3b8; border-radius: 6px; padding: 8px 12px; }"
                )
                setCustomStyleSheet(self.orig_box, orig_qss, orig_qss)
                self.orig_box.setStyleSheet(orig_qss)
                self.viewLayout.addWidget(self.orig_box)

                self.curr_box = QLabel(f"当前转换： {suspect_data['default']}", self)
                curr_qss = (
                    "QLabel { background-color: #bae6fd; color: #000000; font-weight: bold; "
                    "font-size: 16px; border: 2px solid #0284c7; border-radius: 6px; padding: 10px 12px; }"
                )
                setCustomStyleSheet(self.curr_box, curr_qss, curr_qss)
                self.curr_box.setStyleSheet(curr_qss)
                self.viewLayout.addWidget(self.curr_box)

                self.combo = ComboBox()
                combo_qss = (
                    "ComboBox { background-color: #bae6fd; color: #000000; font-weight: bold; "
                    "font-size: 15px; border: 2px solid #0284c7; border-radius: 6px; padding: 6px 31px 6px 12px; } "
                    "ComboBox:hover { background-color: #7dd3fc; color: #000000; border: 2px solid #0369a1; } "
                    "ComboBox:pressed { background-color: #38bdf8; color: #000000; }"
                )
                setCustomStyleSheet(self.combo, combo_qss, combo_qss)
                self.combo.setStyleSheet(combo_qss)
                options = suspect_data.get('options') or [suspect_data.get('default', '')]
                self.combo.addItems(options)
                self.viewLayout.addWidget(self.combo)

                self.custom_edit = LineEdit()
                self.custom_edit.setPlaceholderText("或手动输入正确翻译...")
                le_qss = (
                    "LineEdit { background-color: #e0f2fe; color: #000000; font-weight: bold; "
                    "font-size: 14px; border: 1px solid #0284c7; border-radius: 6px; padding: 4px 10px; } "
                    "LineEdit:focus { background-color: #bae6fd; color: #000000; border: 2px solid #0284c7; }"
                )
                setCustomStyleSheet(self.custom_edit, le_qss, le_qss)
                self.custom_edit.setStyleSheet(le_qss)
                self.viewLayout.addWidget(self.custom_edit)
                self.widget.setMinimumWidth(360)

            def get_correction(self):
                txt = self.custom_edit.text()
                if txt:
                    return txt
                return self.combo.currentText()

        w = CorrectionDialog(suspect, self.window())
        if w.exec():
            correction = w.get_correction()
            if not correction:
                return
            # 去掉 if correction != suspect.get('default'): 的限制，只要 correction 非空就执行
            engine.user_overrides[suspect['simplified']] = correction
            # 实时替换刷新当前预览文本和列表项显示
            try:
                refreshed_via_engine = False
                # 如果有原始文件和临时文件，刷新临时文件（支持多种格式）
                if self.current_original_path and self.current_temp_path and self.current_ext:
                    try:
                        ext = self.current_ext
                        if ext == '.docx':
                            processor._process_docx(self.current_original_path, self.current_temp_path, None)
                        elif ext == '.pdf':
                            processor._process_pdf(self.current_original_path, self.current_temp_path, None)
                        elif ext in ['.doc', '.odt']:
                            processor._process_via_libreoffice(self.current_original_path, self.current_temp_path, None)
                        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp']:
                            processor._process_jpg(self.current_original_path, self.current_temp_path, None)
                            self.rendered_page_paths = []
                            if self.preview_mode == "page":
                                self._load_page_images()
                        # 尝试用原文重新生成完整预览文本
                        try:
                            orig_text = processor.extract_text(self.current_original_path)
                            if orig_text:
                                new_preview = engine.convert_text(orig_text)
                                if new_preview:
                                    try:
                                        if self.preview_mode == "rich":
                                            self._refresh_rich_preview()
                                        else:
                                            self.preview_text.setPlainText(new_preview)
                                    except Exception:
                                        self.preview_text.setPlainText(new_preview)
                                    refreshed_via_engine = True
                        except Exception:
                            pass
                    except Exception:
                        pass
                if not refreshed_via_engine:
                    # 更新当前预览文本中对应的词
                    try:
                        if self.preview_mode == "rich":
                            self._refresh_rich_preview()
                            # 更新列表项显示后继续
                        else:
                            target_word = suspect.get('converted') or suspect.get('default') or ''
                            try:
                                cur = self.preview_text.toPlainText()
                            except Exception:
                                cur = ""
                            if target_word and target_word in cur:
                                new_text = cur.replace(target_word, correction)
                                self.preview_text.setPlainText(new_text)
                            elif suspect.get('simplified') and suspect.get('simplified') in cur:
                                new_text = cur.replace(suspect['simplified'], correction)
                                self.preview_text.setPlainText(new_text)
                    except Exception:
                        pass
                # 刷新原貌页（如在页面模式下）
                try:
                    if self.preview_mode == "page":
                        self.rendered_page_paths = []
                        self._load_page_images()
                except Exception:
                    pass
                # 更新列表项显示
                item.setText(f"{suspect['simplified']} -> {correction}")
                suspect['default'] = correction
                suspect['converted'] = correction  # 同时更新 converted
                item.setData(Qt.ItemDataRole.UserRole, suspect)
                try:
                    self.preview_data = self.preview_text.toPlainText()
                except Exception:
                    pass
            except Exception:
                pass
            InfoBar.success("已记录", f"'{suspect['simplified']}' 将被转换为 '{correction}'，已实时更新预览。", parent=self)

    def load_preview(self, result_dict: dict):
        self._hide_loading_modal()
        self._is_rendering_pages = False
        self.current_original_path = result_dict.get("original_path")
        self.current_temp_path = result_dict.get("temp_output_path")
        self.preview_data = result_dict.get("preview_data", "")
        self.suspects = list(result_dict.get("suspects", []))
        self.rendered_page_paths = []
        self.preview_mode = "rich"
        try:
            self.rich_mode_btn.setEnabled(False)
            self.page_mode_btn.setEnabled(True)
        except Exception:
            pass
        # 清除 hover 红框
        try:
            self.preview_text.hover_red_boxes = []
            self.preview_text.hover_word = ""
            try:
                self.preview_text.viewport().update()
            except Exception:
                pass
        except Exception:
            pass

        if self.current_original_path:
            self.current_ext = os.path.splitext(self.current_original_path)[1].lower()
            fname = os.path.basename(self.current_original_path)
            self.file_info_label.setText(f"原始文件：{fname} | 临时输出：{self.current_temp_path or 'N/A'}")
        else:
            self.current_ext = None
            self.file_info_label.setText("预览信息")

        # 清空预览和列表
        try:
            self.preview_text.clear()
        except Exception:
            pass
        self.image_label.clear()
        self.image_label.setVisible(False)
        try:
            self._clear_page_scroll()
        except Exception:
            pass
        self.page_scroll_area.setVisible(False)
        self.preview_text.setVisible(True)
        self.suspect_list.clear()

        # 按文章出现顺序排序待检字词
        full_text_for_order = self.preview_data or ""
        try:
            if self.current_original_path:
                if self.current_ext == '.docx':
                    try:
                        import docx
                        doc = docx.Document(self.current_original_path)
                        full_text_for_order = "\n".join([p.text for p in doc.paragraphs])
                    except Exception:
                        pass
                elif self.current_ext == '.pdf':
                    try:
                        import pymupdf
                        doc = pymupdf.open(self.current_original_path)
                        ft = ""
                        for page in doc:
                            ft += page.get_text()
                        full_text_for_order = ft or full_text_for_order
                        doc.close()
                    except Exception:
                        pass
                elif self.current_ext in ['.odt', '.doc']:
                    try:
                        extracted = processor.extract_text(self.current_original_path)
                        if extracted:
                            full_text_for_order = extracted
                    except Exception:
                        pass
        except Exception:
            pass

        def _get_first_pos(s):
            try:
                if isinstance(s, dict) and 'first_pos' in s and s['first_pos'] is not None:
                    return int(s['first_pos'])
            except Exception:
                pass
            # 根据在原文或预览文本中首次出现的 index 排序
            simp = s.get('simplified', '') if isinstance(s, dict) else ''
            if not simp:
                return 10**12
            try:
                idx = full_text_for_order.find(simp)
                if idx >= 0:
                    return idx
            except Exception:
                pass
            try:
                idx = self.preview_data.find(simp)
                if idx >= 0:
                    return idx
            except Exception:
                pass
            return 10**12

        try:
            sorted_suspects = sorted(self.suspects, key=_get_first_pos)
        except Exception:
            sorted_suspects = list(self.suspects)

        # 填充嫌疑词汇（统一使用 Qt.ItemDataRole.UserRole）
        seen = set()
        for suspect in sorted_suspects:
            if not isinstance(suspect, dict):
                continue
            simp_key = suspect.get("simplified")
            if not simp_key or simp_key in seen:
                continue
            seen.add(simp_key)
            item = QListWidgetItem(f"{simp_key} -> {suspect.get('default', '')}")
            item.setData(Qt.ItemDataRole.UserRole, dict(suspect))
            self.suspect_list.addItem(item)

        if self.current_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp']:
            # 图片预览：默认显示完整的富文本提取与转换结果（不压缩高度，也不被 520px 缩略图挤压），切换到“按需原貌页渲染”时在可滚动高清视图中展示原位覆盖繁体图
            self.image_label.setVisible(False)
            self.image_label.clear()
            try:
                self.preview_text.setMaximumHeight(16777215)
                self.preview_text.setMinimumHeight(260)
            except Exception:
                pass
            try:
                self._refresh_rich_preview()
            except Exception:
                pass
            if not (hasattr(self.preview_text, 'toHtml') and self.preview_text.toHtml()):
                if self.preview_data:
                    try:
                        self.preview_text.setPlainText(self.preview_data)
                    except Exception:
                        pass
                elif self.current_temp_path and os.path.exists(self.current_temp_path):
                    try:
                        extracted = processor.extract_text(self.current_temp_path) or (processor.extract_text(self.current_original_path) if self.current_original_path else "")
                        if extracted:
                            self.preview_text.setPlainText(extracted)
                        else:
                            self.preview_text.setPlainText("图片已转换，临时文件已生成。")
                    except Exception:
                        self.preview_text.setPlainText("图片已转换，临时文件已生成。")
                else:
                    try:
                        self.preview_text.setPlainText("图片已转换，临时文件已生成。")
                    except Exception:
                        pass
        elif self.current_ext == '.pdf':
            self.image_label.setVisible(False)
            self.image_label.clear()
            try:
                self.preview_text.setMaximumHeight(16777215)
                self.preview_text.setMinimumHeight(260)
            except Exception:
                pass
            # 默认使用富文本预览
            try:
                self._refresh_rich_preview()
            except Exception:
                pass
            if not (hasattr(self.preview_text, 'toHtml') and self.preview_text.toHtml()):
                if self.preview_data:
                    try:
                        self.preview_text.setPlainText(self.preview_data)
                    except Exception:
                        pass
                    try:
                        self.file_info_label.setText((self.file_info_label.text() or "") + " | PDF 已转换并保存到临时目录，可点击“保存输出到指定目录”保存")
                    except Exception:
                        pass
                else:
                    try:
                        self.preview_text.setPlainText(
                            (self.preview_data or "") + "\n\n【预览说明】PDF 已转换并保存到临时目录。\n可点击“保存输出到指定目录”选择目标路径后保存，或使用系统查看器打开临时文件：\n" + (self.current_temp_path or "")
                        )
                    except Exception:
                        pass
        elif self.current_ext in ['.doc', '.odt']:
            self.image_label.setVisible(False)
            self.image_label.clear()
            try:
                self.preview_text.setMaximumHeight(16777215)
                self.preview_text.setMinimumHeight(260)
            except Exception:
                pass
            try:
                self._refresh_rich_preview()
            except Exception:
                pass
            if not (hasattr(self.preview_text, 'toHtml') and self.preview_text.toHtml()):
                if self.preview_data:
                    try:
                        self.preview_text.setPlainText(self.preview_data)
                    except Exception:
                        pass
                    try:
                        self.file_info_label.setText((self.file_info_label.text() or "") + " | 已通过 LibreOffice 流程转换")
                    except Exception:
                        pass
                else:
                    try:
                        self.preview_text.setPlainText(
                            (self.preview_data or "") + "\n\n【预览说明】已通过 LibreOffice 流程转换。\n临时输出路径：" + (self.current_temp_path or "")
                        )
                    except Exception:
                        pass
        else:
            # docx 等文本类，优先使用富文本预览
            self.image_label.setVisible(False)
            self.image_label.clear()
            try:
                self.preview_text.setMaximumHeight(16777215)
                self.preview_text.setMinimumHeight(260)
            except Exception:
                pass
            try:
                self._refresh_rich_preview()
            except Exception:
                pass
            try:
                html_check = self.preview_text.toHtml() if hasattr(self.preview_text, 'toHtml') else ""
            except Exception:
                html_check = ""
            if not html_check:
                if self.preview_data:
                    try:
                        self.preview_text.setPlainText(self.preview_data)
                    except Exception:
                        pass
                elif self.current_temp_path and os.path.exists(self.current_temp_path):
                    try:
                        import docx
                        doc = docx.Document(self.current_temp_path)
                        lines = []
                        for p in doc.paragraphs[:50]:
                            t = p.text.strip()
                            if t:
                                lines.append(t)
                        try:
                            self.preview_text.setPlainText("\n".join(lines) or "（无可预览文本）")
                        except Exception:
                            pass
                    except Exception as e:
                        try:
                            self.preview_text.setPlainText(f"预览加载失败：{e}\n临时文件：{self.current_temp_path}")
                        except Exception:
                            pass
                else:
                    try:
                        self.preview_text.setPlainText("预览数据不可用。")
                    except Exception:
                        pass

    def on_save_clicked(self):
        if not self.current_original_path:
            InfoBar.warning(title="提示", content="没有可保存的原始文件信息", position=InfoBarPosition.TOP, parent=self)
            return
        d, f = os.path.split(self.current_original_path)
        name, ext = os.path.splitext(f)
        default_out = os.path.join(d, f"{name}_繁体{ext}")
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存转换后的文件到指定目录", default_out, f"*{ext}"
        )
        if save_path and self.current_original_path:
            self.save_requested.emit(self.current_original_path, save_path)
            InfoBar.success(title="已提交", content=f"正在保存到：{save_path}", position=InfoBarPosition.TOP, parent=self)

    def on_redo_clicked(self):
        if self.current_original_path:
            self.redo_requested.emit(self.current_original_path)
            InfoBar.info(title="重做转换", content="已重新提交转换任务，请返回转换列表查看进度", position=InfoBarPosition.TOP, parent=self)
        else:
            InfoBar.warning(title="提示", content="没有可重做的原始文件", position=InfoBarPosition.TOP, parent=self)

    def on_back_clicked(self):
        # 通知父窗口切换到转换列表
        parent_win = self.window()
        if hasattr(parent_win, 'switch_to_task_page'):
            parent_win.switch_to_task_page()

    def _on_suspect_list_leave(self, event):
        try:
            self.preview_text.hover_red_boxes = []
            self.preview_text.hover_word = ""
            self.preview_text.viewport().update()
        except Exception:
            pass
        try:
            ListWidget.leaveEvent(self.suspect_list, event)
        except Exception:
            pass

    def on_suspect_hover(self, item):
        if item is None:
            try:
                self.preview_text.hover_red_boxes = []
                self.preview_text.hover_word = ""
                self.preview_text.viewport().update()
            except Exception:
                pass
            return
        suspect = item.data(Qt.ItemDataRole.UserRole)
        if suspect is None:
            suspect = item.data(Qt.UserRole)
        if not isinstance(suspect, dict):
            try:
                self.preview_text.hover_red_boxes = []
                self.preview_text.hover_word = ""
                self.preview_text.viewport().update()
            except Exception:
                pass
            return
        hover_word = suspect.get("converted") or suspect.get("default") or ""
        if not hover_word:
            try:
                hover_word = engine.convert_text(suspect.get("simplified", ""))
            except Exception:
                hover_word = suspect.get("simplified", "")
        try:
            self.preview_text.hover_word = hover_word
            self.preview_text.hover_red_boxes = []
            text = self.preview_text.toPlainText()
            if not text or not hover_word:
                self.preview_text.viewport().update()
                return
            start = 0
            while True:
                idx = text.find(hover_word, start)
                if idx == -1:
                    break
                self.preview_text.hover_red_boxes.append(idx)
                start = idx + len(hover_word)
            if self.preview_text.hover_red_boxes:
                try:
                    cursor = self.preview_text.textCursor()
                    cursor.setPosition(self.preview_text.hover_red_boxes[0])
                    self.preview_text.setTextCursor(cursor)
                    self.preview_text.ensureCursorVisible()
                except Exception:
                    pass
            self.preview_text.viewport().update()
        except Exception:
            try:
                self.preview_text.hover_red_boxes = []
                self.preview_text.viewport().update()
            except Exception:
                pass


# 4. 主任务页面
class TaskPage(QWidget):
    request_preview = Signal(dict)
    switch_to_preview = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.threadpool = QThreadPool.globalInstance()
        self.setAcceptDrops(True)
        self.init_ui()
        self.temp_results = {}  # row -> result dict

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(30, 20, 30, 20)
        self.layout.setSpacing(15)

        # 标题栏
        self.title = TitleLabel("简繁文档转换", self)
        self.layout.addWidget(self.title)

        # 按钮控制栏
        btn_layout = QHBoxLayout()
        self.add_btn = PushButton(FIF.ADD, "添加文件", self)
        self.add_btn.clicked.connect(self.select_files)
        self.start_btn = PrimaryPushButton(FIF.PLAY, "开始全部转换", self)
        self.start_btn.clicked.connect(self.start_all_tasks)
        self.clear_btn = PushButton(FIF.DELETE, "清空列表", self)
        self.clear_btn.clicked.connect(self.clear_tasks)

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch(1)
        self.layout.addLayout(btn_layout)

        # 任务队列区表格（直接作为拖拽目的地，内置背景说明文字）
        self.table = QueueTableWidget(self)
        self.table.file_dropped.connect(self.add_files)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文件名", "格式", "进度", "状态", "操作"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(2, 180)
        self.table.horizontalHeader().setAcceptDrops(False)
        self.table.horizontalHeader().viewport().setAcceptDrops(False)
        self.table.viewport().setAcceptDrops(False)
        self.table.setAcceptDrops(False)
        self.layout.addWidget(self.table)

        self.tasks = []

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = []
        unsupported = []
        supported_exts = ('.docx', '.doc', '.odt', '.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')
        for u in urls:
            if not u.isLocalFile():
                continue
            local_path = u.toLocalFile()
            if local_path.endswith('\\') or local_path.endswith('/'):
                local_path = local_path.rstrip('\\/')
            lower_path = local_path.lower()
            if lower_path.endswith(supported_exts):
                files.append(local_path)
            else:
                unsupported.append(os.path.basename(local_path) or local_path)

        if files:
            event.acceptProposedAction()
            self.add_files(files)
        else:
            event.ignore()

        if unsupported:
            try:
                InfoBar.warning(
                    title="不支持的文件格式",
                    content="仅支持 DOCX/DOC/ODT/PDF/JPG/JPEG/PNG/BMP/TIFF/TIF 格式",
                    position=InfoBarPosition.TOP,
                    parent=self
                )
            except Exception:
                pass

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择待转换文档", "", "Documents (*.docx *.doc *.odt *.pdf *.jpg *.jpeg *.png *.bmp);;All Files (*.*)"
        )
        if files:
            self.add_files(files)

    def add_files(self, file_paths: list[str]):
        for path in file_paths:
            if any(t['path'] == path for t in self.tasks):
                continue
            row = self.table.rowCount()
            self.table.insertRow(row)

            # 文件名 & 格式
            ext = os.path.splitext(path)[1].lower()
            self.table.setItem(row, 0, QTableWidgetItem(os.path.basename(path)))
            self.table.setItem(row, 1, QTableWidgetItem(ext[1:].upper() if ext else ""))

            # 进度条
            pbar = ProgressBar(self)
            pbar.setValue(0)
            self.table.setCellWidget(row, 2, pbar)

            # 状态
            self.table.setItem(row, 3, QTableWidgetItem("待就绪"))
            
            # 操作按钮 (预览/保存)
            action_btn = PushButton("预览并保存", self)
            action_btn.setEnabled(False)
            action_btn.clicked.connect(lambda _, r=row: self.open_preview(r))
            self.table.setCellWidget(row, 4, action_btn)

            self.tasks.append({"path": path, "row": row, "done": False, "btn": action_btn, "result": None})
        self.table.viewport().update()

    def clear_tasks(self):
        self.table.setRowCount(0)
        self.tasks.clear()
        self.temp_results.clear()
        self.table.viewport().update()

    def start_all_tasks(self):
        if not self.tasks:
            InfoBar.warning(
                title="提示", content="请先添加需要转换的文件",
                position=InfoBarPosition.TOP, parent=self
            )
            return

        for task in self.tasks:
            if not task.get("done"):
                # 使用 process_to_temp 生成临时文件和预览数据
                worker = ConvertWorker(task["row"], task["path"], processor.process_to_temp, return_result=True)
                worker.signals.progress.connect(self.update_progress)
                worker.signals.status.connect(self.update_status)
                worker.signals.result.connect(self.on_convert_result)
                self.threadpool.start(worker)

    def update_progress(self, row: int, val: int):
        widget = self.table.cellWidget(row, 2)
        if isinstance(widget, ProgressBar):
            widget.setValue(val)

    def on_convert_result(self, row: int, result: dict):
        self.temp_results[row] = result
        for t in self.tasks:
            if t["row"] == row:
                t["result"] = result
                break

    def update_status(self, row: int, text: str, err: str):
        item = self.table.item(row, 3)
        if item:
            item.setText(text)
        if text == "已完成":
            for t in self.tasks:
                if t["row"] == row and not t.get("done"):
                    t["done"] = True
                    t["btn"].setEnabled(True)

    def open_preview(self, row: int):
        result = self.temp_results.get(row)
        if not result:
            # 尝试从任务中获取
            for t in self.tasks:
                if t["row"] == row:
                    result = t.get("result")
                    break
        if result:
            self.request_preview.emit(result)
            self.switch_to_preview.emit()
        else:
            InfoBar.warning(title="提示", content="预览数据尚未就绪，请稍后再试", position=InfoBarPosition.TOP, parent=self)

    def redo_conversion(self, original_path: str):
        # 查找对应任务行并重新开始转换
        for t in self.tasks:
            if t["path"] == original_path and t.get("done"):
                t["done"] = False
                if t["btn"]:
                    t["btn"].setEnabled(False)
                item = self.table.item(t["row"], 3)
                if item:
                    item.setText("待重做...")
                worker = ConvertWorker(t["row"], t["path"], processor.process_to_temp, return_result=True)
                worker.signals.progress.connect(self.update_progress)
                worker.signals.status.connect(self.update_status)
                worker.signals.result.connect(self.on_convert_result)
                self.threadpool.start(worker)
                InfoBar.info(title="重做转换", content=f"已重新提交：{os.path.basename(original_path)}", position=InfoBarPosition.TOP, parent=self)
                return
        InfoBar.warning(title="提示", content="未找到对应任务或任务未完成", position=InfoBarPosition.TOP, parent=self)

    def save_to_target(self, original_path: str, target_path: str):
        ext = os.path.splitext(original_path)[1].lower()
        # 查找任务行
        task_row = None
        for t in self.tasks:
            if t["path"] == original_path:
                task_row = t["row"]
                break
        if task_row is None:
            InfoBar.error(title="错误", content="无法确定任务行", position=InfoBarPosition.TOP, parent=self)
            return

        try:
            if ext == '.docx':
                worker = ConvertWorker(task_row, original_path, lambda path, cb: processor._process_docx(path, target_path, cb), return_result=False)
            elif ext == '.pdf':
                worker = ConvertWorker(task_row, original_path, lambda path, cb: processor._process_pdf(path, target_path, cb), return_result=False)
            elif ext in ['.doc', '.odt']:
                worker = ConvertWorker(task_row, original_path, lambda path, cb: processor._process_via_libreoffice(path, target_path, cb), return_result=False)
            elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
                worker = ConvertWorker(task_row, original_path, lambda path, cb: processor._process_jpg(path, target_path, cb), return_result=False)
            else:
                worker = ConvertWorker(task_row, original_path, lambda path, cb: processor.process(path, cb), return_result=False)
            worker.signals.progress.connect(self.update_progress)
            worker.signals.status.connect(self.update_status)
            self.threadpool.start(worker)
            InfoBar.success(title="保存中", content=f"正在转换并保存到指定目录：{target_path}", position=InfoBarPosition.TOP, parent=self)
        except Exception as e:
            InfoBar.error(title="保存失败", content=str(e), position=InfoBarPosition.TOP, parent=self)


# 桩函数，供 UI 流程联调使用
def dummy_convert_pipeline(file_path: str, progress_cb):
    import time
    for i in range(1, 101, 20):
        time.sleep(0.1)
        progress_cb(i)


# 5. 主窗口封装
class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("简繁文档专业转换工具")
        self.resize(980, 680)
        self.setAcceptDrops(True)

        # 页面注册
        self.task_interface = TaskPage(self)
        self.task_interface.setObjectName("taskInterface")
        self.addSubInterface(self.task_interface, FIF.DOCUMENT, "转换列表")

        self.preview_interface = PreviewPage(self)
        self.preview_interface.setObjectName("previewInterface")
        self.addSubInterface(self.preview_interface, FIF.VIEW, "预览页")

        # 事件绑定：任务页 <-> 预览页
        self.task_interface.request_preview.connect(self.preview_interface.load_preview)
        self.task_interface.switch_to_preview.connect(self.switch_to_preview_page)
        self.preview_interface.redo_requested.connect(self.task_interface.redo_conversion)
        self.preview_interface.save_requested.connect(self.task_interface.save_to_target)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = []
        unsupported = []
        supported_exts = ('.docx', '.doc', '.odt', '.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')
        for u in urls:
            if not u.isLocalFile():
                continue
            local_path = u.toLocalFile()
            if local_path.endswith('\\') or local_path.endswith('/'):
                local_path = local_path.rstrip('\\/')
            lower_path = local_path.lower()
            if lower_path.endswith(supported_exts):
                files.append(local_path)
            else:
                unsupported.append(os.path.basename(local_path) or local_path)

        if files:
            event.acceptProposedAction()
            try:
                self.switchTo(self.task_interface)
            except Exception:
                pass
            try:
                self.task_interface.add_files(files)
            except Exception:
                pass
        else:
            event.ignore()

        if unsupported:
            try:
                InfoBar.warning(
                    title="不支持的文件格式",
                    content="仅支持 DOCX/DOC/ODT/PDF/JPG/JPEG/PNG/BMP/TIFF/TIF 格式",
                    position=InfoBarPosition.TOP,
                    parent=self
                )
            except Exception:
                pass

    def switch_to_preview_page(self):
        try:
            self.switchTo(self.preview_interface)
        except Exception:
            pass

    def switch_to_task_page(self):
        try:
            self.switchTo(self.task_interface)
        except Exception:
            pass


if __name__ == "__main__":
    app = QApplication(sys.argv)

    def _qt_message_handler(mode, context, message):
        try:
            msg = message or ""
            if "CreateFontFaceFromHDC() failed" in msg or "Fixedsys" in msg:
                return
        except Exception:
            pass
        try:
            sys.stderr.write(f"{message}\n")
        except Exception:
            pass

    try:
        qInstallMessageHandler(_qt_message_handler)
    except Exception:
        pass

    try:
        QFont.insertSubstitution("Fixedsys", "Microsoft YaHei UI")
        QFont.insertSubstitution("System", "Microsoft YaHei UI")
    except Exception:
        pass

    try:
        app.setFont(QFont("Microsoft YaHei UI", 10))
    except Exception:
        try:
            app.setFont(QFont("Microsoft YaHei", 10))
        except Exception:
            pass

    w = MainWindow()
    w.show()
    sys.exit(app.exec())