# main.py
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Inches

try:
    import pygments
    from pygments.formatters import HtmlFormatter
    from pygments.lexer import RegexLexer
    from pygments.token import Text
    print("DEBUG: Pygments imported successfully")      
except ImportError as e:
    print(f"ERROR: Pygments import failed: {e}")
    sys.exit(1)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QPushButton, QLabel, QMessageBox, QFrame, QSizePolicy, QFileDialog, QComboBox,
    QStyleOption, QStyle
)
from PyQt6.QtGui import QFont, QIcon, QDesktopServices, QPixmap, QPainter
from PyQt6.QtCore import QMimeData, Qt, QUrl, QTimer, QSize

# 从我们创建的模块中导入核心处理类
from reference_processor import ReferenceProcessor, DOCX_AVAILABLE
from styles import build_main_window_qss, DIALOG_QSS


class MarqueeLabel(QLabel):
    """单行滚动标签，文本超出时自动滚动显示。"""

    def __init__(self, text="", parent=None, interval=35, step=2, gap="    "):
        super().__init__(text, parent)
        self._full_text = text or ""
        self._interval = interval
        self._step = step
        self._gap = gap
        self._offset = 0
        self._need_scroll = False
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._update_offset)
        self._padding = 12
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.setContentsMargins(self._padding, 0, self._padding, 0)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

    def setText(self, text):
        self._full_text = text or ""
        super().setText(self._full_text)
        self._offset = 0
        self._evaluate_scroll()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._evaluate_scroll()

    def paintEvent(self, event):
        if not self._need_scroll:
            super().paintEvent(event)
            return

        option = QStyleOption()
        option.initFrom(self)

        painter = QPainter(self)
        self.style().drawPrimitive(QStyle.PrimitiveElement.PE_Widget, option, painter, self)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        painter.setPen(self.palette().windowText().color())

        rect = self.rect().adjusted(self._padding, 0, -self._padding, 0)
        painter.setClipRect(rect)
        text = self._full_text + self._gap
        text_width = self.fontMetrics().horizontalAdvance(text)
        baseline = int(
            rect.center().y()
            + (self.fontMetrics().ascent() - self.fontMetrics().descent()) / 2
        )

        x = rect.left() - self._offset
        while x < rect.right():
            painter.drawText(x, baseline, text)
            x += text_width

    def _evaluate_scroll(self):
        available = max(1, self.width() - 2 * self._padding)
        text_width = self.fontMetrics().horizontalAdvance(self._full_text)
        if text_width > available:
            if not self._timer.isActive():
                self._timer.start(self._interval)
            self._need_scroll = True
        else:
            self._need_scroll = False
            if self._timer.isActive():
                self._timer.stop()
            self._offset = 0
        self.update()

    def _update_offset(self):
        text = self._full_text + self._gap
        width = self.fontMetrics().horizontalAdvance(text)
        if width == 0:
            return
        self._offset = (self._offset + self._step) % width
        self.update()


class ReferenceFormatterApp(QMainWindow):
    """
    文献引用导出工具的主窗口类。
    使用 PyQt6 构建GUI，并调用 ReferenceProcessor 处理核心逻辑。
    """

    # ⚠️ 新增：图标和链接常量
    ICON_PATH = "D:/python/pythonProject10/app_icon.png"
    GITHUB_URL = "https://github.com/rhj-flash/Reference-Formatter-Tool"
    GITHUB_ICON_PATH = "D:/python/pythonProject10/github_icon.ico"

    def __init__(self):
        """
        初始化应用，设置处理器和固定格式。
        """
        super().__init__()
        self.processor = ReferenceProcessor()
        # 存储 Word HTML 结果，供复制使用
        self.html_output_for_clipboard = ""
        # 固定使用的格式名称
        self.fixed_format_name = "普通数字"

        # 设置主窗口图标
        self.setWindowIcon(QIcon(self.ICON_PATH))

        self.initUI()
        self._apply_global_style()  # 应用全局样式

    def _open_github_link(self):
        """打开 GitHub 仓库链接。"""
        QDesktopServices.openUrl(QUrl(self.GITHUB_URL))
        print(f"DEBUG: Opening GitHub link: {self.GITHUB_URL}")

    def _apply_global_style(self):
        """
        应用一套大胆、优雅、淡雅的全局 QSS 样式 (Glassmorphism 玻璃磨砂风格)。
        """
        # 设置全局字体，确保中文显示 (Windows 推荐使用 Microsoft YaHei UI)
        font = QFont("Microsoft YaHei UI", 10)
        self.setFont(font)
        self.setStyleSheet(build_main_window_qss())

    def _create_hero_card(self):
        """创建顶部的介绍卡片，展示图标、标题和描述。"""
        hero_card = QFrame()
        hero_card.setObjectName("HeroCard")
        hero_layout = QHBoxLayout(hero_card)
        hero_layout.setContentsMargins(0, 0, 0, 0)
        hero_layout.setSpacing(4)

        icon_label = QLabel()
        pixmap = QPixmap(self.ICON_PATH)
        if not pixmap.isNull():
            icon_label.setPixmap(pixmap.scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation))
        hero_layout.addWidget(icon_label, alignment=Qt.AlignmentFlag.AlignTop)

        text_layout = QVBoxLayout()
        text_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label = QLabel("📚 文献列表格式化工具")
        title_label.setObjectName("TitleLabel")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text_layout.addWidget(title_label)

        subtitle = QLabel("4 步完成从杂乱引用到正式 Word 交付的全流程，让排版、字体、缩进一次搞定。")
        subtitle.setObjectName("HeroSubtitle")
        subtitle.setWordWrap(False)
        subtitle.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        subtitle.setMaximumHeight(subtitle.fontMetrics().height() + 6)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text_layout.addWidget(subtitle)

        hero_layout.addLayout(text_layout, stretch=1)
        return hero_card

    def initUI(self):
        """
        设置主窗口的布局和组件，添加字体设置面板。
        """
        self.setWindowTitle('文献引用格式化工具')

        # 设置更舒适的窗口尺寸比例
        window_width = 1040
        window_height = 500
        self.resize(window_width, window_height)

        # 窗口居中逻辑：使用整块屏幕区域的宽高来计算，使窗口在视觉上居中
        screen = QApplication.primaryScreen()
        if screen is not None:
            screen_geometry = screen.geometry()
            desktop_width = screen_geometry.width()
            desktop_height = screen_geometry.height()

            # 使用当前窗口的实际宽高进行居中（包含窗口边框在内）
            frame_geo = self.frameGeometry()
            window_w = frame_geo.width() or window_width
            window_h = frame_geo.height() or window_height

            x = screen_geometry.x() + max(0, (desktop_width - window_w) // 2)
            # 在几何中心的基础上再向上偏移一段距离，使窗口看起来更接近屏幕正中而不是偏底部
            base_y = screen_geometry.y() + max(0, (desktop_height - window_h) // 2)
            y = max(screen_geometry.y(), base_y - 80)

            # 调试输出：桌面尺寸、窗口尺寸和最终位置
            print(f"DEBUG: Desktop size = {desktop_width}x{desktop_height}")
            print(f"DEBUG: Window size  = {window_w}x{window_h}")
            print(f"DEBUG: Move window to x={x}, y={y}")

            self.move(x, y)

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 调整边距和间距，使布局更紧凑
        main_layout.setContentsMargins(18, 10, 18, 12)
        main_layout.setSpacing(10)

        # 顶部介绍卡片
        main_layout.addWidget(self._create_hero_card())

        # --- 字体设置面板 - 调整布局 ---
        font_settings_layout = QHBoxLayout()
        font_settings_layout.setSpacing(8)

        # 英文字体设置
        english_font_layout = QHBoxLayout()
        english_font_layout.setContentsMargins(0, 0, 0, 0)
        english_font_layout.setSpacing(6)
        english_font_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        # 先放置 GitHub 按钮在最左侧
        github_button = QPushButton()
        github_button.setIcon(QIcon(self.GITHUB_ICON_PATH))
        github_button.setIconSize(QSize(20, 20))
        github_button.setCursor(Qt.CursorShape.PointingHandCursor)
        # 加宽按钮，让图标左右留白多一些
        github_button.setFixedSize(28, 28)

        github_button.setToolTip("访问 GitHub 仓库")
        github_button.setFlat(True)
        # 通过不对称 padding 让图标在按钮内部略微左移
        github_button.setStyleSheet(
            "QPushButton { border: none; background-color: transparent; padding-left: 4px; padding-right: 12px; }"
            "QPushButton:hover { background-color: rgba(0,0,0,0.05); border-radius: 14px; }"
            "QPushButton:pressed { background-color: rgba(0,0,0,0.08); }"
        )

        github_button.clicked.connect(self._open_github_link)
        english_font_layout.addWidget(github_button, alignment=Qt.AlignmentFlag.AlignVCenter)

        # GitHub 按钮右侧预留一段固定空白，用于视觉分隔
        english_font_layout.addSpacing(20)

        # 中间放“英文字体”文字和下拉框
        english_font_label = QLabel("英文字体：")
        english_font_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        english_font_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        english_font_layout.addWidget(english_font_label)

        self.english_font_combo = QComboBox()
        self.english_font_combo.addItems([
            "Times New Roman",
            "Arial",
            "Calibri",
            "Cambria",
            "Georgia",
            "Verdana"
        ])
        self.english_font_combo.setCurrentText("Times New Roman")
        self.english_font_combo.setFixedHeight(34)
        english_font_layout.addWidget(self.english_font_combo)

        # 英文字号
        english_size_layout = QHBoxLayout()
        english_size_label = QLabel("英文字号：")
        english_size_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        english_size_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        english_size_layout.addWidget(english_size_label)

        self.english_size_combo = QComboBox()
        self.english_size_combo.addItems(["10", "10.5", "11", "12", "14", "16", "18"])
        self.english_size_combo.setCurrentText("12")
        self.english_size_combo.setFixedHeight(34)
        english_size_layout.addWidget(self.english_size_combo)

        # 中文字体设置
        chinese_font_layout = QHBoxLayout()
        chinese_font_label = QLabel("中文字体：")
        chinese_font_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        chinese_font_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        chinese_font_layout.addWidget(chinese_font_label)

        self.chinese_font_combo = QComboBox()
        self.chinese_font_combo.addItems([
            "宋体",
            "黑体",
            "微软雅黑",
            "楷体",
            "仿宋",
            "华文宋体"
        ])
        self.chinese_font_combo.setCurrentText("宋体")
        self.chinese_font_combo.setFixedHeight(34)
        chinese_font_layout.addWidget(self.chinese_font_combo)

        # 中文字号
        chinese_size_layout = QHBoxLayout()
        chinese_size_label = QLabel("中文字号：")
        chinese_size_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        chinese_size_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        chinese_size_layout.addWidget(chinese_size_label)

        self.chinese_size_combo = QComboBox()
        self.chinese_size_combo.addItems(["10", "10.5", "11", "12", "14", "16", "18"])
        self.chinese_size_combo.setCurrentText("12")
        self.chinese_size_combo.setFixedHeight(34)
        chinese_size_layout.addWidget(self.chinese_size_combo)

        # 将所有字体设置添加到水平布局
        font_settings_layout.addLayout(english_font_layout)
        font_settings_layout.addLayout(english_size_layout)
        font_settings_layout.addLayout(chinese_font_layout)
        font_settings_layout.addLayout(chinese_size_layout)
        font_settings_layout.addStretch(1)

        main_layout.addLayout(font_settings_layout)

        # 主内容布局：左侧控制面板 + 右侧输入/输出区域
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)  # 减少间距

        # --- 左侧控制面板 - 调整宽度 ---
        control_panel = QFrame()
        control_panel.setObjectName("ControlPanel")
        control_panel.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        control_panel.setFixedWidth(260)
        control_panel_layout = QVBoxLayout(control_panel)
        control_panel_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        control_panel_layout.setSpacing(8)

        # 1. 操作步骤标题
        control_panel_layout.addWidget(QLabel("➡️ **格式化流程 (4 步)**"))
        control_panel_layout.addWidget(self._create_separator())

        # 2. 步骤按钮 - 调整按钮高度
        self.preview_button = QPushButton("1. 检查文献分割")
        self.preview_button.setObjectName("PreviewButton")
        self.preview_button.clicked.connect(self.split_preview)
        self.preview_button.setMinimumHeight(42)
        control_panel_layout.addWidget(self.preview_button)

        self.process_button = QPushButton("2. 统一格式并清洗")
        self.process_button.setObjectName("ProcessButton")
        self.process_button.clicked.connect(self.process_references)
        self.process_button.setMinimumHeight(42)
        control_panel_layout.addWidget(self.process_button)

        self.copy_button = QPushButton("3. 复制 Word 专用格式")
        self.copy_button.setObjectName("CopyButton")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        self.copy_button.setMinimumHeight(42)
        control_panel_layout.addWidget(self.copy_button)

        # 3. 生成Word文件按钮
        self.export_button = QPushButton("4. 生成 Word 文件")
        self.export_button.setObjectName("ExportButton")
        self.export_button.clicked.connect(self.export_to_word_file)
        self.export_button.setMinimumHeight(42)
        control_panel_layout.addWidget(self.export_button)

        control_panel_layout.addSpacing(10)

        # 4. 当前字体设置显示
        font_info_label = QLabel(
            f"**当前字体设置:**\n"
            f"• 英文: {self.english_font_combo.currentText()} {self.english_size_combo.currentText()}pt\n"
            f"• 中文: {self.chinese_font_combo.currentText()} {self.chinese_size_combo.currentText()}pt"
        )
        font_info_label.setObjectName("FormatInfo")
        font_info_label.setWordWrap(True)
        font_info_label.setFixedHeight(58)
        self.font_info_label = font_info_label
        control_panel_layout.addWidget(self.font_info_label)

        control_panel_layout.addStretch(1)

        # 5. 提示
        self.status_label = MarqueeLabel("💡 状态: 等待用户输入原始文献。")
        self.status_label.setObjectName("StatusLabel")
        self.status_label.setFixedHeight(40)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        control_panel_layout.addWidget(self.status_label)

        content_layout.addWidget(control_panel)

        # --- 右侧输入/输出区域 - 调整比例 ---
        io_area_layout = QVBoxLayout()
        io_area_layout.setSpacing(10)

        # 输入区域
        input_label = QLabel("📝 **原始文献输入区**")
        io_area_layout.addWidget(input_label)
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("请将文献列表粘贴到此处。程序将自动处理乱码、多行和旧编号...")
        self.input_text.setFont(QFont("Courier New", 10))
        self.input_text.setMinimumHeight(90)
        self.input_text.setFrameShape(QFrame.Shape.NoFrame)
        io_area_layout.addWidget(self.input_text, 1)  # 权重为2

        # 输出/预览区域
        output_label = QLabel("👁️ **格式化预览区**")
        io_area_layout.addWidget(output_label)
        self.output_text = QTextEdit()
        self.output_text.setObjectName("output_text")
        self.output_text.setReadOnly(True)
        self.output_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.output_text.setMinimumHeight(120)
        self.output_text.setFrameShape(QFrame.Shape.NoFrame)
        io_area_layout.addWidget(self.output_text, 1)  # 权重为3，比输入区域稍大

        content_layout.addLayout(io_area_layout, 1)

        main_layout.addLayout(content_layout)

        # 连接字体变化的信号
        self.english_font_combo.currentTextChanged.connect(self.update_font_info)
        self.english_size_combo.currentTextChanged.connect(self.update_font_info)
        self.chinese_font_combo.currentTextChanged.connect(self.update_font_info)
        self.chinese_size_combo.currentTextChanged.connect(self.update_font_info)

    def update_font_info(self):
        """更新字体设置显示"""
        if hasattr(self, "font_info_label") and self.font_info_label:
            self.font_info_label.setText(
                f"**当前字体设置:**\n"
                f"• 英文: {self.english_font_combo.currentText()} {self.english_size_combo.currentText()}pt\n"
                f"• 中文: {self.chinese_font_combo.currentText()} {self.chinese_size_combo.currentText()}pt"
            )

    def _create_separator(self):
        """创建一个视觉分隔线"""
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Plain)
        # 使用淡雅的分隔线颜色
        line.setStyleSheet("QFrame { background-color: #cccccc; height: 1px; border: none; margin: 10px 0; }")
        return line

    def split_preview(self):
        """
        第一步：调用处理器生成分割预览 HTML。
        """
        raw_text = self.input_text.toPlainText()
        if not raw_text.strip():
            self.status_label.setText("⚠️ 输入为空，请粘贴文献文本。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("输入为空")
            msg.setText("请输入文献文本后进行分割预览。")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()
            return

        try:
            selected_format = self.fixed_format_name

            # 调用处理器获取格式化后的分割预览 HTML
            preview_html = self.processor.get_formatted_split_preview(raw_text, selected_format)

            # 在预览区域显示 HTML 内容
            self.output_text.setHtml(preview_html)

            # 每次分割预览后，清空已有的 HTML 格式化结果，防止用户跳过格式化直接复制
            self.html_output_for_clipboard = ""

            # 更新提示信息
            self.status_label.setText("✅ 步骤 1 完成：请检查右侧的分组结果。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setWindowTitle("检查文献分割 (1/3)")
            msg.setText("✅ 文献分组预览已生成！\n\n"
                        "• 请检查右侧预览区，确认每篇文献是否被正确地分割到独立的彩色区块中。\n\n"
                        "➡️ 下一步：确认无误后，点击 '统一格式并清洗' 按钮。")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()


        except Exception as e:
            self.status_label.setText(f"❌ 步骤 1 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("处理错误")
            msg.setText(f"文献分割时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

    def process_references(self):
        """
        第二步：调用处理器进行完整的格式化，并存储 Word HTML 结果。
        """
        raw_text = self.input_text.toPlainText()
        if not raw_text.strip():
            self.status_label.setText("⚠️ 输入为空，请粘贴文献文本。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("输入为空")
            msg.setText("请输入文献文本后进行格式化。")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()
            return

        try:
            selected_format = self.fixed_format_name

            # 调用核心处理函数
            word_html_output, plain_text_output, was_stripped = self.processor.process_text(raw_text, selected_format)

            # 存储 HTML 结果供复制使用
            self.html_output_for_clipboard = word_html_output

            # 在预览区域显示纯文本格式化结果作为最终确认
            self.output_text.setPlainText(plain_text_output)

            # 弹出提示
            stripped_message = "自动剥离了旧编号" if was_stripped else "未检测到旧编号"
            self.status_label.setText(f"🎉 步骤 2 完成：格式统一，Word结果已就绪! ({stripped_message})")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setWindowTitle("统一格式并清洗 (2/3)")
            msg.setText(f"🎉 文献列表已格式化！\n\n"
                        f"• 当前预览区显示的是最终纯文本结果。\n"
                        f"• {stripped_message}，并应用了中英文分字体等样式。\n"
                        f"• Word专用格式已准备好复制。\n\n"
                        f"➡️ 下一步：点击 '复制 Word 专用格式' 按钮。")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

        except Exception as e:
            self.status_label.setText(f"❌ 步骤 2 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("处理错误")
            msg.setText(f"格式化时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

    def copy_to_clipboard(self):
        """
        第三步：复制格式化结果到剪贴板。
        """
        if not self.html_output_for_clipboard:
            self.status_label.setText("⚠️ 请按顺序先进行 '检查' 和 '统一格式' 操作。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("操作顺序提示")
            msg.setText("请先完成 '检查文献分割' 和 '统一格式并清洗'！")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()
            return

        try:
            # 准备数据对象
            mime_data = QMimeData()

            # 1. 设置纯文本
            mime_data.setText(self.output_text.toPlainText())

            # 2. 设置 HTML 格式 (Word 识别的关键)
            html_data = self.html_output_for_clipboard
            mime_data.setHtml(html_data)

            # 复制到剪贴板
            QApplication.clipboard().setMimeData(mime_data)

            self.status_label.setText("✨ 步骤 3 成功！请在 Word 中粘贴 (Ctrl+V)。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setWindowTitle("复制 Word 专用格式 (3/3)")
            msg.setText("✨ 格式化后的文献列表已复制到剪贴板。\n\n"
                        "请将光标置于 Word 文档中，使用 **Ctrl+V** 进行粘贴。\n\n"
                        "提示：Word会自动应用编号和字体，无需手动调整。粘贴后如果出现 Word 的列表提示，忽略即可。")
            msg.setStyleSheet(DIALOG_QSS)
            # 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

        except Exception as e:
            self.status_label.setText(f"❌ 步骤 3 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            # 修改：新增设置弹窗图标 (修复了原始代码中此处缺少 msg.exec() 的问题)
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.setWindowTitle("处理错误")  # 补充了标题，使弹窗完整
            msg.setText(f"复制到剪贴板时出现错误：\n{str(e)}")  # 补充了文本
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()

    def export_to_word_file(self):
        """
        第四步：将格式化结果导出为Word文件，使用自定义字体设置。
        """
        if not self.html_output_for_clipboard:
            self.status_label.setText("⚠️ 请先完成格式化操作。")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("操作顺序提示")
            msg.setText("请先完成 '检查文献分割' 和 '统一格式并清洗'！")
            msg.setStyleSheet(DIALOG_QSS)
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()
            return

        try:
            # 弹出文件保存对话框
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存 Word 文件",
                "参考文献.docx",
                "Word Documents (*.docx);;All Files (*)"
            )

            if not file_path:
                return

            # 确保文件扩展名是 .docx
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'

            # 获取用户选择的字体设置
            english_font = self.english_font_combo.currentText()
            english_size = float(self.english_size_combo.currentText())
            chinese_font = self.chinese_font_combo.currentText()
            chinese_size = float(self.chinese_size_combo.currentText())

            # 创建自定义格式配置
            custom_format = {
                "language": "chinese",
                "line_spacing": 1.5,
                "font_size": english_size,  # 使用英文字号作为基准
                "chinese_font": chinese_font,
                "english_font": english_font,
                "title_alignment": "center",
                "title_font_size": 16,
                "title_margin_bottom": 20,
                "item_spacing": 6,
                "hanging_indent": 2,
                "requirements": [
                    f"中文文献使用{chinese_font}",
                    f"英文文献使用{english_font}",
                    f"英文字号: {english_size}pt",
                    f"中文字号: {chinese_size}pt",
                    "1.5倍行距",
                    "标题居中，16号字",
                    "悬挂缩进2字符",
                    "文献间间距6磅"
                ]
            }

            # 使用自定义字体设置导出Word文件
            success = self.processor.export_to_word_file_with_custom_font(
                self.html_output_for_clipboard,
                file_path,
                custom_format,
                "普通数字",  # 固定使用普通数字格式
                english_font,
                english_size,
                chinese_font,
                chinese_size
            )

            if success:
                self.word_file_path = file_path
                file_name = os.path.basename(file_path)
                self.status_label.setText(f"✅ Word文件已生成: {file_name}")

                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Icon.Information)
                msg.setWindowTitle("生成 Word 文件 (4/4)")
                msg.setText(f"✨ Word文件已成功生成！\n\n"
                            f"文件位置: {file_path}\n"
                            f"编号格式: 普通数字自动排序\n"
                            f"英文字体: {english_font} {english_size}pt\n"
                            f"中文字体: {chinese_font} {chinese_size}pt")
                msg.setStyleSheet(DIALOG_QSS)
                msg.setWindowIcon(QIcon(self.ICON_PATH))
                msg.exec()
            else:
                raise Exception("Word文件生成失败")

        except PermissionError as e:
            # 专门处理权限错误
            self.status_label.setText(f"❌ 文件保存失败: 权限被拒绝")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("文件保存错误")
            msg.setText(
                f"无法保存文件：\n{str(e)}\n\n请确保：\n1. 文件没有被其他程序打开\n2. 您有该位置的写入权限\n3. 文件路径正确")
            msg.setStyleSheet(DIALOG_QSS)
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()
        except Exception as e:
            self.status_label.setText(f"❌ 导出失败: {str(e)[:50]}...")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("导出错误")
            msg.setText(f"生成Word文件时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 新增：设置应用程序的图标，确保任务栏和文件管理器中显示正确的图标
    # 路径使用 ReferenceFormatterApp 中定义的常量，确保一致性
    app.setWindowIcon(QIcon(ReferenceFormatterApp.ICON_PATH))

    # 为 QApplication 设置中文字体，确保全局中文显示正常
    font = QFont("Microsoft YaHei UI")
    app.setFont(font)

    ex = ReferenceFormatterApp()
    ex.show()
    sys.exit(app.exec())