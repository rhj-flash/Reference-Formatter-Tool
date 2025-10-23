# main.py

import sys
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
    QTextEdit, QPushButton, QLabel, QMessageBox, QFrame, QSizePolicy
)
from PyQt6.QtGui import QFont, QColor, QIcon, QDesktopServices
from PyQt6.QtCore import QMimeData, Qt, QUrl

# 从我们创建的模块中导入核心处理类
from reference_processor import ReferenceProcessor










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
        # 修复/简化：固定使用的格式名称
        self.fixed_format_name = "普通数字 (Word默认)"

        # ⚠️ 恢复默认标题栏：不需要设置 FramelessWindowHint

        # ⚠️ 新增：设置主窗口图标 (对于系统标题栏，这一步就足够了)
        self.setWindowIcon(QIcon(self.ICON_PATH))

        self.initUI()
        self._apply_global_style()  # 应用全局样式

        # ⚠️ 新增：创建并添加 GitHub 图标操作
        self._create_github_action()
        print("DEBUG: GitHub Action added to menu bar.")

    # ⚠️ 新增：创建 GitHub 链接的 Action
    def _create_github_action(self):
        """
        在菜单栏(或工具栏)添加一个可点击的 GitHub 图标。
        使用 QToolBar 实现一个不含文字，仅含图标的按钮。
        """
        # 1. 创建工具栏
        toolbar = self.addToolBar("External Links")
        toolbar.setObjectName("ExternalLinksToolbar")
        toolbar.setMovable(False)  # 禁止用户拖动工具栏
        toolbar.setStyleSheet("QToolBar { padding: 0px; margin: 0px; border: none; }")

        # 2. 创建 Action (可点击的图标)
        github_action = toolbar.addAction(QIcon(self.GITHUB_ICON_PATH), "GitHub")
        github_action.setToolTip("访问 GitHub 仓库")

        # 3. 连接到槽函数
        github_action.triggered.connect(self._open_github_link)

    # ⚠️ 新增：打开 GitHub 链接的槽函数
    def _open_github_link(self):
        """打开 GitHub 仓库链接。"""
        QDesktopServices.openUrl(QUrl(self.GITHUB_URL))
        print(f"DEBUG: Opening GitHub link: {self.GITHUB_URL}")

    def _apply_global_style(self):
        """
        应用一套大胆、优雅、淡雅的全局 QSS 样式 (Glassmorphism 玻璃磨砂风格)。
        """
        # 定义核心莫兰迪颜色
        # ⚠️ 主背景色使用极浅色，保持淡雅统一
        COLOR_BACKGROUND = "white"  # 主背景/左侧面板背景：极浅蓝灰 (模拟玻璃后面的背景色)
        COLOR_CARD_BG = "white"  # 文本编辑区背景：纯白
        COLOR_TEXT_DARK = "#33415c"  # 深色文本：深海军蓝
        COLOR_TEXT_LIGHT = "#7f8c8d"  # 浅色文本：中灰
        COLOR_BORDER = "#d8e1e8"  # 边框/分隔线：浅灰

        # Glassmorphism 按钮核心颜色 (使用 rgba，并设置低饱和度颜色)
        COLOR_PREVIEW_RGBA = "rgba(162, 185, 188, 0.6)"  # 预览 (灰蓝, 60% 透明度)
        COLOR_PROCESS_RGBA = "rgba(197, 197, 138, 0.6)"  # 格式化 (橄榄绿, 60% 透明度)
        COLOR_COPY_RGBA = "rgba(216, 167, 134, 0.6)"  # 复制 (焦糖橙, 60% 透明度)

        # 悬停时颜色 (使用更深的颜色以保持对比和反馈)
        COLOR_PREVIEW_HOVER = "rgba(162, 185, 188, 0.8)"
        COLOR_PROCESS_HOVER = "rgba(197, 197, 138, 0.8)"
        COLOR_COPY_HOVER = "rgba(216, 167, 134, 0.8)"

        qss = f"""
            QMainWindow {{
                background-color: {COLOR_BACKGROUND};
            }}
            QWidget#centralWidget {{
                background-color: {COLOR_BACKGROUND};
            }}
            QLabel {{
                color: {COLOR_TEXT_DARK};
                font-size: 10pt;
            }}
            /* 标题样式 */
            QLabel#TitleLabel {{
                font-size: 18pt;
                font-weight: bold;
                color: {COLOR_TEXT_DARK}; 
                padding: 20px 0 15px 0;
            }}
            /* 提示/状态标签样式 */
            QLabel#StatusLabel {{
                color: {COLOR_TEXT_DARK};
                font-weight: 500;
                font-size: 10pt;
                padding: 12px;
                border: 1px solid {COLOR_BORDER};
                border-radius: 8px;
                background-color: {COLOR_CARD_BG};
                margin-top: 10px;
            }}
            /* 输入/输出区域样式 - 卡片设计 */
            QTextEdit {{
                border: 1px solid {COLOR_BORDER};
                border-radius: 10px;
                padding: 15px;
                background-color: {COLOR_CARD_BG};
                font-size: 10pt;
                line-height: 1.5;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.04);
            }}
            QTextEdit:focus {{
                border: 1px solid rgba(162, 185, 188, 1.0);
            }}
            QTextEdit#output_text {{
                background-color: #f7f9fc;
            }}
            /* 控制面板框架样式 - 模拟与主窗口的统一 */
            QFrame#ControlPanel {{
                background-color: {COLOR_BACKGROUND};
                border: none;
                border-radius: 12px;
                padding: 20px;
                /* 保持阴影，让卡片浮起 */
                box-shadow: 0 6px 15px rgba(0, 0, 0, 0.08); 
            }}
            /* 按钮通用样式 - 玻璃磨砂质感 */
            QPushButton {{
                color: {COLOR_TEXT_DARK}; /* 按钮文本使用深色以增强可读性 */
                border: 1px solid rgba(255, 255, 255, 0.3); /* 半透明白色边框 */
                padding: 14px 15px;
                border-radius: 12px; /* 更圆润的边角 */
                font-weight: 600;
                font-size: 10pt;
                min-height: 35px;
                text-align: left;
                margin-bottom: 12px; /* 增加按钮间距 */
                /* 模拟玻璃高光和立体感的阴影 */
                box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1),
                            0 0 0 1px rgba(255, 255, 255, 0.5) inset; 
            }}
            QPushButton:hover {{
                /* 悬停时去除高光，使用更深的背景色 */
                box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15); 
            }}
            QPushButton:pressed {{
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2); /* 按下时的阴影 */
                padding-top: 15px; /* 模拟按下效果 */
            }}

            /* 步骤 1: 预览按钮 */
            QPushButton#PreviewButton {{
                background-color: {COLOR_PREVIEW_RGBA};
            }}
            QPushButton#PreviewButton:hover {{
                background-color: {COLOR_PREVIEW_HOVER};
            }}
            /* 步骤 2: 格式化按钮 */
            QPushButton#ProcessButton {{
                background-color: {COLOR_PROCESS_RGBA};
            }}
            QPushButton#ProcessButton:hover {{
                background-color: {COLOR_PROCESS_HOVER};
            }}
            /* 步骤 3: 复制按钮 */
            QPushButton#CopyButton {{
                background-color: {COLOR_COPY_RGBA};
            }}
            QPushButton#CopyButton:hover {{
                background-color: {COLOR_COPY_HOVER};
            }}
            /* 禁用按钮样式 */
            QPushButton:disabled {{
                background-color: {COLOR_BORDER};
                color: {COLOR_TEXT_LIGHT};
                border: 1px solid {COLOR_BORDER};
                box-shadow: none;
            }}
            QLabel#FormatInfo {{
                font-size: 9pt;
                color: {COLOR_TEXT_LIGHT};
                padding: 10px 0 0 0;
            }}
        """
        # 设置全局字体，确保中文显示 (Windows 推荐使用 Microsoft YaHei UI)
        font = QFont("Microsoft YaHei UI", 10)
        self.setFont(font)
        self.setStyleSheet(qss)

    def initUI(self):
        """
        设置主窗口的布局和组件，采用简约、美观的布局，并实现窗口居中。
        """
        self.setWindowTitle('文献引用格式化工具 (Word专用)')
        # 初始设置几何尺寸
        self.setGeometry(100, 100, 1100, 750)

        # ⚠️ 需要添加的代码段：窗口居中逻辑
        # 获取屏幕的几何尺寸
        screen_geometry = QApplication.primaryScreen().geometry()
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        # 获取窗口的尺寸
        window_width = self.width()
        window_height = self.height()

        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 移动窗口到计算出的位置
        self.move(x, y)
        # ⚠️ 代码段结束

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(30, 15, 30, 30)  # 增加大体留白
        main_layout.setSpacing(20)

        # 标题
        title_label = QLabel("📚 文献列表格式化工具")
        title_label.setObjectName("TitleLabel")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        # 主内容布局：左侧控制面板 + 右侧输入/输出区域
        content_layout = QHBoxLayout()
        content_layout.setSpacing(25)  # 增大分栏间距

        # --- 左侧控制面板 ---
        control_panel = QFrame()
        control_panel.setObjectName("ControlPanel")
        control_panel.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        control_panel.setFixedWidth(280)  # 略微增加宽度
        control_panel_layout = QVBoxLayout(control_panel)
        control_panel_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        control_panel_layout.setSpacing(15)

        # 1. 操作步骤标题
        control_panel_layout.addWidget(QLabel("➡️ **格式化流程 (3 步)**"))
        control_panel_layout.addWidget(self._create_separator())

        # 2. 步骤按钮
        self.preview_button = QPushButton("1. 检查文献分割")
        self.preview_button.setObjectName("PreviewButton")
        self.preview_button.clicked.connect(self.split_preview)
        control_panel_layout.addWidget(self.preview_button)

        self.process_button = QPushButton("2. 统一格式并清洗")
        self.process_button.setObjectName("ProcessButton")
        self.process_button.clicked.connect(self.process_references)
        control_panel_layout.addWidget(self.process_button)

        self.copy_button = QPushButton("3. 复制 Word 专用格式")
        self.copy_button.setObjectName("CopyButton")
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        control_panel_layout.addWidget(self.copy_button)

        control_panel_layout.addSpacing(30)

        # 3. 固定格式信息
        format_info_label = QLabel(
            f"**格式说明:**<br/>"
            f"• 自动剥离旧编号<br/>"
            f"• 中英文分字体<br/>"
            f"• 固定编号格式: `{self.fixed_format_name}`"
        )
        format_info_label.setObjectName("FormatInfo")
        format_info_label.setWordWrap(True)
        control_panel_layout.addWidget(format_info_label)

        control_panel_layout.addStretch(1)

        # 4. 提示
        self.status_label = QLabel("💡 状态: 等待用户输入原始文献。")
        self.status_label.setObjectName("StatusLabel")
        self.status_label.setWordWrap(True)
        control_panel_layout.addWidget(self.status_label)

        content_layout.addWidget(control_panel)

        # --- 右侧输入/输出区域 ---
        io_area_layout = QVBoxLayout()
        io_area_layout.setSpacing(15)

        # 输入区域
        input_label = QLabel("📝 **原始文献输入区**")
        io_area_layout.addWidget(input_label)
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("请将文献列表粘贴到此处。程序将自动处理乱码、多行和旧编号...")
        self.input_text.setFont(QFont("Courier New", 10))
        io_area_layout.addWidget(self.input_text, 1)

        # 输出/预览区域
        output_label = QLabel("👁️ **格式化预览区**")
        io_area_layout.addWidget(output_label)
        self.output_text = QTextEdit()
        self.output_text.setObjectName("output_text")
        self.output_text.setReadOnly(True)
        self.output_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        io_area_layout.addWidget(self.output_text, 2)

        content_layout.addLayout(io_area_layout, 1)

        main_layout.addLayout(content_layout)

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
        # ⚠️ 样式修改：按钮背景颜色淡化 (使用浅灰蓝 #e0eaf1)
        DIALOG_QSS = (
            "QMessageBox { background-color: white; padding: 20px; }"
            "QLabel { margin-top: 5px; margin-bottom: 5px; }"
            # 针对 QMessageBox 中的 QPushButton 进行样式调整
            "QMessageBox QPushButton {"
            "background-color: #e0eaf1; "  # 淡雅的浅灰蓝
            "color: #33415c; "  # 使用深色文本，保证可读性
            "border: 1px solid #c8d3db; "  # 浅色边框
            "border-radius: 4px; "
            "padding: 5px 15px;"
            "}"
            # 悬停效果保持默认或略微加深
            "QMessageBox QPushButton:hover { background-color: #d1dde8; }"
        )

        raw_text = self.input_text.toPlainText()
        if not raw_text.strip():
            self.status_label.setText("⚠️ 输入为空，请粘贴文献文本。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("输入为空")
            msg.setText("请输入文献文本后进行分割预览。")
            msg.setStyleSheet(DIALOG_QSS)
            # ⚠️ 修改：新增设置弹窗图标
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
            # ⚠️ 修改：新增设置弹窗图标
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
            # ⚠️ 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

    def process_references(self):
        """
        第二步：调用处理器进行完整的格式化，并存储 Word HTML 结果。
        """
        # ⚠️ 样式修改：按钮背景颜色淡化 (使用浅灰蓝 #e0eaf1)
        DIALOG_QSS = (
            "QMessageBox { background-color: white; padding: 20px; }"
            "QLabel { margin-top: 5px; margin-bottom: 5px; }"
            "QMessageBox QPushButton {"
            "background-color: #e0eaf1; "
            "color: #33415c; "
            "border: 1px solid #c8d3db; "
            "border-radius: 4px; "
            "padding: 5px 15px;"
            "}"
            "QMessageBox QPushButton:hover { background-color: #d1dde8; }"
        )

        raw_text = self.input_text.toPlainText()
        if not raw_text.strip():
            self.status_label.setText("⚠️ 输入为空，请粘贴文献文本。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("输入为空")
            msg.setText("请输入文献文本后进行格式化。")
            msg.setStyleSheet(DIALOG_QSS)
            # ⚠️ 修改：新增设置弹窗图标
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
            # ⚠️ 修改：新增设置弹窗图标
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
            # ⚠️ 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

    def copy_to_clipboard(self):
        """
        第三步：复制格式化结果到剪贴板。
        """
        # ⚠️ 样式修改：按钮背景颜色淡化 (使用浅灰蓝 #e0eaf1)
        DIALOG_QSS = (
            "QMessageBox { background-color: white; padding: 20px; }"
            "QLabel { margin-top: 5px; margin-bottom: 5px; }"
            "QMessageBox QPushButton {"
            "background-color: #e0eaf1; "
            "color: #33415c; "
            "border: 1px solid #c8d3db; "
            "border-radius: 4px; "
            "padding: 5px 15px;"
            "}"
            "QMessageBox QPushButton:hover { background-color: #d1dde8; }"
        )

        if not self.html_output_for_clipboard:
            self.status_label.setText("⚠️ 请按顺序先进行 '检查' 和 '统一格式' 操作。")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("操作顺序提示")
            msg.setText("请先完成 '检查文献分割' 和 '统一格式并清洗'！")
            msg.setStyleSheet(DIALOG_QSS)
            # ⚠️ 修改：新增设置弹窗图标
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
            # ⚠️ 修改：新增设置弹窗图标
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.exec()

        except Exception as e:
            self.status_label.setText(f"❌ 步骤 3 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            # ⚠️ 修改：新增设置弹窗图标 (修复了原始代码中此处缺少 msg.exec() 的问题)
            msg.setWindowIcon(QIcon(self.ICON_PATH))
            msg.setWindowTitle("处理错误") # 补充了标题，使弹窗完整
            msg.setText(f"复制到剪贴板时出现错误：\n{str(e)}") # 补充了文本
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # ⚠️ 新增：设置应用程序的图标，确保任务栏和文件管理器中显示正确的图标
    # 路径使用 ReferenceFormatterApp 中定义的常量，确保一致性
    app.setWindowIcon(QIcon(ReferenceFormatterApp.ICON_PATH))

    # 为 QApplication 设置中文字体，确保全局中文显示正常
    font = QFont("Microsoft YaHei UI")
    app.setFont(font)

    ex = ReferenceFormatterApp()
    ex.show()
    sys.exit(app.exec())