# main.py
import os
import re
import sys
from config import Color as c
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
try:
    from window_effect import apply_acrylic_effect
    WINDOWS_EFFECT_AVAILABLE = True
except ImportError:
    WINDOWS_EFFECT_AVAILABLE = False

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
    QStyleOption, QStyle, QGraphicsDropShadowEffect
)
from PyQt6.QtGui import QFont, QDesktopServices, QPainter, QColor, QIcon
from PyQt6.QtCore import QMimeData, Qt, QUrl, QTimer, QSize
from PyQt6.QtCore import QPropertyAnimation, QEasingCurve, pyqtProperty
from PyQt6.QtGui import QColor


# 从我们创建的模块中导入核心处理类
from reference_processor import ReferenceProcessor, DOCX_AVAILABLE
from styles import build_main_window_qss, DIALOG_QSS


# pyinstaller --onefile --windowed --name "文献引用格式化工具" --icon "app_icon.ico" --add-data "app_icon.ico;." main.py



# **[START]** 新增函数：处理打包后的资源路径
def resource_path(relative_path):
    """
    获取资源文件的绝对路径，适配 PyInstaller 打包后的环境。
    如果程序是作为独立应用运行 (打包后)，它会查找临时目录下的资源。
    如果程序是直接从脚本运行 (未打包)，它会查找相对路径。
    """
    try:
        # 检查是否被 PyInstaller 打包
        if getattr(sys, 'frozen', False):
            # 打包后的运行路径
            base_path = sys._MEIPASS
        else:
            # 脚本运行时的路径 (当前脚本所在目录)
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        path = os.path.join(base_path, relative_path)
        if not os.path.exists(path):
            print(f"WARNING: Resource file not found: {path}")
        return path
    except Exception as e:
        print(f"ERROR in resource_path: {str(e)}")
        return relative_path  # 返回原始路径作为后备


# **[END]** 新增函数
class SmoothButton(QPushButton):
    """
    自定义丝滑按钮：
    1. 支持背景色渐变动画 (Hover 效果)
    2. 支持点击时的缩放/回弹效果 (Press 效果)
    """

    def __init__(self, text, parent=None, bg_color="#EBE7DD", hover_color="#D6D2C9", text_color="#2C2420"):
        super().__init__(text, parent)
        self._bg_color = QColor(bg_color)
        self._hover_color = QColor(hover_color)
        self._current_bg = self._bg_color
        self._text_color = text_color

        # 设置初始样式
        self.update_style()

        # 鼠标样式
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 阴影效果 - 更深的阴影
        self._shadow = QGraphicsDropShadowEffect(self)
        self._shadow.setBlurRadius(15)  # 增加模糊半径
        self._shadow.setOffset(0, 4)    # 增加垂直偏移
        self._shadow.setColor(QColor(0, 0, 0, 40))  # 增加不透明度
        self.setGraphicsEffect(self._shadow)

    @pyqtProperty(QColor)
    def background_color(self):
        return self._current_bg

    @background_color.setter
    def background_color(self, color):
        self._current_bg = color
        self.update_style()

    def update_style(self):
        # 动态更新样式表
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {self._current_bg.name()};
                color: {self._text_color};
                border: none;
                border-radius: 10px;
                padding: 12px 16px;
                font-weight: bold;
                font-size: 11pt;
                text-align: center;
                min-width: 200px;
            }}
            QPushButton:hover {{
                background-color: {self._hover_color.name()};
            }}
        """)

    def enterEvent(self, event):
        # 鼠标悬停：颜色渐变动画
        self.anim = QPropertyAnimation(self, b"background_color")
        self.anim.setDuration(200)  # 200ms 丝滑过渡
        self.anim.setStartValue(self._bg_color)
        self.anim.setEndValue(self._hover_color)
        self.anim.setEasingCurve(QEasingCurve.Type.OutQuad)
        self.anim.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        # 鼠标离开：颜色恢复
        self.anim = QPropertyAnimation(self, b"background_color")
        self.anim.setDuration(200)
        self.anim.setStartValue(self._hover_color)
        self.anim.setEndValue(self._bg_color)
        self.anim.setEasingCurve(QEasingCurve.Type.OutQuad)
        self.anim.start()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        # 点击：轻微下沉效果 (改变阴影和位置)
        self._shadow.setOffset(0, 0)
        self._shadow.setBlurRadius(2)
        self.move(self.x(), self.y() + 2)
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        # 释放：回弹
        self._shadow.setOffset(0, 3)
        self._shadow.setBlurRadius(12)
        self.move(self.x(), self.y() - 2)
        super().mouseReleaseEvent(event)


class MarqueeLabel(QLabel):
    """单行滚动标签，文本超出时自动滚动显示。"""

    def _add_shadow(self, widget, blur_radius=20, x_offset=0, y_offset=5, alpha=30):
        """为控件添加阴影效果"""
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(blur_radius)
        shadow.setXOffset(x_offset)
        shadow.setYOffset(y_offset)
        shadow.setColor(QColor(0, 0, 0, alpha))
        widget.setGraphicsEffect(shadow)

    def __init__(self, text="", parent=None, interval=15, step=1, gap="    "):
        super().__init__(text, parent)
        self._full_text = text or ""
        self._interval = interval
        self._step = step
        self._offset = 0
        self._need_scroll = False
        self._gap = gap  # Initialize _gap attribute
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

    # GitHub 仓库链接
    GITHUB_URL = "https://github.com/rhj-flash/Reference-Formatter-Tool"

    def __init__(self):
        super().__init__()
        self.processor = ReferenceProcessor()
        self.html_output_for_clipboard = ""
        self.fixed_format_name = "普通数字"

        # 设置窗口标志，保留最小化、最大化和关闭按钮，移除帮助按钮
        self.setWindowFlags(Qt.WindowType.Window | 
                          Qt.WindowType.WindowMinimizeButtonHint | 
                          Qt.WindowType.WindowMaximizeButtonHint |
                          Qt.WindowType.WindowCloseButtonHint |
                          Qt.WindowType.WindowTitleHint)

        # 1. 初始化 UI
        self.initUI()
        
        self._apply_global_style()

        # 3. 启动入场动画
        self.start_entrance_animation()

    # --- 新增：入场动画方法 ---
    def start_entrance_animation(self):
        """窗口组件的入场动画：从下往上浮现，且透明度渐变"""
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(2000)  # 0.8秒
        self.animation.setStartValue(0)
        self.animation.setEndValue(1)
        self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.animation.start()

        # 也可以让中央部件有个轻微的上浮效果
        central = self.centralWidget()
        self.pos_anim = QPropertyAnimation(central, b"pos")
        # 注意：这里需要获取当前位置稍微往下一点作为起点，需要 careful
        # 简单起见，我们只做 Opacity 渐变，既高级又稳健

    # In the ReferenceFormatterApp class, find the _open_github_link method (around line 288) and replace it with:
    def _open_github_link(self):
        """Open GitHub repository in default browser."""
        try:
            url = QUrl(self.GITHUB_URL)
            if not QDesktopServices.openUrl(url):
                print(f"ERROR: Failed to open URL: {self.GITHUB_URL}")
                self.status_label.setText("❌ 无法打开GitHub链接")
        except Exception as e:
            print(f"ERROR: Exception while opening GitHub link: {str(e)}")
            self.status_label.setText("❌ 打开GitHub链接时出错")

    # Then find where the GitHub button is created (around line 400-445) and replace that section with:
            # 添加GitHub按钮
            self.github_button = QPushButton()
            self.github_button.setObjectName("GithubButton")
            self.github_button.setCursor(Qt.CursorShape.PointingHandCursor)
            self.github_button.setFixedSize(36, 36)
            self.github_button.setToolTip("访问 GitHub 仓库")

            # 加载并设置GitHub图标
            github_icon_path = resource_path("github_icon.ico")
            icon = QIcon(github_icon_path)
            if not icon.isNull():
                self.github_button.setIcon(icon)
                self.github_button.setIconSize(QSize(24, 24))
                print(f"DEBUG: GitHub图标已加载: {github_icon_path}")
            else:
                print(f"WARNING: 无法加载GitHub图标: {github_icon_path}")

            # 设置样式表
            self.github_button.setStyleSheet("""
                QPushButton#GithubButton {
                    background-color: #F8F5ED;
                    border: none;
                    border-radius: 18px;
                    padding: 0;
                    margin: 0;
                    width: 36px;
                    height: 36px;
                }
                QPushButton#GithubButton:hover {
                    background-color: #F2EDE0;
                }
                QPushButton#GithubButton:pressed {
                    background-color: #E5E0D0;
                }
                QPushButton#GithubButton::icon {
                    padding: 0;
                    margin: 0;
                }
            """)

            # 添加阴影效果
            self._add_shadow(self.github_button, blur_radius=30, y_offset=8, alpha=100)
            
            # 连接点击事件
            self.github_button.clicked.connect(self._open_github_link)

    def _apply_global_style(self):
        """
        应用一套大胆、优雅、淡雅的全局 QSS 样式 (Glassmorphism 玻璃磨砂风格)。
        """
        # 设置全局字体，确保中文显示 (Windows 推荐使用 Microsoft YaHei UI)
        font = QFont("Microsoft YaHei UI", 10)
        self.setFont(font)
        self.setStyleSheet(build_main_window_qss())

    def _create_hero_card(self):
        """创建顶部的介绍卡片，背景设置为透明"""
        hero_card = QFrame()
        hero_card.setObjectName("HeroCard")
        hero_card.setStyleSheet("background: transparent;")
        hero_layout = QHBoxLayout(hero_card)
        hero_layout.setContentsMargins(0, 0, 0, 0)
        hero_layout.setSpacing(4)

        text_layout = QVBoxLayout()
        text_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text_layout.setSpacing(0)  # Remove spacing between elements
        text_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins
        
        # Remove the title label completely since it's not needed
        subtitle = QLabel("3 步完成从杂乱引用到正式 Word 交付的全流程，让排版、字体、缩进一次搞定。")
        subtitle.setObjectName("HeroSubtitle")
        font = subtitle.font()
        font.setBold(True)
        font.setItalic(True)
        subtitle.setFont(font)
        subtitle.setWordWrap(False)
        subtitle.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # Remove any extra spacing around the text
        subtitle.setStyleSheet("margin: 0; padding: 0;")
        subtitle.setMaximumHeight(subtitle.fontMetrics().height())
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        text_layout.addWidget(subtitle, alignment=Qt.AlignmentFlag.AlignTop)

        hero_layout.addLayout(text_layout, stretch=1)
        # 添加非常明显的四周阴影效果
        self._add_shadow(hero_card, blur_radius=400, y_offset=0, alpha=180)

        return hero_card

    def initUI(self):
        """
        设置主窗口的布局和组件，添加字体设置面板。
        """
        self.setWindowTitle('文献引用格式化工具@rhj_flash')

        # 设置窗口图标
        # **[START]** 修改主窗口图标加载逻辑
        icon_path = resource_path('app_icon.ico')

        # 设置窗口图标
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            print(f"DEBUG: 主窗口图标已加载: {icon_path}")
        else:
            print(f"ERROR: 主窗口图标文件未找到: {icon_path}")
        # **[END]** 修改主窗口图标加载逻辑

        # 设置更舒适的窗口尺寸比例
        window_width = 1250
        window_height = 780
        self.resize(window_width, window_height)

        # 窗口居中逻辑：使用可用屏幕区域的宽高来计算，使窗口在视觉上居中

        screen = QApplication.primaryScreen()
        if screen is not None:
            # 获取屏幕的几何信息（包括任务栏）
            screen_geometry = screen.geometry()
            # 获取可用屏幕区域（排除任务栏）
            available_geometry = screen.availableGeometry()

            # 计算窗口居中位置
            x = available_geometry.x() + (available_geometry.width() - window_width) // 2
            y = available_geometry.y() + (available_geometry.height() - window_height) // 2

            # 确保窗口不会移出屏幕
            x = max(0, min(x, screen_geometry.width() - window_width))
            y = max(0, min(y, screen_geometry.height() - window_height))

            # 调试输出
            print(f"DEBUG: Screen geometry = {screen_geometry.width()}x{screen_geometry.height()} @ ({screen_geometry.x()}, {screen_geometry.y()})")
            print(f"DEBUG: Available geometry = {available_geometry.width()}x{available_geometry.height()} @ ({available_geometry.x()}, {available_geometry.y()})")
            print(f"DEBUG: Window size = {window_width}x{window_height}")
            print(f"DEBUG: Moving window to x={x}, y={y}")

            self.move(x, y)

        central_widget = QWidget()
        central_widget.setObjectName("centralWidget")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 设置主布局的边距和间距
        main_layout.setContentsMargins(18, 10, 18, 12)
        main_layout.setSpacing(10)

        # 创建顶部栏容器
        top_bar = QWidget()
        top_bar.setObjectName("TopBar")
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(10, 10, 10, 10)
        top_bar_layout.setSpacing(15)

        # 添加GitHub按钮
        self.github_button = QPushButton()
        self.github_button.setObjectName("GithubButton")
        self.github_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.github_button.setFixedSize(36, 36)
        self.github_button.setToolTip("访问 GitHub 仓库")

        # 加载并设置GitHub图标
        github_icon_path = resource_path("github_icon.ico")
        icon = QIcon(github_icon_path)
        if not icon.isNull():
            self.github_button.setIcon(icon)
            self.github_button.setIconSize(QSize(24, 24))
            print(f"DEBUG: GitHub图标已加载: {github_icon_path}")

        # 设置样式表
        self.github_button.setStyleSheet("""
            QPushButton#GithubButton {
                background-color: #F8F5ED;
                border: none;
                border-radius: 18px;
                padding: 0;
                margin: 0;
                width: 36px;
                height: 36px;
            }
            QPushButton#GithubButton:hover {
                background-color: #F2EDE0;
            }
            QPushButton#GithubButton:pressed {
                background-color: #E5E0D0;
            }
            QPushButton#GithubButton::icon {
                padding: 0;
                margin: 0;
            }
        """)

        # 添加阴影效果
        self._add_shadow(self.github_button, blur_radius=30, y_offset=8, alpha=100)

        # 连接点击事件
        self.github_button.clicked.connect(self._open_github_link)

        # 创建标题标签
        title_label = QLabel("文献列表格式化工具")
        title_label.setObjectName("TitleLabel")
        title_label.setStyleSheet("""
            QLabel#TitleLabel {
                font-size: 50px;
                font-weight: bold;
                color: #2C3E50;
                margin: 0;
                padding: 5px 0;
                qproperty-alignment: AlignCenter;
            }
        """)

        # 添加GitHub按钮到左侧
        top_bar_layout.addWidget(self.github_button)

        # 添加弹簧使标题居中
        top_bar_layout.addStretch(1)

        # 添加标题
        top_bar_layout.addWidget(title_label)

        # 添加另一个水平弹簧使标题保持居中
        top_bar_layout.addStretch(1)

        # 添加顶部栏到主布局
        main_layout.insertWidget(0, top_bar)

        hero_card = self._create_hero_card()
        # 移除标题标签，只保留副标题
        hero_card.layout().itemAt(0).itemAt(0).widget().setVisible(False)
        main_layout.addWidget(hero_card)

        # --- 字体设置面板 - 调整布局 ---
        font_settings_layout = QHBoxLayout()
        font_settings_layout.setSpacing(8)

        # 英文字体设置
        english_font_layout = QHBoxLayout()
        english_font_layout.setContentsMargins(0, 0, 0, 0)
        english_font_layout.setSpacing(6)
        english_font_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        # 移除之前添加的 GitHub 按钮相关代码

        # GitHub 按钮右侧预留一段固定空白，用于视觉分隔
        english_font_layout.addSpacing(20)

        # 中间放“英文字体”文字和下拉框
        english_font_label = QLabel("<b><i>英文字体：</i></b>")
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
        self.english_font_combo.setFixedHeight(28)
        self._add_shadow(self.english_font_combo, blur_radius=15, y_offset=3, alpha=40)
        english_font_layout.addWidget(self.english_font_combo)

        # 英文字号
        english_size_layout = QHBoxLayout()
        english_size_label = QLabel("<b><i>英文字号：</i></b>")
        english_size_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        english_size_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        english_size_layout.addWidget(english_size_label)

        self.english_size_combo = QComboBox()
        self.english_size_combo.addItems(["10", "10.5", "11", "12", "14", "16", "18"])
        self.english_size_combo.setCurrentText("12")
        self.english_size_combo.setFixedHeight(28)
        self._add_shadow(self.english_size_combo, blur_radius=15, y_offset=3, alpha=40)
        english_size_layout.addWidget(self.english_size_combo)

        # 中文字体设置
        chinese_font_layout = QHBoxLayout()
        chinese_font_label = QLabel("<b><i>中文字体：</i></b>")
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
        self.chinese_font_combo.setFixedHeight(28)
        self._add_shadow(self.chinese_font_combo, blur_radius=15, y_offset=3, alpha=40)
        chinese_font_layout.addWidget(self.chinese_font_combo)

        # 中文字号
        chinese_size_layout = QHBoxLayout()
        chinese_size_label = QLabel("<b><i>中文字号：</i></b>")
        chinese_size_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        chinese_size_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        chinese_size_layout.addWidget(chinese_size_label)

        self.chinese_size_combo = QComboBox()
        self.chinese_size_combo.addItems(["10", "10.5", "11", "12", "14", "16", "18"])
        self.chinese_size_combo.setCurrentText("12")
        self.chinese_size_combo.setFixedHeight(28)
        self._add_shadow(self.chinese_size_combo, blur_radius=15, y_offset=3, alpha=40)
        chinese_size_layout.addWidget(self.chinese_size_combo)

        # --- 新增：编号格式设置 ---
        num_format_layout = QHBoxLayout()
        num_format_label = QLabel("<b><i>编号格式：</i></b>")
        num_format_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)
        num_format_label.setStyleSheet("font-size: 12pt; font-weight: 600;")
        num_format_layout.addWidget(num_format_label)

        self.num_format_combo = QComboBox()
        # 修改：添加更多格式，并确保 1. 和 [1] 在前两位
        self.num_format_combo.addItems([
            "[1]",  # 学术常用
            "1.",  # Word 默认
            "(1)",  # 括号式
            "1)",  # 半括号式
            "<1>",  # 尖括号
            "{1}"  # 花括号
        ])
        self.num_format_combo.setCurrentText("[1]")
        self.num_format_combo.setFixedHeight(34)
        self._add_shadow(self.num_format_combo, blur_radius=15, y_offset=3, alpha=40)
        num_format_layout.addWidget(self.num_format_combo)

        # 将新的布局加入到主布局 (注意：需要把 num_format_layout 加入 font_settings_layout)
        font_settings_layout.addLayout(english_font_layout)
        font_settings_layout.addLayout(english_size_layout)
        font_settings_layout.addLayout(chinese_font_layout)
        font_settings_layout.addLayout(chinese_size_layout)
        font_settings_layout.addLayout(num_format_layout)
        font_settings_layout.addStretch(1)

        main_layout.addLayout(font_settings_layout)

        # 主内容布局：左侧控制面板 + 右侧输入/输出区域
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)

        # --- 左侧控制面板 - 调整宽度 ---
        control_panel = QFrame()
        control_panel.setObjectName("ControlPanel")
        self._add_shadow(control_panel, blur_radius=35, y_offset=8, alpha=60)
        control_panel.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        control_panel.setFixedWidth(260)
        control_panel_layout = QVBoxLayout(control_panel)
        control_panel_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        control_panel_layout.setSpacing(8)

        # 1. 操作步骤标题
        step_title = QLabel("➡️ <b>格式化流程 (3 步)</b>")
        step_title.setStyleSheet("font-size: 13pt; margin-bottom: 10px;")
        control_panel_layout.addWidget(step_title)
        control_panel_layout.addWidget(self._create_separator())

        # 按钮容器，用于控制间距
        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)
        button_layout.setContentsMargins(8, 20, 8, 20)  # 增加上下边距
        button_layout.setSpacing(16)  # 增加按钮之间的间距

        # 2. 步骤按钮
        # 按钮 1 - 与背景色相同，深色阴影
        self.preview_button = SmoothButton("1. 检查文献分割",
                                         bg_color="#F8F5ED",  # 修正为有效的6位颜色代码
                                         hover_color="#F2EDE0",  # 悬停时深一点
                                         text_color="black")  # 白色文字
        self.preview_button.clicked.connect(self.split_preview)
        self.preview_button.setMinimumHeight(50)
        button_layout.addWidget(self.preview_button)
        button_layout.addSpacing(35)  # 添加间距

        # 按钮 2 - 与背景色相同，深色阴影
        self.process_button = SmoothButton("2. 统一格式并清洗",
                                         bg_color="#F8F5ED",  # 修正为有效的6位颜色代码
                                         hover_color="#F2EDE0",  # 悬停时深一点
                                         text_color="black")  # 白色文字
        self.process_button.clicked.connect(self.process_references)
        self.process_button.setMinimumHeight(50)
        button_layout.addWidget(self.process_button)
        button_layout.addSpacing(35)  # 添加间距

        # 按钮 3 - 与背景色相同，深色阴影
        self.export_button = SmoothButton("3. 生成 Word 文件",
                                        bg_color="#F8F5ED",  # 修正为有效的6位颜色代码
                                        hover_color="#F2EDE0",  # 悬停时深一点
                                        text_color="black")  # 白色文字
        self.export_button.clicked.connect(self.export_to_word_file)
        self.export_button.setMinimumHeight(50)  # 增加高度
        button_layout.addWidget(self.export_button)

        # 添加按钮容器到主布局
        control_panel_layout.addWidget(button_container)
        control_panel_layout.addStretch(1)  # 将按钮推到顶部

        control_panel_layout.addSpacing(10)

        # 4. 当前字体设置显示 - 添加编号格式和分割线
        # 先创建空的标签
        self.font_info_label = QLabel()
        self.font_info_label.setWordWrap(True)
        self.font_info_label.setFixedHeight(110)
        self.font_info_label.setObjectName("FormatInfo")
        # 然后初始化显示内容
        self.update_font_info()
        # 设置点击事件
        self.font_info_label.mousePressEvent = self.on_font_info_clicked
        control_panel_layout.addWidget(self.font_info_label)

        control_panel_layout.addStretch(1)

        # 5. 提示
        self.status_label = MarqueeLabel("    💡状态: 等待用户输入原始文献。      ")
        self.status_label.setObjectName("StatusLabel")
        self.status_label.setFixedHeight(50)  # Increased height from 40 to 50
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # 添加阴影效果
        self.status_label._add_shadow(self.status_label, blur_radius=20, y_offset=5, alpha=50)
        control_panel_layout.addWidget(self.status_label)

        content_layout.addWidget(control_panel)

        # --- 右侧输入/输出区域 - 调整比例 ---
        io_area_layout = QVBoxLayout()
        io_area_layout.setSpacing(10)


        # 输入区域
        input_label = QLabel("<b>📝 原始文献输入区</b>")
        io_area_layout.addWidget(input_label)
        self.input_text = QTextEdit()
        # 使用 HTML 格式的占位文本
        placeholder_html = """
        <div style="color: #666666; font-size: 12px; line-height: 1.5;">
            请将需要处理的参考文献粘贴在此处，例如：<br><br>
            [1] Smith J, Johnson B. Title of the paper. Journal Name. 2020;15(3):123-145.<br>
            [2] 张三, 李四. 论文标题. 期刊名称. 2021;12(4):56-78.<br>
            [3] Wang L, et al. Another example. Nature. 2019;567(7748):305-312.<br><br><br>
            支持以下格式：<br><br>
            • 带编号的文献（如 [1] 或 1. 开头）<br>
            • 无编号的文献列表<br>
            • 中英文混排的文献<br>
            • 多篇文献（每行一个或空行分隔）
            
        </div>
        """
        self.input_text.setPlaceholderText(" ")  # 设置一个空格的占位符
        self.input_text.setHtml(placeholder_html)
        self.input_text.setObjectName("InputTextEdit")

        # 格式化预览区
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setObjectName("PreviewTextEdit")

        # 添加预览区提示
        preview_placeholder = """
        <div style="color: #666666; font-size: 12px; line-height: 1.5; padding: 15px;">
            格式化预览区将显示处理后的文献列表，例如：<br><br>
            [1] Smith J, Johnson B. Title of the paper. Journal Name. 2020;15(3):123-145.<br>
            [2] 张三, 李四. 论文标题. 期刊名称. 2021;12(4):56-78.<br><br>
            使用步骤：<br>
            • 在左侧输入区粘贴您的文献列表<br>
            • 点击"1. 检查文献分割"预览分割效果<br>
            • 确认后点击"2. 统一格式并清洗"进行格式化<br>
            • 最后点击"3. 生成 Word 文件"导出结果<br><br>
            <span style="color: #7f8c8d;">提示：您可以在上方设置中调整中英文字体、字号和编号格式</span>
        </div>
        """
        self.preview_text.setHtml(preview_placeholder)

        # 🟢 修改 2：改用 Consolas 字体，字号加大到 12
        # Consolas 是 Windows 下非常优秀的等宽编程字体，比 Courier New 更好看
        font_style = QFont("Consolas", 12)
        self.input_text.setFont(font_style)

        self._add_shadow(self.input_text, blur_radius=35, alpha=60)
        io_area_layout.addWidget(self.input_text, 1)

        # 输出/预览区域
        output_label = QLabel("<b>👁️ 格式化预览区</b>")
        io_area_layout.addWidget(output_label)
        self.output_text = QTextEdit()
        self.output_text.setObjectName("output_text")
        self.output_text.setReadOnly(True)

# ... (rest of the code remains the same)
        # 🟢 修改 3：同步应用新字体
        self.output_text.setFont(font_style)

        self.output_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self._add_shadow(self.output_text, blur_radius=35, alpha=60)
        io_area_layout.addWidget(self.output_text, 1)

        content_layout.addLayout(io_area_layout, 1)

        main_layout.addLayout(content_layout)

        # 连接字体变化的信号
        self.english_font_combo.currentTextChanged.connect(self.update_font_info)
        self.english_size_combo.currentTextChanged.connect(self.update_font_info)
        self.chinese_font_combo.currentTextChanged.connect(self.update_font_info)
        self.chinese_size_combo.currentTextChanged.connect(self.update_font_info)
        self.num_format_combo.currentTextChanged.connect(self.update_font_info)

    def update_font_info(self):
        """更新字体设置显示，包含编号格式和分割线"""
        if hasattr(self, "font_info_label") and self.font_info_label:
            self.font_info_label.setText(
                f"\n"
                f"**当前格式设置:** 📝 (点击刷新)\n"
                f"———————————————————————————————\n"
                f"• 英文: {self.english_font_combo.currentText()} {self.english_size_combo.currentText()}pt\n"
                f"• 中文: {self.chinese_font_combo.currentText()} {self.chinese_size_combo.currentText()}pt\n"
                f"• 编号: {self.num_format_combo.currentText()}"
                f"———————————————————————————————\n"
                f"\n"
            )
            self.font_info_label.setWordWrap(True)
            self.font_info_label.setFixedHeight(240)  
            # 设置对象名称以便样式应用
            self.font_info_label.setObjectName("FormatInfo")

    def on_font_info_clicked(self, event):
        """点击字体设置区域时刷新显示"""
        self.update_font_info()
        # 可选：给用户一个视觉反馈
        self.status_label.setText("🔄 设置信息已刷新")

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
            msg.exec()


        except Exception as e:
            self.status_label.setText(f"❌ 步骤 1 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("处理错误")
            msg.setText(f"文献分割时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
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
            msg.exec()
            return

        try:
            selected_format = self.fixed_format_name
            # 获取当前选中的编号格式
            selected_numbering = self.num_format_combo.currentText()

            # 传入 numbering_format 参数
            word_html_output, plain_text_output, was_stripped = self.processor.process_text(
                raw_text, selected_format, numbering_format=selected_numbering
            )

            self.html_output_for_clipboard = word_html_output
            self.output_text.setPlainText(plain_text_output)  

            stripped_message = "自动剥离了旧编号" if was_stripped else "未检测到旧编号"
            self.status_label.setText(f"🎉 步骤 2 完成：格式统一，Word结果已就绪! ({stripped_message})")

            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setWindowTitle("统一格式并清洗 (2/3)")
            # 更新提示文本，去掉复制相关的描述
            msg.setText(f"🎉 文献列表已格式化！\n\n"
                        f"• 当前预览区显示的是最终纯文本结果。\n"
                        f"• 编号格式已应用：{selected_numbering}\n"
                        f"• {stripped_message}，并应用了中英文分字体等样式。\n\n"
                        f"➡️ 下一步：点击 '生成 Word 文件' 按钮。")
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()

        except Exception as e:
            self.status_label.setText(f"❌ 步骤 2 错误: {str(e)[:50]}...")
            # 样式应用
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("处理错误")
            msg.setText(f"格式化时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()

    def export_to_word_file(self):
        """
        第三步：将格式化结果导出为Word文件，使用自定义字体设置。
        """
        # 1. 检查是否有数据
        if not self.html_output_for_clipboard:
            self.status_label.setText("⚠️ 请先完成格式化操作。")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("操作顺序提示")
            msg.setText("请先完成 '检查文献分割' 和 '统一格式并清洗'！")
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()
            return

        try:
            # 2. 弹出文件保存对话框
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

            # 3. 获取用户界面上的设置
            english_font = self.english_font_combo.currentText()
            english_size = float(self.english_size_combo.currentText())
            chinese_font = self.chinese_font_combo.currentText()
            chinese_size = float(self.chinese_size_combo.currentText())

            # [新增] 获取用户选择的编号格式
            selected_numbering = self.num_format_combo.currentText()

            # 创建自定义格式配置
            custom_format = {
                "language": "chinese",
                "line_spacing": 1.5,
                "font_size": english_size,
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
                    f"编号格式: {selected_numbering}"
                ]
            }

            # 4. 调用处理器导出 (传入 selected_numbering)
            success = self.processor.export_to_word_file_with_custom_font(
                self.html_output_for_clipboard,
                file_path,
                custom_format,
                selected_numbering,  
                english_font,
                english_size,
                chinese_font,
                chinese_size
            )

            # 5. 处理结果
            if success:
                self.word_file_path = file_path
                file_name = os.path.basename(file_path)
                self.status_label.setText(f"✅ Word文件已生成: {file_name}")

                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Icon.Information)
                msg.setWindowTitle("生成 Word 文件 (3/3)")  
                msg.setText(f"✨ Word文件已成功生成！\n\n"
                            f"文件位置: {file_path}\n"
                            f"编号格式: {selected_numbering}\n"  
                            f"英文字体: {english_font} {english_size}pt\n"
                            f"中文字体: {chinese_font} {chinese_size}pt")
                msg.setStyleSheet(DIALOG_QSS)
                msg.exec()
            else:
                raise Exception("Word文件生成失败")

        except PermissionError as e:
            # 权限错误处理
            self.status_label.setText(f"❌ 文件保存失败: 权限被拒绝")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("文件保存错误")
            msg.setText(
                f"无法保存文件：\n{str(e)}\n\n请确保：\n1. 文件没有被其他程序打开\n2. 您有该位置的写入权限\n3. 文件路径正确")
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()
        except Exception as e:
            # 其他错误处理
            self.status_label.setText(f"❌ 导出失败: {str(e)[:50]}...")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("导出错误")
            msg.setText(f"生成Word文件时出现错误：\n{str(e)}")
            msg.setStyleSheet(DIALOG_QSS)
            msg.exec()

    def _add_shadow(self, widget, blur_radius=10, alpha=30, y_offset=0):
        """
        为控件添加阴影效果
        :param widget: 要添加阴影的控件
        :param blur_radius: 模糊半径
        :param alpha: 阴影透明度 (0-255)
        :param y_offset: y轴偏移量，默认为0（四周阴影）
        """
        shadow = QGraphicsDropShadowEffect(widget)
        shadow.setBlurRadius(blur_radius)
        shadow.setColor(QColor(0, 0, 0, alpha))
        shadow.setOffset(0, y_offset)  # 设置y轴偏移，0表示四周阴影
        widget.setGraphicsEffect(shadow)

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 不设置应用程序图标，使用系统默认 设置中文字体，确保全局中文显示正常
    font = QFont("Microsoft YaHei UI")
    app.setFont(font)

    ex = ReferenceFormatterApp()
    ex.show()
    sys.exit(app.exec())