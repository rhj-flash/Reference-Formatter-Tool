# styles.py

"""
集中存放 UI 主题相关的颜色、QSS 与弹窗样式，便于 main.py 统一引用。
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARROW_ICON_PATH = os.path.join(BASE_DIR, "assets", "down_arrow.svg").replace("\\", "/")

# —— 主题配色（恢复为莫兰迪并加入羊皮卷底色） —— #
COLOR_BACKGROUND = "#fdf3d7"
COLOR_CARD_BG = "#fff9ec"
COLOR_TEXT_DARK = "#554236"
COLOR_TEXT_LIGHT = "#8c756a"
COLOR_BORDER = "#e5d2a9"

COLOR_PREVIEW_RGBA = "rgba(162, 185, 188, 0.6)"
COLOR_PROCESS_RGBA = "rgba(197, 197, 138, 0.6)"
COLOR_COPY_RGBA = "rgba(216, 167, 134, 0.6)"
COLOR_EXPORT_RGBA = "rgba(134, 167, 216, 0.6)"

COLOR_PREVIEW_HOVER = "rgba(162, 185, 188, 0.8)"
COLOR_PROCESS_HOVER = "rgba(197, 197, 138, 0.8)"
COLOR_COPY_HOVER = "rgba(216, 167, 134, 0.8)"
COLOR_EXPORT_HOVER = "rgba(134, 167, 216, 0.8)"

# —— 弹窗样式（恢复原始淡雅样式） —— #
DIALOG_QSS = """
    QMessageBox { background-color: white; padding: 20px; }
    QLabel { margin-top: 5px; margin-bottom: 5px; }
    QMessageBox QPushButton {
        background-color: #e0eaf1;
        color: #33415c;
        border: 1px solid #c8d3db;
        border-radius: 4px;
        padding: 5px 15px;
    }
    QMessageBox QPushButton:hover { background-color: #d1dde8; }
"""


def build_main_window_qss() -> str:
    """返回恢复后的莫兰迪风格全局 QSS。"""

    return f"""
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

        QLabel#TitleLabel {{
            font-size: 18pt;
            font-weight: bold;
            color: {COLOR_TEXT_DARK};
            padding: 12px 0 8px 0;
        }}

        QLabel#HeroSubtitle {{
            color: {COLOR_TEXT_LIGHT};
            font-size: 10pt;
        }}

        QLabel#HeroBadge {{
            background-color: rgba(255, 255, 255, 0.6);
            color: {COLOR_TEXT_DARK};
            border-radius: 10px;
            padding: 3px 10px;
        }}

        QFrame#HeroCard {{
            background-color: rgba(255, 255, 255, 0.65);
            border-radius: 14px;
            padding: 14px;
        }}

        QFrame#FontSettings {{
            background-color: rgba(255, 255, 255, 0.6);
            border: 1px solid {COLOR_BORDER};
            border-radius: 12px;
            padding: 6px 10px;
        }}

        QLabel#StatusLabel {{
            color: {COLOR_TEXT_DARK};
            font-weight: 500;
            font-size: 10pt;
            padding: 4px 12px;
            border: none;
            border-radius: 20px;
            background-color: rgba(255,255,255,0.85);
            margin-top: 8px;
            min-height: 34px;
        }}

        QLabel#FormatInfo {{
            font-size: 9pt;
            color: {COLOR_TEXT_LIGHT};
            padding: 10px 0 0 0;
        }}

        QTextEdit {{
            border: none;
            border-radius: 10px;
            padding: 12px;
            background-color: {COLOR_CARD_BG};
            font-size: 10pt;
            line-height: 1.4;
        }}
        QTextEdit:focus {{
            border: 1px solid rgba(162, 185, 188, 1.0);
        }}
        QTextEdit#output_text {{
            background-color: {COLOR_CARD_BG};
        }}

        QFrame#ControlPanel {{
            background-color: rgba(255, 255, 255, 0.7);
            border: none;
            border-radius: 12px;
            padding: 20px;
        }}

        QPushButton {{
            color: {COLOR_TEXT_DARK};
            border: none;
            padding: 10px 14px;
            border-radius: 12px;
            font-weight: 600;
            font-size: 10pt;
            min-height: 34px;
            text-align: left;
            margin-bottom: 10px;
        }}
        QPushButton:hover {{
        }}
        QPushButton:pressed {{
            padding-top: 13px;
        }}

        QPushButton#PreviewButton {{
            background-color: {COLOR_PREVIEW_RGBA};
        }}
        QPushButton#PreviewButton:hover {{
            background-color: {COLOR_PREVIEW_HOVER};
        }}

        QPushButton#ProcessButton {{
            background-color: {COLOR_PROCESS_RGBA};
        }}
        QPushButton#ProcessButton:hover {{
            background-color: {COLOR_PROCESS_HOVER};
        }}

        QPushButton#CopyButton {{
            background-color: {COLOR_COPY_RGBA};
        }}
        QPushButton#CopyButton:hover {{
            background-color: {COLOR_COPY_HOVER};
        }}

        QPushButton#ExportButton {{
            background-color: {COLOR_EXPORT_RGBA};
        }}
        QPushButton#ExportButton:hover {{
            background-color: {COLOR_EXPORT_HOVER};
        }}

        QPushButton:disabled {{
            background-color: {COLOR_BORDER};
            color: {COLOR_TEXT_LIGHT};
            border: 1px solid {COLOR_BORDER};
        }}

        QComboBox {{
            background-color: rgba(255, 255, 255, 0.92);
            border: none;
            border-radius: 10px;
            padding: 4px 36px 4px 12px;
            font-size: 10pt;
            min-height: 30px;
            margin-bottom: 10px;
            color: {COLOR_TEXT_DARK};
        }}
        QComboBox:hover {{
            border: none;
        }}
        QComboBox:focus {{
            border: none;
        }}
        QComboBox::drop-down {{
            border: none;
            width: 32px;
            background: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 rgba(255, 255, 255, 0.9),
                stop:1 rgba(255, 248, 225, 0.9)
            );
            border-top-right-radius: 10px;
            border-bottom-right-radius: 10px;
        }}
        QComboBox::down-arrow {{
            image: url("{ARROW_ICON_PATH}");
            width: 16px;
            height: 16px;
            margin-right: 8px;
        }}
        QComboBox QAbstractItemView {{
            border: none;
            background: #fffaf0;
            selection-background-color: rgba(162, 185, 188, 0.25);
            selection-color: {COLOR_TEXT_DARK};
            outline: none;
        }}
    """
