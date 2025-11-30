# 文件：styles.py
from config import Color as c

DIALOG_QSS = f"""
    QMessageBox {{ background: {c.BG_CARD}; border: none; border-radius: 16px; }}
    QMessageBox QLabel {{ color: {c.TEXT}; font-family: "Microsoft YaHei"; font-size: 10pt; }}
    QMessageBox QPushButton {{
        background: {c.PRIMARY}; color: white; border: none;
        border-radius: 10px; padding: 10px 30px; font-weight: 600;
    }}
    QMessageBox QPushButton:hover {{ background: {c.PRIMARY_DARK}; }}
"""

def build_main_window_qss() -> str:
    return f"""
    QMainWindow {{
        background: {c.BG_PAPER}; 
        border: none;
    }}
    
    /* 垂直滚动条 */
    QScrollBar:vertical {{
        background: {c.BG_PAPER};
        width: 4px;
        margin: 0px;
    }}
    QScrollBar::handle:vertical {{
        background: #C8C4BB;
        min-height: 20px;
        border-radius: 2px;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    
    /* 水平滚动条 */
    QScrollBar:horizontal {{
        background: {c.BG_PAPER};
        height: 4px;
        margin: 0px;
    }}
    QScrollBar::handle:horizontal {{
        background: #C8C4BB;
        min-width: 20px;
        border-radius: 2px;
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}

    QLabel {{
        color: {c.TEXT};
        font-family: "Microsoft YaHei";
        font-size: 10pt; font-weight: 500;
    }}
    
    QLabel#TitleLabel {{
        font-family: Georgia; font-size: 30pt; font-weight: bold; color: {c.PRIMARY};
    }}

    QFrame#HeroCard, QFrame#ControlPanel {{
        background: {c.BG_CARD}; border: none; border-radius: 18px;
    }}

    QLabel#StatusLabel {{
        background: {c.BG_STATUS}; color: {c.PRIMARY};
        border-radius: 12px; padding: 14px 20px; font-weight: bold; font-size: 11pt;
    }}

    QTextEdit {{
        background: #F8F5ED; border: none; border-radius: 14px;
        padding: 18px; font-size: 10pt;
        selection-background-color: {c.ACCENT};
    }}
    QTextEdit:focus {{ background: #FFF8E8; }}
    QTextEdit#PreviewTextEdit {{ background: #F2ECDE; }}
    QTextEdit#PreviewTextEdit:focus {{ background: #F8F5ED; }}

    QPushButton {{
        background: white; border: none; border-radius: 14px;
        padding: 14px 28px; font-weight: 600; color: {c.TEXT}; font-size: 11pt;
    }}
    QPushButton:hover {{ background: {c.ACCENT}; color: white; }}
    QPushButton:pressed {{ background: {c.PRIMARY_DARK}; }}

    /* ==================== 关键修改：下拉框矮 + ▼ 箭头一定显示 ==================== */
    QComboBox {{
        background: {c.BG_PAPER};  /* 使用主背景色 */
        border: none;  
        border-radius: 8px;
        padding: 4px 30px 4px 12px;
        font-family: "Microsoft YaHei";
        font-size: 10.5pt;
        font-weight: 600;
        color: {c.TEXT};
        min-height: 20px;
        max-height: 28px;
        text-align: center;  /* 使下拉框中的文本居中 */
    }}
    QComboBox:hover, QComboBox:focus {{
        background: {c.BG_PAPER};
    }}

    /* 下拉箭头区域 */
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: center right;
        width: 24px;                            
        border: none;
        padding-right: 8px;
        background: transparent;
    }}

    /* 彻底清除 QComboBox::down-arrow 样式 */
    QComboBox::down-arrow {{
        /* **[START]** 修改：彻底隐藏 QComboBox 箭头 */
        image: none;
        background: none;
        border: none;
        width: 0px;
        height: 0px;
        padding: 0px;
        margin: 0px;
        /* **[END]** 修改 */
    }}

    /* 下拉箭头样式 */
    QComboBox::down-arrow {{
        width: 16px;
        height: 16px;
        margin-right: 4px;
        image: url(dropdown-arrow.svg);
    }}

    QComboBox QAbstractItemView {{
        background: white;
        selection-background-color: {c.PRIMARY};
        selection-color: white;
        border: none;
        border-radius: 10px;
        padding: 6px;
        font-family: "Microsoft YaHei";
        font-size: 11pt;
        outline: none;
    }}
    QComboBox QAbstractItemView::item {{
        text-align: center;
        padding: 4px 8px;
    }}

    /* 三个大按钮 */
    QPushButton#CheckButton,
    QPushButton#FormatButton,
    QPushButton#ExportButton {{
        font-family: "Microsoft YaHei";
        font-size: 13pt;
        font-weight: 700;
        padding: 18px 32px;
        min-height: 40px;
    }}

    QPushButton#GithubButton:hover {{
        background: rgba(53, 82, 74, 0.12);
        border-radius: 16px;
    }}
    """