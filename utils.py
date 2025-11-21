# utils.py
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtGui import QIcon
from styles import DIALOG_QSS

def show_message_box(parent, title, text, icon_type, icon_path=None):
    """
    显示一个标准样式的 QMessageBox。
    
    :param parent: 父窗口
    :param title: 标题
    :param text: 内容文本
    :param icon_type: 图标类型 (QMessageBox.Icon.Warning, QMessageBox.Icon.Information, etc.)
    :param icon_path: 窗口图标路径 (可选)
    """
    msg = QMessageBox(parent)
    msg.setIcon(icon_type)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setStyleSheet(DIALOG_QSS)
    if icon_path:
        msg.setWindowIcon(QIcon(icon_path))
    msg.exec()
