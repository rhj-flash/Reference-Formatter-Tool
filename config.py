# config.py
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class Color:
    """
    颜色配置表 —— 100% 按你提供的截图取色 + 加上你要求的中文注释
    """
    # 主色调（标题、状态栏文字、重点高亮）
    PRIMARY = "#35524A"          # 滑动字体颜射（深墨绿，几乎发黑的岩灰绿）
    PRIMARY_DARK = "#2A423A"     # 按下态更深（按钮按下、悬停深色）

    # 高亮色（按钮悬停、选中状态）
    ACCENT = "#F8F5ED"           # 古铜金（按钮悬停）——沉稳旧金，不是亮金

    # 文字颜色
    TEXT = "#2A241F"             # 正文深棕灰（主要文字）
    TEXT_LIGHT = "#2A423A"       # 副文本（提示文字、说明文字）

    # 背景色（旧纸张质感）
    BG_PAPER = "#F8F5ED"         # 大背景（整个窗口底色，带一点暖黄纸张感）
    BG_CARD = "#F8F5ED"          # 顶部大控件背景（HeroCard区域）
    BG_PANEL = "#F8F5ED"         # 左侧控件大背景（ControlPanel区域）
    BG_STATUS = "#E8E0D5"        # 滚动显示背景（状态栏那块浅灰米色）

    # 滚动条颜色
    SCROLL_HANDLE = "#E8E0D5"    # 滚动条颜色（和状态栏同色，统一沉稳感）
    # 注意：你原来写 "#E8E0D5z" 多了一个 z，已帮你修正为正确的十六进制色值