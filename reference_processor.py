# reference_processor.py

import re
import html
from io import StringIO
# Pygments components for stable HTML formatting and segmentation
from pygments.formatters import HtmlFormatter
from pygments.lexer import RegexLexer
from pygments.token import Text

# 定义一个字典，用于进行全角到半角的转换
CHAR_MAPPING = {
    # 数字
    '０': '0', '１': '1', '２': '2', '３': '3', '四': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
    # 括号和方括号
    '（': '(', '）': ')', '【': '[', '】': ']',
    # 分隔符
    '—': '-', '–': '-', '―': '-', '．': '.', '：': ':', '；': ';',
    # 其他常见全角符号
    '，': ',', '。': '.', '？': '?', '！': '!', '＠': '@', '＃': '#',
    '＄': '$', '％': '%', '＾': '^', '＆': '&', '＊': '*', '＋': '+',
    '＝': '=', '～': '~', '　': ' '
}


# --- Pygments 自定义 Lexer (无需修改) ---

class ReferenceBlockLexer(RegexLexer):
    """
    一个简单的 Pygments Lexer，用于将每一行文本识别为一个块。
    我们只利用 Pygments 的分块能力来配合 Formatter。
    """
    name = 'RefBlock'
    aliases = ['refblock']
    filenames = ['*.ref']

    tokens = {
        'root': [
            # 匹配一整行，并将其标记为 Text.Line 类型
            (r'.*\n', Text.Line),
            (r'.*$', Text.Line),  # 匹配最后一行没有换行符的情况
        ],
    }


# --- Pygments 自定义 Formatter (无需修改) ---

class ReferenceBlockFormatter(HtmlFormatter):
    """
    自定义 Pygments HTML Formatter，用于显示格式化后的文献分割预览。
    """
    # 启用全行着色
    full_lines = True

    def __init__(self, **options):
        super().__init__(style='default', **options)
        # 定义文献块样式
        self.BLOCK_STYLES = [
            {
                'bg_color': '#f0f8ff',  # 浅蓝色
                'border_color': '#b0d0e0',
                'header_bg': '#d0e8f0'
            },
            {
                'bg_color': '#fff8f0',  # 浅橙色
                'border_color': '#e0c0b0',
                'header_bg': '#f0e0d0'
            },
            {
                'bg_color': '#f8f8f0',  # 浅黄色
                'border_color': '#d0d0b0',
                'header_bg': '#e8e8d0'
            }
        ]
        # 从 options 中取出预先计算的行标记数据
        self.lines_with_markers = options.pop('lines_with_markers', [])

    def format_unencoded(self, tokensource, outfile):
        """
        覆盖格式化方法，显示格式化后的文献分割。
        """
        # 如果 outfile 为 None，使用 StringIO 作为缓冲区
        if outfile is None:
            string_buffer = StringIO()
            real_outfile = string_buffer
            should_return_string = True
        else:
            real_outfile = outfile
            should_return_string = False

        # 开始包装div
        real_outfile.write(
            '<div style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.5; padding: 10px;">')

        current_block_lines = []
        block_index = 0

        def render_block(block_lines, block_num):
            """渲染一个完整的格式化文献块"""
            if not block_lines:
                return ""

            # 选择样式（循环使用）
            style = self.BLOCK_STYLES[block_num % len(self.BLOCK_STYLES)]

            # 文献标题 - 显示格式化后的状态
            block_header = (
                f'<div style="background: {style["header_bg"]}; padding: 10px 15px; '
                f'font-weight: bold; border-radius: 8px 8px 0 0; '
                f'border: 2px solid {style["border_color"]}; border-bottom: 1px dashed {style["border_color"]}; '
                f'color: #333; font-size: 12pt;">'
                f'📖 文献 {block_num + 1} </div>'
            )

            # 文献内容
            block_content = []
            for line in block_lines:
                if line.strip():  # 只处理非空行
                    # 注意：这里不需要再转义，因为内容已经是格式化后的
                    line_html = (
                        f'<div style="padding: 8px 15px; '
                        f'border-left: 2px solid {style["border_color"]}; '
                        f'border-right: 2px solid {style["border_color"]}; '
                        f'font-family: \'Courier New\', monospace; '
                        f'white-space: pre-wrap;">'
                        f'{line}'
                        f'</div>'
                    )
                    block_content.append(line_html)

            # 包装整个文献块
            return (
                f'<div style="margin: 25px 0; border-radius: 8px; box-shadow: 0 3px 6px rgba(0,0,0,0.1);">'
                f'{block_header}'
                f'<div style="background-color: {style["bg_color"]}; '
                f'border: 2px solid {style["border_color"]}; border-top: none; '
                f'border-radius: 0 0 8px 8px;">'
                f'{"".join(block_content)}'
                f'</div>'
                f'</div>'
            )

        # 处理所有行
        for i, (line, is_new_block_start) in enumerate(self.lines_with_markers):

            if is_new_block_start and current_block_lines:
                # 渲染当前块
                real_outfile.write(render_block(current_block_lines, block_index))
                # 开始新块
                current_block_lines = [line]
                block_index += 1
            else:
                current_block_lines.append(line)

        # 渲染最后一个块
        if current_block_lines:
            real_outfile.write(render_block(current_block_lines, block_index))

        # 显示统计信息
        if block_index >= 0:
            real_outfile.write(
                f'<div style="text-align: center; color: #666; padding: 15px; '
                f'background: #f0f0f0; border-radius: 5px; margin-top: 20px; '
                f'font-size: 10pt;">'
                f'✅ 共检测到 {block_index + 1} 篇格式化文献'
                f'</div>'
            )

        real_outfile.write('</div>')

        # 如果使用了缓冲区，返回缓冲区内容
        if should_return_string:
            return string_buffer.getvalue()

class ReferenceProcessor:
    """
    一个专门用于处理和格式化学术文献引用的类。
    实现了中英文混合字体格式化、多编号格式支持和智能自动分割预览。
    """

    # 定义支持的编号格式及其配置。
    # ⚠️ 关键修改：只保留格式的描述作为键，移除了“编号”二字
    SUPPORTED_FORMATS = {
        "方括号 (推荐，需替换I.为[1])": {
            "plain_prefix": "[{}] ",
            "html_type": "upper-roman"
        },
        "普通数字 (Word默认)": {
            "plain_prefix": "{}. ",
            "html_type": "decimal"
        },
        "括号数字 (需替换a.为(1))": {
            "plain_prefix": "({}) ",
            "html_type": "lower-alpha"
        },
    }

    # 定义中英文所需的字体
    CHINESE_FONT = 'SimSun'
    ENGLISH_FONT = 'Times New Roman'

    def __init__(self):
        """
        初始化处理器，编译用于检测序号的正则表达式。
        """
        # 正则表达式匹配 [1], (1), 1. 三种格式的序号
        self.numbering_pattern = re.compile(r'^\s*(\[\d+\]|\(\d+\)|\d+\.)\s*')
        # 匹配 CJK 统一汉字，用于判断语言
        self.cjk_pattern = re.compile(r'[\u4e00-\u9fff]')
        print("DEBUG: ReferenceProcessor initialized.")

    def get_supported_formats(self) -> dict:
        """
        返回支持的格式列表，供UI使用。
        """
        return self.SUPPORTED_FORMATS

    def process_text(self, raw_text: str, format_name: str) -> tuple[str, str, bool]:
        """
        处理用户输入的完整原始文本。

        Args:
            raw_text (str): 原始输入的文献文本。
            format_name (str): 用户选择的编号格式名称。

        Returns:
            tuple[str, str, bool]: (Word HTML输出, UI纯文本输出, 是否成功剥离了现有编号)
        """
        print(f"DEBUG: Starting to process raw text with format: {format_name}...")

        if not raw_text.strip():
            print("DEBUG: Raw text is empty.")
            return "", "", False

        lines = raw_text.strip().split('\n')

        styled_html_references = []  # 1. 带 <span> 样式的 HTML (用于 Word)
        plain_text_references = []  # 2. 纯文本 (用于 UI 预览)

        was_stripped = False

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 1. 剥离已存在的序号
            cleaned_line, stripped, _ = self._strip_numbering(line)
            if stripped:
                was_stripped = True

            # 2. 对文献内容进行字符和符号规范化
            normalized_line = self._normalize_characters(cleaned_line)

            # 3. 生成两种输出
            # (用于 UI 预览): 存储转义后的纯文本
            plain_text_references.append(html.escape(normalized_line))

            # (用于 Word): 存储带 <span> 样式的 HTML
            styled_html_content = self._apply_mixed_font_styles(normalized_line)
            styled_html_references.append(styled_html_content)

        print(f"DEBUG: Stripped existing numbering: {was_stripped}")

        # 获取选定的格式配置
        format_config = self.SUPPORTED_FORMATS.get(format_name)
        # ⚠️ 修正了格式配置获取逻辑，如果找不到则使用第一个默认格式
        if not format_config:
            print(f"WARNING: Format name '{format_name}' not found. Using default format.")
            default_key = list(self.SUPPORTED_FORMATS.keys())[0]
            format_config = self.SUPPORTED_FORMATS[default_key]

        # 生成两种格式的文本
        html_output = self._generate_html_list(styled_html_references, format_config["html_type"])
        plain_text_output = self._generate_plain_text_list(plain_text_references, format_config["plain_prefix"])

        print("DEBUG: Processing finished. HTML and Plain Text generated.")
        return html_output, plain_text_output, was_stripped

    def _strip_numbering(self, line: str) -> tuple[str, bool, int]:
        """
        检查并从单个文献行中剥离已存在的序号。

        Args:
            line (str): 原始文献行。

        Returns:
            tuple[str, bool, int]: (清洗后的行, 是否被剥离, 剥离出的序号[整数])
        """
        match = self.numbering_pattern.match(line)
        if match:
            # 提取匹配到的序号字符串，并尝试转换为整数
            number_str = match.group(1).strip('[]() .')

            extracted_number = 0
            try:
                extracted_number = int(number_str)
            except ValueError:
                pass

            return line[match.end():].strip(), True, extracted_number

        return line.strip(), False, 0

    def _is_chinese_line(self, line: str, threshold: float = 0.1) -> bool:
        """
        简单的语言识别：如果文本中汉字字符数超过总有效字符数的阈值，则认为是中文行。
        """
        effective_line = re.sub(r'[\s\W_]', '', line)
        if not effective_line:
            return False

        chinese_chars = len(self.cjk_pattern.findall(line))
        return chinese_chars / len(effective_line) > threshold

    def _normalize_characters(self, line: str) -> str:
        """
        对文献条目进行字符和符号规范化，确保英文使用半角标点，中文保留全角标点。
        """
        normalized_line = line
        # 1. 通用全角转半角 (数字、括号等)
        for full_char, half_char in CHAR_MAPPING.items():
            normalized_line = normalized_line.replace(full_char, half_char)

        parts = []
        # 2. 分割：中文 | 非中文
        pattern = re.compile(r'([\u4e00-\u9fff]+)|([^\u4e00-\u9fff]+)')
        for match in pattern.finditer(normalized_line):
            chinese_part = match.group(1)
            other_part = match.group(2)

            if chinese_part:
                # 对中文部分进行标点全角规范化
                parts.append(self._normalize_chinese_punctuation(chinese_part))
            elif other_part:
                # 对非中文部分（英文、数字、标点）进行标点半角规范化
                parts.append(self._normalize_english_punctuation(other_part))

        normalized_line = "".join(parts)
        # 3. 替换多余的空格为一个空格，并去除首尾空格
        normalized_line = re.sub(r'\s+', ' ', normalized_line).strip()
        return normalized_line

    def _apply_mixed_font_styles(self, line: str) -> str:
        """
        遍历一行文本，使用 <span> 标签为中英文/数字片段分别应用不同字体。
        """
        parts = []
        # 正则表达式分割：中文 | 英文/数字/空格 | 标点符号
        pattern = re.compile(r'([\u4e00-\u9fff]+)|([a-zA-Z0-9\s]+)|([^\u4e00-\u9fff\s]+)')

        for match in pattern.finditer(line):
            chinese_part = match.group(1)
            english_part = match.group(2)
            punctuation_part = match.group(3)

            if chinese_part:
                escaped_content = html.escape(chinese_part)
                parts.append(
                    f'<span style="font-family: \'{self.CHINESE_FONT}\'; '
                    f'mso-hansi-font-family: \'{self.CHINESE_FONT}\'; '
                    f'mso-bidi-font-family: \'{self.CHINESE_FONT}\'; '
                    f'mso-ascii-font-family: \'{self.CHINESE_FONT}\';">'
                    f'{escaped_content}</span>'
                )
            elif english_part:
                escaped_content = html.escape(english_part)
                parts.append(
                    f'<span style="font-family: \'{self.ENGLISH_FONT}\'; '
                    f'mso-hansi-font-family: \'{self.ENGLISH_FONT}\'; '
                    f'mso-bidi-font-family: \'{self.ENGLISH_FONT}\'; '
                    f'mso-ascii-font-family: \'{self.ENGLISH_FONT}\';">'
                    f'{escaped_content}</span>'
                )
            elif punctuation_part:
                escaped_content = html.escape(punctuation_part)
                # 简单判断是否为中文标点 (使用常见全角符号作为判断依据)
                is_chinese_punct = any(c in '，。：；！？' for c in punctuation_part)
                font = self.CHINESE_FONT if is_chinese_punct else self.ENGLISH_FONT
                parts.append(
                    f'<span style="font-family: \'{font}\'; '
                    f'mso-hansi-font-family: \'{font}\'; '
                    f'mso-bidi-font-family: \'{font}\'; '
                    f'mso-ascii-font-family: \'{font}\';">'
                    f'{escaped_content}</span>'
                )

        return "".join(parts)

    def _generate_html_list(self, styled_html_references: list[str], html_type: str) -> str:
        """
        生成 Word 可识别的 HTML 自动编号列表，确保字体和编号格式正确。
        """
        if not styled_html_references:
            return ""

        # 根据选定的格式设置 CSS 列表样式
        if html_type == "upper-roman":
            before_content = "'[' counter(list-counter) '] '"
            mso_list_type = "mso-list: l0 level1 lfo1;"
        elif html_type == "lower-alpha":
            before_content = "'(' counter(list-counter) ') '"
            mso_list_type = "mso-list: l0 level1 lfo1;"
        else:  # decimal
            before_content = "counter(list-counter) '. '"
            mso_list_type = "mso-list: l0 level1 lfo1;"

        # 构建列表项
        list_items = "".join([
            f'<li style="{mso_list_type} font-family: \'{self.ENGLISH_FONT}\'; '
            f'mso-hansi-font-family: \'{self.ENGLISH_FONT}\'; '
            f'mso-bidi-font-family: \'{self.ENGLISH_FONT}\'; '
            f'mso-ascii-font-family: \'{self.ENGLISH_FONT}\';">'
            f'{styled_html_content}</li>'
            for styled_html_content in styled_html_references
        ])

        # 包含 CSS 样式块
        style = f"""
        <style>
            ol {{
                list-style-type: none;
                counter-reset: list-counter;
                margin-left: 0;
                padding-left: 35px;
                font-family: '{self.ENGLISH_FONT}';
                mso-hansi-font-family: '{self.ENGLISH_FONT}';
                mso-bidi-font-family: '{self.ENGLISH_FONT}';
                mso-ascii-font-family: '{self.ENGLISH_FONT}';
            }}
            li {{
                counter-increment: list-counter;
                margin-bottom: 3px;
                position: relative;
                font-family: inherit;
            }}
            li::before {{
                content: {before_content};
                position: absolute;
                left: -35px;
                width: 30px;
                text-align: right;
                display: inline-block;
                font-family: '{self.ENGLISH_FONT}';
                mso-hansi-font-family: '{self.ENGLISH_FONT}';
                mso-bidi-font-family: '{self.ENGLISH_FONT}';
                mso-ascii-font-family: '{self.ENGLISH_FONT}';
            }}
            [style*="mso-list"] {{
                margin-left: 35px;
                font-family: '{self.ENGLISH_FONT}';
                mso-hansi-font-family: '{self.ENGLISH_FONT}';
                mso-bidi-font-family: '{self.ENGLISH_FONT}';
                mso-ascii-font-family: '{self.ENGLISH_FONT}';
            }}
        </style>
        """

        # 最终的 Word HTML 结构
        html_string = f"""
        <html xmlns:o="urn:schemas-microsoft-com:office:office"
              xmlns:w="urn:schemas-microsoft-com:office:word">
        <head>
            <meta charset="UTF-8">
            {style}
        </head>
        <body>
        <ol>{list_items}</ol>
        </body>
        </html>
        """
        return html_string

    def _generate_plain_text_list(self, references: list, plain_prefix: str) -> str:
        """
        生成用于在程序界面中预览的纯文本列表。
        """
        plain_list = [
            f"{plain_prefix.format(i)}{html.unescape(ref)}"
            for i, ref in enumerate(references, 1)
        ]
        return "\n".join(plain_list)

    def get_split_preview(self, raw_text: str) -> str:
        """
        检测潜在的文献块边界，并返回一个包含交替背景色的 HTML 字符串，用于预览区。
        使用 Pygments 的 Formatter 确保 HTML 渲染的可靠性。

        Args:
            raw_text (str): 原始输入的文献文本。

        Returns:
            str: 包含交替颜色分块的 HTML 字符串。
        """
        if not raw_text.strip():
            return ""

        lines = raw_text.split('\n')
        # 1. 检测边界，返回 (行, 是否为新块开始) 列表
        lines_with_markers = self._detect_boundary_lines(lines)

        # 2. 初始化 Pygments Formatter。将边界标记数据传递给自定义 Formatter
        formatter = ReferenceBlockFormatter(lines_with_markers=lines_with_markers)

        # 3. 初始化 Pygments Lexer
        lexer = ReferenceBlockLexer()

        # 4. 生成 HTML
        # 使用 formatter.format 并传入 outfile=None 来获取字符串输出
        html_output = html.unescape(formatter.format(lexer.get_tokens_unprocessed(0, raw_text), outfile=None))

        print("DEBUG: Pygments HTML generation finished.")
        return html_output


    def _is_likely_new_reference(self, line: str) -> bool:
        """
        判断一行是否可能是新文献的开始。
        """
        new_ref_indicators = [
            '作者：', '作者:', 'author:', 'authors:', '标题：', '标题:',
            'title:', 'journal', '期刊', 'vol.', '卷', 'no.', '期',
            'pp.', '页码', 'pages:', '年', 'year:', 'doi:', 'http'
        ]

        line_lower = line.lower()

        # 如果包含多个文献特征词，可能是新文献
        indicator_count = sum(1 for indicator in new_ref_indicators if indicator in line_lower)

        # 同时检查行长度（文献标题通常较短）
        return indicator_count >= 2 or (len(line.strip()) < 100 and indicator_count >= 1)

    def _detect_number_reset(self, prev_line: str, current_line: str) -> bool:
        """
        检测编号是否重置（从高编号跳到低编号）
        """
        _, prev_has_num, prev_num = self._strip_numbering(prev_line.strip())
        _, curr_has_num, curr_num = self._strip_numbering(current_line.strip())

        if prev_has_num and curr_has_num and prev_num > 0 and curr_num > 0:
            # 如果当前编号明显小于前一个编号（考虑编号重置的情况）
            if curr_num < prev_num and (prev_num - curr_num) > 5:
                return True

        return False
    def _normalize_english_punctuation(self, text: str) -> str:
        """
        英文标点规范化：确保使用半角标点
        """
        full_to_half = {
            '，': ',', '。': '.', '：': ':', '；': ';',
            '？': '?', '！': '!', '（': '(', '）': ')',
            '【': '[', '】': ']', '―': '-', '–': '-',
            '—': '-', '．': '.', '·': '.', '∶': ':',
            '＠': '@', '＃': '#', '＄': '$', '％': '%',
            '＾': '^', '＆': '&', '＊': '*', '＋': '+',
            '＝': '=', '～': '~', '　': ' '
        }

        full_digits = '０１２３４５６７８９'
        half_digits = '0123456789'
        digit_mapping = str.maketrans(full_digits, half_digits)

        result = text.translate(digit_mapping)
        for full, half in full_to_half.items():
            result = result.replace(full, half)

        return result

    def _normalize_chinese_punctuation(self, text: str) -> str:
        """
        中文标点规范化：确保使用全角标点
        """
        half_to_full = {
            ',': '，', '.': '。', ':': '：', ';': '；',
            '?': '？', '!': '！', '(': '（', ')': '）',
            '[': '【', ']': '】'
        }

        result = text
        for half, full in half_to_full.items():
            result = result.replace(half, full)

        return result

    def get_formatted_split_preview(self, raw_text: str, format_name: str) -> str:
        """
        先格式化文献，再进行分割预览。

        Args:
            raw_text (str): 原始输入的文献文本
            format_name (str): 用户选择的编号格式名称

        Returns:
            str: 包含格式化后文献分割预览的 HTML 字符串
        """
        if not raw_text.strip():
            return ""

        # 1. 先进行完整的格式化处理
        html_output, plain_text_output, was_stripped = self.process_text(raw_text, format_name)

        # 2. 从纯文本输出中获取格式化后的文献列表
        formatted_lines = plain_text_output.split('\n')

        # 3. 对格式化后的文本进行分割检测
        lines_with_markers = self._detect_boundary_lines(formatted_lines)

        # 4. 使用自定义 Formatter 生成分割预览
        formatter = ReferenceBlockFormatter(lines_with_markers=lines_with_markers)
        lexer = ReferenceBlockLexer()

        # 重新组合文本用于 Pygments 处理
        formatted_text = '\n'.join([line for line, _ in lines_with_markers])

        html_output = html.unescape(formatter.format(lexer.get_tokens_unprocessed(0, formatted_text), outfile=None))

        print("DEBUG: Formatted split preview generation finished.")
        return html_output

    def _detect_boundary_lines(self, lines: list[str]) -> list[tuple[str, bool]]:
        """
        改进的边界检测：基于格式化后的编号识别不同的文献。
        """
        lines_with_markers = []

        if not lines:
            return lines_with_markers

        # 强制第一行总是新块的开始
        lines_with_markers.append((lines[0], True))

        for i in range(1, len(lines)):
            current_line = lines[i]
            current_stripped = current_line.strip()

            is_new_block = False

            # 主要分割条件：检测到格式化后的编号模式
            if current_stripped:
                # 检查是否包含格式化后的文献编号
                if self._has_formatted_numbering(current_stripped):
                    is_new_block = True
                # 检查空行后的内容（可能是新文献）
                elif i > 0 and not lines[i - 1].strip() and current_stripped:
                    is_new_block = True

            # 保持原始行内容，并标记是否为新块开始
            lines_with_markers.append((current_line, is_new_block))

        return lines_with_markers

    def _has_formatted_numbering(self, line: str) -> bool:
        """
        检测行是否包含格式化后的编号模式。
        """
        # 匹配格式化后的常见编号模式
        formatted_patterns = [
            r'^\[\d+\]',  # [1]
            r'^\d+\.',  # 1.
            r'^\(\d+\)',  # (1)
        ]

        for pattern in formatted_patterns:
            if re.match(pattern, line.strip()):
                return True

        return False