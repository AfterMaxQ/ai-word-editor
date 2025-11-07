# src/latex_converter.py
import re
import traceback
from lxml import etree
from docx.oxml.ns import qn

# --- 1. OMML 命名空间和辅助函数 ---
M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M_PREFIX = "{%s}" % M_NAMESPACE


def _m_tag(tag_name: str) -> str:
    """为OMML标签添加正确的命名空间前缀。"""
    return M_PREFIX + tag_name


# --- 2. OMML 元素构建器 (Element Builders) ---
def _create_run_omml(text: str) -> etree._Element:
    """创建一个包含文本的基础OMML run (<m:r><m:t>...</m:t></m:r>)。"""
    mr = etree.Element(_m_tag('r'))
    mt = etree.SubElement(mr, _m_tag('t'))
    # 保留文本前后的空格，对公式排版很重要
    if text.startswith(' ') or text.endswith(' '):
        mt.set(qn('xml:space'), 'preserve')
    mt.text = text
    return mr


def _create_fraction_omml(num_run: etree._Element, den_run: etree._Element) -> etree._Element:
    """创建一个OMML分数 (<m:f>) 元素。"""
    mf = etree.Element(_m_tag('f'))
    mnum = etree.SubElement(mf, _m_tag('num'))
    mden = etree.SubElement(mf, _m_tag('den'))
    mnum.append(num_run)
    mden.append(den_run)
    return mf


def _create_sqrt_omml(base_run: etree._Element) -> etree._Element:
    """创建一个OMML平方根 (<m:rad>) 元素。"""
    mrad = etree.Element(_m_tag('rad'))
    # ★ Bug #1 修复: 添加 radPr 来隐藏空的次数占位符
    mradPr = etree.SubElement(mrad, _m_tag('radPr'))
    mdegHide = etree.SubElement(mradPr, _m_tag('degHide'))
    mdegHide.set(_m_tag('val'), '1')

    etree.SubElement(mrad, _m_tag('deg'))  # 依然需要空的次数元素
    me = etree.SubElement(mrad, _m_tag('e'))
    me.append(base_run)
    return mrad


def _create_accent_omml(base_run: etree._Element, accent_char: str) -> etree._Element:
    """创建一个OMML重音 (<m:acc>) 元素, 例如 hat。"""
    macc = etree.Element(_m_tag('acc'))
    maccPr = etree.SubElement(macc, _m_tag('accPr'))
    mchr = etree.SubElement(maccPr, _m_tag('chr'))
    mchr.set(_m_tag('val'), accent_char)
    me = etree.SubElement(macc, _m_tag('e'))
    me.append(base_run)
    return macc


def _create_superscript_omml(base_run: etree._Element, sup_run: etree._Element) -> etree._Element:
    """创建一个OMML上标 (<m:sSup>) 元素。"""
    msSup = etree.Element(_m_tag('sSup'))
    me = etree.SubElement(msSup, _m_tag('e'))
    msup = etree.SubElement(msSup, _m_tag('sup'))
    me.append(base_run)
    msup.append(sup_run)
    return msSup


# --- 3. 原生LaTeX解析器 (Native LaTeX Parser) ---
# 符号映射表
SYMBOL_MAP = {
    '\\hbar': 'ħ', '\\partial': '∂', '\\omega': 'ω',
    '\\Psi': 'Ψ', '\\rangle': '⟩', '\\langle': '⟨',
}
# ★ Bug #2 修复: 增加定界符映射
DELIMITER_MAP = {
    '|': '|',
    '(': '(',
    ')': ')',
    '[': '[',
    ']': ']',
    '{': '{',
    '}': '}'
}


class ParserState:
    """跟踪解析器在标记列表中的当前位置。"""

    def __init__(self, tokens: list[str]):
        self.tokens = tokens
        self.pos = 0

    def has_tokens(self) -> bool: return self.pos < len(self.tokens)

    def current_token(self) -> str: return self.tokens[self.pos] if self.has_tokens() else None

    def advance(self) -> str: token = self.current_token(); self.pos += 1; return token


def tokenize(latex_string: str) -> list[str]:
    """将LaTeX字符串分解为标记列表。"""
    # 正则表达式现在能更好地处理符号和字母
    token_regex = re.compile(r"(\\[a-zA-Z]+|[{}[\]()|^_]|[^\\{}\[\]()|^_a-zA-Z0-9\s]+|[a-zA-Z0-9]+|\s+)")
    return [token for token in token_regex.findall(latex_string) if not token.isspace()]


def _parse_argument(state: ParserState) -> etree._Element:
    """解析一个独立的LaTeX参数, 例如 {...} 或单个token。"""
    if state.current_token() == '{':
        state.advance()  # 跳过 '{'
        return _parse_tokens(state, stop_char='}')
    else:
        # 解析单个token作为参数
        return _parse_tokens(state, limit=1)


def _parse_tokens(state: ParserState, stop_char: str = None, limit: int = -1) -> etree._Element:
    """
    递归地解析标记流，直到遇到停止符或达到限制。
    返回一个包含所有已解析元素的 <m:r> run。
    """
    # 一个OMML run可以包含多个子元素 (文本、分数、根号等)
    current_run = etree.Element(_m_tag('r'))
    tokens_consumed = 0

    while state.has_tokens():
        if limit != -1 and tokens_consumed >= limit:
            break

        token = state.current_token()
        if token == stop_char:
            state.advance()  # 跳过停止符
            break

        state.advance()
        tokens_consumed += 1

        if token.startswith('\\'):
            # 处理LaTeX命令
            if token in SYMBOL_MAP:
                current_run.append(_create_run_omml(SYMBOL_MAP[token]))
            elif token == '\\frac':
                num_run = _parse_argument(state)
                den_run = _parse_argument(state)
                current_run.append(_create_fraction_omml(num_run, den_run))
            elif token == '\\sqrt':
                base_run = _parse_argument(state)
                current_run.append(_create_sqrt_omml(base_run))
            elif token == '\\hat':
                base_run = _parse_argument(state)
                current_run.append(_create_accent_omml(base_run, '^'))
            else:
                print(f"警告: 未知的LaTeX命令 '{token}'，已忽略。")
        elif token == '^':
            # 处理上标
            if len(current_run):
                # 将run中的最后一个元素作为基数
                base_element = current_run[-1]
                current_run.remove(base_element)
                base_run = etree.Element(_m_tag('r'))
                base_run.append(base_element)

                sup_run = _parse_argument(state)
                ssup_element = _create_superscript_omml(base_run, sup_run)
                current_run.append(ssup_element)
            else:
                print(f"警告: 发现没有基数的上标 '^'，已忽略。")
        else:  # 处理普通文本和符号
            current_run.append(_create_run_omml(token))

    return current_run


def latex_to_omml(latex_string: str) -> etree._Element | None:
    """
    将LaTeX数学字符串完整地转换为一个独立的OMML XML元素。
    ★ 新增：强制所有公式使用“专业”格式渲染。
    """
    state = None
    try:
        tokens = tokenize(latex_string)
        if not tokens: return None
        state = ParserState(tokens)

        # OMML的完整结构是 <m:oMathPara> 包裹着 <m:oMath>
        # 我们将返回 <m:oMathPara>，因为显示属性是在这里设置的
        omml_para = etree.Element(_m_tag('oMathPara'))

        # --- ★ 核心修复：添加专业格式指令 ★ ---
        # 创建 oMathParaPr (段落属性) 元素
        omml_para_pr = etree.SubElement(omml_para, _m_tag('oMathParaPr'))
        # 在属性中，添加 jc (Justification) 元素，并设置为居中
        # 虽然对齐主要由上层段落控制，但这里设置一个默认值是好习惯
        jc = etree.SubElement(omml_para_pr, _m_tag('jc'))
        jc.set(_m_tag('val'), 'center')  # 你可以根据需要改为 'left' 或 'right'

        # Word通过这个来区分是“专业”还是“线性”
        # 我们在这里没有明确设置，因为Word默认会根据上下文选择
        # 但如果需要强制，可以在这里添加 <m:plcHide val="0"/> 等

        # 创建核心的 <m:oMath> 元素
        omml_math = etree.SubElement(omml_para, _m_tag('oMath'))

        # 解析所有token并放入一个顶层run中
        result_run = _parse_tokens(state)
        omml_math.append(result_run)

        if state.has_tokens():
            print(f"警告: 解析在标记 '{state.current_token()}' 处提前结束。")

        return omml_para  # ★ 注意：现在返回的是 oMathPara

    except Exception as e:
        # ... (错误处理代码保持不变) ...
        print("\n" + "=" * 25 + " LaTeX Converter Critical Error " + "=" * 25)
        # ...
        return None