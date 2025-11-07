# src/latex_converter.py
import re
from lxml import etree
from docx.oxml.ns import qn

# --- 1. OMML 命名空间和辅助函数 ---
M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M_PREFIX = "{%s}" % M_NAMESPACE


def _m_tag(tag_name: str) -> str:
    """为OMML标签添加正确的命名空间前缀。"""
    return M_PREFIX + tag_name


# --- 2. OMML 元素构建器 (更一致和健壮) ---
# 这些函数现在都期望接收一个 <m:r> 元素作为参数。

def _create_run_omml(text: str) -> etree._Element:
    """创建一个包含文本的基础OMML run (<m:r><m:t>...</m:t></m:r>)。"""
    mr = etree.Element(_m_tag('r'))
    mt = etree.SubElement(mr, _m_tag('t'))
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
    etree.SubElement(mrad, _m_tag('deg'))
    me = etree.SubElement(mrad, _m_tag('e'))
    me.append(base_run)
    return mrad


def _create_accent_omml(base_run: etree._Element, accent_char: str) -> etree._Element:
    """创建一个OMML重音 (<m:acc>) 元素，例如 hat。"""
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


# --- 3. 全新重构的原生LaTeX解析器 ---

SYMBOL_MAP = {
    '\\hbar': 'ħ', '\\partial': '∂', '\\omega': 'ω',
    '\\Psi': 'Ψ', '\\rangle': '⟩', '\\langle': '⟨',
}


class ParserState:
    """跟踪解析器在标记列表中的当前位置。"""

    def __init__(self, tokens: list[str]):
        self.tokens = tokens
        self.pos = 0

    def has_tokens(self) -> bool:
        return self.pos < len(self.tokens)

    def current_token(self) -> str:
        return self.tokens[self.pos] if self.has_tokens() else None

    def advance(self) -> str:
        token = self.current_token()
        self.pos += 1
        return token


def tokenize(latex_string: str) -> list[str]:
    """将LaTeX字符串分解为标记列表。"""
    token_regex = re.compile(r"(\\[a-zA-Z]+|[{}^]|[a-zA-Z0-9]|.)")
    return [token for token in token_regex.findall(latex_string) if not token.isspace()]


def _parse_argument(state: ParserState) -> etree._Element:
    """
    【核心修正】解析一个独立的LaTeX参数。
    一个参数要么是单个标记，要么是一个完整的 {...} 组。
    """
    if state.current_token() == '{':
        state.advance()  # Consume '{'
        return _parse_tokens(state, stop_char='}')
    else:
        # 解析单个标记作为参数
        return _parse_tokens(state, limit=1)


def _parse_tokens(state: ParserState, stop_char: str = None, limit: int = -1) -> etree._Element:
    """
    解析标记流直到遇到 stop_char 或达到 limit。
    这是递归下降解析器的核心。
    """
    current_run = etree.Element(_m_tag('r'))
    tokens_consumed = 0

    while state.has_tokens():
        if limit != -1 and tokens_consumed >= limit:
            break

        token = state.current_token()
        if token == stop_char:
            state.advance()  # Consume the stop character
            break

        state.advance()
        tokens_consumed += 1

        if token.startswith('\\'):  # --- 处理命令 ---
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

        elif token == '^':  # --- 处理上标 (后缀操作符) ---
            # 它作用于前一个元素，这使得逻辑复杂。
            # 简化：我们假设它作用于前一个子元素 run
            if len(current_run):
                base_element = current_run[-1]
                current_run.remove(base_element)
                base_run = etree.Element(_m_tag('r'))
                base_run.append(base_element)

                sup_run = _parse_argument(state)
                ssup_element = _create_superscript_omml(base_run, sup_run)
                current_run.append(ssup_element)
            else:
                print(f"警告: 发现没有基数的上标 '^'，已忽略。")

        else:  # --- 处理普通文本/字符 ---
            current_run.append(_create_run_omml(token))

    return current_run


def latex_to_omml(latex_string: str) -> etree._Element | None:
    """
    将LaTeX数学字符串完整地转换为一个独立的OMML XML元素。
    """
    try:
        tokens = tokenize(latex_string)
        if not tokens: return None
        state = ParserState(tokens)

        omml_math = etree.Element(_m_tag('oMath'))
        omml_para = etree.SubElement(omml_math, _m_tag('oMathPara'))

        result_run = _parse_tokens(state)
        omml_para.append(result_run)

        if state.has_tokens():
            print(f"警告: 解析在标记 '{state.current_token()}' 处提前结束。")

        return omml_math

    except IndexError:
        print(f"错误: LaTeX语法错误，可能缺少参数或括号不匹配: '{latex_string}'")
        return None
    except Exception as e:
        print(f"错误: 原生LaTeX转换器内部错误 -> {e} (原始LaTeX: '{latex_string}')")
        return None