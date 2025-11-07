# src/latex_converter.py
import re
from lxml import etree

# --- 1. OMML 命名空间和辅助函数 ---
M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M_PREFIX = "{%s}" % M_NAMESPACE


def _m_tag(tag_name: str) -> str:
    """为OMML标签添加正确的命名空间前缀。"""
    return M_PREFIX + tag_name


# --- 2. OMML 元素构建器 ---
# 这些函数负责创建特定类型的OMML XML结构。

def create_run_omml(text: str) -> etree._Element:
    """创建一个包含文本的OMML run (<m:r><m:t>...</m:t></m:r>)。"""
    mr = etree.Element(_m_tag('r'))
    mt = etree.SubElement(mr, _m_tag('t'))
    # Word需要保留空格，尤其是在操作符周围
    if text.startswith(' ') or text.endswith(' '):
        mt.set(qn('xml:space'), 'preserve')
    mt.text = text
    return mr


def create_fraction_omml(num_element: etree._Element, den_element: etree._Element) -> etree._Element:
    """创建一个OMML分数 (<m:f>) 元素。"""
    mf = etree.Element(_m_tag('f'))
    mnum = etree.SubElement(mf, _m_tag('num'))
    mden = etree.SubElement(mf, _m_tag('den'))
    mnum.append(num_element)
    mden.append(den_element)
    return mf


def create_sqrt_omml(base_element: etree._Element) -> etree._Element:
    """创建一个OMML平方根 (<m:rad>) 元素。"""
    mrad = etree.Element(_m_tag('rad'))
    mdeg = etree.SubElement(mrad, _m_tag('deg'))  # 空的deg表示平方根
    me = etree.SubElement(mrad, _m_tag('e'))
    me.append(base_element)
    return mrad


def create_accent_omml(base_element: etree._Element, accent_char: str) -> etree._Element:
    """创建一个OMML重音 (<m:acc>) 元素，例如 hat。"""
    macc = etree.Element(_m_tag('acc'))
    maccPr = etree.SubElement(macc, _m_tag('accPr'))
    mchr = etree.SubElement(maccPr, _m_tag('chr'))
    mchr.set(_m_tag('val'), accent_char)
    me = etree.SubElement(macc, _m_tag('e'))
    me.append(base_element)
    return macc


def create_superscript_omml(base_element: etree._Element, sup_element: etree._Element) -> etree._Element:
    """创建一个OMML上标 (<m:sSup>) 元素。"""
    msSup = etree.Element(_m_tag('sSup'))
    me = etree.SubElement(msSup, _m_tag('e'))
    msup = etree.SubElement(msSup, _m_tag('sup'))
    # 如果基础本身也是一个run，我们需要解包它
    for child in base_element:
        me.append(child)
    msup.append(sup_element)
    return msSup


# --- 3. 原生LaTeX解析器 ---

# 定义LaTeX特殊符号到Unicode的映射
SYMBOL_MAP = {
    '\\hbar': 'ħ', '\\partial': '∂', '\\omega': 'ω',
    '\\Psi': 'Ψ', '\\rangle': '⟩', '\\langle': '⟨',
    # 可以继续添加更多...
}


class ParserState:
    """一个简单的类，用于在解析过程中跟踪当前位置。"""

    def __init__(self, tokens: list[str]):
        self.tokens = tokens
        self.pos = 0

    def current_token(self) -> str | None:
        """获取当前标记，但不前进。"""
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def advance(self):
        """前进到下一个标记。"""
        self.pos += 1


def tokenize(latex_string: str) -> list[str]:
    """
    将LaTeX字符串分解为标记列表。
    例如: '\\frac{a}{b}' -> ['\\frac', '{', 'a', '}', '{', 'b', '}']
    """
    # 这个正则表达式匹配:
    # 1. \\后面跟字母的命令 (\\frac)
    # 2. 花括号、上标符 ({, }, ^)
    # 3. 单个字母或数字
    # 4. 其他任何单个字符 (如 =, +, |)
    token_regex = re.compile(r"(\\[a-zA-Z]+|[{}^]|[a-zA-Z0-9]|.)")
    return [token for token in token_regex.findall(latex_string) if not token.isspace()]


def parse_latex_tokens(state: ParserState) -> etree._Element:
    """
    解析标记列表并生成一个OMML run (<m:r>)。
    这是递归下降解析器的核心。
    """
    # 每个解析作用域的结果都放在一个OMML run中
    current_run = etree.Element(_m_tag('r'))

    while (token := state.current_token()) is not None and token != '}':
        state.advance()

        if token.startswith('\\'):  # --- 处理命令 ---
            if token in SYMBOL_MAP:
                current_run.append(create_run_omml(SYMBOL_MAP[token]))
            elif token == '\\frac':
                num_run = parse_latex_tokens(state)  # 递归解析分子
                den_run = parse_latex_tokens(state)  # 递归解析分母
                current_run.append(create_fraction_omml(num_run, den_run))
            elif token == '\\sqrt':
                base_run = parse_latex_tokens(state)  # 递归解析根号内的内容
                current_run.append(create_sqrt_omml(base_run))
            elif token == '\\hat':
                base_run = parse_latex_tokens(state)  # 递归解析要加hat的内容
                current_run.append(create_accent_omml(base_run, '^'))
            else:
                print(f"警告: 未知的LaTeX命令 '{token}'，已忽略。")

        elif token == '{':  # --- 处理分组 ---
            # 递归解析组内的内容，并将其子元素附加到当前run
            group_run = parse_latex_tokens(state)
            for child in group_run:
                current_run.append(child)

        elif token == '^':  # --- 处理上标 ---
            # 上标修饰前一个元素
            last_element = current_run[-1]
            if last_element is not None:
                current_run.remove(last_element)  # 移除最后一个元素
                sup_run = parse_latex_tokens(state)  # 解析上标内容
                # 创建上标元素并替换
                ssup_element = create_superscript_omml(last_element, sup_run)
                current_run.append(ssup_element)

        else:  # --- 处理普通文本/字符 ---
            current_run.append(create_run_omml(token))

    # 如果我们是因为遇到 '}' 而结束循环，需要消耗掉这个标记
    if state.current_token() == '}':
        state.advance()

    return current_run


def latex_to_omml(latex_string: str) -> etree._Element | None:
    """
    将LaTeX数学字符串完整地转换为一个独立的OMML XML元素。
    这是外部调用的主函数。

    Args:
        latex_string (str): 要转换的LaTeX格式的数学公式字符串。

    Returns:
        etree._Element | None: 转换成功则返回一个代表 <m:oMath> 的lxml元素，
                               否则返回None。
    """
    try:
        # 步骤 1: 词法分析
        tokens = tokenize(latex_string)
        if not tokens:
            return None
        state = ParserState(tokens)

        # 步骤 2: 创建OMML顶层结构
        omml_math = etree.Element(_m_tag('oMath'))
        # Word公式通常包裹在一个 oMathPara 中
        omml_para = etree.SubElement(omml_math, _m_tag('oMathPara'))

        # 步骤 3: 启动递归解析
        result_run = parse_latex_tokens(state)
        omml_para.append(result_run)

        # 步骤 4: 检查解析是否消耗了所有标记
        if state.current_token() is not None:
            print(f"警告: 解析在标记 '{state.current_token()}' 处提前结束，可能存在语法错误。")

        return omml_math

    except IndexError:
        # 当命令需要参数但没找到时 (例如 \frac{a} )，会触发IndexError
        print(f"错误: LaTeX语法错误，可能缺少参数或括号不匹配。原始文本: '{latex_string}'")
        return None
    except Exception as e:
        print(f"错误: 原生LaTeX转换失败 -> {e} (原始LaTeX: '{latex_string}')")
        return None


# --- 4. 导入 lxml.etree.qn 以支持 xml:space ---
# 这是使 set(qn(...)) 工作所必需的
from docx.oxml.ns import qn