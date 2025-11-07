# src/latex_converter.py

# --- 1. Imports (for pylatexenc v2.x compatibility, ADDING ParsingState) ---
from pylatexenc.latexwalker import (
    LatexWalker, LatexCharsNode, LatexMacroNode, LatexGroupNode, LatexMathNode,
    LatexWalkerParseError, ParsingState  # <-- CRITICAL: ParsingState is imported here
)
from lxml import etree
from docx.oxml.ns import qn
import traceback

# --- 2. OMML Namespace and Builder Functions (Unchanged) ---
M_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M_PREFIX = "{%s}" % M_NAMESPACE


# ... (All _create_* helper functions remain exactly the same) ...
def _m_tag(tag_name: str) -> str: return M_PREFIX + tag_name


def _create_run(elements: list = None) -> etree._Element:
    run = etree.Element(_m_tag('r'))
    if elements:
        for elem in elements: run.append(elem)
    return run


def _create_text(text: str) -> etree._Element:
    mt = etree.Element(_m_tag('t'))
    if text.strip() != text: mt.set(qn('xml:space'), 'preserve')
    mt.text = text
    return mt


def _create_fraction(num_run: etree._Element, den_run: etree._Element) -> etree._Element:
    mf = etree.Element(_m_tag('f'))
    mnum, mden = etree.SubElement(mf, _m_tag('num')), etree.SubElement(mf, _m_tag('den'))
    mnum.append(num_run)
    mden.append(den_run)
    return mf


def _create_sqrt(base_run: etree._Element) -> etree._Element:
    mrad = etree.Element(_m_tag('rad'))
    etree.SubElement(mrad, _m_tag('deg'))
    me = etree.SubElement(mrad, _m_tag('e'))
    me.append(base_run)
    return mrad


def _create_accent(base_run: etree._Element, accent_char: str) -> etree._Element:
    macc = etree.Element(_m_tag('acc'))
    maccPr = etree.SubElement(macc, _m_tag('accPr'))
    mchr = etree.SubElement(maccPr, _m_tag('chr'))
    mchr.set(_m_tag('val'), accent_char)
    me = etree.SubElement(macc, _m_tag('e'))
    me.append(base_run)
    return macc


def _create_script(base_run: etree._Element, sub_run: etree._Element, sup_run: etree._Element) -> etree._Element:
    if sub_run is not None and sup_run is not None:
        tag = 'sSubSup'
    elif sub_run is not None:
        tag = 'sSub'
    else:  # sup_run must be not None
        tag = 'sSup'
    elem = etree.Element(_m_tag(tag))
    me = etree.SubElement(elem, _m_tag('e'))
    me.append(base_run)
    if sub_run is not None:
        msub = etree.SubElement(elem, _m_tag('sub'))
        msub.append(sub_run)
    if sup_run is not None:
        msup = etree.SubElement(elem, _m_tag('sup'))
        msup.append(sup_run)
    return elem


# --- 3. The Converter Class using pylatexenc (FINAL v2.x COMPATIBLE FIX) ---
class LatexToOmmlConverter:
    SIMPLE_SYMBOLS = {
        'hbar': 'ħ', 'partial': '∂', 'omega': 'ω', 'Psi': 'Ψ',
        'rangle': '⟩', 'langle': '⟨'
    }

    def __init__(self, latex_string: str):
        lw = LatexWalker(latex_string)

        # --- START OF THE DECISIVE v2.x FIX ---
        # Create a ParsingState object that explicitly starts in math mode.
        math_parsing_state = ParsingState(in_math_mode=True)

        # Pass this state object to the parser. This is the correct v2.x way.
        self.nodelist, _, _ = lw.get_latex_nodes(parsing_state=math_parsing_state)
        # --- END OF THE DECISIVE v2.x FIX ---

    # ... (The rest of the class: convert, _render_nodelist_to_run, _render_node, etc. remains IDENTICAL) ...
    def convert(self) -> etree._Element:
        final_run = self._render_nodelist_to_run(self.nodelist)
        omml_math = etree.Element(_m_tag('oMath'))
        omml_para = etree.SubElement(omml_math, _m_tag('oMathPara'))
        omml_para.append(final_run)
        return omml_math

    def _render_nodelist_to_run(self, nodelist) -> etree._Element:
        main_run = _create_run()
        for node in nodelist:
            element = self._render_node(node)
            if element is None: continue
            if element.tag == _m_tag('r'):
                for child in element: main_run.append(child)
            else:
                main_run.append(element)
        return main_run

    def _render_node(self, node) -> etree._Element | None:
        if node.isNodeType(LatexCharsNode): return _create_text(node.chars)
        if node.isNodeType(LatexGroupNode): return self._render_nodelist_to_run(node.nodelist)
        if node.isNodeType(LatexMacroNode): return self._render_macro_node(node)
        if node.isNodeType(LatexMathNode): return self._render_math_node(node)
        print(f"警告: 未处理的节点类型 '{type(node).__name__}'")
        return None

    def _render_macro_node(self, node: LatexMacroNode) -> etree._Element:
        macro_name = node.macroname
        args = node.nodeargs if node.nodeargs else []
        if macro_name in self.SIMPLE_SYMBOLS and not args: return _create_text(self.SIMPLE_SYMBOLS[macro_name])
        if macro_name == 'frac' and len(args) == 2:
            num_run = self._render_nodelist_to_run([args[0]])
            den_run = self._render_nodelist_to_run([args[1]])
            return _create_fraction(num_run, den_run)
        if macro_name == 'sqrt' and len(args) == 1:
            content_run = self._render_nodelist_to_run([args[0]])
            return _create_sqrt(content_run)
        if macro_name == 'hat' and len(args) == 1:
            content_run = self._render_nodelist_to_run([args[0]])
            return _create_accent(content_run, '^')
        print(f"警告: 不支持的宏或参数数量不正确: '\\{macro_name}'")
        return _create_text(f"[?\\{macro_name}?]")

    def _render_math_node(self, node: LatexMathNode) -> etree._Element:
        base_run, sub_run, sup_run = None, None, None
        if node.base is not None:
            base_run = self._render_nodelist_to_run(node.base.nodelist)
        else:
            return _create_text("?")
        if node.sub is not None: sub_run = self._render_nodelist_to_run(node.sub.nodelist)
        if node.sup is not None: sup_run = self._render_nodelist_to_run(node.sup.nodelist)
        return _create_script(base_run, sub_run, sup_run)


# --- 4. The Public Facade Function (Unchanged from the debugging upgrade) ---
def latex_to_omml(latex_string: str) -> tuple[etree._Element | None, str | None]:
    try:
        cleaned_latex = latex_string.encode('utf-8').decode('unicode_escape')
        converter = LatexToOmmlConverter(cleaned_latex)
        omml_element = converter.convert()
        return (omml_element, None)
    except LatexWalkerParseError as e:
        error_msg = f"LaTeX解析失败: {e}"
        print(f"错误: {error_msg} (原始LaTeX: '{latex_string}')")
        return (None, error_msg)
    except Exception as e:
        error_msg = f"OMML渲染器内部错误: {e}\n{traceback.format_exc()}"
        print(f"错误: {error_msg} (原始LaTeX: '{latex_string}')")
        return (None, f"OMML渲染器内部错误: {e}")