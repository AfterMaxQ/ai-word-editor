# src/doc_parser.py

import io
from typing import Dict, Any, List, Iterator, Union, Optional
from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _get_alignment_str(alignment) -> str:
    if alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return 'center'
    elif alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return 'right'
    elif alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return 'justify'
    return 'left'


def _get_color_hex(color_obj) -> Optional[str]:
    if color_obj and color_obj.rgb:
        return f"#{color_obj.rgb}"
    return None


def _parse_paragraph_properties(p: Paragraph) -> Dict[str, Any]:
    """
    提取段落级别的样式：对齐、间距、行高、缩进等。
    """
    props = {}

    # 1. 样式名
    if p.style and p.style.name:
        props['style'] = p.style.name

    # 2. 对齐方式
    if p.alignment is not None:
        props['alignment'] = _get_alignment_str(p.alignment)

    # 3. 间距 (转换为 px 或 pt 字符串供前端使用)
    p_format = p.paragraph_format
    if p_format.space_after is not None:
        if hasattr(p_format.space_after, 'pt'):
            props['spacing_after'] = f"{p_format.space_after.pt}pt"

    if p_format.space_before is not None:
        if hasattr(p_format.space_before, 'pt'):
            props['spacing_before'] = f"{p_format.space_before.pt}pt"

    if p_format.line_spacing is not None:
        props['line_spacing'] = p_format.line_spacing

    return props


def _parse_run_properties(run: Run) -> Dict[str, Any]:
    """
    提取 Run (文本片段) 级别的样式：字体、字号、颜色、粗体、斜体。
    注意：使用 ._element 访问底层 XML
    """
    props = {}
    font = run.font

    if font.name:
        props['font_family'] = font.name
    # 尝试处理中文字体 (East Asia Theme)
    # 修复：使用 ._element
    elif run._element.rPr is not None and run._element.rPr.rFonts is not None:
        xml = run._element.rPr.rFonts.xml
        if 'w:eastAsia="' in xml:
            try:
                start = xml.index('w:eastAsia="') + 12
                end = xml.index('"', start)
                props['font_family'] = xml[start:end]
            except:
                pass

    if font.size:
        props['font_size'] = f"{font.size.pt}pt"

    if font.bold:
        props['bold'] = True

    if font.italic:
        props['italic'] = True

    color_hex = _get_color_hex(font.color)
    if color_hex:
        props['color'] = color_hex

    return props


def _parse_paragraph(p: Paragraph) -> List[Dict[str, Any]]:
    """
    解析段落，支持分页符拆分和详细样式提取。
    注意：使用 ._element.iterchildren() 遍历
    """
    elements = []

    base_props = _parse_paragraph_properties(p)
    current_text = ""
    dominant_run_props = {}
    found_dominant = False

    # 修复：使用 p._element.iterchildren()
    for child in p._element.iterchildren():
        tag = child.tag

        # 1. 检测分页符/分栏符
        if tag.endswith('br'):
            type_attr = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            if type_attr == 'page':
                if current_text:
                    final_props = {**base_props, **dominant_run_props}
                    elements.append({"type": "paragraph", "text": current_text, "properties": final_props})
                    current_text = ""
                elements.append({"type": "page_break"})
                continue
            elif type_attr == 'column':
                if current_text:
                    final_props = {**base_props, **dominant_run_props}
                    elements.append({"type": "paragraph", "text": current_text, "properties": final_props})
                    current_text = ""
                elements.append({"type": "column_break"})
                continue

        # 2. 处理 Run (w:r)
        if tag.endswith('r'):
            run = Run(child, p)
            # 修复：使用 run._element.xml
            run_xml = run._element.xml

            # 检查 Run 内部的分页符
            if '<w:br' in run_xml and 'w:type="page"' in run_xml:
                if current_text:
                    final_props = {**base_props, **dominant_run_props}
                    elements.append({"type": "paragraph", "text": current_text, "properties": final_props})
                    current_text = ""
                elements.append({"type": "page_break"})
                if run.text: current_text += run.text
            elif '<w:br' in run_xml and 'w:type="column"' in run_xml:
                if current_text:
                    final_props = {**base_props, **dominant_run_props}
                    elements.append({"type": "paragraph", "text": current_text, "properties": final_props})
                    current_text = ""
                elements.append({"type": "column_break"})
                if run.text: current_text += run.text
            else:
                if run.text:
                    current_text += run.text
                    if not found_dominant:
                        run_props = _parse_run_properties(run)
                        if run_props:
                            dominant_run_props = run_props
                            found_dominant = True

    if current_text:
        final_props = {**base_props, **dominant_run_props}
        elements.append({"type": "paragraph", "text": current_text, "properties": final_props})

    if not elements and not current_text:
        elements.append({"type": "paragraph", "text": "", "properties": base_props})

    return elements


def _parse_table(t: Table) -> Dict[str, Any]:
    """
    解析表格及其样式
    """
    data = []
    for row in t.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        data.append(row_data)

    properties = {}
    if t.style and t.style.name:
        properties['style'] = t.style.name

    return {
        "type": "table",
        "data": data,
        "properties": properties
    }


def iter_block_items(parent: DocumentObject) -> Iterator[Union[Paragraph, Table]]:
    """
    按文档顺序迭代 Paragraph 和 Table
    修复：使用 parent._element.body
    """
    if isinstance(parent, DocumentObject):
        # 修复：使用 ._element 获取底层 XML 对象
        parent_elm = parent._element.body
    else:
        raise ValueError("Parent object is not a Document.")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def parse_docx_to_json(doc_bytes: bytes) -> Dict[str, Any]:
    """
    API 入口：解析 DOCX 字节流为 JSON
    """
    doc = Document(io.BytesIO(doc_bytes))

    doc_state: Dict[str, Any] = {
        "page_setup": {},
        "sections": [{"elements": []}]
    }

    current_elements = doc_state["sections"][0]["elements"]

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            parsed_items = _parse_paragraph(block)
            current_elements.extend(parsed_items)
        elif isinstance(block, Table):
            current_elements.append(_parse_table(block))

    return doc_state