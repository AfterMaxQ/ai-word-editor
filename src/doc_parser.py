# src/doc_parser.py

import io
from typing import Dict, Any, List, Iterator, Union
from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph


def _parse_paragraph(p: Paragraph) -> Dict[str, Any]:
    """
    Parses a python-docx Paragraph object into our standard ParagraphElement dictionary.

    Args:
        p (Paragraph): The paragraph object to parse.

    Returns:
        Dict[str, Any]: A dictionary conforming to the ParagraphElement schema.
    """
    properties = {}
    if p.style and p.style.name and p.style.name != 'Normal':
        properties['style'] = p.style.name

    # NOTE: More detailed property parsing can be added here.

    return {
        "type": "paragraph",
        "text": p.text,
        "properties": properties
    }


def _parse_table(t: Table) -> Dict[str, Any]:
    """
    Parses a python-docx Table object into our standard TableElement dictionary.

    Args:
        t (Table): The table object to parse.

    Returns:
        Dict[str, Any]: A dictionary conforming to the TableElement schema.
    """
    data = []
    for row in t.rows:
        row_data = [cell.text for cell in row.cells]
        data.append(row_data)

    properties = {}
    if t.style and t.style.name and t.style.name != 'Table Grid':
        properties['style'] = t.style.name

    return {
        "type": "table",
        "data": data,
        "properties": properties
    }


def parse_docx_to_json(doc_bytes: bytes) -> Dict[str, Any]:
    """
    【已重构 v2】解析上传的 .docx 文件内容为我们的标准 DocumentModel JSON 格式。
    此版本使用正确的 python-docx 高级迭代器 `iter_block_items` 来确保
    段落和表格能够被正确识别和排序，修复了导致解析出0个元素的 bug。

    Args:
        doc_bytes (bytes): .docx 文件的字节内容。

    Returns:
        Dict[str, Any]: 一个符合 DocumentModel schema 的字典。
    """
    doc = Document(io.BytesIO(doc_bytes))

    doc_state: Dict[str, Any] = {
        "page_setup": {},
        "numbering_definitions": [],
        "sections": [{"elements": []}]
    }

    current_elements = doc_state["sections"][0]["elements"]

    # --- [CORE FIX] ---
    # Use the high-level iter_block_items() iterator. This correctly yields
    # Paragraph and Table objects in document order.
    for block in doc.iter_block_items():
        if isinstance(block, Paragraph):
            # The 'block' variable is already a proper Paragraph object.
            # We only process paragraphs that contain some non-whitespace text.
            if block.text.strip():
                current_elements.append(_parse_paragraph(block))
        elif isinstance(block, Table):
            # The 'block' variable is already a proper Table object.
            current_elements.append(_parse_table(block))

    return doc_state


def iter_block_items(parent: DocumentObject) -> Iterator[Union[Paragraph, Table]]:
    """
    【核心工具函数】按文档中的自然顺序迭代块级元素（段落和表格）。

    python-docx 原生不提供按顺序混合访问段落和表格的方法。
    此函数通过遍历底层的 XML 元素子节点，并根据标签类型（w:p 或 w:tbl）
    实例化相应的高级对象来实现这一功能。

    Args:
        parent (DocumentObject): 文档对象。

    Yields:
        Union[Paragraph, Table]: 按顺序产出的段落或表格对象。
    """
    if isinstance(parent, DocumentObject):
        parent_elm = parent.element.body
    else:
        raise ValueError("Parent object is not a Document.")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _parse_paragraph(p: Paragraph) -> Dict[str, Any]:
    """
    Parses a python-docx Paragraph object into our standard ParagraphElement dictionary.
    """
    properties = {}
    if p.style and p.style.name and p.style.name != 'Normal':
        properties['style'] = p.style.name

    return {
        "type": "paragraph",
        "text": p.text,
        "properties": properties
    }


def _parse_table(t: Table) -> Dict[str, Any]:
    """
    Parses a python-docx Table object into our standard TableElement dictionary.
    """
    data = []
    for row in t.rows:
        row_data = [cell.text for cell in row.cells]
        data.append(row_data)

    properties = {}
    if t.style and t.style.name and t.style.name != 'Table Grid':
        properties['style'] = t.style.name

    return {
        "type": "table",
        "data": data,
        "properties": properties
    }


def parse_docx_to_json(doc_bytes: bytes) -> Dict[str, Any]:
    """
    【已重构 v3】解析上传的 .docx 文件内容为我们的标准 DocumentModel JSON 格式。
    此版本集成了自定义的 iter_block_items 迭代器，修复了 AttributeError。

    Args:
        doc_bytes (bytes): .docx 文件的字节内容。

    Returns:
        Dict[str, Any]: 一个符合 DocumentModel schema 的字典。
    """
    doc = Document(io.BytesIO(doc_bytes))

    doc_state: Dict[str, Any] = {
        "page_setup": {},
        "numbering_definitions": [],
        "sections": [{"elements": []}]
    }

    current_elements = doc_state["sections"][0]["elements"]

    # Use our custom iterator
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            if block.text.strip():
                current_elements.append(_parse_paragraph(block))
        elif isinstance(block, Table):
            current_elements.append(_parse_table(block))

    return doc_state