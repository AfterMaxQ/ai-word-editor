# src/doc_generator.py

import json
import os
import shutil
import tempfile
import uuid
import zipfile

import io
from lxml import etree
import re
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT, WD_SECTION_START

from docx.shared import RGBColor
from .ai_parser import translate_latex_to_omml_llm
from .latex_converter import latex_to_omml
from docx.enum.text import WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import xml.etree.ElementTree as ET



# 定义一个从字符串到docx枚举的映射字典
ALIGNMENT_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
}

def load_document_data(filepath):
    """
        从指定的JSON文件中读取文档结构数据。
        Args:
            filepath (str): JSON文件的路径。

        Returns:
            dict: 包含文档结构数据的字典。
                  如果文件不存在或格式错误，则程序会打印错误信息并退出。
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data
    except FileNotFoundError:
        print(f"错误：文件未找到")
    except json.JSONDecodeError:
        print(f"错误：JSON文件格式不正确 -> {filepath}")
        sys.exit(1)


def apply_paragraph_properties(paragraph, properties: dict):
    """
    将properties字典中定义的格式应用到段落对象上。
    - 新增：应用段前/段后间距和行距。
    """
    p_format = paragraph.paragraph_format

    # 设置段前间距
    spacing_before_pt = properties.get('spacing_before')
    if spacing_before_pt is not None:
        p_format.space_before = Pt(spacing_before_pt)

    # 设置段后间距
    spacing_after_pt = properties.get('spacing_after')
    if spacing_after_pt is not None:
        p_format.space_after = Pt(spacing_after_pt)

    # 设置行距
    line_spacing_val = properties.get('line_spacing')
    if line_spacing_val is not None:
        p_format.line_spacing = line_spacing_val

    # 设置首行缩进
    indent_cm = properties.get('first_line_indent')
    if indent_cm is not None:
        p_format.first_line_indent = Cm(indent_cm)

    # --- 设置字体格式 (字体、字号、颜色等), 字体格式需要应用到段落内的Run上。---
    if not paragraph.runs:
        return

    for run in paragraph.runs:
        font = run.font

        font_name = properties.get('font_name', '宋体')

        if font_name:
            font.name = font_name
            from docx.oxml.ns import qn
            font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

        font_size_pt = properties.get('font_size')
        if font_size_pt is not None:
            font.size = Pt(font_size_pt)

        is_bold = properties.get('bold')
        if is_bold is not None:
            font.bold = bool(is_bold)

        font_color_hex = properties.get('font_color')
        if font_color_hex:
            try:
                font.color.rgb = RGBColor.from_string(font_color_hex)
            except ValueError:
                print(f"警告：无效的颜色格式码 '{font_color_hex}'，已跳过颜色设置。")

def add_table_from_data(doc, element: dict):
    """
        根据element字典中的数据，在文档中添加一个表格。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含表格数据的字典。
    """
    properties = element.get("properties", {})
    table_data = element.get('data', [])

    # 1. 验证数据完整性
    if not table_data or not table_data[0]:
        print("警告：表格数据为空或格式不正确，跳过此表格。")
        return

    # 从数据中推断行数和列数，这比从properties中读取更可靠
    cols = len(table_data[0])

    # 2. 创建表格，初始只有一行（用于表头）
    table = doc.add_table(rows=1, cols=cols)
    table.style = 'Table Grid'

    # 3. 填充表头行
    header_cells = table.rows[0].cells
    for j in range(cols):
        header_cells[j].text = str(table_data[0][j])
        # 如果是表头，则加粗
        if properties.get('header'):
            header_cells[j].paragraphs[0].runs[0].font.bold = True

    # 4. 动态添加并填充数据行
    for i in range(1, len(table_data)):
        row_cells = table.add_row().cells
        for j in range(cols):
            row_cells[j].text = str(table_data[i][j])

    # 5. 应用列对齐 (在所有行都添加后进行)
    alignments = properties.get("alignments", [])
    if alignments:
        for col_idx, align_str in enumerate(alignments):
            if col_idx < cols:
                alignment_enum = ALIGNMENT_MAP.get(align_str.lower())
                if alignment_enum:
                    for row in table.rows:
                        row.cells[col_idx].paragraphs[0].paragraph_format.alignment = alignment_enum

def add_list_from_data(doc, element: dict):
    """
        根据element字典中的数据，在文档中添加一个有序或无序列表。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含列表数据的字典。
    """
    properties = element.get('properties', {})
    items = element.get('items', [])
    # 确保items是一个非空列表
    if not isinstance(items, list) or not items:
        print("警告：列表数据为空或格式不正确，跳过此列表。")
        return
    # 根据 'ordered' 属性决定使用哪种段落样式
    is_ordered = properties.get('ordered', False)
    style = "List Number" if is_ordered else 'List Bullet'
    # 遍历所有列表项，并使用指定的样式添加到文档中
    for item_text in items:
        doc.add_paragraph(str(item_text), style=style)

def add_image_from_data(doc, element: dict):
    """
        根据element字典中的数据，在文档中添加一张图片。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含图片数据的字典。
    """
    properties = element.get('properties', {})
    path = properties.get('path')

    if not path:
        print("警告：图片元素缺少'path'属性，跳过此图片。")
        return
    # 从properties中获取宽度和高度
    width_cm = properties.get('width')
    height_cm = properties.get('height')
    # 将Python的None或数值转换为docx的Cm单位对象
    width = Cm(width_cm) if width_cm is not None else None
    height = Cm(height_cm) if height_cm is not None else None

    try:
        doc.add_picture(path, width=width, height=height)
    except FileNotFoundError:
        print(f"警告：图片文件未找到 -> {path} ，跳过此图片。")
    except Exception as e:
        # 捕获其他可能的错误，如文件格式不支持等
        print(f"警告：插入图片时发生错误 -> {path} ({e})，跳过此图片。")

def add_page_number(paragraph):
    """
        在一个段落中添加页码域。
        这部分代码比较底层，直接操作Word文档的XML结构。
    """
    # 创建一个 run (可以理解为段落中的一小段格式统一的文本)
    run = paragraph.add_run()

    # --- 开始添加复杂的页码域 ---
    # 1. 创建 fldChar 并设置其类型为 'begin'
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')

    # 2. 创建 instrText，这是页码的指令
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'  # PAGE 表示当前页码

    # 3. 创建 fldChar 并设置其类型为 'end'
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    # 将这三个元素按顺序添加到 run 中
    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_end)

def add_header_from_data(doc, element: dict, section):
    """
        根据element字典中的数据，在文档中添加页眉。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含页眉数据的字典。
    """
    properties = element.get("properties", {})
    text = properties.get("text", "")

    header = section.header

    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    paragraph.clear()

    #设置对齐
    align_str = properties.get('alignment', 'center')
    paragraph.alignment = ALIGNMENT_MAP.get(align_str.lower(), WD_ALIGN_PARAGRAPH.CENTER)

    if '{PAGE_NUM}' in text:
        parts = text.split('{PAGE_NUM}')
        paragraph.add_run(parts[0])
        add_page_number(paragraph)
        paragraph.add_run(parts[1])
    else:
        paragraph.add_run(text)


def add_footer_from_data(doc, element: dict, section):
    """
    根据element字典中的数据，在文档中添加页脚。

    Args:
        doc: The python-docx Document object.
        element (dict): 包含页脚数据的字典。
    """
    properties = element.get("properties", {})
    text = properties.get("text", "")

    footer = section.footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.clear()

    # 设置对齐
    align_str = properties.get('alignment', 'center')
    paragraph.alignment = ALIGNMENT_MAP.get(align_str.lower(), WD_ALIGN_PARAGRAPH.CENTER)

    # 添加内容，并处理页码占位符
    if '{PAGE_NUM}' in text:
        parts = text.split('{PAGE_NUM}')
        paragraph.add_run(parts[0])
        add_page_number(paragraph)
        paragraph.add_run(parts[1])
    else:
        paragraph.add_run(text)


def post_process_footnotes(docx_bytes, footnotes_to_inject, reference_format: str = "#"):
    """
    通过解压、操作XML、再重新打包的方式，为DOCX文件添加脚注。
    此函数现在会根据提供的格式字符串（如 "[#]"）创建引用标记。

    Args:
        docx_bytes (bytes): 包含占位符的原始DOCX文件字节流。
        footnotes_to_inject (dict): 一个字典，键是占位符，值是脚注文本。
        reference_format (str): 引用标记的格式，'#' 是数字的占位符。
    """
    if not footnotes_to_inject:
        return docx_bytes

    print("\n--- 开始脚注XML后处理阶段 ---")
    print(f"  > 使用引用格式: '{reference_format}'")

    # ... (ns 定义和 temp_dir 创建保持不变) ...
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'content_types': 'http://schemas.openxmlformats.org/package/2006/content-types'
    }
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)

    temp_dir = tempfile.mkdtemp()
    try:
        # ... (解压和文件路径定义保持不变) ...
        with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 定义文件路径
        document_path = os.path.join(temp_dir, 'word', 'document.xml')
        footnotes_path = os.path.join(temp_dir, 'word', 'footnotes.xml')
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')

        # ... (处理 footnotes.xml 和读取 document.xml 的逻辑保持不变) ...
        footnotes_tree = None
        footnotes_root = None
        if os.path.exists(footnotes_path):
            footnotes_tree = ET.parse(footnotes_path)
            footnotes_root = footnotes_tree.getroot()
        else:
            footnotes_root = ET.Element(f'{{{ns["w"]}}}footnotes')
            # 添加Word需要的分隔符脚注
            ET.SubElement(footnotes_root, f'{{{ns["w"]}}}footnote',
                          {f'{{{ns["w"]}}}id': '-1', f'{{{ns["w"]}}}type': 'separator'})
            ET.SubElement(footnotes_root, f'{{{ns["w"]}}}footnote',
                          {f'{{{ns["w"]}}}id': '0', f'{{{ns["w"]}}}type': 'continuationSeparator'})
            footnotes_tree = ET.ElementTree(footnotes_root)

        # 2. 读取 document.xml 内容
        doc_xml_content = ""
        with open(document_path, 'r', encoding='utf-8') as f:
            doc_xml_content = f.read()

        # 3. 遍历需要注入的脚注，替换占位符并构建脚注内容
        for placeholder, footnote_text in footnotes_to_inject.items():
            # ... (确定 new_footnote_id 和构建脚注XML内容的逻辑保持不变) ...
            existing_ids = [int(fn.get(f'{{{ns["w"]}}}id')) for fn in footnotes_root.findall('w:footnote', ns) if
                            int(fn.get(f'{{{ns["w"]}}}id')) > 0]
            new_footnote_id = (max(existing_ids) + 1) if existing_ids else 1

            # 构建脚注XML内容
            footnote_element = ET.SubElement(footnotes_root, f'{{{ns["w"]}}}footnote',
                                             {f'{{{ns["w"]}}}id': str(new_footnote_id)})
            p_element = ET.SubElement(footnote_element, f'{{{ns["w"]}}}p')
            pPr = ET.SubElement(p_element, f'{{{ns["w"]}}}pPr')
            ET.SubElement(pPr, f'{{{ns["w"]}}}pStyle', {f'{{{ns["w"]}}}val': 'FootnoteText'})
            r_ref = ET.SubElement(p_element, f'{{{ns["w"]}}}r')
            rPr_ref = ET.SubElement(r_ref, f'{{{ns["w"]}}}rPr')
            ET.SubElement(rPr_ref, f'{{{ns["w"]}}}rStyle', {f'{{{ns["w"]}}}val': 'FootnoteReference'})
            ET.SubElement(r_ref, f'{{{ns["w"]}}}footnoteRef')
            r_text = ET.SubElement(p_element, f'{{{ns["w"]}}}r')
            t_text = ET.SubElement(r_text, f'{{{ns["w"]}}}t')
            t_text.text = f" {footnote_text}"

            parts = reference_format.split('#')
            prefix = parts[0]
            suffix = parts[1] if len(parts) > 1 else ""

            # 创建一个函数来生成带上标样式的 run
            def create_styled_run(text_content=None):
                run = ET.Element(f'{{{ns["w"]}}}r')
                rpr = ET.SubElement(run, f'{{{ns["w"]}}}rPr')
                ET.SubElement(rpr, f'{{{ns["w"]}}}vertAlign', {f'{{{ns["w"]}}}val': 'superscript'})
                if text_content:
                    t = ET.SubElement(run, f'{{{ns["w"]}}}t')
                    t.text = text_content
                return run

            xml_parts_to_combine = []
            if prefix:
                xml_parts_to_combine.append(create_styled_run(prefix))

            # 创建包含脚注引用的 run
            ref_run = create_styled_run()
            ET.SubElement(ref_run, f'{{{ns["w"]}}}footnoteReference', {f'{{{ns["w"]}}}id': str(new_footnote_id)})
            xml_parts_to_combine.append(ref_run)

            if suffix:
                xml_parts_to_combine.append(create_styled_run(suffix))

            # 将所有 XML 部分转换为字符串并连接
            combined_xml_str = "".join([ET.tostring(part, encoding='unicode') for part in xml_parts_to_combine])
            # ▲▲▲ 修改结束 ▲▲▲

            # 在 document.xml 内容中替换占位符
            if placeholder in doc_xml_content:
                doc_xml_content = doc_xml_content.replace(placeholder, combined_xml_str)
                print(f"  > 找到并替换脚注占位符: '{placeholder}'")
            else:
                print(f"警告: 在 document.xml 中未找到占位符 '{placeholder}'。")

        footnotes_tree.write(footnotes_path, encoding='UTF-8', xml_declaration=True)
        with open(document_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml_content)

        # 5. 更新关系文件 (rels)
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        if not any(rel.get('Target') == 'footnotes.xml' for rel in rels_root):
            existing_rids = [int(r.get('Id')[3:]) for r in rels_root]
            new_rid = f"rId{max(existing_rids) + 1 if existing_rids else 1}"
            ET.SubElement(rels_root, f'{{{ns["rel"]}}}Relationship', {
                'Id': new_rid,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
                'Target': 'footnotes.xml'
            })
            rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True)

        # 6. 更新 [Content_Types].xml
        types_tree = ET.parse(content_types_path)
        types_root = types_tree.getroot()
        if not any(
                ov.get('PartName') == '/word/footnotes.xml' for ov in types_root.findall('content_types:Override', ns)):
            ET.SubElement(types_root, f'{{{ns["content_types"]}}}Override', {
                'PartName': '/word/footnotes.xml',
                'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
            })
            types_tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)

        # 7. 将临时目录重新打包成字节流
        output_stream = io.BytesIO()
        with zipfile.ZipFile(output_stream, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, archive_name)

        print("--- 脚注XML后处理阶段完成 ---")
        return output_stream.getvalue()

    finally:
        shutil.rmtree(temp_dir)


def post_process_endnotes(docx_bytes, endnotes_to_inject, reference_format: str = "#"):
    """
    通过解压、操作XML、再重新打包的方式，为DOCX文件添加尾注。
    此函数现在会根据提供的格式字符串（如 "[#]"）创建引用标记。

    Args:
        docx_bytes (bytes): 包含占位符的原始DOCX文件字节流。
        endnotes_to_inject (dict): 一个字典，键是占位符，值是尾注文本。
        reference_format (str): 引用标记的格式，'#' 是数字的占位符。
    """
    if not endnotes_to_inject:
        return docx_bytes

    print("\n--- 开始尾注XML后处理阶段 ---")
    print(f"  > 使用引用格式: '{reference_format}'")

    # ... (ns 定义和 temp_dir 创建保持不变) ...
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'content_types': 'http://schemas.openxmlformats.org/package/2006/content-types'
    }
    for prefix, uri in ns.items():
        ET.register_namespace(prefix, uri)

    temp_dir = tempfile.mkdtemp()
    try:
        # ... (解压和文件路径定义保持不变) ...
        with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 定义所需文件的路径
        document_path = os.path.join(temp_dir, 'word', 'document.xml')
        endnotes_path = os.path.join(temp_dir, 'word', 'endnotes.xml')  # 尾注XML文件
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')

        endnotes_tree = None
        endnotes_root = None
        if os.path.exists(endnotes_path):
            endnotes_tree = ET.parse(endnotes_path)
            endnotes_root = endnotes_tree.getroot()
        else:
            # 如果文件不存在，则创建一个新的XML结构
            endnotes_root = ET.Element(f'{{{ns["w"]}}}endnotes')
            # 添加Word需要的分隔符尾注
            ET.SubElement(endnotes_root, f'{{{ns["w"]}}}endnote',
                          {f'{{{ns["w"]}}}id': '-1', f'{{{ns["w"]}}}type': 'separator'})
            ET.SubElement(endnotes_root, f'{{{ns["w"]}}}endnote',
                          {f'{{{ns["w"]}}}id': '0', f'{{{ns["w"]}}}type': 'continuationSeparator'})
            endnotes_tree = ET.ElementTree(endnotes_root)

        # 2. 读取主文档XML内容
        doc_xml_content = ""
        with open(document_path, 'r', encoding='utf-8') as f:
            doc_xml_content = f.read()

        # 3. 遍历需要注入的尾注，构建XML并替换占位符
        for placeholder, endnote_text in endnotes_to_inject.items():
            # ... (确定 new_endnote_id 和构建尾注XML内容的逻辑保持不变) ...
            existing_ids = [int(en.get(f'{{{ns["w"]}}}id')) for en in endnotes_root.findall('w:endnote', ns) if
                            int(en.get(f'{{{ns["w"]}}}id')) > 0]
            new_endnote_id = (max(existing_ids) + 1) if existing_ids else 1

            # 构建<w:endnote> XML元素
            endnote_element = ET.SubElement(endnotes_root, f'{{{ns["w"]}}}endnote',
                                            {f'{{{ns["w"]}}}id': str(new_endnote_id)})
            p_element = ET.SubElement(endnote_element, f'{{{ns["w"]}}}p')
            pPr = ET.SubElement(p_element, f'{{{ns["w"]}}}pPr')
            ET.SubElement(pPr, f'{{{ns["w"]}}}pStyle', {f'{{{ns["w"]}}}val': 'EndnoteText'})  # 尾注文本样式
            r_ref = ET.SubElement(p_element, f'{{{ns["w"]}}}r')
            rPr_ref = ET.SubElement(r_ref, f'{{{ns["w"]}}}rPr')
            ET.SubElement(rPr_ref, f'{{{ns["w"]}}}rStyle', {f'{{{ns["w"]}}}val': 'EndnoteReference'})  # 尾注引用样式
            ET.SubElement(r_ref, f'{{{ns["w"]}}}endnoteRef')  # 尾注引用标记
            r_text = ET.SubElement(p_element, f'{{{ns["w"]}}}r')
            t_text = ET.SubElement(r_text, f'{{{ns["w"]}}}t')
            t_text.text = f" {endnote_text}"

            parts = reference_format.split('#')
            prefix = parts[0]
            suffix = parts[1] if len(parts) > 1 else ""

            # 创建一个函数来生成带上标样式的 run
            def create_styled_run(text_content=None):
                run = ET.Element(f'{{{ns["w"]}}}r')
                rpr = ET.SubElement(run, f'{{{ns["w"]}}}rPr')
                ET.SubElement(rpr, f'{{{ns["w"]}}}vertAlign', {f'{{{ns["w"]}}}val': 'superscript'})
                if text_content:
                    t = ET.SubElement(run, f'{{{ns["w"]}}}t')
                    t.text = text_content
                return run

            xml_parts_to_combine = []
            if prefix:
                xml_parts_to_combine.append(create_styled_run(prefix))

            # 创建包含尾注引用的 run
            ref_run = create_styled_run()
            ET.SubElement(ref_run, f'{{{ns["w"]}}}endnoteReference', {f'{{{ns["w"]}}}id': str(new_endnote_id)})
            xml_parts_to_combine.append(ref_run)

            if suffix:
                xml_parts_to_combine.append(create_styled_run(suffix))

            # 将所有 XML 部分转换为字符串并连接
            combined_xml_str = "".join([ET.tostring(part, encoding='unicode') for part in xml_parts_to_combine])

            # 在主文档XML中进行替换
            if placeholder in doc_xml_content:
                # 使用新的带上标属性的 XML 字符串
                doc_xml_content = doc_xml_content.replace(placeholder, combined_xml_str)
                print(f"  > 找到并替换尾注占位符: '{placeholder}'")
            else:
                print(f"警告: 在 document.xml 中未找到尾注占位符 '{placeholder}'。")

        # ... (写回文件、更新rels和打包的逻辑保持不变) ...
        endnotes_tree.write(endnotes_path, encoding='UTF-8', xml_declaration=True)
        with open(document_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml_content)

        # 5. 更新关系文件 (document.xml.rels)
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        if not any(rel.get('Target') == 'endnotes.xml' for rel in rels_root):
            existing_rids = [int(r.get('Id')[3:]) for r in rels_root]
            new_rid = f"rId{max(existing_rids) + 1 if existing_rids else 1}"
            ET.SubElement(rels_root, f'{{{ns["rel"]}}}Relationship', {
                'Id': new_rid,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',  # 尾注关系类型
                'Target': 'endnotes.xml'
            })
            rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True)

        # 6. 更新内容类型文件 ([Content_Types].xml)
        types_tree = ET.parse(content_types_path)
        types_root = types_tree.getroot()
        if not any(
                ov.get('PartName') == '/word/endnotes.xml' for ov in types_root.findall('content_types:Override', ns)):
            ET.SubElement(types_root, f'{{{ns["content_types"]}}}Override', {
                'PartName': '/word/endnotes.xml',
                'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'  # 尾注内容类型
            })
            types_tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)

        # 7. 重新打包成DOCX字节流
        output_stream = io.BytesIO()
        with zipfile.ZipFile(output_stream, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, archive_name)

        print("--- 尾注XML后处理阶段完成 ---")
        return output_stream.getvalue()

    finally:
        shutil.rmtree(temp_dir)


_bookmark_id_counter = 0


def add_bookmark(paragraph, bookmark_name: str):
    """
    为一个段落的完整内容添加书签。
    这需要直接操作OXML来插入<w:bookmarkStart>和<w:bookmarkEnd>标签。
    """
    global _bookmark_id_counter
    # 创建书签的起始标签
    run = paragraph.runs[0]  # 通常在段落的第一个run之前插入
    bookmark_start = OxmlElement('w:bookmarkStart')
    # Word需要一个数字ID和一个名字。我们用计数器生成唯一ID。
    bookmark_start.set(qn('w:id'), str(_bookmark_id_counter))
    bookmark_start.set(qn('w:name'), bookmark_name)
    # 将起始标签插入到段落的XML中
    run._r.addprevious(bookmark_start)

    # 创建书签的结束标签
    run = paragraph.runs[-1]  # 在段落的最后一个run之后插入
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(_bookmark_id_counter))
    # 将结束标签插入到段落的XML中
    run._r.addnext(bookmark_end)

    _bookmark_id_counter += 1


def add_cross_reference_field(paragraph, bookmark_name: str, display_text: str):
    """
    在段落中插入一个带有预显示文本的交叉引用域 (REF field)。
    """
    # 域代码由多个部分组成，每个部分都放在自己的 run 中，这是最稳妥的做法

    # 1. 'begin' 标记：表示一个域的开始
    run_begin = paragraph.add_run()
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    run_begin._r.append(fldChar_begin)

    # 2. 指令文本：告诉Word具体做什么
    run_instr = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = f' REF {bookmark_name} \\h '  # REF <书签名> \h
    run_instr._r.append(instrText)

    # 3. 'separate' 标记：分隔指令和结果
    run_sep = paragraph.add_run()
    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    run_sep._r.append(fldChar_separate)

    # 4. 【关键修复】缓存的显示结果
    # Word 打开文档时会直接显示这里的文本
    run_display = paragraph.add_run()
    run_display.text = display_text

    # 5. 'end' 标记：表示一个域的结束
    run_end = paragraph.add_run()
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run_end._r.append(fldChar_end)


def apply_numbering_formats(docx_bytes, page_setup_data):
    """
    通过修改 settings.xml 来应用脚注和尾注的数字格式。
    """
    endnote_fmt = page_setup_data.get("endnote_number_format")
    footnote_fmt = page_setup_data.get("footnote_number_format")

    if not endnote_fmt and not footnote_fmt:
        return docx_bytes

    print("\n--- 开始应用自定义引用编号格式 ---")

    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    ET.register_namespace('w', ns['w'])

    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')

        if not os.path.exists(settings_path):
            print("警告: settings.xml 未找到，跳过编号格式化。")
            return docx_bytes

        settings_tree = ET.parse(settings_path)
        settings_root = settings_tree.getroot()

        # 处理尾注设置
        if endnote_fmt:
            # 查找或创建 <w:endnotePr>
            endnote_pr = settings_root.find('w:endnotePr', ns)
            if endnote_pr is None:
                endnote_pr = ET.SubElement(settings_root, f'{{{ns["w"]}}}endnotePr')

            # 查找或创建 <w:numFmt> 并设置其值
            num_fmt = endnote_pr.find('w:numFmt', ns)
            if num_fmt is None:
                num_fmt = ET.SubElement(endnote_pr, f'{{{ns["w"]}}}numFmt')
            num_fmt.set(f'{{{ns["w"]}}}val', endnote_fmt)
            print(f"  > 已将尾注编号格式设置为: '{endnote_fmt}'")

        # (可选) 处理脚注设置，逻辑同上
        if footnote_fmt:
            footnote_pr = settings_root.find('w:footnotePr', ns)
            if footnote_pr is None:
                footnote_pr = ET.SubElement(settings_root, f'{{{ns["w"]}}}footnotePr')
            num_fmt = footnote_pr.find('w:numFmt', ns)
            if num_fmt is None:
                num_fmt = ET.SubElement(footnote_pr, f'{{{ns["w"]}}}numFmt')
            num_fmt.set(f'{{{ns["w"]}}}val', footnote_fmt)
            print(f"  > 已将脚注编号格式设置为: '{footnote_fmt}'")

        settings_tree.write(settings_path, encoding='UTF-8', xml_declaration=True)

        # 重新打包
        output_stream = io.BytesIO()
        with zipfile.ZipFile(output_stream, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, archive_name)

        print("--- 自定义引用编号格式应用完毕 ---")
        return output_stream.getvalue()

    finally:
        shutil.rmtree(temp_dir)

def apply_page_setup(doc, page_setup_data: dict):
    """
        根据 page_setup_data 字典中的数据，应用页面布局设置。

        Args:
            doc: The python-docx Document object.
            page_setup_data (dict): 包含页面设置（方向、边距）的字典。
    """
    if not page_setup_data:
        return

    # 在Word文档中，页面布局是“节(Section)”的属性。
    # 对于一个新文档，默认只有一个节。
    section = doc.sections[0]

    #设置页面方向
    orientation_str = page_setup_data.get('orientation')
    if orientation_str == 'landscape':
        section.orientation = WD_ORIENT.LANDSCAPE
        # 手动交换页面宽高，确保设置生效
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height

    # 设置页边距
    margins = page_setup_data.get('margins')
    if isinstance(margins, dict):
        # --- 这是修复后的代码 ---
        # 我们现在不仅检查键是否存在，还检查它的值是否为 None

        top_cm = margins.get('top')
        if top_cm is not None:
            section.top_margin = Cm(top_cm)

        bottom_cm = margins.get('bottom')
        if bottom_cm is not None:
            section.bottom_margin = Cm(bottom_cm)

        left_cm = margins.get('left')
        if left_cm is not None:
            section.left_margin = Cm(left_cm)

        right_cm = margins.get('right')
        if right_cm is not None:
            section.right_margin = Cm(right_cm)

def apply_section_properties(section, section_data: dict):
    """
    根据 section_data 字典中的数据，应用节的属性（如分栏）。
    """
    properties = section_data.get('properties', {})
    if not properties:
        return

    sectPr = section._sectPr

    existing_cols = sectPr.find(qn('w:cols'))
    if existing_cols is not None:
        sectPr.remove(existing_cols)
    # 设置分栏
    columns = properties.get('columns')
    if columns and columns > 1:
        # 创建 <w:cols> 元素
        cols = OxmlElement('w:cols')
        cols.set(qn('w:num'), str(columns))
        # 将 <w:cols> 添加到 <w:sectPr> 中
        sectPr.append(cols)

def add_page_break_from_data(doc, element: dict):
    """
    在文档中添加一个分页符。

    Args:
        doc: The python-docx Document object.
        element (dict): 包含分页符数据的字典 (虽然此元素为空)。
    """
    doc.add_page_break()

def add_column_break_from_data(doc, element: dict):
    """
    在文档中添加一个分栏符。
    """
    # 分栏符必须添加到一个段落中
    if not doc.paragraphs:
        # 极端情况：如果文档中还没有任何段落，则创建一个
        # （这在实际应用中几乎不会发生，因为总会有内容的）
        p = doc.add_paragraph()
        p.add_run().add_break(WD_BREAK.COLUMN)
        return

        # 获取文档中的最后一个段落
    last_paragraph = doc.paragraphs[-1]
    # 在该段落的末尾添加一个新的run，并在其中插入分栏符
    last_paragraph.add_run().add_break(WD_BREAK.COLUMN)


def add_toc_from_data(doc, element: dict):
    """
    在文档中添加一个目录（TOC）。
    这需要直接操作底层的OXML来创建一个复杂的域。
    """
    properties = element.get("properties", {})

    title = properties.get('title')
    if title:
        doc.add_paragraph(title, style='Heading 1')

    paragraph = doc.add_paragraph()
    run = paragraph.add_run()

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = r'TOC \o "1-3" \h \z \u'

    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')

    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)

def get_formula_xml_and_placeholder(element: dict) -> tuple[str | None, str | None]:
    """
        1. 尝试本地转换器
        2. 若失败，回退到 LLM
        3. 对所有输出进行零信任验证
        4. 若全部失败，生成一个格式化的错误信息
    """
    latex_text = element.get("properties", {}).get("text")
    if not latex_text: return None, None

    placeholder = f"__FORMULA_{uuid.uuid4().hex}__"
    omml_str = None

    # --- 防御层 1: 本地转换器 ---
    print(f"⚙️ 尝试使用本地解析器处理: {latex_text}")
    try:
        omml_elem = latex_to_omml(latex_text)
        if omml_elem is not None:
            # 零信任验证: 确保生成的XML是有效的
            temp_str = etree.tostring(omml_elem, encoding='unicode')
            etree.fromstring(temp_str)  # 重新解析以确保有效性
            if omml_elem.tag == qn('m:oMathPara'):
                print("✅ 本地解析器成功并验证通过。")
                omml_str = temp_str
    except Exception as e:
        print(f"⚠️ 本地解析器失败: {e}。回退到 LLM...")
        omml_str = None

    # --- 防御层 2: LLM 兜底 ---
    if not omml_str:
        llm_xml = translate_latex_to_omml_llm(latex_text)
        if llm_xml:
            try:
                # 零信任验证
                root = etree.fromstring(llm_xml)
                if root.tag == qn('m:oMathPara') and root.find(qn('m:oMath')) is not None:
                    print("✅ LLM成功并验证通过。")
                    omml_str = llm_xml
                else:
                    print("❌ LLM返回的XML结构不符合规范 (缺少<m:oMathPara>或<m:oMath>)。")
            except etree.XMLSyntaxError as e:
                print(f"❌ LLM返回的不是有效的XML: {e}")

    # --- 防御层 3: 最终错误提示 ---
    if omml_str:
        return placeholder, omml_str
    else:
        print(f"❌ 所有转换方法均失败，生成错误提示: '{latex_text}'")
        # 创建一个有效的 <m:oMathPara> XML元素，其中包含红色错误文本
        # 这确保了即使转换失败，注入的XML也是有效的，不会损坏文档
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        color = OxmlElement('w:color');
        color.set(qn('w:val'), 'FF0000')  # 红色
        rPr.append(color)
        r.append(rPr)
        t = OxmlElement('w:t');
        t.text = f"[公式渲染失败: '{latex_text}']"
        r.append(t)
        p.append(r)

        return placeholder, etree.tostring(p, encoding='unicode')


def create_document(data: dict) -> tuple[bytes | None, str | None]:
    """
    根据以“节”为单位的、经过验证的JSON数据创建Word文档。
    """
    doc = Document()
    formulas_to_inject = {}
    footnotes_to_inject = {}
    endnotes_to_inject = {}

    page_setup_data = data.get('page_setup', {})
    if page_setup_data:
        apply_page_setup(doc, page_setup_data)

    sections_data = data.get('sections', [])

    for i, section_data in enumerate(sections_data):
        if i == 0:
            section = doc.sections[0]
        else:
            section = doc.add_section(WD_SECTION_START.CONTINUOUS)

        apply_section_properties(section, section_data)

        elements = section_data.get('elements', [])

        # 第一步：预扫描，构建书签->文本的映射
        bookmark_map = {}
        print("\n--- 开始预扫描本节的书签 ---")
        for element in elements:
            if element.get('type') == 'paragraph':
                properties = element.get('properties', {})
                bookmark_id = properties.get('bookmark_id')
                if bookmark_id:
                    # 将书签ID和段落的纯文本内容关联起来
                    text_content = element.get('text', '')
                    bookmark_map[bookmark_id] = text_content
                    print(f"  > 发现书签: '{bookmark_id}' -> '{text_content}'")
        print("--- 预扫描结束 ---\n")
        # 预扫描结束
        #  第二步：正式生成内容
        for element in elements:
            element_type = element.get('type')

            if element_type == "header":
                add_header_from_data(doc, element, section)
            elif element_type == "footer":
                add_footer_from_data(doc, element, section)
            elif element_type == "formula":
                placeholder, omml_xml = get_formula_xml_and_placeholder(element)
                if placeholder and omml_xml:
                    p = doc.add_paragraph()
                    p.add_run(placeholder)
                    formulas_to_inject[placeholder] = omml_xml
            elif element_type == 'paragraph':
                properties = element.get('properties', {})
                style = properties.get("style") if isinstance(properties, dict) else None

                p = doc.add_paragraph(style=style)

                content = element.get('content')
                if content and isinstance(content, list):
                    for run_item in content:
                        run_type = run_item.get('type')
                        run_text = run_item.get('text', '')
                        if run_type == 'text':
                            p.add_run(run_text)
                        elif run_type == 'footnote':
                            placeholder = f"__FOOTNOTE_{uuid.uuid4().hex}__"
                            p.add_run(placeholder)
                            footnotes_to_inject[placeholder] = run_text
                        elif run_type == 'endnote':
                            placeholder = f"__ENDNOTE_{uuid.uuid4().hex}__"
                            p.add_run(placeholder)
                            endnotes_to_inject[placeholder] = run_text
                        elif run_type == 'cross_reference':
                            target_bookmark = run_item.get('target_bookmark')
                            if target_bookmark:
                                display_text = bookmark_map.get(target_bookmark, "[错误: 找不到引用]")
                                add_cross_reference_field(p, target_bookmark, display_text)
                                print(f"  > 成功插入对书签 '{target_bookmark}' 的交叉引用。")
                else:
                    text = element.get('text', '')
                    p.add_run(text)

                if isinstance(properties, dict):
                    apply_paragraph_properties(p, properties)

                bookmark_name = properties.get("bookmark_id")
                if bookmark_name and p.runs:
                    add_bookmark(p, bookmark_name)
                    print(f"> 成功为段落创建书签: '{bookmark_name}")

            elif element_type == "table":
                add_table_from_data(doc, element)
            elif element_type == "list":
                add_list_from_data(doc, element)
            elif element_type == "image":
                add_image_from_data(doc, element)
            elif element_type == "page_break":
                add_page_break_from_data(doc, element)
            elif element_type == "column_break":
                add_column_break_from_data(doc, element)
            elif element_type == "toc":
                add_toc_from_data(doc, element)

    # Step 1: Always save the initial document with placeholders to a byte stream.
    # This stream will be the input for our post-processing steps.
    draft_stream = io.BytesIO()
    doc.save(draft_stream)
    draft_stream.seek(0)
    processed_bytes = draft_stream.getvalue()
    final_xml_body_for_log = "" # We'll populate this for logging purposes.

    # Step 2: Process formulas if they exist.
    # This is a complex operation that unzips, modifies, and re-zips the document.
    if formulas_to_inject:
        print("\n--- 开始XML后处理阶段 (公式) ---")
        final_stream = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(processed_bytes), 'r') as zin, zipfile.ZipFile(final_stream, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    root = etree.fromstring(content)
                    body = root.find(qn('w:body'))
                    paragraphs = body.findall('.//' + qn('w:p'))
                    for p in paragraphs:
                        placeholder_text = p.xpath("string(.)").strip()
                        if placeholder_text in formulas_to_inject:
                            real_xml_str = formulas_to_inject[placeholder_text]
                            try:
                                new_content_element = etree.fromstring(real_xml_str)
                                for child in list(p):
                                    p.remove(child)
                                if new_content_element.tag == qn('m:oMathPara'):
                                    para_props = new_content_element.find(qn('m:oMathParaPr'))
                                    if para_props is not None:
                                        pPr = p.find(qn('w:pPr'))
                                        if pPr is None:
                                            pPr = OxmlElement('w:pPr')
                                            p.insert(0, pPr)
                                        jc = para_props.find(qn('m:jc'))
                                        if jc is not None:
                                            new_jc = OxmlElement('w:jc')
                                            new_jc.set(qn('w:val'), jc.get(qn('m:val')))
                                            pPr.append(new_jc)
                                    oMath = new_content_element.find(qn('m:oMath'))
                                    if oMath is not None:
                                        p.append(oMath)
                                else:
                                    for run in new_content_element.findall('.//' + qn('w:r')):
                                        p.append(run)
                            except etree.XMLSyntaxError as e:
                                print(f"  ❌ 严重错误: 无法解析待注入的XML，跳过: {e}")
                                continue
                    final_xml_content = etree.tostring(root, encoding='utf-8', xml_declaration=True, standalone=True)
                    final_xml_body_for_log = etree.tostring(root, pretty_print=True, encoding='unicode')
                    zout.writestr(item, final_xml_content)
                else:
                    zout.writestr(item, content)
        final_stream.seek(0)
        processed_bytes = final_stream.getvalue() # Update the bytes with the formula-processed version.
        print("--- 公式XML后处理阶段完成 ---\n")

    # ▼▼▼【关键修改】读取自定义格式并传递给后处理函数 ▼▼▼
    page_setup_data_for_refs = data.get('page_setup', {})
    footnote_ref_format = page_setup_data_for_refs.get('footnote_reference_format', "[#]")
    endnote_ref_format = page_setup_data_for_refs.get('endnote_reference_format', "[#]")

    # Step 3: Process footnotes if they exist.
    # This runs on the output of the previous step (or the initial bytes if no formulas).
    if footnotes_to_inject:
        processed_bytes = post_process_footnotes(processed_bytes, footnotes_to_inject, footnote_ref_format)

    if endnotes_to_inject:
        processed_bytes = post_process_endnotes(processed_bytes, endnotes_to_inject, endnote_ref_format)

    if page_setup_data_for_refs:
        processed_bytes = apply_numbering_formats(processed_bytes, page_setup_data_for_refs)

    # 步骤 5: 应用自定义的脚注/尾注编号格式
    page_setup_data_for_numbering = data.get('page_setup', {})
    if page_setup_data_for_numbering:
        processed_bytes = apply_numbering_formats(processed_bytes, page_setup_data_for_numbering)

    # Step 4: If no formulas were processed, we still need to get the XML for logging.
    if not final_xml_body_for_log:
         try:
            # Note: This XML won't show the footnote changes, as they happen at the byte level.
            # It's primarily for diagnosing issues before post-processing.
            final_xml_body_for_log = etree.tostring(doc.element.body, pretty_print=True, encoding='unicode')
         except Exception as e:
            print(f"警告: 无法为日志生成XML: {e}")
            final_xml_body_for_log = "无法为日志生成XML。"

    # Step 5: Return the final, fully processed document bytes and the log.
    return processed_bytes, final_xml_body_for_log


