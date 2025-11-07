# src/doc_generator.py

import json
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT

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

        Args:
            paragraph: python-docx的段落对象。
            properties (dict): 包含格式定义的字典。
    """
    p_format = paragraph.paragraph_format

    # 设置首行缩进
    if 'first_line_indent' in properties:
        p_format.first_line_indent = Cm(properties['first_line_indent'])

    # --- 设置字体格式 (字体、字号等), 字体格式需要应用到段落内的Run上。---
    if paragraph.runs:
        font = paragraph.runs[0].font
        # 设置字体名称
        if 'font_name' in properties:
            font.name = properties['font_name']
            # 导入中文字体所需的包
            from docx.oxml.ns import qn
            # 设置中文字体 (东亚字体)
            font.element.rPr.rFonts.set(qn('w:eastAsia'), properties['font_name'])

        # 设置字体大小
        if "font_size" in properties:
            font.size = Pt(properties['font_size'])
        #设置粗体
        if 'bold' in properties:
            font.bold = bool(properties['bold'])


def add_table_from_data(doc, element: dict):
    """
        根据element字典中的数据，在文档中添加一个表格。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含表格数据的字典。
    """
    properties = element.get("properties",{})
    table_data = element.get('data',[])

    # 1. 获取表格尺寸和验证
    rows = properties.get('rows', 0)
    cols = properties.get('cols', 0)
    if rows==0 or cols==0 or not table_data:
        print("警告：表格数据不完整，跳过此表格。")
        return

    # 定义一个从字符串到docx枚举的映射字典
    ALIGNMENT_MAP = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
    }
    # 获取JSON中定义的对齐方式列表
    alignments = properties.get("alignments", [])

    # 2. 创建表格
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'

    # 3. 填充数据并设置格式
    for i in range(rows):
        for j in range(cols):
            # 获取单元格对象
            cell = table.cell(i, j)
            # 填充文本 (增加边界检查，防止数据行/列数与定义的rows/cols不匹配)
            if i < len(table_data) and j < len(table_data[i]):
                cell.text = str(table_data[i][j])
            # 4. 如果是表头行，则加粗
            if properties.get('header') and i == 0:
                #单元格内的第一个段落的第一个run的字体
                if cell.paragraphs and cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.bold = True

            if j < len(alignments):
                align_str = alignments[j]
                alignment_enum = ALIGNMENT_MAP.get(align_str.lower())
                if alignment_enum is not None and cell.paragraphs:
                    cell.paragraphs[0].paragraph_format.alignment = alignment_enum

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
        print(f"警告：图片文件未找到 -> {path}，跳过此图片。")
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

def add_header_from_data(doc, element: dict):
    """
        根据element字典中的数据，在文档中添加页眉。

        Args:
            doc: The python-docx Document object.
            element (dict): 包含页眉数据的字典。
    """
    properties = element.get("properties", {})
    text = properties.get("text", "")

    section = doc.sections[0]
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


def add_footer_from_data(doc, element: dict):
    """
    根据element字典中的数据，在文档中添加页脚。

    Args:
        doc: The python-docx Document object.
        element (dict): 包含页脚数据的字典。
    """
    properties = element.get("properties", {})
    text = properties.get("text", "")

    section = doc.sections[0]
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


def create_document(data: dict):
    """
        根据传入的数据字典，创建一个Word文档对象。

        Args:
            data (dict): 从JSON文件加载的文档结构数据。

        Returns:
            Document: 一个构建好的python-docx的Document对象。
    """
    doc = Document()

    if "page_setup" in data:
        apply_page_setup(doc, data['page_setup'])

    if 'elements' not in data or not isinstance(data['elements'], list):
        print("错误：JSON数据中缺少'elements'列表。")
        return doc # 返回一个空文档

    for element in data['elements']:
        element_type = element.get('type')

        if element.get('type')=='paragraph':
            text = element.get('text', '')
            properties = element.get('properties', {})
            if not isinstance(properties, dict):
                properties = {}

            style = properties.get("style")

            if style and 'Heading' in style:
                p = doc.add_paragraph(text, style=style)
            else:
                p = doc.add_paragraph(text)
                apply_paragraph_properties(p, properties)

        elif element_type == "table":
            add_table_from_data(doc, element)

        elif element_type == "list":
            add_list_from_data(doc, element)

        elif element_type == "image":
            add_image_from_data(doc, element)

        elif element_type == "header":
            add_header_from_data(doc, element)

        elif element_type == "footer":
            add_footer_from_data(doc, element)

    return doc


