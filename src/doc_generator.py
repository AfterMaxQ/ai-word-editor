# src/doc_generator.py

import json
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
from docx.enum.section import WD_ORIENT

from docx.shared import RGBColor
from .ai_parser import translate_latex_to_omml_llm
from .latex_converter import latex_to_omml

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
    indent_cm = properties.get('first_line_indent')
    if indent_cm is not None:
        p_format.first_line_indent = Cm(indent_cm)

    # --- 设置字体格式 (字体、字号等), 字体格式需要应用到段落内的Run上。---
    if paragraph.runs:
        font = paragraph.runs[0].font
        font_name = properties.get('font_name')
        # 设置字体名称
        if font_name:
            font.name = font_name
            # 导入中文字体所需的包
            from docx.oxml.ns import qn
            # 设置中文字体 (东亚字体)
            font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

        # 设置字体大小
        font_size_pt = properties.get('font_size')
        if font_size_pt is not None:
            font.size = Pt(font_size_pt)
        #设置粗体
        is_bold = properties.get('bold')
        if is_bold is not None:  # 可以是 True 或 False, 但不能是 None
            font.bold = bool(is_bold)



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


def add_page_break_from_data(doc, element: dict):
    """
    在文档中添加一个分页符。

    Args:
        doc: The python-docx Document object.
        element (dict): 包含分页符数据的字典 (虽然此元素为空)。
    """
    doc.add_page_break()

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


# ★★★ 新增核心函数: 三层防御公式处理器 ★★★
def get_formula_xml_and_placeholder(element: dict) -> tuple[str | None, str | None]:
    """
    ★★★【三层防御公式处理器】★★★
    1. 尝试本地转换器 (高可靠性)
    2. 若失败，回退到 LLM (高覆盖性)
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
    使用“占位符与健壮XML注入”两阶段方法，根据经过验证的JSON数据创建Word文档。
    这是文档生成引擎的最终实现，它通过彻底分离高层API操作和底层XML注入，
    从根本上解决了所有“文件损坏”问题，是目前最健壮的方案。
    """
    doc = Document()
    formulas_to_inject = {}

    # --- 阶段一: 使用 python-docx 生成带占位符的草稿文档 ---
    page_setup_data = data.get('page_setup', {})
    elements = data.get('elements', [])

    if page_setup_data:
        apply_page_setup(doc, page_setup_data)

    for element in elements:
        element_type = element.get('type')
        if element_type == "formula":
            placeholder, omml_xml = get_formula_xml_and_placeholder(element)
            if placeholder and omml_xml:
                p = doc.add_paragraph()
                p.add_run(placeholder)
                formulas_to_inject[placeholder] = omml_xml
        # ... (其他 element_type 的处理保持不变)
        elif element_type == "header":
            add_header_from_data(doc, element)
        elif element_type == "footer":
            add_footer_from_data(doc, element)
        elif element_type == 'paragraph':
            text = element.get('text', '')
            properties = element.get('properties', {})
            style = properties.get("style") if isinstance(properties, dict) else None
            p = doc.add_paragraph(text, style=style)
            if isinstance(properties, dict): apply_paragraph_properties(p, properties)
        elif element_type == "table":
            add_table_from_data(doc, element)
        elif element_type == "list":
            add_list_from_data(doc, element)
        elif element_type == "image":
            add_image_from_data(doc, element)
        elif element_type == "page_break":
            add_page_break_from_data(doc, element)
        elif element_type == "toc":
            add_toc_from_data(doc, element)

    if not formulas_to_inject:
        # 如果没有公式，直接保存并返回
        draft_stream = io.BytesIO()
        doc.save(draft_stream)
        draft_stream.seek(0)
        final_xml_body = etree.tostring(doc.element.body, pretty_print=True, encoding='unicode')
        return draft_stream.getvalue(), final_xml_body

    # --- 阶段二: 终极版后处理 (使用lxml树操作) ---
    print("\n--- 开始XML后处理阶段 ---")
    draft_stream = io.BytesIO()
    doc.save(draft_stream)
    draft_stream.seek(0)

    final_stream = io.BytesIO()
    diagnostic_xml = ""

    with zipfile.ZipFile(draft_stream, 'r') as zin, zipfile.ZipFile(final_stream, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)

            if item.filename == 'word/document.xml':
                print("  正在处理: word/document.xml")
                root = etree.fromstring(content)
                body = root.find(qn('w:body'))

                # ★★★ 这是核心逻辑修正点 ★★★
                # 我们不再直接替换 <w:p>，而是修改其内容
                paragraphs = body.findall('.//' + qn('w:p'))
                for p in paragraphs:
                    # 获取段落内所有文本，这是最可靠的比较方法
                    placeholder_text = p.xpath("string(.)").strip()

                    if placeholder_text in formulas_to_inject:
                        real_xml_str = formulas_to_inject[placeholder_text]
                        print(f"  > 找到占位符: '{placeholder_text}'")
                        try:
                            # 将待注入的XML字符串解析为lxml元素
                            new_content_element = etree.fromstring(real_xml_str)

                            # 清空占位符段落<w:p>内的所有内容 (即所有<w:r>元素)
                            for child in list(p):
                                p.remove(child)

                            # 如果注入的是数学段落，我们需要提取其内容(<m:oMath>)
                            if new_content_element.tag == qn('m:oMathPara'):
                                # 将<m:oMathPara>的属性(如居中)复制到<w:p>的属性中
                                para_props = new_content_element.find(qn('m:oMathParaPr'))
                                if para_props is not None:
                                    # 创建或获取<w:pPr>
                                    pPr = p.find(qn('w:pPr'))
                                    if pPr is None:
                                        pPr = OxmlElement('w:pPr')
                                        p.insert(0, pPr)
                                    # 复制对齐属性
                                    jc = para_props.find(qn('m:jc'))
                                    if jc is not None:
                                        new_jc = OxmlElement('w:jc')
                                        new_jc.set(qn('w:val'), jc.get(qn('m:val')))
                                        pPr.append(new_jc)

                                # 提取<m:oMath>元素并注入
                                oMath = new_content_element.find(qn('m:oMath'))
                                if oMath is not None:
                                    p.append(oMath)
                                    print(f"    - 成功将 <m:oMath> 注入到 <w:p> 中。")

                            else:  # 如果注入的是错误信息(已经是<w:p>格式)
                                # 提取其内容(<w:r>)并注入
                                for run in new_content_element.findall('.//' + qn('w:r')):
                                    p.append(run)
                                print(f"    - 成功将错误提示 <w:r> 注入到 <w:p> 中。")

                        except etree.XMLSyntaxError as e:
                            print(f"  ❌ 严重错误: 无法解析待注入的XML，跳过: {e}")
                            continue

                final_xml_content = etree.tostring(root, encoding='utf-8', xml_declaration=True, standalone=True)
                diagnostic_xml = etree.tostring(root, pretty_print=True, encoding='unicode')
                zout.writestr(item, final_xml_content)
            else:
                zout.writestr(item, content)

    print("--- XML后处理阶段完成 ---\n")
    final_stream.seek(0)
    return final_stream.getvalue(), diagnostic_xml


