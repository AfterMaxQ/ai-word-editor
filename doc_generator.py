import json
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

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


def create_document(data: dict):
    """
        根据传入的数据字典，创建一个Word文档对象。

        Args:
            data (dict): 从JSON文件加载的文档结构数据。

        Returns:
            Document: 一个构建好的python-docx的Document对象。
    """
    doc = Document()

    if 'elements' not in data or not isinstance(data['elements'], list):
        print("错误：JSON数据中缺少'elements'列表。")
        return doc # 返回一个空文档

    for element in data['elements']:
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

    return doc


if __name__ == "__main__":
    # 1. 加载数据
    file_path = 'document_structure.json'
    document_data = load_document_data(file_path)
    print("✅ 成功读取JSON文件！")
    # 2. 创建文档
    document_object = create_document(document_data)
    print("✅ 成功创建Word文档对象！")
    # 3. 保存文档
    output_filename = 'output.docx'
    document_object.save(output_filename)
    print(f"✅ 成功将文档保存为 '{output_filename}'！")
    print("\n请打开项目文件夹查看生成的Word文档。")
