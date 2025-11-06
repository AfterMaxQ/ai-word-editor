import json
import sys
from docx import Document

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

            style = properties.get("style", '')

            if style=='Heading 1':
                doc.add_paragraph(text, style='Heading 1')
            else:
                doc.add_paragraph(text)

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
