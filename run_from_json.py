# run_from_json.py

from src import doc_generator
import json

# 定义输入和输出文件名
INPUT_JSON_FILE = 'data/document_structure.json'
OUTPUT_DOCX_FILE = 'output_from_json.docx'


def main():
    """
    从本地JSON文件生成Word文档的主函数。
    """
    print(f"📄 正在从 '{INPUT_JSON_FILE}' 读取数据...")

    # 1. 加载本地JSON数据
    try:
        with open(INPUT_JSON_FILE, 'r', encoding='utf-8') as f:
            document_data = json.load(f)
        print("✅ 成功读取JSON文件！")
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"错误：无法读取或解析JSON文件 -> {e}")
        return

    # 2. 调用核心引擎创建文档
    print("⚙️ 正在调用文档生成引擎...")
    document_object = doc_generator.create_document(document_data)
    print("✅ 成功创建Word文档对象！")

    # 3. 保存文档
    document_object.save(OUTPUT_DOCX_FILE)
    print(f"🎉 成功将文档保存为 '{OUTPUT_DOCX_FILE}'！")


if __name__ == "__main__":
    main()