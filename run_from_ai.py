# run_from_ai.py

from src import doc_generator
from src.ai_parser import parse_natural_language_to_json

# 定义输出文件名
OUTPUT_DOCX_FILE = 'output/output_from_ai.docx'


def main():
    """
    从自然语言指令通过AI生成Word文档的主函数。
    """
    # 1. 定义用户的自然语言指令
    user_command = """
    创建一个标题叫'我的相册'。下面写一段话：'这是我最喜欢的一张照片：'。然后插入一张图片，路径是'assets/test_image.png'，宽度设置为12厘米。
    """

    # 2. 调用AI解析器
    document_data = parse_natural_language_to_json(user_command)

    if not document_data:
        print("❌ 文档生成失败，AI解析步骤出错。")
        return

    # 3. 调用核心引擎创建文档
    print("\n⚙️ 正在调用文档生成引擎...")
    document_object = doc_generator.create_document(document_data)
    print("✅ 成功创建Word文档对象！")

    # 4. 保存文档
    document_object.save(OUTPUT_DOCX_FILE)
    print(f"🎉 成功将文档保存为 '{OUTPUT_DOCX_FILE}'！")


if __name__ == "__main__":
    main()