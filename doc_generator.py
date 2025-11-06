import json
import sys

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

if __name__ == "__main__":
    filepath = 'document_structure.json'
    document_data = load_document_data(filepath)

    print("成功读取JSON文件")
    print("数据类型: ", type(document_data))
    print("内容预览:")
    print(json.dumps(document_data, indent=2, ensure_ascii=False))