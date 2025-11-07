import io
from src.ai_parser import parse_natural_language_to_json
from src.doc_generator import create_document

def generate_document_from_command(user_command: str) -> bytes | None :
    """
        接收用户指令，调用AI和文档引擎，返回Word文档的二进制数据。

        Args:
            user_command (str): 用户的自然语言指令。

        Returns:
            bytes | None: 成功则返回.docx文件的字节流，失败则返回None。
    """

    #调用ai解析器
    document_data = parse_natural_language_to_json(user_command)
    if not document_data:
        return None

    # 调用文档生成引擎
    document_object = create_document(document_data)

    #将文档保存到内存的二进制流, 而不是物理文件
    filestream = io.BytesIO()
    document_object.save(filestream)
    filestream.seek(0)

    return filestream.getvalue()

