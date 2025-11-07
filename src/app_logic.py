import io
import json
from src.ai_parser import parse_natural_language_to_json
from src.doc_generator import create_document

# 2. 修改函数签名和返回类型提示
def generate_document_from_command(user_command: str) -> tuple[bytes | None, str | None]:
    """
        接收用户指令，调用AI和文档引擎，返回Word文档的二进制数据和原始JSON。

        Args:
            user_command (str): 用户的自然语言指令。

        Returns:
            tuple[bytes | None, str | None]: 
                成功则返回 (.docx文件的字节流, JSON字符串), 
                失败则返回 (None, None)。
    """

    # 调用ai解析器
    document_data = parse_natural_language_to_json(user_command)
    if not document_data:
        # 3. 失败时返回两个 None
        return None, None

    # 4. 将解析出的字典转换为格式化的JSON字符串，以便显示
    json_str_for_display = json.dumps(
        document_data, 
        indent=2, 
        ensure_ascii=False
    )

    # 调用文档生成引擎
    document_object = create_document(document_data)

    # 将文档保存到内存的二进制流, 而不是物理文件
    filestream = io.BytesIO()
    document_object.save(filestream)
    filestream.seek(0)

    # 5. 成功时返回字节流和JSON字符串
    return filestream.getvalue(), json_str_for_display