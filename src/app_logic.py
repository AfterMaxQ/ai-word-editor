import io
import json
from src.ai_parser import parse_natural_language_to_json
from src.doc_generator import create_document
from lxml import etree
from jsonschema import validate, ValidationError

SCHEMA = {
    "type": "object",
    "properties": {
        "page_setup": {"type": "object"},
        "sections": {
            "type": "array",
            "minItems": 1,
            "items": {
                "type": "object",
                "properties": {
                    "properties": {
                        "type": "object",
                        "properties": {
                            "columns": {"type": "number", "minimum": 1}
                        }
                    },
                    "elements": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "type": {
                                    "type": "string",
                                    "enum": ["paragraph", "list", "image", "table", "header", "footer", "page_break", "toc", "formula", "column_break"]
                                },
                                "properties": {"type": "object"},
                                "text": {"type": "string"},
                                "items": {"type": "array", "items": {"type": "string"}},
                                "data": {"type": "array"}
                            },
                            "required": ["type"]
                        }
                    }
                },
                "required": ["elements"]
            }
        }
    },
    "required": ["sections"]
}

# 2. 修改函数签名和返回类型提示
def generate_document_from_command(user_command: str) -> tuple[bytes | None, str | None, str | None]:
    """
    接收用户指令，调用AI和文档引擎，返回Word文档、原始JSON和处理日志。
    """
    document_data, log_str = parse_natural_language_to_json(user_command)
    if not document_data:
        return None, None, log_str

    try:
        # 步骤 1: 严格的Schema验证
        validate(instance=document_data, schema=SCHEMA)
        log_str += "\n✅ AI生成的JSON已通过Schema验证。"

        # 步骤 2: 只有在验证成功后，才继续正常的流程
        json_str_for_display = json.dumps(
            document_data,
            indent=2,
            ensure_ascii=False
        )
        document_bytes, final_xml_body = create_document(document_data)

        if final_xml_body:
            try:
                root = etree.fromstring(final_xml_body.encode('utf-8'))
                pretty_xml = etree.tostring(root, pretty_print=True, encoding='unicode')
                log_str += "\n\n--- 最终生成的完整文档XML (用于诊断) ---\n" + pretty_xml
            except Exception:
                log_str += "\n\n--- 最终生成的完整文档XML (用于诊断) ---\n" + final_xml_body

        print("\n--- 最终生成的完整文档XML (用于诊断) ---")
        print(final_xml_body)

        return document_bytes, json_str_for_display, log_str

    except ValidationError as e:
        # 步骤 3: 如果验证失败，进入这个专门的错误处理流程
        error_msg = f"❌ 致命错误：AI生成的JSON结构不符合规范！请检查日志中的JSON并修正您的Prompt。\n错误详情: {e.message}\n错误的JSON路径: {list(e.path)}"
        log_str += f"\n\n{error_msg}"
        print(error_msg)

        # ★★★【关键修复】★★★
        # 在返回之前，为 json_str_for_display 赋值。
        # 这样即使用户看到的是错误，他们也能在UI上看到那个导致错误的JSON是什么样子。
        json_str_for_display = json.dumps(
            document_data,
            indent=2,
            ensure_ascii=False
        )

        # 返回 None 表示文档生成失败，同时返回包含错误原因的日志和错误的JSON字符串
        return None, json_str_for_display, log_str

    # 5. 成功时返回字节流和JSON字符串
    return filestream.getvalue(), json_str_for_display, log_str