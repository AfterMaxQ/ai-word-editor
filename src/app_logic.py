import io
import json
from src.ai_parser import parse_natural_language_to_json
from src.doc_generator import create_document
from lxml import etree
from jsonschema import validate, ValidationError
from src.ai_parser import parse_natural_language_to_json, polish_user_prompt_llm


def generate_document_from_command(user_command: str) -> tuple[bytes | None, str | None, str | None]:
    """
    接收用户指令，调用AI和文档引擎，返回Word文档、原始JSON和处理日志。
    """
    document_data, log_str = parse_natural_language_to_json(user_command)
    if not document_data:
        return None, None, log_str

    log_str += "\n✅ AI生成的JSON已通过Pydantic模型验证（在生成时完成）。"

    json_str_for_display = json.dumps(
        document_data,
        indent=2,
        ensure_ascii=False
    )

    # create_document 内部逻辑不变
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

def polish_command(user_command: str) -> str | None:
    """
    调用AI润色用户指令。
    """
    return polish_user_prompt_llm(user_command)