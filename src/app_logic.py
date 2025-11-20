# src/app_logic.py
import json
from typing import Callable, Optional
import uuid

# 【核心变更】切换回基于 ai_parser 的顺序工作流
from .ai_parser import parse_natural_language_to_json
from .doc_generator import create_document


async def generate_document_from_command(
        user_command: str,
        logger: Optional[Callable[[str], None]] = None
) -> tuple[bytes | None, str | None, str | None]:
    """
    协调完整的文档生成流程，采用简化的顺序Agent流。
    此版本调用 ai_parser 中的工作流，具备规划、排序、生成和自愈能力。

    Args:
        user_command (str): 用户的完整指令。
        logger (Optional[Callable[[str], None]]): 用于流式日志记录的回调函数。

    Returns:
        tuple[bytes | None, str | None, str | None]: 文档字节流、最终JSON字符串和完整日志。
    """
    log_stream = []

    def log(message: str):
        log_stream.append(message)
        if logger:
            logger(message)

    log("🚀 AI工作流启动...")

    # 【核心变更】调用新的 ai_parser.py 中的函数。
    # 这个函数现在内部处理日志记录并通过回调流式传输。
    parsed_json, ai_log = await parse_natural_language_to_json(
        user_command,
        log_callback=logger
    )

    # ai_log 已经包含了所有日志，这里我们不再需要单独处理

    if not parsed_json:
        log("❌ AI工作流执行失败或未能生成有效文档状态，中止文档生成。")
        return None, None, ai_log

    final_json_str = json.dumps(parsed_json, indent=2, ensure_ascii=False)

    log("\n" + "=" * 20 + " 5. 开始生成 DOCX 文档 " + "=" * 20)
    docx_bytes, generator_log = await create_document(parsed_json)
    log("✅ DOCX 文档生成完毕。")

    # 合并AI日志和生成器日志
    full_log = ai_log + "\n\n--- Generator Log ---\n" + (generator_log or "No generator log available.")

    return docx_bytes, final_json_str, full_log