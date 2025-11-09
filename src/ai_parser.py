# src/ai_parser.py

import json
import math
import re
import requests
from lxml import etree

# 导入我们自己的库和 Pydantic 模型
from .ollama_pydantic import create
from .schemas import DocumentModel

"""NAME                     ID              SIZE      MODIFIED
deepseek-coder-v2:16b    63fb193b3a9b    8.9 GB    6 hours ago
qwen3-vl:235b-cloud      86b3322ec200    -         7 hours ago
gpt-oss:120b-cloud       569662207105    -         24 hours ago
gpt-oss:20b              17052f91a42e    13 GB     24 hours ago
qwen3-vl:8b              901cae732162    6.1 GB    24 hours ago
qwen3-vl:4b              1343d82ebee3    3.3 GB    24 hours ago
qwen2.5-coder:14b        9ec8897f747e    9.0 GB    2 days ago
qwen2.5-coder:7b         dae161e27b0e    4.7 GB    2 days ago
deepseek-r1:14b          c333b7232bdb    9.0 GB    2 days ago
deepseek-r1:7b           755ced02ce7b    4.7 GB    2 days ago
llama3:8b                365c0bd3c000    4.7 GB    2 days ago"""

# --- 常量定义 ---
MODEL_NAME = "deepseek-coder-v2:16b"
POLISH_MODEL_NAME = "deepseek-coder-v2:16b"
SYSTEM_PROMPT_FILE = "prompts/system_prompt.txt"
LATEX_PROMPT_FILE = "prompts/prompt_for_latex_convert.txt"
OLLAMA_API_URL = "http://localhost:11434/api/chat"
POLISH_PROMPT_FILE = "prompts/prompt_for_polishing.txt"


def translate_latex_to_omml_llm(latex_string: str) -> str | None:
    """
    (此函数保持不变)
    使用LLM将LaTeX字符串转换为OMML。这是一个独立的辅助函数。
    """
    print(f"🤖 尝试使用LLM转译LaTeX: {latex_string}")
    try:
        with open(LATEX_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        print(f"❌ 致命错误: 无法找到LaTeX转换提示词文件 -> {LATEX_PROMPT_FILE}")
        return None

    user_prompt = f"""
Convert the following LaTeX formula into a centered OMML `<m:oMathPara>` XML block.
LaTeX Input: `{latex_string}`
Alignment: center
"""

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=5000)
        response.raise_for_status()
        response_data = response.json()
        omml_xml_string = response_data.get('message', {}).get('content')

        if not omml_xml_string:
            print("❌ LLM返回内容为空。")
            return None

        try:
            # 清理LLM可能返回的markdown代码块标记
            omml_xml_string = re.sub(r'^```xml\s*|\s*```$', '', omml_xml_string, flags=re.MULTILINE).strip()
            etree.fromstring(omml_xml_string)
            print("✅ LLM转译成功并已通过XML验证。")
            return omml_xml_string
        except etree.XMLSyntaxError as e:
            print(f"❌ LLM返回的不是有效的XML，验证失败: {e}")
            print(f"收到的内容: {omml_xml_string}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"❌ 调用LLM进行公式转译失败: {e}")
        return None


def split_command_into_chunks(user_command: str, max_chunks: int = 30) -> tuple[list[str], str]:
    """
    (此函数保持不变)
    【动态分片核心实现】
    将用户的长指令分割成更小的、符合逻辑的任务块。
    """
    log_messages = []

    print("\n" + "=" * 20 + " 1. 开始智能指令分割 " + "=" * 20)

    # 粗分：根据一个或多个空行来分割
    logical_units = re.split(r'\n\s*\n+', user_command.strip())
    logical_units = [unit.strip() for unit in logical_units if unit.strip()]
    total_units = len(logical_units)

    print(f"[控制台] 粗粒度分割：找到 {total_units} 个逻辑单元。")
    log_messages.append(f"🧠 指令被初步分解为 {total_units} 个逻辑单元。")

    # 如果单元数在限制内，直接返回，无需合并
    if total_units <= max_chunks:
        print(f"[控制台] 逻辑单元数 ({total_units}) <= 最大分块数 ({max_chunks})，无需合并。")
        log_messages.append(f"  - 单元数 ({total_units}) 不超过最大分块数 ({max_chunks})，无需合并。")
        print("=" * 62 + "\n")
        return logical_units, "\n".join(log_messages)

    # 精合：如果单元数过多，则进行智能分组
    print(f"[控制台] 逻辑单元数 ({total_units}) > 最大分块数 ({max_chunks})，开始智能分组。")
    log_messages.append(f"  - 单元数 ({total_units}) 超过最大分块数 ({max_chunks})，开始智能分组...")

    units_per_chunk = math.ceil(total_units / max_chunks)
    print(f"[控制台] 计算得出：每个任务块应包含约 {units_per_chunk} 个逻辑单元。")
    log_messages.append(f"  - 计算得出：每个任务块应包含约 {units_per_chunk} 个逻辑单元。")

    final_chunks = []
    for i in range(0, total_units, units_per_chunk):
        group = logical_units[i:i + units_per_chunk]
        combined_chunk = "\n".join(group)
        final_chunks.append(combined_chunk)

    print(f"[控制台] 成功将 {total_units} 个逻辑单元合并为 {len(final_chunks)} 个最终任务块。")
    log_messages.append(f"✅ 成功将 {total_units} 个逻辑单元合并为 {len(final_chunks)} 个最终任务块。")
    print("=" * 62 + "\n")

    return final_chunks, "\n".join(log_messages)


def parse_natural_language_to_json(user_command: str) -> tuple[dict | None, str]:
    """
    将用户的自然语言指令分块发送给LLM，并返回最终的JSON和详细的处理日志。
    (已重构为使用自研的 ollama_pydantic 库)
    """
    log_messages = []
    try:
        with open(SYSTEM_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        error_msg = f"❌ 错误：系统提示文件未找到 -> {SYSTEM_PROMPT_FILE}"
        print(f"[控制台] {error_msg}")
        return None, error_msg

    chunks, split_log = split_command_into_chunks(user_command, max_chunks=30)
    log_messages.append(split_log)

    aggregated_document_data = {"sections": []}

    print("=" * 20 + " 2. 开始循环处理任务块 " + "=" * 20)
    log_messages.append(f"\n--- 开始逐一调用AI解析器处理 {len(chunks)} 个任务块 ---")

    for i, chunk in enumerate(chunks):
        print(f"\n--- [控制台] 正在处理第 {i + 1}/{len(chunks)} 个任务块 ---")
        log_messages.append(f"\n--- 正在处理第 {i + 1}/{len(chunks)} 个任务块 ---")
        print(f"[控制台] 任务块内容:\n---\n{chunk}\n---")
        log_messages.append(f"📄 指令内容:\n---\n{chunk}\n---")

        context_summary = f"So far, {len(aggregated_document_data.get('sections', []))} sections have been generated."
        chunk_user_prompt = f"""
CRITICAL INSTRUCTION: You are a component in a larger system. 
Your SOLE task is to convert the user's command for THIS specific part into a JSON structure that conforms to the Pydantic model.
You MUST NOT add any content, text, or elements not explicitly requested in the command below.
This is part {i + 1} of a multi-part command.
The user's command for THIS part is:
---
{chunk}
---
CONTEXT: {context_summary}.

YOUR TASK:
1.  Analyze the command for THIS part ONLY.
2.  If the command is purely instructional, transitional (e.g., "Next, do the following"), or a summary, and contains NO concrete content to add to the document, you MUST return a valid JSON that will result in an empty Pydantic model (e.g., {{}} or {{"sections": []}}).
3.  Otherwise, generate the JSON structure strictly for the elements described in THIS part. Do not hallucinate or create extra content.
"""
        print("[控制台] 为此任务块生成的 User Prompt:")
        print(chunk_user_prompt)

        # ▼▼▼【核心调用逻辑】▼▼▼
        # 调用我们自己的 `ollama_pydantic.create` 函数
        chunk_model = create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk_user_prompt}
            ],
            response_model=DocumentModel,
            max_retries=2,
        )
        # ▲▲▲【核心调用逻辑】▲▲▲

        if chunk_model is None:
            # 如果 `create` 函数在所有重试后返回 None，说明彻底失败
            error_msg = f"❌ 致命错误：在处理第 {i + 1} 个块时，AI在多次尝试后仍无法生成有效JSON。"
            print(f"[控制台] {error_msg}")
            log_messages.append(error_msg)
            return None, "\n".join(log_messages)

        # 将返回的 Pydantic 模型实例转换为字典，用于后续的聚合操作
        chunk_json = chunk_model.model_dump(exclude_unset=True)

        print(f"[控制台] AI为块 {i + 1} 返回的已验证JSON片段:")
        print(json.dumps(chunk_json, indent=2, ensure_ascii=False))
        log_messages.append(f"🤖 AI为块 {i + 1} 返回的已验证JSON片段:")
        log_messages.append(json.dumps(chunk_json, indent=2, ensure_ascii=False))

        # 聚合逻辑
        new_sections = chunk_json.get('sections', [])
        if new_sections:
            if 'sections' not in aggregated_document_data:
                aggregated_document_data['sections'] = []
            aggregated_document_data['sections'].extend(new_sections)
            print(f"[控制台] 成功聚合 {len(new_sections)} 个新节(section)。")
            log_messages.append(f"✅ 成功聚合 {len(new_sections)} 个新节(section)。")

        if 'page_setup' in chunk_json:
            if 'page_setup' not in aggregated_document_data:
                aggregated_document_data['page_setup'] = {}
            aggregated_document_data['page_setup'].update(chunk_json['page_setup'])
            print("[控制台] 已更新页面设置。")
            log_messages.append("✅ 已更新页面设置。")

    print("\n" + "=" * 20 + " 3. 所有任务块处理完毕 " + "=" * 20)
    log_messages.append("\n✅ 所有任务块处理完毕，AI解析成功！")
    return aggregated_document_data, "\n".join(log_messages)

def polish_user_prompt_llm(user_command: str) -> str | None:
    """
    使用LLM将用户输入的模糊指令润色成清晰、结构化的指令。
    """
    print(f" polishing user prompt: {user_command}")
    try:
        with open(POLISH_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        print(f"❌ 致命错误: 无法找到指令润色提示词文件 -> {POLISH_PROMPT_FILE}")
        return None

    # 将用户的原始指令附加到系统提示词的末尾
    full_prompt = f"{system_prompt}\n{user_command}"

    payload = {
        "model": POLISH_MODEL_NAME, # 可以为这个任务选择一个不同的、更擅长创意的模型
        "messages": [
            {"role": "system", "content": "You are a helpful assistant that rewrites user commands."}, # 简单的系统角色
            {"role": "user", "content": full_prompt}
        ],
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=5000)
        response.raise_for_status()
        response_data = response.json()
        polished_command = response_data.get('message', {}).get('content')

        if not polished_command:
            print("❌ LLM返回的润色内容为空。")
            return user_command # 如果失败，返回原始指令

        # 简单的清理，移除可能的前后空行
        return polished_command.strip()

    except requests.exceptions.RequestException as e:
        print(f"❌ 调用LLM进行指令润色失败: {e}")
        return user_command # 如果失败，返回原始指令