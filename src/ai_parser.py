# src/ai_parser.py

import requests
import json
import re
from lxml import etree
import math

# --- 常量定义部分保持不变 ---
OLLAMA_API_URL = "http://localhost:11434/api/chat"
MODEL_NAME = "qwen2.5-coder:14b"
SYSTEM_PROMPT_FILE = "prompts/system_prompt.txt"
LATEX_PROMPT_FILE = "prompts/prompt_for_latex_convert.txt"


# --- translate_latex_to_omml_llm 函数保持不变 ---
def translate_latex_to_omml_llm(latex_string: str) -> str | None:
    # ... 此函数已有足够的 print 输出，无需修改 ...
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
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=60)
        response.raise_for_status()
        response_data = response.json()
        omml_xml_string = response_data.get('message', {}).get('content')

        if not omml_xml_string:
            print("❌ LLM返回内容为空。")
            return None

        try:
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


# ★★★ 已添加详细控制台日志 ★★★
def split_command_into_chunks(user_command: str, max_chunks: int = 5) -> tuple[list[str], str]:
    """
    【动态分片核心实现】
    将用户的长指令分割成更小的、符合逻辑的任务块。
    """
    log_messages = []

    print("\n" + "=" * 20 + " 1. 开始智能指令分割 " + "=" * 20)

    # 粗分
    logical_units = re.split(r'\n\s*\n+', user_command.strip())
    logical_units = [unit.strip() for unit in logical_units if unit.strip()]
    total_units = len(logical_units)

    print(f"[控制台] 粗粒度分割：找到 {total_units} 个逻辑单元。")
    log_messages.append(f"🧠 指令被初步分解为 {total_units} 个逻辑单元。")

    if total_units <= max_chunks:
        print(f"[控制台] 逻辑单元数 ({total_units}) <= 最大分块数 ({max_chunks})，无需合并。")
        log_messages.append(f"  - 单元数 ({total_units}) 不超过最大分块数 ({max_chunks})，无需合并。")
        print("=" * 62 + "\n")
        return logical_units, "\n".join(log_messages)

    # 精合
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


# ★★★ 已添加详细控制台日志 ★★★
def parse_natural_language_to_json(user_command: str) -> tuple[dict | None, str]:
    """
    将用户的自然语言指令分块发送给LLM，并返回最终的JSON和详细的处理日志。
    """
    log_messages = []
    try:
        with open(SYSTEM_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        error_msg = f"❌ 错误：系统提示文件未找到 -> {SYSTEM_PROMPT_FILE}"
        print(f"[控制台] {error_msg}")
        return None, error_msg

    chunks, split_log = split_command_into_chunks(user_command, max_chunks=5)
    log_messages.append(split_log)

    aggregated_document_data = {"elements": []}

    print("=" * 20 + " 2. 开始循环处理任务块 " + "=" * 20)
    log_messages.append(f"\n--- 开始逐一调用AI解析器处理 {len(chunks)} 个任务块 ---")

    for i, chunk in enumerate(chunks):
        print(f"\n--- [控制台] 正在处理第 {i + 1}/{len(chunks)} 个任务块 ---")
        log_messages.append(f"\n--- 正在处理第 {i + 1}/{len(chunks)} 个任务块 ---")

        print(f"[控制台] 任务块内容:\n---\n{chunk}\n---")
        log_messages.append(f"📄 指令内容:\n---\n{chunk}\n---")

        # 构建上下文感知的 Prompt
        context_summary = f"So far, {len(aggregated_document_data.get('elements', []))} elements have been generated."
        chunk_user_prompt = f"""
        This is part {i + 1} of a multi-part command.
        The user's command for THIS part is: "{chunk}"
        CONTEXT: {context_summary}. 
        Please generate the JSON structure ONLY for the command in THIS part. Do not repeat or re-generate previous elements."""

        print("[控制台] 为此任务块生成的 User Prompt:")
        print(chunk_user_prompt)

        payload = {
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk_user_prompt}
            ],
            "format": "json",
            "stream": False
        }

        try:
            print("[控制台] 正在向 Ollama API 发送请求...")
            response = requests.post(OLLAMA_API_URL, json=payload, timeout=60)
            response.raise_for_status()
            response_data = response.json()
            message_content = response_data.get('message', {}).get('content')
            print("[控制台] 已收到 AI 响应。")

            if not message_content:
                error_msg = f"❌ 错误：第 {i + 1} 个块的AI响应中找不到内容。"
                print(f"[控制台] {error_msg}")
                log_messages.append(error_msg)
                return None, "\n".join(log_messages)

            chunk_json = json.loads(message_content)

            print(f"[控制台] AI为块 {i + 1} 返回的JSON片段:")
            print(json.dumps(chunk_json, indent=2, ensure_ascii=False))
            log_messages.append(f"🤖 AI为块 {i + 1} 返回的JSON片段:")
            log_messages.append(json.dumps(chunk_json, indent=2, ensure_ascii=False))

            # 聚合 JSON
            new_elements = chunk_json.get('elements', [])
            if new_elements:
                if 'elements' not in aggregated_document_data:
                    aggregated_document_data['elements'] = []
                aggregated_document_data['elements'].extend(new_elements)
                print(f"[控制台] 成功聚合 {len(new_elements)} 个新元素。")
                log_messages.append(f"✅ 成功聚合 {len(new_elements)} 个新元素。")

            if 'page_setup' in chunk_json:
                if 'page_setup' not in aggregated_document_data:
                    aggregated_document_data['page_setup'] = {}
                aggregated_document_data['page_setup'].update(chunk_json['page_setup'])
                print("[控制台] 已更新页面设置。")
                log_messages.append("✅ 已更新页面设置。")

        except requests.exceptions.RequestException as e:
            error_msg = f"❌ 错误：在处理第 {i + 1} 个块时连接Ollama API失败 -> {e}"
            print(f"[控制台] {error_msg}")
            log_messages.append(error_msg)
            return None, "\n".join(log_messages)
        except json.JSONDecodeError as e:
            error_msg = f"❌ 错误：在处理第 {i + 1} 个块时AI返回的不是有效的JSON格式 -> {e}"
            print(f"[控制台] {error_msg}")
            log_messages.append(error_msg)
            print(f"[控制台] 收到的原始响应内容: {message_content}")
            log_messages.append(f"收到的内容: {message_content}")
            return None, "\n".join(log_messages)

    print("\n" + "=" * 20 + " 3. 所有任务块处理完毕 " + "=" * 20)
    log_messages.append("\n✅ 所有任务块处理完毕，AI解析成功！")
    return aggregated_document_data, "\n".join(log_messages)