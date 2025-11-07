# src/ai_parser.py

import requests
import json
import re

# 定义Ollama API的地址和模型名称
OLLAMA_API_URL = "http://localhost:11434/api/chat"

"""
    NAME                 ID              SIZE      MODIFIED
    qwen2.5-coder:14b    9ec8897f747e    9.0 GB    4 minutes ago
    qwen2.5-coder:7b     dae161e27b0e    4.7 GB    14 minutes ago
    deepseek-r1:14b      c333b7232bdb    9.0 GB    2 hours ago
    deepseek-r1:7b       755ced02ce7b    4.7 GB    2 hours ago
    llama3:8b            365c0bd3c000    4.7 GB    16 hours ago
"""

MODEL_NAME = "qwen2.5-coder:7b"
SYSTEM_PROMPT_FILE = "prompts/system_prompt.txt"

def split_command_into_chunks(user_command: str, max_chunks: int = 5):
    """
        将用户的长指令分割成更小的、符合逻辑的块。

        为什么这么做？
        - 我们发现，一次性将一个非常长的指令（例如，包含10个步骤）交给一个7B大小的模型，
          它很容易在生成JSON的过程中“忘记”前面的指令，或者最终的JSON结构会非常混乱。
        - 通过将指令按自然语言的换行符（代表一个独立的步骤）分割，我们可以一次只让模型专注于一个子任务。
          这就像我们指导新手一样，一步一步来，而不是一次性告诉他所有事情。

        Args:
            user_command (str): 用户的完整自然语言指令。
            max_chunks (int): 为了防止指令被过度分割（例如，一个表格的每一行都被分开），
                              我们设置一个最大分块数。超过这个数量，后面的内容会合并到最后一个块中。

        Returns:
            list[str]: 一个包含指令块字符串的列表。
        """
    # 1. 使用正则表达式按一个或多个换行符进行分割
    lines = re.split(r'\n\s*\n*', user_command.strip())
    # 2. 过滤掉所有仅包含空白字符的无效行
    chunks = [chunk.strip() for chunk in lines if chunk.strip()]

    # 3. 如果分割后的块数超过了最大限制
    if len(chunks) > max_chunks:
        print(f"警告：指令被分割成 {len(chunks)} 块，超过最大限制 {max_chunks}。")
        # 将超出的部分合并到最后一个块中
        last_valid_chunk = "\n".join(chunks[max_chunks - 1:])
        chunks = chunks[:max_chunks - 1] + [last_valid_chunk]
        print(f"已将指令合并为 {len(chunks)} 块进行处理。")

    return chunks


def parse_natural_language_to_json(user_command: str) -> dict | None:
    """
        将用户的自然语言指令发送给本地LLM，并解析返回的JSON。
    此函数现在支持将长指令分块，以提高稳定性和处理复杂指令的能力。

    Args:
        user_command (str): 用户的自然语言指令。

    Returns:
        dict | None: 解析成功则返回包含文档结构的字典，否则返回None。
    """

    # 1.读取我们的“prompt”
    try:
        with open(SYSTEM_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        print(f"错误：系统提示文件未找到 -> {SYSTEM_PROMPT_FILE}")
        return None

    # 2. 将用户的完整指令分割成多个块
    chunks = split_command_into_chunks(user_command)

    # 初始化一个最终的JSON对象和所有元素的列表
    aggregated_document_data = {
        "elements": []
    }

    print(f"🧠 指令已被分为 {len(chunks)} 个任务块，开始逐一调用AI解析器...")

    for i, chunk in enumerate(chunks):
        print(f"\n--- 正在处理第 {i + 1}/{len(chunks)} 个任务块 ---")
        print(f"指令内容: \"{chunk}\"")
        # 4. 为每个块构建特定的请求
        chunk_user_prompt = f"""
        This is part {i + 1} of a multi-part command.
        The user's command for THIS part is: "{chunk}"
        CONTEXT: So far, the following number of elements have been generated: {len(aggregated_document_data['elements'])}. 
        Please generate the JSON structure ONLY for the command in THIS part. Do not repeat or re-generate previous elements."""

        payload = {
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk_user_prompt}
            ],
            "format": "json",
            "stream": False
        }

        # 3. 发送HTTP POST请求
        try:
            response = requests.post(OLLAMA_API_URL, json=payload, timeout=60)
            response.raise_for_status() # 如果HTTP状态码是4xx或5xx，则抛出异常
            # 解析返回的响应
            response_data = response.json()
            message_content = response_data.get('message', {}).get('content')

            if not message_content:
                print(f"错误：第 {i + 1} 个块的AI响应中找不到内容。")
                return None

            # 解析当前块返回的JSON片段
            chunk_json = json.loads(message_content)

            print(f"--- AI为块 {i + 1} 返回的JSON片段 ---")
            print(json.dumps(chunk_json, indent=2, ensure_ascii=False))
            print("--------------------------------")

            # 6. JSON聚合：将新生成的元素合并到最终结果中
            #    这是整个流程的关键一步，我们将所有“零件”组装成一个完整的产品。
            new_elements = chunk_json.get('elements', [])
            if new_elements:
                aggregated_document_data['elements'].extend(new_elements)
                print(f"✅ 成功聚合 {len(new_elements)} 个新元素。")

            # 同时，检查是否有页面设置，并更新到主对象中
            # 这允许用户在任何步骤中设置页面格式
            if 'page_setup' in chunk_json:
                # 使用.update()可以合并字典，或添加新键
                if 'page_setup' not in aggregated_document_data:
                    aggregated_document_data['page_setup'] = {}
                aggregated_document_data['page_setup'].update(chunk_json['page_setup'])
                print("✅ 已更新页面设置。")


        except requests.exceptions.RequestException as e:
            print(f"错误：在处理第 {i + 1} 个块时连接Ollama API失败 -> {e}")
            return None
        except json.JSONDecodeError as e:
            print(f"错误：在处理第 {i + 1} 个块时AI返回的不是有效的JSON格式 -> {e}")
            print(f"收到的内容: {message_content}")
            return None

    print("\n✅ 所有任务块处理完毕，AI解析成功！")
    return aggregated_document_data



# --- 测试代码 ---
if __name__ == "__main__":

    test_command = """
    给我一个一级标题叫'销售报告'。
    然后另起一段，内容是'这是第一季度的总结'，宋体小四，首行缩进。
    最后，给我一个3x3的表格，带表头，内容是姓名、年龄、城市，张三、30、北京，李四、25、上海。第一列左对齐，后两列居中。
    """

    document_structure = parse_natural_language_to_json(test_command)
    if document_structure:
        print("\n--- 解析得到的JSON结构 ---")
        print(json.dumps(document_structure, indent=2, ensure_ascii=False))


