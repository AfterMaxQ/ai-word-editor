# src/ai_parser.py

from http.client import responses

import requests
import json

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

MODEL_NAME = "qwen2.5-coder:14b"
SYSTEM_PROMPT_FILE = "prompts/system_prompt.txt"

def parse_natural_language_to_json(user_command: str) -> dict | None:
    """
        将用户的自然语言指令发送给本地LLM，并解析返回的JSON。

        Args:
            user_command (str): 用户的自然语言指令。

        Returns:
            dict | None: 解析成功则返回包含文档结构的字典，否则返回None。
    """
    print("🧠 正在调用AI解析器，请稍候...")

    # 1.读取我们的“prompt”
    try:
        with open(SYSTEM_PROMPT_FILE, 'r', encoding='utf-8') as f:
            system_prompt = f.read()
    except FileNotFoundError:
        print(f"错误：系统提示文件未找到 -> {SYSTEM_PROMPT_FILE}")
        return None

    # 2. 构建发送给Ollama API的数据载荷 (Payload)
    payload = {
        "model": MODEL_NAME,
        "messages":[
            {"role":"system", "content": system_prompt},
            {"role":"user", "content": user_command}
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

        # 1. (调试) 打印出Ollama返回的完整原始JSON，这对于排错至关重要！
        print("--- Ollama Raw Response ---")
        print(json.dumps(response_data, indent=2, ensure_ascii=False))
        print("--------------------------")

        # 2. (修正) 使用新的、更健壮的解析逻辑
        # 新版Ollama通常将内容直接放在 response['message']['content']
        message_content = response_data.get('message', {}).get('content')

        if not message_content:
            print("错误：在Ollama的响应中找不到'message'或'content'。")
            return None

        # Ollama返回的JSON内容本身是一个字符串，需要再次解析
        parsed_json = json.loads(message_content)
        print("✅ AI解析成功！")
        return parsed_json

    except requests.exceptions.RequestException as e:
        print(f"错误：连接Ollama API失败 -> {e}")
        print(f"请确保Ollama服务正在后台运行，并且模型 '{MODEL_NAME}' 已通过 `ollama pull` 下载。")
        return None
    except json.JSONDecodeError as e:
        # 增加对解析失败时内容的打印
        print(f"错误：AI返回的不是有效的JSON格式 -> {e}")
        print(f"收到的内容: {message_content}")
        return None

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


