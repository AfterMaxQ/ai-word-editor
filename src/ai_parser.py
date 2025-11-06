# src/ai_parser.py

from http.client import responses

import requests
import json

# 定义Ollama API的地址和模型名称
OLLAMA_API_URL = "http://localhost:11434/api/chat"
MODEL_NAME = "llama3:8b"
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

        # Ollama返回的JSON内容是一个字符串，需要再次解析
        json_content_str = response_data.get('message', {}).get('content', '{}')

        parsed_json = json.loads(json_content_str)
        print("✅ AI解析成功！")
        return parsed_json
    except requests.exceptions.RequestException as e:
        print(f"错误：连接Ollama API失败 -> {e}")
        print("请确保Ollama服务正在后台运行，并且已通过 `ollama run llama3:8b` 下载了模型。")
        return None
    except json.JSONDecodeError:
        print("错误：AI返回的不是有效的JSON格式。")
        print(f"收到的内容: {json_content_str}")
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


