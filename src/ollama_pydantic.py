# src/ollama_pydantic.py

import requests
import json
from pydantic import BaseModel, ValidationError
from typing import Type, TypeVar, Optional, List, Dict, Any

# 定义一个泛型类型变量 T，它必须是 Pydantic BaseModel 的子类
# 这让我们的函数可以返回与传入的 response_model 类型完全一致的对象，并获得IDE的类型提示
T = TypeVar("T", bound=BaseModel)


def create(
        *,
        response_model: Type[T],
        model: str = "deepseek-coder-v2:16b",
        messages: List[Dict[str, str]],
        max_retries: int = 3,
        ollama_base_url: str = "http://localhost:11434"
) -> Optional[T]:
    """
    使用 Ollama 的原生 API 生成符合 Pydantic 模型的结构化数据。

    这个函数实现了自动验证和修复的逻辑：
    1. 向 Ollama 发送请求，并强制要求返回 JSON 格式。
    2. 接收 Ollama 返回的 JSON 字符串。
    3. 尝试使用传入的 `response_model` (一个Pydantic模型) 来解析和验证该字符串。
    4. 如果验证成功，返回 Pydantic 模型实例。
    5. 如果验证失败 (无论是JSON格式错误还是数据不符合模型)，则构建一个新的提示，
       将错误信息告诉 LLM，让它根据错误修复自己的输出，然后重试。

    Args:
        response_model (Type[T]): 你期望返回的 Pydantic 模型类。
        model (str): 要使用的 Ollama 模型名称。
        messages (List[Dict[str, str]]): 发送给模型的聊天消息列表。
        max_retries (int): 在彻底失败前，允许的最大重试次数。
        ollama_base_url (str): Ollama 服务的根 URL。

    Returns:
        Optional[T]: 如果成功，返回一个经过验证的 Pydantic 模型实例；
                     如果经过所有重试后仍然失败，则返回 None。
    """
    # 使用 session 来复用TCP连接，提高效率
    session = requests.Session()
    # 复制一份原始消息，以防在重试逻辑中被修改
    original_messages = [msg for msg in messages]

    # 总尝试次数 = 1次初始尝试 + max_retries 次重试
    for attempt in range(max_retries + 1):
        # 准备发送给 Ollama 的数据体
        payload = {
            "model": model,
            "messages": messages,
            "format": "json",  # 这是Ollama原生支持的强制JSON输出模式
            "stream": False
        }

        try:
            print(f"--- [尝试 #{attempt + 1}] 正在向 Ollama 发送请求...")

            # 发起网络请求
            response = session.post(
                f"{ollama_base_url}/api/chat",
                json=payload,
                timeout=120  # 设置一个合理的超时时间
            )
            # 如果服务器返回错误状态码 (如 500)，则抛出异常
            response.raise_for_status()

            # 提取 LLM 生成的 JSON 内容字符串
            response_data = response.json()
            content_string = response_data.get('message', {}).get('content')

            if not content_string:
                raise ValueError("Ollama 返回了空内容。")

            print(f"--- [尝试 #{attempt + 1}] 已收到响应，正在使用 Pydantic 模型验证...")
            # 核心步骤：使用 Pydantic 模型从 JSON 字符串进行验证和解析
            # 如果成功，model_validate_json 会返回一个 response_model 的实例
            validated_model = response_model.model_validate_json(content_string)

            print(f"--- [尝试 #{attempt + 1}] 验证成功！")
            return validated_model  # 成功，立即返回结果

        except (requests.RequestException, json.JSONDecodeError, ValidationError, ValueError) as e:
            # 捕获所有可能的错误：网络问题、JSON解析问题、Pydantic验证问题、空内容问题
            print(f"--- [尝试 #{attempt + 1}] 失败。原因: {type(e).__name__}: {e}")

            # 如果这不是最后一次尝试，则准备重试
            if attempt < max_retries:
                print("--- 准备重试，正在构建修复提示...")
                # 核心修复逻辑：构建一个新消息，告诉LLM它犯了什么错
                repair_prompt = f"""
你的上一次输出未能通过验证。错误信息如下：
---
{e}
---
请仔细阅读以上错误，并严格按照 Pydantic 模型的要求修正你的输出。
请只输出修正后的、完整且有效的JSON对象，不要包含任何额外的解释或文本。
这是你上次需要完成的原始指令：
---
{original_messages[-1]['content']}
---
"""
                # 在消息历史中追加这个修复提示
                messages.append({"role": "user", "content": repair_prompt})
            else:
                # 所有尝试都已用尽
                print("--- 已达到最大重试次数，彻底失败。")
                return None

    return None  # 理论上不会执行到这里，但作为安全保障