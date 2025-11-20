# src/ollama_pydantic.py

import json
from typing import Any, Dict, List, Optional, Type, TypeVar

import httpx
from pydantic import BaseModel, ValidationError

T = TypeVar("T", bound=BaseModel)


async def create(
        *,
        response_model: Type[T],
        model: str = "qwen2.5-coder:14b",
        prompt: str,
        max_retries: int = 2,
        ollama_base_url: str = "http://localhost:11434"
) -> tuple[Optional[T], str]:
    """
    【已更新】使用Ollama API异步生成符合Pydantic模型的结构化数据。
    此版本将temperature设置为0，以最大程度地减少AI幻觉，提高输出的确定性。
    """
    log_chunks = []

    messages = [{"role": "system", "content": prompt}]

    for attempt in range(max_retries + 1):
        payload = {
            "model": model,
            "messages": messages,
            "format": "json",
            "stream": False,
            "options": {
                "temperature": 0.0  # 【核心优化】抑制随机性
            }
        }

        log_msg = f"\n[AI-PYDANTIC-CALL] Attempt {attempt + 1}/{max_retries + 1} using model '{model}'"
        log_chunks.append(log_msg)
        print(log_msg)

        try:
            async with httpx.AsyncClient(timeout=180.0) as client:
                response = await client.post(
                    f"{ollama_base_url}/api/chat",
                    json=payload
                )
                response.raise_for_status()

            response_data = response.json()
            content_string = response_data.get('message', {}).get('content')

            if not content_string:
                raise ValueError("Ollama returned an empty content string.")

            log_chunks.append(f"  [RAW-JSON]\n{content_string}\n")

            validated_model = response_model.model_validate_json(content_string)

            success_msg = f"  [SUCCESS] Pydantic model validation successful on attempt {attempt + 1}!"
            log_chunks.append(success_msg)
            print(success_msg)

            return validated_model, "\n".join(log_chunks)

        except (httpx.RequestError, json.JSONDecodeError, ValidationError, ValueError) as e:
            error_details = str(e).replace('\n', '\n    ')
            err_msg = f"  [ERROR] Attempt {attempt + 1} failed. Reason: {type(e).__name__}\n    Details: {error_details}"
            log_chunks.append(err_msg)
            print(err_msg)

            if attempt < max_retries:
                repair_prompt = f"""
Your last JSON output failed validation against the required schema. This is a structural error.
Here is the precise validation error message:
---
{e}
---
Do not apologize or explain. Your task is to analyze this error, review your previous incorrect output, and generate a new, complete, and valid JSON object that strictly fixes this schema violation. The original instructions are in the system prompt.
"""
                if content_string:
                    messages.append({"role": "assistant", "content": content_string})
                messages.append({"role": "user", "content": repair_prompt})
            else:
                fail_msg = "\n[AI-PYDANTIC-CALL] All attempts failed. Max retries reached."
                log_chunks.append(fail_msg)
                print(fail_msg)
                return None, "\n".join(log_chunks)

    return None, "\n".join(log_chunks)