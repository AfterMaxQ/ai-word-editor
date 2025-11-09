# src/ollama_pydantic.py

def create(*, response_model: Type[T], model: str, messages: List[Dict[str, str]], max_retries: int, ollama_base_url: str) -> Optional[T]:
    """
    **[核心函数 - AI交互]** 使用Ollama API生成符合指定Pydantic模型的结构化JSON数据，并内置了自动验证和修复的重试逻辑。

    Args:
        response_model (Type[T]): 期望API返回的JSON所对应的Pydantic模型类。
        model (str): 要使用的Ollama模型名称。
        messages (List[Dict[str, str]]): 发送给模型的聊天消息列表。
        max_retries (int): 在彻底失败前，允许的最大重试次数。
        ollama_base_url (str): Ollama服务的根URL。

    Returns:
        Optional[T]: 如果成功，返回一个经过验证的、类型为 `response_model` 的Pydantic模型实例；如果经过所有重试后仍然失败，则返回 `None`。

    数据处理过程（自愈式调用循环）:
        1.  进入一个循环，总共尝试 `max_retries + 1` 次。
        2.  **请求**: 向Ollama的 `/api/chat` 端点发送POST请求，请求体中包含模型、消息历史，并设置 `format: "json"` 来强制Ollama输出JSON。
        3.  **接收与验证**:
            a. 接收Ollama返回的JSON字符串。
            b. 关键步骤：使用 `response_model.model_validate_json(content_string)` 尝试解析并验证这个字符串。
            c. 如果验证成功，函数立即返回Pydantic模型实例，循环结束。
        4.  **失败与修复**:
            a. 如果在请求或验证过程中发生任何错误（网络错误、JSON解析错误、Pydantic验证失败），则捕获异常。
            b. **构建修复提示**: 创建一个新的用户提示，该提示包含捕获到的错误信息和原始指令，明确要求LLM根据错误修正其上一次的输出。
            c. **重试**: 将这个修复提示追加到 `messages` 列表中，然后继续下一次循环，这样LLM在下一次生成时就会看到自己的错误并尝试改正。
        5.  如果循环结束后仍未成功，则返回 `None`。
    """