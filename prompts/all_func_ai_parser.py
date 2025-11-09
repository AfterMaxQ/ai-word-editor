# src/ai_parser.py

def translate_latex_to_omml_llm(latex_string: str) -> str | None:
    """
    使用大语言模型（LLM）将单个LaTeX字符串转换为OMML XML。这是 `get_formula_xml_and_placeholder` 函数的一个备用方案。

    Args:
        latex_string (str): 需要转换的LaTeX公式字符串。

    Returns:
        str | None: 如果成功，返回包含OMML的XML字符串；否则返回 `None`。

    数据处理过程:
        1. 读取专为LaTeX转换设计的系统提示词。
        2. 将用户的LaTeX字符串格式化为一个用户提示。
        3. 构造请求体，调用Ollama API。
        4. 接收LLM的响应，并从中提取内容字符串。
        5. 对返回的字符串进行清理（例如，移除 ` ```xml ` 这样的markdown标记）。
        6. 使用 `lxml` 验证清理后的字符串是否是有效的XML。
        7. 如果有效，则返回该字符串；否则打印错误并返回 `None`。
    """

def split_command_into_chunks(user_command: str, max_chunks: int = 30) -> tuple[list[str], str]:
    """
    将用户输入的长指令分割成更小的、符合逻辑的、易于LLM处理的任务块。

    Args:
        user_command (str): 用户的完整自然语言指令。
        max_chunks (int): 允许的最大任务块数量。

    Returns:
        tuple[list[str], str]: 返回一个元组，第一个元素是分割后的任务块字符串列表，第二个元素是分割过程的日志。

    数据处理过程（动态分片）:
        1.  **粗分**: 首先使用一个或多个连续的空行作为分隔符，将整个指令分割成逻辑单元。
        2.  **判断**: 如果分割后的单元数小于或等于 `max_chunks`，则直接返回这些单元，无需进一步处理。
        3.  **精合**: 如果单元数过多，则进行智能合并。计算出平均每个最终任务块应该包含多少个逻辑单元 (`ceil(total_units / max_chunks)`)。
        4.  根据计算出的步长，将逻辑单元列表进行分组，并将每组内的单元合并成一个最终的任务块。
    """

def parse_natural_language_to_json(user_command: str) -> tuple[dict | None, str]:
    """
    **[核心函数 - AI解析]** 负责将用户的完整自然语言指令转换为最终的、结构化的JSON文档数据。

    Args:
        user_command (str): 用户的完整自然语言指令。

    Returns:
        tuple[dict | None, str]: 返回一个元组，第一个元素是代表整个文档的聚合后的JSON字典，如果失败则为 `None`；第二个元素是详细的处理日志。

    数据处理过程:
        1.  **分块**: 调用 `split_command_into_chunks` 将用户指令分解为任务块。
        2.  **初始化**: 创建一个空的 `aggregated_document_data` 字典，用于累积所有任务块的结果。
        3.  **循环处理**: 遍历每一个任务块。
            a. 为当前任务块构建一个特定的用户提示，该提示包含了任务块内容、上下文信息（如这是第几个块），并强调了严格遵守JSON格式和“空指令返回空JSON”的规则。
            b. **调用AI**: 调用 `ollama_pydantic.create` 函数，将当前任务块的提示发送给LLM，并指定 `DocumentModel` 作为期望的返回数据结构。这一步利用了 `create` 函数的自愈能力。
            c. 如果 `create` 返回 `None`，说明AI在多次尝试后依然失败，整个流程中止。
            d. **聚合**: 如果成功，将返回的Pydantic模型实例转换为字典。然后，将这个字典中的 `sections` 列表追加到 `aggregated_document_data` 的 `sections` 中，并将 `page_setup` 信息更新到主字典中。
        4.  **完成**: 循环结束后，返回聚合了所有任务块结果的 `aggregated_document_data` 和完整的日志。
    """