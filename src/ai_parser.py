# src/ai_parser.py

import asyncio
import json
import re
import traceback
from typing import Callable, Optional, Dict, Any, List, cast, Literal, Union
from inspect import signature, Parameter
from difflib import get_close_matches
import networkx as nx
import yaml
from pydantic import ValidationError

from .doc_builder import DocumentBuilder, ParagraphProxy, TableProxy
from .ollama_pydantic import create as ollama_create
from .schemas import CommandBlockContainer, ToolCallContainer, LogicalCommandBlock, ToolCall, DocumentModel


"""qwen3-coder:480b-cloud                                e30e45586389    -         42 minutes ago
phi3:mini                                             4f2222927938    2.2 GB    20 hours ago
deepseek-coder:latest                                 3ddd2d3fc8d2    776 MB    21 hours ago
smallthinker:3b                                       945eb1864589    3.6 GB    21 hours ago
my-qwen2.5-coder:3b-editor                            3627b3e4392e    1.9 GB    29 hours ago
qwen2.5-coder:3b                                      f72c60cabf62    1.9 GB    31 hours ago
my-qwen2.5-coder:1.5b-editor                          43a9af82069b    986 MB    31 hours ago
my-qwen2.5-coder:0.5b-editor                          f06ee692be06    397 MB    31 hours ago
qwen2.5-coder:0.5b                                    4ff64a7f502a    397 MB    31 hours ago
qwen2.5-coder:1.5b                                    d7372fd82851    986 MB    31 hours ago
my-qwen3:4b-editor-                                   0ac0a525a906    2.5 GB    31 hours ago
my-qwen-editor-q4:latest                              bc76e8c03f5a    4.8 GB    31 hours ago
qwen2.5:7b-instruct-q4_K_M                            845dbda0ea48    4.7 GB    2 days ago
doc-ai-compiler-deepseek-r1:1.5b_base_v1              ba69d2d602c2    4.7 GB    3 days ago
doc-ai-compiler-qwen2.5-coder:7b-v1                   ba69d2d602c2    4.7 GB    3 days ago
deepseek-r1:1.5b                                      e0979632db5a    1.1 GB    3 days ago
doc-ai-compiler-qwen2.5-coder:14b_base-v1             38c5edf6d9fd    9.0 GB    3 days ago
doc-ai-compiler-deepseekcoderv2_16b_base-v1:latest    7d6c0f1ae496    8.9 GB    3 days ago
qwen3-vl:2b                                           0635d9d857d4    1.9 GB    5 days ago
deepseek-coder-v2:16b                                 63fb193b3a9b    8.9 GB    7 days ago
qwen3-vl:235b-cloud                                   86b3322ec200    -         7 days ago
gpt-oss:120b-cloud                                    569662207105    -         7 days ago
qwen3-vl:8b                                           901cae732162    6.1 GB    7 days ago
qwen3-vl:4b                                           1343d82ebee3    3.3 GB    7 days ago
qwen2.5-coder:14b                                     9ec8897f747e    9.0 GB    9 days ago
qwen2.5-coder:7b                                      dae161e27b0e    4.7 GB    9 days ago
deepseek-r1:14b                                       c333b7232bdb    9.0 GB    9 days ago
deepseek-r1:7b                                        755ced02ce7b    4.7 GB    9 days ago"""


# --- 静态配置 ---
try:
    with open('config.yaml', 'r', encoding='utf-8') as f:
        MODEL_CONFIG = yaml.safe_load(f)['models']
    print("✅ config.yaml 加载成功。")
except (FileNotFoundError, KeyError, yaml.YAMLError):
    print("⚠️ config.yaml 未找到或格式不正确，使用硬编码 MODEL_CONFIG。")
    MODEL_CONFIG = {
        "planner": "deepseek-coder-v2:16b",
        "generator": "deepseek-coder-v2:16b",
        "corrector": "deepseek-coder-v2:16b"
    }

POLISHING_PROMPT_FILE = "prompts/prompt_for_polishing.txt"
TOOL_CALL_PROMPT_FILE = "prompts/prompt_for_tool_calls.txt"
CORRECTION_PROMPT_FILE = "prompts/prompt_for_correction.txt"
POSTPROCESS_PATTERN = re.compile(
    r'(\$[^$]+\$|\{\{footnote:[^}]+\}\}|\{\{endnote:[^}]+\}\}|\{\{cross_reference:[^}]+\}\})')


# --- 辅助函数 ---

def _normalize_and_validate_tool_call(builder: DocumentBuilder, call: ToolCall) -> tuple[
    Optional[Callable], dict, Optional[str]]:
    """
    【已重构 v2】验证单个工具调用的名称、参数和类型。
    此版本增加了智能错误反馈，当工具名称无效时，会尝试提供一个 "Did you mean...?" 建议，
    以帮助 AI 自我修正。

    Args:
        builder (DocumentBuilder): DocumentBuilder的实例，用于检查方法是否存在。
        call (ToolCall): Pydantic模型，表示一个工具调用。

    Returns:
        tuple[Optional[Callable], dict, Optional[str]]:
            返回一个元组，包含可调用的方法、经过验证的参数字典和一个错误消息（如果验证失败）。
    """
    tool_name_orig = call.tool_name
    tool_input = call.tool_input.copy()

    tool_aliases = {
        'create_paragraph': 'add_paragraph', 'insert_paragraph': 'add_paragraph',
        'create_table': 'add_table', 'insert_table': 'add_table',
        'create_list': 'add_list', 'insert_list': 'add_list',
        'set_page_setup': 'set_page_orientation'
    }
    valid_tools = {
        'add_paragraph', 'add_list', 'add_table', 'update_properties', 'update_table',
        'set_page_orientation', 'set_margins_cm', 'define_numbering',
        'add_header', 'add_footer', 'add_page_break', 'add_toc', 'no_op'
    }

    normalized_name = tool_name_orig.strip().lower()
    if normalized_name in valid_tools:
        tool_name = normalized_name
    elif normalized_name in tool_aliases:
        tool_name = tool_aliases[normalized_name]
    else:
        error_msg = f"Unknown tool '{tool_name_orig}'."
        suggestions = get_close_matches(normalized_name, list(valid_tools), n=1, cutoff=0.7)
        if suggestions:
            error_msg += f" Did you mean '{suggestions[0]}'?"
        return None, {}, error_msg

    if tool_name in ['no_op', 'update_properties']:
        return None, tool_input, None

    method = getattr(builder, tool_name, None)
    if not callable(method):
        return None, {}, f"Tool '{tool_name}' is valid but not implemented in DocumentBuilder."

    try:
        sig = signature(method)
        for param_name, param in sig.parameters.items():
            if param_name not in tool_input and param.default is Parameter.empty:
                provided_keys = list(tool_input.keys())
                error_msg = (
                    f"Missing required argument '{param_name}' for tool '{tool_name}'. "
                    f"Provided arguments were: {provided_keys}."
                )
                return None, {}, error_msg

        allowed_args = set(sig.parameters.keys())
    except Exception as e:
        return None, {}, f"Argument validation failed for tool '{tool_name}': {e}"


def _execute_tool_calls(builder: DocumentBuilder, tool_calls: List[ToolCall]) -> None:
    """
    按顺序执行一系列经过验证的工具调用。

    Args:
        builder (DocumentBuilder): DocumentBuilder的实例，用于执行调用。
        tool_calls (List[ToolCall]): 要执行的工具调用列表。
    """
    last_proxy: Optional[Union[ParagraphProxy, TableProxy]] = None
    for call in tool_calls:
        method_to_call, final_input, error = _normalize_and_validate_tool_call(builder, call)

        if error:
            print(f"Warning: Tool call validation failed. {error}. Skipping.")
            builder.add_paragraph(text=f"[AI Tool Call Failed: {error}]")
            continue

        if call.tool_name == 'no_op':
            print(f"Info: Skipping no_op tool. Reason: {call.tool_input.get('reason', 'N/A')}")
            continue

        # update_properties 是一个特殊情况，它作用于上一个元素代理
        if call.tool_name == 'update_properties':
            if isinstance(last_proxy, (ParagraphProxy, TableProxy)):
                props = call.tool_input.get('properties', {})
                if 'alignment' in props:
                    last_proxy.set_alignment(cast(Literal['left', 'center', 'right'], props['alignment']))
                if isinstance(last_proxy, ParagraphProxy) and 'bookmark_id' in props:
                    last_proxy.bookmark(props['bookmark_id'])
            else:
                print(f"Warning: 'update_properties' called without a valid preceding element proxy.")
            continue

        try:
            result = method_to_call(**final_input)
            if isinstance(result, (ParagraphProxy, TableProxy)):
                last_proxy = result
            else:
                last_proxy = None  # 重置代理
        except Exception as e:
            print(f"Error during execution of tool '{call.tool_name}': {e}")
            builder.add_paragraph(text=f"[AI Tool Execution Failed for '{call.tool_name}': {e}]")
            last_proxy = None


def _post_process_and_resolve_state(document_state: Dict) -> Dict:
    """
    【已重构 v3】对最终的文档状态JSON进行后处理。
    此版本实现了正确的内联元素处理，将包含特殊语法的段落转换为
    一个带有多类型“运行”(run)的段落元素，而不是错误地将其拆分为多个块级元素。

    Args:
        document_state (Dict): 由DocumentBuilder生成的原始文档状态。

    Returns:
        Dict: 经过处理和解析的最终文档状态。
    """
    PREFIX_FOOTNOTE = '{{footnote:'
    PREFIX_ENDNOTE = '{{endnote:'
    PREFIX_CROSS_REF = '{{cross_reference:'

    new_document_state = document_state.copy()
    for section in new_document_state.get('sections', []):
        final_elements = []
        for element in section.get('elements', []):
            if element.get('type') != 'paragraph' or 'text' not in element:
                final_elements.append(element)
                continue

            text_to_process = element.get('text', "")
            if not text_to_process:
                # Keep empty paragraphs that might have properties (e.g., for spacing)
                if element.get('properties'):
                    final_elements.append(element)
                continue

            if 'properties' in element and 'bookmark_id' in element['properties']:
                text_to_process = re.sub(r'\{\{bookmark:[^}]+\}\}', '', text_to_process).strip()

            parts = [p for p in POSTPROCESS_PATTERN.split(text_to_process) if p]

            if len(parts) <= 1:
                # No special syntax found, just append the original element
                final_elements.append(element)
                continue

            # --- [NEW LOGIC] ---
            # If special syntax is found, create a SINGLE paragraph with a content array of runs.
            new_paragraph = {k: v for k, v in element.items() if k != 'text'}
            new_paragraph['content'] = []

            for part in parts:
                if not POSTPROCESS_PATTERN.fullmatch(part):
                    new_paragraph['content'].append({'type': 'text', 'text': part})
                    continue

                if part.startswith('$') and part.endswith('$'):
                    # Create a formula RUN, not a formula ELEMENT
                    new_paragraph['content'].append({'type': 'formula', 'text': part[1:-1]})
                elif part.startswith(PREFIX_FOOTNOTE):
                    content = part[len(PREFIX_FOOTNOTE):-2]
                    new_paragraph['content'].append({'type': 'footnote', 'text': content})
                elif part.startswith(PREFIX_ENDNOTE):
                    content = part[len(PREFIX_ENDNOTE):-2]
                    new_paragraph['content'].append({'type': 'endnote', 'text': content})
                elif part.startswith(PREFIX_CROSS_REF):
                    content = part[len(PREFIX_CROSS_REF):-2]
                    new_paragraph['content'].append({'type': 'cross_reference', 'target_bookmark': content})

            if new_paragraph.get('content'):
                final_elements.append(new_paragraph)

        section['elements'] = final_elements
    return new_document_state


async def parse_natural_language_to_json(
        user_command: str,
        log_callback: Optional[Callable[[str], None]] = None
) -> tuple[dict | None, str]:
    """
    【核心重构】采用简化的顺序Agent流，将自然语言转换为文档状态JSON。
    此函数协调规划、排序、生成和执行的整个过程。

    Args:
        user_command (str): 用户的原始自然语言指令。
        log_callback (Optional[Callable[[str], None]]): 用于流式日志记录的回调函数。

    Returns:
        tuple[dict | None, str]: 最终文档状态JSON和完整日志。
    """
    full_log = []

    def log(message: str):
        print(message, flush=True)
        full_log.append(message)
        if log_callback:
            try:
                log_callback(message)
            except Exception as e:
                # 防止日志回调中的错误中断主流程
                print(f"[LOGGING_WARN] Log callback failed: {e}")

    try:
        with open(POLISHING_PROMPT_FILE, 'r', encoding='utf-8') as f:
            polishing_prompt_template = f.read()
        with open(TOOL_CALL_PROMPT_FILE, 'r', encoding='utf-8') as f:
            tool_call_prompt_template = f.read()
        with open(CORRECTION_PROMPT_FILE, 'r', encoding='utf-8') as f:
            correction_prompt_template = f.read()
    except FileNotFoundError as e:
        log(f"❌ 致命错误: 未找到提示词文件: {e}")
        return None, "\n".join(full_log)

    # --- 阶段 1: Planner Agent - 指令规整与分解 ---
    log("\n" + "=" * 20 + " 1. Agent 1: 指令规整与分解 " + "=" * 20)
    command_block_container, agent1_log = await ollama_create(
        response_model=CommandBlockContainer,
        prompt=polishing_prompt_template.format(command=user_command),
        model=MODEL_CONFIG["planner"]
    )
    log(agent1_log)

    if not command_block_container:
        log("❌ 阶段一失败：未能生成有效的逻辑指令块。")
        return None, "\n".join(full_log)

    log(f"✅ 阶段一成功：指令被分解为 {len(command_block_container.command_blocks)} 个逻辑块。")

    # --- 阶段 1.5: Sorter - 依赖排序 (with ID-to-Index Conversion) ---
    try:
        blocks = command_block_container.command_blocks
        num_blocks = len(blocks)

        id_to_index_map: Dict[str, int] = {}
        for i, block in enumerate(blocks):
            if block.id in id_to_index_map:
                raise ValueError(f"AI generated duplicate block ID: '{block.id}'. Aborting.")
            id_to_index_map[block.id] = i
        log(f"✅ 块ID到索引的映射构建成功: {id_to_index_map}")

        graph = nx.DiGraph()
        for i in range(num_blocks):
            graph.add_node(i)

        # Step B: Convert dependency IDs to indices and build the graph.
        for i, block in enumerate(blocks):
            dependency_ids = block.dependencies or []
            for dep_id in dependency_ids:
                if dep_id not in id_to_index_map:
                    raise ValueError(f"Block '{block.id}' has an invalid dependency ID: '{dep_id}'. It does not exist.")

                dep_index = id_to_index_map[dep_id]
                graph.add_edge(dep_index, i)

        if not nx.is_directed_acyclic_graph(graph):
            # This can still happen if the AI creates a logical loop (A -> B, B -> A)
            # but it eliminates self-dependency errors.
            raise ValueError("检测到循环依赖，无法确定执行顺序。")

        sorted_indices = list(nx.topological_sort(graph))
        sorted_blocks = [blocks[i] for i in sorted_indices]
        log(f"✅ 依赖排序成功：执行顺序为 {sorted_indices}")
    except (nx.NetworkXError, ValueError) as e:
        log(f"❌ 依赖关系图错误: {e}")
        return None, "\n".join(full_log)
    builder = DocumentBuilder()

    # --- 阶段 2: Generator Agent - 并发生成工具调用 ---
    log("\n" + "=" * 20 + " 2. Agent 2: 并发生成工具调用 " + "=" * 20)

    async def generate_calls_for_block(block: LogicalCommandBlock, index: int):
        """为单个块生成工具调用的异步任务。"""
        log(f"  - 开始处理块 {index}: '{block.primary_command[:40]}...'")
        prompt = tool_call_prompt_template.format(command_block=block.model_dump_json(indent=2))
        container, agent_log = await ollama_create(
            response_model=ToolCallContainer,
            prompt=prompt,
            model=MODEL_CONFIG["generator"]
        )
        # 将日志和结果关联起来
        return container, agent_log, block

    tasks = [generate_calls_for_block(block, idx) for idx, block in enumerate(sorted_blocks)]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    # --- 阶段 2.5 & 3: Executor & Corrector - 执行与自我修正循环 ---
    log("\n" + "=" * 20 + " 3. 执行与验证工具调用 " + "=" * 20)

    for i, res in enumerate(results):
        block_info = f"块 {i} ('{sorted_blocks[i].primary_command[:30]}...')"
        if isinstance(res, Exception):
            log(f"❌ {block_info} 的工具调用生成失败: {res}")
            # 应用一个 fallback 策略，生成一个错误段落
            tool_call_container = ToolCallContainer(calls=[
                ToolCall(tool_name="add_paragraph",
                         tool_input={"text": f"[AI Generation Failed for: {sorted_blocks[i].primary_command}]"})
            ])
            agent2_log = f"[Fallback Applied for Block {i}]"
            block_context = sorted_blocks[i]
        else:
            tool_call_container, agent2_log, block_context = res

        log(agent2_log)
        if not tool_call_container:
            log(f"⚠️ {block_info} 未能生成工具调用，应用 fallback。")
            tool_call_container = ToolCallContainer(calls=[
                ToolCall(tool_name="add_paragraph",
                         tool_input={"text": f"[AI Generation Failed for: {block_context.primary_command}]"})
            ])

        log(f"✅ 已为 {block_info} 生成 {len(tool_call_container.calls)} 个工具调用。")

        # 执行与修正循环
        max_exec_retries = 2
        for attempt in range(max_exec_retries + 1):
            try:
                log(f"\n  --- {block_info} 执行尝试 #{attempt + 1} ---")
                _execute_tool_calls(builder, tool_call_container.calls)

                # 临时获取状态以进行Pydantic验证
                temp_doc_state = builder.get_document_state()
                DocumentModel.model_validate(temp_doc_state)

                log(f"  ✅ {block_info} 执行成功并通过Schema验证。")
                break  # 成功，退出循环
            except (ValidationError, Exception) as e:
                error_traceback = traceback.format_exc()
                log(f"  ❌ {block_info} 执行/验证失败: {type(e).__name__}: {e}")

                if attempt < max_exec_retries:
                    log("    🔧 启动执行层自我修正...")
                    correction_prompt = correction_prompt_template.format(
                        command_block=block_context.model_dump_json(indent=2),
                        failed_tool_calls=tool_call_container.model_dump_json(indent=2),
                        error_message=error_traceback
                    )
                    corrected_container, correction_log = await ollama_create(
                        response_model=ToolCallContainer,
                        prompt=correction_prompt,
                        model=MODEL_CONFIG["corrector"]
                    )
                    log(correction_log)

                    if corrected_container:
                        log("    ✅ 已生成修正后的工具调用，准备重试。")
                        tool_call_container = corrected_container  # 更新为修正后的调用
                    else:
                        log("    ❌ 修正失败，中止此块的执行。")
                        break  # 修正失败，退出循环
                else:
                    log(f"  ❌ {block_info} 已达到最大执行重试次数，中止。")
                    break  # 达到最大次数，退出循环

    # --- 阶段 4: 后处理 ---
    log("\n" + "=" * 20 + " 4. 文档状态后处理 " + "=" * 20)
    final_doc_state = builder.get_document_state()
    processed_final_state = _post_process_and_resolve_state(final_doc_state)
    log("✅ 后处理完成。")

    final_json_log = json.dumps(processed_final_state, indent=2, ensure_ascii=False)
    log(f"\n[最终JSON] 最终聚合生成的JSON对象如下:\n{final_json_log}")

    return processed_final_state, "\n".join(full_log)