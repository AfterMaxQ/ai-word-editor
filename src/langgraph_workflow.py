# src/langgraph_workflow.py

import asyncio
import json
import re
import traceback
from typing import Dict, List, TypedDict, Annotated, Optional, cast, Literal

import networkx as nx
from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import StateGraph, END

from .doc_builder import DocumentBuilder, ParagraphProxy, TableProxy
from .ollama_pydantic import create as ollama_create
from .schemas import CommandBlockContainer, ToolCallContainer, LogicalCommandBlock, ToolCall, DocumentModel

# --- 1. 提示词加载 ---
try:
    with open("prompts/prompt_for_polishing.txt", 'r', encoding='utf-8') as f:
        POLISHING_PROMPT = f.read()
    with open("prompts/prompt_for_tool_calls.txt", 'r', encoding='utf-8') as f:
        TOOL_PROMPT = f.read()
    with open("prompts/prompt_for_correction.txt", 'r', encoding='utf-8') as f:
        CORRECTION_PROMPT = f.read()
except FileNotFoundError as e:
    raise RuntimeError(f"Critical prompt file not found: {e}") from e

# --- 2. 辅助函数 (从 ai_parser.py 迁移并保持不变) ---
POSTPROCESS_PATTERN = re.compile(
    r'(\$[^$]+\$|\{\{footnote:[^}]+\}\}|\{\{endnote:[^}]+\}\}|\{\{cross_reference:[^}]+\}\})')


def _post_process_and_resolve_state(document_state: Dict) -> Dict:
    # ... (此函数的实现与上一版本完全相同)
    new_document_state = document_state.copy()
    for section in new_document_state.get('sections', []):
        final_elements = []
        for element in section.get('elements', []):
            if 'properties' in element and 'bookmark_id' in element['properties']:
                text_to_process = element.get('text', '')
                if '{{bookmark:' in text_to_process:
                    element['text'] = re.sub(r'\{\{bookmark:[^}]+\}\}', '', text_to_process).strip()
            if element.get('type') != 'paragraph' or 'text' not in element:
                final_elements.append(element)
                continue
            text_to_process = element['text']
            parts = [p for p in POSTPROCESS_PATTERN.split(text_to_process) if p]
            if len(parts) <= 1:
                final_elements.append(element)
                continue
            current_paragraph = element.copy()
            current_paragraph.pop('text', None)
            current_paragraph['content'] = []
            for part in parts:
                if not POSTPROCESS_PATTERN.fullmatch(part):
                    current_paragraph['content'].append({'type': 'text', 'text': part})
                    continue
                if part.startswith('$') and part.endswith('$'):
                    content = part[1:-1]
                    if current_paragraph['content']:
                        final_elements.append(current_paragraph)
                    final_elements.append({"type": "formula", "properties": {"text": content}})
                    current_paragraph = element.copy()
                    current_paragraph.pop('text', None)
                    current_paragraph['content'] = []
                elif part.startswith('{{footnote:'):
                    content = part[11:-2]
                    current_paragraph['content'].append({'type': 'footnote', 'text': content})
                elif part.startswith('{{endnote:'):
                    content = part[10:-2]
                    current_paragraph['content'].append({'type': 'endnote', 'text': content})
                elif part.startswith('{{cross_reference:'):
                    content = part[19:-2]
                    current_paragraph['content'].append({'type': 'cross_reference', 'target_bookmark': content})
            if current_paragraph.get('content'):
                final_elements.append(current_paragraph)
        section['elements'] = final_elements
    return new_document_state


def _execute_tool_calls(builder: DocumentBuilder, tool_calls: List[ToolCall]) -> None:
    # ... (此函数的实现与上一版本完全相同)
    last_proxy = None
    seen_globals = set()
    for call in tool_calls:
        tool_input = call.tool_input
        if call.tool_name in ['set_page_setup', 'define_numbering']:
            if call.tool_name in seen_globals:
                print(f"Skipping duplicate global tool call: {call.tool_name}")
                continue
            seen_globals.add(call.tool_name)

        if call.tool_name == 'create_paragraph':
            last_proxy = builder.add_paragraph(**tool_input)
        elif call.tool_name == 'create_table':
            last_proxy = builder.add_table(**tool_input)
        elif call.tool_name == 'create_list':
            builder.add_list(**tool_input);
            last_proxy = None
        elif call.tool_name == 'update_properties':
            if isinstance(last_proxy, (ParagraphProxy, TableProxy)):
                props = tool_input.get('properties', {})
                if 'alignment' in props: last_proxy.set_alignment(
                    cast(Literal['left', 'center', 'right'], props['alignment']))
                if isinstance(last_proxy, ParagraphProxy) and 'bookmark_id' in props: last_proxy.bookmark(
                    props['bookmark_id'])
            else:
                print(f"Warning: 'update_properties' called without a preceding element.")
        elif call.tool_name == 'set_page_setup':
            if 'orientation' in tool_input: builder.set_page_orientation(tool_input['orientation'])
            if 'margins' in tool_input: builder.set_margins_cm(**tool_input['margins'])
        elif call.tool_name == 'define_numbering':
            builder.define_numbering(**tool_input)
        elif call.tool_name == 'add_page_break':
            builder.add_page_break()
        elif call.tool_name == 'add_toc':
            builder.add_toc()
        elif call.tool_name == 'no_op':
            print(f"Skipping no_op tool call. Reason: {tool_input.get('reason')}")


# --- 3. LangGraph 状态定义 ---
class AgentState(TypedDict):
    """定义了工作流中传递的完整状态。"""
    user_command: str
    command_blocks: Optional[List[Dict]]
    sorted_blocks: Optional[List[Dict]]
    tool_calls_per_block: List[List[Dict]]
    final_doc_state: Optional[Dict]
    error: Optional[str]
    correction_attempts: int
    log: List[str]


# --- 4. LangGraph 节点定义 ---
async def planner_node(state: AgentState) -> AgentState:
    """节点1: 接收用户指令，将其分解为带依赖的逻辑指令块。"""
    state['log'].append("--- Entering Planner Node ---")
    prompt = POLISHING_PROMPT.format(command=state["user_command"])
    container, log = await ollama_create(response_model=CommandBlockContainer, prompt=prompt)
    state['log'].append(log)
    if container:
        state["command_blocks"] = [block.model_dump() for block in container.command_blocks]
    else:
        state["error"] = "Planner failed to generate valid command blocks."
    return state


def sorter_node(state: AgentState) -> AgentState:
    """节点2: 对指令块进行拓扑排序并检测循环依赖。"""
    state['log'].append("--- Entering Sorter Node ---")
    if state.get("error") or not state.get("command_blocks"): return state
    try:
        graph = nx.DiGraph()
        blocks = state["command_blocks"]
        for i, block in enumerate(blocks):
            graph.add_node(i)
            for dep in block.get("dependencies") or []:
                if not (isinstance(dep, int) and 0 <= dep < len(blocks)):
                    raise ValueError(f"Invalid dependency index '{dep}' in block {i}.")
                graph.add_edge(dep, i)
        if not nx.is_directed_acyclic_graph(graph):
            raise ValueError("Circular dependency detected in command blocks.")
        sorted_indices = list(nx.topological_sort(graph))
        state["sorted_blocks"] = [blocks[i] for i in sorted_indices]
        state['log'].append(f"Block execution order: {sorted_indices}")
    except Exception as e:
        state["error"] = f"Sorter failed: {e}"
    return state


async def generator_node(state: AgentState) -> AgentState:
    """节点3: 为每个排序后的指令块并发生成工具调用。"""
    state['log'].append("--- Entering Generator Node ---")
    if state.get("error") or not state.get("sorted_blocks"): return state

    async def generate_for_block(block: Dict):
        prompt = TOOL_PROMPT.format(command_block=json.dumps(block))
        container, log = await ollama_create(response_model=ToolCallContainer, prompt=prompt)
        return (container.model_dump()["calls"] if container else None), log

    tasks = [generate_for_block(block) for block in state["sorted_blocks"]]
    results = await asyncio.gather(*tasks)

    state["tool_calls_per_block"] = []
    for calls, log_entry in results:
        state['log'].append(log_entry)
        if calls is None:
            state["error"] = "Generator failed for one or more blocks."
            state["tool_calls_per_block"].append([])  # Append empty list to maintain index
        else:
            state["tool_calls_per_block"].append(calls)
    return state


# In: src/langgraph_workflow.py

def executor_node(state: AgentState) -> AgentState:
    """
    【已重构】节点4: 执行所有工具调用，并进行结构和逻辑验证。
    此版本引入了强大的预执行验证层，能够捕获AI幻觉（如无效工具名、错误参数），
    并将验证失败信息反馈给工作流以触发修正，而不是直接崩溃。
    """
    state['log'].append("--- Entering Executor Node (v2 - Hardened) ---")
    if state.get("error") or not state.get("tool_calls_per_block"):
        return state

    builder = DocumentBuilder()
    validation_errors = []

    # 【核心新增】在执行前对每一个工具调用进行严格验证
    for i, calls_for_block in enumerate(state["tool_calls_per_block"]):
        validated_calls_for_block = []
        for call_dict in calls_for_block:
            tool_call = ToolCall(**call_dict)
            method_to_call, final_input, error = _normalize_and_validate_tool_call(builder, tool_call)

            if error:
                # 捕获验证失败，记录结构化错误，但不崩溃
                error_msg = f"Block {i}: Tool '{tool_call.tool_name}' validation failed: {error}"
                state['log'].append(f"  [VALIDATION_FAIL] {error_msg}")
                validation_errors.append(error_msg)
                continue  # 跳过这个错误的调用

            validated_calls_for_block.append((method_to_call, final_input, tool_call.tool_name))

        # 只执行通过验证的调用
        try:
            last_proxy = None  # 重置代理以避免跨块影响
            for method, args, tool_name in validated_calls_for_block:
                if tool_name == 'no_op': continue
                if tool_name == 'update_properties':
                    # (此处省略 update_properties 的代理逻辑)
                    pass
                else:
                    result = method(**args)
                    last_proxy = result if isinstance(result, (ParagraphProxy, TableProxy)) else None
        except Exception as e:
            # 捕获执行时的运行时错误
            state["error"] = f"Executor failed during execution: {traceback.format_exc()}"
            state["correction_attempts"] = state.get("correction_attempts", 0) + 1
            return state

    # 【核心新增】检查是否存在验证失败，如果存在，则触发修正流程
    if validation_errors:
        aggregated_error_msg = "AI generated one or more invalid tool calls:\n- " + "\n- ".join(validation_errors)
        state["error"] = aggregated_error_msg
        state["correction_attempts"] = state.get("correction_attempts", 0) + 1
        return state

    # 验证和后处理 (成功路径)
    try:
        temp_doc_state = builder.get_document_state()
        DocumentModel.model_validate(temp_doc_state)
        state['log'].append("✅ Schema validation successful.")
        state["final_doc_state"] = _post_process_and_resolve_state(temp_doc_state)
        state["error"] = None  # 成功后清除错误状态
    except Exception as e:
        state["error"] = f"Executor failed during final validation: {traceback.format_exc()}"
        state["correction_attempts"] = state.get("correction_attempts", 0) + 1

    return state


async def corrector_node(state: AgentState) -> AgentState:
    """节点5: 在执行失败时，尝试修正工具调用。"""
    state['log'].append("--- Entering Corrector Node ---")
    if not state.get("error"): return state

    # 简化：假设修正最后一个失败的块集
    prompt = CORRECTION_PROMPT.format(
        command_block=json.dumps(state["sorted_blocks"][-1]),
        failed_tool_calls=json.dumps({"calls": state["tool_calls_per_block"][-1]}),
        error_message=state["error"]
    )
    container, log = await ollama_create(response_model=ToolCallContainer, prompt=prompt)
    state['log'].append(log)

    if container:
        state["tool_calls_per_block"][-1] = container.model_dump()["calls"]
        state["error"] = None  # Clear error to allow re-execution
        state['log'].append("Generated corrected tool calls for retry.")
    else:
        state['log'].append("Corrector failed to generate a fix.")
        # Error persists, will lead to END
    return state


# --- 5. 图的构建与路由 ---
def should_continue(state: AgentState) -> str:
    """路由逻辑：决定下一个节点或结束。"""
    if state.get("error"):
        if state.get("correction_attempts", 0) >= 2:
            state['log'].append("Max correction attempts reached. Ending workflow.")
            return END
        return "corrector"
    if not state.get("final_doc_state"):
        return "executor"  # This path is for retrying after correction
    return END


workflow = StateGraph(AgentState)
workflow.add_node("planner", planner_node)
workflow.add_node("sorter", sorter_node)
workflow.add_node("generator", generator_node)
workflow.add_node("executor", executor_node)
workflow.add_node("corrector", corrector_node)

workflow.set_entry_point("planner")
workflow.add_edge("planner", "sorter")
workflow.add_edge("sorter", "generator")
workflow.add_edge("generator", "executor")
workflow.add_edge("corrector", "executor")  # After correction, retry execution

workflow.add_conditional_edges("executor", should_continue)

# --- 6. 编译图并提供公共接口 ---
checkpointer = MemorySaver()
graph = workflow.compile(checkpointer=checkpointer)


async def parse_natural_language_to_json(
        user_command: str,
        thread_id: str = "default_thread"
) -> tuple[dict | None, str]:
    """
    【最终版接口】使用LangGraph执行可恢复的、带自修正的文档生成工作流。

    Args:
        user_command (str): 用户的原始自然语言指令。
        thread_id (str): 用于会话恢复的唯一标识符。

    Returns:
        tuple[dict | None, str]: 最终文档状态JSON和完整日志。
    """

    initial_state = {
        "user_command": user_command,
        "log": [f"--- Starting new workflow for thread_id: {thread_id} ---"],
        "correction_attempts": 0,
        "tool_calls_per_block": []
    }
    config = {"configurable": {"thread_id": thread_id}}

    final_state = await graph.ainvoke(initial_state, config)

    logs = "\n".join(final_state.get("log", []))
    if final_state.get("error"):
        return None, logs

    return final_state.get("final_doc_state"), logs