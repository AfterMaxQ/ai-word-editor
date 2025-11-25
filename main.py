# main.py

import asyncio
import base64
import json
import os

import httpx
from fastapi import FastAPI, File, HTTPException, UploadFile, Request, Form
from starlette.exceptions import HTTPException as StarletteHTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, Response
from pydantic import BaseModel
from starlette.requests import ClientDisconnect

from src.app_logic import generate_document_from_command

from src.doc_parser import parse_docx_to_json
from src.formatting_engine import apply_formatting

from src.doc_generator import create_document

os.environ['NO_PROXY'] = 'localhost,127.0.0.1'


class CommandRequest(BaseModel):
    command: str


class PolishRequest(BaseModel):
    text: str


app = FastAPI(
    title="AI 文档生成器 API",
    description="一个用于通过自然语言生成 Word 文档的 API",
    version="1.0.0",
)

origins = [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OLLAMA_API_URL = "http://localhost:11434/api/chat"


@app.get("/")
def read_root():
    """
    根路径，用于检查API服务是否正常运行。

    Returns:
        dict: 包含欢迎信息的字典。
    """
    return {"message": "AI 文档生成器 API 运行正常！"}


@app.post("/polish-text")
async def polish_text_endpoint(request: PolishRequest):
    """
    Receives text, polishes it using a local LLM, and returns the result.
    This endpoint is stateless and now includes a strict language-adherence constraint.
    """
    if not request.text.strip():
        raise HTTPException(status_code=400, detail="Text cannot be empty.")

    # NEW: The prompt now includes a strict language matching constraint.
    system_prompt = """You are a world-class editor and copywriter. Your task is to meticulously polish the user-provided text.
"Polishing" means improving clarity, fluency, conciseness, and correcting any grammatical or spelling errors, while preserving the original core meaning.

INTERNAL THOUGHT PROCESS (You must follow these steps before generating the output):
1.  **Analyze**: First, silently analyze the input text to identify its **language**, type (e.g., academic, technical, business, creative, casual), and original tone.
2.  **Strategize**: Based on the analysis, determine the most appropriate polishing strategy.
3.  **Execute**: Perform the polishing based on your chosen strategy.

CRITICAL OUTPUT INSTRUCTIONS:
1.  **LANGUAGE MATCH**: The output language MUST STRICTLY match the input language. DO NOT TRANSLATE. If the input is Chinese, the output MUST be Chinese. If the input is English, the output MUST be English.
2.  **FORMAT**: You MUST return ONLY the final, polished text. Do not include your analysis, any explanations, apologies, or markdown formatting like ```.
"""

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": request.text}
    ]

    payload = {
        "model": "qwen2.5-coder:14b",
        "messages": messages,
        "stream": False,
        "options": {
            "temperature": 0.7
        }
    }

    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(OLLAMA_API_URL, json=payload)
            response.raise_for_status()

        response_data = response.json()
        polished_content = response_data.get("message", {}).get("content", "").strip()

        if polished_content.startswith('"') and polished_content.endswith('"'):
            polished_content = polished_content[1:-1]

        if not polished_content:
            return {"polished_text": request.text}

        return {"polished_text": polished_content}

    except httpx.RequestError as e:
        raise HTTPException(status_code=503, detail=f"Failed to connect to Ollama service: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred: {e}")


@app.post("/generate")
async def generate_endpoint(req: Request, request: CommandRequest):
    """
    【已重构】接收用户指令，以流式响应的方式生成并返回Word文档。
    此版本实现了对大文件（Base64编码后）的分块传输，并增强了对后台任务异常的捕获。

    Args:
        req (Request): FastAPI的Request对象，用于检测连接状态。
        request (CommandRequest): 包含用户指令`command`的请求体。

    Returns:
        StreamingResponse: 一个流式响应，处理过程中可被客户端中止。
    """
    log_queue = asyncio.Queue()
    generation_task = None

    async def stream_generator():
        nonlocal generation_task

        def logger(message: str):
            try:
                log_queue.put_nowait(message)
            except asyncio.QueueFull:
                pass

        generation_task = asyncio.create_task(
            generate_document_from_command(request.command, logger)
        )

        try:
            # 阶段一：实时流式传输日志
            while not generation_task.done() or not log_queue.empty():
                if await req.is_disconnected():
                    raise ClientDisconnect()
                try:
                    log_line = await asyncio.wait_for(log_queue.get(), timeout=0.1)
                    yield f"data: {json.dumps({'type': 'log', 'content': log_line})}\n\n"
                    await asyncio.sleep(0.01)
                except asyncio.TimeoutError:
                    continue

            # 【核心修正】在获取结果前，显式检查后台任务是否因异常而终止
            if generation_task.done() and generation_task.exception():
                # 重新抛出异常，以便被外层except块捕获并发送错误信息
                raise generation_task.exception()

            # 阶段二：获取并分块发送最终结果
            document_bytes, json_str, full_log = await generation_task

            if document_bytes:
                yield f"data: {json.dumps({'type': 'full_log', 'content': full_log or ''})}\n\n"
                yield f"data: {json.dumps({'type': 'final_json', 'content': json_str or ''})}\n\n"

                encoded_string = base64.b64encode(document_bytes).decode('utf-8')
                CHUNK_SIZE = 32 * 1024

                for i in range(0, len(encoded_string), CHUNK_SIZE):
                    chunk = encoded_string[i:i + CHUNK_SIZE]
                    yield f"data: {json.dumps({'type': 'file_chunk', 'content': chunk})}\n\n"
                    await asyncio.sleep(0.01)

                file_metadata = {
                    "file_name": "generated_document.docx",
                    "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
                yield f"data: {json.dumps({'type': 'file_end', 'content': file_metadata})}\n\n"

            else:
                error_payload = {"type": "error", "content": full_log or "文档生成失败，未提供日志信息。"}
                yield f"data: {json.dumps(error_payload)}\n\n"

        except (ClientDisconnect, asyncio.CancelledError):
            print("[CANCEL] 流式生成器已捕获取消信号。")
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"[ERROR] 流式生成器发生未捕获的异常: {e}")
            error_message = f"在文档生成过程中发生严重错误: {type(e).__name__}: {e}"
            yield f"data: {json.dumps({'type': 'error', 'content': error_message})}\n\n"
        finally:
            if generation_task and not generation_task.done():
                generation_task.cancel()
                print("[CANCEL] 后台生成任务已被强制取消。")
            if generation_task:
                try:
                    await generation_task
                except asyncio.CancelledError:
                    print("[CANCEL] 任务取消已确认。")

    return StreamingResponse(stream_generator(), media_type="text/event-stream")


@app.post("/recognize-formula")
async def recognize_formula_endpoint(file: UploadFile = File(...)):
    """
    接收上传的公式图片，识别并返回其LaTeX表示。

    Args:
        file (UploadFile): 用户上传的图片文件。

    Raises:
        HTTPException: 如果上传的不是图片、读取失败、调用Ollama服务失败或未识别到公式。

    Returns:
        dict: 包含识别出的LaTeX代码 `{"latex": "..."}`。
    """
    if not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="上传的文件不是图片类型。")

    try:
        image_bytes = await file.read()
        image_b64 = base64.b64encode(image_bytes).decode("utf-8")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取或编码图片失败: {e}")

    prompt = "Transcribe the formula into a single line of raw LaTeX. Ensure matrices (e.g., pmatrix) are correctly formatted. Output ONLY the code, no markdown or explanations. Return 'None' if no formula found."
    payload = {
        "model": "qwen3-vl:4b",
        "messages": [
            {
                "role": "user",
                "content": prompt,
                "images": [image_b64]
            }
        ],
        "stream": False
    }

    try:
        # 【修正】移除 'proxy' 参数
        async with httpx.AsyncClient(timeout=300.0) as client:
            response = await client.post(OLLAMA_API_URL, json=payload)
            response.raise_for_status()

        response_data = response.json()
        content = response_data.get("message", {}).get("content", "").strip()

        if not content or content == "None":
            raise HTTPException(status_code=404, detail="未检测到数学公式")

        return {"latex": content}

    except httpx.RequestError as e:
        raise HTTPException(status_code=503, detail=f"调用Ollama服务失败: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"处理响应时发生未知错误: {e}")


@app.post("/parse-document")
async def parse_document_endpoint(file: UploadFile = File(...)):
    """
    接收 .docx 文件，解析为 JSON 结构并返回。
    主要用于前端导入现有文档内容到编辑器中。
    """
    # 简单的文件类型校验
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .docx file.")

    try:
        # 读取文件字节流
        docx_bytes = await file.read()

        # 调用核心解析逻辑
        doc_state = parse_docx_to_json(docx_bytes)

        # 返回标准中间态 JSON (IR)
        return doc_state
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Parsing failed: {str(e)}")


@app.post("/format-document")
async def format_document_endpoint(
        file: UploadFile = File(...),
        rules: str = Form(...)
):
    """
    【已重构 v2】接收 .docx 文件和格式规则，以流式响应的方式
    进行处理并返回新的 .docx 文件，同时实时推送详细的诊断日志。
    """
    log_queue = asyncio.Queue()

    async def stream_formatter():
        def logger(message: str, stage: str = "INFO"):
            """Helper to print to console and send to frontend via queue."""
            log_message = f"[{stage}] {message}"
            print(log_message)
            try:
                log_queue.put_nowait(log_message)
            except asyncio.QueueFull:
                print("[WARNING] Log queue is full. A message was dropped.")

        # This inner function will run the main logic and use the logger
        async def run_formatting_logic():
            try:
                logger("Workflow started.", "START")

                # --- Stage 1: Validation and Parsing Rules ---
                logger("Validating input file and rules...", "VALIDATE")
                try:
                    formatting_rules = json.loads(rules)
                except json.JSONDecodeError:
                    raise ValueError("Invalid JSON format for rules.")

                if not file.filename.endswith('.docx'):
                    raise ValueError("Invalid file type. Please upload a .docx file.")
                logger("Input validation successful.", "VALIDATE")

                # --- Stage 2: Parsing DOCX to JSON IR ---
                logger("Reading uploaded DOCX file...", "PARSE")
                docx_bytes = await file.read()
                logger(f"File read successfully ({len(docx_bytes)} bytes).", "PARSE")

                logger("Parsing DOCX structure into JSON Intermediate Representation (IR)...", "PARSE")
                doc_state_ir = parse_docx_to_json(docx_bytes)
                element_count = len(doc_state_ir.get("sections", [{}])[0].get("elements", []))
                logger(f"Parsing complete. Found {element_count} block-level elements.", "PARSE")
                logger(f"Initial IR JSON:\n{json.dumps(doc_state_ir, indent=2, ensure_ascii=False)}", "DEBUG")

                # --- Stage 3: Applying Formatting Rules ---
                logger("Applying formatting rules to the JSON IR...", "FORMAT")
                modified_ir = apply_formatting(doc_state_ir, formatting_rules)
                logger("Formatting rules applied successfully.", "FORMAT")
                logger(f"Modified IR JSON:\n{json.dumps(modified_ir, indent=2, ensure_ascii=False)}", "DEBUG")

                # --- Stage 4: Generating New DOCX from Modified IR ---
                logger("Regenerating .docx file from the modified IR...", "GENERATE")
                new_docx_bytes, generator_log = await create_document(modified_ir)
                if not new_docx_bytes:
                    raise RuntimeError("The document generator returned an empty file.")
                logger(f"New .docx file generated successfully ({len(new_docx_bytes)} bytes).", "GENERATE")
                if generator_log:
                    logger(f"Generator Log:\n--- START ---\n{generator_log}\n--- END ---", "DEBUG")

                return new_docx_bytes

            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                logger(f"Workflow failed with an error: {type(e).__name__}: {e}", "ERROR")
                logger(f"Full traceback:\n{error_details}", "DEBUG")
                return None  # Indicate failure

        # Task to run the logic in the background
        formatting_task = asyncio.create_task(run_formatting_logic())

        # Stream logs from the queue
        while not formatting_task.done() or not log_queue.empty():
            try:
                log_line = await asyncio.wait_for(log_queue.get(), timeout=0.1)
                yield f"data: {json.dumps({'type': 'log', 'content': log_line})}\n\n"
            except asyncio.TimeoutError:
                continue

        # Get the final result
        final_result_bytes = await formatting_task

        if final_result_bytes:
            # Stream the file content
            encoded_string = base64.b64encode(final_result_bytes).decode('utf-8')
            CHUNK_SIZE = 32 * 1024
            for i in range(0, len(encoded_string), CHUNK_SIZE):
                chunk = encoded_string[i:i + CHUNK_SIZE]
                yield f"data: {json.dumps({'type': 'file_chunk', 'content': chunk})}\n\n"

            file_metadata = {
                "file_name": f"formatted_{file.filename}",
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            yield f"data: {json.dumps({'type': 'file_end', 'content': file_metadata})}\n\n"
        else:
            # Signal an error to the frontend
            yield f"data: {json.dumps({'type': 'error', 'content': 'Formatting workflow failed. See logs for details.'})}\n\n"

    return StreamingResponse(stream_formatter(), media_type="text/event-stream")