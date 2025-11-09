# main.py
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import io

# 导入你现有的业务逻辑
from src.app_logic import polish_command, generate_document_from_command


# --- Pydantic 模型定义 ---
# 定义API请求体的数据结构
class CommandRequest(BaseModel):
    command: str


# --- FastAPI 应用实例 ---
app = FastAPI(
    title="AI 文档生成器 API",
    description="一个用于通过自然语言生成 Word 文档的 API",
    version="1.0.0",
)


# --- API 端点定义 ---

@app.get("/")
def read_root():
    """根路径，用于检查服务是否在线。"""
    return {"message": "AI 文档生成器 API 运行正常！"}


@app.post("/polish", response_model=dict)
def polish_endpoint(request: CommandRequest):
    """
    接收用户指令并返回润色后的版本。
    """
    polished_text = polish_command(request.command)
    return {"polished_command": polished_text}


@app.post("/generate")
def generate_endpoint(request: CommandRequest):
    """
    接收用户指令，生成 Word 文档并以文件流形式返回。
    """
    # 注意：generate_document_from_command 返回三个值
    document_bytes, _, _ = generate_document_from_command(request.command)

    if document_bytes:
        # 使用 StreamingResponse 高效地返回文件内容
        return StreamingResponse(
            io.BytesIO(document_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=generated_document.docx"}
        )
    else:
        return {"error": "文档生成失败"}, 400