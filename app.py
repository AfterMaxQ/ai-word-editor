# src/app.py

import streamlit as st
import traceback
from src.app_logic import generate_document_from_command, polish_command

# 设置页面标题和图标
st.set_page_config(page_title="AI 文档编辑器", page_icon="✍️")


# --- 回调函数定义 ---
# 将状态更新的逻辑封装到函数中
def handle_polish_click():
    """当“润色指令”按钮被点击时执行此函数"""
    if st.session_state.user_command:
        with st.spinner("✍️ 正在为您优化指令..."):
            polished = polish_command(st.session_state.user_command)
            if polished:
                # 在回调中更新 session_state 是安全的
                st.session_state.user_command = polished


# --- 主应用流程 ---

st.title("✍️ AI 文档生成器")
st.caption("只需用自然语言描述，即可生成格式精准的Word文档！")

# 初始化 session_state，使用空字符串，并将示例文字放入 placeholder
if 'user_command' not in st.session_state:
    st.session_state.user_command = ""

# 绑定 text_area 到 session_state
st.text_area(
    "请输入您的文档生成指令：",
    height=200,
    key='user_command',  # 关键：设置一个key
    placeholder="例如：创建一个标题叫'项目报告'，然后另起一段，内容是'这是第一季度的总结'，宋体小四，首行缩进。"
)

# 使用列布局来并排显示按钮
col1, col2 = st.columns([1, 5])  # 调整比例

with col1:
    # 使用 on_click 参数将按钮与回调函数关联
    st.button(
        "✨ 润色指令",
        on_click=handle_polish_click
    )

with col2:
    if st.button("🚀 生成文档", type="primary"):
        # 读取 session_state 中的最新值
        user_command = st.session_state.user_command
        if user_command:
            with st.spinner("🧠 AI正在思考，引擎正在构建，请稍候..."):
                try:
                    document_bytes, json_str, log_str = generate_document_from_command(user_command)

                    if log_str:
                        with st.expander("查看AI处理日志 📝"):
                            st.code(log_str, language="log")

                    if document_bytes:
                        st.success("🎉 文档生成成功！请点击下方按钮下载。")
                        st.download_button(
                            label="📥 下载 Word 文档",
                            data=document_bytes,
                            file_name="generated_document.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        if json_str:
                            with st.expander("查看AI生成的最终JSON结构 👀"):
                                st.code(json_str, language="json")
                    else:
                        st.error("❌ 文档生成失败。请检查您的指令或Ollama服务是否正常运行。")

                except Exception as e:
                    st.error(f"发生错误：{e}")
                    with st.expander("查看详细错误信息 🐛"):
                        error_traceback = traceback.format_exc()
                        st.code(error_traceback, language="python")
        else:
            st.warning("请输入指令后再点击生成！")