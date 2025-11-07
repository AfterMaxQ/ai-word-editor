import streamlit as st
from src.app_logic import generate_document_from_command

# 设置页面标题和图标
st.set_page_config(page_title="AI 文档编辑器", page_icon="✍️")

#主标题
st.title("✍️ AI 文档生成器")
st.caption("只需用自然语言描述，即可生成格式精准的Word文档！")

user_command = st.text_area(
    "请输入您的文档生成指令：",
    height=200,
    placeholder="例如：创建一个标题叫'项目报告'，然后另起一段，内容是'这是第一季度的总结'，宋体小四，首行缩进。"
)

if st.button("🚀 生成文档", type="primary"):
    if user_command:
        # 使用 spinner 显示加载状态
        with st.spinner("🧠 AI正在思考，引擎正在构建，请稍候..."):
            try:
                document_bytes = generate_document_from_command(user_command)
                if document_bytes:
                    st.success("🎉 文档生成成功！请点击下方按钮下载。")

                    # 下载按钮
                    st.download_button(
                        label="📥 下载 Word 文档",
                        data=document_bytes,
                        file_name="generate_document.docx",
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("❌ 文档生成失败。请检查您的指令或Ollama服务是否正常运行。")
            except Exception as e:
                st.error(f"发生错误：{e}")
    else:
        st.warning("请输入指令后再点击生成！")
