import streamlit as st
import os
import tempfile

# 这里导入你写好的三个引擎（确保这三个py文件和app.py在同一个目录下）
# 需要把你之前的脚本稍微改一下，把 process_single_english_ppt 这类函数导出来
from EN import process_single_english_ppt
# from geo_batch import process_single_geo_ppt
# from main import process_single_physics_ppt # 假设物理叫这个

# --- 页面配置 ---
st.set_page_config(page_title="自动化讲义生成中心", page_icon="📝", layout="centered")

st.title("📚 自动化讲义生成中心")
st.markdown("欢迎！请选择对应的学科，并上传您的 PPTX 文件，系统将自动生成 Word 讲义。")

# --- 侧边栏：学科选择 ---
subject = st.sidebar.radio(
    "1. 请选择学科模型",
    ("英语阅读讲义", "地理讲义", "物理讲义")
)

# --- 主界面：文件上传 ---
uploaded_file = st.file_uploader("2. 上传 PPTX 文件", type=["pptx"])

if uploaded_file is not None:
    st.info(f"已上传: {uploaded_file.name}，正在准备处理...")
    
    if st.button("🚀 开始生成讲义"):
        with st.spinner('AI 引擎正在奋力排版中，请耐心等待（约需 10-30 秒）...'):
            try:
                # 1. 因为上传的是内存文件，需先保存到服务器临时目录
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
                    tmp_ppt.write(uploaded_file.getvalue())
                    tmp_ppt_path = tmp_ppt.name
                
                # 2. 准备输出路径
                output_docx_name = uploaded_file.name.replace(".pptx", f"_{subject}.docx")
                # 这里为了简单，我们让子脚本生成固定名字，然后我们重命名
                
                # 3. 呼叫对应的学科引擎
                if subject == "英语阅读讲义":
                    # 假定你的 EN.py 被稍作修改，允许传入输入路径和输出路径
                    process_single_english_ppt(tmp_ppt_path, '模板.docx', 'icons', output_docx_name)
                elif subject == "地理讲义":
                    # process_single_geo_ppt(...) 
                    pass
                elif subject == "物理讲义":
                    # process_single_physics_ppt(...)
                    pass

                # 4. 读取生成好的 Word 提供下载
                if os.path.exists(output_docx_name):
                    with open(output_docx_name, "rb") as file:
                        btn = st.download_button(
                            label="✅ 处理完成！点击下载 Word 讲义",
                            data=file,
                            file_name=output_docx_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    st.balloons() # 撒花特效
                else:
                    st.error("生成失败，未找到输出文件。")
                    
            except Exception as e:
                st.error(f"处理过程中出现错误：{e}")
                
st.sidebar.markdown("---")
st.sidebar.caption("👨‍💻 研发支持：您的名字/部门")