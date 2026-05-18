import streamlit as st
import os
import tempfile
from EN import process_single_english_ppt
from wl import process_single_physics_ppt  # 确保文件名和函数名对应
from geo import process_single_geo_ppt     # 确保文件名和函数名对应

# --- 页面配置 ---
st.set_page_config(page_title="全学科自动化讲义中心", page_icon="📝", layout="centered")

st.title("📚 全学科自动化讲义中心")
st.markdown("上传您的 PPTX 文件，系统将根据学科模型自动生成标准 Word 讲义。")

# --- 侧边栏：学科模型选择 ---
subject = st.sidebar.radio(
    "1. 请选择学科模型",
    ("英语阅读讲义", "物理讲义", "地理讲义")
)

# --- 主界面：文件上传 ---
uploaded_file = st.file_uploader(f"2. 上传【{subject}】PPTX 文件", type=["pptx"])

if uploaded_file is not None:
    if st.button("🚀 开始生成讲义", type="primary"):
        with st.spinner('AI 引擎正在排版中，请耐心等待...'):
            try:
                # 保存上传的临时文件
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
                    tmp_ppt.write(uploaded_file.getvalue())
                    tmp_ppt_path = tmp_ppt.name
                
                # 定义输出文件名
                output_name = uploaded_file.name.replace(".pptx", f"_{subject}.docx")
                
                # --- 根据选择呼叫对应的引擎 ---
                if subject == "英语阅读讲义":
                    process_single_english_ppt(tmp_ppt_path, '模板.docx', 'icons', output_name)
                elif subject == "物理讲义":
                    process_single_physics_ppt(tmp_ppt_path, '模板.docx', 'icons', output_name)
                elif subject == "地理讲义":
                    process_single_geo_ppt(tmp_ppt_path, '模板.docx', 'icons', output_name)

                # 下载按钮
                if os.path.exists(output_name):
                    with open(output_name, "rb") as file:
                        st.download_button(
                            label="✅ 处理完成！点击下载讲义",
                            data=file,
                            file_name=output_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    st.balloons()
                    # 清理产生的临时文件
                    os.remove(output_name)
                
                os.remove(tmp_ppt_path)
                    
            except Exception as e:
                st.error(f"处理出错：{e}")

st.sidebar.markdown("---")
st.sidebar.info("💡 提示：如果遇到处理卡顿，请点击右侧菜单选择 'Rerun' 或联系管理员。")
