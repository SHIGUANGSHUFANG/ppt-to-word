import streamlit as st
from pathlib import Path
import shutil
import tempfile

# 导入你写好的各个学科处理引擎
from EN import process_single_english_ppt
from listen import process_single_listening_ppt
from writing import process_single_writing_ppt
from geo import process_single_geo_ppt
from wl import process_single_physics_ppt

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "模板.docx"
ICON_DIR = BASE_DIR / "icons"

PROCESSORS = {
    "英语阅读讲义": process_single_english_ppt,
    "英语听力讲义": process_single_listening_ppt,
    "英语作文讲义": process_single_writing_ppt,
    "地理讲义": process_single_geo_ppt,
    "物理讲义": process_single_physics_ppt,
}

# --- 页面配置 ---
st.set_page_config(page_title="自动化讲义生成中心", page_icon="📝", layout="centered")

st.title("📚 自动化讲义生成中心")
st.markdown("欢迎！请选择对应的学科，并上传您的 PPTX 文件，系统将自动生成 Word 讲义。")

# --- 侧边栏：学科选择 ---
subject = st.sidebar.radio(
    "1. 请选择学科模型",
    tuple(PROCESSORS.keys())
)

# --- 主界面：文件上传 ---
uploaded_file = st.file_uploader("2. 上传 PPTX 文件", type=["pptx"])

if uploaded_file is not None:
    st.info(f"已上传: {uploaded_file.name}，正在准备处理...")
    
    if st.button("🚀 开始生成讲义"):
        with st.spinner('AI 引擎正在奋力排版中，请耐心等待（约需 10-30 秒）...'):
            try:
                if not TEMPLATE_PATH.exists():
                    raise FileNotFoundError(f"未找到模板文件：{TEMPLATE_PATH}")
                if not ICON_DIR.exists():
                    raise FileNotFoundError(f"未找到图标目录：{ICON_DIR}")

                with tempfile.TemporaryDirectory() as work_dir:
                    work_dir_path = Path(work_dir)
                    tmp_ppt_path = work_dir_path / uploaded_file.name
                    run_icon_dir = work_dir_path / "icons"
                    output_docx_name = f"{Path(uploaded_file.name).stem}_{subject}.docx"
                    output_docx_path = work_dir_path / output_docx_name
                    shutil.copytree(ICON_DIR, run_icon_dir)

                    # 1. 保存上传的内存文件到服务器临时目录
                    with open(tmp_ppt_path, "wb") as tmp_ppt:
                        tmp_ppt.write(uploaded_file.getvalue())

                    # 2. 呼叫对应的学科引擎
                    processor = PROCESSORS[subject]
                    processor(
                        str(tmp_ppt_path),
                        str(TEMPLATE_PATH),
                        str(run_icon_dir),
                        str(output_docx_path),
                    )

                    # 3. 读取生成好的 Word 提供下载
                    if not output_docx_path.exists():
                        st.error(f"生成失败，{subject} 引擎没有生成输出文件。")
                    else:
                        output_bytes = output_docx_path.read_bytes()
                        st.download_button(
                            label="✅ 处理完成！点击下载 Word 讲义",
                            data=output_bytes,
                            file_name=output_docx_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                        st.balloons()
                    
            except Exception as e:
                st.error(f"处理过程中出现错误：{e}")
                
st.sidebar.markdown("---")
st.sidebar.caption("👨‍💻 研发支持：您的名字/部门")
