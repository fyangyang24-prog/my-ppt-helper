import streamlit as st
import os
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 核心逻辑：直接调用你刚刚跑通的无损填空代码 ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
            if key in paragraph.text:
                orig_font_name = paragraph.runs[0].font.name if paragraph.runs else "Microsoft YaHei"
                orig_font_size = paragraph.runs[0].font.size if paragraph.runs else Pt(14)
                orig_font_color = paragraph.runs[0].font.color.rgb if paragraph.runs and paragraph.runs[0].font.color else None
                
                paragraph.text = paragraph.text.replace(key, value)
                for run in paragraph.runs:
                    run.font.name = orig_font_name
                    run.font.size = orig_font_size
                    if orig_font_color:
                        run.font.color.rgb = orig_font_color

def build_ppt_from_web(uploaded_file, project_title, user, date, desc, solve, duty, deadline):
    template_path = "template.pptx"
    if not os.path.exists(template_path):
        return None
        
    prs = Presentation(template_path)
    slide = prs.slides[0]

    # 将手机端输入的数据做成映射
    data = {
        "{{title}}": project_title,
        "{{user}}": user,
        "{{date}}": date,
        "{{desc}}": desc,
        "{{solve}}": solve,
        "{{duty}}": duty,
        "{{deadline}}": deadline
    }

    # 开始无损替换
    for shape in slide.shapes:
        if shape.has_text_frame:
            for key, val in data.items():
                safe_replace_text(shape.text_frame, key, val)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for key, val in data.items():
                        safe_replace_text(cell.text_frame, key, val)

    # 📸 处理现场拍的照片
    if uploaded_file is not None:
        # 将手机上传的照片临时存下来
        temp_path = "temp_mobile_pic.jpg"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        # 精准贴在左侧格子上
        slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))

    output_path = "手机生成的巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 📱 手机端网页界面设计 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (手机版)")
st.write("在现场拍完照，直接在此录入，一键生成和电脑端一样完美的 PPT。")

st.divider()

# 1. 拍照/上传组件（手机上打开时直接唤起相机！）
uploaded_file = st.file_uploader("📷 点击拍照 / 上传问题照片", type=["jpg", "jpeg", "png"])

# 2. 录入表单
project_title = st.text_input("项目名称", value="独立路壹号项目")
user = st.text_input("检查人", value="刘璐")
date = st.text_input("检查时间", value="2026/06/23")
desc = st.text_area("问题描述", value="空间、布局、尺度：\n1、项目的各项配置标准应满足《住宅大盘区景观分档》 不满足")
solve = st.text_area("解决措施", value="通过各个宅间增加价值树、软装弱化问题。")
duty = st.text_input("责任人", value="刘璐")
deadline = st.text_input("要求完成时间", value="2026/07/10")

st.divider()

# 3. 生成与下载
if st.button("🚀 现场一键生成 PPT 报告"):
    if not os.path.exists("template.pptx"):
        st.error("❌ 错误：请确保文件夹里有 template.pptx 模板文件！")
    else:
        with st.spinner("正在智能排版中..."):
            out_file = build_ppt_from_web(uploaded_file, project_title, user, date, desc, solve, duty, deadline)
            if out_file:
                st.success("🎉 PPT 生成成功！")
                with open(out_file, "rb") as file:
                    st.download_button(
                        label="📥 立即点击下载 PPT 到手机",
                        data=file,
                        file_name=f"{project_title}-巡场报告.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )