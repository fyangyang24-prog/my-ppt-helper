import streamlit as st
import os
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 核心无损替换发动机 ---
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

def build_ppt_from_web(uploaded_file, project_title, user, date_str, desc, solve, duty, deadline_str, decision_text):
    template_path = "template.pptx"
    if not os.path.exists(template_path):
        return None
        
    prs = Presentation(template_path)
    slide = prs.slides[0]

    # 映射表，把单选框组合好的符号文本塞给 {{decision}}
    data = {
        "{{title}}": project_title,
        "{{user}}": user,
        "{{date}}": date_str,
        "{{desc}}": desc,
        "{{solve}}": solve,
        "{{duty}}": duty,
        "{{deadline}}": deadline_str,
        "{{decision}}": decision_text  # 👈 新增：整改决定的打勾文本
    }

    for shape in slide.shapes:
        if shape.has_text_frame:
            for key, val in data.items():
                safe_replace_text(shape.text_frame, key, val)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for key, val in data.items():
                        safe_replace_text(cell.text_frame, key, val)

    if uploaded_file is not None:
        temp_path = "temp_mobile_pic.jpg"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))

    output_path = "手机生成的巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 📱 终极完美版手机网页界面 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (终极完美版)")

st.divider()

# 1. 照片拍照/上传
uploaded_file = st.file_uploader("📷 点击拍照 / 上传问题照片", type=["jpg", "jpeg", "png"])

# 2. 基础信息填写
project_title = st.text_input("项目名称", value="独立路壹号项目")

# 快捷选择检查人
user_options = ["樊洋洋", "刘璐", "其他人员"]
selected_user = st.selectbox("检查人", options=user_options, index=0)
if selected_user == "其他人员":
    user = st.text_input("请输入实际检查人姓名")
else:
    user = selected_user

# 弹出日历选时间
check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# 问题与措施（默认留空）
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...")

st.divider()

# 快捷选择责任人
duty_options = ["樊洋洋", "刘璐", "外部单位/其他人"]
selected_duty = st.selectbox("责任人", options=duty_options, index=0)
if selected_duty == "外部单位/其他人":
    duty = st.text_input("请输入实际责任人姓名")
else:
    duty = selected_duty

# 🌟 新增优化点：整改决定单选框（手机端点击切换）
decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True)
# 根据选择，自动拼装成带符号的文字输出给 PPT
if decision_choice == "整改":
    decision_text = "整改  √ \n 不整改 ▢"
else:
    decision_text = "整改  ▢ \n 不整改 √"

# 弹出日历选完成时间
deadline_date = st.date_input("要求完成时间")
deadline_str = deadline_date.strftime("%Y/%m/%d")

st.divider()

# 3. 生成与下载
if st.button("🚀 现场一键生成 PPT 报告"):
    if not os.path.exists("template.pptx"):
        st.error("❌ 错误：请确保云端有名为 template.pptx 的模板文件！")
    else:
        with st.spinner("正在智能排版中..."):
            out_file = build_ppt_from_web(uploaded_file, project_title, user, date_str, desc, solve, duty, deadline_str, decision_text)
            if out_file:
                st.success("🎉 PPT 生成成功！")
                with open(out_file, "rb") as file:
                    st.download_button(
                        label="📥 立即点击下载 PPT 到手机",
                        data=file,
                        file_name=f"{project_title}-巡场报告.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )