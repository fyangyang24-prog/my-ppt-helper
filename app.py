import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 基础配置与初始化 ---
st.set_page_config(page_title="现场巡场助理", layout="centered")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 2. 核心替换逻辑 ---
def replace_text_in_shape(shape, data):
    """最稳健的文本替换，只替换文本，不改变形状结构"""
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for key, val in data.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(val))
    elif hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if hasattr(cell, "has_text_frame") and cell.has_text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        for key, val in data.items():
                            if key in paragraph.text:
                                paragraph.text = paragraph.text.replace(key, str(val))

# --- 3. 稳健生成 PPT ---
def build_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    
    # 直接在现有的 PPT 幻灯片里填数据
    # 如果你有 N 个问题，确保 template.pptx 预先有 N 页
    for i, prob in enumerate(problem_list):
        if i < len(prs.slides):
            slide = prs.slides[i]
            data = {
                "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
                "{{type}}": prob.get("type", ""), "{{desc}}": prob["desc"],
                "{{solve}}": prob["solve"], "{{duty}}": prob["duty"],
                "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
            }
            for shape in slide.shapes:
                replace_text_in_shape(shape, data)
            
            # 尝试添加图片
            if prob.get("img_base64"):
                try:
                    img_bytes = base64.b64decode(prob["img_base64"])
                    temp_path = f"temp_{i}.jpg"
                    with open(temp_path, "wb") as f: f.write(img_bytes)
                    slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
                except: pass
                
    output_path = "巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 4. 界面展示 ---
st.title("📱 现场巡场助理")
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

prob_type = st.radio("问题分类", ["底线", "严控事项", "设计红黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True)
uploaded_file = st.file_uploader("照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision = st.radio("整改", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "type": prob_type, "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": final_duty, "decision": f"{decision} √", "deadline": deadline
    })
    st.rerun()

if st.button("🚀 生成报告"):
    out = build_ppt(project_title, final_user, check_date, st.session_state.problem_list)
    if out:
        with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="巡场报告.pptx")