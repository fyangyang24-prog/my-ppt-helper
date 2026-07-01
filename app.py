import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 1. PPT 处理核心函数 ---
def safe_replace_text(text_frame, key, value):
    if text_frame is None: return
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            paragraph.text = paragraph.text.replace(key, str(value))

def duplicate_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in source_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shape.element))
    return new_slide

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{desc}}": prob["desc"], "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items(): safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items(): safe_replace_text(cell.text_frame, key, val)
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_path = f"temp_{index}.jpg"
                with open(temp_path, "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
    prs.save("最终报告.pptx")
    return "最终报告.pptx"

# --- 2. 界面逻辑 ---
st.set_page_config(page_title="巡场助手", layout="centered")
st.title("📱 现场巡场助理")

if "problem_list" not in st.session_state: st.session_state.problem_list = []

# 输入区
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

st.divider()
st.subheader(f"📷 录入问题 (已暂存 {len(st.session_state.problem_list)} 条)")

uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    if not desc or not final_duty:
        st.error("请填写描述和责任人！")
    else:
        img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
        st.session_state.problem_list.append({
            "img_base64": img_b64, "desc": desc, "solve": solve, 
            "duty": final_duty, "decision": decision, "deadline": deadline
        })
        st.rerun()

if st.button("🚀 生成 PPT 报告"):
    if not st.session_state.problem_list:
        st.error("请先添加问题！")
    else:
        out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
        if out:
            with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="巡场报告.pptx")
        else: st.error("未找到 template.pptx 模板文件！")