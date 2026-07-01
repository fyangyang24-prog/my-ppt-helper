import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 💡 安全组件 ---
try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

# --- 核心：无损替换函数 (锁定字体防止乱码) ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            ref_name = "Microsoft YaHei"
            ref_size = Pt(14)
            if paragraph.runs:
                ref_name = paragraph.runs[0].font.name or ref_name
                ref_size = paragraph.runs[0].font.size or ref_size
            
            paragraph.text = paragraph.text.replace(key, str(value))
            
            for run in paragraph.runs:
                run.font.name = ref_name
                run.font.size = ref_size

# --- 核心：PPT 生成引擎 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        # 严格控制每一页的独立性
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = prs.slides.add_slide(source_slide.slide_layout)
            for shape in source_slide.shapes:
                new_shape = current_slide.shapes.add_shape(shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height)
                if shape.has_text_frame:
                    new_shape.text = shape.text
        
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{type}}": prob.get("type", ""), "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"], "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
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
            
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 界面逻辑 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")

# 初始化状态
if "problem_list" not in st.session_state: st.session_state.problem_list = []

st.title("📱 现场巡场助理")
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

st.divider()
st.subheader(f"📷 录入巡场问题 (已添加 {len(st.session_state.problem_list)} 条)")

# 单选逻辑
prob_type = st.radio("巡场问题分类", options=["底线", "严控事项", "设计红黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True)

uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision_choice = st.radio("整改决定", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    if not desc or not final_duty:
        st.error("请完善信息！")
    else:
        img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
        st.session_state.problem_list.append({
            "type": prob_type, "img_base64": img_b64, "desc": desc, "solve": solve, 
            "duty": final_duty, "decision": f"{decision_choice} √", "deadline": deadline
        })
        st.rerun() # 自动清除输入框并刷新

if st.button("🚀 生成 PPT 报告"):
    if not st.session_state.problem_list:
        st.error("请先添加问题！")
    else:
        out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
        if out:
            with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="报告.pptx")