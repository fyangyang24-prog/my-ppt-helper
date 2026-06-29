import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

# --- 核心：读取技术规定 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 功能：PPT 生成与处理 ---
def safe_replace_text(text_frame, key, value):
    for p in text_frame.paragraphs:
        if key in p.text:
            for r in p.runs:
                if key in r.text: r.text = r.text.replace(key, value)
            p.text = p.text.replace(key, value)

def duplicate_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    for index, prob in enumerate(problem_list):
        slide = prs.slides[0] if index == 0 else duplicate_slide(prs, prs.slides[0])
        data = {"{{title}}": project_title, "{{user}}": user, "{{date}}": date_str, "{{desc}}": prob["desc"], 
                "{{solve}}": prob["solve"], "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]}
        for shape in slide.shapes:
            if shape.has_text_frame:
                for k, v in data.items(): safe_replace_text(shape.text_frame, k, v)
    output_path = "summary.pptx"
    prs.save(output_path)
    return output_path

# --- 初始化环境 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_val" not in st.session_state: st.session_state.desc_val = ""
all_rules = load_technical_rules()

st.title("📱 现场巡场助理 (完整版)")

# --- 1. 公共信息 ---
project_title = st.text_input("项目名称", value="独立路壹号项目")
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]
selected_user = st.selectbox("检查人", options=team_members)
final_user = st.text_input("输入检查人", value="") if selected_user == "✍️ 手动输入..." else selected_user
check_date = st.date_input("检查时间")

st.divider()

# --- 2. 录入与搜索 ---
st.subheader("📷 录入巡场问题")
uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "jpeg", "png"])

search_kw = st.text_input("🔍 搜索技术规定（定位后自动填入）")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched)
        if st.button("一键填入"):
            st.session_state.desc_val = f"【技术规定依据】：{chosen}\n【现场实况】：\n"
            st.rerun()

# 描述与措施绑定 State
desc = st.text_area("问题描述", value=st.session_state.desc_val, key="desc_input")
st.session_state.desc_val = desc # 同步
solve = st.text_area("解决措施")
selected_duty = st.selectbox("责任人", options=team_members, index=7)
decision_choice = st.radio("整改决定", options=["整改", "不整改"], horizontal=True)

# 确认添加逻辑
if st.button("➕ 确认添加"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": selected_duty, "decision": decision_choice, "deadline": str(check_date)
    })
    st.session_state.desc_val = "" # 清空描述
    st.success("🎉 添加成功！")
    st.rerun()

# --- 3. 打包导出 ---
if len(st.session_state.problem_list) > 0:
    if st.button("🚀 生成 PPT"):
        out_file = build_multi_page_ppt(project_title, final_user, str(check_date), st.session_state.problem_list)
        if out_file:
            with open(out_file, "rb") as f:
                st.download_button("📥 点击下载 PPT (若微信内无法下载，请点右上角在浏览器打开)", f, "最终报告.pptx")