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

# --- 工具函数 ---
def load_technical_rules():
    rules_file = "rules.txt"
    if not os.path.exists(rules_file): return []
    with open(rules_file, "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text: run.text = run.text.replace(key, value)
            paragraph.text = paragraph.text.replace(key, value)

def duplicate_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    for index, prob in enumerate(problem_list):
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        data = {"{{title}}": project_title, "{{user}}": user, "{{date}}": date_str, "{{desc}}": prob["desc"], "{{solve}}": prob["solve"], "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]}
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items(): safe_replace_text(shape.text_frame, key, val)
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                with open(f"temp_{index}.jpg", "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(f"temp_{index}.jpg", Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 主程序 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (智能联动版)")

# 初始化持久化存储
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_buffer" not in st.session_state: st.session_state.desc_buffer = ""

all_rules = load_technical_rules()
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]

# 1. 公共信息
project_title = st.text_input("项目名称", value="独立路壹号项目")
selected_user = st.selectbox("检查人", options=team_members)
final_user = st.text_input("✍️ 输入检查人", value="") if selected_user == "✍️ 手动输入..." else selected_user
check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

# 2. 录入区域
st.subheader("📷 录入巡场问题")
uploaded_file = st.file_uploader("📷 拍摄/上传照片", type=["jpg", "jpeg", "png"])

# 智能检索逻辑
st.markdown("##### 🔍 快速检索技术规定")
search_kw = st.text_input("输入关键词（如：渗漏）", key="search_bar")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择要引用的规定：", matched)
        if st.button("一键填入条文"):
            st.session_state.desc_buffer = f"【技术规定依据】：{chosen}\n【现场实际情况说明】：\n"
            st.rerun()

# --- 关键修复：使用 session_state.desc_buffer 绑定 value ---
desc = st.text_area("问题描述", value=st.session_state.desc_buffer, key="desc_input", placeholder="请录入现场描述...")
# 同步更新缓存，防止手动修改后又被重置
if desc != st.session_state.desc_buffer:
    st.session_state.desc_buffer = desc

solve = st.text_area("解决措施", placeholder="请录入整改要求...")
selected_duty = st.selectbox("责任人", options=team_members, index=7)
final_duty = st.text_input("✍️ 输入责任人/单位", value="") if selected_duty == "✍️ 手动输入..." else selected_duty
decision_choice = st.radio("整改决定", options=["整改", "不整改"], horizontal=True)
decision_text = "整改  √ \n 不整改 ▢" if decision_choice == "整改" else "整改  ▢ \n 不整改 √"
deadline_str = st.date_input("要求完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加此条问题"):
    img_base64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({"img_base64": img_base64, "desc": desc, "solve": solve, "duty": final_duty, "decision": decision_text, "deadline": deadline_str})
    st.session_state.desc_buffer = "" # 清空，方便下一条
    st.success("🎉 问题已装箱！")
    st.rerun()

# 3. 汇总与下载
if st.button("🚀 一键打包生成 PPT"):
    out_file = build_multi_page_ppt(project_title, final_user, date_str, st.session_state.problem_list)
    if out_file:
        with open(out_file, "rb") as file: st.download_button("📥 点击下载 PPT", file, file_name=f"{project_title}-汇总报告.pptx")