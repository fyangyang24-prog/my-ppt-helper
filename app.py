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

# --- 核心辅助函数 ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

def duplicate_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in source_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shape.element))
    return new_slide

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        data = {
            "{{title}}": project_title,
            "{{user}}": user,
            "{{date}}": date_str,
            "{{type}}": prob.get("type", ""), # 新增：问题分类
            "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items(): safe_replace_text(shape.text_frame, key, val)
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                with open(f"temp_{index}.jpg", "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(f"temp_{index}.jpg", Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
    prs.save("最终汇总巡场报告.pptx")
    return "最终汇总巡场报告.pptx"

# --- 界面 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理")

if "problem_list" not in st.session_state: st.session_state.problem_list = []

# 恢复逻辑
if JS_AVAILABLE and "saved_problem_list" not in st.session_state:
    try:
        data = st_javascript("localStorage.getItem('xuncha_backup_v2');")
        if data and data != "null":
            backup = json.loads(data)
            st.session_state.problem_list = backup.get("problem_list", [])
            st.session_state.saved_project_title = backup.get("project_title", "独立路壹号项目")
            st.session_state.saved_user = backup.get("user", "樊洋洋")
            st.session_state.saved_problem_list = True
    except: pass

team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]

project_title = st.text_input("项目名称", value=st.session_state.get("saved_project_title", "独立路壹号项目"))
selected_user = st.selectbox("检查人", options=team_members, index=team_members.index(st.session_state.get("saved_user", "樊洋洋")) if st.session_state.get("saved_user") in team_members else 11)
final_user = st.text_input("✍️ 输入检查人", value=st.session_state.get("saved_user", "")) if selected_user == "✍️ 手动输入..." else selected_user
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

st.divider()

# 录入区域
idx = len(st.session_state.problem_list)
st.subheader(f"📷 录入问题 ({idx}个)")
issue_type = st.radio("巡场问题分类", options=["底线", "严控事项", "设计红条", "设计黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True, key=f"type_{idx}")
uploaded_file = st.file_uploader("📷 拍摄照片", type=["jpg", "png"], key=f"file_{idx}")
desc = st.text_area("问题描述", key=f"desc_{idx}")
solve = st.text_area("解决措施", key=f"solve_{idx}")
selected_duty = st.selectbox("责任人", options=team_members, index=7, key=f"duty_{idx}")
final_duty = st.text_input("✍️ 输入责任人", key=f"duty_in_{idx}") if selected_duty == "✍️ 手动输入..." else selected_duty
decision = st.radio("整改决定", options=["整改", "不整改"], horizontal=True, key=f"dec_{idx}")
deadline = st.date_input("完成时间", key=f"dead_{idx}").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    if not final_user or not final_duty: st.error("检查人和责任人不能为空！")
    else:
        img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
        prob = {"type": issue_type, "img_base64": img_b64, "desc": desc, "solve": solve, "duty": final_duty, "decision": f"{decision} √", "deadline": deadline}
        st.session_state.problem_list.append(prob)
        if JS_AVAILABLE: st_javascript(f"localStorage.setItem('xuncha_backup_v2', '{json.dumps({'project_title': project_title, 'user': final_user, 'problem_list': st.session_state.problem_list}, ensure_ascii=False)}');")
        st.rerun()

st.divider()
if st.button("🚀 一键打包生成 PPT"):
    out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
    if out:
        with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="报告.pptx")