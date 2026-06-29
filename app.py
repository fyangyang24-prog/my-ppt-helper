import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Pt

# --- 核心辅助函数 (保持原有功能) ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

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
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str, 
            "{{desc}}": prob["desc"], "{{solve}}": prob["solve"], 
            "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
        }
        for shape in slide.shapes:
            if shape.has_text_frame:
                for k, v in data.items(): safe_replace_text(shape.text_frame, k, v)
        # 图片处理逻辑
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                with open("temp.jpg", "wb") as f: f.write(img_bytes)
                slide.shapes.add_picture("temp.jpg", left=Pt(50), top=Pt(150), width=Pt(200))
            except: pass
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 主程序 UI (还原最初简洁风格) ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")

if "problem_list" not in st.session_state: st.session_state.problem_list = []

st.title("📱 现场巡场助理")

# 1. 基础信息
project_title = st.text_input("项目名称", value="独立路壹号项目")
user_name = st.text_input("检查人")
check_date = st.date_input("检查时间")

# 2. 技术规定检索插件 (独立板块，不影响原有逻辑)
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 [辅助] 检索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched)
        st.info(f"选中的条文：{chosen}")
        st.caption("提示：你可以手动复制上面的条文粘贴到下方的描述框中。")

# 3. 录入区
uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "jpeg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
duty = st.text_input("责任人")
decision = st.radio("决定", ["整改", "不整改"], horizontal=True)

# 4. 核心按钮 (恢复到最简最稳模式)
if st.button("➕ 确认并添加此条问题"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({"img_base64": img_b64, "desc": desc, "solve": solve, "duty": duty, "decision": decision, "deadline": str(check_date)})
    st.success("添加成功！")

if st.button("🚀 生成并下载 PPT"):
    out_file = build_multi_page_ppt(project_title, user_name, str(check_date), st.session_state.problem_list)
    if out_file:
        with open(out_file, "rb") as f:
            st.download_button("📥 点击下载 PPT", f, "报告.pptx")