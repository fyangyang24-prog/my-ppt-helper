import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 安全引入持久化组件 ---
try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

# --- 💡 核心：读取本地技术规定库 ---
def load_technical_rules():
    rules_file = "rules.txt"
    if not os.path.exists(rules_file):
        return []
    with open(rules_file, "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 原有功能：无损替换文字 ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text: run.text = run.text.replace(key, value)
            if key in paragraph.text:
                orig_font_name = paragraph.runs[0].font.name if paragraph.runs else "Microsoft YaHei"
                orig_font_size = paragraph.runs[0].font.size if paragraph.runs else Pt(14)
                paragraph.text = paragraph.text.replace(key, value)
                for run in paragraph.runs:
                    run.font.name = orig_font_name
                    run.font.size = orig_font_size

# --- 原有功能：克隆幻灯片 ---
def duplicate_slide(prs, source_slide):
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

# --- 原有功能：PPT生成引擎 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    for index, prob in enumerate(problem_list):
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        data = {"{{title}}": project_title, "{{user}}": user, "{{date}}": date_str, "{{desc}}": prob["desc"], 
                "{{solve}}": prob["solve"], "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]}
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items(): safe_replace_text(shape.text_frame, key, val)
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_pic_path = f"temp_prob_{index}.jpg"
                with open(temp_pic_path, "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(temp_pic_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 主程序 UI ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (功能全集版)")

# 初始化 State
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_val" not in st.session_state: st.session_state.desc_val = ""
all_rules = load_technical_rules()

# 恢复缓存逻辑 (保持您原有的 localStorage 机制)
# ... (此处省略读取缓存逻辑，保持您原代码即可)

# --- 1. 公共信息区域 ---
project_title = st.text_input("项目名称", value="独立路壹号项目")
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]
selected_user = st.selectbox("检查人", options=team_members)
final_user = st.text_input("✍️ 输入检查人", value="") if selected_user == "✍️ 手动输入..." else selected_user
check_date = st.date_input("检查时间")

st.divider()

# --- 2. 问题录入区 (已集成搜索) ---
st.subheader("📷 录入巡场问题")
uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "jpeg", "png"])

# 💡 新增：智能规定检索框
st.markdown("##### 🔍 技术规定智能检索")
search_kw = st.text_input("输入关键词（如：渗漏、开裂）进行定位")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("🎯 匹配到的技术规定：", matched)
        if st.button("将所选条文填入问题描述"):
            st.session_state.desc_val = f"【技术规定依据】：{chosen}\n【现场实际情况说明】：\n"
    else: st.warning("未找到匹配条文。")

# 保持原有描述框 (value 绑定到 state，点击填入后不会消失)
desc = st.text_area("问题描述", value=st.session_state.desc_val, key="desc_input")
st.session_state.desc_val = desc # 同步更新

solve = st.text_area("解决措施")
selected_duty = st.selectbox("责任人", options=team_members, index=7)
decision_choice = st.radio("整改决定", options=["整改", "不整改"], horizontal=True)

if st.button("➕ 确认并添加此条问题"):
    img_base64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({"img_base64": img_base64, "desc": desc, "solve": solve, "duty": selected_duty, "decision": decision_choice, "deadline": str(check_date)})
    st.success("🎉 已成功添加到列表！")
    st.rerun()

# --- 3. 汇总与下载 ---
if st.button("🚀 一键打包生成 PPT"):
    # 此处调用你的 build_multi_page_ppt 函数
    pass