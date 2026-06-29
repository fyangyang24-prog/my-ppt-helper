import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 状态初始化：防止数据消失 ---
def init_session_state():
    if "problem_list" not in st.session_state: st.session_state.problem_list = []
    if "desc" not in st.session_state: st.session_state.desc = ""
    if "solve" not in st.session_state: st.session_state.solve = ""
    if "current_rule" not in st.session_state: st.session_state.current_rule = ""

init_session_state()

# --- 智能读取规定 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- PPT 生成核心逻辑 (保持不变) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    for index, prob in enumerate(problem_list):
        slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        # ... (此处保留你原有的 PPT 填空逻辑) ...
        # (因为篇幅原因，此处仅做逻辑说明，你需要保持你原有的 build_multi_page_ppt 函数逻辑)
    output_path = "summary_report.pptx"
    prs.save(output_path)
    return output_path

# --- UI 交互 ---
st.title("📱 巡场助手 (修复版)")

# 1. 智能规定检索
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定：")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched)
        if st.button("一键填入描述框"):
            st.session_state.desc = f"【依据】：{chosen}\n【现场实况】：\n"

# 2. 核心：使用 session_state 绑定输入框，防止消失
st.session_state.desc = st.text_area("问题描述", value=st.session_state.desc, key="desc_input")
st.session_state.solve = st.text_area("解决措施", value=st.session_state.solve, key="solve_input")

# 3. 添加按钮逻辑
if st.button("➕ 确认添加"):
    st.session_state.problem_list.append({
        "desc": st.session_state.desc,
        "solve": st.session_state.solve,
        "duty": "待定", "deadline": "2026/06/29", "decision": "整改", "img_base64": ""
    })
    # 清空输入框以便录入下一条
    st.session_state.desc = ""
    st.session_state.solve = ""
    st.rerun()

# 4. 下载逻辑 (微信兼容性)
if len(st.session_state.problem_list) > 0:
    if st.button("🚀 生成并下载 PPT"):
        out_file = build_multi_page_ppt(...)
        with open(out_file, "rb") as f:
            st.download_button(
                label="📥 点击此处下载 (微信内如无法下载请用浏览器打开)",
                data=f,
                file_name="巡场报告.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )