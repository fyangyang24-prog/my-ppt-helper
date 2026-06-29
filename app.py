import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Pt

# --- 核心辅助函数 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    # ... (保持你原有的 PPT 处理逻辑) ...
    output_path = "summary_report.pptx"
    prs.save(output_path)
    return output_path

# --- 初始化环境 ---
st.set_page_config(page_title="现场巡场助手", layout="centered")
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_val" not in st.session_state: st.session_state.desc_val = ""

st.title("📱 现场巡场助理 (修复版)")

# --- 1. 智能检索 (使用回调函数修复点击无效) ---
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")

def insert_rule_callback():
    st.session_state.desc_val = f"【技术规定依据】：{st.session_state.selected_rule_box}\n【现场实况说明】：\n"

if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        # 使用 key 绑定选择框，并设置 on_change 回调
        st.selectbox("选择条文：", matched, key="selected_rule_box")
        st.button("✅ 点此插入条文到描述框", on_click=insert_rule_callback)

# --- 2. 录入区 ---
# 描述框直接使用 session_state 确保不丢失
desc = st.text_area("问题描述", value=st.session_state.desc_val, key="desc_input")
solve = st.text_area("解决措施", key="solve_input")

if st.button("➕ 确认并添加此条问题"):
    st.session_state.problem_list.append({"desc": desc, "solve": solve})
    st.session_state.desc_val = "" # 清空描述
    st.rerun()

# --- 3. 下载区 (微信优化版) ---
st.divider()
if len(st.session_state.problem_list) > 0:
    if st.button("🚀 生成 PPT 报告"):
        out_file = build_multi_page_ppt(...)
        st.session_state.final_file = out_file
    
    if "final_file" in st.session_state and os.path.exists(st.session_state.final_file):
        with open(st.session_state.final_file, "rb") as f:
            st.download_button(
                label="📥 立即下载 (如微信无法下载，请点右上角选择：在浏览器打开)",
                data=f,
                file_name="巡场报告.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )