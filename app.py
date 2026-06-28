import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 💡 新增：智能读取与匹配规定库 ---
def load_technical_rules():
    rules_file = "rules.txt"
    if not os.path.exists(rules_file):
        return []
    with open(rules_file, "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 保持你现有的其他函数（如 build_multi_page_ppt 等不变） ---
# (此处省略中间重复部分，直接整合核心逻辑)

st.title("📱 智能巡场助手 (规定联动版)")

# 加载规则库
all_rules = load_technical_rules()

# --- 问题录入区域优化 ---
st.subheader("📷 录入巡场问题")

# 1. 🔍 关键词检索挂载
search_kw = st.text_input("🔍 搜索技术规定关键词 (如：防水、坡道、高度)", placeholder="输入关键字后，下方自动匹配...")
selected_rule = ""

if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("🎯 匹配到的规定 (点击选择自动填入下方):", matched)
        selected_rule = f"【依据】：{chosen}\n【现场实际情况】："
    else:
        st.warning("未找到匹配规定，请手动录入。")

# 2. 自动填入与自主完善
# 注意：我们将 selected_rule 的值作为文本框的默认值，但仅在用户搜索时改变
# 为了实现“选完后还能改”，我们使用 session_state 记录描述
if "prob_desc" not in st.session_state:
    st.session_state.prob_desc = ""

if selected_rule:
    st.session_state.prob_desc = selected_rule

desc = st.text_area("问题描述（支持语音输入）", value=st.session_state.prob_desc, height=150)
st.session_state.prob_desc = desc # 同步更新

# ... (后续你的其余 PPT 生成逻辑不变)