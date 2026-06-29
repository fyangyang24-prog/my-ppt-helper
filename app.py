import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 辅助函数 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 初始化 ---
st.set_page_config(page_title="现场巡场助手", layout="centered")

if "desc_text" not in st.session_state:
    st.session_state.desc_text = ""

# --- UI 录入区 ---
project_title = st.text_input("项目名称", value="独立路壹号项目")

# 智能检索 (这是关键)
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched, key="rule_sel")
        if st.button("✅ 插入到问题描述"):
            # 直接赋值给 session_state
            st.session_state.desc_text = f"【技术规定依据】：{chosen}\n【现场实况说明】：\n"
            st.rerun()

# 问题描述区
desc = st.text_area("问题描述", value=st.session_state.desc_text, key="main_desc_area")

# 关键修复：如果你在框里输入，实时更新 session_state
if desc != st.session_state.desc_text:
    st.session_state.desc_text = desc

# 下面的解决措施、责任人等保持不变...
solve = st.text_area("解决措施")
# ... 省略其他代码 ...