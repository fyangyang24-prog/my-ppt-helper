import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 辅助函数 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 2. 状态初始化 (必须在程序最上方) ---
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_buffer" not in st.session_state: st.session_state.desc_buffer = ""

# --- 3. 核心回调函数 (强制更新逻辑) ---
def update_desc_with_rule():
    # 获取下拉框选中的值并直接存入缓冲区
    if "rule_selector" in st.session_state:
        st.session_state.desc_buffer = f"【技术规定依据】：{st.session_state.rule_selector}\n【现场实际情况说明】：\n"

# --- 4. 主 UI 界面 ---
st.set_page_config(page_title="现场巡场助手", layout="centered")
st.title("📱 现场巡场助理 (完美修复版)")

# 公共信息
project_title = st.text_input("项目名称", value="独立路壹号项目")
check_date = st.date_input("检查时间")

# 拍照
uploaded_file = st.file_uploader("📷 拍摄/上传照片", type=["jpg", "jpeg", "png"])

# 技术规定检索
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        # 这里用 rule_selector 作为 key，并绑定回调函数 update_desc_with_rule
        st.selectbox("选择条文：", matched, key="rule_selector")
        st.button("✅ 插入到问题描述", on_click=update_desc_with_rule)

# 问题描述区 (使用 key="main_desc" 和 session_state 绑定)
# 这里的 value 总是指向 desc_buffer，无论页面如何刷新，值都不会丢失
desc = st.text_area("问题描述", value=st.session_state.desc_buffer, key="main_desc", placeholder="请录入现场描述...")

# 实时同步：当用户在 text_area 中手动输入时，更新 buffer
if desc != st.session_state.desc_buffer:
    st.session_state.desc_buffer = desc

# 其他字段
solve = st.text_area("解决措施")
duty = st.text_input("责任人")
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)

# 确认添加
if st.button("➕ 确认并添加此条问题"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "desc": desc, "solve": solve, "duty": duty, "decision": decision, 
        "deadline": str(check_date), "img_base64": img_b64
    })
    st.session_state.desc_buffer = "" # 清空缓冲区
    st.success("🎉 添加成功！")
    st.rerun()

st.divider()

# PPT 生成按钮 (占位，你可以填回你原来的生成函数)
if st.button("🚀 生成汇总 PPT"):
    st.write("调用生成逻辑...")