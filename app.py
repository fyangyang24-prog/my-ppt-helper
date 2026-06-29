import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Pt

# --- 核心辅助函数 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 1. 回调函数：处理点击事件，强制更新状态 ---
def callback_insert_rule():
    if "temp_rule" in st.session_state and st.session_state.temp_rule:
        st.session_state.desc_content = f"【技术规定依据】：{st.session_state.temp_rule}\n【现场实况说明】：\n"

def callback_add_problem():
    # 执行添加逻辑
    img_b64 = base64.b64encode(st.session_state.temp_img.getbuffer()).decode("utf-8") if st.session_state.temp_img else ""
    st.session_state.problem_list.append({
        "img_base64": img_b64, 
        "desc": st.session_state.desc_content, 
        "solve": st.session_state.solve_content, 
        "duty": st.session_state.duty_content, 
        "decision": st.session_state.decision_content, 
        "deadline": str(st.session_state.date_content)
    })
    st.session_state.desc_content = "" # 清空
    st.success("🎉 添加成功！")

# --- 2. 初始化 ---
st.set_page_config(page_title="现场巡场助手", layout="centered")
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_content" not in st.session_state: st.session_state.desc_content = ""

st.title("📱 现场巡场助理")

# 3. 布局区
project_title = st.text_input("项目名称", value="独立路壹号项目")
st.session_state.duty_content = st.selectbox("检查人", ["樊洋洋", "付长春", "李新宇", "顾宇"], key="duty_content")
st.session_state.date_content = st.date_input("检查时间")

# 拍照位置调整
st.file_uploader("📷 拍摄/上传照片", type=["jpg", "jpeg", "png"], key="temp_img")

# 技术规定检索
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        # 将选中的项存入 temp_rule
        st.selectbox("选择条文：", matched, key="temp_rule")
        st.button("✅ 插入条文到描述框", on_click=callback_insert_rule)

# 问题描述 (绑定 state)
st.text_area("问题描述", key="desc_content")
st.text_area("解决措施", key="solve_content")
st.radio("整改决定", ["整改", "不整改"], key="decision_content", horizontal=True)

# 核心按钮 (直接绑定回调，不使用 if 判断)
st.button("➕ 确认并添加此条问题", on_click=callback_add_problem)

# --- 4. 生成区 ---
st.divider()
if st.button("🚀 生成汇总 PPT"):
    # 此处调用你原有的生成 PPT 逻辑
    st.write("正在打包生成...")