import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 保持你原有的工具函数不变 ---
try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text: run.text = run.text.replace(key, value)
            paragraph.text = paragraph.text.replace(key, value)

def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    # ... (保持你原有的 PPT 填充逻辑) ...
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 主程序逻辑 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")

# 初始化 Session State
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_val" not in st.session_state: st.session_state.desc_val = ""

st.title("📱 现场巡场助理")

# 1. 公共信息 (保持不变)
project_title = st.text_input("项目名称", value="独立路壹号项目")
# ... (其余人员名单及输入框) ...

# 2. 录入区域 (BUG 修复点：确保逻辑平铺，不被嵌套)
st.subheader("📷 录入巡场问题")
uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "jpeg", "png"])

# 修复点 1：智能检索逻辑平铺，确保点击立刻生效
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched, key="rule_select")
        if st.button("✅ 插入条文到描述框"):
            st.session_state.desc_val = f"【技术规定依据】：{chosen}\n【现场实况】：\n"
            st.rerun()

# 保持你原有的输入框绑定
desc = st.text_area("问题描述", value=st.session_state.desc_val, key="desc_input")
st.session_state.desc_val = desc 
solve = st.text_area("解决措施")
# ... (保持原有的责任人/整改决定等逻辑) ...

# 修复点 2：确认添加按钮（移至最外层，确保可见）
if st.button("➕ 确认并添加此条问题"):
    # 你的原有添加逻辑
    st.session_state.problem_list.append({"desc": desc, "solve": solve, ...})
    st.session_state.desc_val = "" # 清空临时值
    st.success("🎉 添加成功！")
    st.rerun()

st.divider()

# 3. 汇总与下载 (修复点 3：即使列表为空，按钮也不消失，只在点击时检查)
st.subheader("🚀 一键打包生成 PPT")
if st.button("🚀 生成汇总报告"):
    if len(st.session_state.problem_list) == 0:
        st.warning("暂存箱为空，请先添加问题！")
    else:
        with st.spinner("生成中..."):
            out_file = build_multi_page_ppt(...)
            if out_file:
                st.session_state.final_file = out_file

# 单独的下载按钮显示
if "final_file" in st.session_state:
    with open(st.session_state.final_file, "rb") as f:
        st.download_button("📥 下载 PPT (如微信无法下载，请点右上角在浏览器打开)", f, "报告.pptx")