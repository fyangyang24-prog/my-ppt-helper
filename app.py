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

# --- PPT 生成逻辑 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    # ... (保持原有的生成逻辑) ...
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 主程序 UI ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")

# 初始化 Session State
if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_content" not in st.session_state: st.session_state.desc_content = ""

st.title("📱 现场巡场助理")

# 1. 公共信息
project_title = st.text_input("项目名称", value="独立路壹号项目")
user_name = st.text_input("检查人")
check_date = st.date_input("检查时间")

# 2. 拍照放到最前面
uploaded_file = st.file_uploader("📷 拍摄/上传照片", type=["jpg", "jpeg", "png"])

# 3. 检索区 (逻辑修复：通过定义函数确保填入)
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")

if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched, key="rule_select")
        
        # 修复点：定义一个回调函数，专门负责更新描述框的内容
        def insert_text():
            st.session_state.desc_content = f"【技术规定依据】：{chosen}\n【现场实况说明】：\n"
        
        st.button("✅ 插入条文到描述框", on_click=insert_text)

# 4. 描述框 (使用 session_state 同步)
# 当用户在框内手动输入时，更新 buffer；当点击插入按钮时，更新 buffer 并显示
desc = st.text_area("问题描述", value=st.session_state.desc_content, key="desc_input")
st.session_state.desc_content = desc # 实时同步到 buffer

solve = st.text_area("解决措施")
duty = st.text_input("责任人")
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)

# 5. 添加与导出按钮
if st.button("➕ 确认并添加此条问题"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": duty, "decision": decision, "deadline": str(check_date)
    })
    st.session_state.desc_content = "" # 添加后清空
    st.success("🎉 添加成功！")
    st.rerun()

st.divider()

if st.button("🚀 生成汇总 PPT"):
    out_file = build_multi_page_ppt(project_title, user_name, str(check_date), st.session_state.problem_list)
    if out_file:
        with open(out_file, "rb") as f:
            st.download_button("📥 下载 PPT (若微信无法下载，请点右上角选择：在浏览器打开)", f, "报告.pptx")