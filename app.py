import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Pt

# --- 1. 核心辅助函数 ---
def load_technical_rules():
    if not os.path.exists("rules.txt"): return []
    with open("rules.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f.readlines() if line.strip()]

# --- 2. PPT 生成引擎 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    if not os.path.exists("template.pptx"): return None
    prs = Presentation("template.pptx")
    # ... 此处保留你原有的 PPT 生成逻辑 ...
    output_path = "summary_report.pptx"
    prs.save(output_path)
    return output_path

# --- 3. 状态初始化 ---
st.set_page_config(page_title="现场巡场助手", layout="centered")

if "problem_list" not in st.session_state: st.session_state.problem_list = []
if "desc_content" not in st.session_state: st.session_state.desc_content = ""

# --- 4. 页面布局与逻辑 ---
st.title("📱 现场巡场助理")

# 公共信息 (不使用 session_state 赋值)
project_title = st.text_input("项目名称", value="独立路壹号项目")
check_person = st.selectbox("检查人", ["樊洋洋", "付长春", "李新宇", "顾宇", "✍️ 手动输入..."])
check_date = st.date_input("检查时间")

# 拍照/上传 (放到前面)
uploaded_file = st.file_uploader("📷 拍摄/上传照片", type=["jpg", "jpeg", "png"])

# 技术规定检索 (独立逻辑，不产生状态冲突)
all_rules = load_technical_rules()
search_kw = st.text_input("🔍 搜索技术规定")
if search_kw:
    matched = [r for r in all_rules if search_kw in r]
    if matched:
        chosen = st.selectbox("选择条文：", matched)
        if st.button("✅ 插入条文到描述框"):
            st.session_state.desc_content = f"【技术规定依据】：{chosen}\n【现场实况说明】：\n"
            st.rerun()

# 录入区 (通过 key 自动同步)
desc = st.text_area("问题描述", value=st.session_state.desc_content, key="desc_input")
solve = st.text_area("解决措施")
duty = st.selectbox("责任人", ["樊洋洋", "付长春", "李新宇", "顾宇", "✍️ 手动输入..."])
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)

# 确认添加按钮 (核心逻辑)
if st.button("➕ 确认并添加此条问题"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": duty, "decision": decision, "deadline": str(check_date)
    })
    st.session_state.desc_content = "" # 清空缓存
    st.success("🎉 添加成功！")
    st.rerun()

st.divider()

# 汇总生成 PPT
if st.button("🚀 生成汇总 PPT"):
    out_file = build_multi_page_ppt(project_title, check_person, str(check_date), st.session_state.problem_list)
    if out_file:
        with open(out_file, "rb") as f:
            st.download_button(
                label="📥 下载 PPT (微信内无法下载请用浏览器打开)", 
                data=f, 
                file_name="巡场报告.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )