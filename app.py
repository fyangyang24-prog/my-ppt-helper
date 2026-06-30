import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# ... (safe_replace_text, duplicate_slide 函数保持原样不变) ...

# --- 核心：多页 PPT 填空与生成发动机 (同步分类字段) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        data = {
            "{{title}}": project_title,
            "{{user}}": user,
            "{{date}}": date_str,
            "{{type}}": prob.get("type", "普通问题"), # 新增：填充分类
            "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        # ... (替换逻辑保持原样) ...
        # (请确保此处代码段与您原有的替换逻辑一致)
        # ...
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 界面 ---
# ... (初始化代码保持原样) ...

# --- 2. 问题录入区域 (修改版) ---
current_prob_idx = len(st.session_state.problem_list)
st.subheader(f"📷 第二步：录入巡场问题")

# 【新增】单选项
issue_type = st.radio("巡场问题分类", 
                      options=["底线", "严控事项", "设计红条", "设计黑条", "图纸错漏碰缺", "落地效果问题"], 
                      horizontal=True, key="issue_type_radio")

uploaded_file = st.file_uploader("📷 拍摄照片", type=["jpg", "jpeg", "png"])
desc = st.text_area("问题描述", placeholder="请录入现场问题描述...")
solve = st.text_area("解决措施", placeholder="请录入整改要求与措施...")
selected_duty = st.selectbox("责任人", options=team_members, index=7)
final_duty = st.text_input("✍️ 手动输入责任人") if selected_duty == "✍️ 手动输入..." else selected_duty
decision_choice = st.radio("整改决定", options=["整改", "不整改"], horizontal=True)
deadline_date = st.date_input("要求完成时间")

if st.button("➕ 确认并添加此条问题到列表"):
    if not final_user or not final_duty:
        st.error("⚠️ 检查人和责任人不能为空！")
    else:
        img_base64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
        prob_data = {
            "type": issue_type, # 存入分类
            "img_base64": img_base64,
            "desc": desc,
            "solve": solve,
            "duty": final_duty,
            "decision": "整改  √" if decision_choice=="整改" else "不整改 √",
            "deadline": deadline_date.strftime("%Y/%m/%d")
        }
        st.session_state.problem_list.append(prob_data)
        st.success(f"✅ 第 {len(st.session_state.problem_list)} 条已加入暂存箱！")
        st.rerun() # 强制刷新页面