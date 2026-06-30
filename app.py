import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 💡 安全引入持久化组件 ---
try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

# --- 核心：无损替换单个页面文字的函数 ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
            if key in paragraph.text:
                orig_font_name = paragraph.runs[0].font.name if paragraph.runs else "Microsoft YaHei"
                orig_font_size = paragraph.runs[0].font.size if paragraph.runs else Pt(14)
                orig_font_color = paragraph.runs[0].font.color.rgb if paragraph.runs and paragraph.runs[0].font.color else None
                paragraph.text = paragraph.text.replace(key, value)
                for run in paragraph.runs:
                    run.font.name = orig_font_name
                    run.font.size = orig_font_size
                    if orig_font_color:
                        run.font.color.rgb = orig_font_color

# --- 核心：深度克隆幻灯片 ---
def duplicate_slide(prs, source_slide):
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in source_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shape.element))
    return new_slide

# --- 核心：修复后的多页 PPT 生成逻辑 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        # 每一页都必须基于模板进行操作，不是第一页则克隆
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
        
        # 关键修复：确保每一页使用各自独立的问题数据
        data = {
            "{{title}}": project_title,
            "{{user}}": user,
            "{{date}}": date_str,
            "{{type}}": prob.get("type", "普通问题"), # 新增分类映射
            "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        
        # 执行独立页面的文本替换
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items(): safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items(): safe_replace_text(cell.text_frame, key, val)
        
        # 插入图片
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_path = f"temp_{index}.jpg"
                with open(temp_path, "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
            
    prs.save("最终汇总巡场报告.pptx")
    return "最终汇总巡场报告.pptx"

# --- 界面 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理")

if "problem_list" not in st.session_state: st.session_state.problem_list = []

team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]

project_title = st.text_input("项目名称", value="独立路壹号项目")
selected_user = st.selectbox("检查人", options=team_members)
final_user = st.text_input("✍️ 输入检查人") if selected_user == "✍️ 手动输入..." else selected_user
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

st.divider()

# 录入区域
st.subheader(f"📷 录入巡场问题 ({len(st.session_state.problem_list)}个)")
# 【新增】巡场问题分类单选
issue_type = st.radio("巡场问题分类", options=["底线", "严控事项", "设计红条", "设计黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True)

uploaded_file = st.file_uploader("📷 拍摄照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
selected_duty = st.selectbox("责任人", options=team_members, index=7)
final_duty = st.text_input("✍️ 输入责任人") if selected_duty == "✍️ 手动输入..." else selected_duty
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "type": issue_type, # 存入分类
        "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": final_duty, "decision": f"{decision} √", "deadline": deadline
    })
    st.rerun()

if st.button("🚀 一键打包生成 PPT"):
    out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
    if out:
        with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="报告.pptx")
    else: st.error("未找到模板文件！")