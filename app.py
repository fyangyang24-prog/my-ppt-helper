import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 💡 工具支持 ---
try:
    from streamlit_javascript import st_javascript
    JS_AVAILABLE = True
except ImportError:
    JS_AVAILABLE = False

# --- 核心：确保文本替换保留格式 ---
def safe_replace_text(text_frame, key, value):
    for paragraph in text_frame.paragraphs:
        if key in paragraph.text:
            # 记录原格式
            orig_font_name = paragraph.runs[0].font.name if paragraph.runs else "Microsoft YaHei"
            orig_font_size = paragraph.runs[0].font.size if paragraph.runs else Pt(14)
            orig_font_color = paragraph.runs[0].font.color.rgb if paragraph.runs and paragraph.runs[0].font.color else None
            
            # 替换文字
            paragraph.text = paragraph.text.replace(key, str(value))
            
            # 还原格式
            for run in paragraph.runs:
                run.font.name = orig_font_name
                run.font.size = orig_font_size
                if orig_font_color:
                    run.font.color.rgb = orig_font_color

# --- 核心：克隆幻灯片 ---
def duplicate_slide(prs, source_slide):
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(el)
    return new_slide

# --- 核心：PPT 生成引擎 (修复了每一页数据覆盖的问题) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
        
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        # 1. 获取当前要操作的页面
        current_slide = source_slide if index == 0 else duplicate_slide(prs, source_slide)
            
        # 2. 准备数据
        data = {
            "{{title}}": project_title,
            "{{user}}": user,
            "{{date}}": date_str,
            "{{type}}": prob.get("type", ""),
            "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        
        # 3. 填充该页面的所有形状
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items():
                    safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.has_text_frame:
                            for key, val in data.items():
                                safe_replace_text(cell.text_frame, key, val)
                                    
        # 4. 插入图片
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_pic = f"temp_{index}.jpg"
                with open(temp_pic, "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(temp_pic, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
            
    output_path = "最终汇总报告.pptx"
    prs.save(output_path)
    return output_path

# --- 界面逻辑 ---
st.title("📱 现场巡场助理")

if "problem_list" not in st.session_state: st.session_state.problem_list = []

# 输入区
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

prob_type = st.text_input("问题分类", placeholder="（如底线、严控事项、设计红黑条，图纸错漏碰缺、落地效果问题等）")
uploaded_file = st.file_uploader("照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "type": prob_type, "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": final_duty, "decision": f"{decision} √", "deadline": deadline
    })
    st.rerun()

if st.button("🚀 生成 PPT"):
    if not st.session_state.problem_list:
        st.error("请先添加问题！")
    else:
        out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
        with open(out, "rb") as f:
            st.download_button("📥 点击下载报告", f, file_name="巡场报告.pptx")