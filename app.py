import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 确保 session_state 初始化，防止页面白屏 ---
if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 2. 安全文本替换函数 (兼容表格和文本框) ---
def force_replace_text(shape, key, value):
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))
    elif hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if hasattr(cell, "has_text_frame") and cell.has_text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, str(value))

# --- 3. 稳健的 PPT 生成引擎 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        if index == 0:
            current_slide = source_slide
        else:
            # 创建新页面，使用模板的第一页布局
            current_slide = prs.slides.add_slide(source_slide.slide_layout)
            # 安全复制模板中的形状，跳过会导致 AttributeError 的复杂对象
            for shape in source_slide.shapes:
                if hasattr(shape, "auto_shape_type"):
                    try:
                        new_shape = current_slide.shapes.add_shape(
                            shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height
                        )
                        if shape.has_text_frame: new_shape.text = shape.text
                    except: continue

        # 准备填充数据
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{type}}": prob.get("type", "普通"), "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"], "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
        }
        
        # 填充当前页所有形状
        for shape in current_slide.shapes:
            for key, val in data.items():
                force_replace_text(shape, key, val)
        
        # 插入图片
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_path = f"temp_{index}.jpg"
                with open(temp_path, "wb") as f: f.write(img_bytes)
                current_slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except: pass
            
    output_path = "最终报告.pptx"
    prs.save(output_path)
    return output_path

# --- 4. 界面展示 ---
st.title("📱 现场巡场助理")
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")

# 修复后的日期输入 (注意：必须在同一行完成 .strftime，避免语法错误)
check_date_obj = st.date_input("检查时间")
check_date = check_date_obj.strftime("%Y/%m/%d")

prob_type = st.radio("问题分类", ["底线", "严控事项", "设计红黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True)
uploaded_file = st.file_uploader("照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision = st.radio("整改", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
    st.session_state.problem_list.append({
        "type": prob_type, "img_base64": img_b64, "desc": desc, "solve": solve, 
        "duty": final_duty, "decision": f"{decision} √", "deadline": deadline
    })
    st.rerun()

if st.button("🚀 生成报告"):
    out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
    if out:
        with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="报告.pptx")