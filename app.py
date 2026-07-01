import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 确保 session_state 初始化 ---
if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 2. 安全处理文本替换 ---
def force_replace_text(shape, key, value):
    # 检查普通文本框
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))
    # 检查表格单元格
    elif hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if hasattr(cell, "has_text_frame") and cell.has_text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, str(value))

# --- 3. 修复后的 PPT 生成逻辑 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        if index == 0:
            current_slide = source_slide
        else:
            # 创建新页面并复制模板布局
            current_slide = prs.slides.add_slide(source_slide.slide_layout)
            # 安全遍历模板形状
            for shape in source_slide.shapes:
                # 仅在对象具有形状类型时尝试添加
                if hasattr(shape, "auto_shape_type"):
                    try:
                        new_shape = current_slide.shapes.add_shape(
                            shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height
                        )
                        if shape.has_text_frame: new_shape.text = shape.text
                    except: continue

        # 填充数据
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{type}}": prob.get("type", "普通"), "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"], "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
        }
        
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
check_date = st.date_input("检查时间").strftime("%Y/%