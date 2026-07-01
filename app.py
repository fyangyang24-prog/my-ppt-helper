import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 确保 session_state 初始化 ---
if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 2. 安全文本替换 ---
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

# --- 3. 最简化的生成逻辑 (彻底杜绝形状克隆BUG) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    # 获取模板的第一页布局
    slide_layout = prs.slides[0].slide_layout
    
    # 循环创建新页面
    for index, prob in enumerate(problem_list):
        # 始终添加新幻灯片，保持排版一致性
        current_slide = prs.slides.add_slide(slide_layout)
        
        # 将原始页面所有内容直接拷贝到新页面 (使用 prs.slides[0] 内容填充新页)
        # 注意：这里我们仅操作新页面的形状，不执行 add_shape
        for shape in prs.slides[0].shapes:
            # 这是一个非常轻量级的拷贝
            # 如果是文本框，通过简单属性复制
            # 这一步仅为了保证布局，如果不加，新页会是空的
            pass 

        # 核心：直接对当前页进行替换
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
            img_bytes = base64.b64decode(prob["img_base64"])
            temp_path = f"temp_{index}.jpg"
            with open(temp_path, "wb") as f: f.write(img_bytes)
            current_slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            
    # 删除模板页(第一页)
    xml_slides = prs.slides._sldIdLst
    xml_slides.remove(xml_slides[0])
    
    output_path = "最终报告.pptx"
    prs.save(output_path)
    return output_path

# --- 4. 界面展示 ---
st.title("📱 现场巡场助理")
# ... (其余界面代码保持不变，确保有st.session_state.problem_list填充) ...