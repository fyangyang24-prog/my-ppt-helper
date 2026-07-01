import streamlit as st
import os
import copy
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

# --- 初始化 ---
st.set_page_config(page_title="现场巡场助理", layout="centered")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 核心：安全文本替换 ---
def force_replace_text(shape, key, value):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.has_text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, str(value))

# --- 核心：PPT 生成引擎 (修复克隆 Bug) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        # 第一页直接用，后续页克隆
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = prs.slides.add_slide(source_slide.slide_layout)
            # 安全地复制形状
            for shape in source_slide.shapes:
                try:
                    # 只复制有明确形状类型的元素
                    if hasattr(shape, "auto_shape_type"):
                        new_shape = current_slide.shapes.add_shape(
                            shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height
                        )
                        if shape.has_text_frame:
                            new_shape.text = shape.text
                except:
                    continue # 跳过无法复制的复杂元素
        
        # 填充数据
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{desc}}": prob["desc"], "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"], "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        
        for shape in current_slide.shapes:
            force_replace_text(shape, "{{title}}", project_title)
            force_replace_text(shape, "{{user}}", user)
            force_replace_text(shape, "{{date}}", date_str)
            force_replace_text(shape, "{{desc}}", prob["desc"])
            force_replace_text(shape, "{{solve}}", prob["solve"])
            force_replace_text(shape, "{{duty}}", prob["duty"])
            force_replace_text(shape, "{{deadline}}", prob["deadline"])
            force_replace_text(shape, "{{decision}}", prob["decision"])
        
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

# --- 界面 ---
st.title("📱 现场巡场助理")

project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

st.divider()
st.subheader(f"📷 录入问题 (已暂存 {len(st.session_state.problem_list)} 条)")

uploaded_file = st.file_uploader("拍摄照片", type=["jpg", "png"])
desc = st.text_area("问题描述")
solve = st.text_area("解决措施")
final_duty = st.text_input("责任人")
decision = st.radio("整改决定", ["整改", "不整改"], horizontal=True)
deadline = st.date_input("完成时间").strftime("%Y/%m/%d")

if st.button("➕ 确认并添加"):
    if not desc or not final_duty:
        st.error("请完善信息！")
    else:
        img_b64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8") if uploaded_file else ""
        st.session_state.problem_list.append({
            "img_base64": img_b64, "desc": desc, "solve": solve, 
            "duty": final_duty, "decision": decision, "deadline": deadline
        })
        st.rerun()

if st.button("🚀 生成 PPT 报告"):
    if not st.session_state.problem_list:
        st.error("请先添加问题！")
    else:
        out = build_multi_page_ppt(project_title, final_user, check_date, st.session_state.problem_list)
        if out:
            with open(out, "rb") as f: st.download_button("📥 下载 PPT", f, file_name="最终报告.pptx")