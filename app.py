import streamlit as st
import os
import base64
from pptx import Presentation
from pptx.util import Inches

# --- 1. 基础配置 ---
st.set_page_config(page_title="现场巡场助理", layout="centered")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 2. 增强型文本替换 ---
def replace_in_text_frame(text_frame, data):
    """替换文本框或单元格中的文本"""
    for paragraph in text_frame.paragraphs:
        for key, val in data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(val))

def process_shape(shape, data):
    """检查形状类型并进行替换"""
    if shape.has_text_frame:
        replace_in_text_frame(shape.text_frame, data)
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.has_text_frame:
                    replace_in_text_frame(cell.text_frame, data)

# --- 3. 最稳健生成逻辑 ---
def build_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
    
    prs = Presentation(template_path)
    
    # 遍历每一个录入的问题
    for i, prob in enumerate(problem_list):
        # 假设模板中每页对应一个问题
        if i < len(prs.slides):
            slide = prs.slides[i]
            
            # 准备数据字典
            data = {
                "{{title}}": project_title,
                "{{user}}": user,
                "{{date}}": date_str,
                "{{type}}": prob.get("type", ""), # 这里处理分类
                "{{desc}}": prob["desc"],
                "{{solve}}": prob["solve"],
                "{{duty}}": prob["duty"],
                "{{deadline}}": prob["deadline"],
                "{{decision}}": prob["decision"]
            }
            
            # 执行替换
            for shape in slide.shapes:
                process_shape(shape, data)
            
            # 添加图片
            if prob.get("img_base64"):
                try:
                    img_bytes = base64.b64decode(prob["img_base64"])
                    temp_path = f"temp_{i}.jpg"
                    with open(temp_path, "wb") as f: f.write(img_bytes)
                    slide.shapes.add_picture(temp_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
                except Exception as e:
                    st.warning(f"图片插入失败: {e}")
                    
    output_path = "巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 4. 界面展示 ---
st.title("📱 现场巡场助理")
project_title = st.text_input("项目名称", value="独立路壹号项目")
final_user = st.text_input("检查人", value="樊洋洋")
check_date = st.date_input("检查时间").strftime("%Y/%m/%d")

# 单选逻辑
prob_type = st.radio("问题分类", ["底线", "严控事项", "设计红黑条", "图纸错漏碰缺", "落地效果问题"], horizontal=True)
uploaded_file = st.file_uploader("拍摄现场照片", type=["jpg", "png"])
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

if st.button("🚀 生成报告"):
    if not st.session_state.problem_list:
        st.error("请先添加问题！")
    else: