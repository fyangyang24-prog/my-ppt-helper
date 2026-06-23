import streamlit as st
import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt

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

# --- 核心：深度克隆幻灯片的底层函数 ---
def duplicate_slide(prs, source_slide):
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        el = shape.element
        new_el = copy.deepcopy(el)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

# --- 核心：多页 PPT 填空与生成发动机 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path):
        return None
        
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = duplicate_slide(prs, source_slide)
            
        data = {
            "{{title}}": project_title,
            "{{user}}": user,
            "{{date}}": date_str,
            "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"],
            "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"],
            "{{decision}}": prob["decision"]
        }
        
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items():
                    safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items():
                            safe_replace_text(cell.text_frame, key, val)
                            
        if prob["img_bytes"] is not None:
            temp_pic_path = f"temp_prob_{index}.jpg"
            with open(temp_pic_path, "wb") as f:
                f.write(prob["img_bytes"])
            current_slide.shapes.add_picture(temp_pic_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 📱 现场巡场助理 (下拉自适应极简版) ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (多页稳定版)")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# 固定备选人员名单
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]

# --- 1. 公共信息区域 ---
st.subheader("🏢 第一步：填写项目公共信息")
project_title = st.text_input("项目名称", value="独立路壹号项目")

selected_user = st.selectbox("检查人", options=team_members, index=0)
if selected_user == "✍️ 手动输入...":
    user = st.text_input("✍️ 请输入实际检查人姓名", value="", placeholder="例如：新同事名字")
else:
    user = selected_user

check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# --- 2. 问题录入区域 ---
current_prob_idx = len(st.session_state.problem_list)
st.subheader(f"📷 第二步：录入巡场问题 (当前已录入 {current_prob_idx} 个问题)")

uploaded_file = st.file_uploader("📷 拍摄/上传当前问题照片", type=["jpg", "jpeg", "png"], key=f"file_{current_prob_idx}")
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...", key=f"desc_{current_prob_idx}")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...", key=f"solve_{current_prob_idx}")

# 责任人下拉菜单，默认第 7 位（刘璐）
selected_duty = st.selectbox("责任人", options=team_members, index=7, key=f"duty_select_{current_prob_idx}")
if selected_duty == "✍️ 手动输入...":
    duty = st.text_input("✍️ 请输入实际责任人/总包单位", value="", placeholder="例如：某某中建总包/李四", key=f"duty_input_{current_prob_idx}")
else:
    duty = selected_duty

decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True, key=f"decision_{current_prob_idx}")
decision_text = "整改  √ \n 不整改 ▢" if decision_choice == "整改" else "整改  ▢ \n 不整改 √"

deadline_date = st.date_input("要求完成时间", key=f"deadline_{current_prob_idx}")
deadline_str = deadline_date.strftime("%Y/%m/%d")

# 暂存按钮
if st.button("➕ 确认并添加此条问题到列表"):
    if not duty or not user:
        st.error("⚠️ 请确保检查人和责任人的姓名已填写完整！")