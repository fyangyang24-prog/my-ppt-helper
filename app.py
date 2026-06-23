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

# --- 核心：深度克隆幻灯片的底层函数（确保样式边框不丢） ---
def duplicate_slide(prs, source_slide):
    # 使用 PPT 底层布局克隆一页新的幻灯片
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    
    # 将原页面的所有元素（表格、文本框）复制到新页面
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
    # 模板自带的第一页作为我们的“源动力页面”
    source_slide = prs.slides[0]
    
    # 遍历用户在手机上录入的所有问题列表
    for index, prob in enumerate(problem_list):
        # 如果是第一条问题，直接用模板自带的第一页；如果是后续问题，克隆一页新 PPT
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = duplicate_slide(prs, source_slide)
            
        # 准备当前页面的填空数据
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
        
        # 对当前页面进行文字替换
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items():
                    safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items():
                            safe_replace_text(cell.text_frame, key, val)
                            
        # 📸 处理当前页面的照片
        if prob["img_bytes"] is not None:
            # 临时存下该页面的照片
            temp_pic_path = f"temp_prob_{index}.jpg"
            with open(temp_pic_path, "wb") as f:
                f.write(prob["img_bytes"])
            # 精准贴到左侧
            current_slide.shapes.add_picture(temp_pic_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 📱 终极多页手机网页界面 ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (多页批量汇总版)")

# 使用 Streamlit 的 session_state 来在后台建立一个“临时问题暂存箱”
if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 1. 公共信息区域（整份 PPT 共享） ---
st.subheader("🏢 第一步：填写项目公共信息")
project_title = st.text_input("项目名称", value="独立路壹号项目")

user_options = ["樊洋洋", "刘璐", "其他人员"]
selected_user = st.selectbox("检查人", options=user_options, index=0)
if selected_user == "其他人员":
    user = st.text_input("请输入实际检查人姓名")
else:
    user = selected_user

check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# --- 2. 问题录入区域（支持循环添加） ---
st.subheader(f"📷 第二步：录入巡场问题 (当前已录入 {len(st.session_state.problem_list)} 个问题)")

uploaded_file = st.file_uploader("📷 拍摄/上传当前问题照片", type=["jpg", "jpeg", "png"], key=f"file_{len(st.session_state.problem_list)}")
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...", key=f"desc_{len(st.session_state.problem_list)}")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...", key=f"solve_{len(st.session_state.problem_list)}")

duty_options = ["樊洋洋", "刘璐", "外部单位/其他人"]
selected_duty = st.selectbox("责任人", options=duty_options, index=0, key=f"duty_opt_{len(st.session_state.problem_list)}")
if selected_duty == "外部单位/其他人":
    duty = st.text_input("请输入实际责任人姓名", key=f"duty_name_{len(st.session_state.problem_list)}")
else:
    duty = selected_duty

decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True, key=f"decision_{len(st.session_state.problem_list)}")
if decision_choice == "整改":
    decision_text = "整改  √ \n 不整改 ▢"
else:
    decision_text = "整改  ▢ \n 不整改 √"

deadline_date = st.date_input("要求完成时间", key=f"deadline_{len(st.session_state.problem_list)}")
deadline_str = deadline_date.strftime("%Y/%m/%d")

# 🌟 核心：【暂存并添加下一个】按钮
if st.button("➕ 确认并添加此条问题到列表"):
    # 把当前填写的这一页数据打包存进列表
    prob_data = {
        "img_bytes": uploaded_file.getbuffer() if uploaded_file is not None else None,
        "desc": desc,
        "solve": solve,
        "duty": duty,
        "decision": decision_text,
        "deadline": deadline_str
    }
    st.session_state.problem_list.append(prob_data)
    st.success(f"🎉 成功！第 {len(st.session_state.problem_list)} 个问题已成功装箱！请继续在上方录入下一个问题。")
    st.rerun() # 刷新网页清空当前框，方便录入下一条

st.divider()

# --- 3. 汇总与清空区域 ---
st.subheader("🚀 第三步：一键打包汇总并下载")

# 展示当前已经添加的所有问题简述，让你心里有数
if len(st.session_state.problem_list) > 0:
    st.write("📋 当前暂存箱内的问题列表：")
    for i, p in enumerate(st.session_state.problem_list):
        st.info(f"问题 {i+1}: {p['desc'][:20]}... (责任人: {p['duty']})")
        
    if st.button("🚀 一键打包生成完整多页 PPT 报告"):
        if not os.path.exists("template.pptx"):
            st.error("❌ 错误：请确保云端有名为 template.pptx 的模板文件！")
        else:
            with st.spinner("正在将所有问题进行多页批量排版中..."):
                out_file = build_multi_page_ppt(project_title, user, date_str, st.session_state.problem_list)
                if out_file:
                    st.success("🎉 完美！多页汇总 PPT 已成功合体！")
                    with open(out_file, "rb") as file:
                        st.download_button(
                            label="📥 立即点击下载多页汇总 PPT 到手机",
                            data=file,
                            file_name=f"{project_title}-汇总巡场报告.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        
    # 允许一键清空重来
    if st.button("🗑️ 清空暂存箱（重新开始）", type="secondary"):
        st.session_state.problem_list = []
        st.success("暂存箱已清空。")
        st.rerun()
else:
    st.warning("⚠️ 暂存箱目前是空的，请至少在上方添加一个问题后再生成 PPT。")