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

# --- 📱 现场巡场助理 (完美输入版) ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (多页稳定版)")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# 全局常用备选人员名单
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强"]

# --- 1. 公共信息区域 ---
st.subheader("🏢 第一步：填写项目公共信息")
project_title = st.text_input("项目名称", value="独立路壹号项目")

# 🌟 修复方案：直接用 text_input 接收最终名字，同时提供快捷联想按钮
st.write("✍️ **检查人** (可在下方直接输入任意新名字，或点击快捷键填入)")
user = st.text_input("当前检查人姓名", value="樊洋洋", label_visibility="collapsed")
# 快捷小标签按钮，点一下就能把名字飞进输入框
user_cols = st.columns(6)
for idx, member in enumerate(team_members[:12]):
    with user_cols[idx % 6]:
        if st.button(member, key=f"user_btn_{member}"):
            user = member
            st.rerun()

check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# --- 2. 问题录入区域 ---
st.subheader(f"📷 第二步：录入巡场问题 (当前已录入 {len(st.session_state.problem_list)} 个问题)")

uploaded_file = st.file_uploader("📷 拍摄/上传当前问题照片", type=["jpg", "jpeg", "png"], key=f"file_{len(st.session_state.problem_list)}")
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...", key=f"desc_{len(st.session_state.problem_list)}")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...", key=f"solve_{len(st.session_state.problem_list)}")

# 🌟 修复方案：责任人同样改为“可输可点”模式，完美支持名单外的任何人或总包单位
st.write("✍️ **责任人** (直接在下方输入任意新名字/单位，或点击快捷键填入)")
current_prob_idx = len(st.session_state.problem_list)
duty = st.text_input("当前责任人姓名/单位", value="刘璐", key=f"duty_raw_{current_prob_idx}", label_visibility="collapsed")

duty_cols = st.columns(6)
for idx, member in enumerate(team_members):
    with duty_cols[idx % 6]:
        if st.button(member, key=f"duty_btn_{member}_{current_prob_idx}"):
            st.session_state[f"duty_raw_{current_prob_idx}"] = member
            st.rerun()

decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True, key=f"decision_{current_prob_idx}")
decision_text = "整改  √ \n 不整改 ▢" if decision_choice == "整改" else "整改  ▢ \n 不整改 √"

deadline_date = st.date_input("要求完成时间", key=f"deadline_{current_prob_idx}")
deadline_str = deadline_date.strftime("%Y/%m/%d")

# 暂存按钮
if st.button("➕ 确认并添加此条问题到列表"):
    # 从session_state里获取最新打字输入的值，防止丢失
    final_user = user
    final_duty = st.session_state.get(f"duty_raw_{current_prob_idx}", duty)
    
    if not final_duty or not final_user:
        st.error("⚠️ 请确保检查人和责任人的姓名已填写完整！")
    else:
        prob_data = {
            "img_bytes": uploaded_file.getbuffer() if uploaded_file is not None else None,
            "desc": desc,
            "solve": solve,
            "duty": final_duty,
            "decision": decision_text,
            "deadline": deadline_str
        }
        st.session_state.problem_list.append(prob_data)
        st.success(f"🎉 成功！第 {len(st.session_state.problem_list)} 个问题已成功装箱！")
        st.rerun()

st.divider()

# --- 3. 汇总与清空区域 ---
st.subheader("🚀 第三步：一键打包汇总并下载")

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
                        
    if st.button("🗑️ 清空暂存箱（重新开始）", type="secondary"):
        st.session_state.problem_list = []
        st.success("暂存箱已清空。")
        st.rerun()
else:
    st.warning("⚠️ 暂存箱目前是空的，请至少在上方添加一个问题后再生成 PPT。")