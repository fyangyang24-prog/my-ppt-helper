import streamlit as st
import os
import copy
import json
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from streamlit_javascript import st_javascript

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
                            
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                temp_pic_path = f"temp_prob_{index}.jpg"
                with open(temp_pic_path, "wb") as f:
                    f.write(img_bytes)
                current_slide.shapes.add_picture(temp_pic_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except Exception as e:
                pass
            
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path

# --- 📱 现场巡场助理 (防刷新断线自动暂存版) ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (抗断线稳定版)")

# --- 💾 本地持久化核心逻辑 ---
# 尝试利用 JavaScript 从浏览器本地缓存中读取历史数据
try:
    local_data_json = st_javascript("localStorage.getItem('xuncha_backup');")
    if local_data_json and local_data_json != "null" and "saved_problem_list" not in st.session_state:
        backup = json.loads(local_data_json)
        st.session_state.problem_list = backup.get("problem_list", [])
        st.session_state.saved_project_title = backup.get("project_title", "独立路壹号项目")
        st.session_state.saved_user = backup.get("user", "樊洋洋")
        st.session_state.saved_problem_list = True
except Exception as e:
    pass

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# 固定备选人员名单
team_members = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]

# --- 1. 公共信息区域 ---
st.subheader("🏢 第一步：填写项目公共信息")
default_title = st.session_state.get("saved_project_title", "独立路壹号项目")
project_title = st.text_input("项目名称", value=default_title)

saved_user_name = st.session_state.get("saved_user", "樊洋洋")
default_user_idx = team_members.index(saved_user_name) if saved_user_name in team_members else team_members.index("✍️ 手动输入...")

selected_user = st.selectbox("检查人", options=team_members, index=default_user_idx)
if selected_user == "✍️ 手动输入...":
    user_input_val = st.text_input("✍️ 请输入实际检查人姓名", value=saved_user_name if saved_user_name not in team_members else "", placeholder="例如：新同事名字", key="user_manual_widget")
    final_user = user_input_val
else:
    final_user = selected_user

check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# --- 2. 问题录入区域 ---
current_prob_idx = len(st.session_state.problem_list)
st.subheader(f"📷 第二步：录入巡场问题 (当前已录入 {current_prob_idx} 个问题)")

uploaded_file = st.file_uploader("📷 拍摄/上传当前问题照片", type=["jpg", "jpeg", "png"], key=f"file_{current_prob_idx}")
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...", key=f"desc_{current_prob_idx}")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...", key=f"solve_{current_prob_idx}")

selected_duty = st.selectbox("责任人", options=team_members, index=7, key=f"duty_select_{current_prob_idx}")
if selected_duty == "✍️ 手动输入...":
    duty_input_val = st.text_input("✍️ 请输入实际责任人/总包单位", value="", placeholder="例如：某某中建总包/李四", key=f"duty_manual_widget_{current_prob_idx}")
    final_duty = duty_input_val
else:
    final_duty = selected_duty

decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True, key=f"decision_{current_prob_idx}")
decision_text = "整改  √ \n 不整改 ▢" if decision_choice == "整改" else "整改  ▢ \n 不整改 √"

deadline_date = st.date_input("要求完成时间", key=f"deadline_{current_prob_idx}")
deadline_str = deadline_date.strftime("%Y/%m/%d")

# 暂存按钮
if st.button("➕ 确认并添加此条问题到列表"):
    final_user = final_user.strip() if final_user else ""
    final_duty = final_duty.strip() if final_duty else ""
    
    if not final_user:
        st.error("⚠️ 检查人不能为空！")
    elif not final_duty:
        st.error("⚠️ 责任人不能为空！")
    else:
        # 图片转换为 Base64 字符串以支持 JSON 本地持久化保存
        img_base64 = ""
        if uploaded_file is not None:
            img_base64 = base64.b64encode(uploaded_file.getbuffer()).decode("utf-8")

        prob_data = {
            "img_base64": img_base64,
            "desc": desc,
            "solve": solve,
            "duty": final_duty,
            "decision": decision_text,
            "deadline": deadline_str
        }
        st.session_state.problem_list.append(prob_data)
        
        # 🌟 核心：每次成功添加，立刻同步写入浏览器本地硬盘缓存，防止断网刷新丢失
        backup_data = {
            "project_title": project_title,
            "user": final_user,
            "problem_list": st.session_state.problem_list
        }
        js_save = f"localStorage.setItem('xuncha_backup', '{json.dumps(backup_data, ensure_ascii=False)}');"
        st_javascript(js_save)
        
        st.success(f"🎉 成功！第 {len(st.session_state.problem_list)} 个问题已成功装箱并自动本地备份！")
        st.rerun()

st.divider()

# --- 3. 汇总与清空区域 ---
st.subheader("🚀 第三步：一键打包汇总并下载")

if len(st.session_state.problem_list) > 0:
    st.write("📋 当前已安全暂存的问题列表：")
    for i, p in enumerate(st.session_state.problem_list):
        st.info(f"问题 {i+1}: {p['desc'][:20]}... (责任人: {p['duty']})")
        
    if st.button("🚀 一键打包生成完整多页 PPT 报告"):
        if not os.path.exists("template.pptx"):
            st.error("❌ 错误：请确保云端有名为 template.pptx 的模板文件！")
        else:
            with st.spinner("正在将所有问题进行多页批量排版中..."):
                out_file = build_multi_page_ppt(project_title, final_user, date_str, st.session_state.problem_list)
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
        # 清空浏览器本地缓存
        st_javascript("localStorage.removeItem('xuncha_backup');")
        if "saved_problem_list" in st.session_state:
            del st.session_state["saved_problem_list"]
        st.success("暂存箱及本地备份已清空。")
        st.rerun()
else:
    st.warning("⚠️ 暂存箱目前是空的，请至少在上方添加一个问题后再生成 PPT。")