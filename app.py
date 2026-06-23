import streamlit as st
import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import base64

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

# --- 📱 现场巡场助理 (极致修复稳定版) ---
st.set_page_config(page_title="设计师巡场助手", layout="centered")
st.title("📱 现场巡场助理 (完美画笔版)")

if "problem_list" not in st.session_state:
    st.session_state.problem_list = []

# --- 1. 公共信息区域 ---
st.subheader("🏢 第一步：填写项目公共信息")
project_title = st.text_input("项目名称", value="独立路壹号项目")

user_options = ["樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "刘璐", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]
selected_user = st.selectbox("检查人", options=user_options, index=0)
user = st.text_input("请输入实际检查人姓名", value="", placeholder="例：张三") if selected_user == "✍️ 手动输入..." else selected_user

check_date = st.date_input("检查时间")
date_str = check_date.strftime("%Y/%m/%d")

st.divider()

# --- 2. 问题录入与照片标注区域 ---
st.subheader(f"📷 第二步：录入巡场问题并标注 (当前已录入 {len(st.session_state.problem_list)} 个)")

# 拍照上传
bg_image_file = st.file_uploader("📷 第一步：拍摄/上传现场照片", type=["jpg", "jpeg", "png"], key=f"bg_file_{len(st.session_state.problem_list)}")

marked_image_bytes = None

if bg_image_file is not None:
    st.write("🎨 **第二步：请在下方照片上滑动手指进行标注**")
    
    # 🌟 终极修复方案：转成 Base64 网页链接喂给画板，彻底解决底层崩溃
    bg_image = Image.open(bg_image_file)
    w, h = bg_image.size
    display_width = 400
    display_height = int(h * (display_width / w))
    bg_image_resized = bg_image.resize((display_width, display_height))
    
    buffered = io.BytesIO()
    bg_image_resized.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    b64_bg_url = f"data:image/png;base64,{img_str}"
    
    # 标注工具栏选择
    tool_mode = st.radio("选择标注工具：", ("画笔(自由画)", "矩形红框", "拉线/箭头", "橡皮擦"), index=0, horizontal=True)
    
    drawing_mode = "freedraw"
    if tool_mode == "矩形红框": drawing_mode = "rect"
    elif tool_mode == "拉线/箭头": drawing_mode = "line"
    elif tool_mode == "橡皮擦": drawing_mode = "transform"

    # 唤起手机触屏画板（使用 b64_bg_url 完美绕过兼容报错）
    canvas_result = st_canvas(
        fill_color="rgba(255, 0, 0, 0)", 
        stroke_width=3,
        stroke_color="rgb(255, 0, 0)",  
        background_image=b64_bg_url, # 👈 注入安全通道
        update_streamlit=True,
        height=display_height,
        width=display_width,
        drawing_mode=drawing_mode,
        key=f"canvas_{len(st.session_state.problem_list)}"
    )
    
    # 实时捕获画完后的图像数据并拼接
    if canvas_result.image_data is not None:
        canvas_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
        final_marked_img = bg_image_resized.convert("RGBA")
        final_marked_img = Image.alpha_composite(final_marked_img, canvas_img).convert("RGB")
        
        img_byte_arr = io.BytesIO()
        final_marked_img.save(img_byte_arr, format='JPEG')
        marked_image_bytes = img_byte_arr.getvalue()
        st.caption("✨ 标注已实时同步，可直接点击下方添加。")

st.divider()
desc = st.text_area("问题描述", value="", placeholder="请录入现场问题描述（支持手机语音转文字）...", key=f"desc_{len(st.session_state.problem_list)}")
solve = st.text_area("解决措施", value="", placeholder="请录入整改要求与措施...", key=f"solve_{len(st.session_state.problem_list)}")

duty_options = ["刘璐", "樊洋洋", "付长春", "李新宇", "顾宇", "王硕", "郝思仆", "张晓莉", "吕山", "王凤国", "夏友强", "✍️ 手动输入..."]
selected_duty = st.selectbox("责任人", options=duty_options, index=0, key=f"duty_opt_{len(st.session_state.problem_list)}")
duty = st.text_input("请输入实际责任人姓名", value="", placeholder="例：某某总包单位", key=f"duty_name_{len(st.session_state.problem_list)}") if selected_duty == "✍️ 手动输入..." else selected_duty

decision_choice = st.radio("整改决定", options=["整改", "不整改"], index=0, horizontal=True, key=f"decision_{len(st.session_state.problem_list)}")
decision_text = "整改  √ \n 不整改 ▢" if decision_choice == "整改" else "整改  ▢ \n 不整改 √"

deadline_date = st.date_input("要求完成时间", key=f"deadline_{len(st.session_state.problem_list)}")
deadline_str = deadline_date.strftime("%Y/%m/%d")

# 暂存并添加
if st.button("➕ 确认并添加此条问题到列表"):
    if not duty or not user:
        st.error("⚠️ 请确保检查人和责任人的姓名已填写完整！")
    else:
        # 如果用户上传了图但没有做任何画笔点击，默认用缩放后的原图
        if marked_image_bytes is None and bg_image_file is not None:
            img_byte_arr = io.BytesIO()
            bg_image_resized.convert("RGB").save(img_byte_arr, format='JPEG')
            marked_image_bytes = img_byte_arr.getvalue()

        prob_data = {
            "img_bytes": marked_image_bytes,  
            "desc": desc,
            "solve": solve,
            "duty": duty,
            "decision": decision_text,
            "deadline": deadline_str
        }
        st.session_state.problem_list.append(prob_data)
        st.success(f"🎉 成功！带标注的第 {len(st.session_state.problem_list)} 个问题已成功装箱！")
        st.rerun()

st.divider()

# --- 3. 汇总与下载区域 ---
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