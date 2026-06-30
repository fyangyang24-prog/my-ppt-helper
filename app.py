# --- 核心：多页 PPT 填空与生成发动机 (已修复循环填充逻辑) ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path):
        return None
        
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    for index, prob in enumerate(problem_list):
        # 核心逻辑：如果是第一页直接用，不是第一页就克隆第一页
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = duplicate_slide(prs, source_slide)
            
        # 明确每一个索引对应的具体数据
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
        
        # 遍历当前页面的所有形状进行替换，确保互不干扰
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items():
                    safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items():
                            safe_replace_text(cell.text_frame, key, val)
                            
        # 处理图片（确保每一页添加的是对应的图片）
        if prob.get("img_base64"):
            try:
                img_bytes = base64.b64decode(prob["img_base64"])
                # 这里给文件名增加索引，防止覆盖
                temp_pic_path = f"temp_prob_{index}.jpg"
                with open(temp_pic_path, "wb") as f:
                    f.write(img_bytes)
                current_slide.shapes.add_picture(temp_pic_path, Inches(0.52), Inches(2.3), width=Inches(2.95))
            except Exception:
                pass
            
    output_path = "最终汇总巡场报告.pptx"
    prs.save(output_path)
    return output_path