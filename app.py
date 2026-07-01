# --- 核心：多页 PPT 填空与生成发动机 ---
def build_multi_page_ppt(project_title, user, date_str, problem_list):
    template_path = "template.pptx"
    if not os.path.exists(template_path): return None
        
    prs = Presentation(template_path)
    source_slide = prs.slides[0]
    
    # 清空除第一页以外的页面（如果需要重新生成）
    # 或者直接基于第一页的布局添加新幻灯片
    for index, prob in enumerate(problem_list):
        if index == 0:
            current_slide = source_slide
        else:
            current_slide = prs.slides.add_slide(source_slide.slide_layout)
            # 手动复制source_slide的所有形状到新页面
            for shape in source_slide.shapes:
                new_shape = current_slide.shapes.add_shape(shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height)
                if shape.has_text_frame:
                    new_shape.text = shape.text # 复制原文本占位符
        
        # 准备当前页的数据
        data = {
            "{{title}}": project_title, "{{user}}": user, "{{date}}": date_str,
            "{{type}}": prob.get("type", "无"), "{{desc}}": prob["desc"],
            "{{solve}}": prob["solve"], "{{duty}}": prob["duty"],
            "{{deadline}}": prob["deadline"], "{{decision}}": prob["decision"]
        }
        
        # 对当前页面进行精确替换
        for shape in current_slide.shapes:
            if shape.has_text_frame:
                for key, val in data.items():
                    safe_replace_text(shape.text_frame, key, val)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, val in data.items():
                            safe_replace_text(cell.text_frame, key, val)
        # ... (图片逻辑保持不变)