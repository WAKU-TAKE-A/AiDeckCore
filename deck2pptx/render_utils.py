def _set_text_frame_text(text_frame, content, font_name=None, font_size=None, font_color=None):
    text_frame.clear()
    lines = str(content).split('\n')
    for i, line in enumerate(lines):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.text = line
        if font_name:
            p.font.name = font_name
        if font_size:
            p.font.size = font_size
        if font_color:
            p.font.color.rgb = font_color
