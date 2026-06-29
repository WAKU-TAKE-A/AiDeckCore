def _set_text_frame_text(text_frame, content, font_name=None, font_size=None, font_color=None):
    text_frame.clear()
    single_para_text = str(content).replace('\n', '\x0b')
    p = text_frame.paragraphs[0]
    p.text = single_para_text
    if font_name:
        p.font.name = font_name
    if font_size:
        p.font.size = font_size
    if font_color:
        p.font.color.rgb = font_color
