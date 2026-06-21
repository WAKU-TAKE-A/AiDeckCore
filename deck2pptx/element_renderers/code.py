from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ..render_context import SlideContext
from ..height_estimator import ELEMENT_GAP

def render(element, ctx: SlideContext, x, y, w, h) -> float:
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    start_x = ph.left if ph else x
    start_y = ph.top if ph else y
    width = ph.width if ph else w
    
    if element.caption or element.language:
        tb = ctx.slide.shapes.add_textbox(start_x, start_y, width, Inches(0.4))
        p = tb.text_frame.paragraphs[0]
        p.text = element.caption if element.caption else f"Language: {element.language}"
        p.font.name = ctx.theme.font_name
        p.font.size = ctx.theme.size_body_small
        p.font.italic = True
        start_y += Inches(0.4)
    
    line_count = len(element.code.splitlines()) if element.code else 1
    box_height = Inches(max(1.0, line_count * 0.25 + 0.2))
    
    shape = ctx.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x, start_y, width, box_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = ctx.theme.color_surface
    shape.line.color.rgb = ctx.theme.color_border
    
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = element.code
    p.font.name = "Consolas"
    p.font.size = Pt(14)
    p.font.color.rgb = ctx.theme.color_text
    p.alignment = PP_ALIGN.LEFT
    
    if not ph:
        rendered_height = getattr(element, 'height_hint', None)
        if rendered_height is None:
            rendered_height = (Inches(0.4) if element.caption or element.language else 0) + box_height
        return y + rendered_height + ELEMENT_GAP
    return y
