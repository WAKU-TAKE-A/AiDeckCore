from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from ..render_context import SlideContext
from ..height_estimator import ELEMENT_GAP

def render(element, ctx: SlideContext, x, y, w, h) -> float:
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    num_cols = len(element.columns)
    if num_cols == 0:
        return y
    
    col_width = (ph.width if ph else w) / num_cols
    start_x = ph.left if ph else x
    start_y = ph.top if ph else y
    
    if element.title:
        tb = ctx.slide.shapes.add_textbox(start_x, start_y, col_width * num_cols, Inches(0.5))
        p = tb.text_frame.paragraphs[0]
        p.text = element.title
        p.font.name = ctx.theme.font_name
        p.font.size = ctx.theme.size_body
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        start_y += Inches(0.5)

    box_height = Inches(1.0)
    for i, col in enumerate(element.columns):
        col_x = start_x + (i * col_width)
        
        max_items = max((len(c.items) for c in element.columns), default=0)
        box_height = Inches(max(1.0, (1 + max_items) * 0.25 + 0.2))
        
        tb = ctx.slide.shapes.add_textbox(col_x, start_y, col_width, box_height)
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = col.label
        p.font.name = ctx.theme.font_name
        p.font.size = ctx.theme.size_body
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        
        for item in col.items:
            p = tf.add_paragraph()
            p.text = item
            p.level = 1
            p.font.name = ctx.theme.font_name
            p.font.size = ctx.theme.size_body_small
            
    if not ph:
        rendered_height = getattr(element, 'height_hint', None)
        if rendered_height is None:
            rendered_height = (Inches(0.5) if element.title else 0) + box_height
        return y + rendered_height + ELEMENT_GAP
    return y

def render_timeline(element, ctx: SlideContext, x, y, w, h) -> float:
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    start_x = ph.left if ph else x
    start_y = ph.top if ph else y
    width = ph.width if ph else w
    
    event_height = Inches(0.8)
    for i, ev in enumerate(element.events):
        ev_y = start_y + (i * event_height)
        
        # Label
        tb = ctx.slide.shapes.add_textbox(start_x, ev_y, Inches(1.5), event_height)
        p = tb.text_frame.paragraphs[0]
        p.text = ev.label
        p.font.name = ctx.theme.font_name
        p.font.size = ctx.theme.size_body
        p.font.bold = True
        p.font.color.rgb = ctx.theme.color_flow_line
        
        # Line
        line = ctx.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x + Inches(1.6), ev_y + Inches(0.1), Inches(0.05), event_height - Inches(0.2))
        line.fill.solid()
        line.fill.fore_color.rgb = ctx.theme.color_flow_line
        line.line.color.rgb = ctx.theme.color_flow_line
        
        # Content
        tb = ctx.slide.shapes.add_textbox(start_x + Inches(1.8), ev_y, width - Inches(1.8), event_height)
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = ev.title
        p.font.name = ctx.theme.font_name
        p.font.size = ctx.theme.size_body
        p.font.bold = True
        p.font.color.rgb = ctx.theme.color_flow_text
        
        if ev.description:
            p2 = tf.add_paragraph()
            p2.text = ev.description
            p2.font.name = ctx.theme.font_name
            p2.font.size = ctx.theme.size_body_small
            p2.font.color.rgb = ctx.theme.color_text_light
            
    if not ph:
        rendered_height = getattr(element, 'height_hint', None)
        if rendered_height is None:
            rendered_height = len(element.events) * event_height
        return y + rendered_height + ELEMENT_GAP
    return y
