from pptx.util import Pt
from ..render_context import SlideContext
from ..height_estimator import _estimate_element_height, ELEMENT_GAP

def render(element, ctx: SlideContext, x, y, w, h) -> float:
    from ..render_utils import _set_text_frame_text
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    if ph and ph.has_text_frame:
        _set_text_frame_text(ph.text_frame, element.content)
        return y
    else:
        target_x = ph.left if ph else x
        target_y = ph.top if ph else y
        target_w = ph.width if ph else w
        rendered_height = _estimate_element_height(element, target_w)
        txBox = ctx.slide.shapes.add_textbox(target_x, target_y, target_w, rendered_height)
        tf = txBox.text_frame
        tf.word_wrap = True
        _set_text_frame_text(
            tf,
            element.content,
            font_name=ctx.theme.font_name,
            font_size=Pt(ctx.deck.font_size_l0) if ctx.deck.font_size_l0 else ctx.theme.size_body,
            font_color=ctx.theme.color_text,
        )
        if not ph:
            return y + rendered_height + ELEMENT_GAP
        return y

def render_bullet(element, ctx: SlideContext, x, y, w, h) -> float:
    from ..models import ListItem
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    if ph and ph.has_text_frame:
        tf = ph.text_frame
        tf.clear()
        for i, item in enumerate(element.items):
            is_li = isinstance(item, ListItem)
            text = item.text if is_li else str(item)
            level = item.level if is_li else 0
            
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = text
            if level > 0: p.level = level
            
            font_size = None
            if level == 0 and ctx.deck.font_size_l0: font_size = ctx.deck.font_size_l0
            elif level == 1 and ctx.deck.font_size_l1: font_size = ctx.deck.font_size_l1
            elif level == 2 and ctx.deck.font_size_l2: font_size = ctx.deck.font_size_l2
            elif level == 3 and ctx.deck.font_size_l3: font_size = ctx.deck.font_size_l3
            elif level >= 4 and ctx.deck.font_size_l4: font_size = ctx.deck.font_size_l4
            
            if font_size:
                p.font.size = Pt(font_size)
        return y
    else:
        target_x = ph.left if ph else x
        target_y = ph.top if ph else y
        target_w = ph.width if ph else w
        rendered_height = _estimate_element_height(element, target_w)
        txBox = ctx.slide.shapes.add_textbox(target_x, target_y, target_w, rendered_height)
        tf = txBox.text_frame
        tf.word_wrap = True
        for i, item in enumerate(element.items):
            is_li = isinstance(item, ListItem)
            text = item.text if is_li else str(item)
            level = item.level if is_li else 0
            
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = text
            if level > 0: p.level = level
            p.font.name = ctx.theme.font_name
            
            font_size = None
            if level == 0 and ctx.deck.font_size_l0: font_size = ctx.deck.font_size_l0
            elif level == 1 and ctx.deck.font_size_l1: font_size = ctx.deck.font_size_l1
            elif level == 2 and ctx.deck.font_size_l2: font_size = ctx.deck.font_size_l2
            elif level == 3 and ctx.deck.font_size_l3: font_size = ctx.deck.font_size_l3
            elif level >= 4 and ctx.deck.font_size_l4: font_size = ctx.deck.font_size_l4
            
            p.font.size = Pt(font_size) if font_size else (ctx.theme.size_body if level == 0 else ctx.theme.size_body_small)
            p.font.color.rgb = ctx.theme.color_text
        if not ph:
            return y + rendered_height + ELEMENT_GAP
        return y
