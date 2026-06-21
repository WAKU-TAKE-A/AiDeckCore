import os
from ..render_context import SlideContext
from ..models import Image, CodeBlock

def render(element, ctx: SlideContext, x, y, w, h) -> float:
    from . import code, image
    from .. import mermaid_handler
    
    ph = ctx.find_placeholder(getattr(element, 'placeholder', None))
    if mermaid_handler.has_mermaid_cli():
        tmp_img = None
        try:
            tmp_img = mermaid_handler.render_to_temp_image(element.code)
            img_elem = Image(source=tmp_img, placeholder=getattr(element, 'placeholder', None), height_hint=getattr(element, 'height_hint', None))
            return image.render(img_elem, ctx, x, y, w, h)
        except Exception as e:
            cb = CodeBlock(code=f"[Mermaid Error: {e}]\n{element.code}", placeholder=getattr(element, 'placeholder', None), height_hint=getattr(element, 'height_hint', None))
            return code.render(cb, ctx, x, y, w, h)
        finally:
            if tmp_img and os.path.exists(tmp_img):
                try:
                    os.unlink(tmp_img)
                except OSError:
                    pass
    else:
        cb = CodeBlock(code=element.code, placeholder=getattr(element, 'placeholder', None), height_hint=getattr(element, 'height_hint', None))
        return code.render(cb, ctx, x, y, w, h)
