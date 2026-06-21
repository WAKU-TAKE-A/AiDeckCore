from ..models import (Text, BulletList, Image, Gallery, Table,
                      Flow, Comparison, Timeline, CodeBlock, Mermaid, Tree, Split)
from . import text, image, table, flow, tree, mermaid, code, comparison, split

def render_element(element, ctx, x, y, w, h) -> float:
    if isinstance(element, Text):         return text.render(element, ctx, x, y, w, h)
    if isinstance(element, BulletList):   return text.render_bullet(element, ctx, x, y, w, h)
    if isinstance(element, Image):        return image.render(element, ctx, x, y, w, h)
    if isinstance(element, Gallery):      return image.render_gallery(element, ctx, x, y, w, h)
    if isinstance(element, Table):        return table.render(element, ctx, x, y, w, h)
    if isinstance(element, Flow):         return flow.render(element, ctx, x, y, w, h)
    if isinstance(element, CodeBlock):    return code.render(element, ctx, x, y, w, h)
    if isinstance(element, Mermaid):      return mermaid.render(element, ctx, x, y, w, h)
    if isinstance(element, Tree):         return tree.render(element, ctx, x, y, w, h)
    if isinstance(element, Split):        return split.render(element, ctx, x, y, w, h)
    if type(element).__name__ == 'Comparison': return comparison.render(element, ctx, x, y, w, h)
    if type(element).__name__ == 'Timeline':   return comparison.render_timeline(element, ctx, x, y, w, h)
    return y  # unknown: no-op
