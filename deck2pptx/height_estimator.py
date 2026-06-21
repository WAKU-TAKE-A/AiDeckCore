from pptx.util import Inches
from .models import Text, BulletList, Image, Table, Gallery, Flow, Split, CodeBlock, Mermaid, Tree

ELEMENT_GAP = Inches(0.15)          # uniform gap between consecutive elements
_TEXT_LINE_HEIGHT = Inches(0.35)    # estimated height per wrapped line of body text
_BULLET_LINE_HEIGHT = Inches(0.32)  # estimated height per bullet item line
_TEXT_MIN_HEIGHT = Inches(0.4)      # minimum text box height (approx 1 line)
_BULLET_MIN_HEIGHT = Inches(0.8)    # minimum bullet list height

def _estimate_element_height(element, content_width):
    """Return estimated rendered height (EMU) for an element."""
    if getattr(element, 'height_hint', None) is not None:
        return element.height_hint

    if isinstance(element, Text):
        estimated_lines = element.content.count('\n') + 1 + len(element.content) // 60
        return max(_TEXT_MIN_HEIGHT, estimated_lines * _TEXT_LINE_HEIGHT)

    if isinstance(element, BulletList):
        from .models import ListItem
        weighted = sum(
            1.0 if (item.level if isinstance(item, ListItem) else 0) == 0 else 0.8
            for item in element.items
        )
        return max(_BULLET_MIN_HEIGHT, weighted * _BULLET_LINE_HEIGHT)

    if isinstance(element, CodeBlock):
        lines = len(element.code.splitlines()) if element.code else 1
        box_h = Inches(max(1.0, lines * 0.25 + 0.2))
        caption_h = Inches(0.4) if (getattr(element, 'caption', None) or getattr(element, 'language', None)) else 0
        return caption_h + box_h
        
    if isinstance(element, Mermaid):
        from . import mermaid_handler
        if mermaid_handler.has_mermaid_cli():
            return getattr(element, 'height_hint', None) or Inches(3.0)
        else:
            lines = len(element.code.splitlines()) if element.code else 1
            box_h = Inches(max(1.0, lines * 0.25 + 0.2))
            return box_h

    if isinstance(element, Tree):
        def count_leaves(node):
            if not node.children:
                return 1
            return sum(count_leaves(child) for child in node.children)
        leaf_count = count_leaves(element.root)
        node_height = Inches(0.4)
        vertical_gap = Inches(0.15)
        return max(Inches(1.0), leaf_count * node_height + (leaf_count - 1) * vertical_gap)

    # Image, Gallery, Table, Flow, Comparison, Timeline, Split:
    # height depends on runtime data — caller handles these separately.
    return Inches(1.0)

def get_adjusted_height(elements_list, current_idx, total_bottom_y, current_y, content_width):
    """Estimate available height for elements_list[current_idx]."""
    remaining_h = total_bottom_y - current_y
    if remaining_h < Inches(1): remaining_h = Inches(1)

    from . import mermaid_handler
    has_mmdc = mermaid_handler.has_mermaid_cli()

    # Reserve height for all text-like elements that come AFTER the current one.
    future_elements = elements_list[current_idx + 1:]
    reserved_text_h = sum(
        _estimate_element_height(e, content_width) + ELEMENT_GAP
        for e in future_elements
        if isinstance(e, (Text, BulletList, CodeBlock)) or (isinstance(e, Mermaid) and not has_mmdc)
    )

    remaining_imgs = sum(
        1 for e in elements_list[current_idx:]
        if isinstance(e, (Image, Gallery, Split, Table)) or (isinstance(e, Mermaid) and has_mmdc)
    )

    available_img_h = remaining_h - reserved_text_h
    if available_img_h < Inches(1): available_img_h = Inches(1)

    current_element = elements_list[current_idx]
    if (isinstance(current_element, (Image, Gallery, Split, Table)) or (isinstance(current_element, Mermaid) and has_mmdc)) and remaining_imgs > 0:
        return available_img_h / remaining_imgs
    else:
        return remaining_h
