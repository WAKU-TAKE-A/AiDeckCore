from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pathlib import Path
from .models import Deck, Slide, Text, BulletList, Image, Table, Gallery, Flow
from .layout import Layout, get_slide_layout_type
from .theme import Theme

def _remove_existing_slides(prs):
    """Use a template for layouts/theme without carrying starter slides forward."""
    slide_id_list = prs.slides._sldIdLst
    for slide_id in list(slide_id_list):
        prs.part.drop_rel(slide_id.rId)
        slide_id_list.remove(slide_id)

def _name_matches(actual_name: str, requested_name: str) -> bool:
    actual = actual_name.strip().casefold()
    requested = requested_name.strip().casefold()
    return bool(requested) and actual.startswith(requested)

def _name_equals(actual_name: str, requested_name: str) -> bool:
    return actual_name.strip().casefold() == requested_name.strip().casefold()

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

def render_deck(deck: Deck, output_path: str, base_dir: Path = Path('.'), template_path: str = None):
    # Initialize Presentation
    if template_path:
        prs = Presentation(template_path)
        _remove_existing_slides(prs)
    else:
        prs = Presentation()
    
    base_dir = Path(base_dir)
    theme = Theme(deck.theme)
    
    # Set slide dimensions
    if deck.orientation == 'portrait':
        prs.slide_width = Inches(7.5)
        prs.slide_height = Inches(13.333)
    else:
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
    layout = Layout(prs.slide_width, prs.slide_height)
    
    render_slides = list(deck.slides)
    
    if deck.toc:
        toc_title_text = deck.toc_title or "Table of Contents"
        toc_items = []
        for i, s in enumerate(deck.slides):
            if i == 0 and get_slide_layout_type(s) == "title":
                continue
            if s.title:
                toc_items.append(s.title)
        
        toc_slide = Slide(
            title=toc_title_text,
            layout_hint="Title and Content",
            elements=[BulletList(items=toc_items)]
        )
        
        if len(render_slides) > 0:
            render_slides.insert(1, toc_slide)
        else:
            render_slides.append(toc_slide)
    
    for slide_model in render_slides:
        layout_type = get_slide_layout_type(slide_model)
        
        # Determine layout
        slide_layout = None
        if slide_model.layout_hint:
            # Try to find by name
            for ly in prs.slide_layouts:
                if _name_matches(ly.name, slide_model.layout_hint):
                    slide_layout = ly
                    break
        
        if not slide_layout:
            # fallback to blank (index 6 usually) or title (index 0)
            if layout_type == "title" and len(prs.slide_layouts) > 0:
                slide_layout = prs.slide_layouts[0]
            elif len(prs.slide_layouts) > 6:
                slide_layout = prs.slide_layouts[6]
            elif len(prs.slide_layouts) > 0:
                slide_layout = prs.slide_layouts[-1]
                
        slide = prs.slides.add_slide(slide_layout)
        
        # Add slide notes if any
        if slide_model.notes and slide.has_notes_slide:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_model.notes
            
        def find_placeholder(name):
            if not name: return None
            for shape in slide.shapes:
                if _name_equals(shape.name, name):
                    return shape
            matching_layout_idx = None
            for layout_shape in slide.slide_layout.placeholders:
                if _name_matches(layout_shape.name, name):
                    matching_layout_idx = layout_shape.placeholder_format.idx
                    break
            if matching_layout_idx is not None:
                for shape in slide.placeholders:
                    if shape.placeholder_format.idx == matching_layout_idx:
                        return shape
            for shape in slide.shapes:
                if _name_matches(shape.name, name):
                    return shape
            return None

        # Try to use default title/subtitle placeholders if no manual coords
        title_ph = find_placeholder("Title")
        subtitle_ph = find_placeholder("Subtitle")
        for shape in slide.shapes:
            if shape.is_placeholder:
                if not title_ph and shape.placeholder_format.type in (1, 3): # TITLE or CENTER_TITLE
                    title_ph = shape
                elif not subtitle_ph and shape.placeholder_format.type == 4: # SUBTITLE
                    subtitle_ph = shape
        
        if slide_model.title:
            if title_ph:
                title_ph.text = slide_model.title
            else:
                title_y = layout.slide_height / 3 if layout_type == "title" else layout.title_y
                txBox = slide.shapes.add_textbox(layout.title_x, title_y, layout.title_width, Inches(1.5) if layout_type == "title" else layout.title_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = slide_model.title
                if layout_type == "title": p.alignment = PP_ALIGN.CENTER
                p.font.name = theme.font_name
                p.font.size = theme.size_title
                p.font.color.rgb = theme.color_primary if layout_type != "title" else theme.color_text
        
        if slide_model.subtitle:
            if subtitle_ph:
                subtitle_ph.text = slide_model.subtitle
            else:
                # If we made a manual title box, add paragraph there if title existed, else make new box
                # For simplicity, fallback to adding a new textbox below title
                sub_y = (layout.slide_height / 3 + Inches(1.5)) if layout_type == "title" else (layout.title_y + layout.title_height)
                txBox = slide.shapes.add_textbox(layout.title_x, sub_y, layout.title_width, Inches(1))
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = slide_model.subtitle
                if layout_type == "title": p.alignment = PP_ALIGN.CENTER
                p.font.name = theme.font_name
                p.font.size = theme.size_subtitle
                p.font.color.rgb = theme.color_text_light
        
        # Simple rendering layout flow
        current_y = layout.content_y
        content_x = layout.content_x
        def render_element(element, content_x, current_y, layout_content_width, layout_content_height, ph=None):
            ph = find_placeholder(getattr(element, 'placeholder', None))
            
            if isinstance(element, Text):
                if ph and ph.has_text_frame:
                    _set_text_frame_text(ph.text_frame, element.content)
                else:
                    target_x = ph.left if ph else content_x
                    target_y = ph.top if ph else current_y
                    target_w = ph.width if ph else layout_content_width
                    txBox = slide.shapes.add_textbox(target_x, target_y, target_w, Inches(1))
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    _set_text_frame_text(
                        tf,
                        element.content,
                        font_name=theme.font_name,
                        font_size=Pt(deck.font_size_l0) if deck.font_size_l0 else theme.size_body,
                        font_color=theme.color_text,
                    )
                    if not ph: current_y += Inches(0.5)
                
            elif isinstance(element, BulletList):
                from .models import ListItem
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
                        if level == 0 and deck.font_size_l0: font_size = deck.font_size_l0
                        elif level == 1 and deck.font_size_l1: font_size = deck.font_size_l1
                        elif level == 2 and deck.font_size_l2: font_size = deck.font_size_l2
                        elif level == 3 and deck.font_size_l3: font_size = deck.font_size_l3
                        elif level >= 4 and deck.font_size_l4: font_size = deck.font_size_l4
                        
                        if font_size:
                            p.font.size = Pt(font_size)
                else:
                    target_x = ph.left if ph else content_x
                    target_y = ph.top if ph else current_y
                    target_w = ph.width if ph else layout_content_width
                    txBox = slide.shapes.add_textbox(target_x, target_y, target_w, Inches(2))
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    for i, item in enumerate(element.items):
                        is_li = isinstance(item, ListItem)
                        text = item.text if is_li else str(item)
                        level = item.level if is_li else 0
                        
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = text
                        if level > 0: p.level = level
                        p.font.name = theme.font_name
                        
                        font_size = None
                        if level == 0 and deck.font_size_l0: font_size = deck.font_size_l0
                        elif level == 1 and deck.font_size_l1: font_size = deck.font_size_l1
                        elif level == 2 and deck.font_size_l2: font_size = deck.font_size_l2
                        elif level == 3 and deck.font_size_l3: font_size = deck.font_size_l3
                        elif level >= 4 and deck.font_size_l4: font_size = deck.font_size_l4
                        
                        p.font.size = Pt(font_size) if font_size else (theme.size_body if level == 0 else theme.size_body_small)
                        p.font.color.rgb = theme.color_text
                    if not ph: current_y += Inches(2)
                
            elif isinstance(element, Image):
                img_path = base_dir / element.source
                try:
                    if ph and hasattr(ph, 'insert_picture'):
                        pic = ph.insert_picture(str(img_path))
                        if element.caption:
                            tb = slide.shapes.add_textbox(pic.left, pic.top + pic.height, pic.width, Inches(0.3))
                            p = tb.text_frame.paragraphs[0]
                            p.text = element.caption
                            p.font.name = theme.font_name
                            p.font.size = Pt(12)
                            p.alignment = PP_ALIGN.CENTER
                            p.font.color.rgb = theme.color_text_light
                    else:
                        target_x = ph.left if ph else content_x
                        target_y = ph.top if ph else current_y
                        max_w = ph.width if ph else layout_content_width
                        max_h = ph.height if ph else layout_content_height
                        if element.caption:
                            max_h -= Inches(0.3)
                            
                        from PIL import Image as PILImage
                        try:
                            with PILImage.open(str(img_path)) as pil_img:
                                w, h = pil_img.size
                            ratio = min(max_w / w, max_h / h)
                            new_w = w * ratio
                            new_h = h * ratio
                        except:
                            new_w = max_w
                            new_h = None
                            
                        pic = slide.shapes.add_picture(str(img_path), target_x, target_y, width=new_w, height=new_h)
                        if element.caption:
                            tb = slide.shapes.add_textbox(target_x, target_y + (new_h if new_h else pic.height), new_w, Inches(0.3))
                            p = tb.text_frame.paragraphs[0]
                            p.text = element.caption
                            p.font.name = theme.font_name
                            p.font.size = Pt(12)
                            p.alignment = PP_ALIGN.CENTER
                            p.font.color.rgb = theme.color_text_light
                            
                        if not ph: current_y += (new_h if new_h else pic.height) + (Inches(0.4) if element.caption else Inches(0.1))
                except Exception as e:
                    print(f"Failed to load image {img_path}: {e}")
                    
            elif isinstance(element, Table):
                rows = len(element.rows) + 1 if element.headers else len(element.rows)
                cols = len(element.headers) if element.headers else (len(element.rows[0]) if element.rows else 1)
                
                target_x = ph.left if ph else content_x
                target_y = ph.top if ph else current_y
                target_w = ph.width if ph else layout_content_width
                
                table_shape = slide.shapes.add_table(rows, cols, target_x, target_y, target_w, Inches(1))
                table = table_shape.table
                
                row_offset = 0
                if element.headers:
                    for i, header in enumerate(element.headers):
                        table.cell(0, i).text = header
                    row_offset = 1
                        
                for r_idx, row in enumerate(element.rows):
                    for c_idx, cell in enumerate(row):
                        cell_obj = table.cell(r_idx + row_offset, c_idx)
                        cell_obj.text = str(cell)
                        for p in cell_obj.text_frame.paragraphs:
                            p.font.name = theme.font_name
                            p.font.size = Pt(deck.font_size_l1) if deck.font_size_l1 else theme.size_body_small
                        
                if not ph: current_y += Inches(2)
                    
            elif isinstance(element, Gallery):
                num_images = len(element.images)
                if num_images == 0:
                    return current_y
                    
                cols = getattr(element, 'columns', None)
                rows_ct = getattr(element, 'rows', None)
                
                if cols is None and rows_ct is None:
                    if num_images == 1: cols, rows_ct = 1, 1
                    elif num_images == 2: cols, rows_ct = 2, 1
                    elif num_images == 3: cols, rows_ct = 3, 1
                    elif num_images == 4: cols, rows_ct = 2, 2
                    else: cols, rows_ct = 3, 2
                elif cols is None:
                    import math
                    cols = math.ceil(num_images / rows_ct)
                elif rows_ct is None:
                    import math
                    rows_ct = math.ceil(num_images / cols)
                    
                cell_width = (ph.width if ph else layout_content_width) / cols
                cell_height = (ph.height if ph else layout_content_height) / rows_ct
                
                for i, img in enumerate(element.images[:cols*rows_ct]):
                    r = i // cols
                    c = i % cols
                    x = (ph.left if ph else content_x) + (c * cell_width)
                    y = (ph.top if ph else current_y) + (r * cell_height)
                    
                    max_w = cell_width - Inches(0.1)
                    max_h = cell_height - Inches(0.1)
                    if img.caption:
                        max_h -= Inches(0.3)
                        
                    img_path = base_dir / img.source
                    try:
                        from PIL import Image as PILImage
                        with PILImage.open(str(img_path)) as pil_img:
                            w, h = pil_img.size
                        ratio = min(max_w / w, max_h / h)
                        new_w = w * ratio
                        new_h = h * ratio
                        
                        center_x = x + (cell_width - new_w) / 2
                        center_y = y + (cell_height - new_h - (Inches(0.3) if img.caption else 0)) / 2
                        
                        slide.shapes.add_picture(str(img_path), center_x, center_y, width=new_w, height=new_h)
                        
                        if img.caption:
                            tb = slide.shapes.add_textbox(x, center_y + new_h, cell_width, Inches(0.3))
                            p = tb.text_frame.paragraphs[0]
                            p.text = img.caption
                            p.font.name = theme.font_name
                            p.font.size = Pt(12)
                            p.alignment = PP_ALIGN.CENTER
                            p.font.color.rgb = theme.color_text_light
                    except Exception as e:
                        print(f"Failed to load image {img_path}: {e}")
                
                if not ph: current_y += (cell_height * rows_ct)
                
            elif isinstance(element, Flow):
                node_width = Inches(1.5)
                node_height = Inches(0.5)
                
                node_positions = {}
                
                def style_flow_node(shape):
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = theme.color_flow_fill
                    shape.line.color.rgb = theme.color_flow_line
                    for p in shape.text_frame.paragraphs:
                        p.font.name = theme.font_name
                        p.font.size = theme.size_body_small
                        p.font.color.rgb = theme.color_flow_text
                
                def style_flow_arrow(shape):
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = theme.color_flow_line
                    shape.line.color.rgb = theme.color_flow_line
                
                if element.direction == 'horizontal':
                    for i, node in enumerate(element.nodes):
                        x = content_x + i * (node_width + Inches(0.5))
                        y = current_y
                        shape = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            x, y, node_width, node_height
                        )
                        shape.text = node.label
                        style_flow_node(shape)
                        node_positions[node.id] = (x + node_width, y + node_height / 2) # right center for out, left center for in
                        # Actually let's just store the right-center and left-center
                        node_positions[f"{node.id}_out"] = (x + node_width, y + node_height / 2)
                        node_positions[f"{node.id}_in"] = (x, y + node_height / 2)
                        
                    for edge in element.edges:
                        if f"{edge.from_node}_out" in node_positions and f"{edge.to_node}_in" in node_positions:
                            fx, fy = node_positions[f"{edge.from_node}_out"]
                            tx, ty = node_positions[f"{edge.to_node}_in"]
                            
                            if tx >= fx:
                                arrow_x = fx
                                arrow_width = max(tx - fx, Inches(0.1))
                                shape_type = MSO_SHAPE.RIGHT_ARROW
                            else:
                                arrow_x = tx
                                arrow_width = max(fx - tx, Inches(0.1))
                                shape_type = MSO_SHAPE.LEFT_ARROW
                                
                            arrow = slide.shapes.add_shape(
                                shape_type,
                                arrow_x, fy - Inches(0.1), arrow_width, Inches(0.2)
                            )
                            style_flow_arrow(arrow)
                else: # vertical
                    for i, node in enumerate(element.nodes):
                        x = content_x
                        y = current_y + i * (node_height + Inches(0.5))
                        shape = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            x, y, node_width, node_height
                        )
                        shape.text = node.label
                        style_flow_node(shape)
                        node_positions[f"{node.id}_out"] = (x + node_width / 2, y + node_height)
                        node_positions[f"{node.id}_in"] = (x + node_width / 2, y)
                
                    for edge in element.edges:
                        if f"{edge.from_node}_out" in node_positions and f"{edge.to_node}_in" in node_positions:
                            fx, fy = node_positions[f"{edge.from_node}_out"]
                            tx, ty = node_positions[f"{edge.to_node}_in"]
                            # Draw down/up arrow
                            if ty >= fy:
                                arrow_y = fy
                                arrow_height = max(ty - fy, Inches(0.1))
                                shape_type = MSO_SHAPE.DOWN_ARROW
                            else:
                                arrow_y = ty
                                arrow_height = max(fy - ty, Inches(0.1))
                                shape_type = MSO_SHAPE.UP_ARROW
                                
                            arrow = slide.shapes.add_shape(
                                shape_type,
                                fx - Inches(0.1), arrow_y, Inches(0.2), arrow_height
                            )
                            style_flow_arrow(arrow)

                current_y += Inches(3)

            elif type(element).__name__ == 'Comparison':
                num_cols = len(element.columns)
                if num_cols > 0:
                    col_width = (ph.width if ph else layout_content_width) / num_cols
                    start_x = ph.left if ph else content_x
                    start_y = ph.top if ph else current_y
                    
                    if element.title:
                        tb = slide.shapes.add_textbox(start_x, start_y, col_width * num_cols, Inches(0.5))
                        p = tb.text_frame.paragraphs[0]
                        p.text = element.title
                        p.font.name = theme.font_name
                        p.font.size = theme.size_body
                        p.font.bold = True
                        p.alignment = PP_ALIGN.CENTER
                        start_y += Inches(0.5)

                    for i, col in enumerate(element.columns):
                        x = start_x + (i * col_width)
                        
                        tb = slide.shapes.add_textbox(x, start_y, col_width, Inches(2))
                        tf = tb.text_frame
                        tf.word_wrap = True
                        
                        p = tf.paragraphs[0]
                        p.text = col.label
                        p.font.name = theme.font_name
                        p.font.size = theme.size_body
                        p.font.bold = True
                        p.alignment = PP_ALIGN.CENTER
                        
                        for item in col.items:
                            p = tf.add_paragraph()
                            p.text = item
                            p.level = 1
                            p.font.name = theme.font_name
                            p.font.size = theme.size_body_small
                            
                    if not ph: current_y += Inches(2.5)
            
            elif type(element).__name__ == 'Timeline':
                start_x = ph.left if ph else content_x
                start_y = ph.top if ph else current_y
                width = ph.width if ph else layout_content_width
                
                event_height = Inches(0.8)
                for i, ev in enumerate(element.events):
                    y = start_y + (i * event_height)
                    
                    # Label
                    tb = slide.shapes.add_textbox(start_x, y, Inches(1.5), event_height)
                    p = tb.text_frame.paragraphs[0]
                    p.text = ev.label
                    p.font.name = theme.font_name
                    p.font.size = theme.size_body
                    p.font.bold = True
                    p.font.color.rgb = theme.color_flow_line
                    
                    # Line
                    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x + Inches(1.6), y + Inches(0.1), Inches(0.05), event_height - Inches(0.2))
                    line.fill.solid()
                    line.fill.fore_color.rgb = theme.color_flow_line
                    line.line.color.rgb = theme.color_flow_line
                    
                    # Content
                    tb = slide.shapes.add_textbox(start_x + Inches(1.8), y, width - Inches(1.8), event_height)
                    tf = tb.text_frame
                    p = tf.paragraphs[0]
                    p.text = ev.title
                    p.font.name = theme.font_name
                    p.font.size = theme.size_body
                    p.font.bold = True
                    p.font.color.rgb = theme.color_flow_text
                    
                    if ev.description:
                        p2 = tf.add_paragraph()
                        p2.text = ev.description
                        p2.font.name = theme.font_name
                        p2.font.size = theme.size_body_small
                        p2.font.color.rgb = theme.color_text_light
                        
                if not ph: current_y += len(element.events) * event_height + Inches(0.5)

            elif type(element).__name__ == 'CodeBlock':
                start_x = ph.left if ph else content_x
                start_y = ph.top if ph else current_y
                width = ph.width if ph else layout_content_width
                
                if element.caption or element.language:
                    tb = slide.shapes.add_textbox(start_x, start_y, width, Inches(0.4))
                    p = tb.text_frame.paragraphs[0]
                    p.text = element.caption if element.caption else f"Language: {element.language}"
                    p.font.name = theme.font_name
                    p.font.size = theme.size_body_small
                    p.font.italic = True
                    start_y += Inches(0.4)
                
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x, start_y, width, Inches(2))
                shape.fill.solid()
                shape.fill.fore_color.rgb = theme.color_surface
                shape.line.color.rgb = theme.color_border
                
                tf = shape.text_frame
                tf.word_wrap = False
                p = tf.paragraphs[0]
                p.text = element.code
                p.font.name = "Consolas"
                p.font.size = Pt(14)
                p.font.color.rgb = theme.color_text
                
                if not ph: current_y += Inches(2.5)

            elif type(element).__name__ == 'Tree':
                start_x = ph.left if ph else content_x
                start_y = ph.top if ph else current_y
                width = ph.width if ph else layout_content_width
                
                tb = slide.shapes.add_textbox(start_x, start_y, width, Inches(2))
                tf = tb.text_frame
                
                def render_tree_node(node, level):
                    if level == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    
                    p.text = node.label
                    if level > 0:
                        p.level = level
                        
                    p.font.name = theme.font_name
                    p.font.size = Pt(18 - (level * 2)) if level < 4 else Pt(12)
                    
                    for child in node.children:
                        render_tree_node(child, level + 1)
                        
                render_tree_node(element.root, 0)
                if not ph: current_y += Inches(2.5)


            elif type(element).__name__ == 'Split':
                num_panels = len(element.panels)
                if num_panels == 0:
                    return current_y
                target_x = ph.left if ph else content_x
                target_y = ph.top if ph else current_y
                target_w = ph.width if ph else layout_content_width
                target_h = ph.height if ph else layout_content_height
                gap = Inches(0.2)
                if element.direction == 'horizontal':
                    panel_w = (target_w - (gap * (num_panels - 1))) / num_panels
                    max_bottom = target_y
                    for i, panel in enumerate(element.panels):
                        px = target_x + i * (panel_w + gap)
                        py = target_y
                        if panel.title:
                            tb = slide.shapes.add_textbox(px, py, panel_w, Inches(0.4))
                            p = tb.text_frame.paragraphs[0]
                            p.text = panel.title
                            p.font.name = theme.font_name
                            p.font.size = theme.size_body
                            p.font.bold = True
                            p.font.color.rgb = theme.color_primary
                            py += Inches(0.5)
                        end_y = py
                        for pe in panel.elements:
                            pe_ph = find_placeholder(getattr(pe, 'placeholder', None))
                            if pe_ph:
                                render_element(pe, pe_ph.left, pe_ph.top, pe_ph.width, pe_ph.height, pe_ph)
                            else:
                                end_y = render_element(pe, px, end_y, panel_w, target_h - (py - target_y), None)
                        if end_y > max_bottom:
                            max_bottom = end_y
                    if not ph: current_y = max_bottom + Inches(0.2)
                else:
                    panel_w = target_w
                    panel_h = (target_h - (gap * (num_panels - 1))) / num_panels
                    py = target_y
                    for i, panel in enumerate(element.panels):
                        px = target_x
                        panel_start_y = py
                        if panel.title:
                            tb = slide.shapes.add_textbox(px, py, panel_w, Inches(0.4))
                            p = tb.text_frame.paragraphs[0]
                            p.text = panel.title
                            p.font.name = theme.font_name
                            p.font.size = theme.size_body
                            p.font.bold = True
                            p.font.color.rgb = theme.color_primary
                            py += Inches(0.5)
                        for pe in panel.elements:
                            pe_ph = find_placeholder(getattr(pe, 'placeholder', None))
                            if pe_ph:
                                render_element(pe, pe_ph.left, pe_ph.top, pe_ph.width, pe_ph.height, pe_ph)
                            else:
                                py = render_element(pe, px, py, panel_w, panel_h - (py - panel_start_y), None)
                        py = panel_start_y + panel_h + gap
                    if not ph: current_y = py
            return current_y

        current_y = layout.content_y
        content_x = layout.content_x
        for element in slide_model.elements:
            ph = find_placeholder(getattr(element, "placeholder", None))
            current_y = render_element(element, content_x, current_y, layout.content_width, layout.content_height, ph)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
