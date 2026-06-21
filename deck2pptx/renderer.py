from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pathlib import Path
from .models import Deck, Slide, BulletList
from .layout import Layout, get_slide_layout_type
from .theme import Theme

from .render_context import SlideContext
from .height_estimator import get_adjusted_height
from .element_renderers import render_element

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



def render_deck(deck: Deck, output_path: str, base_dir: Path = Path('.'), template_path: str = None, calib_first_slide: bool = False):
    calibrated_metrics = {}
    # Initialize Presentation
    if template_path:
        prs = Presentation(template_path)
        
        # Calibration Extraction
        if calib_first_slide and len(prs.slides) > 0:
            calib_slide = prs.slides[0]
            for shape in calib_slide.shapes:
                if shape.has_text_frame:
                    text = shape.text
                    lines = text.count('\n') + 1
                    if lines >= 2:
                        font_size_pt = None
                        for p in shape.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size:
                                    font_size_pt = run.font.size.pt
                                    break
                            if font_size_pt: break
                        
                        if font_size_pt:
                            # Calculate actual line height based on font size (approx 1.2x)
                            # 1 pt = 12700 EMU. 12700 * 1.2 = 15240
                            height_per_line = int(font_size_pt * 15240)
                            first_para_text = shape.text_frame.paragraphs[0].text
                            cpi = len(first_para_text) / (shape.width / 914400.0) if shape.width else 60.0 / 6.0
                            calibrated_metrics[font_size_pt] = {
                                'height': height_per_line,
                                'cpi': cpi
                            }
                            
        level_fonts = {}
        if calibrated_metrics:
            sorted_fonts = sorted(calibrated_metrics.keys(), reverse=True)
            for i, fs in enumerate(sorted_fonts):
                level_fonts[i] = fs

        _remove_existing_slides(prs)
    else:
        prs = Presentation()
        level_fonts = {}
    
    base_dir = Path(base_dir)
    theme = Theme(deck.theme)
    
    # Set slide dimensions (only when no template; templates keep their own size)
    if not template_path:
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
        
        ctx = SlideContext(
            slide=slide,
            theme=theme,
            deck=deck,
            base_dir=base_dir,
            layout=layout,
            find_placeholder=find_placeholder,
            calibrated_metrics=calibrated_metrics,
            level_fonts=level_fonts
        )

        align_val = getattr(slide_model, 'content_align', None) or getattr(deck, 'content_align', None)
        y_offset = 0
        if align_val:
            val = align_val.lower()
            if val in ('top',): y_offset = -Inches(0.6)
            elif val in ('semi-top', 'high'): y_offset = -Inches(0.3)
            elif val in ('semi-bottom', 'low'): y_offset = Inches(0.3)
            elif val in ('bottom',): y_offset = Inches(0.6)
            
        current_y = layout.content_y + y_offset
        content_x = layout.content_x
        for idx_el, element in enumerate(slide_model.elements):
            ph = find_placeholder(getattr(element, "placeholder", None))
            adj_h = layout.content_height if ph else get_adjusted_height(
                slide_model.elements, idx_el,
                layout.content_y + layout.content_height,
                current_y, layout.content_width
            )
            current_y = render_element(element, ctx, content_x, current_y, layout.content_width, adj_h)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
