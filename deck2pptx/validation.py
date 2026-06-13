import os
from pathlib import Path
from .models import Deck, Image, Gallery, Flow
from .errors import DeckValidationError

class ValidationError(Exception):
    pass

def validate_deck(deck: Deck, base_dir: str | Path):
    errors = []
    
    if deck.orientation not in ('landscape', 'portrait'):
        errors.append({
            "code": "invalid_orientation",
            "message": f"Invalid Deck orientation '{deck.orientation}'. Must be 'landscape' or 'portrait'.",
            "field": "orientation"
        })
        
    base_dir = Path(base_dir)
    for slide_idx, slide in enumerate(deck.slides):
        for elem_idx, element in enumerate(slide.elements):
            if isinstance(element, Image):
                img_path = base_dir / element.source
                if not img_path.is_file():
                    errors.append({
                        "code": "image_not_found",
                        "message": f"Slide {slide_idx+1}: Image file not found: {element.source}",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "Image",
                        "field": "source"
                    })
                    
            elif isinstance(element, Gallery):
                for img_idx, img in enumerate(element.images):
                    img_path = base_dir / img.source
                    if not img_path.is_file():
                        errors.append({
                            "code": "image_not_found",
                            "message": f"Slide {slide_idx+1}: Gallery image file not found: {img.source}",
                            "slide_index": slide_idx,
                            "slide_title": slide.title,
                            "element_index": elem_idx,
                            "element_type": "Gallery",
                            "field": f"images[{img_idx}].source"
                        })
                        
            elif isinstance(element, Flow):
                if element.direction not in ('horizontal', 'vertical'):
                    errors.append({
                        "code": "invalid_flow_direction",
                        "message": f"Slide {slide_idx+1}: Invalid Flow direction '{element.direction}'. Must be 'horizontal' or 'vertical'.",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "Flow",
                        "field": "direction"
                    })
                
                node_ids = {node.id for node in element.nodes}
                for edge in element.edges:
                    if edge.from_node not in node_ids:
                        errors.append({
                            "code": "invalid_flow_edge",
                            "message": f"Slide {slide_idx+1}: Flow edge references unknown from_node '{edge.from_node}'",
                            "slide_index": slide_idx,
                            "slide_title": slide.title,
                            "element_index": elem_idx,
                            "element_type": "Flow",
                            "field": "edges.from_node"
                        })
                    if edge.to_node not in node_ids:
                        errors.append({
                            "code": "invalid_flow_edge",
                            "message": f"Slide {slide_idx+1}: Flow edge references unknown to_node '{edge.to_node}'",
                            "slide_index": slide_idx,
                            "slide_title": slide.title,
                            "element_index": elem_idx,
                            "element_type": "Flow",
                            "field": "edges.to_node"
                        })
            elif type(element).__name__ == 'Comparison':
                if len(element.columns) < 2:
                    errors.append({
                        "code": "invalid_comparison_columns",
                        "message": f"Slide {slide_idx+1}: Comparison element must have at least 2 columns.",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "Comparison",
                        "field": "columns"
                    })
            elif type(element).__name__ == 'Timeline':
                if len(element.events) < 1:
                    errors.append({
                        "code": "invalid_timeline_events",
                        "message": f"Slide {slide_idx+1}: Timeline element must have at least 1 event.",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "Timeline",
                        "field": "events"
                    })
            elif type(element).__name__ == 'CodeBlock':
                if not element.code.strip():
                    errors.append({
                        "code": "invalid_code_block",
                        "message": f"Slide {slide_idx+1}: CodeBlock element must have code content.",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "CodeBlock",
                        "field": "code"
                    })
            elif type(element).__name__ == 'Tree':
                if not element.root:
                    errors.append({
                        "code": "invalid_tree",
                        "message": f"Slide {slide_idx+1}: Tree element must have a root node.",
                        "slide_index": slide_idx,
                        "slide_title": slide.title,
                        "element_index": elem_idx,
                        "element_type": "Tree",
                        "field": "root"
                    })

    if errors:
        # Preserve backwards compatibility by throwing ValidationError with the first error string if simple code expects it,
        # but attach the structured errors for the CLI.
        exc = ValidationError(errors[0]["message"])
        exc.errors = errors
        raise exc
