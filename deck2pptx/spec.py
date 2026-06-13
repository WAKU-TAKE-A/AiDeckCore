def get_spec() -> dict:
    return {
        "description": "Deck to PPTX MVP Format Spec",
        "concept": "The deck2pptx tool accepts YAML or Markdown input files. Both are only adapters; the canonical model is the Deck object.",
        "deck_metadata": {
            "title": "The presentation title (string).",
            "orientation": "'landscape' (default) or 'portrait'.",
            "theme": "Theme name, e.g. 'default'.",
            "toc": "Generate table of contents slide automatically (boolean).",
            "toc_title": "Custom TOC slide title (string).",
            "indent": "List indentation mapping for Markdown parsing. Specifies how many spaces equal one list level (integer).",
            "font_size_l0": "Font size for level 0 text (integer). Overrides theme defaults.",
            "font_size_l1": "Font size for level 1 text (integer). Overrides theme defaults.",
            "font_size_l2": "Font size for level 2 text (integer). Overrides theme defaults.",
            "font_size_l3": "Font size for level 3 text (integer). Overrides theme defaults.",
            "font_size_l4": "Font size for level 4 text (integer). Overrides theme defaults."
        },
        "slide_metadata": {
            "title": "Slide title (string).",
            "subtitle": "Slide subtitle (string).",
            "notes": "Presenter notes (string).",
            "layout_hint": "Target PPTX layout name, or built-in hints like 'title', 'content'."
        },
        "elements": {
            "description": "All elements support an optional `placeholder` field (string) to target PPTX placeholders.",
            "text": "A simple text block (Markdown normal paragraph).",
            "bullet_list": "A list of strings (Markdown `-` or `*` list).",
            "image": "A relative path to an image file (Markdown `![alt](path)`).",
            "table": "An object with `headers` and `rows` (Markdown tables with `|`).",
            "gallery": "An object with `images` (Markdown consecutive images).",
            "flow": {
                "description": "A flowchart.",
                "fields": {
                    "direction": "'horizontal' or 'vertical'.",
                    "nodes": "List of objects with `id` and `label`.",
                    "edges": "List of objects with `from` and `to` matching node IDs."
                }
            },
            "comparison": {
                "description": "A comparison matrix.",
                "fields": {
                    "columns": "List of objects with `label` and `items` (list of strings).",
                    "title": "Optional title string."
                }
            },
            "timeline": {
                "description": "A timeline of events.",
                "fields": {
                    "events": "List of objects with `label`, `title`, and optional `description`."
                }
            },
            "code_block": {
                "description": "A source code block.",
                "fields": {
                    "code": "The raw code string.",
                    "language": "Optional programming language name.",
                    "caption": "Optional caption string."
                }
            },
            "tree": {
                "description": "A hierarchical tree.",
                "fields": {
                    "root": "A node object with `label` and optional `children` (list of nodes)."
                }
            }
        },
        "markdown_notes": "In Markdown, use `<!-- layout=\"Name\" -->`, `<!-- new_page=\"Name\" -->`, `<!-- subtitle=\"Text\" -->`, and `<!-- place=\"Name\" -->` for structural controls.\nUse code blocks for business elements: ```comparison, ```timeline, ```code <lang>, ```tree.",
        "validation_rules": [
            "Deck orientation must be 'landscape' or 'portrait'.",
            "Image paths and gallery image paths must exist relative to the input file.",
            "Flow direction must be supported.",
            "Flow edges must reference known node IDs.",
            "Comparison must have at least 2 columns.",
            "Timeline must have at least 1 event.",
            "CodeBlock must have code content.",
            "Tree must have a root node."
        ],
        "non_goals": [
            "AsciiDoc adapter, Natural Language adapter.",
            "Full template/theme systems or broad visual redesigns.",
            "Manual coordinate specification (x, y, width, height) in the Deck model."
        ]
    }

def explain_spec_text() -> str:
    spec = get_spec()
    text = f"{spec['description']}\n\n{spec['concept']}\n\nSupported Deck-level metadata:\n"
    for k, v in spec["deck_metadata"].items():
        text += f"- `{k}`: {v}\n"
        
    text += "\nSupported Slide-level metadata:\n"
    for k, v in spec["slide_metadata"].items():
        text += f"- `{k}`: {v}\n"
        
    text += "\nSupported elements inside `elements`:\n"
    for k, v in spec["elements"].items():
        if isinstance(v, dict):
            text += f"- `{k}`: {v['description']}\n"
            for fk, fv in v["fields"].items():
                text += f"    - `{fk}`: {fv}\n"
        else:
            text += f"- `{k}`: {v}\n"
            
    text += f"\n({spec['markdown_notes']})\n"
    text += "\nValidation Rules:\n"
    for rule in spec["validation_rules"]:
        text += f"- {rule}\n"
        
    text += "\nNon-Goals:\n"
    for goal in spec["non_goals"]:
        text += f"- {goal}\n"
        
    return text
