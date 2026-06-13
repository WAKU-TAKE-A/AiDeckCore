import re
import yaml
from pathlib import Path
from .models import Deck, Slide, Text, BulletList, Image, Table, Gallery, Flow, FlowNode, FlowEdge

def load_markdown(file_path: str | Path) -> Deck:
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    deck = Deck()
    
    # 1. Parse Front Matter
    if content.startswith('---'):
        parts = content.split('---', 2)
        if len(parts) >= 3:
            try:
                fm = yaml.safe_load(parts[1]) or {}
                deck.title = fm.get('title')
                deck.orientation = fm.get('orientation', 'landscape')
                deck.theme = fm.get('theme', 'default')
                deck.toc = fm.get('toc', False)
                deck.toc_title = fm.get('toc_title')
                deck.indent = fm.get('indent')
                deck.font_size_l0 = fm.get('font_size_l0')
                deck.font_size_l1 = fm.get('font_size_l1')
                deck.font_size_l2 = fm.get('font_size_l2')
                deck.font_size_l3 = fm.get('font_size_l3')
                deck.font_size_l4 = fm.get('font_size_l4')
            except:
                pass
            content = parts[2].lstrip()
            
    # 2. Split into slides
    lines = content.splitlines()
    slides_data = []
    current_slide_lines = []
    
    # Look ahead for layout comments before headings
    pending_layout = None

    for line in lines:
        sline = line.strip()
        layout_match = re.match(r'^<!--\s*layout="(.*?)"\s*-->$', sline)
        if layout_match:
            pending_layout = layout_match.group(1)
            continue
            
        new_page_match = re.match(r'^<!--\s*new_page(?:="(.*?)")?\s*-->$', sline)
        if new_page_match:
            if current_slide_lines:
                if not all(not l.strip() for l in current_slide_lines):
                    slides_data.append((pending_layout, current_slide_lines))
            pending_layout = new_page_match.group(1)
            current_slide_lines = [""]
            continue
            
        if line.startswith('# ') or line.startswith('## '):
            if current_slide_lines:
                if all(not l.strip() for l in current_slide_lines):
                    current_slide_lines = [line]
                else:
                    slides_data.append((pending_layout, current_slide_lines))
                    current_slide_lines = [line]
                    pending_layout = None
            else:
                current_slide_lines = [line]
        else:
            current_slide_lines.append(line)
            
    if current_slide_lines:
        slides_data.append((pending_layout, current_slide_lines))
        
    for layout_hint, slide_lines in slides_data:
        if not slide_lines:
            continue
            
        header_line = slide_lines[0]
        title = header_line.lstrip('#').strip()
        
        slide = Slide(title=title, layout_hint=layout_hint)
        
        i = 1
        current_text = []
        current_bullets = []
        current_table = []
        current_placeholder = None
        
        def commit_text():
            if current_text:
                slide.elements.append(Text(content=' '.join(current_text), placeholder=current_placeholder))
                current_text.clear()
        
        def commit_bullets():
            if current_bullets:
                slide.elements.append(BulletList(items=list(current_bullets), placeholder=current_placeholder))
                current_bullets.clear()
                
        def commit_table():
            if current_table:
                # First row is header, second is divider, rest are rows
                headers = []
                rows = []
                if len(current_table) > 2 and '---' in current_table[1]:
                    headers = [c.strip() for c in current_table[0].strip('|').split('|')]
                    for r in current_table[2:]:
                        if r.strip('|'):
                            rows.append([c.strip() for c in r.strip('|').split('|')])
                else:
                    for r in current_table:
                        if r.strip('|'):
                            rows.append([c.strip() for c in r.strip('|').split('|')])
                slide.elements.append(Table(headers=headers if headers else None, rows=rows, placeholder=current_placeholder))
                current_table.clear()
                
        while i < len(slide_lines):
            raw_line = slide_lines[i]
            line = raw_line.strip()
            
            sub_match = re.match(r'^<!--\s*subtitle="(.*?)"\s*-->$', line)
            if sub_match:
                slide.subtitle = sub_match.group(1)
                i += 1
                continue
                
            place_match = re.match(r'^<!--\s*place(?:holder)?="(.*?)"\s*-->$', line)
            if place_match:
                current_placeholder = place_match.group(1)
                i += 1
                continue
            
            if not line:
                commit_text()
                commit_bullets()
                commit_table()
                i += 1
                continue
                
            # Table
            if line.startswith('|') and line.endswith('|'):
                commit_text()
                commit_bullets()
                current_table.append(line)
                i += 1
                continue
            else:
                commit_table()
                
            # Bullets
            if line.startswith('- ') or line.startswith('* '):
                commit_text()
                indent_spaces = len(raw_line) - len(raw_line.lstrip())
                indent_size = deck.indent if deck.indent else 2
                level = indent_spaces // indent_size
                from .models import ListItem
                current_bullets.append(ListItem(text=line[2:].strip(), level=level))
                i += 1
                continue
            else:
                commit_bullets()
                
            # Flow Block
            if line.startswith('```flow'):
                commit_text()
                direction = 'horizontal'
                if 'vertical' in line:
                    direction = 'vertical'
                
                i += 1
                nodes = []
                edges = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    fl = slide_lines[i].strip()
                    if '->' in fl:
                        fr, to = fl.split('->')
                        edges.append(FlowEdge(from_node=fr.strip(), to_node=to.strip()))
                    elif ':' in fl:
                        nid, nlabel = fl.split(':', 1)
                        nodes.append(FlowNode(id=nid.strip(), label=nlabel.strip()))
                    i += 1
                
                if nodes:
                    slide.elements.append(Flow(direction=direction, nodes=nodes, edges=edges, placeholder=current_placeholder))
                i += 1
                continue
                
            # Comparison Block
            if line.startswith('```comparison'):
                commit_text()
                i += 1
                from .models import Comparison, ComparisonColumn
                columns = []
                current_col_label = None
                current_col_items = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    cl = slide_lines[i].strip()
                    if cl.endswith(':'):
                        if current_col_label is not None:
                            columns.append(ComparisonColumn(label=current_col_label, items=current_col_items))
                        current_col_label = cl[:-1].strip()
                        current_col_items = []
                    elif cl.startswith('- ') or cl.startswith('* '):
                        current_col_items.append(cl[2:].strip())
                    i += 1
                if current_col_label is not None:
                    columns.append(ComparisonColumn(label=current_col_label, items=current_col_items))
                slide.elements.append(Comparison(columns=columns, placeholder=current_placeholder))
                i += 1
                continue
                
            # Timeline Block
            if line.startswith('```timeline'):
                commit_text()
                i += 1
                from .models import Timeline, TimelineEvent
                events = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    tl = slide_lines[i].strip()
                    if ':' in tl:
                        label, rest = tl.split(':', 1)
                        title = rest.strip()
                        desc = None
                        if ' - ' in title:
                            title, desc = title.split(' - ', 1)
                        events.append(TimelineEvent(label=label.strip(), title=title.strip(), description=desc.strip() if desc else None))
                    i += 1
                slide.elements.append(Timeline(events=events, placeholder=current_placeholder))
                i += 1
                continue
                
            # CodeBlock
            if line.startswith('```code'):
                commit_text()
                lang = line[7:].strip()
                i += 1
                from .models import CodeBlock
                code_lines = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    code_lines.append(slide_lines[i])
                    i += 1
                slide.elements.append(CodeBlock(code='\n'.join(code_lines), language=lang if lang else None, placeholder=current_placeholder))
                i += 1
                continue
                
            # Tree Block
            if line.startswith('```tree'):
                commit_text()
                i += 1
                from .models import Tree, TreeNode
                tree_lines = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    if slide_lines[i].strip():
                        tree_lines.append(slide_lines[i])
                    i += 1

                positive_indents = [
                    len(tree_line) - len(tree_line.lstrip())
                    for tree_line in tree_lines
                    if len(tree_line) - len(tree_line.lstrip()) > 0
                ]
                tree_indent_size = min(positive_indents) if positive_indents else (deck.indent if deck.indent else 2)

                nodes_by_level = {}
                root = None
                for raw_tree_line in tree_lines:
                    indent_spaces = len(raw_tree_line) - len(raw_tree_line.lstrip())
                    level = indent_spaces // tree_indent_size
                    node = TreeNode(label=raw_tree_line.strip())
                    nodes_by_level[level] = node
                    if level == 0 and not root:
                        root = node
                    elif level > 0 and (level - 1) in nodes_by_level:
                        nodes_by_level[level - 1].children.append(node)
                if root:
                    slide.elements.append(Tree(root=root, placeholder=current_placeholder))
                i += 1
                continue
                
            # Image
            img_match = re.match(r'^!\[.*?\]\((.*?)\)$', line)
            if img_match:
                commit_text()
                img_src = img_match.group(1)
                
                # Check if last element is Gallery with same placeholder, and append. If Image, convert to Gallery.
                if slide.elements and isinstance(slide.elements[-1], Gallery) and slide.elements[-1].placeholder == current_placeholder:
                    slide.elements[-1].images.append(Image(source=img_src, placeholder=current_placeholder))
                elif slide.elements and isinstance(slide.elements[-1], Image) and slide.elements[-1].placeholder == current_placeholder:
                    prev_img = slide.elements.pop()
                    slide.elements.append(Gallery(images=[prev_img, Image(source=img_src, placeholder=current_placeholder)], placeholder=current_placeholder))
                else:
                    slide.elements.append(Image(source=img_src, placeholder=current_placeholder))
                i += 1
                continue
                
            # Plain Text
            current_text.append(line)
            i += 1
            
        commit_text()
        commit_bullets()
        commit_table()
        
        deck.slides.append(slide)

    # Default layout hint for title slide if not explicitly set
    if deck.slides and deck.slides[0].title and not deck.slides[0].elements and not deck.slides[0].layout_hint:
        deck.slides[0].layout_hint = 'title'

    return deck
