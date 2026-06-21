import re
import yaml
from pathlib import Path
from .models import Deck, Slide, Text, BulletList, Image, Table, Gallery, Flow, FlowNode, FlowEdge

def load_markdown(file_path: str | Path) -> Deck:
    with open(file_path, 'r', encoding='utf-8-sig') as f:
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
                deck.content_align = fm.get('content_align')
            except:
                pass
            content = parts[2].lstrip()
            
    def parse_html_comment(line):
        import shlex
        m = re.match(r'^<!--(.*?)-->$', line.strip())
        if not m: return []
        content = m.group(1).strip()
        
        commands = []
        current = []
        in_quotes = False
        for char in content:
            if char == '"':
                in_quotes = not in_quotes
                current.append(char)
            elif char == ';' and not in_quotes:
                commands.append(''.join(current).strip())
                current = []
            else:
                current.append(char)
        if current:
            commands.append(''.join(current).strip())
            
        parsed_cmds = []
        for cmd in commands:
            if not cmd: continue
            try:
                tokens = shlex.split(cmd)
                if not tokens: continue
                
                if '=' in tokens[0] and tokens[0] != '=':
                    parts = tokens[0].split('=', 1)
                    cmd_name = parts[0].lower()
                    args = [parts[1]] + tokens[1:]
                else:
                    cmd_name = tokens[0].lower()
                    args = tokens[1:]
                    
                if cmd_name in ('l', 'layout'): cmd_name = 'layout'
                elif cmd_name in ('sub', 'subtitle'): cmd_name = 'subtitle'
                elif cmd_name in ('ph', 'place', 'placeholder'): cmd_name = 'placeholder'
                elif cmd_name in ('new', 'new_page', 'newpage'): cmd_name = 'newpage'
                elif cmd_name in ('align', 'content_align', 'valign'): cmd_name = 'content_align'
                elif cmd_name in ('gal', 'gallery'): cmd_name = 'gallery'
                
                parsed_args = []
                for arg in args:
                    if '=' in arg:
                        k, v = arg.split('=', 1)
                        parsed_args.append((k, v))
                    else:
                        parsed_args.append(arg)
                parsed_cmds.append((cmd_name, parsed_args))
            except:
                pass
        return parsed_cmds

    # 2. Split into slides
    lines = content.splitlines()
    slides_data = []
    current_slide_lines = []
    pending_layout = None

    def append_current_slide():
        if current_slide_lines and not all(not l.strip() for l in current_slide_lines):
            slides_data.append((pending_layout, list(current_slide_lines)))

    def has_heading(lines):
        return any(re.match(r'^(#{1,3})\s+', l) for l in lines)

    def layout_arg(args):
        if not args:
            return None
        if type(args[0]) == tuple:
            return args[0][1]
        return args[0]

    for line in lines:
        sline = line.strip()
        cmds = parse_html_comment(sline)
        has_newpage = any(cmd == 'newpage' for cmd, _ in cmds)
        has_layout = any(cmd == 'layout' for cmd, _ in cmds)
        
        if has_newpage:
            append_current_slide()
            new_layout = None
            for cmd, args in cmds:
                if cmd in ('newpage', 'layout') and args:
                    new_layout = layout_arg(args)
            pending_layout = new_layout
            current_slide_lines = [""]
            other_cmds = [(c, a) for c, a in cmds if c not in ('newpage', 'layout')]
            if other_cmds:
                current_slide_lines.append(line)
            continue
            
        if has_layout and not has_newpage and len(cmds) == 1:
            for cmd, args in cmds:
                if cmd == 'layout' and args:
                    if has_heading(current_slide_lines) and current_slide_lines[-1].strip():
                        current_slide_lines.append(line)
                    else:
                        append_current_slide()
                        pending_layout = layout_arg(args)
                        current_slide_lines = []
            continue
            
        if re.match(r'^(#{1,3})\s+', line):
            if current_slide_lines:
                has_content = False
                for l in current_slide_lines:
                    if l.strip() and not l.strip().startswith('<!--'):
                        has_content = True
                        break
                if not has_content:
                    current_slide_lines.append(line)
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
            
        title_line_idx = -1
        for idx, ln in enumerate(slide_lines):
            if re.match(r'^(#{1,3})\s+', ln):
                title_line_idx = idx
                break
        
        if title_line_idx >= 0:
            title = slide_lines[title_line_idx].lstrip('#').strip()
        else:
            title = ""
            
        i = 0
            
        slide = Slide(title=title, layout_hint=layout_hint)
        
        current_text = []
        current_bullets = []
        current_table = []
        current_placeholder = None
        
        active_split = None
        active_panel = None
        active_gallery = None
        
        def get_target_list():
            if active_panel:
                return active_panel.elements
            return slide.elements
        
        def commit_text():
            nonlocal active_gallery
            if current_text:
                get_target_list().append(Text(content=join_text_lines(current_text), placeholder=current_placeholder))
                current_text.clear()
                active_gallery = None

        def join_text_lines(lines):
            content = ""
            for text in lines:
                if not content:
                    content = text
                elif content.endswith('\n') or text.startswith('\n'):
                    content += text
                else:
                    content += ' ' + text
            return content

        def normalize_text_line(raw_text):
            text = raw_text.strip()
            hard_break = len(raw_text) - len(raw_text.rstrip(' ')) >= 2
            text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
            text = text.replace('\\n', '\n')
            if hard_break and not text.endswith('\n'):
                text += '\n'
            return text
        
        def commit_bullets():
            nonlocal active_gallery
            if current_bullets:
                get_target_list().append(BulletList(items=list(current_bullets), placeholder=current_placeholder))
                current_bullets.clear()
                active_gallery = None
                
        def commit_table():
            nonlocal active_gallery
            if current_table:
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
                get_target_list().append(Table(headers=headers, rows=rows, placeholder=current_placeholder))
                current_table.clear()
                active_gallery = None
                
        while i < len(slide_lines):
            if i == title_line_idx:
                i += 1
                continue

            raw_line = slide_lines[i]
            line = raw_line.strip()
            
            cmds = parse_html_comment(line)
            if cmds:
                handled_line = False
                for cmd, args in cmds:
                    if cmd == 'subtitle':
                        slide.subtitle = args[0] if args else None
                        handled_line = True
                    elif cmd == 'layout':
                        slide.layout_hint = args[0] if args else None
                        handled_line = True
                    elif cmd == 'content_align':
                        slide.content_align = args[0] if args else None
                        handled_line = True
                    elif cmd == 'placeholder':
                        current_placeholder = args[0] if args else None
                        if len(args) > 1:
                            hidden_val = normalize_text_line(args[1])
                            commit_text()
                            commit_bullets()
                            commit_table()
                            get_target_list().append(Text(content=hidden_val, placeholder=current_placeholder))
                            current_placeholder = None
                        handled_line = True
                    elif cmd in ('v', 'value'):
                        if current_placeholder and args:
                            hidden_val = normalize_text_line(args[0])
                            commit_text()
                            commit_bullets()
                            commit_table()
                            get_target_list().append(Text(content=hidden_val, placeholder=current_placeholder))
                            current_placeholder = None
                        handled_line = True
                    elif cmd == 'gallery':
                        commit_text()
                        commit_bullets()
                        commit_table()
                        cols = None
                        if args:
                            try:
                                cols = int(args[0])
                            except:
                                pass
                        active_gallery = Gallery(images=[], columns=cols, placeholder=current_placeholder)
                        get_target_list().append(active_gallery)
                        handled_line = True
                    elif cmd == 'split':
                        commit_text()
                        commit_bullets()
                        commit_table()
                        from .models import Split, Panel
                        direction = 'horizontal'
                        if args:
                            arg = args[0]
                            if isinstance(arg, tuple) and arg[0] == 'direction':
                                direction = 'horizontal' if arg[1] in ('h', 'horizontal') else 'vertical'
                            elif isinstance(arg, str):
                                direction = 'horizontal' if arg in ('h', 'horizontal') else 'vertical'
                        active_split = Split(direction=direction, panels=[], placeholder=current_placeholder)
                        slide.elements.append(active_split)
                        active_panel = Panel()
                        active_split.panels.append(active_panel)
                        handled_line = True
                    elif cmd == 'panel':
                        commit_text()
                        commit_bullets()
                        commit_table()
                        from .models import Panel
                        title_val = None
                        if args:
                            arg = args[0]
                            if isinstance(arg, tuple) and arg[0] == 'title':
                                title_val = arg[1]
                            elif isinstance(arg, str):
                                title_val = arg
                        if active_split:
                            if len(active_split.panels) == 1 and not active_split.panels[0].elements and active_split.panels[0].title is None:
                                active_split.panels.clear()
                            active_panel = Panel(title=title_val)
                            active_split.panels.append(active_panel)
                        handled_line = True
                    elif cmd == '/split':
                        commit_text()
                        commit_bullets()
                        commit_table()
                        active_split = None
                        active_panel = None
                        handled_line = True
                if handled_line:
                    i += 1
                    continue
            
            if not line:
                commit_text()
                commit_bullets()
                commit_table()
                active_gallery = None
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
                active_gallery = None
                direction = 'horizontal'
                if 'vertical' in line:
                    direction = 'vertical'
                
                start_i = i
                i += 1
                nodes = []
                edges = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    fl = slide_lines[i].strip()
                    edge_match = re.search(r'(.+?)\s*-+>\s*(.+)', fl)
                    if edge_match:
                        edges.append(FlowEdge(from_node=edge_match.group(1).strip(), to_node=edge_match.group(2).strip()))
                    else:
                        node_match = re.match(r'^([^\[\(\{]+?)\s*[\[\(\{](.+?)[\]\)\}]\s*$', fl)
                        if node_match:
                            nid = node_match.group(1).strip()
                            nlabel = node_match.group(2).strip()
                            if nlabel.startswith('"') and nlabel.endswith('"'):
                                nlabel = nlabel[1:-1]
                            nodes.append(FlowNode(id=nid, label=nlabel))
                    i += 1
                
                if nodes:
                    get_target_list().append(Flow(direction=direction, nodes=nodes, edges=edges, placeholder=current_placeholder))
                else:
                    from .models import CodeBlock
                    raw_code = '\n'.join(slide_lines[start_i:i+1])
                    get_target_list().append(CodeBlock(code=raw_code, placeholder=current_placeholder))
                i += 1
                continue
                
            # Mermaid Block
            if line.startswith('```mermaid'):
                commit_text()
                active_gallery = None
                start_i = i
                i += 1
                code_lines = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    code_lines.append(slide_lines[i])
                    i += 1
                
                from .models import Mermaid
                get_target_list().append(Mermaid(code='\n'.join(code_lines), placeholder=current_placeholder))
                i += 1
                continue
                
            # Comparison Block
            if line.startswith('```comparison'):
                commit_text()
                active_gallery = None
                start_i = i
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
                if columns:
                    get_target_list().append(Comparison(columns=columns, placeholder=current_placeholder))
                else:
                    from .models import CodeBlock
                    raw_code = '\n'.join(slide_lines[start_i:i+1])
                    get_target_list().append(CodeBlock(code=raw_code, placeholder=current_placeholder))
                i += 1
                continue
                
            # Timeline Block
            if line.startswith('```timeline'):
                commit_text()
                active_gallery = None
                start_i = i
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
                if events:
                    get_target_list().append(Timeline(events=events, placeholder=current_placeholder))
                else:
                    from .models import CodeBlock
                    raw_code = '\n'.join(slide_lines[start_i:i+1])
                    get_target_list().append(CodeBlock(code=raw_code, placeholder=current_placeholder))
                i += 1
                continue
                
            # CodeBlock
            if line.startswith('```code'):
                commit_text()
                active_gallery = None
                lang = line[7:].strip()
                i += 1
                from .models import CodeBlock
                code_lines = []
                while i < len(slide_lines) and not slide_lines[i].strip().startswith('```'):
                    code_lines.append(slide_lines[i])
                    i += 1
                get_target_list().append(CodeBlock(code='\n'.join(code_lines), language=lang if lang else None, placeholder=current_placeholder))
                i += 1
                continue
                
            # Tree Block
            if line.startswith('```tree'):
                commit_text()
                active_gallery = None
                start_i = i
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
                    get_target_list().append(Tree(root=root, placeholder=current_placeholder))
                else:
                    from .models import CodeBlock
                    raw_code = '\n'.join(slide_lines[start_i:i+1])
                    get_target_list().append(CodeBlock(code=raw_code, placeholder=current_placeholder))
                i += 1
                continue
                
            # Image
            img_match = re.match(r'^!\[(.*?)\]\((.*?)\)$', line)
            if img_match:
                commit_text()
                img_alt = img_match.group(1).strip()
                img_src = img_match.group(2).strip()
                caption = img_alt if img_alt else None
                
                if active_gallery and active_gallery.placeholder == current_placeholder:
                    active_gallery.images.append(Image(source=img_src, caption=caption, placeholder=current_placeholder))
                else:
                    get_target_list().append(Image(source=img_src, caption=caption, placeholder=current_placeholder))
                i += 1
                continue
                
            # Plain Text
            current_text.append(normalize_text_line(raw_line))
            i += 1
            
        commit_text()
        commit_bullets()
        commit_table()
        
        if active_split:
            raise ValueError(f"Slide '{title}': Unclosed Split detected. A <!-- split --> was started but never closed with <!-- /split -->.")
        
        deck.slides.append(slide)

    # Default layout hint for title slide if not explicitly set
    if deck.slides and deck.slides[0].title and not deck.slides[0].elements and not deck.slides[0].layout_hint:
        deck.slides[0].layout_hint = 'title'

    return deck
