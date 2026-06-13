import yaml
from pathlib import Path
from .models import Deck, Slide, Text, BulletList, Image, Table, Gallery, Flow, FlowNode, FlowEdge

def load_yaml(file_path: str | Path) -> Deck:
    with open(file_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    
    deck = Deck(
        title=data.get('title'),
        orientation=data.get('orientation', 'landscape'),
        theme=data.get('theme', 'default'),
        toc=data.get('toc', False),
        toc_title=data.get('toc_title'),
        indent=data.get('indent'),
        font_size_l0=data.get('font_size_l0'),
        font_size_l1=data.get('font_size_l1'),
        font_size_l2=data.get('font_size_l2'),
        font_size_l3=data.get('font_size_l3'),
        font_size_l4=data.get('font_size_l4')
    )
    for slide_data in data.get('slides', []):
        slide = Slide(
            title=slide_data.get('title', ''),
            subtitle=slide_data.get('subtitle'),
            notes=slide_data.get('notes'),
            layout_hint=slide_data.get('layout_hint')
        )
        
        for elem_data in slide_data.get('elements', []):
            placeholder = elem_data.get('placeholder')
            if 'text' in elem_data:
                slide.elements.append(Text(content=elem_data['text'], placeholder=placeholder))
            elif 'bullet_list' in elem_data:
                from .models import ListItem
                parsed_items = []
                for item in elem_data['bullet_list']:
                    if isinstance(item, dict):
                        parsed_items.append(ListItem(text=item.get('text', ''), level=item.get('level', 0)))
                    else:
                        parsed_items.append(str(item))
                slide.elements.append(BulletList(items=parsed_items, placeholder=placeholder))
            elif 'image' in elem_data:
                slide.elements.append(Image(source=elem_data['image'], placeholder=placeholder))
            elif 'table' in elem_data:
                table_data = elem_data['table']
                slide.elements.append(Table(
                    headers=table_data.get('headers', []),
                    rows=table_data.get('rows', []),
                    placeholder=placeholder
                ))
            elif 'gallery' in elem_data:
                gallery_data = elem_data['gallery']
                images = [Image(source=img) for img in gallery_data.get('images', [])]
                slide.elements.append(Gallery(images=images, placeholder=placeholder))
            elif 'flow' in elem_data:
                flow_data = elem_data['flow']
                nodes = [FlowNode(id=n['id'], label=n['label']) for n in flow_data.get('nodes', [])]
                edges = [FlowEdge(from_node=e['from'], to_node=e['to']) for e in flow_data.get('edges', [])]
                slide.elements.append(Flow(
                    direction=flow_data.get('direction', 'horizontal'),
                    nodes=nodes,
                    edges=edges,
                    placeholder=placeholder
                ))
            elif 'comparison' in elem_data:
                from .models import Comparison, ComparisonColumn
                comp_data = elem_data['comparison']
                columns = [ComparisonColumn(label=c.get('label', ''), items=c.get('items', [])) for c in comp_data.get('columns', [])]
                slide.elements.append(Comparison(
                    columns=columns,
                    title=comp_data.get('title'),
                    placeholder=placeholder
                ))
            elif 'timeline' in elem_data:
                from .models import Timeline, TimelineEvent
                tl_data = elem_data['timeline']
                events = [TimelineEvent(label=e.get('label', ''), title=e.get('title', ''), description=e.get('description')) for e in tl_data.get('events', [])]
                slide.elements.append(Timeline(
                    events=events,
                    placeholder=placeholder
                ))
            elif 'code_block' in elem_data:
                from .models import CodeBlock
                cb_data = elem_data['code_block']
                slide.elements.append(CodeBlock(
                    code=cb_data.get('code', ''),
                    language=cb_data.get('language'),
                    caption=cb_data.get('caption'),
                    placeholder=placeholder
                ))
            elif 'tree' in elem_data:
                from .models import Tree, TreeNode
                tree_data = elem_data['tree']
                def parse_tree_node(node_data):
                    return TreeNode(
                        label=node_data.get('label', ''),
                        children=[parse_tree_node(c) for c in node_data.get('children', [])]
                    )
                if 'root' in tree_data:
                    root = parse_tree_node(tree_data['root'])
                    slide.elements.append(Tree(root=root, placeholder=placeholder))
        deck.slides.append(slide)
    
    return deck
