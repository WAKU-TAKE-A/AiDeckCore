from dataclasses import dataclass, field
from typing import List, Optional, Union

@dataclass
class Text:
    content: str
    placeholder: Optional[str] = None

@dataclass
class ListItem:
    text: str
    level: int = 0

@dataclass
class BulletList:
    items: List[Union[str, ListItem]]
    placeholder: Optional[str] = None

@dataclass
class Image:
    source: str
    placeholder: Optional[str] = None

@dataclass
class Table:
    headers: List[str]
    rows: List[List[str]]
    placeholder: Optional[str] = None

@dataclass
class Gallery:
    images: List[Image]
    placeholder: Optional[str] = None

@dataclass
class FlowNode:
    id: str
    label: str

@dataclass
class FlowEdge:
    from_node: str
    to_node: str

@dataclass
class Flow:
    direction: str  # 'horizontal' or 'vertical'
    nodes: List[FlowNode]
    edges: List[FlowEdge]
    placeholder: Optional[str] = None

@dataclass
class ComparisonColumn:
    label: str
    items: List[str]

@dataclass
class Comparison:
    columns: List[ComparisonColumn]
    title: Optional[str] = None
    placeholder: Optional[str] = None

@dataclass
class TimelineEvent:
    label: str
    title: str
    description: Optional[str] = None

@dataclass
class Timeline:
    events: List[TimelineEvent]
    placeholder: Optional[str] = None

@dataclass
class CodeBlock:
    code: str
    language: Optional[str] = None
    caption: Optional[str] = None
    placeholder: Optional[str] = None

@dataclass
class TreeNode:
    label: str
    children: List['TreeNode'] = field(default_factory=list)

@dataclass
class Tree:
    root: TreeNode
    placeholder: Optional[str] = None

Element = Union[Text, BulletList, Image, Table, Gallery, Flow, Comparison, Timeline, CodeBlock, Tree]

@dataclass
class Slide:
    title: str
    subtitle: Optional[str] = None
    notes: Optional[str] = None
    layout_hint: Optional[str] = None
    elements: List[Element] = field(default_factory=list)

@dataclass
class Deck:
    title: Optional[str] = None
    orientation: str = "landscape"
    theme: str = "default"
    toc: bool = False
    toc_title: Optional[str] = None
    indent: Optional[int] = None
    font_size_l0: Optional[int] = None
    font_size_l1: Optional[int] = None
    font_size_l2: Optional[int] = None
    font_size_l3: Optional[int] = None
    font_size_l4: Optional[int] = None
    slides: List[Slide] = field(default_factory=list)
