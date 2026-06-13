import pytest
from pathlib import Path
from deck2pptx.yaml_adapter import load_yaml
from deck2pptx.markdown_adapter import load_markdown
from deck2pptx.models import Comparison, Timeline, CodeBlock, Tree
from deck2pptx.adapters import load_deck
from deck2pptx.renderer import render_deck

def test_yaml_business_elements(tmp_path):
    content = """
title: Test
slides:
  - title: Slide 1
    elements:
      - comparison:
          columns:
            - label: Left
              items: ["a", "b"]
            - label: Right
              items: ["c", "d"]
      - timeline:
          events:
            - label: "2026"
              title: "Phase 1"
              description: "Desc"
      - code_block:
          code: "print('hello')"
          language: "python"
      - tree:
          root:
            label: "A"
            children:
              - label: "B"
"""
    f = tmp_path / "test.yaml"
    f.write_text(content, encoding='utf-8')
    deck = load_yaml(f)
    
    assert len(deck.slides) == 1
    elements = deck.slides[0].elements
    assert len(elements) == 4
    
    comp = elements[0]
    assert isinstance(comp, Comparison)
    assert len(comp.columns) == 2
    assert comp.columns[0].label == "Left"
    assert comp.columns[0].items == ["a", "b"]
    
    tl = elements[1]
    assert isinstance(tl, Timeline)
    assert len(tl.events) == 1
    assert tl.events[0].label == "2026"
    assert tl.events[0].title == "Phase 1"
    
    cb = elements[2]
    assert isinstance(cb, CodeBlock)
    assert cb.code == "print('hello')"
    assert cb.language == "python"
    
    tree = elements[3]
    assert isinstance(tree, Tree)
    assert tree.root.label == "A"
    assert len(tree.root.children) == 1
    assert tree.root.children[0].label == "B"

def test_markdown_business_elements(tmp_path):
    content = """
# Slide 1

```comparison
Left:
- a
- b
Right:
- c
- d
```

```timeline
2026: Phase 1 - Desc
```

```code python
print('hello')
```

```tree
A
  B
```
"""
    f = tmp_path / "test.md"
    f.write_text(content, encoding='utf-8')
    deck = load_markdown(f)
    
    assert len(deck.slides) == 1
    elements = deck.slides[0].elements
    assert len(elements) == 4
    
    comp = elements[0]
    assert isinstance(comp, Comparison)
    assert len(comp.columns) == 2
    assert comp.columns[0].label == "Left"
    assert comp.columns[0].items == ["a", "b"]
    
    tl = elements[1]
    assert isinstance(tl, Timeline)
    assert len(tl.events) == 1
    assert tl.events[0].label == "2026"
    assert tl.events[0].title == "Phase 1"
    assert tl.events[0].description == "Desc"
    
    cb = elements[2]
    assert isinstance(cb, CodeBlock)
    assert cb.code == "print('hello')"
    assert cb.language == "python"
    
    tree = elements[3]
    assert isinstance(tree, Tree)
    assert tree.root.label == "A"
    assert len(tree.root.children) == 1
    assert tree.root.children[0].label == "B"

def test_markdown_tree_uses_block_indent_unit(tmp_path):
    content = """---
indent: 4
---
# Slide 1

```tree
Root
  Child
    Grandchild
```
"""
    f = tmp_path / "tree.md"
    f.write_text(content, encoding='utf-8')
    deck = load_markdown(f)

    tree = deck.slides[0].elements[0]
    assert isinstance(tree, Tree)
    assert tree.root.label == "Root"
    assert tree.root.children[0].label == "Child"
    assert tree.root.children[0].children[0].label == "Grandchild"

def test_business_element_examples_build(tmp_path):
    for sample in [Path("examples/sample_advanced.md"), Path("examples/sample_advanced.deck.yaml")]:
        deck = load_deck(sample)
        output = tmp_path / f"{sample.stem}.pptx"
        render_deck(deck, str(output), base_dir=sample.parent)
        assert output.exists()
