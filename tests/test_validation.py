import pytest
import os
from pathlib import Path
from deck2pptx.models import Deck, Slide, Text, Image, Flow, Gallery, FlowNode, FlowEdge
from deck2pptx.adapters import load_deck
from deck2pptx.validation import validate_deck, ValidationError

def test_deck_model_has_no_coordinates():
    t = Text(content="hello")
    assert not hasattr(t, 'x')
    assert not hasattr(t, 'y')
    assert not hasattr(t, 'width')
    assert not hasattr(t, 'height')

def test_valid_deck_yaml():
    sample_path = Path('examples/sample.deck.yaml')
    if sample_path.exists():
        deck = load_deck(sample_path)
        # Should not raise any ValidationError
        validate_deck(deck, sample_path.parent)

def test_valid_deck_markdown():
    sample_path = Path('examples/sample.md')
    if sample_path.exists():
        deck = load_deck(sample_path)
        # Should not raise any ValidationError
        validate_deck(deck, sample_path.parent)

def test_nested_validation():
    nested_path = Path('tests/fixtures/nested/fixture.deck.yaml')
    if nested_path.exists():
        deck = load_deck(nested_path)
        validate_deck(deck, nested_path.parent)

def test_portrait_validation():
    sample_path = Path('tests/fixtures/portrait.deck.yaml')
    if sample_path.exists():
        deck = load_deck(sample_path)
        assert deck.orientation == 'portrait'
        validate_deck(deck, sample_path.parent)

def test_invalid_orientation():
    deck = Deck(orientation='diagonal', slides=[Slide(title="Test")])
    with pytest.raises(ValidationError, match="Invalid Deck orientation 'diagonal'"):
        validate_deck(deck, Path('.'))

def test_missing_image_rejected():
    deck = Deck(slides=[
        Slide(title="Test", elements=[Image(source="nonexistent.png")])
    ])
    with pytest.raises(ValidationError, match="Image file not found"):
        validate_deck(deck, Path('.'))

def test_missing_gallery_image_rejected():
    deck = Deck(slides=[
        Slide(title="Test", elements=[Gallery(images=[Image(source="nonexistent.png")])])
    ])
    with pytest.raises(ValidationError, match="Gallery image file not found"):
        validate_deck(deck, Path('.'))

def test_invalid_flow_direction_rejected():
    deck = Deck(slides=[
        Slide(title="Test", elements=[Flow(direction="diagonal", nodes=[], edges=[])])
    ])
    with pytest.raises(ValidationError, match="Invalid Flow direction 'diagonal'"):
        validate_deck(deck, Path('.'))

def test_invalid_flow_edge_from_node_rejected():
    deck = Deck(slides=[
        Slide(title="Test", elements=[Flow(
            direction="horizontal", 
            nodes=[FlowNode(id="1", label="1")], 
            edges=[FlowEdge(from_node="2", to_node="1")]
        )])
    ])
    with pytest.raises(ValidationError, match="references unknown from_node '2'"):
        validate_deck(deck, Path('.'))

def test_invalid_flow_edge_to_node_rejected():
    deck = Deck(slides=[
        Slide(title="Test", elements=[Flow(
            direction="horizontal", 
            nodes=[FlowNode(id="1", label="1")], 
            edges=[FlowEdge(from_node="1", to_node="2")]
        )])
    ])
    with pytest.raises(ValidationError, match="references unknown to_node '2'"):
        validate_deck(deck, Path('.'))

def test_multi_error_validation():
    sample_path = Path('tests/fixtures/multi_error.deck.yaml')
    if sample_path.exists():
        deck = load_deck(sample_path)
        with pytest.raises(ValidationError) as excinfo:
            validate_deck(deck, sample_path.parent)
        
        errors = excinfo.value.errors
        assert len(errors) == 4
        codes = [e['code'] for e in errors]
        assert 'invalid_orientation' in codes
        assert 'image_not_found' in codes
        assert 'invalid_flow_direction' in codes
        assert 'invalid_flow_edge' in codes

def test_invalid_comparison_rejected():
    from deck2pptx.models import Comparison, ComparisonColumn
    deck = Deck(slides=[
        Slide(title="Test", elements=[Comparison(columns=[ComparisonColumn(label="Only One", items=[])])])
    ])
    with pytest.raises(ValidationError, match="Comparison element must have at least 2 columns"):
        validate_deck(deck, Path('.'))

def test_invalid_timeline_rejected():
    from deck2pptx.models import Timeline
    deck = Deck(slides=[
        Slide(title="Test", elements=[Timeline(events=[])])
    ])
    with pytest.raises(ValidationError, match="Timeline element must have at least 1 event"):
        validate_deck(deck, Path('.'))

def test_invalid_code_block_rejected():
    from deck2pptx.models import CodeBlock
    deck = Deck(slides=[
        Slide(title="Test", elements=[CodeBlock(code="   ")])
    ])
    with pytest.raises(ValidationError, match="CodeBlock element must have code content"):
        validate_deck(deck, Path('.'))

def test_invalid_tree_rejected():
    from deck2pptx.models import Tree
    deck = Deck(slides=[
        Slide(title="Test", elements=[Tree(root=None)])
    ])
    with pytest.raises(ValidationError, match="Tree element must have a root node"):
        validate_deck(deck, Path('.'))
