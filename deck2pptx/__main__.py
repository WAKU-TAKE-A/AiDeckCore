import argparse
import sys
import json
from pathlib import Path
from .adapters import load_deck
from .validation import validate_deck, ValidationError
from .renderer import render_deck
from .spec import get_spec, explain_spec_text
from .inspect import inspect_deck

def explain_spec(args):
    if getattr(args, 'format', None) == 'json':
        print(json.dumps(get_spec(), indent=2, ensure_ascii=False))
    else:
        print(explain_spec_text())

def inspect_cmd(args):
    try:
        deck = load_deck(args.input_file, format=args.format)
        data = inspect_deck(deck)
        if getattr(args, 'output_format', None) == 'json':
            print(json.dumps(data, indent=2, ensure_ascii=False))
        else:
            print(json.dumps(data, indent=2, ensure_ascii=False)) # Default to json for inspect for now
    except Exception as e:
        if getattr(args, 'output_format', None) == 'json':
            print(json.dumps({"ok": False, "errors": [{"message": str(e)}]}), file=sys.stderr)
        else:
            print(f"Inspect failed: {e}", file=sys.stderr)
        sys.exit(1)

def validate_cmd(args):
    is_json = getattr(args, 'output_format', None) == 'json'
    try:
        deck = load_deck(args.input_file, format=args.format)
        validate_deck(deck, Path(args.input_file).parent)
        if is_json:
            print(json.dumps({"ok": True, "errors": []}, indent=2, ensure_ascii=False))
        else:
            print(f"Validation successful. Found {len(deck.slides)} slides.")
    except ValidationError as e:
        if is_json:
            print(json.dumps({"ok": False, "errors": getattr(e, 'errors', [{"message": str(e)}])}, indent=2, ensure_ascii=False))
        else:
            errors = getattr(e, 'errors', [{"message": str(e)}])
            print(f"Validation failed with {len(errors)} errors:", file=sys.stderr)
            for err in errors:
                print(f"- {err['message']}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        if is_json:
            print(json.dumps({"ok": False, "errors": [{"message": str(e)}]}), file=sys.stderr)
        else:
            print(f"Validation failed: {e}", file=sys.stderr)
        sys.exit(1)

def build_cmd(args):
    try:
        deck = load_deck(args.input_file, format=args.format)
        validate_deck(deck, Path(args.input_file).parent)
        render_deck(deck, args.output_file, base_dir=Path(args.input_file).parent, template_path=getattr(args, 'template', None))
        print(f"Built PPTX successfully to {args.output_file}")
    except Exception as e:
        print(f"Build failed: {e}", file=sys.stderr)
        sys.exit(1)

def inspect_template_cmd(args):
    from .inspect_template import inspect_template
    inspect_template(args.template_file, args.format)

def main():
    parser = argparse.ArgumentParser(description="Generate PowerPoint from Semantic Deck Format")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")
    
    # Explain Spec
    p_explain = subparsers.add_parser('explain-spec', help="Output the AI-facing model schema")
    p_explain.add_argument('--format', help="Output format (e.g. json)", default=None)
    p_explain.set_defaults(func=explain_spec)
    
    # Inspect
    p_inspect = subparsers.add_parser('inspect', help="Inspect input as normalized Deck")
    p_inspect.add_argument('input_file', help="Path to input YAML or MD file")
    p_inspect.add_argument('--format', dest="output_format", help="Output format (e.g. json)", default=None)
    p_inspect.add_argument('--input-format', dest="format", help="Force input format (yaml or markdown)", default=None)
    p_inspect.set_defaults(func=inspect_cmd)

    # Inspect Template
    p_inspect_tmpl = subparsers.add_parser('inspect-template', help="Inspect PPTX template layouts and placeholders")
    p_inspect_tmpl.add_argument('template_file', help="Path to PPTX template")
    p_inspect_tmpl.add_argument('--format', choices=["json", "text"], default="text", help="Output format")
    p_inspect_tmpl.set_defaults(func=inspect_template_cmd)
    
    # Validate
    p_validate = subparsers.add_parser('validate', help="Validate an input deck")
    p_validate.add_argument('input_file', help="Path to input YAML or MD file")
    p_validate.add_argument('--format', dest="output_format", help="Output format (e.g. json)", default=None)
    p_validate.add_argument('--input-format', dest="format", help="Force input format (yaml or markdown)", default=None)
    p_validate.set_defaults(func=validate_cmd)
    
    # Build
    p_build = subparsers.add_parser('build', help="Build PPTX from input deck")
    p_build.add_argument('input_file', help="Path to input YAML or MD file")
    p_build.add_argument('output_file', help="Path to output PPTX file")
    p_build.add_argument('--template', help="Path to PPTX template file to use for rendering", default=None)
    p_build.add_argument('--input-format', dest="format", help="Force input format (yaml or markdown)", default=None)
    p_build.set_defaults(func=build_cmd)
    
    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)
        
    args.func(args)

if __name__ == "__main__":
    main()
