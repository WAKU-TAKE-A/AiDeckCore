# README_AI

You are generating a PowerPoint deck for this repository.

Use `deck2pptx` through the canonical `Deck` model. Do not manually specify PowerPoint coordinates, text box positions, or shape dimensions. YAML and Markdown are only input adapters. The renderer consumes `Deck`, not adapter-specific syntax.

## Core Rule

The `Deck` model is canonical.

- YAML is only an input adapter.
- Markdown is only an input adapter.
- Future AsciiDoc or Natural Language inputs must target the same `Deck` model.
- Renderer changes must not be required just because a new input adapter is added.
- PowerPoint is only one renderer.

## Current Supported Elements

**Supported block-level elements**:
  - `Text`: plain paragraphs
  - `BulletList`: generated from `-` or `*`
  - `Table`: standard Markdown tables
  - `Image`: `![alt](path)` (automatically preserves alt text as a caption)
  - `Gallery`: explicit grid layout of multiple images using `<!-- gallery [cols] -->` (e.g. `<!-- gallery 3 -->`).
  - `Flow`: simple flowchart (` ```flow `)
  - `Comparison`: ` ```comparison ` block with columns labeled by `Label:` and `- item` lists
  - `Timeline`: ` ```timeline ` block with `Date: Title - Description`
  - `CodeBlock`: ` ```code python ` block for source code
  - `Tree`: ` ```tree ` block for hierarchical structures
  - `Split`: a multi-panel layout splitting the slide area `horizontal` or `vertical`.

Avoid using generic tables and bullets for comparisons, timelines, code, or hierarchy. Use the semantic elements designed for business presentations.

## Markdown Control Comments

Markdown uses HTML comments for slide controls and structure. Multiple commands can be combined with `;`. String values should be quoted (e.g. `"TitleLayout"`).

- **Layout & Targeting**: `<!-- layout "Name" -->`, `<!-- subtitle "Text" -->`, `<!-- placeholder "Name" -->`
- **Short aliases**: `<!-- l "Name" -->`, `<!-- sub "Text" -->`, `<!-- ph "Name" -->`
- **Alignment**: `<!-- align "top" -->` to set vertical alignment (`top`, `semi-top`, `normal`, `semi-bottom`, `bottom`).
- **Slide logic**: `#` headings start new slides. `<!-- newpage "Name" -->` (or `new`) explicitly creates a new slide.
- **Split/Panel Layouts**: Use `<!-- split h -->` or `<!-- split v -->` to create multi-panel regions. Add `<!-- panel "Title" -->` to define sections. Close with `<!-- /split -->`. Nested splits are not supported. Note: The `style` property, font size controls, and weighted panel rendering are future work and are NOT implemented.

## Required Authoring Loop

Run commands with the repo-local virtual environment:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx explain-spec --format json
```

Use the returned schema as the source of truth.

When a PowerPoint template is involved, always inspect its available layouts and placeholders first:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx inspect-template template.pptx --format json
```

Then:

1. Generate an input file in YAML or Markdown.
   - Do not modify project code unless the user explicitly asks for code changes. Prefer writing/converting presentation input files over changing code.
   - If Markdown cannot express the requested deck cleanly, convert the context into YAML targeting the same Deck model.
   - Element-level placeholder binding: `<!-- placeholder="ExactName" -->` to route an element to a specific placeholder shape on the slide layout. Layout and placeholder names are matched by case-insensitive prefix. The renderer maps matching layout placeholder names to the generated slide placeholder with the same idx. This handles PowerPoint-created names such as `Subtitle 2`. Note: Missing placeholders fall back to default auto-positioning.

### Template-Aware Workflow

You can author presentations targeting a specific user PPTX template. Follow this workflow:

1. **Inspect the Template**
   ```powershell
   deck2pptx inspect-template template.pptx --format json
   ```
   This reveals the exact layout names and placeholder names available.

2. **Author the Input (Markdown/YAML)**
   Use the exact layout and placeholder names discovered, or a stable prefix of those names. Matching is case-insensitive, so `title&body` can target `Title&BodyLayout`, and `sub` can target `Subtitle 2`. Do NOT guess aliases like `subtitle -> SubTitle` or fix typos like `SubTitile` unless explicitly requested. In Markdown, `#`, `##`, and `###` start new slides; `####` and deeper headings stay in the slide body.
   ```markdown
   <!-- layout="TitleLayout" -->
   # My Presentation
   <!-- subtitle="SubTitile" -->
   ```

3. **Inspect the Parsed Deck**
   ```powershell
   deck2pptx inspect input.md --format json
   ```

4. **Validate the Deck**
   ```powershell
   deck2pptx validate input.md --format json
   ```

5. **Build Using the Template**
   ```powershell
   deck2pptx build input.md output.pptx --template template.pptx
   ```
   Fallback behavior is active: if a layout or placeholder name is not found in the template, the default behavior will be used without crashing.

## Architecture Guidelines

1. Inspect the parsed Deck model:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx inspect your_file.md --format json
```

2. Validate the input:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx validate your_file.md --format json
```

3. If validation fails, repair the input using the structured `errors` list. Do not guess.
4. Build the PPTX only after validation passes:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx build your_file.md output.pptx
```

## Input Guidance

Prefer Markdown when the user wants a simple text-first deck. Use YAML for complex structures.
Concierge-layer guidance: If the user provides custom or ad-hoc Markdown formatting rules for a deck, briefly explain how their format maps to the supported elements, and automatically convert their intent to valid Markdown/YAML. Do not ask the user to rewrite their input first.

### Document Structure & Typography
You can control the overall deck typography and structure using YAML front matter (or Markdown `---` block):
- `toc`: Set to `true` to automatically generate a Table of Contents slide containing all slide titles.
- `toc_title`: Override the default "Table of Contents" title.
- `indent`: Controls list hierarchy mapping in Markdown. Specifies how many spaces equal one list indentation level (default is 2).
- `font_size_l0` to `font_size_l4`: Override the default font sizes for text levels (e.g., `font_size_l0: 24`, `font_size_l1: 20`). These propagate to bullet list levels, text boxes, and tables.

For images, use paths relative to the input file location.

For Flow, define nodes and edges using supported IDs. Always validate because edges must reference existing node IDs.

## Quality Gate

Before reporting completion, run at least:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx validate your_file.md --format json
.\.venv\Scripts\python.exe -m deck2pptx build your_file.md output.pptx
```

For repository-level verification, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\verify_release.ps1
```

## Do Not

- Do not modify project code unless asked.
- Do not add `x`, `y`, `width`, or `height` fields to Deck input.
- Do not make the renderer parse YAML or Markdown directly.
- Do not invent unsupported elements without updating the Deck model, adapters, validation, renderer, spec, docs, and tests together.
- Do not ignore validation errors.
- Do not claim a PPTX is ready until validation and build have both succeeded.

## For System Developers / Maintainers

If the user asks you to modify the `deck2pptx` core system (e.g., Python code, renderer logic, adding new elements), you are acting as a **System Developer**, not just a Presentation Author.

In this case, you MUST read the strict repository policies located in the `docs/` directory before making changes or reporting completion:
- `docs/source-baseline.md`: Rules for repository hygiene, tracked files, and allowed artifacts.
- `docs/release-verification.md`: Rules for running the release gate pipeline (`scripts/verify_release.ps1`) to ensure your code changes do not break the system.
