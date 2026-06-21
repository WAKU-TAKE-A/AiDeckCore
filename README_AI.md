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
  - `Mermaid`: ` ```mermaid ` block for advanced flowcharts and diagrams
  - `Tree`: ` ```tree ` block for hierarchical structures
  - `Split`: a multi-panel layout splitting the slide area `horizontal` or `vertical`.

*Note on Mermaid*:
- Mermaid diagrams (`` `mermaid ``) require the Mermaid CLI (`mmdc`) to render. If not installed or misconfigured, the renderer automatically falls back to rendering the raw code inside a `CodeBlock` to prevent crashes.
- `` `flow `` blocks are rendered natively as PowerPoint auto-shapes and do NOT require the Mermaid CLI.
- On Windows PowerShell, resolve npm binaries with `.cmd` (e.g. `npx.cmd` instead of `npx`) to bypass script execution policy blocks.

Avoid using generic tables and bullets for comparisons, timelines, code, or hierarchy. Use the semantic elements designed for business presentations.

## Markdown Control Comments

Markdown uses HTML comments for slide controls and structure. Multiple commands can be combined with `;`. String values should be quoted (e.g. `"TitleLayout"`).

| Command | Aliases | Example |
|---|---|---|
| `layout` | `l` | `<!-- l "Title&Body" -->` |
| `subtitle` | `sub` | `<!-- sub "My Subtitle" -->` |
| `placeholder` | `ph`, `place` | `<!-- ph "Body" -->` or `<!-- ph "Footer" "Hidden Text" -->` |
| `value` | `v` | `<!-- ph "Footer"; v "Hidden Text" -->` |
| `newpage` | `new`, `new_page` | `<!-- newpage "LayoutName" -->` |
| `align` | `content_align`, `valign` | `<!-- align "top" -->` |
| `gallery` | `gal` | `<!-- gallery 3 -->` |
| `split` | — | `<!-- split h -->` |
| `panel` | — | `<!-- panel "Title" -->` |
| `/split` | — | `<!-- /split -->` |

- **Hidden text injection**: You can inject text into placeholders without it appearing in the Markdown preview. Pass a second argument to `ph` or use the `v` command. `\n` or `<br>` are converted to newlines. Example: `<!-- ph "Footer" "Line 1\nLine 2" -->`
- **Slide logic**: `#`, `##`, `###` headings start new slides. `####` and deeper stay in the slide body.
- **Alignment values**: `top`, `semi-top`, `normal`, `semi-bottom`, `bottom`.
- **Split/Panel**: Use `<!-- split h -->` or `<!-- split v -->` to create multi-panel regions. Nested splits are not supported. The `style` property and weighted panel rendering are future work and NOT implemented.

## Authoring Workflow

Run commands with the repo-local virtual environment:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx explain-spec --format json
```

Use the returned schema as the source of truth.

### Basic Workflow

1. Generate an input file in YAML or Markdown.
   - Prefer Markdown for simple text-first decks. Use YAML for complex structures.
   - If the user provides custom formatting rules, automatically convert their intent to valid Markdown/YAML. Do not ask the user to rewrite their input first.
   - Do not modify project code unless the user explicitly asks.
2. Inspect the parsed Deck model:
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx inspect your_file.md --format json
   ```
3. Validate:
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx validate your_file.md --format json
   ```
4. If validation fails, repair the input using the structured `errors` list. Do not guess.
5. Build:
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx build your_file.md output.pptx
   ```

### Template-Aware Workflow

When a PowerPoint template is involved, always inspect it first:

1. **Inspect the Template**
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx inspect-template template.pptx --format json
   ```
   This reveals the exact layout names and placeholder names available.

2. **Author the Input**
   Use the exact layout and placeholder names discovered, or a stable prefix. Matching is case-insensitive, so `title&body` can target `Title&BodyLayout`, and `sub` can target `Subtitle 2`. Do NOT guess aliases or fix typos unless explicitly requested.
   ```markdown
   <!-- l "TitleLayout" -->
   # My Presentation
   <!-- sub "SubTitile" -->
   ```

3. **Inspect → Validate → Build** (same as basic workflow, adding `--template`):
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx build input.md output.pptx --template template.pptx
   ```
   **Calibration Flag**: If the user asks for exact text wrapping and bounding box heights based on the template's typography (e.g., matching the font heights and line spacing exactly), append the `--calib-first-slide` flag. This flag extracts physical dimension metrics from the text frames on the 1st slide of the template.
   ```powershell
   .\.venv\Scripts\python.exe -m deck2pptx build input.md output.pptx --template template.pptx --calib-first-slide
   ```
   Fallback behavior is active: if a layout or placeholder name is not found in the template, the default behavior will be used without crashing.

### Document Structure & Typography

You can control the overall deck typography and structure using YAML front matter (or Markdown `---` block):
- `toc`: Set to `true` to automatically generate a Table of Contents slide.
- `toc_title`: Override the default "Table of Contents" title.
- `indent`: Controls list hierarchy mapping in Markdown. How many spaces equal one list level (default: 2).


For images, use paths relative to the input file location.

For Flow, define nodes and edges using supported IDs. Always validate because edges must reference existing node IDs.

## Quality Gate

Before reporting completion, run at least:

```powershell
.\.venv\Scripts\python.exe -m deck2pptx validate your_file.md --format json
.\.venv\Scripts\python.exe -m deck2pptx build your_file.md output.pptx
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

In this case, you MUST read the following repository policies before making changes or reporting completion:
- `docs/source-baseline.md`: Defines which files are tracked source, generated artifacts, and operational data. Used by the hygiene checker.
- `docs/release-verification.md`: Describes the release gate pipeline and its stages.

Key scripts for system development:
- `scripts/verify_release.ps1`: Full release verification pipeline (tests → quality gate → negative validation → visual export → clean env → hygiene check).
- `scripts/check_source_hygiene.ps1`: Validates that the repository contains no unexpected untracked or modified files.
