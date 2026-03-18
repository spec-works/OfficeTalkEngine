---
name: officetalk-cli
description: >
  Apply deterministic operations to Microsoft Office documents using the OfficeTalk CLI tool.
  Use this skill when the user wants to modify Word (.docx), Excel (.xlsx), or PowerPoint (.pptx)
  documents programmatically — fixing typos, updating headings, reformatting tables, inserting
  sections, or applying bulk changes. Activate when the user mentions OfficeTalk, .otk files,
  document editing via CLI, or wants to script changes to Office documents. Also useful when an
  LLM needs to generate structured document modifications.
license: MIT
compatibility: Requires .NET 9.0 or later SDK. Works on Windows, macOS, and Linux.
metadata:
  author: spec-works
  version: "0.4"
  repository: https://github.com/spec-works/OfficeTalkEngine
---

# OfficeTalk CLI — Scripted Office Document Editing

Apply deterministic, repeatable operations to Microsoft Office documents (Word, Excel, PowerPoint) using a human-readable, LLM-friendly grammar.

OfficeTalk is a structured document format (`application/officetalk`, file extension `.otk`) that describes precise modifications to Office documents. The CLI tool parses `.otk` files, validates them, resolves addresses against target documents, and applies operations using the Office Open XML SDK.

## Prerequisites — Installing .NET

The OfficeTalk CLI requires .NET 9.0 or later. Install the SDK for your platform:

### Windows

```powershell
# Using winget (recommended)
winget install Microsoft.DotNet.SDK.9

# Or download from https://dotnet.microsoft.com/download/dotnet/9.0
```

### macOS

```bash
# Using Homebrew (recommended)
brew install dotnet-sdk

# Or download from https://dotnet.microsoft.com/download/dotnet/9.0
```

### Linux (Ubuntu/Debian)

```bash
wget https://packages.microsoft.com/config/ubuntu/24.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y dotnet-sdk-9.0
```

### Linux (Fedora/RHEL)

```bash
sudo dnf install dotnet-sdk-9.0
```

### Verify Installation

```bash
dotnet --version
# Should show 9.x.x or later
```

## Installing the CLI Tool

```bash
dotnet tool install --global SpecWorks.OfficeTalk.CLI
```

Verify it works:

```bash
officetalk version
```

## OfficeTalk Document Format

An OfficeTalk document (`.otk` file) is a UTF-8 text file with a line-oriented grammar. Every document starts with a version header and document type, followed by operation blocks.

### Document Structure

```
OFFICETALK/1.0
DOCTYPE word|excel|powerpoint

# Comments start with #

AT <address>
<OPERATION> [arguments]
<OPERATION> [arguments]

AT <address>
<OPERATION> [arguments]
```

### Document Types

| DOCTYPE | Target | File Extension |
|---------|--------|----------------|
| `word` | Word documents | .docx |
| `excel` | Excel workbooks | .xlsx |
| `powerpoint` | PowerPoint presentations | .pptx |

## Addressing

Addresses are `/`-separated paths that locate content within an Office document. They use a simplified syntax — no XML namespaces, no Open XML internals.

### Path Segments

> **Important:** `heading` and `paragraph` are distinct segment types. A heading is a paragraph with a heading style, but `paragraph[1]` does **not** count headings — it finds the first non-heading paragraph. Use `heading[...]` when targeting headings (e.g., document titles, section headers) and `paragraph[...]` when targeting body text. Using `paragraph` to reach a heading will target the wrong element.

#### Word

| Segment | Description |
|---------|-------------|
| `body` | Document body |
| `paragraph` | A paragraph |
| `heading` | A paragraph with a heading style |
| `run` | A text run within a paragraph |
| `table` | A table |
| `row` | A table row |
| `cell` | A table cell |
| `list` | A list |
| `item` | A list item |
| `image` | An inline or anchored image |
| `section` | A document section |
| `header` | Page header |
| `footer` | Page footer |
| `bookmark` | A named bookmark |
| `content-control` | A structured document tag |

#### Excel

| Segment | Description |
|---------|-------------|
| `sheet` | A worksheet |
| `range` | A cell range |
| `column` | A column |
| `row` | A table row |
| `cell` | A single cell |

#### PowerPoint

| Segment | Description |
|---------|-------------|
| `slide` | A slide |
| `shape` | A shape on a slide |
| `title` | Title placeholder |
| `subtitle` | Subtitle placeholder |
| `notes` | Speaker notes |

### Predicates

Predicates filter elements and go in square brackets after a segment.

#### Positional (1-based)

```
paragraph[3]              # third paragraph
table[1]/row[2]           # second row of first table
slide[5]                  # fifth slide
```

#### Key-Value

```
heading[level=2]                      # H2 heading
heading[level=2, text="Background"]   # specific H2 by text
table[caption="Results"]              # table by caption
sheet["Revenue"]                      # sheet by name (shorthand)
shape[name="Logo"]                    # shape by name
header[type=default]                  # default page header
content-control[tag="author-name"]    # content control by tag
```

#### Text Matching

| Operator | Meaning | Example |
|----------|---------|---------|
| `text="..."` | Exact match | `paragraph[text="Hello World"]` |
| `text~="..."` | Regex match (I-Regexp) | `paragraph[text~="^Chapter \d+"]` |
| `text^="..."` | Starts with | `paragraph[text^="In conclusion"]` |
| `text$="..."` | Ends with | `paragraph[text$="respectively."]` |
| `text*="..."` | Contains | `paragraph[text*="important"]` |

#### Disambiguating Multiple Matches

```
paragraph[text*="revenue"][2]    # second paragraph containing "revenue"
```

### Address Examples

```
# Word
body/paragraph[1]
body/heading[level=1]
body/heading[level=2, text="Methods"]
body/paragraph[text*="conclusion"]
body/table[1]/row[3]/cell[2]
body/table[caption="Results"]/row[1]
body/bookmark["references"]
header[type=default]

# Excel
sheet["Revenue"]/B7
sheet["Revenue"]/A1:D10
sheet[1]/row[5]

# PowerPoint
slide[1]/title
slide[3]/shape[name="Chart1"]
slide[2]/notes
```

## Operations

### Content Operations

#### SET — Replace element content

```
AT body/paragraph[3]
SET "New paragraph text."

AT body/paragraph[3]
SET <<<
This paragraph now contains
multiple lines of text.
>>>
```

#### REPLACE — Find and replace text

```
AT body/paragraph[1]
REPLACE "FY2024" WITH "FY2025"

# Replace all occurrences
AT body
REPLACE ALL "colour" WITH "color"
```

#### INSERT BEFORE / INSERT AFTER — Add sibling content

```
AT body/heading[text="Conclusion"]
INSERT BEFORE <<<
Recommendations

Based on the findings presented in this report, we recommend:

1. Increase investment in renewable energy.
2. Expand the remote work program.
>>>

AT body/paragraph[1]
INSERT AFTER "A new paragraph after the first one."
```

#### INSERT IMAGE — Embed an image

```
AT body/paragraph[3]
INSERT IMAGE AFTER "diagram.png"
  alt="Architecture diagram"
  width=5in
  height=3in

AT body/paragraph[1]
INSERT IMAGE BEFORE "https://example.com/logo.png"
  alt="Company logo"
  width=2in
```

#### INSERT TABLE — Create a new table

```
AT body/paragraph[5]
INSERT TABLE AFTER rows=4, columns=3

AT body/table[1]/row[1]
SET CELLS "Name", "Role", "Department"

AT body/table[1]/row[2]
SET CELLS "Alice", "Engineer", "Platform"
```

#### INSERT LIST — Create ordered or unordered lists

```
AT body/paragraph[3]
INSERT LIST AFTER unordered
  ITEM "First bullet point"
  ITEM "Second bullet point"
  ITEM "Sub-item under second" nested
  ITEM "Third bullet point"

AT body/paragraph[5]
INSERT LIST AFTER ordered
  ITEM "Step one"
  ITEM "Step two"
  ITEM "Step three"
```

#### LINK — Create a hyperlink

```
AT body/paragraph[2]
LINK "https://example.com"

AT body/paragraph[2]/run[3]
LINK "https://docs.example.com"
```

#### SET RUNS — Replace paragraph content with individually formatted runs

```
AT body/paragraph[1]
SET RUNS
  RUN "This is "
  RUN "bold" bold=true
  RUN " and "
  RUN "italic" italic=true
  RUN " text."

# Inline code with background shading
AT body/paragraph[2]
SET RUNS
  RUN "Run "
  RUN "npm start" font-name="Consolas" font-size=10pt background-color=#F0F0F0
  RUN " to launch the server."

# Hyperlink within formatted runs
AT body/paragraph[3]
SET RUNS
  RUN "Visit "
  RUN "our website" href="https://example.com" color=#0563C1 underline=single
  RUN " for details."
```

#### DELETE — Remove an element

```
AT body/paragraph[text*="DRAFT"]
DELETE
```

#### APPEND / PREPEND — Add to existing content

```
AT body/paragraph[1]
APPEND " (see Appendix A)"

AT sheet["Notes"]/A1
PREPEND "UPDATED: "
```

### Formatting Operations

#### FORMAT — Apply formatting properties

```
AT body/paragraph[1]
FORMAT bold=true, font-size=14pt, color=#2B579A

AT body/heading[level=1]
FORMAT font-name="Aptos Display", font-size=28pt
FORMAT color=#1F3864, spacing-after=12pt
```

When FORMAT follows a content operation in the same block, it formats the produced content:

```
AT body/heading[text="Summary"]
INSERT AFTER "Key Findings"
FORMAT bold=true, font-size=16pt
```

#### STYLE — Apply a named style

```
AT body/paragraph[5]
STYLE "Heading 2"

AT body/table[1]
STYLE "Grid Table 4 - Accent 1"
```

### Structural Operations

#### Table Operations

```
AT body/table[1]/row[3]
INSERT ROW AFTER
SET CELLS "Product D", "450", "12%", "Active"

AT body/table[1]/row[5]
DELETE ROW

AT body/table[1]/row[2]/cell[1]
MERGE CELLS TO row[2]/cell[3]
```

#### Slide Operations (PowerPoint)

```
AT slide[3]
INSERT SLIDE AFTER
SET title "New Section"
SET subtitle "Overview of Changes"

AT slide[5]
DELETE SLIDE

AT slide[3]
DUPLICATE SLIDE
```

#### Sheet Operations (Excel)

```
ADD SHEET "Summary"

AT sheet["Old Name"]
RENAME SHEET "New Name"
```

### Metadata Operations

```
PROPERTY title="Quarterly Report Q1 2026"
PROPERTY author="Finance Team"
```

## Bulk Operations with AT EACH

Use `AT EACH` to apply operations to all matching elements:

```
AT EACH body/paragraph[text*="TODO"]
FORMAT highlight=#FFFF00, bold=true

AT EACH body/heading[level=2]
FORMAT color=#2B579A
```

**Important:** `AT EACH` requires at least one match — zero matches is an error. Use it for bulk formatting, styling, and replacement. Exercise caution with destructive operations like DELETE.

## Data Types

| Type | Examples |
|------|---------|
| Strings | `"Hello"`, `"She said \"hi\""`, `"Line 1\nLine 2"` |
| Numbers | `42`, `3.14`, `-7` |
| Booleans | `true`, `false` |
| Colors | `#2B579A`, `#FF0000`, `red`, `cornflowerblue` |
| Lengths | `12pt`, `1.5in`, `2.54cm`, `50%`, `914400emu` |
| Content blocks | `<<<` ... `>>>` (multi-line text, each on its own line) |

## Formatting Properties Reference

### Text Properties

`font-name`, `font-size`, `bold`, `italic`, `underline` (single/double/dotted/dashed/wavy/none), `strikethrough`, `color`, `highlight` (named colors: yellow/green/cyan/magenta/red/blue/etc.), `background-color` (arbitrary hex shading, e.g. #F0F0F0), `superscript`, `subscript`, `small-caps`, `all-caps`, `href` (hyperlink URL on a run, use `none` to remove)

### Paragraph Properties

`alignment` (left/center/right/justify), `spacing-before`, `spacing-after`, `line-spacing`, `indent-left`, `indent-right`, `indent-first-line`, `indent-hanging`, `keep-with-next`, `page-break-before`, `outline-level` (0-8), `border-top`, `border-bottom`, `border-left`, `border-right` (single/double/thick/dashed/dotted/none), `border-color`, `border-width`

### Table/Cell Properties

`width`, `alignment`, `border`, `border-color`, `border-width`, `cell-padding`, `fill-color`, `vertical-align`, `wrap-text`, `number-format`

### Image Properties

`width`, `height`, `alt`, `position` (inline/anchor)

## Processing Model — Snapshot Semantics

**Critical concept:** All addresses are resolved against the original document BEFORE any operations are applied. This means:

- `paragraph[5]` always refers to the fifth paragraph in the original document
- Operations that insert or delete elements do NOT shift indices for later blocks
- Every block operates on the document as it existed before any modifications

This makes OfficeTalk documents safe to produce — the author reasons about the current state of the document, not a partially-modified intermediate state.

## CLI Commands

### `officetalk apply` — Execute .otk against a document

```bash
# Apply changes in-place
officetalk apply -i changes.otk -t report.docx

# Write to a new file
officetalk apply -i changes.otk -t report.docx -o report-updated.docx

# Dry run — validate and resolve without modifying
officetalk apply -i changes.otk -t report.docx --dry-run

# Verbose output
officetalk apply -i changes.otk -t report.docx -v

# Pipe OfficeTalk from stdin
echo "OFFICETALK/1.0
DOCTYPE word
AT body/heading[1]
SET \"Updated Title\"" | officetalk apply -t report.docx
```

| Option | Alias | Description |
|--------|-------|-------------|
| `--input` | `-i` | Path to .otk file (or pipe via stdin) |
| `--target` | `-t` | Target Office document (required) |
| `--output` | `-o` | Output path (default: modify in-place) |
| `--force` | — | Overwrite output if exists |
| `--dry-run` | — | Validate without applying |
| `--verbose` | `-v` | Show detailed operation log |

### `officetalk validate` — Check .otk for errors

```bash
# Syntax-only validation (no target doc needed)
officetalk validate -i changes.otk --syntax-only

# Full validation with semantic checks
officetalk validate -i changes.otk -t report.docx

# Machine-readable JSON output
officetalk validate -i changes.otk -f json
```

| Option | Alias | Description |
|--------|-------|-------------|
| `--input` | `-i` | Path to .otk file (required) |
| `--target` | `-t` | Target document for semantic validation |
| `--syntax-only` | — | Skip semantic validation |
| `--format` | `-f` | Output format: `text` (default) or `json` |

### `officetalk parse` — Dump AST as JSON

```bash
officetalk parse -i changes.otk
officetalk parse -i changes.otk --pretty
```

### `officetalk inspect` — Explore addresses

```bash
# See what an address resolves to
officetalk inspect -t report.docx -a "body/heading[level=1]"

# With surrounding context
officetalk inspect -t report.docx -a "body/paragraph[text*='revenue']" -c 2
```

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Runtime/IO error |
| 2 | Parse or validation error |
| 3 | Address resolution failure |

## Writing OfficeTalk Documents — Best Practices

### 1. Always Start with Header

Every `.otk` file must begin with exactly:
```
OFFICETALK/1.0
DOCTYPE word
```
(Replace `word` with `excel` or `powerpoint` as needed.)

### 2. Use Comments to Document Intent

```
# Fix the typo reported in issue #42
AT body/paragraph[text*="teh company"]
REPLACE "teh" WITH "the"
```

### 3. Prefer Text Predicates Over Positional

Text-based addresses are more robust — they survive document edits:
```
# Good — survives if paragraphs are added/removed above
AT body/heading[text="Conclusion"]

# Fragile — breaks if document structure changes
AT body/heading[5]
```

### 4. Use Specific Addresses

Be as specific as possible to avoid ambiguity errors:
```
# Good — specific
AT body/heading[level=2, text="Methods"]

# Risky — may match multiple H2 headings
AT body/heading[level=2]
```

### 5. Use `heading` for Headings, `paragraph` for Body Text

The `heading` and `paragraph` segments have **separate numbering**. A common mistake is using `paragraph[1]` to target a document title — this will miss the heading and hit the first body paragraph instead:
```
# WRONG — targets first body paragraph, NOT the title
AT body/paragraph[1]
FORMAT color=#008000

# CORRECT — targets the document title (first heading)
AT body/heading[1]
FORMAT color=#008000

# ALSO CORRECT — target by text for precision
AT body/heading[text="My Document Title"]
FORMAT color=#008000
```

### 6. Group Related Changes

Put related formatting in the same block:
```
AT body/heading[level=1]
SET "Annual Report — FY2026"
FORMAT font-size=28pt, color=#1F3864
STYLE "Title"
```

### 7. Use Content Blocks for Multi-line Text

```
AT body/paragraph[text="placeholder"]
SET <<<
First paragraph of the new content.

Second paragraph, separated by a blank line.
This continues the second paragraph.
>>>
```

### 8. Use AT EACH for Bulk Operations

```
# Highlight all TODO markers
AT EACH body/paragraph[text*="TODO"]
FORMAT highlight=#FFFF00

# Restyle all H2 headings
AT EACH body/heading[level=2]
FORMAT color=#2B579A, font-size=18pt
```

### 9. Validate Before Applying

```bash
# Always validate first, especially for LLM-generated .otk files
officetalk validate -i changes.otk -t report.docx
# If exit code 0, safe to apply
officetalk apply -i changes.otk -t report.docx
```

## Complete Examples

### Fix Typos and Update Metadata

```
OFFICETALK/1.0
DOCTYPE word

# Fix typo in introduction
AT body/paragraph[text*="teh company"]
REPLACE "teh" WITH "the"

# Update the document title
AT body/heading[level=1]
SET "Annual Report — FY2026"
FORMAT font-size=28pt, color=#1F3864

# Remove draft notice
AT body/paragraph[text="DRAFT — DO NOT DISTRIBUTE"]
DELETE

# Update metadata
PROPERTY title="Annual Report — FY2026"
PROPERTY author="Finance Team"
```

### Insert a New Section

```
OFFICETALK/1.0
DOCTYPE word

AT body/heading[text="Conclusion"]
INSERT BEFORE <<<
Recommendations

Based on the findings presented in this report, the committee
recommends the following actions:

1. Increase investment in renewable energy infrastructure.
2. Expand the remote work pilot program.
3. Commission an independent audit of supply chain practices.
>>>

AT body/paragraph[text="Recommendations"]
STYLE "Heading 2"
```

### Update a Table

```
OFFICETALK/1.0
DOCTYPE word

# Add a new data row
AT body/table[caption="Quarterly Results"]/row[4]
INSERT ROW AFTER
SET CELLS "Q4", "$2.1M", "$1.8M", "16.7%"

# Highlight the header row
AT body/table[caption="Quarterly Results"]/row[1]
FORMAT bold=true, fill-color=#2B579A, color=#FFFFFF
```

### Update an Excel Spreadsheet

```
OFFICETALK/1.0
DOCTYPE excel

AT sheet["Revenue"]/B7
SET "1250.00"
FORMAT number-format="#,##0.00", bold=true

AT sheet["Revenue"]/A1:F1
FORMAT font-name="Calibri", font-size=12pt, bold=true
FORMAT fill-color=#4472C4, color=#FFFFFF

AT sheet["Revenue"]/row[15]
INSERT ROW AFTER
SET CELLS "Product E", "North", "450", "12.5%", "Active", "2026-01-15"
```

### Update a PowerPoint Presentation

```
OFFICETALK/1.0
DOCTYPE powerpoint

AT slide[1]/title
SET "Q1 2026 Business Review"
FORMAT font-size=40pt, bold=true, color=#1F3864

AT slide[1]/subtitle
SET "Prepared by the Strategy Team — March 2026"

AT slide[3]
INSERT SLIDE AFTER
SET title "Key Metrics"
SET body <<<
Revenue grew 12% year-over-year.
Customer retention improved to 94%.
Three new enterprise clients onboarded.
>>>
```

## Limitations — What OfficeTalk Cannot Do

1. **No code execution** — OfficeTalk is purely declarative; no formulas, macros, or scripts
2. **No external references** — content blocks are literal text only; no URL fetching or file includes (image URLs are fetched at execution time by the engine)
3. **No chart creation** — charts must exist in the document already
4. **No template application** — cannot apply .dotx/.potx templates
5. **Excel/PowerPoint support is partial** — Word documents have the most complete implementation
6. **No document creation** — OfficeTalk edits existing documents; the calling tool creates blank/template documents first

## Composing with Other Tools

```bash
# Convert markdown to Word, then apply OfficeTalk edits
markmyword convert -i doc.md -o doc.docx && officetalk apply -i polish.otk -t doc.docx

# Generate an .otk file with an LLM, validate, then apply
llm "Fix all typos" --format officetalk > fixes.otk
officetalk validate -i fixes.otk -t report.docx && officetalk apply -i fixes.otk -t report.docx

# Validate OfficeTalk scripts in CI
for f in scripts/*.otk; do officetalk validate -i "$f" --syntax-only; done
```

## Installing This Skill

### GitHub Copilot CLI (personal)

```bash
git clone https://github.com/spec-works/OfficeTalkEngine.git /tmp/OfficeTalkEngine
mkdir -p ~/.copilot/skills/officetalk-cli
cp -r /tmp/OfficeTalkEngine/skills/officetalk-cli/* ~/.copilot/skills/officetalk-cli/
```

### GitHub Copilot CLI (project)

```bash
mkdir -p .github/skills/officetalk-cli
cp -r /path/to/OfficeTalkEngine/skills/officetalk-cli/* .github/skills/officetalk-cli/
git add .github/skills/
git commit -m "Add OfficeTalk CLI agent skill"
```

### Claude Code (personal)

```bash
mkdir -p ~/.claude/skills/officetalk-cli
cp -r /path/to/OfficeTalkEngine/skills/officetalk-cli/* ~/.claude/skills/officetalk-cli/
```

### Claude Code (project)

```bash
mkdir -p .claude/skills/officetalk-cli
cp -r /path/to/OfficeTalkEngine/skills/officetalk-cli/* .claude/skills/officetalk-cli/
```

### VS Code / Cursor (project)

```bash
mkdir -p .github/skills/officetalk-cli
cp -r /path/to/OfficeTalkEngine/skills/officetalk-cli/* .github/skills/officetalk-cli/
```

After installing, restart your agent session to pick up the new skill.

## Further Reference

- [OfficeTalk Specification](https://github.com/spec-works/OfficeTalk) — Full specification with formal grammar
- [OfficeTalkEngine](https://github.com/spec-works/OfficeTalkEngine) — Execution engine source code
