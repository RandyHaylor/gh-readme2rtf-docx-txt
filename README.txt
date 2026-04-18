# :rocket: gh-readme2rtf-docx-txt — Auto-Convert GitHub READMEs to RTF, DOCX & TXT :tada:

> **Your README.md, rendered as shareable RTF, DOCX, and TXT files — automatically, on every push.**

```
             ┌──▶  README.rtf  ── Word / LibreOffice / WordPad / TextEdit
README.md  ──┼──▶  README.docx ── Microsoft Word / Google Docs / Pages
             └──▶  README.txt  ── Plain text, any editor
```

:sparkles: **Embedded Images** | :art: **Syntax-Highlighted Code (RTF + DOCX)** | :link: **Clickable Links & Footnotes** | :warning: **GitHub Alerts** | :memo: **Full GFM Support** | :rocket: **Emoji That Actually Render**

---

## Why?

GitHub READMEs look great *on GitHub*. But what about:

- :briefcase: **Sharing with non-developers** who don't have GitHub accounts?
- :page_facing_up: **Attaching to proposals, grants, or reports** that need a real document?
- :airplane: **Reading offline** without a browser?
- :printer: **Printing** with proper formatting?

This action auto-generates rich, formatted `.rtf`, `.docx`, and `.txt` files from your `README.md` every time you push. The outputs open in any word processor on any OS — Word, LibreOffice, Google Docs, Pages, WordPad, TextEdit — no browser required.

## What You Get

| Feature | RTF | DOCX | TXT |
|---------|:---:|:---:|:---:|
| Headings (h1–h6) with sized styles | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| **Bold**, *italic*, ~~strikethrough~~, <ins>underline</ins>, <sub>sub</sub>, <sup>sup</sup> | :white_check_mark: | :white_check_mark: | — |
| Clickable hyperlinks | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| Internal section-anchor links (bookmarks) | :white_check_mark: | :white_check_mark: | — |
| Relative link resolution to GitHub blob URLs | :white_check_mark: | :white_check_mark: | :white_check_mark: |
| `@mention` → GitHub profile links | :white_check_mark: | :white_check_mark: | — |
| `#issue` → GitHub issue links | :white_check_mark: | :white_check_mark: | — |
| Fenced code blocks with **syntax highlighting** | :white_check_mark: | :white_check_mark: | — |
| Inline code with background shading | :white_check_mark: | :white_check_mark: | — |
| Embedded images with auto-scaling | :white_check_mark: | :white_check_mark: | — |
| Oversized image clamping to page bounds | :white_check_mark: | :white_check_mark: | — |
| Tables with borders + column alignment | :white_check_mark: | :white_check_mark: | — |
| Ordered, unordered, and task lists | :white_check_mark: | :white_check_mark: | — |
| GitHub alerts (Note, Tip, Important, Warning, Caution) | :white_check_mark: | :white_check_mark: | — |
| Blockquotes (including nested) | :white_check_mark: | :white_check_mark: | — |
| Footnotes with clickable references | :white_check_mark: | :white_check_mark: | — |
| Emoji shortcodes (`:rocket:` → 🚀) | :white_check_mark: | :white_check_mark: | — |
| Horizontal rules | :white_check_mark: | :white_check_mark: | — |

TXT output is a link-resolved plain-text render — markdown structure preserved, link targets rewritten, no binary formatting.

## Quick Start

### 1. Copy the action into your repo

```bash
mkdir -p .github/actions/gh-readme2rtf-docx-txt
cp action.yml gh-readme2rtf-docx-txt.py rtf_image_embedder.py \
   .github/actions/gh-readme2rtf-docx-txt/
cp gh-readme2rtf-docx-txt-settings.json .
```

### 2. Configure which files and formats to generate

Edit `gh-readme2rtf-docx-txt-settings.json` at the repo root:

```json
{
  "files": [
    {
      "input": "README.md",
      "output_formats": ["rtf", "docx", "txt"]
    }
  ]
}
```

Each entry picks any combination of `rtf`, `docx`, `txt`. Multiple `files` entries are supported — e.g. a top-level README plus submodule READMEs.

> **Settings location note:** the settings file lives at the repo root by default. If you prefer a different location (e.g. `config/readme-docs.json`), update the path in two places in `.github/workflows/gh-readme2rtf-docx-txt.yml`: the `paths:` trigger entry **and** the `SETTINGS_PATH` env var of the sync step. The workflow's self-sync logic keeps everything else in step from there.

### 3. Add the workflow

Copy `.github/workflows/gh-readme2rtf-docx-txt.yml` from this repo into your own — it's already wired up. Key pieces:

```yaml
on:
  workflow_dispatch:
  push:
    paths:
      - 'gh-readme2rtf-docx-txt-settings.json'
      # ↓ additional entries are auto-appended by the sync step below,
      #   mirroring every `input` path from the settings file.
      - 'README.md'

jobs:
  convert:
    steps:
      - uses: actions/checkout@v4

      - name: Sync workflow trigger paths to settings file
        # Reads the settings JSON, rewrites this file's `paths:` block if it
        # drifts. Commits a small workflow-only update when it changes.
        shell: python
        ...

      - name: Convert README
        uses: ./.github/actions/gh-readme2rtf-docx-txt
      # ... stage, commit, push generated files
```

See the workflow file in this repo for the full listing.

### 4. Push and done

The workflow is triggered by **any push that modifies the settings file or one of the tracked input files**. It regenerates the configured outputs and commits them back. You can also trigger it manually from the Actions tab.

**How the trigger stays in sync:** The `paths:` block starts seeded with just the settings file. On every run, the first step reads the settings file, computes the set of tracked inputs, and — if the workflow's `paths:` list doesn't already match — rewrites that block and commits the workflow file. So the next time you add or remove an `input` entry in the settings file, the workflow re-triggers itself, self-heals, and then future edits to the newly-tracked paths start firing the workflow automatically.

Converter source-code changes do **not** auto-trigger the workflow — run it manually from the Actions tab when you want to regenerate outputs after a code change.

## How It Works

Two Python files, zero heavy dependencies:

1. **`gh-readme2rtf-docx-txt.py`** — Single converter driven by a data-driven rule engine. Parses GitHub-Flavored Markdown and emits RTF, DOCX, or TXT based on the output extension. Uses [Pygments](https://pygments.org/) for syntax highlighting across both RTF and DOCX. Resolves `@mentions`, `#issues`, and relative links to GitHub URLs automatically.

2. **`rtf_image_embedder.py`** — Post-processor for the RTF path only. Finds `[Image: ...]` placeholders, reads referenced local images, downscales them with [Pillow](https://pillow.readthedocs.io/) to fit page bounds, and embeds them as hex-encoded `\pict` blocks. DOCX image embedding is handled inside the main converter using the OOXML package's native image relationships.

The composite action (`action.yml`) reads `gh-readme2rtf-docx-txt-settings.json` and dispatches the converter per file/format.

### Dependencies

- **Python 3** (preinstalled on GitHub Actions runners)
- **Pygments** (`pip install pygments`) — syntax highlighting for 500+ languages
- **Pillow** (`pip install pillow`) — image processing for embedding and downscaling

### Run Locally

```bash
pip install pygments pillow

# RTF (two steps: convert, then embed images)
python3 gh-readme2rtf-docx-txt.py README.md README.rtf
python3 rtf_image_embedder.py README.rtf

# DOCX (single step — images embedded inline)
python3 gh-readme2rtf-docx-txt.py README.md README.docx

# TXT (link-resolved plain text)
python3 gh-readme2rtf-docx-txt.py README.md README.txt
```

## Architecture at a Glance

```
README.md  ──▶  [ PARSER ]  ──▶  [ CONVERSION TABLE ]  ──▶  [ FORMAT POST-PASS ]  ──▶  output
                 block + inline     GFM element ─┬─▶ rtf       one cleanup pass
                 line by line                    ├─▶ docx      per file format
                                                 └─▶ txt
```

### 1. Parsing — read it line by line

The converter walks `README.md` top-to-bottom. Each line tries the **block handlers** in order (heading, fenced code, table, list, blockquote, footnote-def, paragraph…); first match wins. The matched block's text then flows through the **inline conversion table** for emphasis, links, emoji, `@mentions`, and so on.

### 2. The conversion table — strategy pattern, expressed as data

Instead of three parallel converters (`MarkdownToRtf`, `MarkdownToDocx`, `MarkdownToTxt`) with overridden methods, the codebase uses **one table** where each row is a GFM element and each column is a pure function that emits that element in one target format:

```python
INLINE_RULES = [
    # (name,           (pattern,                    {'rtf': ...,       'docx': ...}))
    ('bold_star',      (r'\*\*(.+?)\*\*',           {'rtf': r'{\b \1}',       'docx': docx_bold})),
    ('md_link',        (r'\[([^\]]+)\]\(([^)]+)\)', {'rtf': rtf_link,         'docx': docx_link})),
    ('emoji',          (r':\w+:',                   {'rtf': rtf_emoji,        'docx': docx_emoji})),
    ('footnote_ref',   (r'\[\^([^\]]+)\]',          {'rtf': rtf_footnote_ref, 'docx': docx_footnote_ref})),
    # ...
]

BLOCK_RULES = {
    'rtf':  [block_heading, block_fenced_code, block_table, block_list, ...],
    'docx': [docx_block_heading, docx_block_fenced_code, docx_block_table, ...],
}
```

The dispatcher is four lines: **look up the active format in the row's dict, call the function (or substitute the string), move on**. That's the **strategy pattern implemented functionally** — each strategy is a small pure function (or an `re.sub` replacement string), chosen by dict key at runtime. No base class, no subclass hierarchy, no visitor.

What this feels like when you edit the code:

- **Add a markdown feature** → add one row. (This is how `<sub>`, `<sup>`, `<ins>`, and the emoji shortcode map were added.)
- **Add a new output format** → add one key to every row. Dispatcher doesn't change.
- **Tweak how bold renders in DOCX** → edit one cell. Other formats untouched.
- Rules run in **phases** (HTML strip → escapes → links/images → inline code stash → GitHub refs → formatting), so later rows can assume earlier phases have already happened.

### 3. Post-cleanup per doc type

The conversion table handles *content*. Format-specific *plumbing* lives in a small post-pass so the core stays format-agnostic:

| Format | Post-pass | Why |
|---|---|---|
| **RTF** | `rtf_image_embedder.py` swaps `[Image: ...]` markers for hex-encoded `\pict` blocks | Images get sized against page bounds only after layout is known |
| **DOCX** | `<w:sectPr>` gets injected into the last paragraph's `<w:pPr>` | A body-level sectPr makes Word render a ghost trailing blank page |
| **TXT** | `_txt_resolve_relative_links_only` rewrites relative links to full GitHub URLs | Plain text has nothing else to process |

## Viewing Tips

When opening the RTF in LibreOffice Writer:

- **Disable red squiggles**: Tools > Automatic Spell Checking (toggle off)
- **Read-only mode**: Edit > Edit Mode (toggle off) for a clean reading view
- **Full screen**: `Ctrl+Shift+J`

The DOCX opens natively in Microsoft Word, Google Docs, and Pages — no squiggle toggling needed.

---

# :test_tube: Test Content

> Everything below exercises every GitHub-Flavored Markdown element supported by the converter. The generated `README.rtf`, `README.docx`, and `README.txt` in this repo are living proof it all works — open any of them and compare!

## Heading Level 2

### Heading Level 3

#### Heading Level 4

##### Heading Level 5

###### Heading Level 6

## Text Formatting

This is **bold text** and this is __also bold__. This is *italic text* and this is _also italic_. This is ***bold and italic*** together. This is ~~strikethrough text~~. This is <sub>subscript</sub> and this is <sup>superscript</sup>. This is <ins>underlined</ins> text.

Escaped special characters: \*not italic\* and \#not a heading.

Line break with two trailing spaces:  
This is on the next line.

Line break with backslash:\
This is also on the next line.

And a <br/>hard break via HTML.

## Links

Here is an [inline link](https://example.com) and a [link to a section](#tables).

Here is a relative link to [another file](https://github.com/RandyHaylor/gh-readme2rtf-docx-txt/blob/main/SKILL.md).

Bare URL autolink: https://github.com

GitHub references: @octocat and issue #42.

## Images

Standard markdown image (no size control):

![Alt text for an image](https://github.com/RandyHaylor/gh-readme2rtf-docx-txt/blob/main/test-image.png)

HTML image with custom width:

<img src="test-image.png" width="300" alt="Image at 300px wide">

HTML image with custom height only (width should auto-calculate from aspect ratio):

<img src="test-image.png" height="100" alt="Image at 100px tall">

HTML image with custom width and height:

<img src="test-image.png" width="150" height="75" alt="Image at 150x75">

HTML image with percentage width:

<img src="test-image.png" width="50%" alt="Image at 50% width">

Picture element for light/dark theme variants:

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="test-image.png">
  <source media="(prefers-color-scheme: light)" srcset="test-image.png">
  <img alt="Theme-aware image" src="test-image.png" width="200">
</picture>

Large wide image (should be clamped to page width):

<img src="large-img-wide.png" alt="Large wide image">

Large tall image (should be clamped to page height):

<img src="large-img-tall.png" alt="Large tall image">

## Blockquotes

> This is a blockquote.
> It can span multiple lines.
>
> > This is a nested blockquote with actual content inside it.

## Alerts

> [!NOTE]
> This is a note alert with helpful information.

> [!TIP]
> This is a tip alert with a suggestion.

> [!IMPORTANT]
> This is an important alert.

> [!WARNING]
> This is a warning alert.

> [!CAUTION]
> This is a caution alert.

## Unordered Lists

- First item
- Second item
  - Nested item A
  - Nested item B
    - Deeply nested item
- Third item

Alternative markers:

* Star item one
* Star item two

+ Plus item one
+ Plus item two

## Ordered Lists

1. First step
2. Second step
   1. Sub-step A
   2. Sub-step B
3. Third step

## Task Lists

- [x] Completed task
- [ ] Incomplete task
- [x] Another done task
  - [ ] Nested incomplete subtask
  - [x] Nested complete subtask

## Inline Code

Use the `println()` function to print output. The config file is at `~/.config/app.json`.

## Fenced Code Blocks

```kotlin
fun main() {
    val greeting = "Hello, World!"
    println(greeting)
}
```

```xml
<uses-permission android:name="android.permission.INTERNET" />
```

```bash
#!/bin/bash
echo "No language-specific highlighting"
```

```
Plain code block with no language specified.
```

## Tables

| Feature | Status | Notes |
|---------|--------|-------|
| Bold | Supported | Uses `**text**` |
| Italic | Supported | Uses `*text*` |
| Tables | Supported | You're looking at one |

Left, center, and right alignment:

| Left | Center | Right |
|:-----|:------:|------:|
| L1 | C1 | R1 |
| L2 | C2 | R2 |

## Horizontal Rules

---

***

___

## Footnotes

Here is a sentence with a footnote reference[^1]. And another[^longnote].

[^1]: This is the first footnote.
[^longnote]: This is a longer footnote with multiple words.

## Emoji

Emoji shortcodes: `:rocket:` :rocket:, `:tada:` :tada:, `:warning:` :warning:

## Color Codes

GitHub renders color swatches for these inline code formats:

`#FF5733` `#28A745` `rgb(54, 95, 145)` `hsl(210, 45%, 39%)`

## HTML Comments

<!-- This comment should not appear in rendered output -->

Visible text after a hidden comment.

## Mixed Nesting

1. Ordered item with **bold** and `code`
   - Unordered sub-item with a [link](https://example.com)
   - Another sub-item with *italic*
     > A blockquote inside a list
2. Second ordered item
   ```python
   # Code block inside a list
   print("hello")
   ```
3. Third item with ~~strikethrough~~
