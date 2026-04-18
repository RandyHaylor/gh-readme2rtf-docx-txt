🚀 README to RTF — Auto-Convert GitHub READMEs to Rich Documents 🎉
=================================================================

  Your README.md, beautifully rendered as a shareable RTF file — automatically, on every push.

    README.md  ──▶  README.rtf  ──▶  Open anywhere. Share with anyone.

✨ Embedded Images | :art: Syntax-Highlighted Code Blocks | :link: Clickable Links & Footnotes | ⚠️ GitHub Alerts | 📝 Full GFM Support | 🚀 Emoji That Actually Render

────────────────────────────────────────

Why?
----

GitHub READMEs look great on GitHub. But what about:

• :briefcase: Sharing with non-developers who don't have GitHub accounts?
• :page_facing_up: Attaching to proposals, grants, or reports that need a real document?
• :airplane: Reading offline without a browser?
• :printer: Printing with proper formatting?

This project auto-generates a rich, formatted RTF file from your README.md every time you push. The RTF opens in LibreOffice, Word, WordPad, TextEdit — any word processor on any OS.

What You Get
------------

Feature                                                         Status
--------------------------------------------------------------  ------
Headings (h1-h6) with sizes                                     ✅
Bold, italic, strikethrough                                     ✅
Clickable hyperlinks                                            ✅
Internal document links (section anchors)                       ✅
@mention (https://github.com/mention) links to GitHub profiles  ✅
#issue links to repo issues                                     ✅
Fenced code blocks with syntax highlighting                     ✅
Inline code with background shading                             ✅
Embedded images with auto-scaling                               ✅
Oversized image clamping to page bounds                         ✅
Tables with borders                                             ✅
Ordered, unordered, and task lists                              ✅
GitHub alerts (Note, Tip, Important, Warning, Caution)          ✅
Blockquotes                                                     ✅
Footnotes with clickable references                             ✅
Emoji via surrogate pairs                                       ✅
Horizontal rules                                                ✅

Quick Start
-----------

### 1. Copy the action into your repo

    mkdir -p .github/actions/readme-to-rtf
    cp action.yml gfm_markdown_to_rtf.py rtf_image_embedder.py .github/actions/readme-to-rtf/

### 2. Add the workflow

Create .github/workflows/readme-to-rtf.yml:

    name: Convert README to RTF
    
    on:
      workflow_dispatch:
      push:
        paths:
          - 'README.md'
    
    jobs:
      convert:
        runs-on: ubuntu-latest
        permissions:
          contents: write
        steps:
          - uses: actions/checkout@v4
    
          - name: Convert README to RTF
            uses: ./.github/actions/readme-to-rtf
    
          - name: Configure git
            shell: bash
            run: git config user.name "github-actions[bot]"
    
          - name: Configure git email
            shell: bash
            run: git config user.email "github-actions[bot]@users.noreply.github.com"
    
          - name: Stage RTF
            shell: bash
            run: git add README.rtf
    
          - name: Commit if changed
            shell: bash
            run: 'git diff --cached --quiet || git commit -m "auto-regenerate README.rtf from README.md"'
    
          - name: Push
            shell: bash
            run: git push

### 3. Push and done

Every time README.md changes, the workflow generates a fresh README.rtf and commits it to your repo. You can also trigger it manually from the Actions tab.

How It Works
------------

Two Python scripts, zero heavy dependencies:

1. gfm_markdown_to_rtf.py — Parses GitHub-Flavored Markdown and generates RTF with a data-driven rule engine. Uses Pygments (https://pygments.org/) for syntax highlighting. Resolves @mentions (https://github.com/mentions) and #issues to GitHub URLs automatically.

2. rtf_image_embedder.py — Standalone module that finds [Image: ...] placeholders in RTF, reads the referenced local images, downscales them to fit page bounds using Pillow (https://pillow.readthedocs.io/), and embeds them as hex-encoded \pict blocks.

### Dependencies

• Python 3 (preinstalled on GitHub Actions runners)
• Pygments (pip install pygments) — syntax highlighting for 500+ languages
• Pillow (pip install pillow) — image processing for embedding and downscaling

### Run Locally

    pip install pygments pillow
    python3 gfm_markdown_to_rtf.py README.md README.rtf
    python3 rtf_image_embedder.py README.rtf

Viewing Tips
------------

When opening the RTF in LibreOffice Writer:

• Disable red squiggles: Tools > Automatic Spell Checking (toggle off)
• Read-only mode: Edit > Edit Mode (toggle off) for a clean reading view
• Full screen: Ctrl+Shift+J

────────────────────────────────────────

:test_tube: Test Content
========================

  Everything below exercises every GitHub-Flavored Markdown element supported by the converter. The generated README.rtf in this repo is living proof it all works — open it and compare!

Heading Level 2
---------------

### Heading Level 3

#### Heading Level 4

##### Heading Level 5

###### Heading Level 6

Text Formatting
---------------

This is bold text and this is also bold. This is italic text and this is also italic. This is bold and italic together. This is strikethrough text. This is subscript and this is superscript. This is underlined text.

Escaped special characters: not italic and #not a heading.

Line break with two trailing spaces:   This is on the next line.

Line break with backslash:\ This is also on the next line.

And a 
hard break via HTML.

Links
-----

Here is an inline link (https://example.com) and a link to a section.

Here is a relative link to another file (../SKILL.md).

Bare URL autolink: https://github.com

GitHub references: @octocat (https://github.com/octocat) and issue #42 (https://github.com/RandyHaylor/gh-readme2rtf-docx-txt/issues/42).

Images
------

Standard markdown image (no size control):

[Image: Alt text for an image]

HTML image with custom width:

[Image: Image at 300px wide]

HTML image with custom height only (width should auto-calculate from aspect ratio):

[Image: Image at 100px tall]

HTML image with custom width and height:

[Image: Image at 150x75]

HTML image with percentage width:

[Image: Image at 50% width]

Picture element for light/dark theme variants:



Large wide image (should be clamped to page width):

[Image: Large wide image]

Large tall image (should be clamped to page height):

[Image: Large tall image]

Blockquotes
-----------

  This is a blockquote.
  It can span multiple lines.

  > This is a nested blockquote with actual content inside it.

Alerts
------

[NOTE]
  This is a note alert with helpful information.

[TIP]
  This is a tip alert with a suggestion.

[IMPORTANT]
  This is an important alert.

[WARNING]
  This is a warning alert.

[CAUTION]
  This is a caution alert.

Unordered Lists
---------------

• First item
• Second item
  • Nested item A
  • Nested item B
    • Deeply nested item
• Third item

Alternative markers:

• Star item one
• Star item two

• Plus item one
• Plus item two

Ordered Lists
-------------

1. First step
2. Second step
  1. Sub-step A
  2. Sub-step B
3. Third step

Task Lists
----------

[x] Completed task
[ ] Incomplete task
[x] Another done task
  [ ] Nested incomplete subtask
  [x] Nested complete subtask

Inline Code
-----------

Use the println() function to print output. The config file is at ~/.config/app.json.

Fenced Code Blocks
------------------

    fun main() {
        val greeting = "Hello, World!"
        println(greeting)
    }

    <uses-permission android:name="android.permission.INTERNET" />

    #!/bin/bash
    echo "No language-specific highlighting"

    Plain code block with no language specified.

Tables
------

Feature  Status     Notes
-------  ---------  ---------------------
Bold     Supported  Uses text
Italic   Supported  Uses text
Tables   Supported  You're looking at one

Left, center, and right alignment:

Left  Center  Right
----  ------  -----
L1    C1      R1
L2    C2      R2

Horizontal Rules
----------------

────────────────────────────────────────

────────────────────────────────────────

────────────────────────────────────────

Footnotes
---------

Here is a sentence with a footnote reference[1]. And another[longnote].


Emoji
-----

Emoji shortcodes: 🚀 🚀, 🎉 🎉, ⚠️ ⚠️

Color Codes
-----------

GitHub renders color swatches for these inline code formats:

#FF5733 #28A745 rgb(54, 95, 145) hsl(210, 45%, 39%)

HTML Comments
-------------



Visible text after a hidden comment.

Mixed Nesting
-------------

1. Ordered item with bold and code
  • Unordered sub-item with a link (https://example.com)
  • Another sub-item with italic
     > A blockquote inside a list
2. Second ordered item
   ``python    # Code block inside a list    print("hello")    ``
3. Third item with strikethrough


────────────────────────────────────────
[1] This is the first footnote.
[longnote] This is a longer footnote with multiple words.