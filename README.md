# Heading Level 1

Paragraph text with a blank line above and below. This tests basic paragraph rendering.

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

Here is a relative link to [another file](../SKILL.md).

Bare URL autolink: https://github.com

GitHub references: @octocat and issue #42.

## Images

Standard markdown image (no size control):

![Alt text for an image](test-image.png)

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

Emoji shortcodes: :rocket: :tada: :warning:

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
