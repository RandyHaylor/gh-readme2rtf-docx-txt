#!/usr/bin/env python3
"""
gfm_markdown_to_rtf.py — Convert GitHub-Flavored Markdown to RTF.
Zero external dependencies. Designed to be embeddable inline in a GitHub Actions YAML.

Usage: python3 gfm_markdown_to_rtf.py <input.md> [output.rtf]

Architecture: Data-driven conversion engine. Inline and block rules are defined
as ordered lists of (pattern, handler) tuples. The engine iterates the rules
in order; first match wins. This makes adding/fixing conversions a matter of
editing the rule lists, not the engine logic.
"""
import re
import sys

try:
    from pygments import lex
    from pygments.lexers import get_lexer_by_name, TextLexer
    from pygments.token import Token
    PYGMENTS_AVAILABLE = True
except ImportError:
    PYGMENTS_AVAILABLE = False

# ---------------------------------------------------------------------------
# RTF DOCUMENT TEMPLATE
# ---------------------------------------------------------------------------
RTF_HEADER = r"""{\rtf1\ansi\deff0
{\fonttbl
{\f0\fswiss\fcharset0 Calibri;}
{\f1\fmodern\fcharset0 Consolas;}
}
{\colortbl;
\red0\green0\blue0;
\red54\green95\blue145;
\red102\green102\blue102;
\red136\green0\blue0;
\red0\green112\blue32;
\red186\green33\blue33;
\red188\green122\blue0;
\red230\green240\blue250;
\red255\green243\blue205;
\red255\green230\blue230;
\red230\green255\blue230;
\red255\green245\blue230;
\red31\green111\blue200;
\red26\green127\blue55;
\red130\green80\blue223;
\red191\green135\blue0;
\red207\green34\blue46;
\red0\green0\blue128;
\red163\green21\blue21;
\red0\green128\blue0;
\red43\green145\blue175;
\red128\green0\blue128;
\red128\green128\blue128;
\red0\green0\blue255;
}
"""
RTF_FOOTER = r"}"

# Color index reference:
# 1=black 2=blue 3=gray 4=dark-red 5=green 6=red 7=amber
# 8=light-blue-bg 9=light-yellow-bg 10=light-red-bg 11=light-green-bg 12=light-orange-bg
# 13=note-bar-blue 14=tip-bar-green 15=important-bar-purple 16=warning-bar-amber 17=caution-bar-red
# 18=syntax-keyword(navy) 19=syntax-string(brown) 20=syntax-comment(green)
# 21=syntax-type(teal) 22=syntax-builtin(purple) 23=syntax-punctuation(gray) 24=syntax-number(blue)

# Pygments token type -> RTF color index
SYNTAX_COLOR_MAP = {
    Token.Keyword:              18,
    Token.Keyword.Constant:     18,
    Token.Keyword.Declaration:  18,
    Token.Keyword.Namespace:    18,
    Token.Keyword.Type:         21,
    Token.Name.Builtin:         22,
    Token.Name.Class:           21,
    Token.Name.Function:        1,
    Token.Name.Decorator:       22,
    Token.Name.Tag:             18,
    Token.Name.Attribute:       21,
    Token.Literal.String:       19,
    Token.Literal.String.Doc:   19,
    Token.Literal.String.Single: 19,
    Token.Literal.String.Double: 19,
    Token.Literal.Number:       24,
    Token.Literal.Number.Integer: 24,
    Token.Literal.Number.Float: 24,
    Token.Comment:              20,
    Token.Comment.Single:       20,
    Token.Comment.Multiline:    20,
    Token.Comment.Preproc:      22,
    Token.Operator:             1,
    Token.Punctuation:          23,
} if PYGMENTS_AVAILABLE else {}

# ---------------------------------------------------------------------------
# RTF ESCAPE
# ---------------------------------------------------------------------------
# Placeholder for escaped markdown chars — must not collide with real text
_ESCAPE_PLACEHOLDER = '\x00ESC:'
_RTF_STAR_PLACEHOLDER = '\x00RTFSTAR'  # Protects \* in RTF fields from italic regex

def rtf_escape(text):
    """Escape special RTF characters and handle unicode."""
    result = []
    for char in text:
        if char == '\\':
            result.append('\\\\')
        elif char == '{':
            result.append('\\{')
        elif char == '}':
            result.append('\\}')
        elif char == '\t':
            result.append('\\tab ')
        elif ord(char) > 127:
            result.append(f'\\u{ord(char)}?')
        else:
            result.append(char)
    return ''.join(result)

# ---------------------------------------------------------------------------
# EMOJI MAP (common shortcodes -> unicode codepoints)
# ---------------------------------------------------------------------------
EMOJI_MAP = {
    ':rocket:': '\U0001F680', ':tada:': '\U0001F389', ':warning:': '\u26A0\uFE0F',
    ':star:': '\u2B50', ':fire:': '\U0001F525', ':bug:': '\U0001F41B',
    ':check:': '\u2705', ':x:': '\u274C', ':heart:': '\u2764\uFE0F',
    ':thumbsup:': '\U0001F44D', ':thumbsdown:': '\U0001F44E', ':eyes:': '\U0001F440',
    ':memo:': '\U0001F4DD', ':bulb:': '\U0001F4A1', ':gear:': '\u2699\uFE0F',
    ':lock:': '\U0001F512', ':key:': '\U0001F511', ':boom:': '\U0001F4A5',
    ':construction:': '\U0001F6A7', ':sparkles:': '\u2728', ':zap:': '\u26A1',
    ':white_check_mark:': '\u2705', ':heavy_check_mark:': '\u2714\uFE0F',
    ':arrow_right:': '\u27A1\uFE0F', ':point_right:': '\U0001F449',
    ':100:': '\U0001F4AF', ':wave:': '\U0001F44B', ':pray:': '\U0001F64F',
}

# ---------------------------------------------------------------------------
# INLINE FORMATTING ENGINE
# ---------------------------------------------------------------------------
# Each rule: (name, pattern, replacement_or_handler)
# Rules are applied in order. First-match-wins for overlapping patterns.
# Use a callable for complex replacements; string for simple re.sub.
#
# IMPORTANT: Order matters! Rules that produce RTF markup (like links) must
# come before rules that could match inside that markup (like bold/italic).
# Escaped chars must use placeholders to avoid re-matching.

def _handle_html_img(match):
    """Convert <img> tag to RTF text placeholder with dimensions."""
    tag = match.group(0)
    src_match = re.search(r'src="([^"]*)"', tag)
    alt_match = re.search(r'alt="([^"]*)"', tag)
    width_match = re.search(r'width="([^"]*)"', tag)
    height_match = re.search(r'height="([^"]*)"', tag)
    alt_text = alt_match.group(1) if alt_match else 'image'
    src_text = src_match.group(1) if src_match else ''
    size_parts = []
    if width_match:
        size_parts.append(f'w:{width_match.group(1)}')
    if height_match:
        size_parts.append(f'h:{height_match.group(1)}')
    size_info = f' ({", ".join(size_parts)})' if size_parts else ''
    return f'{{\\cf3 [Image: {rtf_escape(alt_text)}{size_info} \\u8212? {rtf_escape(src_text)}]}}'

def _handle_bare_url(match):
    """Convert bare URL to RTF hyperlink field."""
    url = match.group(0)
    return f'{{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK "{url}"}}{{\\fldrslt \\cf2 {url}}}}}'

def _handle_md_link(match):
    """Convert [text](url) to RTF hyperlink field. Internal #anchors use \\l flag."""
    text = match.group(1)
    url = match.group(2)
    if url.startswith('#'):
        bookmark = url[1:]  # strip the #
        return f'{{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK \\\\l "{bookmark}"}}{{\\fldrslt \\cf2 {text}}}}}'
    return f'{{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK "{url}"}}{{\\fldrslt \\cf2 {text}}}}}'

def _handle_emoji(match):
    """Convert :shortcode: to unicode emoji character."""
    shortcode = match.group(0)
    emoji_char = EMOJI_MAP.get(shortcode)
    if emoji_char:
        return rtf_escape(emoji_char)
    return shortcode

def _handle_escaped_char(match):
    """Replace \\* with placeholder so it doesn't match bold/italic later."""
    return _ESCAPE_PLACEHOLDER + match.group(1)

# ---------------------------------------------------------------------------
# GITHUB REPO CONTEXT (for resolving @mentions and #issues to URLs)
# ---------------------------------------------------------------------------
_GITHUB_REPO_SLUG = None  # e.g. "RandyHaylor/gfm-to-rtf-test"

def _detect_github_repo():
    """Try to detect owner/repo from git remote. Returns 'owner/repo' or None."""
    global _GITHUB_REPO_SLUG
    if _GITHUB_REPO_SLUG is not None:
        return _GITHUB_REPO_SLUG if _GITHUB_REPO_SLUG else None
    import subprocess
    try:
        result = subprocess.run(['git', 'remote', 'get-url', 'origin'],
                                capture_output=True, text=True, timeout=5)
        url = result.stdout.strip()
        # Handle SSH: git@github.com:owner/repo.git
        ssh_match = re.search(r'github\.com[:/](.+?)(?:\.git)?$', url)
        if ssh_match:
            _GITHUB_REPO_SLUG = ssh_match.group(1)
            return _GITHUB_REPO_SLUG
        # Handle HTTPS: https://github.com/owner/repo.git
        https_match = re.search(r'github\.com/(.+?)(?:\.git)?$', url)
        if https_match:
            _GITHUB_REPO_SLUG = https_match.group(1)
            return _GITHUB_REPO_SLUG
    except Exception:
        pass
    _GITHUB_REPO_SLUG = ''  # cache the miss
    return None

def _handle_mention(match):
    """Convert @username to a hyperlink to github.com/username."""
    username = match.group(1)
    url = f'https://github.com/{username}'
    return f'{{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK "{url}"}}{{\\fldrslt \\cf2 @{username}}}}}'

def _handle_issue_ref(match):
    """Convert #42 to a hyperlink to the repo's issue, or just colored text if no repo context."""
    number = match.group(1)
    repo = _detect_github_repo()
    if repo:
        url = f'https://github.com/{repo}/issues/{number}'
        return f'{{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK "{url}"}}{{\\fldrslt \\cf2 #{number}}}}}'
    return f'{{\\cf2 #{number}}}'

INLINE_RULES = [
    # --- Phase 1: Strip/transform HTML and structural elements ---
    ('html_comment',    (r'<!--.*?-->',                         '', re.DOTALL)),
    ('html_picture',    (r'<picture>.*?<img\s+([^>]*)>.*?</picture>', r'<img \1>', re.DOTALL)),
    ('html_source',     (r'<source[^>]*>',                      '')),
    ('html_img',        (r'<img\s+[^>]*>',                      _handle_html_img)),
    ('html_sub',        (r'<sub>(.*?)</sub>',                    r'{\\sub \1}')),
    ('html_sup',        (r'<sup>(.*?)</sup>',                    r'{\\super \1}')),
    ('html_ins',        (r'<ins>(.*?)</ins>',                    r'{\\ul \1}')),
    ('html_br',         (r'<br\s*/?>',                           r'\\line ')),

    # --- Phase 2: Escaped chars -> placeholders (before any markdown matching) ---
    ('escaped_char',    (r'\\([*#_~`\[\]\\])',                   _handle_escaped_char)),

    # --- Phase 3: Images and links (before inline code, which uses backticks) ---
    ('md_image',        (r'!\[([^\]]*)\]\(([^)]+)\)',            r'{\\cf3 [Image: \1 \\u8212? \2]}')),
    ('md_link',         (r'\[([^\]]+)\]\(([^)]+)\)',             _handle_md_link)),
    ('bare_url',        (r'(?<!["\(])https?://[^\s<>\)]+',       _handle_bare_url)),

    # --- Phase 4: GitHub-specific inline elements ---
    ('mention',         (r'@(\w[\w/-]*)',                         _handle_mention)),
    ('issue_ref',       (r'(?<![&A-Fa-f0-9])#(\d+)\b',          _handle_issue_ref)),
    ('footnote_ref',    (r'\[\^([^\]]+)\]',                      lambda m: f'{{\\super {{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK \\\\l "fn-{m.group(1)}"}}{{\\fldrslt \\cf2  [{m.group(1)}] }}}}}}')),
    ('emoji',           (r':\w+:',                                _handle_emoji)),

    # --- Phase 5: Inline code (before bold/italic to protect backtick content) ---
    ('inline_code',     (r'`([^`]+)`',                           r'{\\f1\\fs20\\chshdng1\\chcbpat8 \1}')),

    # --- Phase 6: Text formatting ---
    ('bold_italic',     (r'\*\*\*(.+?)\*\*\*',                   r'{\\b\\i \1}')),
    ('bold_star',       (r'\*\*(.+?)\*\*',                        r'{\\b \1}')),
    ('bold_under',      (r'__(.+?)__',                             r'{\\b \1}')),
    ('italic_star',     (r'\*(.+?)\*',                             r'{\\i \1}')),
    ('italic_under',    (r'(?<!\w)_(.+?)_(?!\w)',                  r'{\\i \1}')),
    ('strikethrough',   (r'~~(.+?)~~',                             r'{\\strike \1}')),
]


def apply_inline_rules(text):
    """Apply all inline conversion rules in order."""
    for rule_name, rule_def in INLINE_RULES:
        if callable(rule_def):
            # rule_def is a standalone callable — shouldn't happen with current structure
            text = rule_def(text)
        elif len(rule_def) == 3 and isinstance(rule_def[2], int):
            # (pattern, replacement, flags)
            pattern, replacement, flags = rule_def
            if callable(replacement):
                text = re.sub(pattern, replacement, text, flags=flags)
            else:
                text = re.sub(pattern, replacement, text, flags=flags)
        else:
            # (pattern, replacement)
            pattern, replacement = rule_def[0], rule_def[1]
            if callable(replacement):
                text = re.sub(pattern, replacement, text)
            else:
                text = re.sub(pattern, replacement, text)

    # Final phase: restore placeholders
    text = text.replace(_ESCAPE_PLACEHOLDER, '')
    text = text.replace(_RTF_STAR_PLACEHOLDER, '\\*')

    return text


# ---------------------------------------------------------------------------
# ALERT TYPES
# ---------------------------------------------------------------------------
# Alert types: (label, bar_color_index, label_text_color_index)
# Bar colors: 13=blue, 14=green, 15=purple, 16=amber, 17=red
ALERT_TYPES = {
    '[!NOTE]':      ('Note',      13, 13),
    '[!TIP]':       ('Tip',       14, 14),
    '[!IMPORTANT]': ('Important', 15, 15),
    '[!WARNING]':   ('Warning',   16, 16),
    '[!CAUTION]':   ('Caution',   17, 17),
}

# ---------------------------------------------------------------------------
# BLOCK DETECTION HELPERS
# ---------------------------------------------------------------------------
HEADING_FONT_SIZES = {1: 48, 2: 36, 3: 30, 4: 26, 5: 24, 6: 22}


def detect_list_item(line):
    """Detect if line is a list item. Returns (nest_level, marker_type, content) or None."""
    unordered_match = re.match(r'^(\s*)([-*+])\s+(.*)', line)
    if unordered_match:
        indent = len(unordered_match.group(1))
        return (indent // 2, 'unordered', unordered_match.group(3))

    ordered_match = re.match(r'^(\s*)(\d+)\.\s+(.*)', line)
    if ordered_match:
        indent = len(ordered_match.group(1))
        return (indent // 3, f'ordered:{ordered_match.group(2)}', ordered_match.group(3))

    return None


def detect_task_checkbox(content):
    """Check if list content starts with [ ] or [x]. Returns (is_task, is_checked, text)."""
    task_match = re.match(r'^\[([ xX])\]\s*(.*)', content)
    if task_match:
        return (True, task_match.group(1).lower() == 'x', task_match.group(2))
    return (False, False, content)


# ---------------------------------------------------------------------------
# BLOCK RULES — Each returns (rtf_string, lines_consumed) or None
# ---------------------------------------------------------------------------

def block_blank_line(lines, index):
    if lines[index].strip() == '':
        return ('', 1)
    return None


def block_horizontal_rule(lines, index):
    if re.match(r'^(\s*[-*_]\s*){3,}$', lines[index]):
        rtf = '{\\pard\\sb120\\sa120\\brdrb\\brdrs\\brdrw15\\brsp40\\par}'
        return (rtf, 1)
    return None


def _heading_to_bookmark_id(text):
    """Convert heading text to GitHub-style anchor slug."""
    slug = text.lower().strip()
    slug = re.sub(r'[^\w\s-]', '', slug)
    slug = re.sub(r'\s+', '-', slug)
    return slug


def block_heading(lines, index):
    heading_match = re.match(r'^(#{1,6})\s+(.*)', lines[index])
    if heading_match:
        level = len(heading_match.group(1))
        raw_text = heading_match.group(2)
        bookmark_id = _heading_to_bookmark_id(raw_text)
        text = rtf_escape(raw_text)
        text = apply_inline_rules(text)
        font_size = HEADING_FONT_SIZES.get(level, 24)
        spacing_before = max(360 - (level * 40), 120)
        rtf = (f'{{\\pard\\sb{spacing_before}\\sa120\\keepn\\f0\\fs{font_size}\\cf2\\b '
               f'{{\\*\\bkmkstart {bookmark_id}}}{text}{{\\*\\bkmkend {bookmark_id}}}\\par}}')
        return (rtf, 1)
    return None


def block_html_picture(lines, index):
    if not re.match(r'^\s*<picture', lines[index], re.IGNORECASE):
        return None
    collected = [lines[index]]
    consumed = 1
    while index + consumed < len(lines) and '</picture>' not in collected[-1].lower():
        collected.append(lines[index + consumed])
        consumed += 1
    joined = ' '.join(l.strip() for l in collected)
    formatted = apply_inline_rules(joined)
    rtf = f'{{\\pard\\sb60\\sa60\\f0\\fs22 {formatted}\\par}}'
    return (rtf, consumed)


def block_html_img(lines, index):
    if not re.match(r'^\s*<img\s', lines[index], re.IGNORECASE):
        return None
    formatted = apply_inline_rules(lines[index].strip())
    rtf = f'{{\\pard\\sb60\\sa60\\f0\\fs22 {formatted}\\par}}'
    return (rtf, 1)


def _syntax_highlight_to_rtf(code_text, language):
    """Use Pygments to tokenize code and wrap each token in RTF color commands."""
    if not PYGMENTS_AVAILABLE or not language:
        return rtf_escape(code_text).replace('\n', '\\line\n')

    try:
        lexer = get_lexer_by_name(language, stripall=False)
    except Exception:
        return rtf_escape(code_text).replace('\n', '\\line\n')

    rtf_parts = []
    for token_type, token_value in lex(code_text, lexer):
        escaped = rtf_escape(token_value).replace('\n', '\\line\n')
        # Walk up the token type hierarchy to find a color
        color_index = None
        tt = token_type
        while tt and not color_index:
            color_index = SYNTAX_COLOR_MAP.get(tt)
            tt = tt.parent
        if color_index:
            rtf_parts.append(f'\\cf{color_index} {escaped}\\cf1 ')
        else:
            rtf_parts.append(escaped)
    return ''.join(rtf_parts)


def block_fenced_code(lines, index):
    fence_match = re.match(r'^(\s*)```(\w*)', lines[index])
    if not fence_match:
        return None
    indent = fence_match.group(1)
    language = fence_match.group(2)
    consumed = 1
    code_lines = []
    while index + consumed < len(lines):
        current_line = lines[index + consumed]
        if current_line.strip().startswith('```'):
            consumed += 1
            break
        # Strip leading indent that matches the fence indent
        if indent and current_line.startswith(indent):
            current_line = current_line[len(indent):]
        code_lines.append(current_line)
        consumed += 1
    raw_code = '\n'.join(code_lines)
    highlighted = _syntax_highlight_to_rtf(raw_code, language)
    rtf = f'{{\\pard\\sb100\\sa100\\f1\\fs20\\cbpat8\\li360\\ri360 {highlighted}\\par}}'
    return (rtf, consumed)


def block_table(lines, index):
    line = lines[index]
    if not ('|' in line and line.strip().startswith('|')):
        return None
    next_line = lines[index + 1] if index + 1 < len(lines) else ''
    if not re.match(r'^[\s|:-]+$', next_line):
        return None

    rows = []
    consumed = 0
    while index + consumed < len(lines) and '|' in lines[index + consumed]:
        row_text = lines[index + consumed].strip().strip('|')
        cells = [cell.strip() for cell in row_text.split('|')]
        rows.append(cells)
        consumed += 1

    if len(rows) < 2:
        return None

    header_row = rows[0]
    data_rows = rows[2:]
    num_columns = len(header_row)
    col_width = 9000 // num_columns

    rtf_parts = []

    # Header
    rtf_parts.append('{\\trowd')
    for ci in range(num_columns):
        rtf_parts.append(f'\\clbrdrt\\brdrs\\clbrdrb\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx{col_width * (ci + 1)}')
    rtf_parts.append('\\pard\\intbl')
    for cell in header_row:
        rtf_parts.append(f'{{\\b {apply_inline_rules(cell)}}}\\cell')
    rtf_parts.append('\\row}')

    # Data rows
    for row in data_rows:
        rtf_parts.append('{\\trowd')
        for ci in range(num_columns):
            rtf_parts.append(f'\\clbrdrb\\brdrs\\clbrdrl\\brdrs\\clbrdrr\\brdrs\\cellx{col_width * (ci + 1)}')
        rtf_parts.append('\\pard\\intbl')
        for ci in range(num_columns):
            cell = row[ci] if ci < len(row) else ''
            rtf_parts.append(f'{apply_inline_rules(cell)}\\cell')
        rtf_parts.append('\\row}')

    return ('\n'.join(rtf_parts), consumed)


def block_blockquote(lines, index):
    """Parse a single blockquote block (stops at blank line not followed by >)."""
    if not lines[index].startswith('>'):
        return None

    quote_lines = []
    consumed = 0
    while index + consumed < len(lines):
        current = lines[index + consumed]
        if current.startswith('>'):
            stripped = re.sub(r'^>\s?', '', current)
            quote_lines.append(stripped)
            consumed += 1
        elif current.strip() == '':
            # Blank line: continue only if next > line is NOT a new alert
            next_idx = index + consumed + 1
            if next_idx < len(lines) and lines[next_idx].startswith('>'):
                next_stripped = re.sub(r'^>\s?', '', lines[next_idx])
                if next_stripped.strip() in ALERT_TYPES:
                    break  # New alert block — stop here
                quote_lines.append('')
                consumed += 1
            else:
                break
        else:
            break

    # Check for alert
    alert_label = ''
    alert_bar_color = 0
    alert_text_color = 0
    is_alert = False
    if quote_lines and quote_lines[0].strip() in ALERT_TYPES:
        alert_label, alert_bar_color, alert_text_color = ALERT_TYPES[quote_lines[0].strip()]
        quote_lines = quote_lines[1:]
        is_alert = True

    # Process nested blockquotes
    content_parts = []
    for ql in quote_lines:
        nested_match = re.match(r'^>\s?(.*)', ql)
        if nested_match:
            content_parts.append(f'{{\\li240 {apply_inline_rules(rtf_escape(nested_match.group(1)))}}}')
        else:
            content_parts.append(apply_inline_rules(rtf_escape(ql)))
    joined = '\\line\n'.join(content_parts)

    if is_alert:
        # 2-row table: narrow left column (solid color bar), wide right column (label + content)
        bar_width = 120  # narrow bar in twips
        content_width = 8880  # rest of page
        no_border = '\\clbrdrt\\brdrnone\\clbrdrb\\brdrnone\\clbrdrl\\brdrnone\\clbrdrr\\brdrnone'
        bar_cell = f'{no_border}\\clcbpat{alert_bar_color}\\cellx{bar_width}'
        content_cell = f'{no_border}\\cellx{bar_width + content_width}'
        # Row 1: bar + label
        rtf = (
            f'{{\\trowd\\trbrdrt\\brdrnone\\trbrdrb\\brdrnone\\trbrdrl\\brdrnone\\trbrdrr\\brdrnone\\trgaph80\n'
            f'{bar_cell}\n'
            f'{content_cell}\n'
            f'\\pard\\intbl \\cell\n'
            f'{{\\pard\\intbl\\sb60\\sa0\\f0\\fs24\\b\\cf{alert_text_color} {alert_label}}}\\cell\n'
            f'\\row}}\n'
            # Row 2: bar + content
            f'{{\\trowd\\trbrdrt\\brdrnone\\trbrdrb\\brdrnone\\trbrdrl\\brdrnone\\trbrdrr\\brdrnone\\trgaph80\n'
            f'{bar_cell}\n'
            f'{content_cell}\n'
            f'\\pard\\intbl \\cell\n'
            f'{{\\pard\\intbl\\sb0\\sa60\\f0\\fs22 {joined}}}\\cell\n'
            f'\\row}}'
        )
    else:
        rtf = (f'{{\\pard\\sb60\\sa60\\li480\\brdrl\\brdrs\\brdrw20\\brsp80\\cf3\\f0\\fs22 '
               f'{joined}\\par}}')

    return (rtf, consumed)


def block_list(lines, index):
    if not detect_list_item(lines[index]):
        return None

    rtf_parts = []
    consumed = 0
    while index + consumed < len(lines):
        list_info = detect_list_item(lines[index + consumed])
        if not list_info:
            break
        nest_level, marker_type, content = list_info
        is_task, is_checked, task_content = detect_task_checkbox(content)

        display_content = task_content if is_task else content
        formatted = apply_inline_rules(rtf_escape(display_content))
        left_indent = 480 + (nest_level * 360)

        if is_task:
            bullet = '\\u9745?' if is_checked else '\\u9744?'
            bullet_text = f'{bullet} '
        elif marker_type == 'unordered':
            bullet_text = '\\u8226?  '
        else:
            number = marker_type.split(':')[1]
            bullet_text = f'{number}.  '

        rtf_parts.append(
            f'{{\\pard\\sb36\\sa36\\li{left_indent}\\fi-360\\f0\\fs22 '
            f'{bullet_text}{formatted}\\par}}'
        )
        consumed += 1

    return ('\n'.join(rtf_parts), consumed)


# Collected footnote definitions — populated during parsing, emitted at end of doc
_collected_footnotes = []

def block_footnote_def(lines, index):
    match = re.match(r'^\[\^([^\]]+)\]:\s*(.*)', lines[index])
    if not match:
        return None
    fid = match.group(1)
    ftext = match.group(2)
    _collected_footnotes.append((fid, ftext))
    return ('', 1)  # consume the line but emit nothing here


def block_paragraph(lines, index):
    """Collect contiguous non-blank, non-block lines into a paragraph."""
    collected = []
    consumed = 0
    while index + consumed < len(lines):
        current = lines[index + consumed]
        if current.strip() == '':
            break
        # Check if this line starts a different block type
        if consumed > 0 and (
            re.match(r'^#{1,6}\s', current) or
            re.match(r'^(\s*)```', current) or
            re.match(r'^(\s*[-*_]\s*){3,}$', current) or
            current.startswith('>') or
            re.match(r'^\s*<(picture|img)\s', current, re.IGNORECASE) or
            detect_list_item(current) or
            ('|' in current and current.strip().startswith('|'))
        ):
            break
        # Handle trailing-space line breaks
        if current.endswith('  '):
            collected.append(current.rstrip() + '\\line ')
        elif current.endswith('\\'):
            collected.append(current[:-1] + '\\line ')
        else:
            collected.append(current)
        consumed += 1

    if not collected:
        return None

    paragraph_text = ' '.join(collected)
    paragraph_text = rtf_escape(paragraph_text)
    paragraph_text = apply_inline_rules(paragraph_text)
    rtf = f'{{\\pard\\sb0\\sa120\\f0\\fs22 {paragraph_text}\\par}}'
    return (rtf, consumed)


# Ordered list of block rules — first match wins
BLOCK_RULES = [
    block_blank_line,
    block_horizontal_rule,
    block_heading,
    block_html_picture,
    block_html_img,
    block_fenced_code,
    block_table,
    block_blockquote,
    block_list,
    block_footnote_def,
    block_paragraph,
]


# ---------------------------------------------------------------------------
# MAIN CONVERSION ENGINE
# ---------------------------------------------------------------------------

def _build_footnotes_section():
    """Build the footnotes section RTF from collected definitions."""
    if not _collected_footnotes:
        return ''
    parts = []
    parts.append('{\\pard\\sb360\\sa120\\brdrb\\brdrs\\brdrw10\\brsp20\\par}')
    parts.append('{\\pard\\sb120\\sa60\\f0\\fs20\\b Footnotes\\par}')
    for fid, ftext in _collected_footnotes:
        escaped_fid = rtf_escape(fid)
        escaped_ftext = apply_inline_rules(rtf_escape(ftext))
        bookmark = f'fn-{fid}'
        parts.append(
            f'{{\\pard\\sb30\\sa30\\li360\\f0\\fs18 '
            f'{{\\*\\bkmkstart {bookmark}}}{{\\*\\bkmkend {bookmark}}}'
            f'{{\\b {escaped_fid}.}} {escaped_ftext}\\par}}'
        )
    return '\n'.join(parts)


def convert_markdown_to_rtf(markdown_text):
    """Convert a GFM markdown string to RTF string."""
    _collected_footnotes.clear()
    lines = markdown_text.split('\n')
    rtf_body_parts = []
    line_index = 0

    while line_index < len(lines):
        matched = False
        for block_rule in BLOCK_RULES:
            result = block_rule(lines, line_index)
            if result is not None:
                rtf_content, lines_consumed = result
                if rtf_content:
                    rtf_body_parts.append(rtf_content)
                line_index += lines_consumed
                matched = True
                break
        if not matched:
            line_index += 1  # skip unrecognized line

    # Append footnotes section at end of document
    footnotes_rtf = _build_footnotes_section()
    if footnotes_rtf:
        rtf_body_parts.append(footnotes_rtf)

    return RTF_HEADER + '\n'.join(rtf_body_parts) + '\n' + RTF_FOOTER


# ---------------------------------------------------------------------------
# CLI ENTRY POINT
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f'Usage: {sys.argv[0]} <input.md> [output.rtf]', file=sys.stderr)
        sys.exit(1)

    input_markdown_path = sys.argv[1]
    output_rtf_path = sys.argv[2] if len(sys.argv) > 2 else input_markdown_path.rsplit('.', 1)[0] + '.rtf'

    with open(input_markdown_path, 'r', encoding='utf-8') as markdown_file:
        markdown_content = markdown_file.read()

    rtf_output = convert_markdown_to_rtf(markdown_content)

    with open(output_rtf_path, 'w', encoding='utf-8') as rtf_file:
        rtf_file.write(rtf_output)

    print(f'Converted: {input_markdown_path} -> {output_rtf_path}')
