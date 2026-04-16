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
_INLINE_CODE_STASH = {}  # Stash inline code content to protect from further rules
_INLINE_CODE_PREFIX = '\x00CODE:'

def _xml_escape(text):
    """Escape XML special characters."""
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

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
        elif ord(char) > 0xFFFF:
            # Surrogate pair for codepoints above U+FFFF
            cp = ord(char) - 0x10000
            high = 0xD800 + (cp >> 10)
            low = 0xDC00 + (cp & 0x3FF)
            # RTF \u uses signed 16-bit
            import struct
            high_signed = struct.unpack('h', struct.pack('H', high))[0]
            low_signed = struct.unpack('h', struct.pack('H', low))[0]
            result.append(f'\\u{high_signed}?\\u{low_signed}?')
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
# DOCX-SPECIFIC HANDLERS
# ---------------------------------------------------------------------------
# DOCX hyperlinks need relationship IDs. We collect them during inline processing
# and resolve them when assembling the final document.xml.rels.
_DOCX_HYPERLINK_COUNTER = 0
_DOCX_HYPERLINKS = {}  # {rId: url}

def _docx_register_hyperlink(url):
    """Register a hyperlink URL and return its relationship ID."""
    global _DOCX_HYPERLINK_COUNTER
    _DOCX_HYPERLINK_COUNTER += 1
    rid = f'rIdLink{_DOCX_HYPERLINK_COUNTER}'
    _DOCX_HYPERLINKS[rid] = url
    return rid

def _docx_handle_md_link(match):
    """Convert [text](url) to DOCX hyperlink XML."""
    text = match.group(1)
    url = match.group(2)
    if url.startswith('#'):
        bookmark = url[1:]
        return f'<w:hyperlink w:anchor="{bookmark}"><w:r><w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:hyperlink>'
    rid = _docx_register_hyperlink(url)
    return f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:hyperlink>'

def _docx_handle_bare_url(match):
    """Convert bare URL to DOCX hyperlink XML."""
    url = match.group(0)
    rid = _docx_register_hyperlink(url)
    return f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">{url}</w:t></w:r></w:hyperlink>'

def _docx_handle_mention(match):
    """Convert @username to DOCX hyperlink to GitHub profile."""
    username = match.group(1)
    url = f'https://github.com/{username}'
    rid = _docx_register_hyperlink(url)
    return f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">@{username}</w:t></w:r></w:hyperlink>'

def _docx_handle_issue_ref(match):
    """Convert #42 to DOCX hyperlink if repo context available."""
    number = match.group(1)
    repo = _detect_github_repo()
    if repo:
        url = f'https://github.com/{repo}/issues/{number}'
        rid = _docx_register_hyperlink(url)
        return f'<w:hyperlink r:id="{rid}"><w:r><w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">#{number}</w:t></w:r></w:hyperlink>'
    return f'<w:r><w:rPr><w:color w:val="365F91"/></w:rPr><w:t>#{number}</w:t></w:r>'

def _docx_handle_html_img(match):
    """Convert <img> tag to DOCX image placeholder text."""
    tag = match.group(0)
    alt_match = re.search(r'alt="([^"]*)"', tag)
    src_match = re.search(r'src="([^"]*)"', tag)
    alt_text = alt_match.group(1) if alt_match else 'image'
    src_text = src_match.group(1) if src_match else ''
    return f'<w:r><w:rPr><w:color w:val="666666"/></w:rPr><w:t>[Image: {alt_text} — {src_text}]</w:t></w:r>'

def _docx_handle_emoji(match):
    """Convert :shortcode: to emoji character for DOCX."""
    shortcode = match.group(0)
    emoji_char = EMOJI_MAP.get(shortcode)
    if emoji_char:
        return f'<w:r><w:t>{emoji_char}</w:t></w:r>'
    return shortcode

def _stash_inline_code(content):
    """Stash inline code content behind a placeholder so later rules can't touch it."""
    key = f'{_INLINE_CODE_PREFIX}{len(_INLINE_CODE_STASH)}'
    _INLINE_CODE_STASH[key] = content
    return key

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

# Active output format — controls which replacement is used from multi-format rules
_ACTIVE_FORMAT = 'rtf'

INLINE_RULES = [
    # --- Phase 1: Strip/transform HTML and structural elements ---
    # These are format-neutral (same for all outputs)
    ('html_comment',    (r'<!--.*?-->',                         {'rtf': '', 'docx': '', 'pdf': ''}, re.DOTALL)),
    ('html_picture',    (r'<picture>.*?<img\s+([^>]*)>.*?</picture>', {'rtf': r'<img \1>', 'docx': r'<img \1>'}, re.DOTALL)),
    ('html_source',     (r'<source[^>]*>',                      {'rtf': '', 'docx': ''})),
    ('html_img',        (r'<img\s+[^>]*>',                      {'rtf': _handle_html_img, 'docx': _docx_handle_html_img})),
    ('html_sub',        (r'<sub>(.*?)</sub>',                    {'rtf': r'{\\sub \1}',
                                                                    'docx': r'<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>\1</w:t></w:r>'})),
    ('html_sup',        (r'<sup>(.*?)</sup>',                    {'rtf': r'{\\super \1}',
                                                                    'docx': r'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>\1</w:t></w:r>'})),
    ('html_ins',        (r'<ins>(.*?)</ins>',                    {'rtf': r'{\\ul \1}',
                                                                    'docx': r'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>\1</w:t></w:r>'})),
    ('html_br',         (r'<br\s*/?>',                           {'rtf': r'\\line ',
                                                                    'docx': r'<w:r><w:br/></w:r>'})),

    # --- Phase 2: Escaped chars -> placeholders (before any markdown matching) ---
    ('escaped_char',    (r'\\([*#_~`\[\]\\])',                   {'rtf': _handle_escaped_char, 'docx': _handle_escaped_char})),

    # --- Phase 3: Images and links (before inline code, which uses backticks) ---
    ('md_image',        (r'!\[([^\]]*)\]\(([^)]+)\)',            {'rtf': r'{\\cf3 [Image: \1 \\u8212? \2]}',
                                                                    'docx': r'<w:r><w:rPr><w:color w:val="666666"/></w:rPr><w:t>[Image: \1 — \2]</w:t></w:r>'})),
    ('md_link',         (r'\[([^\]]+)\]\(([^)]+)\)',             {'rtf': _handle_md_link, 'docx': _docx_handle_md_link})),
    ('bare_url',        (r'(?<!["\(])https?://[^\s<>\)]+',       {'rtf': _handle_bare_url, 'docx': _docx_handle_bare_url})),

    # --- Phase 4: Inline code (stash content to protect from emoji/mention rules) ---
    ('inline_code',     (r'`([^`]+)`',                           {'rtf': lambda m: _stash_inline_code(m.group(1)),
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/><w:shd w:val="clear" w:fill="E6F0FA"/></w:rPr><w:t xml:space="preserve">{m.group(1)}</w:t></w:r>'})),

    # --- Phase 5: GitHub-specific inline elements ---
    ('mention',         (r'@(\w[\w/-]*)',                         {'rtf': _handle_mention, 'docx': _docx_handle_mention})),
    ('issue_ref',       (r'(?<![&A-Fa-f0-9])#(\d+)\b',          {'rtf': _handle_issue_ref, 'docx': _docx_handle_issue_ref})),
    ('footnote_ref',    (r'\[\^([^\]]+)\]',                      {'rtf': lambda m: f'{{\\super {{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK \\\\l "fn-{m.group(1)}"}}{{\\fldrslt \\cf2  [{m.group(1)}] }}}}}}',
                                                                    'docx': lambda m: f'<w:hyperlink w:anchor="fn-{m.group(1)}"><w:r><w:rPr><w:vertAlign w:val="superscript"/><w:color w:val="365F91"/></w:rPr><w:t>[{m.group(1)}]</w:t></w:r></w:hyperlink>'})),
    ('emoji',           (r':\w+:',                                {'rtf': _handle_emoji, 'docx': _docx_handle_emoji})),

    # --- Phase 6: Text formatting ---
    ('bold_italic',     (r'\*\*\*(.+?)\*\*\*',                   {'rtf': r'{\\b\\i \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
    ('bold_star',       (r'\*\*(.+?)\*\*',                        {'rtf': r'{\\b \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
    ('bold_under',      (r'__(.+?)__',                             {'rtf': r'{\\b \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
    ('italic_star',     (r'\*(.+?)\*',                             {'rtf': r'{\\i \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
    ('italic_under',    (r'(?<!\w)_(.+?)_(?!\w)',                  {'rtf': r'{\\i \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
    ('strikethrough',   (r'~~(.+?)~~',                             {'rtf': r'{\\strike \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">{_xml_escape(m.group(1))}</w:t></w:r>'})),
]


def apply_inline_rules(text, fmt=None):
    """Apply all inline conversion rules in order for the given format."""
    target_fmt = fmt or _ACTIVE_FORMAT
    for rule_name, rule_def in INLINE_RULES:
        # rule_def is (pattern, format_dict) or (pattern, format_dict, flags)
        pattern = rule_def[0]
        format_dict = rule_def[1]
        flags = rule_def[2] if len(rule_def) == 3 and isinstance(rule_def[2], int) else 0

        # Get replacement for the target format, fall back to 'rtf'
        replacement = format_dict.get(target_fmt, format_dict.get('rtf'))
        if replacement is None:
            continue

        if callable(replacement):
            text = re.sub(pattern, replacement, text, flags=flags)
        else:
            text = re.sub(pattern, replacement, text, flags=flags)

    # Final phase: restore placeholders
    text = text.replace(_ESCAPE_PLACEHOLDER, '')
    text = text.replace(_RTF_STAR_PLACEHOLDER, '\\*')

    # Restore stashed inline code
    for key, content in _INLINE_CODE_STASH.items():
        text = text.replace(key, f'{{\\f1\\fs20\\chshdng1\\chcbpat8 {content}}}')
    _INLINE_CODE_STASH.clear()

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


# ---------------------------------------------------------------------------
# DOCX BLOCK RULES
# ---------------------------------------------------------------------------

def docx_block_blank_line(lines, index):
    if lines[index].strip() == '':
        return ('', 1)
    return None

def docx_block_horizontal_rule(lines, index):
    if re.match(r'^(\s*[-*_]\s*){3,}$', lines[index]):
        return ('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>', 1)
    return None

def docx_block_heading(lines, index):
    heading_match = re.match(r'^(#{1,6})\s+(.*)', lines[index])
    if heading_match:
        level = len(heading_match.group(1))
        raw_text = heading_match.group(2)
        bookmark_id = _heading_to_bookmark_id(raw_text)
        bid = hash(bookmark_id) % 10000
        # Heading text — escape for XML, no inline formatting needed for headings
        escaped = raw_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        xml = (f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr>'
               f'<w:bookmarkStart w:id="{bid}" w:name="{bookmark_id}"/>'
               f'<w:bookmarkEnd w:id="{bid}"/>'
               f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>')
        return (xml, 1)
    return None

def docx_block_html_picture(lines, index):
    if not re.match(r'^\s*<picture', lines[index], re.IGNORECASE):
        return None
    collected = [lines[index]]
    consumed = 1
    while index + consumed < len(lines) and '</picture>' not in collected[-1].lower():
        collected.append(lines[index + consumed])
        consumed += 1
    joined = ' '.join(l.strip() for l in collected)
    formatted = apply_inline_rules(joined, fmt='docx')
    return (f'<w:p>{formatted}</w:p>', consumed)

def docx_block_html_img(lines, index):
    if not re.match(r'^\s*<img\s', lines[index], re.IGNORECASE):
        return None
    formatted = apply_inline_rules(lines[index].strip(), fmt='docx')
    return (f'<w:p>{formatted}</w:p>', 1)

def docx_block_fenced_code(lines, index):
    fence_match = re.match(r'^(\s*)```(\w*)', lines[index])
    if not fence_match:
        return None
    indent = fence_match.group(1)
    consumed = 1
    code_lines = []
    while index + consumed < len(lines):
        current_line = lines[index + consumed]
        if current_line.strip().startswith('```'):
            consumed += 1
            break
        if indent and current_line.startswith(indent):
            current_line = current_line[len(indent):]
        code_lines.append(current_line)
        consumed += 1
    # Each line is a paragraph with monospace font and shaded background
    parts = []
    for line in code_lines:
        escaped = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        parts.append(
            f'<w:p><w:pPr><w:shd w:val="clear" w:fill="E6F0FA"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr>'
            f'<w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'
        )
    return ('\n'.join(parts), consumed)

def docx_block_table(lines, index):
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
    num_cols = len(header_row)
    col_width = 9000 // num_cols

    parts = ['<w:tbl><w:tblPr><w:tblBorders>'
             '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
             '</w:tblBorders></w:tblPr>']

    # Header row
    parts.append('<w:tr>')
    for cell in header_row:
        escaped = cell.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        parts.append(f'<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                     f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p></w:tc>')
    parts.append('</w:tr>')

    for row in data_rows:
        parts.append('<w:tr>')
        for ci in range(num_cols):
            cell = row[ci] if ci < len(row) else ''
            formatted = apply_inline_rules(cell, fmt='docx')
            wrapped = _docx_wrap_plain_text_in_runs(formatted)
            parts.append(f'<w:tc><w:p>{wrapped}</w:p></w:tc>')
        parts.append('</w:tr>')

    parts.append('</w:tbl>')
    return ('\n'.join(parts), consumed)

def docx_block_blockquote(lines, index):
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
            next_idx = index + consumed + 1
            if next_idx < len(lines) and lines[next_idx].startswith('>'):
                next_stripped = re.sub(r'^>\s?', '', lines[next_idx])
                if next_stripped.strip() in ALERT_TYPES:
                    break
                quote_lines.append('')
                consumed += 1
            else:
                break
        else:
            break

    # Check for alert
    is_alert = False
    alert_label = ''
    if quote_lines and quote_lines[0].strip() in ALERT_TYPES:
        alert_label, _, _ = ALERT_TYPES[quote_lines[0].strip()]
        quote_lines = quote_lines[1:]
        is_alert = True

    parts = []
    for ql in quote_lines:
        text = apply_inline_rules(ql, fmt='docx') if ql.strip() else ''
        wrapped = _docx_wrap_plain_text_in_runs(text) if text else ''
        if is_alert:
            parts.append(f'<w:p><w:pPr><w:ind w:left="720"/><w:shd w:val="clear" w:fill="E6F0FA"/></w:pPr>'
                         f'{wrapped}</w:p>')
        else:
            parts.append(f'<w:p><w:pPr><w:ind w:left="720"/><w:pBdr>'
                         f'<w:left w:val="single" w:sz="12" w:space="4" w:color="CCCCCC"/>'
                         f'</w:pBdr></w:pPr>{wrapped}</w:p>')
    if is_alert:
        alert_para = (f'<w:p><w:pPr><w:ind w:left="720"/><w:shd w:val="clear" w:fill="E6F0FA"/></w:pPr>'
                      f'<w:r><w:rPr><w:b/><w:color w:val="365F91"/></w:rPr>'
                      f'<w:t>{alert_label}</w:t></w:r></w:p>')
        parts.insert(0, alert_para)

    return ('\n'.join(parts), consumed)

def docx_block_list(lines, index):
    if not detect_list_item(lines[index]):
        return None
    parts = []
    consumed = 0
    while index + consumed < len(lines):
        list_info = detect_list_item(lines[index + consumed])
        if not list_info:
            break
        nest_level, marker_type, content = list_info
        is_task, is_checked, task_content = detect_task_checkbox(content)
        display_content = task_content if is_task else content
        formatted = apply_inline_rules(display_content, fmt='docx')
        indent = 720 + (nest_level * 360)

        if is_task:
            checkbox = '\u2611' if is_checked else '\u2610'
            prefix = f'{checkbox} '
        elif marker_type == 'unordered':
            prefix = '\u2022 '
        else:
            number = marker_type.split(':')[1]
            prefix = f'{number}. '

        wrapped = _docx_wrap_plain_text_in_runs(formatted)
        parts.append(
            f'<w:p><w:pPr><w:ind w:left="{indent}"/><w:spacing w:after="40"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{prefix}</w:t></w:r>{wrapped}</w:p>'
        )
        consumed += 1
    return ('\n'.join(parts), consumed)

def docx_block_footnote_def(lines, index):
    match = re.match(r'^\[\^([^\]]+)\]:\s*(.*)', lines[index])
    if not match:
        return None
    fid = match.group(1)
    ftext = match.group(2)
    _collected_footnotes.append((fid, ftext))
    return ('', 1)

def _docx_wrap_plain_text_in_runs(text):
    """Wrap any plain text segments (not already in <w:r> tags) into <w:r><w:t> runs."""
    # Split on existing <w:r> and <w:hyperlink> elements, wrap the gaps
    parts = re.split(r'(<w:(?:r|hyperlink|bookmarkStart|bookmarkEnd)[^>]*>.*?</w:(?:r|hyperlink)>|<w:(?:r|bookmarkStart|bookmarkEnd)\s*/>)', text, flags=re.DOTALL)
    result = []
    for part in parts:
        if not part or part.startswith('<w:'):
            result.append(part or '')
        else:
            # Plain text — wrap in a run
            escaped = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            if escaped.strip():
                result.append(f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>')
            elif escaped:
                result.append(f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>')
    return ''.join(result)

def docx_block_paragraph(lines, index):
    collected = []
    consumed = 0
    while index + consumed < len(lines):
        current = lines[index + consumed]
        if current.strip() == '':
            break
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
        collected.append(current)
        consumed += 1
    if not collected:
        return None
    paragraph_text = ' '.join(collected)
    formatted = apply_inline_rules(paragraph_text, fmt='docx')
    wrapped = _docx_wrap_plain_text_in_runs(formatted)
    return (f'<w:p>{wrapped}</w:p>', consumed)


# Ordered list of block rules — first match wins
BLOCK_RULES = {
    'rtf': [
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
    ],
    'docx': [
        docx_block_blank_line,
        docx_block_horizontal_rule,
        docx_block_heading,
        docx_block_html_picture,
        docx_block_html_img,
        docx_block_fenced_code,
        docx_block_table,
        docx_block_blockquote,
        docx_block_list,
        docx_block_footnote_def,
        docx_block_paragraph,
    ],
}


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
        for block_rule in BLOCK_RULES['rtf']:
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
# DOCX CONVERSION ENGINE
# ---------------------------------------------------------------------------

DOCX_STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
    <w:pPr><w:keepNext/><w:spacing w:before="480" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/><w:color w:val="365F91"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
    <w:pPr><w:keepNext/><w:spacing w:before="280" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="365F91"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="30"/><w:color w:val="365F91"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/>
    <w:pPr><w:keepNext/><w:spacing w:before="200" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="365F91"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="365F91"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/>
    <w:pPr><w:keepNext/><w:spacing w:before="120" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="365F91"/></w:rPr></w:style>
</w:styles>'''


def _build_docx_footnotes_section():
    """Build footnotes section as DOCX XML paragraphs."""
    if not _collected_footnotes:
        return ''
    parts = []
    parts.append('<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>')
    parts.append('<w:p><w:pPr><w:spacing w:before="120"/></w:pPr>'
                 '<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>Footnotes</w:t></w:r></w:p>')
    for fid, ftext in _collected_footnotes:
        formatted = apply_inline_rules(ftext, fmt='docx')
        bookmark = f'fn-{fid}'
        bid = hash(bookmark) % 10000
        parts.append(
            f'<w:p><w:pPr><w:ind w:left="360"/><w:spacing w:after="40"/></w:pPr>'
            f'<w:bookmarkStart w:id="{bid}" w:name="{bookmark}"/>'
            f'<w:bookmarkEnd w:id="{bid}"/>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>{fid}.</w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> {formatted}</w:t></w:r></w:p>'
        )
    return '\n'.join(parts)


def convert_markdown_to_docx(markdown_text, output_path):
    """Convert a GFM markdown string to a DOCX file."""
    import zipfile
    global _DOCX_HYPERLINK_COUNTER
    _DOCX_HYPERLINK_COUNTER = 0
    _DOCX_HYPERLINKS.clear()
    _collected_footnotes.clear()

    lines = markdown_text.split('\n')
    body_parts = []
    line_index = 0

    while line_index < len(lines):
        matched = False
        for block_rule in BLOCK_RULES['docx']:
            result = block_rule(lines, line_index)
            if result is not None:
                xml_content, lines_consumed = result
                if xml_content:
                    body_parts.append(xml_content)
                line_index += lines_consumed
                matched = True
                break
        if not matched:
            line_index += 1

    # Append footnotes
    footnotes_xml = _build_docx_footnotes_section()
    if footnotes_xml:
        body_parts.append(footnotes_xml)

    # Build document.xml
    body_content = '\n'.join(body_parts)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body>'
        f'{body_content}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        '</w:sectPr></w:body></w:document>'
    )

    # Build document.xml.rels (styles + hyperlinks)
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
    for rid, url in _DOCX_HYPERLINKS.items():
        rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{url}" TargetMode="External"/>')
    rels.append('</Relationships>')
    doc_rels_xml = '\n'.join(rels)

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', root_rels)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/_rels/document.xml.rels', doc_rels_xml)
        zf.writestr('word/styles.xml', DOCX_STYLES_XML)


# ---------------------------------------------------------------------------
# CLI ENTRY POINT
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f'Usage: {sys.argv[0]} <input.md> [output.rtf|output.docx]', file=sys.stderr)
        sys.exit(1)

    input_markdown_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_markdown_path.rsplit('.', 1)[0] + '.rtf'

    with open(input_markdown_path, 'r', encoding='utf-8') as markdown_file:
        markdown_content = markdown_file.read()

    if output_path.endswith('.docx'):
        convert_markdown_to_docx(markdown_content, output_path)
    else:
        rtf_output = convert_markdown_to_rtf(markdown_content)
        with open(output_path, 'w', encoding='utf-8') as rtf_file:
            rtf_file.write(rtf_output)

    print(f'Converted: {input_markdown_path} -> {output_path}')
