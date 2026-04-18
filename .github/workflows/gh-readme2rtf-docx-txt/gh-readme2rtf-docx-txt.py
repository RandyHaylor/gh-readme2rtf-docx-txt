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

# ---------------------------------------------------------------------------
# DOCX TEXT PLACEHOLDER SYSTEM
# User-facing text is stashed behind placeholders during inline rule processing.
# After all XML structure is built, placeholders are replaced with XML-escaped text.
# This prevents raw user text (with &, <, >, etc.) from breaking XML structure.
# ---------------------------------------------------------------------------
_DOCX_TEXT_PLACEHOLDER_PREFIX = '\x00DOCXTXT_'
_DOCX_TEXT_PLACEHOLDER_STASH = {}  # {placeholder_key: raw_text}
_DOCX_TEXT_PLACEHOLDER_COUNTER = 0

def docx_stash_user_text(raw_text):
    """Stash raw user text behind a placeholder. Returns the placeholder string.
    The escape-char placeholder `\\x00ESC:` and the RTF star placeholder
    `\\x00RTFSTAR` are RTF-pipeline internals — strip them before stashing so
    they never reach the final DOCX output (they're XML-illegal null bytes).
    """
    global _DOCX_TEXT_PLACEHOLDER_COUNTER
    cleaned_text = raw_text.replace(_ESCAPE_PLACEHOLDER, '').replace(_RTF_STAR_PLACEHOLDER, '')
    placeholder_key = f'{_DOCX_TEXT_PLACEHOLDER_PREFIX}{_DOCX_TEXT_PLACEHOLDER_COUNTER}'
    _DOCX_TEXT_PLACEHOLDER_STASH[placeholder_key] = cleaned_text
    _DOCX_TEXT_PLACEHOLDER_COUNTER += 1
    return placeholder_key

def docx_restore_all_stashed_text(docx_xml_with_placeholders):
    """Replace all text placeholders with their XML-escaped content. Call once at the end.
    Uses a regex-driven sweep so placeholder prefixes never collide — the
    placeholder key `\\x00DOCXTXT_4` must not accidentally match inside
    `\\x00DOCXTXT_45`. Each match is replaced with the XML-escaped stashed text.
    The sweep loops until no placeholder remains, since stashed content may itself
    contain placeholder keys (e.g. when bold wraps an already-stashed inline code).
    """
    placeholder_regex = re.compile(
        re.escape(_DOCX_TEXT_PLACEHOLDER_PREFIX) + r'(\d+)'
    )

    def _restore_single_placeholder(match):
        placeholder_key = match.group(0)
        if placeholder_key in _DOCX_TEXT_PLACEHOLDER_STASH:
            return _xml_escape(_DOCX_TEXT_PLACEHOLDER_STASH[placeholder_key])
        # Unknown placeholder — leave it as-is so it's debuggable in the output.
        return placeholder_key

    result = docx_xml_with_placeholders
    # Fixed-point iteration with a sanity-cap so a bug can't infinite-loop us.
    for _sweep in range(16):
        new_result, substitution_count = placeholder_regex.subn(_restore_single_placeholder, result)
        if substitution_count == 0:
            return new_result
        result = new_result
    return result

def docx_reset_text_placeholder_stash():
    """Clear the stash and reset counter. Call at the start of each conversion."""
    global _DOCX_TEXT_PLACEHOLDER_COUNTER
    _DOCX_TEXT_PLACEHOLDER_STASH.clear()
    _DOCX_TEXT_PLACEHOLDER_COUNTER = 0

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
    ':briefcase:': '\U0001F4BC', ':page_facing_up:': '\U0001F4C4',
    ':airplane:': '\u2708\uFE0F', ':printer:': '\U0001F5A8\uFE0F',
    ':test_tube:': '\U0001F9EA', ':link:': '\U0001F517', ':art:': '\U0001F3A8',
    ':hammer:': '\U0001F528', ':wrench:': '\U0001F527', ':package:': '\U0001F4E6',
    ':clipboard:': '\U0001F4CB', ':mag:': '\U0001F50D', ':books:': '\U0001F4DA',
    ':book:': '\U0001F4D6', ':bookmark:': '\U0001F516', ':label:': '\U0001F3F7\uFE0F',
    ':mailbox:': '\U0001F4EB', ':email:': '\U0001F4E7', ':phone:': '\u260E\uFE0F',
    ':computer:': '\U0001F4BB', ':floppy_disk:': '\U0001F4BE', ':cd:': '\U0001F4BF',
    ':file_folder:': '\U0001F4C1', ':open_file_folder:': '\U0001F4C2',
    ':rainbow:': '\U0001F308', ':sunny:': '\u2600\uFE0F', ':cloud:': '\u2601\uFE0F',
    ':information_source:': '\u2139\uFE0F', ':question:': '\u2753', ':exclamation:': '\u2757',
    ':speech_balloon:': '\U0001F4AC', ':pencil:': '\u270F\uFE0F', ':pencil2:': '\u270F\uFE0F',
    ':scroll:': '\U0001F4DC', ':newspaper:': '\U0001F4F0', ':calendar:': '\U0001F4C5',
    ':date:': '\U0001F4C6', ':chart:': '\U0001F4C8', ':bar_chart:': '\U0001F4CA',
    ':chart_with_upwards_trend:': '\U0001F4C8', ':chart_with_downwards_trend:': '\U0001F4C9',
    ':globe_with_meridians:': '\U0001F310', ':earth_americas:': '\U0001F30E',
    ':house:': '\U0001F3E0', ':office:': '\U0001F3E2', ':factory:': '\U0001F3ED',
    ':rotating_light:': '\U0001F6A8', ':no_entry:': '\u26D4',
    ':heavy_plus_sign:': '\u2795', ':heavy_minus_sign:': '\u2796',
    ':heavy_multiplication_x:': '\u2716\uFE0F', ':heavy_division_sign:': '\u2797',
    ':arrow_left:': '\u2B05\uFE0F', ':arrow_up:': '\u2B06\uFE0F',
    ':arrow_down:': '\u2B07\uFE0F', ':arrow_up_down:': '\u2195\uFE0F',
    ':arrow_left_right:': '\u2194\uFE0F', ':recycle:': '\u267B\uFE0F',
    ':shield:': '\U0001F6E1\uFE0F', ':trophy:': '\U0001F3C6',
    ':medal:': '\U0001F3C5', ':first_place_medal:': '\U0001F947',
    ':dart:': '\U0001F3AF', ':crystal_ball:': '\U0001F52E',
    ':muscle:': '\U0001F4AA', ':raised_hands:': '\U0001F64C',
    ':clap:': '\U0001F44F', ':ok_hand:': '\U0001F44C',
    ':bell:': '\U0001F514', ':no_bell:': '\U0001F515',
    ':microphone:': '\U0001F3A4', ':speaker:': '\U0001F508',
    ':loudspeaker:': '\U0001F4E2', ':mega:': '\U0001F4E3',
    ':hourglass:': '\u231B', ':stopwatch:': '\u23F1\uFE0F',
    ':alarm_clock:': '\u23F0', ':watch:': '\u231A',
    ':smile:': '\U0001F604', ':smiley:': '\U0001F603', ':grin:': '\U0001F601',
    ':joy:': '\U0001F602', ':laughing:': '\U0001F606', ':blush:': '\U0001F60A',
    ':wink:': '\U0001F609', ':sunglasses:': '\U0001F60E', ':thinking:': '\U0001F914',
    ':confused:': '\U0001F615', ':neutral_face:': '\U0001F610',
    ':disappointed:': '\U0001F61E', ':sob:': '\U0001F62D',
    ':rage:': '\U0001F621', ':scream:': '\U0001F631',
    ':shrug:': '\U0001F937', ':ok_woman:': '\U0001F646\u200D\u2640\uFE0F',
    ':handshake:': '\U0001F91D', ':tools:': '\U0001F6E0\uFE0F',
    ':nut_and_bolt:': '\U0001F529',
    ':repeat:': '\U0001F501', ':arrows_clockwise:': '\U0001F503',
    ':hash:': '#\uFE0F\u20E3',
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

def _resolve_relative_url_if_needed(url):
    """If `url` is a repo-relative link starting with ../, rewrite it to a full
    GitHub blob URL. Absolute/anchor/mailto URLs are returned unchanged."""
    if url.startswith(('http://', 'https://', '#', 'mailto:')):
        return url
    repo = _detect_github_repo()
    if not repo:
        return url
    if url.startswith('../'):
        tail = url[3:]
        if not tail or tail.startswith('..'):
            return url
        return f'https://github.com/{repo}/blob/main/{tail}'
    return url

def _handle_md_link(match):
    """Convert [text](url) to RTF hyperlink field. Internal #anchors use \\l flag."""
    text = match.group(1)
    url = _resolve_relative_url_if_needed(match.group(2))
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

def _docx_inject_rpr(xml, extra_rpr):
    """Inject `extra_rpr` markup into every <w:rPr> block already in `xml`,
    and add a <w:rPr> to any <w:r> that currently lacks one.
    Used so bold/italic wrappers compose correctly with inner inline-code or
    link runs that were already converted by an earlier phase of apply_inline_rules.
    """
    # Inject into existing rPr blocks (insert immediately after <w:rPr>)
    result = xml.replace('<w:rPr>', '<w:rPr>' + extra_rpr)
    # Add rPr to runs that have none: <w:r> or <w:r ...> NOT followed by <w:rPr>
    result = re.sub(
        r'(<w:r(?:\s[^>]*)?>)(?!<w:rPr>)',
        lambda m, ep=extra_rpr: m.group(1) + '<w:rPr>' + ep + '</w:rPr>',
        result
    )
    return result

def _docx_wrap_or_inject(captured, rpr_xml):
    """If `captured` already contains DOCX runs/hyperlinks, inject rpr_xml into
    each run's properties (compose). Otherwise stash and emit a fresh single run."""
    if '<w:r' in captured or '<w:hyperlink' in captured:
        return _docx_inject_rpr(captured, rpr_xml)
    placeholder = docx_stash_user_text(captured)
    return f'<w:r><w:rPr>{rpr_xml}</w:rPr><w:t xml:space="preserve">{placeholder}</w:t></w:r>'

# DOCX hyperlinks use the field-code approach (mirrors RTF \field\fldinst HYPERLINK),
# embedding URLs directly in <w:instrText> rather than relationship IDs.
# This has better LibreOffice compatibility and needs no rels registration.
_DOCX_HYPERLINKS = {}  # kept only for image relationship tracking (no longer used for links)

def _docx_field_hyperlink(url_or_bookmark, display_placeholder, is_anchor=False, display_rpr=None):
    """Build DOCX field-code hyperlink — direct mirror of RTF {\field{\fldinst HYPERLINK "url"}{\fldrslt ...}}.
    is_anchor=True emits  HYPERLINK \\l "bookmark"  (internal document jump).
    URL/bookmark goes into <w:instrText> (XML-escaped now); display text is a stash placeholder.
    display_rpr overrides the default run properties for the display run (e.g. superscript for footnotes).
    """
    rpr = display_rpr if display_rpr is not None else '<w:rPr><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr>'
    if is_anchor:
        instr = f' HYPERLINK \\l "{_xml_escape(url_or_bookmark)}" '
    else:
        instr = f' HYPERLINK "{_xml_escape(url_or_bookmark)}" '
    return (
        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve">{instr}</w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{display_placeholder}</w:t></w:r>'
        f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


# ---------------------------------------------------------------------------
# DOCX IMAGE MARKER STASH
# ---------------------------------------------------------------------------
# Image markers use a distinct stash so the image pre-pass can find and swap
# them regardless of what XML escaping the text placeholder restore does later.
# Each marker records (alt_text, src_path, requested_width, requested_height).
# The pre-pass replaces the whole marker run with a <w:drawing> block and files
# the image bytes for later writing into the zip under word/media/.
_DOCX_IMAGE_MARKER_PREFIX = '\x00DOCXIMG_'
_DOCX_IMAGE_MARKERS = []  # list of dicts: alt, src, req_w, req_h

def _docx_stash_image_marker(alt_text, src_path, requested_width=0, requested_height=0):
    """Record an image marker for later pre-pass substitution. Returns a placeholder key."""
    marker_index = len(_DOCX_IMAGE_MARKERS)
    _DOCX_IMAGE_MARKERS.append({
        'alt': alt_text,
        'src': src_path,
        'req_w': requested_width,
        'req_h': requested_height,
    })
    return f'{_DOCX_IMAGE_MARKER_PREFIX}{marker_index}'

def _docx_reset_image_marker_stash():
    """Clear the image marker stash. Call at the start of each conversion."""
    _DOCX_IMAGE_MARKERS.clear()

def _docx_handle_md_link(match):
    """Convert [text](url) to DOCX field-code hyperlink. Mirrors RTF {\fldinst HYPERLINK}."""
    text = match.group(1)
    url = _resolve_relative_url_if_needed(match.group(2))
    text_placeholder = docx_stash_user_text(text)
    if url.startswith('#'):
        return _docx_field_hyperlink(url[1:], text_placeholder, is_anchor=True)
    return _docx_field_hyperlink(url, text_placeholder)

def _docx_handle_bare_url(match):
    """Convert bare URL to DOCX field-code hyperlink."""
    url = match.group(0)
    url_display_placeholder = docx_stash_user_text(url)
    return _docx_field_hyperlink(url, url_display_placeholder)

def _docx_handle_mention(match):
    """Convert @username to DOCX field-code hyperlink to GitHub profile."""
    username = match.group(1)
    url = f'https://github.com/{username}'
    display_text_placeholder = docx_stash_user_text(f'@{username}')
    return _docx_field_hyperlink(url, display_text_placeholder)

def _docx_handle_issue_ref(match):
    """Convert #42 to DOCX field-code hyperlink if repo context available."""
    number = match.group(1)
    repo = _detect_github_repo()
    display_text_placeholder = docx_stash_user_text(f'#{number}')
    if repo:
        url = f'https://github.com/{repo}/issues/{number}'
        return _docx_field_hyperlink(url, display_text_placeholder)
    return (f'<w:r><w:rPr><w:color w:val="365F91"/></w:rPr>'
            f'<w:t xml:space="preserve">{display_text_placeholder}</w:t></w:r>')

def _docx_handle_html_img(match):
    """Convert <img> tag to a DOCX image marker run. The image pre-pass later
    swaps the whole run for a <w:drawing> or a text fallback run.
    Marker run shape: <w:r><w:t>\\x00DOCXIMG_N</w:t></w:r>
    """
    tag = match.group(0)
    alt_match = re.search(r'alt="([^"]*)"', tag)
    src_match = re.search(r'src="([^"]*)"', tag)
    width_match = re.search(r'width="([^"]*)"', tag)
    height_match = re.search(r'height="([^"]*)"', tag)
    alt_text = alt_match.group(1) if alt_match else 'image'
    src_text = src_match.group(1) if src_match else ''
    requested_w = int(width_match.group(1)) if width_match and width_match.group(1).isdigit() else 0
    requested_h = int(height_match.group(1)) if height_match and height_match.group(1).isdigit() else 0
    marker_placeholder = _docx_stash_image_marker(alt_text, src_text, requested_w, requested_h)
    # Emit a full <w:r> so _docx_wrap_plain_text_in_runs leaves it alone.
    # The marker placeholder is ASCII (\x00-prefixed), so survives XML escape unchanged.
    return f'<w:r><w:t xml:space="preserve">{marker_placeholder}</w:t></w:r>'

def _docx_handle_md_image(match):
    """Convert ![alt](src) to a DOCX image marker run."""
    alt_text = match.group(1)
    src_text = match.group(2)
    marker_placeholder = _docx_stash_image_marker(alt_text, src_text, 0, 0)
    return f'<w:r><w:t xml:space="preserve">{marker_placeholder}</w:t></w:r>'

def _docx_handle_emoji(match):
    """Convert :shortcode: to emoji character for DOCX. Emoji text stashed."""
    shortcode = match.group(0)
    emoji_char = EMOJI_MAP.get(shortcode)
    if emoji_char:
        emoji_placeholder = docx_stash_user_text(emoji_char)
        return f'<w:r><w:t xml:space="preserve">{emoji_placeholder}</w:t></w:r>'
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
    import os as _os
    env_slug = _os.environ.get('GITHUB_REPOSITORY', '').strip()
    if env_slug:
        _GITHUB_REPO_SLUG = env_slug
        return _GITHUB_REPO_SLUG
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
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">{docx_stash_user_text(m.group(1))}</w:t></w:r>'})),
    ('html_sup',        (r'<sup>(.*?)</sup>',                    {'rtf': r'{\\super \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">{docx_stash_user_text(m.group(1))}</w:t></w:r>'})),
    ('html_ins',        (r'<ins>(.*?)</ins>',                    {'rtf': r'{\\ul \1}',
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">{docx_stash_user_text(m.group(1))}</w:t></w:r>'})),
    ('html_br',         (r'<br\s*/?>',                           {'rtf': r'\\line ',
                                                                    'docx': r'<w:r><w:br/></w:r>'})),

    # --- Phase 2: Escaped chars -> placeholders (before any markdown matching) ---
    ('escaped_char',    (r'\\([*#_~`\[\]\\])',                   {'rtf': _handle_escaped_char, 'docx': _handle_escaped_char})),

    # --- Phase 3: Images and links (before inline code, which uses backticks) ---
    ('md_image',        (r'!\[([^\]]*)\]\(([^)]+)\)',            {'rtf': r'{\\cf3 [Image: \1 \\u8212? \2]}',
                                                                    'docx': _docx_handle_md_image})),
    ('md_link',         (r'\[([^\]]+)\]\(([^)]+)\)',             {'rtf': _handle_md_link, 'docx': _docx_handle_md_link})),
    ('bare_url',        (r'(?<!["\(])https?://[^\s<>\)]+',       {'rtf': _handle_bare_url, 'docx': _docx_handle_bare_url})),

    # --- Phase 4: Inline code (stash content to protect from emoji/mention rules) ---
    ('inline_code',     (r'`([^`]+)`',                           {'rtf': lambda m: _stash_inline_code(m.group(1)),
                                                                    'docx': lambda m: f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/><w:shd w:val="clear" w:fill="E6F0FA"/></w:rPr><w:t xml:space="preserve">{docx_stash_user_text(m.group(1))}</w:t></w:r>'})),

    # --- Phase 5: GitHub-specific inline elements ---
    ('mention',         (r'@(\w[\w/-]*)',                         {'rtf': _handle_mention, 'docx': _docx_handle_mention})),
    ('issue_ref',       (r'(?<![&A-Fa-f0-9])#(\d+)\b',          {'rtf': _handle_issue_ref, 'docx': _docx_handle_issue_ref})),
    ('footnote_ref',    (r'\[\^([^\]]+)\]',                      {'rtf': lambda m: f'{{\\super {{\\field{{{_RTF_STAR_PLACEHOLDER}\\fldinst HYPERLINK \\\\l "fn-{m.group(1)}"}}{{\\fldrslt \\cf2  [{m.group(1)}] }}}}}}',
                                                                    'docx': lambda m: _docx_field_hyperlink(f'fn-{m.group(1)}', docx_stash_user_text(f'[{m.group(1)}]'), is_anchor=True, display_rpr='<w:rPr><w:vertAlign w:val="superscript"/><w:color w:val="365F91"/><w:u w:val="single"/></w:rPr>')})),
    ('emoji',           (r':\w+:',                                {'rtf': _handle_emoji, 'docx': _docx_handle_emoji})),

    # --- Phase 6: Text formatting ---
    ('bold_italic',     (r'\*\*\*(.+?)\*\*\*',                   {'rtf': r'{\\b\\i \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:b/><w:i/>')})),
    ('bold_star',       (r'\*\*(.+?)\*\*',                        {'rtf': r'{\\b \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:b/>')})),
    ('bold_under',      (r'__(.+?)__',                             {'rtf': r'{\\b \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:b/>')})),
    ('italic_star',     (r'\*(.+?)\*',                             {'rtf': r'{\\i \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:i/>')})),
    ('italic_under',    (r'(?<!\w)_(.+?)_(?!\w)',                  {'rtf': r'{\\i \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:i/>')})),
    ('strikethrough',   (r'~~(.+?)~~',                             {'rtf': r'{\\strike \1}',
                                                                    'docx': lambda m: _docx_wrap_or_inject(m.group(1), '<w:strike/>')})),
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
# Alert types: (label, rtf_bar_color_index, rtf_label_text_color_index, docx_hex_color)
# RTF bar colors: 13=blue, 14=green, 15=purple, 16=amber, 17=red
# DOCX hex colors match GitHub GFM alert palette
ALERT_TYPES = {
    '[!NOTE]':      ('Note',      13, 13, '0969DA'),
    '[!TIP]':       ('Tip',       14, 14, '1A7F37'),
    '[!IMPORTANT]': ('Important', 15, 15, '8250DF'),
    '[!WARNING]':   ('Warning',   16, 16, '9A6700'),
    '[!CAUTION]':   ('Caution',   17, 17, 'CF222E'),
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
        alert_label, alert_bar_color, alert_text_color, _ = ALERT_TYPES[quote_lines[0].strip()]
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
        formatted_heading_inline = apply_inline_rules(raw_text, fmt='docx')
        wrapped_heading_runs = _docx_wrap_plain_text_in_runs(formatted_heading_inline)
        xml = (f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr>'
               f'<w:bookmarkStart w:id="{bid}" w:name="{bookmark_id}"/>'
               f'<w:bookmarkEnd w:id="{bid}"/>'
               f'{wrapped_heading_runs}</w:p>')
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

# Map RTF color index -> DOCX hex (mirrors the colortbl entries above)
_DOCX_SYNTAX_HEX_BY_RTF_INDEX = {
    1:  '000000',  # black / default
    18: '000080',  # keyword (navy)
    19: 'A31515',  # string (brown)
    20: '008000',  # comment (green)
    21: '2B91AF',  # type (teal)
    22: '800080',  # builtin (purple)
    23: '808080',  # punctuation (gray)
    24: '0000FF',  # number (blue)
}

def _syntax_highlighted_docx_runs_by_line(code_text, language):
    """Tokenize `code_text` with Pygments and return a list of lines, where each
    line is a list of (text_segment, hex_color_or_None) tuples. Lines are split
    on '\n' inside tokens so callers can emit one <w:p> per line."""
    if not PYGMENTS_AVAILABLE or not language:
        return [[(seg, None)] for seg in code_text.split('\n')]
    try:
        lexer = get_lexer_by_name(language, stripall=False)
    except Exception:
        return [[(seg, None)] for seg in code_text.split('\n')]

    lines_of_runs = [[]]
    for token_type, token_value in lex(code_text, lexer):
        # Walk up the token hierarchy to find a mapped RTF color index
        color_rtf_index = None
        tt = token_type
        while tt and not color_rtf_index:
            color_rtf_index = SYNTAX_COLOR_MAP.get(tt)
            tt = tt.parent
        hex_color = _DOCX_SYNTAX_HEX_BY_RTF_INDEX.get(color_rtf_index) if color_rtf_index else None
        # Split the token's text on newlines, starting new paragraph buckets as we go
        segments = token_value.split('\n')
        for seg_idx, seg_text in enumerate(segments):
            if seg_idx > 0:
                lines_of_runs.append([])
            if seg_text:
                lines_of_runs[-1].append((seg_text, hex_color))
    return lines_of_runs

def docx_block_fenced_code(lines, index):
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
        if indent and current_line.startswith(indent):
            current_line = current_line[len(indent):]
        code_lines.append(current_line)
        consumed += 1
    raw_code = '\n'.join(code_lines)
    lines_of_runs = _syntax_highlighted_docx_runs_by_line(raw_code, language)

    paragraph_xml_parts = []
    for run_tuples in lines_of_runs:
        run_xml_parts = []
        if not run_tuples:
            # empty line — emit a single empty run so paragraph still has shading
            run_tuples = [('', None)]
        for seg_text, hex_color in run_tuples:
            seg_placeholder = docx_stash_user_text(seg_text)
            color_rpr_xml = f'<w:color w:val="{hex_color}"/>' if hex_color else ''
            run_xml_parts.append(
                f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>'
                f'<w:sz w:val="20"/>{color_rpr_xml}</w:rPr>'
                f'<w:t xml:space="preserve">{seg_placeholder}</w:t></w:r>'
            )
        paragraph_xml_parts.append(
            f'<w:p><w:pPr><w:shd w:val="clear" w:fill="E6F0FA"/>'
            f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            + ''.join(run_xml_parts) +
            f'</w:p>'
        )
    return ('\n'.join(paragraph_xml_parts), consumed)

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

    # Header row — stash cell text (preserves original behavior of not running
    # inline rules on headers) and emit a bold run.
    parts.append('<w:tr>')
    for cell in header_row:
        header_cell_placeholder = docx_stash_user_text(cell)
        parts.append(f'<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                     f'<w:r><w:rPr><w:b/></w:rPr>'
                     f'<w:t xml:space="preserve">{header_cell_placeholder}</w:t></w:r></w:p></w:tc>')
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
    alert_hex_color = 'CCCCCC'
    if quote_lines and quote_lines[0].strip() in ALERT_TYPES:
        alert_label, _, _, alert_hex_color = ALERT_TYPES[quote_lines[0].strip()]
        quote_lines = quote_lines[1:]
        is_alert = True

    def _alert_ppr(hex_color, indent=720):
        return (f'<w:pPr><w:ind w:left="{indent}"/><w:pBdr>'
                f'<w:left w:val="single" w:sz="16" w:space="4" w:color="{hex_color}"/>'
                f'</w:pBdr></w:pPr>')

    parts = []
    for ql in quote_lines:
        text = apply_inline_rules(ql, fmt='docx') if ql.strip() else ''
        wrapped = _docx_wrap_plain_text_in_runs(text) if text else ''
        if is_alert:
            parts.append(f'<w:p>{_alert_ppr(alert_hex_color)}{wrapped}</w:p>')
        else:
            parts.append(f'<w:p><w:pPr><w:ind w:left="720"/><w:pBdr>'
                         f'<w:left w:val="single" w:sz="12" w:space="4" w:color="CCCCCC"/>'
                         f'</w:pBdr></w:pPr>{wrapped}</w:p>')
    if is_alert:
        label_placeholder = docx_stash_user_text(alert_label)
        alert_para = (f'<w:p>{_alert_ppr(alert_hex_color)}'
                      f'<w:r><w:rPr><w:b/><w:color w:val="{alert_hex_color}"/></w:rPr>'
                      f'<w:t>{label_placeholder}</w:t></w:r></w:p>')
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
    """Wrap any plain text segments (not already in <w:r> tags) into <w:r><w:t> runs.
    Plain text is routed through docx_stash_user_text() so the final restore pass
    performs the single XML escape step. This keeps the stash as the single source
    of truth for escaping user content.

    Uses an index-walk rather than a regex split so `<w:hyperlink>...</w:hyperlink>`
    (which contains nested `<w:r>...</w:r>` children) is kept intact as a single
    protected element. A naive non-greedy alternation would match the inner
    `</w:r>` as the hyperlink close and strand the real `</w:hyperlink>` text.
    """
    result_pieces = []
    cursor = 0
    text_length = len(text)

    # Precompiled matchers for each protected element kind.
    bookmark_self_close_regex = re.compile(r'<w:(?:bookmarkStart|bookmarkEnd)\b[^>]*/>')
    hyperlink_open_regex = re.compile(r'<w:hyperlink\b[^>]*>')
    run_open_regex = re.compile(r'<w:r\b[^>]*>')
    run_self_close_regex = re.compile(r'<w:r\b[^>]*/>')

    def _find_matching_close(haystack, start_pos, open_tag_name, close_tag_name):
        """Scan from start_pos through haystack, tracking nesting depth of
        <open_tag_name>...</open_tag_name>. Returns the index just past the
        matching close tag, or -1 if not found.
        """
        open_pattern = re.compile(r'<' + open_tag_name + r'\b[^>]*>')
        close_pattern = re.compile(r'</' + open_tag_name + r'>')
        depth = 1
        scan_pos = start_pos
        while depth > 0 and scan_pos < len(haystack):
            next_open = open_pattern.search(haystack, scan_pos)
            next_close = close_pattern.search(haystack, scan_pos)
            if next_close is None:
                return -1
            if next_open is not None and next_open.start() < next_close.start():
                depth += 1
                scan_pos = next_open.end()
            else:
                depth -= 1
                scan_pos = next_close.end()
        return scan_pos

    while cursor < text_length:
        # Try each kind of protected element at cursor.
        hyperlink_match = hyperlink_open_regex.match(text, cursor)
        run_match = run_open_regex.match(text, cursor)
        run_self_close_match = run_self_close_regex.match(text, cursor)
        bookmark_match = bookmark_self_close_regex.match(text, cursor)

        if hyperlink_match is not None:
            close_end = _find_matching_close(text, hyperlink_match.end(), 'w:hyperlink', 'w:hyperlink')
            if close_end == -1:
                close_end = text_length
            result_pieces.append(text[cursor:close_end])
            cursor = close_end
            continue

        if run_self_close_match is not None:
            result_pieces.append(run_self_close_match.group(0))
            cursor = run_self_close_match.end()
            continue

        if run_match is not None:
            close_end = _find_matching_close(text, run_match.end(), 'w:r', 'w:r')
            if close_end == -1:
                close_end = text_length
            result_pieces.append(text[cursor:close_end])
            cursor = close_end
            continue

        if bookmark_match is not None:
            result_pieces.append(bookmark_match.group(0))
            cursor = bookmark_match.end()
            continue

        # Not at a protected element — accumulate plain text until the next one.
        next_protected_positions = []
        for regex in (hyperlink_open_regex, run_open_regex, run_self_close_regex, bookmark_self_close_regex):
            probe = regex.search(text, cursor)
            if probe is not None:
                next_protected_positions.append(probe.start())
        next_stop = min(next_protected_positions) if next_protected_positions else text_length
        plain_text_segment = text[cursor:next_stop]
        if plain_text_segment:
            stashed_placeholder = docx_stash_user_text(plain_text_segment)
            result_pieces.append(f'<w:r><w:t xml:space="preserve">{stashed_placeholder}</w:t></w:r>')
        cursor = next_stop

    return ''.join(result_pieces)

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
    """Build footnotes section as DOCX XML paragraphs. User text routes through
    the placeholder stash so the single final restore pass handles XML escaping.
    """
    if not _collected_footnotes:
        return ''
    parts = []
    parts.append('<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>')
    parts.append('<w:p><w:pPr><w:spacing w:before="120"/></w:pPr>'
                 '<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>Footnotes</w:t></w:r></w:p>')
    for fid, ftext in _collected_footnotes:
        formatted_footnote_xml = apply_inline_rules(ftext, fmt='docx')
        wrapped_footnote_xml = _docx_wrap_plain_text_in_runs(' ' + formatted_footnote_xml)
        bookmark = f'fn-{fid}'
        bid = hash(bookmark) % 10000
        fid_label_placeholder = docx_stash_user_text(f'{fid}.')
        parts.append(
            f'<w:p><w:pPr><w:ind w:left="360"/><w:spacing w:after="40"/></w:pPr>'
            f'<w:bookmarkStart w:id="{bid}" w:name="{bookmark}"/>'
            f'<w:bookmarkEnd w:id="{bid}"/>'
            f'<w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr>'
            f'<w:t xml:space="preserve">{fid_label_placeholder}</w:t></w:r>'
            f'{wrapped_footnote_xml}</w:p>'
        )
    return '\n'.join(parts)


def _docx_inject_sectpr_into_last_paragraph(body_xml, section_properties_xml):
    """Move the section properties element into the last paragraph's pPr.

    Why: a body-level `<w:sectPr>` causes Word to render an implicit trailing
    empty paragraph, which appears as a blank final page. Embedding it inside
    the last paragraph's `<w:pPr>` defines the section without that ghost.

    If the last paragraph already has a `<w:pPr>`, append the sectPr just
    before `</w:pPr>`. If it has no pPr, inject a new pPr containing only the
    sectPr right after the opening `<w:p>` tag.

    Fallback: if no `</w:p>` is found (empty document), return the body_xml
    with the sectPr appended directly — matches the old behavior.
    """
    last_paragraph_close_index = body_xml.rfind('</w:p>')
    if last_paragraph_close_index == -1:
        return body_xml + section_properties_xml
    last_paragraph_open_index = body_xml.rfind('<w:p>', 0, last_paragraph_close_index)
    last_paragraph_open_with_attrs_index = body_xml.rfind('<w:p ', 0, last_paragraph_close_index)
    paragraph_start_index = max(last_paragraph_open_index, last_paragraph_open_with_attrs_index)
    if paragraph_start_index == -1:
        return body_xml + section_properties_xml
    last_paragraph_xml = body_xml[paragraph_start_index:last_paragraph_close_index]
    existing_ppr_close_index = last_paragraph_xml.find('</w:pPr>')
    if existing_ppr_close_index != -1:
        rebuilt_last_paragraph = (
            last_paragraph_xml[:existing_ppr_close_index]
            + section_properties_xml
            + last_paragraph_xml[existing_ppr_close_index:]
        )
    else:
        paragraph_open_tag_end_index = last_paragraph_xml.find('>') + 1
        rebuilt_last_paragraph = (
            last_paragraph_xml[:paragraph_open_tag_end_index]
            + f'<w:pPr>{section_properties_xml}</w:pPr>'
            + last_paragraph_xml[paragraph_open_tag_end_index:]
        )
    return (
        body_xml[:paragraph_start_index]
        + rebuilt_last_paragraph
        + body_xml[last_paragraph_close_index:]
    )


def _docx_cleanup_structural_markers(body_xml):
    """Run once after all blocks are assembled, before text restoration.
    Rewrites any `[tag]`-style structural markers to proper nested XML. Currently
    a no-op placeholder — no block rule emits structural markers today, but this
    hook is wired into the pipeline so future refactors can use it without
    touching the caller.
    """
    return body_xml


# EMU (English Metric Units) per pixel at 96 dpi: 914400 EMU / inch, 96 px / inch
_DOCX_EMU_PER_PIXEL = 9525
# Default page content area used to clamp oversized images. Same 6.5" x 9" as RTF.
_DOCX_MAX_IMAGE_WIDTH_PX = 468
_DOCX_MAX_IMAGE_HEIGHT_PX = 648


def _docx_compute_display_dimensions(native_w, native_h, requested_w, requested_h):
    """Replicate rtf_image_embedder._replace_placeholder sizing math.
    Returns (display_w_px, display_h_px).
    """
    w = requested_w or native_w
    h = requested_h
    if w and not h and native_w and native_h:
        h = round(w * native_h / native_w)
    elif h and not w and native_w and native_h:
        w = round(h * native_w / native_h)
    elif not h:
        h = native_h
    if w > _DOCX_MAX_IMAGE_WIDTH_PX:
        scale = _DOCX_MAX_IMAGE_WIDTH_PX / w
        w = _DOCX_MAX_IMAGE_WIDTH_PX
        h = round(h * scale)
    if h > _DOCX_MAX_IMAGE_HEIGHT_PX:
        scale = _DOCX_MAX_IMAGE_HEIGHT_PX / h
        h = _DOCX_MAX_IMAGE_HEIGHT_PX
        w = round(w * scale)
    return w, h


def _docx_build_drawing_xml_for_inline_image(relationship_id, drawing_unique_id, alt_text, display_width_px, display_height_px):
    """Build a <w:drawing> XML block for a single inline image reference."""
    emu_w = display_width_px * _DOCX_EMU_PER_PIXEL
    emu_h = display_height_px * _DOCX_EMU_PER_PIXEL
    # Alt text inside XML attributes needs escaping now — the main stash/restore
    # pass has already run by the time this XML is stitched in. Escape inline.
    safe_alt = _xml_escape(alt_text)
    return (
        '<w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<wp:extent cx="{emu_w}" cy="{emu_h}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{drawing_unique_id}" name="Picture {drawing_unique_id}" descr="{safe_alt}"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="{drawing_unique_id}" name="Picture {drawing_unique_id}" descr="{safe_alt}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip r:embed="{relationship_id}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu_w}" cy="{emu_h}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing></w:r>'
    )


def _docx_build_image_rels_from_stashed_embeds(base_dir_for_image_paths):
    """Iterate the image marker stash, resolve each image file, downscale via
    rtf_image_embedder private helpers, and produce the data needed to swap
    markers for <w:drawing> runs and to write image bytes into the zip.

    Returns a list of dicts, one per stashed image marker:
        {
          'marker_key': str,          # the \\x00DOCXIMG_N token to find in XML
          'resolved': bool,           # True if the file was readable
          'alt_text': str,
          'drawing_xml': str,         # only if resolved
          'fallback_text': str,       # only if not resolved
          'zip_member_path': str,     # e.g. 'word/media/image1.png' (resolved only)
          'image_bytes': bytes,       # (resolved only)
          'extension': str,           # 'png' or 'jpeg' (resolved only)
          'relationship_id': str,     # e.g. 'rIdImage1' (resolved only)
        }
    """
    import os
    import rtf_image_embedder as embedder_module
    results = []
    image_sequence_number = 0
    for marker_index, meta in enumerate(_DOCX_IMAGE_MARKERS):
        marker_key = f'{_DOCX_IMAGE_MARKER_PREFIX}{marker_index}'
        alt_text = meta['alt']
        src_path = meta['src']
        req_w = meta['req_w']
        req_h = meta['req_h']
        # Resolve relative paths against base_dir, same as rtf_image_embedder does.
        if os.path.isabs(src_path):
            absolute_image_path = src_path
        else:
            absolute_image_path = os.path.join(base_dir_for_image_paths, src_path)
        if not os.path.exists(absolute_image_path):
            results.append({
                'marker_key': marker_key,
                'resolved': False,
                'alt_text': alt_text,
                'fallback_text': f'[Image not found: {src_path}]',
            })
            continue
        extension_lower = os.path.splitext(src_path)[1].lower()
        if extension_lower in ('.jpg', '.jpeg'):
            zip_extension = 'jpeg'
        elif extension_lower == '.png':
            zip_extension = 'png'
        else:
            # Unsupported extension — fall back to text marker.
            results.append({
                'marker_key': marker_key,
                'resolved': False,
                'alt_text': alt_text,
                'fallback_text': f'[Image: {alt_text} \u2014 {src_path}]',
            })
            continue
        native_w, native_h = embedder_module._read_image_native_size(absolute_image_path)
        display_w, display_h = _docx_compute_display_dimensions(native_w, native_h, req_w, req_h)
        # Downscale to display dimensions; the helper returns the bytes as-is.
        downscaled_bytes = embedder_module._downscale_image(absolute_image_path, display_w, display_h)
        image_sequence_number += 1
        relationship_id = f'rIdImage{image_sequence_number}'
        zip_member_path = f'word/media/image{image_sequence_number}.{zip_extension}'
        drawing_unique_id = 1000 + image_sequence_number
        drawing_xml = _docx_build_drawing_xml_for_inline_image(
            relationship_id,
            drawing_unique_id,
            alt_text,
            display_w,
            display_h,
        )
        results.append({
            'marker_key': marker_key,
            'resolved': True,
            'alt_text': alt_text,
            'drawing_xml': drawing_xml,
            'zip_member_path': zip_member_path,
            'image_bytes': downscaled_bytes,
            'extension': zip_extension,
            'relationship_id': relationship_id,
        })
    return results


def _docx_substitute_image_markers_in_xml(body_xml, image_embed_records):
    """Replace each image marker run with either a <w:drawing> run or a text
    fallback run. Consumes the marker run `<w:r>...<w:t ...>KEY</w:t></w:r>`
    as a whole to avoid leftover empty runs.
    """
    result_xml = body_xml
    for record in image_embed_records:
        marker_key = record['marker_key']
        # Match the exact run we emitted: <w:r><w:t xml:space="preserve">KEY</w:t></w:r>
        # Use a tolerant regex so future tweaks to the run wrapper still match.
        marker_run_pattern = re.compile(
            r'<w:r>\s*<w:t[^>]*>' + re.escape(marker_key) + r'</w:t>\s*</w:r>'
        )
        if record['resolved']:
            replacement_run = record['drawing_xml']
        else:
            safe_fallback = _xml_escape(record['fallback_text'])
            replacement_run = (
                '<w:r><w:rPr><w:color w:val="666666"/></w:rPr>'
                f'<w:t xml:space="preserve">{safe_fallback}</w:t></w:r>'
            )
        result_xml, substitution_count = marker_run_pattern.subn(replacement_run, result_xml)
        # If the run shape didn't match (e.g., wrapped differently by a future
        # block rule), fall back to swapping just the bare marker token.
        if substitution_count == 0:
            if record['resolved']:
                result_xml = result_xml.replace(marker_key, record['drawing_xml'])
            else:
                result_xml = result_xml.replace(marker_key, _xml_escape(record['fallback_text']))
    return result_xml


def convert_markdown_to_docx(markdown_text, output_path):
    """Convert a GFM markdown string to a DOCX file."""
    import zipfile
    import os
    _DOCX_HYPERLINKS.clear()  # image rels only; hyperlinks now use field-code embedding
    _collected_footnotes.clear()
    docx_reset_text_placeholder_stash()
    _docx_reset_image_marker_stash()

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

    # Assemble body, run structural-marker cleanup, then the image pre-pass,
    # then the single final text-placeholder restore.
    body_content = '\n'.join(body_parts)
    body_content = _docx_cleanup_structural_markers(body_content)

    base_dir_for_image_paths = os.path.dirname(os.path.abspath(output_path))
    image_embed_records = _docx_build_image_rels_from_stashed_embeds(base_dir_for_image_paths)
    body_content = _docx_substitute_image_markers_in_xml(body_content, image_embed_records)

    # Single final escape pass — restore all stashed user text with XML escaping.
    body_content = docx_restore_all_stashed_text(body_content)

    # Embed sectPr inside the last paragraph's pPr instead of as a direct <w:body>
    # child. A body-level sectPr causes Word to render an implicit trailing empty
    # paragraph, which shows up as a blank final page. Embedding it in the last
    # paragraph's pPr defines the section without adding a ghost paragraph.
    section_properties_xml = (
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        '</w:sectPr>'
    )
    body_content_with_section_properties = _docx_inject_sectpr_into_last_paragraph(
        body_content, section_properties_xml
    )
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body>'
        f'{body_content_with_section_properties}'
        '</w:body></w:document>'
    )

    # Build document.xml.rels (styles + hyperlinks + images)
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
    # Hyperlinks now use field-code embedding (no rels entries needed for links).
    for record in image_embed_records:
        if not record['resolved']:
            continue
        target_relative = record['zip_member_path'][len('word/'):]  # "media/imageN.ext"
        rels.append(
            f'<Relationship Id="{record["relationship_id"]}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="{target_relative}"/>'
        )
    rels.append('</Relationships>')
    doc_rels_xml = '\n'.join(rels)

    # Content types — declare image defaults only when we actually embedded them.
    any_png_embedded = any(r['resolved'] and r['extension'] == 'png' for r in image_embed_records)
    any_jpeg_embedded = any(r['resolved'] and r['extension'] == 'jpeg' for r in image_embed_records)
    content_types_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
    ]
    if any_png_embedded:
        content_types_parts.append('<Default Extension="png" ContentType="image/png"/>')
    if any_jpeg_embedded:
        content_types_parts.append('<Default Extension="jpeg" ContentType="image/jpeg"/>')
    content_types_parts.extend([
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
        '</Types>',
    ])
    content_types = ''.join(content_types_parts)

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
        for record in image_embed_records:
            if record['resolved']:
                zf.writestr(record['zip_member_path'], record['image_bytes'])


# ---------------------------------------------------------------------------
# TXT CONVERSION
# Markdown is passed through verbatim. The only transformation is resolving
# relative URLs in [text](url) links to full GitHub blob URLs so an offline
# reader has complete, clickable paths instead of broken relative references.
# ---------------------------------------------------------------------------

def _txt_resolve_relative_links_only(markdown_text, input_file_path=None):
    """Return markdown unchanged except: relative URLs in [text](url) links
    are resolved to full GitHub blob URLs using the detected repo slug.
    Absolute URLs (http/https), anchors (#), and mailto: are untouched.
    """
    import posixpath

    repo = _detect_github_repo()
    if not repo:
        return markdown_text

    # Determine the directory of the input file relative to the repo root
    # so we can resolve paths like ../SKILL.md correctly.
    input_dir_relative_to_repo_root = ''
    if input_file_path:
        import subprocess as _sp
        try:
            repo_root = _sp.run(
                ['git', 'rev-parse', '--show-toplevel'],
                capture_output=True, text=True
            ).stdout.strip()
            abs_input_dir = os.path.dirname(os.path.abspath(input_file_path))
            input_dir_relative_to_repo_root = os.path.relpath(abs_input_dir, repo_root).replace(os.sep, '/')
            if input_dir_relative_to_repo_root == '.':
                input_dir_relative_to_repo_root = ''
        except Exception:
            pass

    base_github_blob_url = f'https://github.com/{repo}/blob/main'

    def _resolve_one_link(match):
        link_text = match.group(1)
        url = match.group(2)
        if url.startswith(('http://', 'https://', '#', 'mailto:')):
            return match.group(0)  # leave absolute links and anchors unchanged
        # ../file.txt → strip leading ../ and resolve against repo root
        if url.startswith('../'):
            resolved_path = url[3:]
        elif input_dir_relative_to_repo_root:
            resolved_path = posixpath.normpath(
                posixpath.join(input_dir_relative_to_repo_root, url)
            )
        else:
            resolved_path = posixpath.normpath(url)
        if not resolved_path or resolved_path.startswith('..'):
            return match.group(0)
        full_url = f'{base_github_blob_url}/{resolved_path}'
        return f'[{link_text}]({full_url})'

    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', _resolve_one_link, markdown_text)


def convert_markdown_to_txt(markdown_text, input_file_path=None):
    """Return markdown verbatim with only relative URLs in [text](url) links resolved."""
    return _txt_resolve_relative_links_only(markdown_text, input_file_path)

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
    elif output_path.endswith('.txt'):
        txt_output = convert_markdown_to_txt(markdown_content, input_markdown_path)
        with open(output_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write(txt_output)
    else:
        rtf_output = convert_markdown_to_rtf(markdown_content)
        with open(output_path, 'w', encoding='utf-8') as rtf_file:
            rtf_file.write(rtf_output)

    print(f'Converted: {input_markdown_path} -> {output_path}')
