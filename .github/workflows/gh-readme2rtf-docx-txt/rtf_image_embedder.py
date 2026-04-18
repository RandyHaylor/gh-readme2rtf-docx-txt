#!/usr/bin/env python3
"""
rtf_image_embedder.py — Embed local images into RTF files.
Standalone module. Finds [Image: ...] placeholders in RTF content,
reads the referenced image files, downscales if needed, and replaces
the placeholders with embedded \pict blocks.

Placeholder format (produced by gfm_markdown_to_rtf.py):
  [Image: alt text (w:300, h:200) — path/to/image.png]
  [Image: alt text (w:300) — path/to/image.png]
  [Image: alt text — path/to/image.png]

Usage as module:
  from rtf_image_embedder import embed_images_in_rtf
  rtf_with_images = embed_images_in_rtf(rtf_content, base_dir='/path/to/images')

Usage as CLI:
  python3 rtf_image_embedder.py <input.rtf> [output.rtf] [--base-dir /path/to/images]
"""
import re
import sys
import os
import io
import struct

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

# Max page dimensions in pixels (letter page 8.5x11 minus 1" margins each side)
MAX_WIDTH_PX = 468   # 6.5 inches * 72 dpi
MAX_HEIGHT_PX = 648  # 9 inches * 72 dpi

# RTF twips per pixel (approx)
TWIPS_PER_PX = 15

# Regex to match [Image: ...] placeholders in RTF
# The RTF escape for em-dash is \u8212? and — is literal
IMAGE_PLACEHOLDER_PATTERN = re.compile(
    r'\[Image:\s*([^(\]]*?)\s*'           # alt text
    r'(?:\(([^)]*)\)\s*)?'                # optional (w:300, h:200)
    r'(?:\\u8212\?|—)\s*'                 # em-dash separator (RTF escaped or literal)
    r'([^\]]+?)\s*\]'                     # image path
)


def _parse_dimensions(dim_string):
    """Parse '(w:300, h:200)' or '(w:300)' into (width, height) tuple."""
    w, h = 0, 0
    if not dim_string:
        return w, h
    w_match = re.search(r'w:(\d+)', dim_string)
    h_match = re.search(r'h:(\d+)', dim_string)
    if w_match:
        w = int(w_match.group(1))
    if h_match:
        h = int(h_match.group(1))
    return w, h


def _read_image_native_size(filepath):
    """Read native width/height of an image file."""
    if PILLOW_AVAILABLE:
        with Image.open(filepath) as img:
            return img.width, img.height
    # Fallback: read PNG header directly
    with open(filepath, 'rb') as f:
        header = f.read(24)
        if len(header) >= 24 and header[:4] == b'\x89PNG':
            w = struct.unpack('>I', header[16:20])[0]
            h = struct.unpack('>I', header[20:24])[0]
            return w, h
    return 0, 0


def _downscale_image(filepath, target_w, target_h):
    """Downscale image to target dimensions. Returns PNG bytes."""
    if not PILLOW_AVAILABLE:
        with open(filepath, 'rb') as f:
            return f.read()
    with Image.open(filepath) as img:
        if img.width > target_w or img.height > target_h:
            img.thumbnail((target_w, target_h), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        return buf.getvalue()


def _build_pict_block(image_bytes, display_w, display_h, is_jpeg=False):
    """Build an RTF \\pict block from image bytes."""
    hex_data = image_bytes.hex()
    pic_type = 'jpegblip' if is_jpeg else 'pngblip'
    size_params = ''
    if display_w:
        size_params += f'\\picwgoal{display_w * TWIPS_PER_PX}'
    if display_h:
        size_params += f'\\pichgoal{display_h * TWIPS_PER_PX}'
    return f'{{\\pict\\{pic_type}{size_params} {hex_data}}}'


def _replace_placeholder(match, base_dir):
    """Replace a single [Image: ...] placeholder with an embedded \\pict block."""
    alt_text = match.group(1).strip()
    dim_string = match.group(2)
    src_path = match.group(3).strip()

    # Resolve path
    image_path = os.path.join(base_dir, src_path) if not os.path.isabs(src_path) else src_path
    if not os.path.exists(image_path):
        return f'[Image not found: {src_path}]'

    # Determine format
    ext = os.path.splitext(src_path)[1].lower()
    is_jpeg = ext in ('.jpg', '.jpeg')

    # Get native size
    native_w, native_h = _read_image_native_size(image_path)

    # Parse requested dimensions
    req_w, req_h = _parse_dimensions(dim_string)

    # Calculate display dimensions
    w = req_w or native_w
    h = req_h
    if w and not h and native_w and native_h:
        h = round(w * native_h / native_w)
    elif h and not w and native_w and native_h:
        w = round(h * native_w / native_h)
    elif not h:
        h = native_h

    # Clamp to page dimensions
    if w > MAX_WIDTH_PX:
        scale = MAX_WIDTH_PX / w
        w = MAX_WIDTH_PX
        h = round(h * scale)
    if h > MAX_HEIGHT_PX:
        scale = MAX_HEIGHT_PX / h
        h = MAX_HEIGHT_PX
        w = round(w * scale)

    # Read and downscale image
    image_bytes = _downscale_image(image_path, w, h)

    return _build_pict_block(image_bytes, w, h, is_jpeg)


def embed_images_in_rtf(rtf_content, base_dir='.'):
    """Find all [Image: ...] placeholders in RTF content and replace with embedded images."""
    return IMAGE_PLACEHOLDER_PATTERN.sub(
        lambda m: _replace_placeholder(m, base_dir),
        rtf_content
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Embed images into RTF files')
    parser.add_argument('input', help='Input RTF file')
    parser.add_argument('output', nargs='?', help='Output RTF file (default: overwrite input)')
    parser.add_argument('--base-dir', default=None, help='Base directory for resolving image paths')
    args = parser.parse_args()

    output_path = args.output or args.input
    base_dir = args.base_dir or os.path.dirname(os.path.abspath(args.input))

    with open(args.input, 'r', encoding='utf-8') as f:
        rtf_content = f.read()

    result = embed_images_in_rtf(rtf_content, base_dir)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result)

    print(f'Embedded images: {args.input} -> {output_path}')
