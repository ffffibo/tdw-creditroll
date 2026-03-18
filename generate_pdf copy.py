#!/usr/bin/env python3
"""
Generate a single-page PDF with all unique credit names separated by " / ".
Font weight per character is determined by a brightness map image:
  - White areas  → Roman (light)
  - Gray areas   → Medium
  - Black areas  → Bold
"""

import pandas as pd
from PIL import Image
import numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import html
import os

# ── Configuration ──────────────────────────────────────────────────────────
INPUT_CSV = "tdw-creditroll-2026-02-17.csv"
OUTPUT_PDF = "tdw-creditroll-2026-02-17.pdf"
IMAGE_MAP = "tdw-gallpeters.png"

FONT_DIR = "neue-haas-grotesk-display-pro"
FONT_ROMAN = os.path.join(FONT_DIR, "NeueHaasDisplayLight.ttf")
FONT_MEDIUM = os.path.join(FONT_DIR, "NeueHaasDisplayMediu.ttf")
FONT_BOLD   = os.path.join(FONT_DIR, "NeueHaasDisplayBlack.ttf")

FONT_NAME_ROMAN = "NHGDisplayRoman"
FONT_NAME_MEDIUM = "NHGDisplayMedium"
FONT_NAME_BOLD = "NHGDisplayBold"

# Brightness thresholds (0 = black, 255 = white)
THRESHOLD_BLACK = 85    # pixel brightness <= this → Bold (Black)
THRESHOLD_WHITE = 200   # pixel brightness >= this → Light
                        # in between → Medium

# Custom page size: 295mm x 325mm
PAGE_WIDTH = 295 * mm
PAGE_HEIGHT = 325 * mm
PAGE_SIZE = (PAGE_WIDTH, PAGE_HEIGHT)

MARGIN_X = 10 * mm
MARGIN_Y = 10 * mm
SEPARATOR = " / "
LINE_SPACING_FACTOR = 1.05  # tight line spacing


def register_fonts():
    """Register the three font weights."""
    pdfmetrics.registerFont(TTFont(FONT_NAME_ROMAN, FONT_ROMAN))
    pdfmetrics.registerFont(TTFont(FONT_NAME_MEDIUM, FONT_MEDIUM))
    pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, FONT_BOLD))
    print("Fonts registered successfully.")


def load_brightness_map(image_path, target_w, target_h):
    """
    Load the image and convert to a grayscale numpy array
    resized to match the text area dimensions (in pixels at 72 dpi ≈ points).
    Returns a 2D numpy array of brightness values (0–255).
    """
    img = Image.open(image_path).convert("L")  # grayscale
    img_resized = img.resize((int(target_w), int(target_h)), Image.LANCZOS)
    return np.array(img_resized)


def get_font_for_brightness(brightness):
    """Return the font name based on pixel brightness."""
    if brightness >= THRESHOLD_WHITE:
        return FONT_NAME_ROMAN
    elif brightness <= THRESHOLD_BLACK:
        return FONT_NAME_BOLD
    else:
        return FONT_NAME_MEDIUM


def find_font_size(c, text, text_width, text_height, min_size=3, max_size=20):
    """
    Binary search for the largest font size that fits all text
    within the given text_width x text_height area.
    Uses the Roman font for measurement (widest/narrowest differences are small).
    """
    def text_fits(font_size):
        leading = font_size * LINE_SPACING_FACTOR
        # Measure how many lines the text would take
        words = text.split(' ')
        x = 0
        lines = 1
        for word in words:
            word_w = c.stringWidth(word + ' ', FONT_NAME_ROMAN, font_size)
            if x + word_w > text_width and x > 0:
                lines += 1
                x = 0
            x += word_w
        total_height = lines * leading
        return total_height <= text_height

    # Binary search
    lo, hi = min_size, max_size
    best = min_size
    while lo <= hi:
        mid = (lo + hi) / 2
        if text_fits(mid):
            best = mid
            lo = mid + 0.25
        else:
            hi = mid - 0.25
    return best


def layout_characters(c, text, font_size, text_width, text_height, x_start, y_start):
    """
    Lay out all characters with word-wrapping.
    Returns a list of (char, x, y) tuples for each character.
    """
    line_spacing = font_size * LINE_SPACING_FACTOR
    chars = []
    x = 0
    y = 0
    
    words = text.split(' ')
    for w_idx, word in enumerate(words):
        # Measure word width using Roman (approximate, but close enough for layout)
        word_w = c.stringWidth(word, FONT_NAME_ROMAN, font_size)
        space_w = c.stringWidth(' ', FONT_NAME_ROMAN, font_size)
        
        # Check if word fits on current line
        if x + word_w > text_width and x > 0:
            x = 0
            y += line_spacing
        
        # Place each character of the word
        for ch in word:
            ch_w = c.stringWidth(ch, FONT_NAME_ROMAN, font_size)
            chars.append((ch, x_start + x, y_start - y, ch_w))
            x += ch_w
        
        # Add space after word (except last)
        if w_idx < len(words) - 1:
            chars.append((' ', x_start + x, y_start - y, space_w))
            x += space_w
    
    return chars


def generate_pdf():
    print(f"Reading {INPUT_CSV}...")
    try:
        df = pd.read_csv(INPUT_CSV)
    except FileNotFoundError:
        print(f"Error: {INPUT_CSV} not found.")
        return

    # Extract unique names
    unique_names = sorted(df['Name'].dropna().unique())
    print(f"Found {len(unique_names)} unique names.")

    # Build the full text
    full_text = SEPARATOR.join(unique_names)

    # Register fonts
    register_fonts()

    # Page setup
    width, height = PAGE_SIZE

    # Match text area to image aspect ratio
    img = Image.open(IMAGE_MAP)
    img_w, img_h = img.size  # 842 x 596
    img_aspect = img_w / img_h  # ~1.41 (landscape)
    img.close()

    # Fit image aspect ratio into the page
    available_w = width - 2 * MARGIN_X
    available_h = height - 2 * MARGIN_Y

    # Try fitting by width first
    text_width = available_w
    text_height = text_width / img_aspect
    if text_height > available_h:
        # Fit by height instead
        text_height = available_h
        text_width = text_height * img_aspect

    # Center the text area on the page
    x_start = (width - text_width) / 2
    y_start = height - (height - text_height) / 2  # top of text area

    print(f"Page size: {width/mm:.0f}x{height/mm:.0f}mm")
    print(f"Text area: {text_width/mm:.1f}x{text_height/mm:.1f}mm (aspect {text_width/text_height:.3f})")
    print(f"Image aspect: {img_aspect:.3f}")

    # Create canvas
    c = canvas.Canvas(OUTPUT_PDF, pagesize=PAGE_SIZE)

    # Find optimal font size
    print("Calculating optimal font size...")
    font_size = find_font_size(c, full_text, text_width, text_height)
    print(f"Font size: {font_size:.1f}pt")

    # Layout all characters
    print("Laying out characters...")
    char_positions = layout_characters(c, full_text, font_size, text_width, text_height, x_start, y_start)
    print(f"Total characters: {len(char_positions)}")

    # Load brightness map scaled to text area
    print(f"Loading brightness map: {IMAGE_MAP}")
    brightness_map = load_brightness_map(IMAGE_MAP, text_width, text_height)
    map_h, map_w = brightness_map.shape
    print(f"Brightness map size: {map_w}x{map_h}")

    # Draw each character with the appropriate font
    print("Drawing characters...")
    for ch, cx, cy, ch_w in char_positions:
        if ch == ' ':
            continue  # skip drawing spaces

        # Map character position to brightness map pixel
        # cx is absolute x on page, cy is absolute y on page
        # Convert to relative position within text area
        rel_x = cx - x_start
        rel_y = y_start - cy  # invert y (PDF y goes up, image y goes down)

        # Clamp to map bounds
        px = int(min(max(rel_x, 0), map_w - 1))
        py = int(min(max(rel_y, 0), map_h - 1))

        brightness = brightness_map[py, px]
        font_name = get_font_for_brightness(brightness)

        c.setFont(font_name, font_size)
        c.drawString(cx, cy, ch)

    c.save()
    print(f"PDF generated: {OUTPUT_PDF}")


if __name__ == "__main__":
    generate_pdf()
