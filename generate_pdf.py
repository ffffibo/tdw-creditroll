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
import glob

from datetime import datetime

# ── Configuration ──────────────────────────────────────────────────────────
INPUT_EXCEL = "tdw-creditroll-2026-03-03.xlsx"
TODAY = datetime.now().strftime("%Y-%m-%d")
OUTPUT_BASE_ALL = f"tdw-creditroll-{TODAY}-all"
OUTPUT_BASE_MARKED = f"tdw-creditroll-{TODAY}-marked"
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
JUSTIFY = True  # True = Blocksatz (justified), False = linksbündig (left-aligned)

def get_next_output_pdf(output_base):
    """Find the next available counter for the PDF filename."""
    existing = glob.glob(f"{output_base}-*.pdf")
    if not existing:
        return f"{output_base}-01.pdf"
    counters = []
    for f in existing:
        base = os.path.splitext(f)[0]
        suffix = base.replace(output_base + "-", "")
        try:
            counters.append(int(suffix))
        except ValueError:
            pass
    next_num = max(counters) + 1 if counters else 1
    return f"{output_base}-{next_num:02d}.pdf"


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
            word_w = c.stringWidth(word + ' ', FONT_NAME_BOLD, font_size)
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


def split_into_lines(c, text, font_size, text_width):
    """
    Pass 1: Determine which words go on which line using Bold metrics
    (widest font) to guarantee no overflow.
    Returns a list of lines, where each line is a list of words.
    """
    words = text.split(' ')
    lines = [[]]
    x = 0
    for word in words:
        word_w = c.stringWidth(word + ' ', FONT_NAME_BOLD, font_size)
        if x + word_w > text_width and x > 0:
            lines.append([])
            x = 0
        lines[-1].append(word)
        x += word_w
    return lines


def create_pdf(unique_names, output_base):
    print(f"--- Generating PDF for {output_base} ---")
    print(f"Found {len(unique_names)} unique names.")

    # Build the full text
    full_text = SEPARATOR.join(unique_names)

    # Register fonts
    register_fonts()

    # Page setup — image now matches page aspect ratio, use full area
    width, height = PAGE_SIZE
    text_width = width - 2 * MARGIN_X
    text_height = height - 2 * MARGIN_Y
    x_start = MARGIN_X
    y_start = height - MARGIN_Y  # top of text area

    print(f"Page size: {width/mm:.0f}x{height/mm:.0f}mm")
    print(f"Text area: {text_width/mm:.1f}x{text_height/mm:.1f}mm")

    # Determine output filename
    output_pdf = get_next_output_pdf(output_base)
    print(f"Output: {output_pdf}")

    # Create canvas
    c = canvas.Canvas(output_pdf, pagesize=PAGE_SIZE)

    # Find optimal font size
    print("Calculating optimal font size...")
    font_size = find_font_size(c, full_text, text_width, text_height)
    print(f"Font size: {font_size:.1f}pt")

    # Load brightness map scaled to text area
    print(f"Loading brightness map: {IMAGE_MAP}")
    brightness_map = load_brightness_map(IMAGE_MAP, text_width, text_height)
    map_h, map_w = brightness_map.shape
    print(f"Brightness map size: {map_w}x{map_h}")

    # Pass 1: determine line breaks using Bold widths
    print("Pass 1: Line breaking...")
    lines = split_into_lines(c, full_text, font_size, text_width)
    print(f"Total lines: {len(lines)}")

    # Pass 2: position and draw each character using actual font widths
    print(f"Pass 2: Drawing characters ({'justified' if JUSTIFY else 'left-aligned'})...")
    line_spacing = font_size * LINE_SPACING_FACTOR
    total_chars = 0

    for line_idx, words in enumerate(lines):
        y = y_start - line_idx * line_spacing
        x = x_start

        line_text = ' '.join(words)

        # Calculate extra space for justification
        extra_space_per_gap = 0
        if JUSTIFY and line_idx < len(lines) - 1 and len(words) > 1:
            # Measure natural line width using actual fonts
            natural_width = 0
            for ch in line_text:
                rel_x_est = natural_width
                rel_y_est = y_start - y
                px_est = int(min(max(rel_x_est, 0), map_w - 1))
                py_est = int(min(max(rel_y_est, 0), map_h - 1))
                br = brightness_map[py_est, px_est]
                fn = get_font_for_brightness(br)
                natural_width += c.stringWidth(ch, fn, font_size)
            space_count = line_text.count(' ')
            if space_count > 0:
                extra_space_per_gap = (text_width - natural_width) / space_count

        for ch in line_text:
            # Determine position relative to text area for brightness lookup
            rel_x = x - x_start
            rel_y = y_start - y  # invert for image coords

            # Clamp to map bounds
            px = int(min(max(rel_x, 0), map_w - 1))
            py = int(min(max(rel_y, 0), map_h - 1))

            brightness = brightness_map[py, px]
            font_name = get_font_for_brightness(brightness)

            # Get character width using the ACTUAL font for this character
            ch_w = c.stringWidth(ch, font_name, font_size)

            if ch != ' ':
                c.setFont(font_name, font_size)
                c.drawString(x, y, ch)
                total_chars += 1

            # Advance x by the actual font's character width
            x += ch_w

            # Add extra justification space after spaces
            if ch == ' ' and JUSTIFY:
                x += extra_space_per_gap

    print(f"Total visible characters drawn: {total_chars}")
    c.save()
    print(f"PDF generated: {output_pdf}")


def main():
    print(f"Reading {INPUT_EXCEL}...")
    try:
        df = pd.read_excel(INPUT_EXCEL)
    except FileNotFoundError:
        print(f"Error: {INPUT_EXCEL} not found.")
        return

    # 1. All unique names
    all_names = sorted(df['Name'].dropna().astype(str).unique())
    
    # 2. Marked unique names
    # Ensure 'Markiert' column exists
    if 'Markiert' in df.columns:
        marked_df = df[df['Markiert'].astype(str).str.strip().str.lower() == 'ja']
        marked_names = sorted(marked_df['Name'].dropna().astype(str).unique())
    else:
        print("Warning: Column 'Markiert' not found. Marked PDF will be empty.")
        marked_names = []

    if all_names:
        create_pdf(all_names, OUTPUT_BASE_ALL)
    
    if marked_names:
        create_pdf(marked_names, OUTPUT_BASE_MARKED)
    else:
        print("No marked names found to generate the second PDF.")

if __name__ == "__main__":
    main()
