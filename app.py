#!/usr/bin/env python3
"""
Credit PDF Generator — Flask Web Application
Provides a GUI for configuring and generating credit PDFs
with brightness-map-based font weight mapping.
"""

import json
import os
import glob
import unicodedata
from datetime import datetime
import zipfile
import io

import pandas as pd
import numpy as np
from PIL import Image
from flask import Flask, render_template, request, jsonify, send_file
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import base64

# ── App Setup ──────────────────────────────────────────────────────────────
APP_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)

# ── Font Names ─────────────────────────────────────────────────────────────
FONT_NAME_ROMAN = "NHGDisplayRoman"
FONT_NAME_MEDIUM = "NHGDisplayMedium"
FONT_NAME_BOLD = "NHGDisplayBold"
FONT_NAME_FALLBACK = "FallbackUnicode"

FALLBACK_FONT_PATHS = [
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
]


# ── Unicode Helpers ────────────────────────────────────────────────────────
def strip_accents_for_sort(text):
    """Normalize text to remove diacritics for alphabetical sorting. Preserves original case-insensitive base characters."""
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII').lower()

def has_non_latin_chars(text):
    for ch in str(text):
        if not ch.isalpha():
            continue
        cp = ord(ch)
        if 0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF:
            return True
        if 0x20000 <= cp <= 0x2A6DF:
            return True
        if 0x3040 <= cp <= 0x309F or 0x30A0 <= cp <= 0x30FF:
            return True
        if 0xAC00 <= cp <= 0xD7AF or 0x1100 <= cp <= 0x11FF:
            return True
        if 0x0600 <= cp <= 0x06FF or 0x0750 <= cp <= 0x077F:
            return True
        if 0xFB50 <= cp <= 0xFDFF or 0xFE70 <= cp <= 0xFEFF:
            return True
        if 0x0E00 <= cp <= 0x0E7F:
            return True
        if 0x0900 <= cp <= 0x097F or 0x0980 <= cp <= 0x09FF:
            return True
        if 0x3000 <= cp <= 0x303F:
            return True
    return False


# ── PDF Generation ─────────────────────────────────────────────────────────
    has_fallback = False
    for fpath in FALLBACK_FONT_PATHS:
        if os.path.exists(fpath):
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME_FALLBACK, fpath))
                has_fallback = True
                break
            except Exception:
                pass
    return has_fallback


def load_brightness_map(image_source, target_w, target_h):
    img = Image.open(image_source).convert("L")
    img_resized = img.resize((int(target_w), int(target_h)), Image.LANCZOS)
    return np.array(img_resized)


def get_font_for_brightness(brightness, cfg):
    if brightness >= cfg["threshold_white"]:
        return FONT_NAME_ROMAN
    elif brightness <= cfg["threshold_black"]:
        return FONT_NAME_BOLD
    else:
        return FONT_NAME_MEDIUM


def get_exact_word_width(c, word, font_size, x_start, y_start, y_current, map_w, map_h, brightness_map, cfg, has_fallback=False):
    """
    Simulates drawing a word to calculate its precise width 
    based on the brightness map at its projected X/Y position.
    """
    word_width = 0
    tracking_offset = cfg.get("tracking_em", 0.0) * font_size
    current_x = x_start
    
    for ch in word:
        rel_x = current_x - x_start
        rel_y = y_start - y_current
        
        px = int(min(max(rel_x, 0), map_w - 1))
        py = int(min(max(rel_y, 0), map_h - 1))
        
        brightness = brightness_map[py, px]
        font_name = get_font_for_brightness(brightness, cfg)
        if has_fallback and has_non_latin_chars(ch):
            font_name = FONT_NAME_FALLBACK
            
        ch_w = c.stringWidth(ch, font_name, font_size) + tracking_offset
        word_width += ch_w
        current_x += ch_w
        
    return word_width

def get_exact_space_width(c, font_size, x_start, y_start, y_current, map_w, map_h, brightness_map, cfg):
    """Calculates the width of a single space character with tracking."""
    rel_y = y_start - y_current
    py = int(min(max(rel_y, 0), map_h - 1))
    # Approximation: use the X coordinate for the start of the space
    px = int(min(max(0, 0), map_w - 1)) 
    
    brightness = brightness_map[py, px]
    font_name = get_font_for_brightness(brightness, cfg)
    return c.stringWidth(" ", font_name, font_size) + cfg.get("tracking_em", 0.0) * font_size


def find_font_size(c, text, text_width, text_height, cfg, brightness_map, has_fallback=False):
    line_spacing_factor = cfg["line_spacing_factor"]
    min_size = cfg["min_font_size"]
    max_size = cfg["max_font_size"]
    map_h, map_w = brightness_map.shape

    def text_fits(font_size):
        leading = font_size * line_spacing_factor
        words = text.split(" ")
        x = 0
        lines = 1
        y_current = text_height # we simulate from top down, where top is text_height 
        
        space_w = get_exact_space_width(c, font_size, 0, text_height, y_current, map_w, map_h, brightness_map, cfg)

        for word in words:
            word_w = get_exact_word_width(c, word, font_size, x, text_height, y_current, map_w, map_h, brightness_map, cfg, has_fallback)
            
            if x + word_w > text_width and x > 0:
                lines += 1
                x = 0
                y_current -= leading
                # Must recalculate word width on the new line because brightness changed
                word_w = get_exact_word_width(c, word, font_size, x, text_height, y_current, map_w, map_h, brightness_map, cfg, has_fallback)
                space_w = get_exact_space_width(c, font_size, 0, text_height, y_current, map_w, map_h, brightness_map, cfg)

            x += word_w + space_w
            
        return lines * leading <= text_height

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


def split_into_lines(c, text, font_size, text_width, text_height, brightness_map, cfg, has_fallback=False):
    words = text.split(" ")
    lines = [[]]
    x = 0
    y_current = text_height
    leading = font_size * cfg["line_spacing_factor"]
    map_h, map_w = brightness_map.shape

    for word in words:
        space_w = get_exact_space_width(c, font_size, x, text_height, y_current, map_w, map_h, brightness_map, cfg) if x > 0 else 0
        word_w = get_exact_word_width(c, word, font_size, x + space_w, text_height, y_current, map_w, map_h, brightness_map, cfg, has_fallback)
        
        if x + space_w + word_w > text_width and x > 0:
            lines.append([])
            x = 0
            y_current -= leading
            word_w = get_exact_word_width(c, word, font_size, x, text_height, y_current, map_w, map_h, brightness_map, cfg, has_fallback)
            
        lines[-1].append(word)
        x += word_w + (get_exact_space_width(c, font_size, x + word_w, text_height, y_current, map_w, map_h, brightness_map, cfg) if len(lines[-1]) > 0 else 0)
        
    return lines


def create_pdf(names, image_source, cfg, has_fallback=False):
    separator = cfg.get("separator", " / ")
    text_before = cfg.get("text_before", "")
    text_after = cfg.get("text_after", "")

    parts = []
    if text_before.strip():
        parts.append(text_before.strip())
    parts.append(separator.join(names))
    if text_after.strip():
        parts.append(text_after.strip())
    full_text = separator.join(parts) if len(parts) > 1 else parts[0]

    page_w = float(cfg.get("page_width_mm", 295)) * mm
    page_h = float(cfg.get("page_height_mm", 325)) * mm
    margin_x = float(cfg.get("margin_x_mm", 10)) * mm
    margin_y = float(cfg.get("margin_y_mm", 10)) * mm
    text_width = page_w - 2 * margin_x
    text_height = page_h - 2 * margin_y
    x_start = margin_x
    y_start = page_h - margin_y

    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=(page_w, page_h))
    
    brightness_map = load_brightness_map(image_source, text_width, text_height)
    map_h, map_w = brightness_map.shape

    font_size = find_font_size(c, full_text, text_width, text_height, cfg, brightness_map, has_fallback)
    lines = split_into_lines(c, full_text, font_size, text_width, text_height, brightness_map, cfg, has_fallback)

    line_spacing = font_size * cfg["line_spacing_factor"]
    justify = cfg["justify"]
    tracking_offset = cfg.get("tracking_em", 0.0) * font_size
    total_chars = 0

    for line_idx, words in enumerate(lines):
        y = y_start - line_idx * line_spacing
        x = x_start
        line_text = " ".join(words)

        extra_space_per_gap = 0
        # Only justify lines that aren't the very last line
        if justify and line_idx < len(lines) - 1 and len(words) > 1:
            natural_width = 0
            for ch in line_text:
                rel_x_est = natural_width
                rel_y_est = y_start - y
                px_est = int(min(max(rel_x_est, 0), map_w - 1))
                py_est = int(min(max(rel_y_est, 0), map_h - 1))
                br = brightness_map[py_est, px_est]
                fn = get_font_for_brightness(br, cfg)
                if has_fallback and has_non_latin_chars(ch):
                    fn = FONT_NAME_FALLBACK
                natural_width += c.stringWidth(ch, fn, font_size) + tracking_offset
            
            space_count = line_text.count(" ")
            if space_count > 0:
                extra_space_per_gap = (text_width - natural_width) / space_count

        for ch in line_text:
            rel_x = x - x_start
            rel_y = y_start - y
            px = int(min(max(rel_x, 0), map_w - 1))
            py = int(min(max(rel_y, 0), map_h - 1))

            brightness = brightness_map[py, px]
            font_name = get_font_for_brightness(brightness, cfg)
            if has_fallback and has_non_latin_chars(ch):
                font_name = FONT_NAME_FALLBACK

            ch_w = c.stringWidth(ch, font_name, font_size) + tracking_offset

            if ch != " ":
                c.setFont(font_name, font_size)
                c.drawString(x, y, ch)
                total_chars += 1

            x += ch_w
            if ch == " " and justify and line_idx < len(lines) - 1:
                x += extra_space_per_gap

    c.save()
    pdf_buffer.seek(0)
    
    return {
        "buffer": pdf_buffer,
        "font_size": round(font_size, 1),
        "total_lines": len(lines),
        "total_chars": total_chars,
        "name_count": len(names),
    }


# ── Routes ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/load-excel", methods=["POST"])
def load_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
        
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    try:
        df = pd.read_excel(file.stream)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 400

    names = []
    for idx, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        if not name or name == "nan":
            continue
            
        row_id = f"row_{idx}_{name}"
        verwenden = row.get("Verwenden", False)
        is_verwenden = verwenden is True or str(verwenden).strip().upper() in ("WAHR", "TRUE")
        markiert = str(row.get("Markiert", "")).strip().lower()
        typ = str(row.get("Typ", ""))
        is_non_latin = has_non_latin_chars(name)

        names.append({
            "id": row_id,
            "name": name,
            "verwenden": is_verwenden,
            "markiert": markiert == "ja",
            "typ": typ,
            "non_latin": is_non_latin,
            "excluded": False,  # Config handling is now fully clienside
        })

    return jsonify({
        "filename": file.filename,
        "total_rows": len(df),
        "names": names,
    })


@app.route("/api/generate", methods=["POST"])
def generate():
    settings_json = request.form.get("settings", "{}")
    try:
        cfg = json.loads(settings_json)
    except Exception:
        return jsonify({"error": "Invalid settings JSON"}), 400

    selected_names = cfg.get("selected_names", [])
    if not selected_names:
        return jsonify({"error": "No names selected"}), 400
        
    font_dir = cfg.get("font_dir", os.path.join(APP_DIR, "neue-haas-grotesk-display-pro"))
    if not os.path.isdir(font_dir):
        return jsonify({"error": f"Font directory not found: {font_dir}"}), 400
        
    image_source = None
    if "image" in request.files and request.files["image"].filename != "":
        image_source = request.files["image"].stream
    else:
        # Fallback to bundled standard image
        default_image = os.path.join(APP_DIR, "tdw-gallpeters.png")
        if os.path.exists(default_image):
            image_source = default_image
        else:
            return jsonify({"error": "No image uploaded and default missing"}), 400

    try:
        has_fallback = register_fonts(font_dir)
    except Exception as e:
        return jsonify({"error": f"Font registration failed: {str(e)}"}), 500

    today = datetime.now().strftime("%Y-%m-%d")
    variant = cfg.get("variant", "both")
    
    latin_names = sorted([n for n in selected_names if not has_non_latin_chars(n)], key=strip_accents_for_sort)
    all_names = sorted(selected_names, key=strip_accents_for_sort)
    non_latin_names = [n for n in selected_names if has_non_latin_chars(n)]
    
    results = []
    
    try:
        if variant in ("latin", "both") and latin_names:
            res = create_pdf(latin_names, image_source, cfg, has_fallback=False)
            # recreate image_source since stream might have been exhausted
            if hasattr(image_source, "seek"): 
                image_source.seek(0)
            res["filename"] = f"tdw-creditroll-{today}-latin.pdf"
            res["variant"] = "Latin only"
            results.append(res)
            
        if variant in ("all", "both") and all_names:
            res = create_pdf(all_names, image_source, cfg, has_fallback=has_fallback)
            res["filename"] = f"tdw-creditroll-{today}-all.pdf"
            res["variant"] = "All scripts"
            results.append(res)
    except Exception as e:
        return jsonify({"error": f"PDF generation failed: {str(e)}"}), 500

    # If only 1 PDF was generated, stream it directly
    if len(results) == 1:
        return send_file(
            results[0]["buffer"], 
            mimetype="application/pdf", 
            as_attachment=True, 
            download_name=results[0]["filename"]
        )
        
    # If multiple, zip them
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            zf.writestr(r["filename"], r["buffer"].getvalue())
            
    zip_buffer.seek(0)
    return send_file(
        zip_buffer, 
        mimetype="application/zip", 
        as_attachment=True, 
        download_name=f"tdw-creditrolls-{today}.zip"
    )

# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"Starting Credit PDF Generator on http://localhost:5001")
    app.run(debug=False, port=5001, host="0.0.0.0")
