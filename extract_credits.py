import os
import json
import csv
import re
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

MASTER_JSON = "master_table-2026-02-20-15-42.json"
IN_DIR = "in"

# Roles to extract outside of Creditblock
EXTRA_ROLES = [
    {"search": "Name Produktion", "role": "Name der Produktion"},
    {"search": "Künstler:in / Kompanie", "role": "Künstler:in/Kompanie"},
    {"search": "Kurator:in", "role": "Kurator:in"}
]

def get_green_highlighted_text(cell):
    """
    Returns (cleaned_text, is_marked_green, original_text_with_newlines if needed)
    """
    text_content = ""
    is_marked = False
    
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            # Check for WD_COLOR_INDEX.BRIGHT_GREEN which is 4
            highlight = getattr(run.font, "highlight_color", None)
            if highlight == WD_COLOR_INDEX.BRIGHT_GREEN:
                is_marked = True
            text_content += run.text
        text_content += "\n"
        
    text_content = text_content.strip()
    return text_content, is_marked

def parse_docx(filepath, master_lookup):
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    filename = os.path.basename(filepath)
    # Match ID like A_018
    match = re.search(r'(A_\d{3})', filename)
    doc_id = match.group(1) if match else "UNKNOWN"
    
    # Base Info from master table
    master_info = master_lookup.get(doc_id, {})
    chron = master_info.get("Chronologische_Sortierung", "")
    titel = master_info.get("Titel", "")
    art = master_info.get("Art", "")
    zielgruppe = master_info.get("Zielgruppe", "")
    herkunftsland = master_info.get("Herkunftsland", "")  # Or we could extract it from docx if preferred

    doc = Document(filepath)
    results = []
    
    found_extra = {r["role"]: False for r in EXTRA_ROLES}
    is_credit_block = False
    
    for child in doc.element.body:
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            text = p.text.strip().lower()
            if "creditblock" in text:
                is_credit_block = True
            elif "biografien" in text or "kritiken" in text or "pressematerial" in text:
                is_credit_block = False
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            
            for row_idx, row in enumerate(table.rows):
                if len(row.cells) >= 2:
                    cell_left = row.cells[0]
                    
                    # Deduplicate cells correctly
                    if len(row.cells) > 1 and row.cells[1]._tc != cell_left._tc:
                        cell_right = row.cells[1]
                    elif len(row.cells) > 2 and row.cells[2]._tc != cell_left._tc:
                        cell_right = row.cells[2]
                    else:
                        continue # merged cell across columns
                        
                    left_text, left_marked = get_green_highlighted_text(cell_left)
                    right_text, right_marked = get_green_highlighted_text(cell_right)
                    
                    if not left_text and not right_text:
                        continue
                        
                    # Check for extra roles
                    matched_extra = False
                    for extra in EXTRA_ROLES:
                        if extra["search"].lower() in left_text.lower() and left_text.strip():
                            results.append({
                                "Markiert": "Ja" if (left_marked or right_marked) else "Nein",
                                "Name": right_text,
                                "Rolle": extra["role"],
                                "Original_Rolle": left_text,
                                "ID": doc_id,
                                "Titel": titel,
                                "Art": art,
                                "Zielgruppe": zielgruppe,
                                "Chronologische_Sortierung": chron,
                                "Herkunftsland": herkunftsland
                            })
                            found_extra[extra["role"]] = True
                            matched_extra = True
                            break
                            
                    # Focus on credit block
                    if is_credit_block and not matched_extra:
                        if left_text.lower() not in ["medium / datei", ""] and "credit" not in left_text.lower():
                            results.append({
                                "Markiert": "Ja" if (left_marked or right_marked) else "Nein",
                                "Name": right_text,
                                "Rolle": left_text, 
                                "Original_Rolle": left_text,
                                "ID": doc_id,
                                "Titel": titel,
                                "Art": art,
                                "Zielgruppe": zielgruppe,
                                "Chronologische_Sortierung": chron,
                                "Herkunftsland": herkunftsland
                            })
                            
            if is_credit_block:
                # To be safe, any table immediately following a credit block without texts like "Biografien" 
                # belongs to it, but standard template uses one table. We will leave is_credit_block True just in case
                # it's split.
                pass
                
    return results

def main():
    # Load Master Lookup
    master_lookup = {}
    if os.path.exists(MASTER_JSON):
        with open(MASTER_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
            for row in data:
                if "ID" in row:
                    master_lookup[row["ID"]] = row
                    
    all_credits = []
    
    if not os.path.exists(IN_DIR):
        print(f"Directory {IN_DIR} not found.")
        return

    for file in os.listdir(IN_DIR):
        if file.endswith(".docx") and not file.startswith("~$"):
            filepath = os.path.join(IN_DIR, file)
            print(f"Parsing: {file}")
            file_credits = parse_docx(filepath, master_lookup)
            all_credits.extend(file_credits)
            
    # Export to JSON
    with open("credits.json", "w", encoding="utf-8") as f:
        json.dump(all_credits, f, ensure_ascii=False, indent=2)
        
    # Export to CSV
    if all_credits:
        keys = ["Markiert", "Name", "Rolle", "Original_Rolle", "ID", "Titel", "Art", "Zielgruppe", "Chronologische_Sortierung", "Herkunftsland"]
        with open("credits.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            writer.writerows(all_credits)

    # Export to XLSX
    try:
        import pandas as pd
        df = pd.DataFrame(all_credits)
        df.to_excel("credits.xlsx", index=False)
        print("Generated credits.xlsx")
    except ImportError:
        print("pandas not installed, skipping XLSX export.")

if __name__ == "__main__":
    main()
