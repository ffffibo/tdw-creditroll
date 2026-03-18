from extract_credits import parse_docx, MASTER_JSON
import json, os

master_lookup = {}
if os.path.exists(MASTER_JSON):
    with open(MASTER_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
        for row in data:
            if "ID" in row:
                master_lookup[row["ID"]] = row

res = parse_docx("in/A_018_Water_Words;_Parole_d'eau_Dossier.docx", master_lookup)
for r in res:
    print(f"Rolle: {r['Rolle'][:20]} | Name: {r['Name'][:30]} | Markiert: {r['Markiert']}")
