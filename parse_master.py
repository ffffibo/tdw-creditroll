import json
with open("master_table-2026-02-20-15-42.json", "r") as f:
    data = json.load(f)

for k in list(data[0].keys())[:20]:
    print(k)
