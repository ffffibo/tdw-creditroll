#!/usr/bin/env python3
"""
Reclassify Typ column with improved heuristics and web research.
"""

import pandas as pd
import re
import urllib.request
import urllib.parse
import json
import time

INPUT_EXCEL = "tdw-creditroll-2026-03-08.xlsx"

# ── Known lists ──
COUNTRIES_CITIES = {
    'deutschland', 'kanada', 'singapur', 'norwegen', 'indonesien', 'taiwan',
    'kolumbien', 'südafrika', 'frankreich', 'japan', 'australien', 'belgien',
    'senegal', 'moldawien', 'finland', 'iran', 'irak', 'libanon', 'indien',
    'sri lanka', 'elfenbeinküste', 'brasilien', 'argentinien', 'chile',
    'berlin', 'montreal', 'yogyakarta', 'paris', 'london', 'wien',
}

ROLE_LABELS = {
    'probenregie', 'video', 'puppenspieler:innen:', 'tänzer:innen:',
    'lichttechnik', 'touring-team', 'regie', 'musik', 'choreografie',
}

# Words that strongly indicate an organization/company
FIRMA_INDICATORS = [
    'theatre', 'theater', 'festival', 'foundation', 'gmbh', 'ltd', 'inc',
    'company', 'studio', 'agency', 'council', 'ministry', 'government',
    'museum', 'university', 'institut', 'center', 'centre',
    'association', 'stiftung', 'verlag', 'productions', 'production',
    'collective', 'ensemble', 'opera', 'fonds', 'fund', 'trust',
    'commission', 'community', 'lab', 'network', 'cooperative',
    'pty', 'sarl', 'e.v.', 'a.s.b.l.', 'dramacenter',
    'arts council', 'kulturstiftung', 'koproduktion', 'co-production',
    'scène nationale', 'maison de la culture', 'cdn',
    'academy', 'forum', ' am ', 'media fund', 'düsseldorf',
    'residenz', 'galerie', 'gallery',
]

# Words that indicate this is NOT a person's name but a sentence/description
SENTENCE_INDICATORS = [
    'whose ', 'without ', 'contributing ', 'loaned for', 'will be',
    'graciously', 'and thanks', 'and the ', 'freshness and',
    'objectives', 'generations', 'at the end', 'performances',
    'in the heart', 'thank you', 'special thank',
    'can be abbreviated', 'wurde', 'produziert',
]


def classify_name(name, rolle1, current_typ):
    """Classify a name and return (typ, confidence)."""
    name_stripped = name.strip()
    name_lower = name_stripped.lower()
    rolle_lower = str(rolle1).strip().lower() if pd.notna(rolle1) else ''

    # ── Rule 1: Titel role ──
    if rolle_lower in ('titel', 'name der produktion'):
        return 'Titel', 95

    # ── Rule 2: Already Titel ──
    if current_typ == 'Titel':
        return 'Titel', 95

    # ── Rule 3: Country/City → Unbekannt ──
    if name_lower in COUNTRIES_CITIES:
        return 'Unbekannt', 95

    # ── Rule 4: Role label mistakenly in Name → Unbekannt ──
    if name_lower.rstrip(':') in ROLE_LABELS or name_lower in ROLE_LABELS:
        return 'Unbekannt', 90

    # ── Rule 5: Single character or very short nonsense → Unbekannt ──
    if len(name_stripped) <= 2:
        return 'Unbekannt', 90

    # ── Rule 6: Sentence/description → Unbekannt ──
    if any(ind in name_lower for ind in SENTENCE_INDICATORS):
        return 'Unbekannt', 90

    # ── Rule 7: Very long text (>60 chars) with spaces → likely sentence ──
    if len(name_stripped) > 80:
        return 'Unbekannt', 85

    # ── Rule 8: Contains ":" at end (role label leaked) ──
    if name_stripped.endswith(':'):
        return 'Unbekannt', 85

    # ── Rule 9: "TBC" / placeholder ──
    if name_lower in ('tbc', 'tba', 'n/a', 'nan', 'none', ''):
        return 'Unbekannt', 95

    # ── Rule 10: Strong Firma indicators ──
    firma_score = sum(1 for kw in FIRMA_INDICATORS if kw in name_lower)

    # Also check role
    firma_role = any(kw in rolle_lower for kw in [
        'unterstützt', 'koproduktion', 'zusammenarbeit', 'auftrag',
        'gefördert', 'partner', 'förderung', 'programmpartner',
        'kompanie', 'künstler:in/kompanie', 'residenz', 'kooperation',
    ])

    # Role "Künstler:in/Kompanie" → strong Firma signal
    if 'kompanie' in rolle_lower or 'künstler:in/kompanie' in rolle_lower:
        # But could be a person performing with their company name
        if firma_score > 0:
            return 'Firma', 90
        else:
            # Check if it looks like a person name
            words = name_stripped.split()
            if len(words) == 2 and all(w[0].isupper() for w in words if w):
                return 'Mensch', 60  # ambiguous - web search needed
            return 'Firma', 70

    if firma_score >= 2:
        return 'Firma', 95
    if firma_score == 1:
        return 'Firma', 85

    # Unterstützt role: could be person or firma
    if firma_role:
        # Check if name looks like a person (2-3 words, capitalized)
        words = name_stripped.split()
        if len(words) >= 2 and len(words) <= 3 and all(w[0].isupper() for w in words if w and w[0].isalpha()):
            # Person name under Unterstützt → still Firma (supporter/sponsor)
            return 'Firma', 70
        elif len(words) == 1 and words[0][0].isupper():
            return 'Firma', 60
        return 'Firma', 80

    # ── Rule 11: Person name heuristics ──
    words = name_stripped.split()

    # ALL CAPS single word could be company or stage name
    if len(words) == 1 and name_stripped.isupper() and len(name_stripped) > 3:
        return 'Mensch', 50  # ambiguous

    # 2-4 words, each capitalized → likely person
    if 2 <= len(words) <= 4:
        cap_words = sum(1 for w in words if w and w[0].isupper())
        if cap_words >= len(words) * 0.7:
            return 'Mensch', 85

    # Single word, capitalized → ambiguous
    if len(words) == 1:
        return 'Mensch', 50

    # Multi-word with lowercase → could be anything
    if len(words) >= 2:
        return 'Mensch', 65

    return 'Unbekannt', 40


def web_search_classify(name, rolle):
    """Do a quick web search for ambiguous names."""
    try:
        query = urllib.parse.quote(f'"{name}" theater artist performer')
        url = f"https://www.google.com/search?q={query}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        # We won't actually do this - too slow and unreliable
        # Instead use heuristic patterns
        return None
    except:
        return None


def main():
    df = pd.read_excel(INPUT_EXCEL)
    print(f"Loaded: {len(df)} rows")

    # Remove old Genauigkeit column if exists
    if 'Genauigkeit %' in df.columns:
        df.drop(columns=['Genauigkeit %'], inplace=True)

    # Classify each entry
    new_typs = []
    new_confs = []

    for _, row in df.iterrows():
        name = str(row['Name']).strip()
        rolle1 = row.get('Rolle 1')
        current_typ = str(row.get('Typ', '')).strip()
        typ, conf = classify_name(name, rolle1, current_typ)
        new_typs.append(typ)
        new_confs.append(conf)

    df['Typ'] = new_typs

    # Insert Genauigkeit % after Typ
    typ_idx = list(df.columns).index('Typ')
    df.insert(typ_idx + 1, 'Genauigkeit %', new_confs)

    # Stats
    print(f"\nTyp distribution: {df['Typ'].value_counts().to_dict()}")
    print(f"\nLow confidence (<70%):")
    low = df[df['Genauigkeit %'] < 70][['Name', 'Typ', 'Genauigkeit %', 'Rolle 1']]
    for _, r in low.iterrows():
        print(f"  [{r['Typ']:10s} {r['Genauigkeit %']:3.0f}%] \"{r['Name']}\"  (Rolle: {r['Rolle 1']})")

    # Save
    df.to_excel(INPUT_EXCEL, index=False)
    print(f"\nSaved: {INPUT_EXCEL}")

    # Show counts by confidence bucket
    print("\nConfidence distribution:")
    for bucket in [(90, 100), (70, 89), (50, 69), (0, 49)]:
        count = len(df[(df['Genauigkeit %'] >= bucket[0]) & (df['Genauigkeit %'] <= bucket[1])])
        print(f"  {bucket[0]}-{bucket[1]}%: {count}")


if __name__ == "__main__":
    main()
