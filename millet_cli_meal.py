#!/usr/bin/env python3
"""
================================================================================
MILLET ALTERNATIVE FINDER — MEAL MODE (multi-dish x multi-condition)
================================================================================
Pick several dishes you'd normally eat AND several health conditions you have.
For each dish the tool finds the millet alternatives and ranks them by how well
a single choice satisfies ALL your conditions at once, then names one pick.

HOW THE "SATISFY ALL CONDITIONS" LOGIC WORKS
    1. Each condition contributes nutrient directions (fibre up, protein down...).
    2. Those are merged into one combined requirement across all your conditions.
       - If conditions AGREE on a nutrient, that direction is reinforced.
       - If two conditions DISAGREE (e.g. CKD wants protein DOWN, Osteoporosis
         wants protein UP), that nutrient is a CONFLICT and is reported — no
         single food can satisfy both, so it is excluded from the pass/fail test
         and flagged to the user.
    3. Every millet alternative is scored against EACH condition (helps / neutral
       / hurts), then ranked so the recommended one:
         (a) hurts the fewest conditions (ideally zero = "satisfies all"),
         (b) then helps the most,
         (c) then has the best overall nutrient score.

FILES REQUIRED (same folder as this script)
    Millet_Mapping_Final_1.xlsx                       (colour-coded mapping)
    Final Nutrition per 100g - Full Dataset.csv       (per-100g nutrition)
    155_Final.xlsx                                    (for Multi Millet blends)

RUN
    pip install pandas openpyxl        # once
    python millet_cli_meal.py          # interactive (multi-select)
    python millet_cli_meal.py --foods "Samosas,Crab Fried Rice,Bhel Puri" \
                              --diseases "diabetes,kidney,hypertension"
    python millet_cli_meal.py --list-foods
    python millet_cli_meal.py --list-diseases
================================================================================
"""
import os, sys, glob, re, difflib, argparse
import pandas as pd, numpy as np, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
NUTKEYS = ['kcal', 'carb', 'protein', 'fat', 'fibre']
NUTLABEL = {'kcal': 'Energy', 'carb': 'Carb', 'protein': 'Protein', 'fat': 'Fat', 'fibre': 'Fibre'}

def find(patterns, label):
    for pat in patterns:
        hits = glob.glob(os.path.join(HERE, pat))
        if hits:
            return hits[0]
    sys.exit("ERROR: could not find the %s here:\n  %s\n  looked for: %s\n  present: %s"
             % (label, HERE, ', '.join(patterns),
                [f for f in os.listdir(HERE) if f.lower().endswith(('.csv', '.xlsx'))]))

MAP_FILE = find(['*Mapping*Final*.xlsx', '*Millet_Mapping*.xlsx'], 'colour-coded mapping')
NUT_FILE = find(['*Nutrition*100*.csv', '*Nutrition*.csv', '*utrition*.csv'], 'nutrition CSV')
MIL_FILE = find(['*155*.xlsx', '*Final*.xlsx'], 'millet recipe file')

# ---------- nutrition ----------
def _num(v):
    s = str(v).strip().replace('–', '-')
    if s in ('', '-', 'nan') or 'source' in s.lower():
        return np.nan
    m = re.findall(r'\d+\.?\d*', s)
    return round(sum(map(float, m)) / len(m), 2) if m else np.nan

def load_nutrition():
    nut = pd.read_csv(NUT_FILE)
    for c, k in {'Energy (kcal/100g)': 'kcal', 'Carbohydrates (g/100g)': 'carb',
                 'Protein (g/100g)': 'protein', 'Fat (g/100g)': 'fat',
                 'Dietary Fibre (g/100g)': 'fibre'}.items():
        nut[k] = nut[c].map(_num)
    by_id = nut.set_index('ID')
    table = []
    for _, r in nut.iterrows():
        for nm in {str(r['Standard Ingredient']).lower(), str(r['Form / Variety']).lower()}:
            if nm and nm != 'nan':
                table.append((nm, {k: (None if pd.isna(r[k]) else float(r[k])) for k in NUTKEYS},
                              str(r['Standard Ingredient'])))
    return nut, by_id, table

NUT, NUT_BY_ID, NAME_TABLE = load_nutrition()

def grain_nutrition(name):
    if not isinstance(name, str) or not name.strip():
        return None
    nm = name.lower(); best = None; score = 0; label = None
    for cand, vals, lab in NAME_TABLE:
        sc = difflib.SequenceMatcher(None, nm, cand).ratio()
        if nm in cand or cand in nm:
            sc = max(sc, 0.9)
        if sc > score:
            score, best, label = sc, vals, lab
    return dict(best, matched_as=label) if score >= 0.6 else None

# ---------- millet nutrition (blends handled) ----------
BASE_MILLET = {'Sorghum': 'A005', 'Pearl Millet': 'A003', 'Finger Millet': 'A010',
               'Little Millet': 'A016', 'Kodo Millet': 'A017', 'Proso Millet': 'NEW-01',
               'Foxtail Millet': 'NEW-03', 'Barnyard Millet': 'NEW-04'}
NAME2TYPE = {'sorghum': 'Sorghum', 'finger millet': 'Finger Millet', 'pearl millet': 'Pearl Millet',
             'foxtail': 'Foxtail Millet', 'little millet': 'Little Millet', 'kodo millet': 'Kodo Millet',
             'proso millet': 'Proso Millet', 'barnyard millet': 'Barnyard Millet'}

def _base(t):
    r = NUT_BY_ID.loc[BASE_MILLET[t]]
    return {k: round(float(r[k]), 2) for k in NUTKEYS}

def load_millet():
    mil = pd.read_excel(MIL_FILE)
    return dict(zip(mil.recipe_name, mil.ingredients_raw)), dict(zip(mil.recipe_name, mil.millet_type))
RAW_OF, TYPE_OF = load_millet()

def millet_nutrition(recipe_name, millet_type):
    t = millet_type or TYPE_OF.get(recipe_name)
    if t and t != 'Multi Millet':
        return dict(_base(t), src=t)
    m = re.search(r'\(([^)]*)\)', str(RAW_OF.get(recipe_name, '')))
    comps = []
    if m:
        for nm, ty in NAME2TYPE.items():
            if nm in m.group(1).lower() and ty not in comps:
                comps.append(ty)
    if not comps:
        comps = ['Sorghum', 'Finger Millet', 'Pearl Millet', 'Foxtail Millet']
    return dict({k: round(sum(_base(c)[k] for c in comps) / len(comps), 2) for k in NUTKEYS},
                src=' + '.join(comps))

# ---------- colour-coded mapping ----------
COLOUR_STATUS = {'FF00FF00': 'ALTERNATIVE', 'FFFFE599': 'ALTERNATIVE', '00000000': 'ALTERNATIVE', None: 'ALTERNATIVE',
                 'FF4A86E8': 'RECOMMENDATION', 'FF00FFFF': 'RECOMMENDATION',
                 'FFFF0000': 'EXCLUDE', 'FFE06666': 'EXCLUDE', 'FFEA9999': 'EXCLUDE'}

def load_mapping():
    ws = openpyxl.load_workbook(MAP_FILE)['Top 3 Candidates']
    hdr = [c.value for c in ws[1]]; idx = {h: i for i, h in enumerate(hdr)}
    out = []
    for ri in range(2, ws.max_row + 1):
        rgb = getattr(ws.cell(ri, 1).fill.fgColor, 'rgb', None)
        status = COLOUR_STATUS.get(rgb, 'ALTERNATIVE')
        if status == 'EXCLUDE':
            continue
        row = {h: ws.cell(ri, idx[h] + 1).value for h in hdr}
        oid = row.get('original_recipe_id')
        out.append(dict(
            original_id=str(int(oid)) if isinstance(oid, (int, float)) and not pd.isna(oid) else str(oid),
            original=row.get('original_recipe'), food_family=row.get('food_family'),
            millet_recipe=row.get('millet_recipe'), millet_type=row.get('millet_type'),
            replaced=row.get('replaced_ingredient'), status=status))
    return out
MAPPING = load_mapping()

# ---------- disease rules ----------
DISEASES = {
    'Type 2 Diabetes':        {'fibre': 1, 'carb': -1, 'fat': -1},
    'High Cholesterol':       {'fibre': 1, 'fat': -1, 'kcal': -1},
    'Cardiovascular Disease': {'fibre': 1, 'fat': -1},
    'Hypertension':           {'fibre': 1, 'kcal': -1},
    'Chronic Kidney Disease': {'protein': -1},
    'MASLD / Fatty Liver':    {'kcal': -1, 'carb': -1, 'fat': -1, 'fibre': 1},
    'Obesity / Overweight':   {'kcal': -1, 'fibre': 1},
    'Gout':                   {'carb': -1, 'fat': -1, 'fibre': 1, 'protein': -1},
    'Osteoporosis':           {'protein': 1},
    'Dental Caries':          {},
}

def resolve_disease(text):
    t = text.strip().lower()
    for name in DISEASES:
        if t in name.lower() or name.lower().startswith(t):
            return name
    hits = difflib.get_close_matches(text, list(DISEASES), n=1, cutoff=0.4)
    return hits[0] if hits else None

# ---------- combined requirement + conflict detection ----------
def combine_requirements(diseases):
    """Merge nutrient directions across the selected conditions.
       Returns (agreed {nutrient:dir}, conflicts [nutrient...])."""
    wants = {}   # nutrient -> set of directions requested
    for d in diseases:
        for nut, dir in DISEASES[d].items():
            wants.setdefault(nut, set()).add(dir)
    agreed, conflicts = {}, []
    for nut, dirs in wants.items():
        if len(dirs) == 1:
            agreed[nut] = next(iter(dirs))
        else:
            conflicts.append(nut)       # both +1 and -1 requested
    return agreed, conflicts

# ---------- per-disease verdict for one millet pairing ----------
def verdict_for(pair, disease):
    """Return 'help' / 'neutral' / 'hurt' / 'n/a' for this pairing under one disease."""
    rule = DISEASES[disease]
    if not rule:
        return 'n/a'
    gn = grain_nutrition(pair['replaced']) if pair['replaced'] else None
    if gn is None:
        return 'n/a'
    mn = millet_nutrition(pair['millet_recipe'], pair['millet_type'])
    score = 0
    for nut, dir in rule.items():
        if mn.get(nut) is None or gn.get(nut) is None:
            continue
        dv = mn[nut] - gn[nut]
        score += 1 if dv * dir > 0 else (-1 if dv * dir < 0 else 0)
    return 'help' if score > 0 else ('hurt' if score < 0 else 'neutral')

def deltas(pair):
    gn = grain_nutrition(pair['replaced']) if pair['replaced'] else None
    if gn is None:
        return None
    mn = millet_nutrition(pair['millet_recipe'], pair['millet_type'])
    return {k: (None if (mn.get(k) is None or gn.get(k) is None) else round(mn[k] - gn[k], 1)) for k in NUTKEYS}, mn, gn

# ---------- rank a dish's alternatives across all selected diseases ----------
def rank_alternatives(pairs, diseases):
    scored = []
    for p in pairs:
        verds = {d: verdict_for(p, d) for d in diseases}
        helps = sum(1 for v in verds.values() if v == 'help')
        hurts = sum(1 for v in verds.values() if v == 'hurt')
        na = sum(1 for v in verds.values() if v == 'n/a')
        # total nutrient score across computable diseases (tie-breaker)
        total = 0
        for d in diseases:
            rule = DISEASES[d]; dl = deltas(p)
            if rule and dl:
                dd = dl[0]
                for nut, dir in rule.items():
                    if dd.get(nut) is None: continue
                    total += 1 if dd[nut] * dir > 0 else (-1 if dd[nut] * dir < 0 else 0)
        scored.append(dict(pair=p, verdicts=verds, helps=helps, hurts=hurts, na=na,
                           satisfies_all=(hurts == 0), total=total))
    # rank: fewest hurts, then most helps, then best total score, alternatives before recommendations
    scored.sort(key=lambda s: (s['hurts'], -s['helps'], -s['total'],
                               0 if s['pair']['status'] == 'ALTERNATIVE' else 1))
    # one row per millet recipe: keep its best-ranked occurrence (drops mapping duplicates)
    seen, deduped = set(), []
    for s in scored:
        nm = s['pair']['millet_recipe']
        if nm in seen:
            continue
        seen.add(nm); deduped.append(s)
    return deduped

# ---------- CLI ----------
def list_foods():
    for n in sorted({p['original'] for p in MAPPING if p['original']}):
        print(' ', n)

def list_diseases():
    for d in DISEASES:
        print(' ', d, '' if DISEASES[d] else '  (not computable from current nutrients)')

MARK = {'help': 'HELP', 'hurt': 'HURT', 'neutral': 'even', 'n/a': ' -- '}

def run(foods, diseases):
    dnames, unknown_d = [], []
    for d in foods_split(diseases):
        r = resolve_disease(d)
        (dnames if r else unknown_d).append(r or d)
    dnames = [d for d in dict.fromkeys(dnames) if d]
    if unknown_d:
        print("Ignoring unknown conditions:", ', '.join(unknown_d))
    if not dnames:
        print("No valid conditions. Try --list-diseases"); return

    agreed, conflicts = combine_requirements(dnames)
    print('\n' + '#' * 72)
    print('YOUR CONDITIONS :', ', '.join(dnames))
    print('COMBINED TARGET :', ', '.join('%s %s' % (NUTLABEL[n], '↑' if d > 0 else '↓')
                                         for n, d in agreed.items()) or '(none computable)')
    if conflicts:
        print('CONFLICT        : ' + ', '.join(NUTLABEL[c] for c in conflicts) +
              ' — your conditions disagree here, so no single food can satisfy all; this nutrient is not used for pass/fail.')
    print('#' * 72)

    for food in foods_split(foods):
        mine = [p for p in MAPPING if p['original'] and food.lower() in p['original'].lower()]
        if not mine:
            print('\n>> "%s": no dish matches. (try --list-foods)' % food); continue
        target = sorted({p['original'] for p in mine})[0]
        mine = [p for p in mine if p['original'] == target]
        ranked = rank_alternatives(mine, dnames)

        print('\n' + '=' * 72)
        print('DISH: %s   (%d millet option%s)' % (target, len(ranked), '' if len(ranked) == 1 else 's'))
        print('=' * 72)
        # header row of condition initials
        print('   %-34s %s  verdict' % ('millet option', '  '.join('%-5s' % d.split()[0][:5] for d in dnames)))
        best = ranked[0]
        # a confident pick satisfies every condition AND helps at least one
        confident = best['satisfies_all'] and best['helps'] > 0
        for s in ranked:
            p = s['pair']
            cells = '  '.join('%-5s' % MARK[s['verdicts'][d]] for d in dnames)
            if s['satisfies_all']:
                tag = 'satisfies all' if s['helps'] > 0 else 'safe, no gain'
            else:
                tag = 'hurts %d' % s['hurts']
            if s is best:
                star = ' <== PICK' if confident else ' <== best trade-off'
            else:
                star = ''
            kind = '' if p['status'] == 'ALTERNATIVE' else ' (rec)'
            print('   %-34s %s  %s%s' % ((p['millet_recipe'][:33] + kind), cells, tag, star))
        # explain the pick
        bp = best['pair']
        if confident:
            dd = (deltas(bp)[0] if bp['replaced'] else None)
            highs = [NUTLABEL[n] for n in NUTKEYS if (dd and n in agreed and dd.get(n) is not None and dd[n] * agreed[n] > 0)] if dd else []
            print('   -> PICK: %s (satisfies all your conditions%s)'
                  % (bp['millet_recipe'], '; better on ' + ', '.join(highs) if highs else ''))
        elif best['satisfies_all']:
            print('   -> Best: %s is safe for all your conditions but shows no clear nutritional gain over the original.' % bp['millet_recipe'])
        else:
            print('   -> No option satisfies every condition; best trade-off: %s (helps %d, hurts %d).'
                  % (bp['millet_recipe'], best['helps'], best['hurts']))
    print()

def foods_split(s):
    if isinstance(s, list):
        return s
    return [x.strip() for x in str(s).split(',') if x.strip()]

def interactive():
    print('\nMILLET MEAL MODE  (multi-dish x multi-condition)\n')
    print('Conditions available:'); list_diseases()
    d = input('\nEnter your conditions (comma-separated): ').strip()
    print('\nType dish names (comma-separated). Partial names ok; --list-foods for the full list.')
    f = input('Your dishes: ').strip()
    if f and d:
        run(f, d)

def main():
    ap = argparse.ArgumentParser(description='Millet meal recommender (multi-dish x multi-condition)')
    ap.add_argument('--foods', help='comma-separated dish names')
    ap.add_argument('--diseases', help='comma-separated conditions')
    ap.add_argument('--list-foods', action='store_true')
    ap.add_argument('--list-diseases', action='store_true')
    a = ap.parse_args()
    print('mapping: %s | nutrition: %s | millets: %s' %
          (os.path.basename(MAP_FILE), os.path.basename(NUT_FILE), os.path.basename(MIL_FILE)))
    print('%d usable pairings loaded.' % len(MAPPING))
    if a.list_foods:
        list_foods()
    elif a.list_diseases:
        list_diseases()
    elif a.foods and a.diseases:
        run(a.foods, a.diseases)
    else:
        interactive()

if __name__ == '__main__':
    main()
