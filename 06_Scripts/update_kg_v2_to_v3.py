# -*- coding: utf-8 -*-
"""KG v2 -> v3: add 15 BB papers + their authors from BB JSON."""
import sys, io, json, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
KG_PATH = os.path.join(ROOT, '06_Scripts', 'author_kg_data', 'author_kg_v2.json')
BB_PATH = os.path.join(ROOT, '06_Scripts', 'canonical_entries_BB_07062026.json')
OUT = os.path.join(ROOT, '06_Scripts', 'author_kg_data', 'author_kg_v3.json')

with open(KG_PATH, 'r', encoding='utf-8') as f:
    kg = json.load(f)
with open(BB_PATH, 'r', encoding='utf-8') as f:
    bb_data = json.load(f)

bb_entries = bb_data if isinstance(bb_data, list) else bb_data.get('entries', list(bb_data.values()))
print(f'BB entries to process: {len(bb_entries)}')

existing_ids = {n['id'] for n in kg['nodes']}
new_nodes, new_edges = [], []

for entry in bb_entries:
    bb_id = entry.get('id', '')
    desc = entry.get('descriptor', '')
    # Parse author from descriptor like "Xu_2021_China"
    parts = desc.split('_')
    if len(parts) >= 2:
        surname = parts[0]
        year = next((p for p in parts if p.isdigit() and len(p) == 4), '')
        country = parts[-1] if not parts[-1].isdigit() else ''
    else:
        surname = desc
        year = ''
        country = ''

    au_id = f'AU_{surname.upper()}_BB'
    pa_id = f'PA_BB_{bb_id}'

    if au_id not in existing_ids:
        new_nodes.append({
            'id': au_id,
            'type': 'Author',
            'label': surname,
            'surname': surname,
            'is_vietnamese': str(surname in ['Nguyen', 'Pham', 'Tran', 'Le',
                                            'Vu', 'Hoang', 'Ngo', 'Bui',
                                            'Dao', 'Do']),
            'n_papers': 1,
            'country_hint': country,
            'source': 'BB collection (11-bai-ban-dau-va-mo-rong)',
        })
        existing_ids.add(au_id)

    new_nodes.append({
        'id': pa_id,
        'type': 'Paper',
        'label': f'{surname} {year} — BB collection',
        'authors': [au_id],
        'year': int(year) if year.isdigit() else None,
        'pdf_path': entry.get('pdf_path', ''),
        'topic': entry.get('topic', ''),
        'bb_id': bb_id,
        'set': '11-bai-ban-dau-va-mo-rong',
    })
    existing_ids.add(pa_id)
    new_edges.append({'type': 'AUTHORED_BY', 'source': pa_id, 'target': au_id})

# Add Karasu QT074+QT075 papers if not already (they're in v2)
# Add Small & Blanc 2021 paper if not (it IS in v2 as PA_SMALL_BLANC_2021)

kg['nodes'].extend(new_nodes)
kg['edges'].extend(new_edges)

# Update meta
kg['meta']['last_updated'] = '2026-06-07'
kg['meta']['n_nodes'] = len(kg['nodes'])
kg['meta']['n_edges'] = len(kg['edges'])
kg['meta']['node_types'] = {
    'Author': sum(1 for n in kg['nodes'] if n.get('type') == 'Author'),
    'Paper': sum(1 for n in kg['nodes'] if n.get('type') == 'Paper'),
    'Country': sum(1 for n in kg['nodes'] if n.get('type') == 'Country'),
    'Topic': sum(1 for n in kg['nodes'] if n.get('type') == 'Topic'),
}
kg['meta'].setdefault('v3_changelog', []).append(
    '2026-06-07: Added 15 BB papers (11-bai-ban-dau-va-mo-rong) + their authors')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(kg, f, ensure_ascii=False, indent=2)

print(f'Saved KG v3: {OUT}')
print(f'New nodes: {len(new_nodes)}, new edges: {len(new_edges)}')
print(f'Total nodes: {kg["meta"]["n_nodes"]} (was 260)')
print(f'Total edges: {kg["meta"]["n_edges"]} (was 824)')
print(f'Authors: {kg["meta"]["node_types"]["Author"]}')
print(f'Papers: {kg["meta"]["node_types"]["Paper"]}')
