# -*- coding: utf-8 -*-
"""
Phase F (Visualization) — Create 4 views for Author KG:
1. Interactive collaboration network (pyvis HTML)
2. Author × Topic heatmap (matplotlib)
3. Top 20 authors + Country distribution (bar charts)
4. Timeline Year × Topic (heatmap)
"""
import os, sys, io, json
from collections import defaultdict, Counter
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')
VIZ_DIR = os.path.join(OUT_DIR, 'visualizations')
os.makedirs(VIZ_DIR, exist_ok=True)

with open(os.path.join(OUT_DIR, 'authors_normalized.json'), encoding='utf-8') as f:
    authors = json.load(f)
with open(os.path.join(OUT_DIR, 'author_kg_v1.json'), encoding='utf-8') as f:
    kg = json.load(f)
with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

# ============================================================
# VIEW 1: Interactive Collaboration Network (pyvis)
# ============================================================
def build_network_html():
    try:
        from pyvis.network import Network
    except ImportError:
        print('  pyvis not installed. Skipping interactive HTML.')
        return

    net = Network(height='800px', width='100%', bgcolor='#ffffff', font_color='#222222',
                  directed=False, notebook=False)
    # Improved physics
    net.force_atlas_2based(gravity=-50, central_gravity=0.01, spring_length=100, spring_strength=0.08)

    # Country color map
    country_colors = {
        'Vietnam': '#dc2626', 'USA': '#2563eb', 'China': '#d97706', 'Australia': '#16a34a',
        'UK': '#7c3aed', 'Japan': '#db2777', 'Korea': '#ea580c', 'Ireland': '#059669',
        'Netherlands': '#0891b2', 'India': '#84cc16', 'Indonesia': '#f59e0b',
        'Philippines': '#e11d48', 'Sri Lanka': '#4f46e5', 'Saudi Arabia': '#0284c7',
        'Norway': '#1e3a8a', 'Ethiopia': '#a855f7', 'USA/India': '#8b5cf6',
        'Australia/Vietnam': '#f97316',
    }

    # Add Author nodes — only those with ≥ 1 paper or priority
    for aid, rec in authors.items():
        n_papers = rec.get('n_papers', 0)
        is_priority = rec.get('is_priority', False)
        if n_papers == 0 and not is_priority:
            continue

        country = rec.get('country_primary') or rec.get('country_hint', '')
        color = country_colors.get(country, '#9ca3af')
        size = 15 + min(n_papers * 5, 30)
        if is_priority:
            border = '#000000'
            border_width = 3
        else:
            border = color
            border_width = 1

        # Tooltip
        papers = rec.get('papers', [])
        expertise = rec.get('expertise', [])
        tooltip = f"{rec['full_name']}\n"
        if country: tooltip += f"Country: {country}\n"
        tooltip += f"Papers: {n_papers} ({', '.join(papers[:5])}{'...' if len(papers) > 5 else ''})"
        if expertise:
            tooltip += f"\nExpertise: {', '.join(expertise[:4])}"
        if rec.get('role_notes'):
            tooltip += f"\nRole: {rec['role_notes']}"
        if is_priority:
            tooltip += '\n⭐ Priority author'

        net.add_node(aid,
                     label=rec['full_name'][:30],
                     title=tooltip,
                     color={'background': color, 'border': border},
                     borderWidth=border_width,
                     size=size)

    # Add CO_AUTHORED edges
    for e in kg['edges']:
        if e.get('type') == 'CO_AUTHORED':
            source, target = e['source'], e['target']
            # Avoid duplicate (A→B and B→A)
            if source > target:
                continue
            strength = int(e.get('strength', 1))
            net.add_edge(source, target,
                         width=1 + strength,
                         title=f"Co-authored {strength} paper(s)",
                         color='#cccccc')

    # Save HTML
    out = os.path.join(VIZ_DIR, 'author_collab_network.html')
    try:
        net.write_html(out, open_browser=False)
        print(f'✓ Collaboration network saved: {out}')
    except Exception as e:
        print(f'  WARN pyvis failed: {e}')

build_network_html()

# ============================================================
# VIEW 2: Author × Topic Heatmap
# ============================================================
def build_topic_heatmap():
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np
    except ImportError:
        print('  matplotlib not installed. Skipping heatmap.')
        return

    # Extract STUDIES edges
    author_topic = defaultdict(lambda: defaultdict(int))
    for e in kg['edges']:
        if e.get('type') == 'STUDIES':
            author_topic[e['source']][e['target']] += int(e.get('paper_count', 1))

    # Top 25 authors by n_papers
    top_authors = sorted([a for a in authors.values() if a['n_papers'] > 0],
                         key=lambda a: (-a['n_papers'], a['surname']))[:25]
    # Add priority authors if not in top (ensure key figures visible)
    priority_ids = {a['canonical_id'] for a in authors.values() if a.get('is_priority')}
    top_author_ids = {a['canonical_id'] for a in top_authors}
    for pid in priority_ids:
        if pid not in top_author_ids and len(top_authors) < 30:
            top_authors.append(authors[pid])
            top_author_ids.add(pid)

    # Topics with any STUDIES edge
    all_topics = set()
    for t_map in author_topic.values():
        all_topics.update(t_map.keys())
    topics = sorted(all_topics)
    topic_labels = [t.replace('TOP_', '').replace('_', ' ').title() for t in topics]

    if not top_authors or not topics:
        print('  No author-topic data. Skipping heatmap.')
        return

    # Build matrix
    matrix = np.zeros((len(top_authors), len(topics)))
    for i, a in enumerate(top_authors):
        aid = a['canonical_id']
        for j, t in enumerate(topics):
            matrix[i, j] = author_topic[aid].get(t, 0)

    fig, ax = plt.subplots(figsize=(14, 12))
    im = ax.imshow(matrix, cmap='YlOrRd', aspect='auto')

    ax.set_xticks(range(len(topics)))
    ax.set_xticklabels(topic_labels, rotation=45, ha='right', fontsize=9)
    ax.set_yticks(range(len(top_authors)))
    labels = []
    for a in top_authors:
        lbl = a['full_name'][:32]
        if a.get('is_priority'):
            lbl = '⭐ ' + lbl
        labels.append(lbl)
    ax.set_yticklabels(labels, fontsize=9)

    # Cell values
    for i in range(len(top_authors)):
        for j in range(len(topics)):
            val = int(matrix[i, j])
            if val > 0:
                ax.text(j, i, str(val), ha='center', va='center',
                        color='white' if val >= 2 else 'black', fontsize=8)

    plt.colorbar(im, ax=ax, label='Số paper liên quan topic')
    ax.set_title('Heatmap: Tác giả × Chủ đề nghiên cứu (25 tác giả hàng đầu + priority)',
                 fontsize=13, pad=15)
    plt.tight_layout()

    out = os.path.join(VIZ_DIR, 'author_topic_heatmap.png')
    plt.savefig(out, dpi=120, bbox_inches='tight')
    plt.close()
    print(f'✓ Heatmap saved: {out}')

build_topic_heatmap()

# ============================================================
# VIEW 3: Bar charts — Top authors, Country, Institutions
# ============================================================
def build_bar_charts():
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
    except ImportError:
        return

    fig, axes = plt.subplots(1, 3, figsize=(22, 8))

    # --- 3.1 Top 20 authors by paper count ---
    top_20 = sorted([a for a in authors.values() if a['n_papers'] > 0],
                    key=lambda a: -a['n_papers'])[:20]
    names = [a['full_name'][:32] + (' ⭐' if a.get('is_priority') else '') for a in top_20]
    counts = [a['n_papers'] for a in top_20]
    bars = axes[0].barh(range(len(names)), counts,
                        color=['#dc2626' if a.get('is_priority') else '#3b82f6' for a in top_20])
    axes[0].set_yticks(range(len(names)))
    axes[0].set_yticklabels(names, fontsize=8)
    axes[0].invert_yaxis()
    axes[0].set_xlabel('Số paper')
    axes[0].set_title('Top 20 tác giả (⭐ = priority)', fontsize=12)
    for bar, c in zip(bars, counts):
        axes[0].text(bar.get_width() + 0.05, bar.get_y() + bar.get_height()/2,
                     str(c), va='center', fontsize=8)

    # --- 3.2 Country distribution ---
    country_counter = Counter()
    for a in authors.values():
        c = a.get('country_primary') or a.get('country_hint', '')
        if c: country_counter[c] += 1
    top_countries = country_counter.most_common(12)
    countries = [c[0] for c in top_countries]
    ncounts = [c[1] for c in top_countries]
    axes[1].barh(range(len(countries)), ncounts, color='#16a34a')
    axes[1].set_yticks(range(len(countries)))
    axes[1].set_yticklabels(countries, fontsize=9)
    axes[1].invert_yaxis()
    axes[1].set_xlabel('Số tác giả')
    axes[1].set_title('Phân bố quốc gia (top 12)', fontsize=12)
    for i, c in enumerate(ncounts):
        axes[1].text(c + 0.2, i, str(c), va='center', fontsize=8)

    # --- 3.3 Institution distribution (priority authors only — known affiliations) ---
    inst_counter = Counter()
    for a in authors.values():
        for inst in a.get('affiliations', []):
            inst_counter[inst] += 1
    top_inst = inst_counter.most_common(12)
    if top_inst:
        insts = [i[0][:40] for i in top_inst]
        icounts = [i[1] for i in top_inst]
        axes[2].barh(range(len(insts)), icounts, color='#7c3aed')
        axes[2].set_yticks(range(len(insts)))
        axes[2].set_yticklabels(insts, fontsize=8)
        axes[2].invert_yaxis()
        axes[2].set_xlabel('Số tác giả')
        axes[2].set_title('Top đơn vị (từ priority authors)', fontsize=12)
        for i, c in enumerate(icounts):
            axes[2].text(c + 0.05, i, str(c), va='center', fontsize=8)
    else:
        axes[2].text(0.5, 0.5, 'Chưa có dữ liệu institutions\n(cần Phase D augment thêm)',
                     ha='center', va='center', fontsize=10, transform=axes[2].transAxes)
        axes[2].axis('off')

    plt.tight_layout()
    out = os.path.join(VIZ_DIR, 'author_distributions.png')
    plt.savefig(out, dpi=120, bbox_inches='tight')
    plt.close()
    print(f'✓ Bar charts saved: {out}')

build_bar_charts()

# ============================================================
# VIEW 4: Year × Topic Timeline Heatmap
# ============================================================
def build_timeline_heatmap():
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np
        import re
    except ImportError:
        return

    # Extract year from canonical_index descriptors
    year_topic = defaultdict(lambda: defaultdict(int))
    TOPIC_KEYWORDS_RE = {
        'Anxiety': r'(anxiety|lo âu|GAD|SAD|DASS)',
        'Depression': r'(depression|trầm cảm|CESD|PHQ|MDD)',
        'CBT': r'(CBT|iCBT|Maya)',
        'MHPSS': r'(MHPSS|MentalHealth|Psychosocial)',
        'Adolescent': r'(Adolescent|VTN|student|youth)',
        'COVID': r'(COVID|pandemic)',
        'Digital': r'(Digital|DMHI|app|online|Mobile|internet)',
        'Resilience': r'(Resilience)',
        'Bullying': r'(Bullying|bắt nạt|Cyberbullying)',
        'School Intervention': r'(School|classroom|BESST|HappyHouse)',
        'Prevalence': r'(Prevalence|Epidemiology|survey)',
    }

    for cid, meta in canon.items():
        descr = meta.get('descriptor', '')
        year_m = re.search(r'(\d{4})', descr)
        if not year_m: continue
        year = int(year_m.group(1))
        if year < 2010 or year > 2030: continue
        for topic, pattern in TOPIC_KEYWORDS_RE.items():
            if re.search(pattern, descr, re.IGNORECASE):
                year_topic[year][topic] += 1

    if not year_topic:
        print('  No year-topic data. Skipping timeline.')
        return

    years = sorted(year_topic.keys())
    topics = list(TOPIC_KEYWORDS_RE.keys())

    matrix = np.zeros((len(topics), len(years)))
    for i, t in enumerate(topics):
        for j, y in enumerate(years):
            matrix[i, j] = year_topic[y].get(t, 0)

    fig, ax = plt.subplots(figsize=(max(10, len(years) * 0.9), 7))
    im = ax.imshow(matrix, cmap='Blues', aspect='auto')

    ax.set_xticks(range(len(years)))
    ax.set_xticklabels(years, fontsize=10)
    ax.set_yticks(range(len(topics)))
    ax.set_yticklabels(topics, fontsize=10)

    for i in range(len(topics)):
        for j in range(len(years)):
            v = int(matrix[i, j])
            if v > 0:
                ax.text(j, i, str(v), ha='center', va='center',
                        color='white' if v >= 3 else 'black', fontsize=9)

    plt.colorbar(im, ax=ax, label='Số paper')
    ax.set_title('Timeline: Năm xuất bản × Chủ đề nghiên cứu (68 papers)', fontsize=13, pad=15)
    ax.set_xlabel('Năm')
    plt.tight_layout()

    out = os.path.join(VIZ_DIR, 'timeline_year_topic.png')
    plt.savefig(out, dpi=120, bbox_inches='tight')
    plt.close()
    print(f'✓ Timeline saved: {out}')

build_timeline_heatmap()

print()
print('='*70)
print('VISUALIZATIONS COMPLETE')
print('='*70)
print(f'Output directory: {VIZ_DIR}')
for f in sorted(os.listdir(VIZ_DIR)):
    size = os.path.getsize(os.path.join(VIZ_DIR, f))
    print(f'  {f} ({size:,} bytes)')
