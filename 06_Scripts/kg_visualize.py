# -*- coding: utf-8 -*-
"""Visualize KG: pyvis interactive HTML + static PNG subgraphs."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import networkx as nx
from pyvis.network import Network
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
KG_DIR = os.path.join(os.path.dirname(__file__), 'kg_data')

G = nx.read_graphml(os.path.join(KG_DIR, 'kg_v1.graphml'))
print(f'Loaded: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges')

# ============================================================
# Interactive HTML with pyvis
# ============================================================
print('\nBuilding pyvis interactive graph...')
net = Network(height='900px', width='100%', bgcolor='#ffffff',
              font_color='#222222', directed=True, cdn_resources='in_line')
net.force_atlas_2based(gravity=-50, central_gravity=0.01, spring_length=100)

# Color by type
color_map = {
    'Paper': '#1f77b4',       # blue
    'Author': '#ff7f0e',      # orange
    'Journal': '#2ca02c',     # green
    'Country': '#d62728',     # red
    'Method': '#9467bd',      # purple
    'Outcome': '#8c564b',     # brown
    'Scale': '#e377c2',       # pink
    'EffectSize': '#bcbd22',  # yellow-green
    'SampleSize': '#17becf',  # cyan
    'Year': '#7f7f7f',        # gray
}

for node, data in G.nodes(data=True):
    t = data.get('type', 'Unknown')
    label = data.get('descriptor') or data.get('name') or data.get('value') or node
    if t == 'Paper':
        label = node  # Just show ID for papers
        size = 30
    elif t == 'EffectSize':
        label = f'{data.get("es_type", "")}={data.get("value", "")}'
        size = 10
    else:
        label = str(label)[:25]
        size = 15
    title = f'{t}: {label}'
    net.add_node(node, label=label, title=title,
                 color=color_map.get(t, '#cccccc'), size=size)

for u, v, d in G.edges(data=True):
    rel = d.get('rel', '')
    net.add_edge(u, v, title=rel, label=rel[:10], arrows='to')

net.set_options("""
var options = {
  "physics": {
    "forceAtlas2Based": {
      "gravitationalConstant": -50,
      "centralGravity": 0.01,
      "springLength": 100,
      "damping": 0.4
    },
    "minVelocity": 0.75,
    "solver": "forceAtlas2Based"
  },
  "interaction": {
    "hover": true,
    "tooltipDelay": 100
  }
}
""")

out_html = os.path.join(KG_DIR, 'kg_interactive.html')
# Manually write with utf-8
html = net.generate_html()
with open(out_html, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Saved interactive: {out_html}')

# ============================================================
# Static visualization: Vietnam-focused subgraph
# ============================================================
print('\nBuilding VN-focused static graph...')

# Find papers in Vietnam
vn_country = 'Country::Vietnam'
if vn_country in G.nodes:
    # Get predecessors (papers pointing to Vietnam)
    vn_papers = [u for u, v, d in G.in_edges(vn_country, data=True)
                 if d.get('rel') == 'CONDUCTED_IN']
    # Subgraph: these papers + their neighbors
    nodes_to_include = set(vn_papers)
    nodes_to_include.add(vn_country)
    for p in vn_papers:
        for succ in G.successors(p):
            nodes_to_include.add(succ)
    subG = G.subgraph(nodes_to_include)
    print(f'VN subgraph: {subG.number_of_nodes()} nodes, {subG.number_of_edges()} edges')

    fig, ax = plt.subplots(figsize=(16, 12), dpi=150)
    pos = nx.spring_layout(subG, k=2.0, iterations=50, seed=42)

    # Draw nodes by type
    for t, color in color_map.items():
        nodes = [n for n in subG.nodes if G.nodes[n].get('type') == t]
        if nodes:
            nx.draw_networkx_nodes(subG, pos, nodelist=nodes, node_color=color,
                                   node_size=300, ax=ax, label=t, alpha=0.8)

    # Draw edges
    nx.draw_networkx_edges(subG, pos, alpha=0.3, arrows=True, ax=ax,
                          edge_color='gray', width=0.5, arrowsize=8)

    # Labels (short)
    labels = {}
    for n in subG.nodes:
        t = G.nodes[n].get('type', '')
        if t == 'Paper':
            labels[n] = n
        elif t == 'Country':
            labels[n] = G.nodes[n].get('name', '')
        elif t in ('Method', 'Outcome', 'Scale'):
            labels[n] = G.nodes[n].get('name', '')[:15]
    nx.draw_networkx_labels(subG, pos, labels, font_size=6, ax=ax,
                           font_family='Times New Roman')

    ax.set_title('Knowledge Graph — Các bài nghiên cứu tại Việt Nam (subgraph)',
                 fontsize=14, fontweight='bold')
    ax.legend(loc='upper right', fontsize=9)
    ax.axis('off')
    plt.tight_layout()
    out_png = os.path.join(KG_DIR, 'kg_vietnam_subgraph.png')
    plt.savefig(out_png, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f'Saved VN subgraph: {out_png}')

# ============================================================
# Methods-Outcomes bipartite view
# ============================================================
print('\nBuilding Methods-Outcomes heatmap-style view...')

method_nodes = [n for n in G.nodes if G.nodes[n].get('type') == 'Method']
outcome_nodes = [n for n in G.nodes if G.nodes[n].get('type') == 'Outcome']
paper_nodes = [n for n in G.nodes if G.nodes[n].get('type') == 'Paper']

# Count papers per (method, outcome) pair
method_outcome_count = {}
for p in paper_nodes:
    methods = [v for u, v, d in G.out_edges(p, data=True) if d.get('rel') == 'USED_METHOD']
    outcomes = [v for u, v, d in G.out_edges(p, data=True) if d.get('rel') == 'MEASURED']
    for m in methods:
        for o in outcomes:
            method_outcome_count[(m, o)] = method_outcome_count.get((m, o), 0) + 1

if method_nodes and outcome_nodes:
    import numpy as np
    m_names = sorted([G.nodes[n].get('name', n) for n in method_nodes])
    o_names = sorted([G.nodes[n].get('name', n) for n in outcome_nodes])

    matrix = np.zeros((len(m_names), len(o_names)))
    for (m, o), count in method_outcome_count.items():
        mn = G.nodes[m].get('name', '')
        on = G.nodes[o].get('name', '')
        if mn in m_names and on in o_names:
            matrix[m_names.index(mn), o_names.index(on)] = count

    fig, ax = plt.subplots(figsize=(10, 8), dpi=150)
    im = ax.imshow(matrix, cmap='YlOrRd', aspect='auto')
    ax.set_xticks(range(len(o_names)))
    ax.set_xticklabels(o_names, rotation=45, ha='right', fontsize=9)
    ax.set_yticks(range(len(m_names)))
    ax.set_yticklabels(m_names, fontsize=9)

    # Annotate cells
    for i in range(len(m_names)):
        for j in range(len(o_names)):
            if matrix[i, j] > 0:
                ax.text(j, i, int(matrix[i, j]), ha='center', va='center',
                       color='black' if matrix[i, j] < matrix.max()/2 else 'white',
                       fontsize=10, fontweight='bold')

    plt.colorbar(im, ax=ax, label='Số bài nghiên cứu')
    ax.set_title('Ma trận Phương pháp × Kết cục — 60 bài trong hệ thống',
                 fontsize=12, fontweight='bold')
    plt.tight_layout()
    out_heatmap = os.path.join(KG_DIR, 'kg_method_outcome_heatmap.png')
    plt.savefig(out_heatmap, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f'Saved heatmap: {out_heatmap}')

print('\n=== DONE ===')
