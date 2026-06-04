# -*- coding: utf-8 -*-
"""
Ban do Viet Nam voi du lieu Geo THAT (geopandas + GeoJSON tu world-geojson).
Hien thi 11 dia diem da co nghien cuu khao sat lo au + 4 NC quoc gia (note).
"""
import sys, os, io, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import geopandas as gpd
from shapely.geometry import Point

plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.unicode_minus'] = False

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
GEO_DIR = os.path.join(os.path.dirname(__file__), 'geo_data')
OUT = os.path.join(os.path.dirname(__file__), 'charts')
os.makedirs(OUT, exist_ok=True)

# Load Vietnam GeoJSON (49 features — provinces)
gdf = gpd.read_file(os.path.join(GEO_DIR, 'vn_world_geojson.json'))
print(f'Loaded {len(gdf)} features')

# Locations of research studies (lat, lon, list of studies)
LOCATIONS = [
    {'name': 'Hà Nội', 'lat': 21.03, 'lon': 105.85,
     'studies': ['VN01 Hoa 2024 (n=3.910 GAD-7 lo âu 40,6 %)',
                 'VN18 Đào Thị Ngoãn 2025 (n=196 SV Y năm 4 DASS-21)',
                 'Trần Thị Mỵ Lương 2020 (n=540 chuyên DASS-42)',
                 'Trần Nguyễn Ngọc 2018 (n=170 BN HAM-A — người lớn)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)',
                 'Happy House Tran et al. 2023 (RCT n=1.084 lớp 10)']},
    {'name': 'Hải Phòng (Vĩnh Bảo)', 'lat': 20.74, 'lon': 106.49,
     'studies': ['VN15 Phạm Thị Ngọc 2024 (n=420 DASS-21 lo âu 53,8 %)']},
    {'name': 'Hải Dương', 'lat': 20.94, 'lon': 106.32,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh, n=8.473 tổng)']},
    {'name': 'Vĩnh Phúc', 'lat': 21.36, 'lon': 105.60,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Lạng Sơn (DTTS)', 'lat': 21.85, 'lon': 106.76,
     'studies': ['VN05 Ngô Anh Vinh 2024 (n=845 DASS-21 lo âu 54,4 % — DTTS nội trú)']},
    {'name': 'Thừa Thiên Huế', 'lat': 16.46, 'lon': 107.59,
     'studies': ['VN03 Phạm 2024 (n=273+273 hỗ trợ XH)',
                 'VN09 Hồ Thị Trúc Quỳnh 2025 (n=685 lo âu 65,8 %)',
                 'VN11 Trần Thảo Vi 2024 (NC dọc 3 năm n=611 ESSA)']},
    {'name': 'TPHCM', 'lat': 10.82, 'lon': 106.63,
     'studies': ['VN10 Trần Hồ Vĩnh Lộc 2024 (n=976 DASS-Y)',
                 'VN19 Duong 2025 (n=2.631 4 THPT + 4 GDTX, lo âu 50,3 %)',
                 'Nguyễn Thị Vân 2020 (n=558 STAI 15-18,5 %)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Đồng Nai', 'lat': 10.95, 'lon': 106.82,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Long An (Tân An)', 'lat': 10.53, 'lon': 106.41,
     'studies': ['VN16 Trần Đức Sĩ 2025 (HS THPT lo âu 57,2 %)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Vĩnh Long', 'lat': 10.25, 'lon': 105.97,
     'studies': ['VN14 Nguyễn Thanh Truyền 2024 (n=919 CES-D trầm cảm 12,2 %)']},
    {'name': 'An Giang', 'lat': 10.39, 'lon': 105.43,
     'studies': ['VN08 An Giang 2025 (HS THPT)']},
]

NATIONAL_STUDIES = [
    'VN02 V-NAMHS 2022: Khảo sát quốc gia 5.996 VTN tại 38 tỉnh/thành',
    'VN12 UNICEF 2022: School factors multi-province',
    'VN13 Nguyễn LX 2023: 5.730 sinh viên ĐH toàn quốc COVID GAD-7',
    'VN17 Dinh 2021: School factors multi-province',
]

# Convert locations to GeoDataFrame
loc_gdf = gpd.GeoDataFrame(
    LOCATIONS,
    geometry=[Point(l['lon'], l['lat']) for l in LOCATIONS],
    crs='EPSG:4326'
)
loc_gdf['n_studies'] = [len(l['studies']) for l in LOCATIONS]

# ============================================================
# Create map
# ============================================================
fig, ax = plt.subplots(figsize=(11, 14), dpi=160)

# Plot Vietnam provinces (real geo data)
gdf.plot(ax=ax, color='#E8F4F8', edgecolor='#2C5F8D', linewidth=0.6, zorder=1)

# Highlight provinces with research (light pink fill)
research_provinces = {'Hà Nội', 'Hải Phòng', 'Hải Dương', 'Vĩnh Phúc', 'Lạng Sơn',
                      'Thừa Thiên Huế', 'Hồ Chí Minh', 'TPHCM', 'Đồng Nai',
                      'Long An', 'Vĩnh Long', 'An Giang', 'Thua Thien Hue', 'Quang Tri'}
# Note: GeoJSON province names may differ — we just use markers, not fills

# Plot research location markers
for _, row in loc_gdf.iterrows():
    n = row['n_studies']
    size = 150 + n * 90
    color = '#D7263D' if n >= 3 else ('#F1A208' if n == 2 else '#06A77D')
    ax.scatter(row.geometry.x, row.geometry.y, s=size, c=color,
               edgecolors='black', linewidth=1.5, zorder=5, alpha=0.88)
    ax.text(row.geometry.x, row.geometry.y, str(n),
            ha='center', va='center', fontsize=11, fontweight='bold',
            color='white', zorder=6)

# Add labels with smart positioning
label_offsets = {
    'Hà Nội': (1.0, 0.4),
    'Hải Phòng (Vĩnh Bảo)': (1.5, -0.3),
    'Hải Dương': (1.8, -0.7),
    'Vĩnh Phúc': (-2.5, 0.4),
    'Lạng Sơn (DTTS)': (1.0, 0.5),
    'Thừa Thiên Huế': (1.5, 0.0),
    'TPHCM': (1.5, 0.5),
    'Đồng Nai': (2.0, -0.3),
    'Long An (Tân An)': (-2.8, -0.5),
    'Vĩnh Long': (-2.8, -0.4),
    'An Giang': (-2.8, 0.3),
}
for _, row in loc_gdf.iterrows():
    dx, dy = label_offsets.get(row['name'], (0.6, 0.6))
    ax.annotate(row['name'],
                xy=(row.geometry.x, row.geometry.y),
                xytext=(row.geometry.x + dx, row.geometry.y + dy),
                fontsize=9, fontweight='bold',
                arrowprops=dict(arrowstyle='-', color='gray', lw=0.6),
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                          edgecolor='gray', linewidth=0.5, alpha=0.92))

# Region labels
ax.text(104.0, 22.3, 'MIỀN BẮC', fontsize=12, color='#444', style='italic',
        fontweight='bold', ha='center', alpha=0.6)
ax.text(108.5, 16.5, 'MIỀN TRUNG', fontsize=12, color='#444', style='italic',
        fontweight='bold', ha='center', alpha=0.6)
ax.text(106.0, 9.5, 'MIỀN NAM\n(ĐB SCL)', fontsize=11, color='#444', style='italic',
        fontweight='bold', ha='center', alpha=0.6)

# Title
ax.set_title(
    'BẢN ĐỒ NGHIÊN CỨU KHẢO SÁT LO ÂU Ở VỊ THÀNH NIÊN VIỆT NAM\n'
    '11 địa điểm có nghiên cứu — Tổng hợp 19 bài VN trong hệ thống dự án (cập nhật 11/04/2026)',
    fontsize=13, fontweight='bold', pad=15)

# Legend
legend_elements = [
    Line2D([0], [0], marker='o', color='w', markerfacecolor='#06A77D',
           markersize=12, label='1 nghiên cứu', markeredgecolor='black'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='#F1A208',
           markersize=15, label='2 nghiên cứu', markeredgecolor='black'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor='#D7263D',
           markersize=18, label='≥ 3 nghiên cứu', markeredgecolor='black'),
]
ax.legend(handles=legend_elements, loc='upper right', fontsize=10,
          frameon=True, fancybox=True, shadow=True,
          title='Số nghiên cứu tại địa điểm', title_fontsize=10)

# Set bounds
ax.set_xlim(101.5, 113.5)
ax.set_ylim(8.0, 23.7)
ax.set_xlabel('Kinh độ (°E)', fontsize=10)
ax.set_ylabel('Vĩ độ (°N)', fontsize=10)
ax.set_aspect('equal')
ax.grid(True, alpha=0.2, linestyle=':')

# Note about national studies
note_text = ('NGHIÊN CỨU QUỐC GIA / ĐA TỈNH:\n' +
             '\n'.join('• ' + s for s in NATIONAL_STUDIES))
ax.text(101.7, 9.0, note_text, fontsize=8, va='bottom',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FFF8DC',
                  edgecolor='#888', linewidth=0.8))

# Hoang Sa & Truong Sa note
ax.text(112.3, 12.5, 'Hoàng Sa\nTrường Sa\n(VN)', fontsize=8,
        ha='center', style='italic', color='#444',
        bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                  edgecolor='#666', linewidth=0.5))

# Data source attribution
ax.text(0.99, 0.01, 'Bản đồ: GeoJSON 49 tỉnh VN (world-geojson)',
        transform=ax.transAxes, fontsize=7, ha='right', va='bottom',
        color='#888', style='italic')

plt.tight_layout()
out_path = os.path.join(OUT, 'ban_do_VN_NC_lo_au_geo.png')
plt.savefig(out_path, dpi=180, bbox_inches='tight', facecolor='white')
plt.close()

print(f'Saved: {out_path}')
print(f'Locations: {len(LOCATIONS)}')
print(f'Total studies plotted (including dups across multi-site): {sum(len(l["studies"]) for l in LOCATIONS)}')
