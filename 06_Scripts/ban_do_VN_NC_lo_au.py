# -*- coding: utf-8 -*-
"""
Tao ban do Viet Nam danh dau cac dia diem da co nghien cuu khao sat ve lo au.
Du lieu rut tu 19 bai NC VN trong he thong + bai cua cong su (Nguyen Thi Van, Tran Thi My Luong).

Sources:
- VN01 Hoa 2024 — Ha Noi
- VN02 V-NAMHS 2022 — 38 tinh thanh (national, mau dai dien)
- VN03 Pham 2024 — Hue
- VN04 Hoang Trung Hoc 2025 — 6 tinh: TPHCM, Dong Nai, Long An, Ha Noi, Hai Duong, Vinh Phuc
- VN05 Ngo Anh Vinh 2024 — Lang Son (DTTS Tay Nguyen rural)
- VN06 Bao Quyen 2025 — ???
- VN07 Nguyen Danh Lam 2022 — ???
- VN08 An Giang 2025 — An Giang
- VN09 Ho Thi Truc Quynh 2025 — Hue (Thua Thien Hue)
- VN10 Tran Ho Vinh Loc 2024 — TPHCM
- VN11 Tran Thao Vi 2024 J Rural Med — Hue (longitudinal 3 nam)
- VN12 UNICEF 2022 — multi-province
- VN13 Nguyen LX 2023 — quoc gia (5,730 SV)
- VN14 Nguyen Thanh Truyen 2024 — Vinh Long
- VN15 Pham Thi Ngoc 2024 — Hai Phong (Vinh Bao)
- VN16 Tran Duc Si 2025 — Long An (Tan An)
- VN17 Dinh 2021 — VN (locations TBD)
- VN18 Dao Thi Ngoan 2025 — Ha Noi (DH Y Ha Noi, SV nam 4)
- VN19 Duong 2025 — TPHCM (4 truong THPT + 4 GDTX)

Bonus from cong su:
- Nguyen Thi Van 2020 — TPHCM
- Tran Thi My Luong 2020 — Ha Noi (DH Su pham Ha Noi area)
- Tran Nguyen Ngoc 2018 — Ha Noi (BV Bach Mai, nguoi lon)
"""
import sys, os, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
import numpy as np

plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.unicode_minus'] = False

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'charts')
os.makedirs(OUT, exist_ok=True)

# Vietnam outline (simplified — main coastline + borders, lat/lon)
# Approximate Vietnam boundary points (clockwise from northwest)
VN_BOUNDARY = [
    (102.14, 22.40),  # NW corner Lai Chau
    (102.50, 22.78),  # near Phong Tho
    (103.49, 22.79),  # Lao Cai border
    (104.10, 22.83),  # Ha Giang
    (105.30, 23.36),  # Northern border
    (106.20, 22.95),  # Cao Bang
    (106.84, 22.79),  # Lang Son
    (107.10, 22.45),
    (108.05, 21.55),  # Mong Cai
    (108.40, 21.05),  # Hai Phong area coast
    (107.15, 20.95),
    (106.70, 20.74),  # Ninh Binh
    (106.20, 20.10),
    (106.00, 19.50),
    (105.80, 18.90),  # Vinh
    (105.65, 18.30),  # Ha Tinh
    (106.10, 17.80),  # Quang Binh
    (106.55, 17.30),  # Quang Tri
    (107.35, 16.85),  # Hue
    (108.20, 16.10),  # Da Nang
    (109.10, 15.45),  # Quang Ngai
    (109.30, 14.50),  # Quy Nhon
    (109.25, 13.80),  # Tuy Hoa
    (109.20, 13.05),  # Nha Trang
    (109.05, 12.20),  # Phan Rang
    (108.65, 11.55),  # Phan Thiet
    (107.85, 10.70),  # Vung Tau
    (106.95, 10.40),  # Mekong delta east
    (106.30, 10.30),
    (105.70, 10.05),  # Ca Mau peninsula north
    (105.00, 9.50),
    (104.85, 8.60),   # Ca Mau
    (104.45, 8.50),   # Ca Mau
    (104.50, 9.20),   # West Mekong
    (105.05, 9.95),
    (105.50, 10.65),  # Long Xuyen
    (105.20, 10.95),  # An Giang
    (104.85, 10.95),  # Cambodia border
    (104.40, 11.50),  # Cambodia border
    (105.85, 12.50),  # Cambodia border
    (106.40, 12.85),  # Cambodia border
    (107.30, 13.85),  # Central highlands
    (107.70, 14.40),  # Kon Tum
    (107.55, 15.10),  # Quang Nam border
    (107.30, 15.85),  # A Luoi
    (106.65, 16.55),  # Quang Tri
    (106.55, 17.25),  # Quang Binh border
    (105.80, 18.45),  # Nghe An
    (104.90, 19.70),  # Nghe An mountains
    (104.10, 19.85),  # Thanh Hoa
    (103.55, 20.65),  # Hoa Binh
    (103.10, 20.95),  # Son La
    (102.85, 21.55),  # Dien Bien
    (102.14, 22.40),  # close
]

# Locations of research studies
LOCATIONS = [
    # (name, lat, lon, list of studies)
    {'name': 'Hà Nội', 'lat': 21.03, 'lon': 105.85,
     'studies': ['VN01 Hoa 2024 (n=3.910 GAD-7 40,6%)',
                 'VN18 Đào Thị Ngoãn 2025 (n=196 SV Y4 DASS-21 lo âu 43,9%)',
                 'Trần Thị Mỵ Lương 2020 (n=540 chuyên DASS-42 14,2%)',
                 'Trần Nguyễn Ngọc 2018 (n=170 BN HAM-A — người lớn)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Hải Phòng (Vĩnh Bảo)', 'lat': 20.74, 'lon': 106.49,
     'studies': ['VN15 Phạm Thị Ngọc 2024 (n=420 DASS-21 lo âu 53,8%)']},
    {'name': 'Hải Dương', 'lat': 20.94, 'lon': 106.32,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh, n=8.473 tổng)']},
    {'name': 'Vĩnh Phúc', 'lat': 21.36, 'lon': 105.60,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Lạng Sơn (DTTS)', 'lat': 21.85, 'lon': 106.76,
     'studies': ['VN05 Ngô Anh Vinh 2024 (n=845 DASS-21 lo âu 54,4% — DTTS)']},
    {'name': 'Thừa Thiên Huế', 'lat': 16.46, 'lon': 107.59,
     'studies': ['VN03 Phạm 2024 (n=273+273 hỗ trợ XH)',
                 'VN09 Hồ Thị Trúc Quỳnh 2025 (n=685 lo âu 65,8%)',
                 'VN11 Trần Thảo Vi 2024 (NC dọc 3 năm n=611 ESSA)']},
    {'name': 'TPHCM', 'lat': 10.82, 'lon': 106.63,
     'studies': ['VN10 Trần Hồ Vĩnh Lộc 2024 (n=976 DASS-Y)',
                 'VN19 Duong 2025 (n=2.631 4 THPT+4 GDTX, lo âu 50,3%)',
                 'Nguyễn Thị Vân 2020 (n=558 STAI 15-18,5%)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Đồng Nai', 'lat': 10.95, 'lon': 106.82,
     'studies': ['VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Long An (Tân An)', 'lat': 10.53, 'lon': 106.41,
     'studies': ['VN16 Trần Đức Sĩ 2025 (HS THPT lo âu 57,2%)',
                 'VN04 Hoàng Trung Học 2025 (1/6 tỉnh)']},
    {'name': 'Vĩnh Long', 'lat': 10.25, 'lon': 105.97,
     'studies': ['VN14 Nguyễn Thanh Truyền 2024 (n=919 CES-D trầm cảm 12,2%)']},
    {'name': 'An Giang', 'lat': 10.39, 'lon': 105.43,
     'studies': ['VN08 An Giang 2025 (HS THPT)']},
]

NATIONAL_STUDIES = [
    'VN02 V-NAMHS 2022: Khảo sát quốc gia 5.996 VTN tại 38 tỉnh/thành',
    'VN12 UNICEF 2022: Multi-province school factors',
    'VN13 Nguyễn LX 2023: 5.730 sinh viên ĐH toàn quốc COVID GAD-7',
    'VN17 Dinh 2021: Multi-province school factors',
]

# ============================================================
# Create map
# ============================================================
fig, ax = plt.subplots(figsize=(10, 13), dpi=160)

# Draw Vietnam outline
xs, ys = zip(*VN_BOUNDARY)
ax.fill(xs, ys, color='#E8F4F8', edgecolor='#2C5F8D', linewidth=1.8, zorder=1)

# Add region labels
ax.text(105.0, 22.0, 'Miền Bắc', fontsize=11, color='#666', style='italic', ha='center')
ax.text(108.0, 16.0, 'Miền Trung', fontsize=11, color='#666', style='italic', ha='center')
ax.text(105.5, 9.8, 'ĐB SCL\n(Mekong)', fontsize=10, color='#666', style='italic', ha='center')

# Draw research location markers
for loc in LOCATIONS:
    n_studies = len(loc['studies'])
    # Marker size based on number of studies
    size = 100 + n_studies * 80
    color = '#D7263D' if n_studies >= 3 else ('#F1A208' if n_studies == 2 else '#06A77D')
    ax.scatter(loc['lon'], loc['lat'], s=size, c=color, edgecolors='black',
               linewidth=1.5, zorder=5, alpha=0.85)
    # Number badge
    ax.text(loc['lon'], loc['lat'], str(n_studies),
            ha='center', va='center', fontsize=11, fontweight='bold',
            color='white', zorder=6)

# Add labels with offset
label_offsets = {
    'Hà Nội': (0.8, 0.4),
    'Hải Phòng (Vĩnh Bảo)': (1.0, -0.3),
    'Hải Dương': (1.5, -0.5),
    'Vĩnh Phúc': (-2.0, 0.4),
    'Lạng Sơn (DTTS)': (0.5, 0.5),
    'Thừa Thiên Huế': (1.0, 0.0),
    'TPHCM': (1.0, 0.5),
    'Đồng Nai': (1.5, -0.3),
    'Long An (Tân An)': (-2.5, -0.5),
    'Vĩnh Long': (-2.5, -0.4),
    'An Giang': (-2.5, 0.3),
}
for loc in LOCATIONS:
    dx, dy = label_offsets.get(loc['name'], (0.5, 0.5))
    ax.annotate(loc['name'],
                xy=(loc['lon'], loc['lat']),
                xytext=(loc['lon'] + dx, loc['lat'] + dy),
                fontsize=9, fontweight='bold',
                arrowprops=dict(arrowstyle='-', color='gray', lw=0.5),
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                          edgecolor='gray', linewidth=0.5, alpha=0.9))

# Title
ax.set_title('BẢN ĐỒ NGHIÊN CỨU KHẢO SÁT LO ÂU Ở VỊ THÀNH NIÊN VIỆT NAM\n'
             '(11 địa điểm nghiên cứu — Tổng hợp 19 bài VN trong hệ thống dự án)',
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
ax.set_xlim(101.5, 113.0)
ax.set_ylim(8.0, 23.7)
ax.set_xlabel('Kinh độ (°E)', fontsize=10)
ax.set_ylabel('Vĩ độ (°N)', fontsize=10)
ax.set_aspect('equal')
ax.grid(True, alpha=0.2, linestyle=':')

# Add note about national studies
note_text = ('NGHIÊN CỨU QUỐC GIA / ĐA TỈNH (không hiển thị marker):\n' +
             '\n'.join('• ' + s for s in NATIONAL_STUDIES))
ax.text(101.7, 9.0, note_text, fontsize=8, va='bottom',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FFF8DC',
                  edgecolor='#888', linewidth=0.8))

# Add Hoang Sa & Truong Sa note
ax.text(112.5, 12.0, 'Hoàng Sa\nTrường Sa', fontsize=8,
        ha='center', style='italic', color='#666',
        bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                  edgecolor='#666', linewidth=0.5))

plt.tight_layout()
out_path = os.path.join(OUT, 'ban_do_VN_NC_lo_au.png')
plt.savefig(out_path, dpi=180, bbox_inches='tight', facecolor='white')
plt.close()

print(f'Saved: {out_path}')
print(f'Locations: {len(LOCATIONS)}')
print(f'Total studies plotted: {sum(len(l["studies"]) for l in LOCATIONS)}')
print(f'National studies (no marker): {len(NATIONAL_STUDIES)}')
