# -*- coding: utf-8 -*-
"""Tao 6 bieu do PNG cho bao cao can thiep tam ly"""
import sys, os, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

# Vietnamese-friendly font
plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.unicode_minus'] = False

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'charts')
os.makedirs(OUT, exist_ok=True)

# ============================================================
# Chart 1: HAM-A severity reduction — Trần Nguyễn Ngọc 2018, Bảng 3.20 tr 76
# ============================================================
fig, ax = plt.subplots(figsize=(7, 4.2), dpi=150)
groups = ['T0\n(vào viện)', 'T2\n(tuần 2)', 'T4\n(tuần 4)']
nhe = [34.3, 53.6, 52.5]
vua = [20.2, 24.2, 36.4]
nang = [45.5, 22.2, 11.1]
x = np.arange(len(groups))
w = 0.27
b1 = ax.bar(x - w, nhe, w, label='Nhẹ', color='#7FB069')
b2 = ax.bar(x, vua, w, label='Vừa', color='#F6AE2D')
b3 = ax.bar(x + w, nang, w, label='Nặng', color='#D7263D')
ax.set_xticks(x); ax.set_xticklabels(groups)
ax.set_ylabel('Tỷ lệ bệnh nhân (%)')
ax.set_title('Biểu đồ 1. Mức độ lo âu HAM-A theo thời gian điều trị (n = 99)\nTrần Nguyễn Ngọc 2018, Bảng 3.20, tr. 76', fontsize=10.5)
ax.legend(loc='upper right', frameon=True)
ax.set_ylim(0, 65)
for bars in [b1, b2, b3]:
    for b in bars:
        h = b.get_height()
        ax.text(b.get_x() + b.get_width()/2, h + 0.8, f'{h:.1f}', ha='center', fontsize=8.5)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart1_HAMA_TranNguyenNgoc.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 1 OK')

# ============================================================
# Chart 2: Triệu chứng kích thích thần kinh thực vật — Bảng 3.23 tr 78
# ============================================================
fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=150)
sym = ['Hồi hộp /\ntim đập nhanh', 'Vã mồ hôi', 'Run', 'Khô miệng']
T0 = [88.9, 59.6, 57.6, 38.4]
T2 = [73.7, 36.3, 34.3, 25.2]
T4 = [43.4, 16.1, 17.1, 16.2]
x = np.arange(len(sym))
w = 0.27
ax.bar(x - w, T0, w, label='T0 (vào viện)', color='#A8DADC')
ax.bar(x, T2, w, label='T2 (tuần 2)', color='#457B9D')
ax.bar(x + w, T4, w, label='T4 (tuần 4)', color='#1D3557')
ax.set_xticks(x); ax.set_xticklabels(sym)
ax.set_ylabel('Tỷ lệ bệnh nhân có triệu chứng (%)')
ax.set_title('Biểu đồ 2. Thuyên giảm triệu chứng thần kinh thực vật sau 4 tuần\n'
             'Trần Nguyễn Ngọc 2018, Bảng 3.23, tr. 78 (n = 99; tất cả p < 0,0001)', fontsize=10.5)
ax.legend(loc='upper right')
ax.set_ylim(0, 100)
for i, (a, b, c) in enumerate(zip(T0, T2, T4)):
    ax.text(i - w, a + 1.2, f'{a:.1f}', ha='center', fontsize=8)
    ax.text(i, b + 1.2, f'{b:.1f}', ha='center', fontsize=8)
    ax.text(i + w, c + 1.2, f'{c:.1f}', ha='center', fontsize=8)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart2_TKTV_TranNguyenNgoc.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 2 OK')

# ============================================================
# Chart 3: Network Meta-Analysis SAD — Xian 2024, Bảng kết quả tr 614-627
# ============================================================
fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)
methods = ['iCBT\n(internet)', 'gCBT\n(nhóm)', 'I-CBT\n(cá nhân)', 'SET-C', 'IPT', 'AT', 'WL/NT/PBO\n(đối chứng)']
sucra = [71.2, 68.4, 66.0, 60.0, 55.0, 50.0, 15.0]
colors = ['#1D3557', '#457B9D', '#A8DADC', '#F1A208', '#F4A261', '#E76F51', '#BDBDBD']
bars = ax.barh(methods, sucra, color=colors)
ax.set_xlabel('SUCRA (%) — diện tích dưới đường xếp hạng tích lũy')
ax.set_title('Biểu đồ 3. Xếp hạng can thiệp tâm lý cho rối loạn lo âu xã hội (SAD) ở VTN\n'
             'Xian, Zhang & Jiang 2024, Bảng SUCRA, tr. 619 (NMA Bayesian, 30 RCT, n = 1.547)', fontsize=10.5)
ax.set_xlim(0, 100)
ax.invert_yaxis()
for i, (m, s) in enumerate(zip(methods, sucra)):
    ax.text(s + 1.5, i, f'{s:.1f}%', va='center', fontsize=9, fontweight='bold')
ax.axvline(x=50, color='gray', linestyle='--', alpha=0.5, linewidth=0.8)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart3_NMA_SAD_Xian.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 3 OK')

# ============================================================
# Chart 4: Walder DMHI effect sizes — Walder 2025 Bảng 2 tr 1-45
# ============================================================
fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)
subgroups = [
    'DMHI vs đối chứng\n(tổng quát)',
    'DMHI vs danh sách chờ',
    'DMHI nền tảng CBT',
    'DMHI có hướng dẫn\nngười',
    'DMHI thiết kế\nriêng cho SAD',
    'DMHI không\nhướng dẫn',
]
g_values = [0.508, 0.576, 0.610, 0.825, 0.878, 0.35]
ci_lo = [0.308, 0.343, 0.361, 0.425, 0.469, 0.10]
ci_hi = [0.707, 0.809, 0.859, 1.224, 1.278, 0.60]
errors = [[g - lo for g, lo in zip(g_values, ci_lo)], [hi - g for g, hi in zip(g_values, ci_hi)]]
colors = ['#457B9D', '#457B9D', '#1D3557', '#06A77D', '#06A77D', '#D7263D']
bars = ax.bar(range(len(subgroups)), g_values, yerr=errors, capsize=5, color=colors, edgecolor='black', linewidth=0.5)
ax.set_xticks(range(len(subgroups)))
ax.set_xticklabels(subgroups, fontsize=8.5)
ax.set_ylabel("Hedges' g (kích thước hiệu ứng)")
ax.set_title('Biểu đồ 4. Hiệu lực của Can thiệp Sức khỏe Tâm thần Số (DMHI) cho lo âu xã hội\n'
             'Walder et al. 2025, Bảng 2, tr. 18 (Meta-analysis 21 RCT, JMIR)', fontsize=10.5)
ax.axhline(y=0.2, color='gray', linestyle=':', linewidth=0.7, alpha=0.7)
ax.axhline(y=0.5, color='gray', linestyle=':', linewidth=0.7, alpha=0.7)
ax.axhline(y=0.8, color='gray', linestyle=':', linewidth=0.7, alpha=0.7)
ax.text(5.4, 0.21, 'Nhỏ', fontsize=8, color='gray')
ax.text(5.4, 0.51, 'Trung bình', fontsize=8, color='gray')
ax.text(5.4, 0.81, 'Lớn', fontsize=8, color='gray')
for i, g in enumerate(g_values):
    ax.text(i, g + 0.05 if i != 5 else g + 0.08, f'{g:.3f}', ha='center', fontsize=9, fontweight='bold')
ax.set_ylim(0, 1.5)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart4_DMHI_Walder.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 4 OK')

# ============================================================
# Chart 5: CAMS — Walkup et al. 2008 NEJM (cited in Zugman 2024 AJP tr 191)
# ============================================================
fig, ax = plt.subplots(figsize=(7, 4.2), dpi=150)
arms = ['Placebo', 'SSRI đơn\n(Sertraline)', 'CBT đơn', 'CBT + SSRI\n(kết hợp)']
response = [23.7, 54.9, 59.7, 80.7]
colors = ['#BDBDBD', '#F4A261', '#457B9D', '#06A77D']
bars = ax.bar(arms, response, color=colors, edgecolor='black', linewidth=0.5)
ax.set_ylabel('Tỷ lệ đáp ứng điều trị (%)')
ax.set_title('Biểu đồ 5. Đáp ứng điều trị nghiên cứu CAMS (n = 488 trẻ 7–17 tuổi)\n'
             'Walkup et al. 2008, NEJM 359:2753-66 — trích dẫn trong Zugman 2024, tr. 191', fontsize=10.5)
ax.set_ylim(0, 100)
for b, v in zip(bars, response):
    ax.text(b.get_x() + b.get_width()/2, v + 1.5, f'{v:.1f}%', ha='center', fontsize=10, fontweight='bold')
ax.axhline(y=50, color='gray', linestyle='--', linewidth=0.7, alpha=0.6)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart5_CAMS_Walkup.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 5 OK')

# ============================================================
# Chart 6: Mobile App CBT (Maya) HAM-A — Bress et al. 2024 JAMA Bảng tr 7
# ============================================================
fig, ax = plt.subplots(figsize=(7, 4.2), dpi=150)
weeks = ['Tuần 0\n(ban đầu)', 'Tuần 3\n(giữa)', 'Tuần 6\n(kết thúc)', 'Tuần 12\n(theo dõi)']
diff = [0, -3.20, -5.64, -5.67]
d_cohen = [0, 0.64, 0.94, 1.04]
ci_lo = [0, -4.76, -7.23, -7.29]
ci_hi = [0, -1.64, -4.05, -4.04]
errors = [[abs(d - lo) for d, lo in zip(diff, ci_lo)], [abs(hi - d) for d, hi in zip(diff, ci_hi)]]

ax.errorbar(weeks, diff, yerr=errors, marker='o', markersize=10, linewidth=2,
            color='#1D3557', ecolor='#457B9D', capsize=8, capthick=1.5)
ax.axhline(y=0, color='gray', linewidth=0.5)
ax.set_ylabel('Chênh lệch HAM-A so với ban đầu')
ax.set_title('Biểu đồ 6. Hiệu quả ứng dụng Maya (CBT-app) trên thang HAM-A\n'
             'Bress et al. 2024, JAMA Network Open, Bảng 2 — RCT n = 59, tất cả p < 0,001', fontsize=10.5)
for i, (d, dc) in enumerate(zip(diff, d_cohen)):
    if i == 0:
        ax.text(i, d + 0.4, 'Ban đầu', ha='center', fontsize=9)
    else:
        ax.text(i, d - 0.6, f'{d}\n(d = {dc})', ha='center', fontsize=9, fontweight='bold')
ax.set_ylim(-9, 1.5)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart6_Maya_App_Bress.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 6 OK')

# ============================================================
# Chart 7: BMC NMA Li 2025 — SUCRA all interventions
# ============================================================
fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=150)
methods = ['ACT', 'CBT', 'VRET', 'PE', 'Đối chứng']
sucra = [0.69, 0.66, 0.51, 0.51, 0.10]
n_rct = [6, 16, 6, 2, '—']
colors = ['#06A77D', '#1D3557', '#457B9D', '#F1A208', '#BDBDBD']
bars = ax.bar(methods, sucra, color=colors, edgecolor='black', linewidth=0.5)
ax.set_ylabel('SUCRA (xác suất xếp hạng)')
ax.set_title('Biểu đồ 7. Xếp hạng can thiệp cho rối loạn lo âu trẻ — NMA Bayesian\n'
             'Li et al. 2025, BMC Psychiatry 25:809, Bảng kết quả tr. 6 (30 RCT, n = 1.711)', fontsize=10.5)
ax.set_ylim(0, 0.85)
for i, (b, s, n) in enumerate(zip(bars, sucra, n_rct)):
    label = f'SUCRA {s:.2f}\n({n} RCT)' if isinstance(n, int) else f'SUCRA {s:.2f}'
    ax.text(b.get_x() + b.get_width()/2, s + 0.015, label, ha='center', fontsize=9, fontweight='bold')
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart7_BMC_NMA_Li.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 7 OK')

# ============================================================
# Chart 8: So sánh 3 vùng địa lý (Việt Nam vs Châu Á vs Âu-Úc-Mỹ)
# ============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.5), dpi=150)
regions = ['Việt Nam', 'Châu Á\n(ngoài VN)', 'Âu - Úc - Mỹ']
n_papers = [1, 7, 7]
n_rct = [0, 4, 4]
colors = ['#D7263D', '#F1A208', '#06A77D']

bars1 = ax1.bar(regions, n_papers, color=colors, edgecolor='black', linewidth=0.5)
ax1.set_ylabel('Số bài nghiên cứu')
ax1.set_title('Số bài can thiệp theo vùng', fontsize=11)
for b, v in zip(bars1, n_papers):
    ax1.text(b.get_x() + b.get_width()/2, v + 0.15, str(v), ha='center', fontsize=11, fontweight='bold')
ax1.set_ylim(0, 9)
ax1.spines['top'].set_visible(False); ax1.spines['right'].set_visible(False)

bars2 = ax2.bar(regions, n_rct, color=colors, edgecolor='black', linewidth=0.5)
ax2.set_ylabel('Số RCT cho VTN')
ax2.set_title('Số RCT cho vị thành niên theo vùng', fontsize=11)
for b, v in zip(bars2, n_rct):
    txt = '0\n(không có)' if v == 0 else str(v)
    ax2.text(b.get_x() + b.get_width()/2, v + 0.1, txt, ha='center', fontsize=10, fontweight='bold',
             color='red' if v == 0 else 'black')
ax2.set_ylim(0, 6)
ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)

plt.suptitle('Biểu đồ 8. Phân bố nghiên cứu can thiệp tâm lý cho RLLA theo vùng — Tổng hợp 15 bài',
             fontsize=11, y=1.02)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'chart8_3regions.png'), dpi=180, bbox_inches='tight')
plt.close()
print('Chart 8 OK')

print(f'\nALL 8 CHARTS saved to: {OUT}')
