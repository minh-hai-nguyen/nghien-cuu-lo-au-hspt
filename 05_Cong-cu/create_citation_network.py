# -*- coding: utf-8 -*-
"""
Create interactive citation network as static HTML.
Hover over nodes to see paper details.
"""
from pyvis.network import Network
import os

BASE = r"c:\Users\OS\OneDrive\read_books\Lo-au"

# ===== DATA: 11 core papers =====
CORE_PAPERS = {
    "Jenkins2023": {
        "label": "Jenkins\n(2023)",
        "title": "<b>Jenkins et al. (2023)</b><br>"
                 "<i>Depression and anxiety among multiethnic middle school students</i><br>"
                 "📖 International Journal of Social Psychiatry, Q2, IF~2.2<br>"
                 "🔬 N=75, Hoa Kỳ, PHQ-9A + GAD-10 + dân tộc học<br>"
                 "📊 Lo âu: 50.6%, Trầm cảm: 44%<br>"
                 "⭐ Phương pháp hỗn hợp sáng tạo. Nữ > nam (p=0.002).<br>"
                 "Bạo lực giới, COVID-19, phân biệt chủng tộc là yếu tố chính.",
        "color": "#E91E63", "size": 30,
    },
    "Saikia2023": {
        "label": "Saikia\n(2023)",
        "title": "<b>Saikia et al. (2023)</b><br>"
                 "<i>Mental Health Morbidities among Adolescents in Kamrup, Assam</i><br>"
                 "📖 Indian J Community Medicine, Q3, IF~0.8<br>"
                 "🔬 N=360, Ấn Độ, DASS-21<br>"
                 "📊 Lo âu: 24.4%, Trầm cảm: 22.2%, Stress: 6.9%<br>"
                 "⭐ NAM > nữ về lo âu (30% vs 18.9%). Nghiên cứu đầu tiên Đông Bắc Ấn Độ.",
        "color": "#FF9800", "size": 28,
    },
    "Mandaknalli2021": {
        "label": "Mandaknalli\n(2021)",
        "title": "<b>Mandaknalli & Malusare (2021)</b><br>"
                 "<i>Prevalence of anxiety among municipality school area</i><br>"
                 "📖 MedPulse Int J Psychology, không IF<br>"
                 "🔬 N=450, Ấn Độ, GAD-7<br>"
                 "📊 Lo âu: ~100% (nhẹ 49.4%, TB 43.3%, nặng 7.3%)<br>"
                 "⭐ Tỷ lệ cao bất thường — điểm cắt quá thấp?",
        "color": "#FF9800", "size": 20,
    },
    "NSCH2020": {
        "label": "NSCH\n(2020)",
        "title": "<b>NSCH — National Survey of Children's Health (2020)</b><br>"
                 "<i>Dữ liệu quốc gia Hoa Kỳ về sức khỏe trẻ em</i><br>"
                 "📖 HRSA/Census Bureau, uy tín rất cao<br>"
                 "🔬 N=55,162, Hoa Kỳ, khảo sát quốc gia<br>"
                 "📊 Lo âu: 16.1%, Trầm cảm: 8.4% (chẩn đoán)<br>"
                 "⭐ Lo âu tăng 61% (2016-2023). Nữ 20.1% vs nam 12.3%.",
        "color": "#E91E63", "size": 35,
    },
    "Alharbi2019": {
        "label": "Alharbi\n(2019)",
        "title": "<b>Alharbi et al. (2019)</b><br>"
                 "<i>Depression and anxiety among high school student at Qassim</i><br>"
                 "📖 J Family Med Prim Care, Q3, IF~1.5<br>"
                 "🔬 N=1,245, Ả Rập Saudi, PHQ-9 + GAD-7<br>"
                 "📊 Lo âu: 63.5%, Trầm cảm: 74%<br>"
                 "⭐ Tỷ lệ cực cao. Nữ > nam (P<0.001). Bài đầu tiên dùng PHQ-9/GAD-7 ở Saudi.",
        "color": "#4CAF50", "size": 30,
    },
    "Nakie2022": {
        "label": "Nakie\n(2022)",
        "title": "<b>Nakie et al. (2022)</b><br>"
                 "<i>Depression, anxiety, stress among high school students, Ethiopia</i><br>"
                 "📖 BMC Psychiatry, Q1, IF~4.4<br>"
                 "🔬 N=849, Ethiopia, DASS-21<br>"
                 "📊 Lo âu: 66.7%, Trầm cảm: 41.4%, Stress: 52.2%<br>"
                 "⭐ Khat (AOR=5.6), hút thuốc (AOR=4.8). Đầu tiên ở châu Phi đánh giá cả 3.",
        "color": "#9C27B0", "size": 32,
    },
    "Chen2023": {
        "label": "Chen\n(2023)",
        "title": "<b>Chen et al. (2023)</b><br>"
                 "<i>Depressive and anxiety symptoms among Chinese secondary school students</i><br>"
                 "📖 BMC Psychiatry, Q1, IF~4.4<br>"
                 "🔬 N=63,205, Trung Quốc (Tự Cống), PHQ-9 + GAD-7 + PSQI + IGDS9-SF<br>"
                 "📊 Lo âu: 13.9%, Trầm cảm: 23.0%<br>"
                 "⭐ Bắt nạt, giấc ngủ kém, nghiện game là yếu tố nguy cơ. Đầu tiên từ miền Tây TQ.",
        "color": "#FFC107", "size": 35,
    },
    "Wen2020": {
        "label": "Wen\n(2020)",
        "title": "<b>Wen et al. (2020)</b><br>"
                 "<i>Latent Profile Analysis of Anxiety, rural China</i><br>"
                 "📖 Int J Environ Res Public Health, Q2, IF~4.6<br>"
                 "🔬 N=900, Trung Quốc (nông thôn Giang Tây), LPA<br>"
                 "📊 Lo âu nặng: 24.78%<br>"
                 "⭐ NAM > nữ lo âu. Hỗ trợ tâm thần tại trường là yếu tố bảo vệ. Đầu tiên dùng LPA.",
        "color": "#FFC107", "size": 28,
    },
    "Qiu2022": {
        "label": "Qiu\n(2022)",
        "title": "<b>Qiu et al. (2022)</b><br>"
                 "<i>Parenting Style and Resilience with Depression/Anxiety, China</i><br>"
                 "📖 Frontiers in Psychology, Q2, IF~3.8<br>"
                 "🔬 N=2,079, Trung Quốc (Hợp Phì), EMBU + CES-D + SAS + SRSMSS<br>"
                 "📊 Lo âu: 13.4%, Trầm cảm: 26.0%<br>"
                 "⭐ Nuôi dạy tiêu cực: OR=1.82-2.01. Phục hồi thấp: OR=6.74. Tương tác không ý nghĩa.",
        "color": "#FFC107", "size": 28,
    },
    "Xu2021": {
        "label": "Xu\n(2021)",
        "title": "<b>Xu et al. (2021)</b><br>"
                 "<i>Anxiety during COVID-19: 373,216 students in China</i><br>"
                 "📖 J Affective Disorders, Q1, IF~6.6<br>"
                 "🔬 N=373,216, Trung Quốc (Hà Nam), GAD-7<br>"
                 "📊 Lo âu: 9.89%<br>"
                 "⭐ Lớn nhất toàn cầu. NAM > nữ (10.11% vs 9.66%). Nông thôn cao nhất 12.8%.",
        "color": "#FFC107", "size": 40,
    },
    "Bhardwaj2020": {
        "label": "Bhardwaj\n(2020)",
        "title": "<b>Bhardwaj et al. (2020)</b><br>"
                 "<i>Depression, Anxiety & Stress, Government schools Chandigarh</i><br>"
                 "📖 J IPHA Chandigarh, tạp chí địa phương<br>"
                 "🔬 N=288, Ấn Độ, DASS-21<br>"
                 "📊 Lo âu: 73.3%, Trầm cảm: 64.9%, Stress: 74.7%<br>"
                 "⭐ Tỷ lệ cao nhất trong 11 bài. Trường công lập, kinh tế thấp.",
        "color": "#FF9800", "size": 22,
    },
}

# ===== CITED BY our papers (references they share) =====
EXTERNAL_PAPERS = {
    "Racine2021": {
        "label": "Racine\n(2021)",
        "title": "<b>Racine et al. (2021)</b><br><i>Global Prevalence of Depression/Anxiety in Children During COVID-19: Meta-analysis</i><br>📖 JAMA Pediatrics, Q1, IF~26.0<br>⭐ 1/4 thanh thiếu niên có trầm cảm, 1/5 có lo âu trong COVID-19.",
        "color": "#90CAF9", "size": 25,
    },
    "Polanczyk2015": {
        "label": "Polanczyk\n(2015)",
        "title": "<b>Polanczyk et al. (2015)</b><br><i>Worldwide prevalence of mental disorders in children: Meta-analysis</i><br>📖 J Child Psychol Psychiatry, Q1<br>⭐ Tỷ lệ rối loạn tâm thần toàn cầu ở trẻ em: 13.4%",
        "color": "#90CAF9", "size": 22,
    },
    "Sandal2017": {
        "label": "Sandal\n(2017)",
        "title": "<b>Sandal et al. (2017)</b><br><i>Depression, anxiety, stress among school going adolescent in Chandigarh</i><br>📖 J Family Med Prim Care<br>⭐ DASS-21 tại Chandigarh: trầm cảm 65.53%, lo âu 80.85%",
        "color": "#90CAF9", "size": 20,
    },
    "Kumar2019": {
        "label": "Kumar\n(2019)",
        "title": "<b>Kumar et al. (2019)</b><br><i>Depression, anxiety, stress among school going adolescents in Delhi</i><br>📖 Int J Community Med Public Health<br>⭐ Nghiên cứu tại Delhi, Ấn Độ",
        "color": "#90CAF9", "size": 18,
    },
    "AlGelban2007": {
        "label": "Al-Gelban\n(2007)",
        "title": "<b>Al-Gelban (2007)</b><br><i>Depression, anxiety, stress among Saudi adolescent school boys</i><br>📖 J R Soc Promot Health<br>⭐ 38.2% trầm cảm, 48.9% lo âu, 35.5% stress ở nam sinh Saudi",
        "color": "#90CAF9", "size": 18,
    },
    "GBD2020": {
        "label": "GBD Study\n(2020)",
        "title": "<b>GBD 2020 — Global Burden of Disease</b><br><i>Mental disorders burden in India/globally</i><br>📖 Lancet Psychiatry, Q1<br>⭐ Gánh nặng bệnh tâm thần toàn cầu",
        "color": "#90CAF9", "size": 25,
    },
    "Spitzer2006": {
        "label": "Spitzer\n(2006)",
        "title": "<b>Spitzer et al. (2006)</b><br><i>GAD-7: A brief measure for generalized anxiety disorder</i><br>📖 Arch Internal Medicine, Q1<br>⭐ Bài gốc phát triển công cụ GAD-7",
        "color": "#B0BEC5", "size": 30,
    },
    "Kroenke2001": {
        "label": "Kroenke\n(2001)",
        "title": "<b>Kroenke et al. (2001)</b><br><i>PHQ-9: Validity of a brief depression severity measure</i><br>📖 J Gen Intern Med, Q1<br>⭐ Bài gốc phát triển công cụ PHQ-9",
        "color": "#B0BEC5", "size": 30,
    },
    "Lovibond1995": {
        "label": "Lovibond\n(1995)",
        "title": "<b>Lovibond & Lovibond (1995)</b><br><i>DASS: Structure of negative emotional states</i><br>📖 Behav Res Ther, Q1<br>⭐ Bài gốc phát triển công cụ DASS-21",
        "color": "#B0BEC5", "size": 28,
    },
    # Papers citing our 11
    "Girma2021": {
        "label": "Girma\n(2021)",
        "title": "<b>Girma et al. (2021)</b><br><i>Depression and its determinants among adolescents in Jimma, Ethiopia</i><br>📖 PLOS ONE<br>⭐ Trích dẫn Nakie 2022, Alharbi 2019",
        "color": "#C8E6C9", "size": 18,
    },
    "Karachi2022": {
        "label": "Prevalence\nKarachi (2022)",
        "title": "<b>Anxiety/depression in high school students of Karachi, Pakistan (2022)</b><br>📖 PubMed<br>⭐ Trích dẫn Alharbi 2019",
        "color": "#C8E6C9", "size": 16,
    },
    "Sudan2025": {
        "label": "Anxiety\nSudan (2025)",
        "title": "<b>Anxiety scores and academic performance, Northern Sudan (2025)</b><br>📖 PMC<br>⭐ Trích dẫn Alharbi 2019",
        "color": "#C8E6C9", "size": 16,
    },
    "QassimCOVID2023": {
        "label": "Qassim\nCOVID (2023)",
        "title": "<b>Depression/Anxiety among Qassim University Students During COVID-19 (2023)</b><br>📖 PMC<br>⭐ Trích dẫn Alharbi 2019",
        "color": "#C8E6C9", "size": 16,
    },
    "PLOSONE2025": {
        "label": "DAS\nSecondary (2025)",
        "title": "<b>Prevalence and determinants of DAS among secondary school students (2025)</b><br>📖 PLOS ONE<br>⭐ Trích dẫn Jenkins 2023, Nakie 2022",
        "color": "#C8E6C9", "size": 18,
    },
}

# ===== EDGES =====
EDGES = [
    # Internal citations (11 papers citing each other via shared references)
    ("Nakie2022", "Alharbi2019", "Nakie trích dẫn Alharbi"),
    ("Nakie2022", "Saikia2023", "Cùng nghiên cứu DASS-21"),
    ("Saikia2023", "Sandal2017", "Saikia trích dẫn Sandal (Chandigarh)"),
    ("Saikia2023", "Kumar2019", "Saikia trích dẫn Kumar (Delhi)"),
    ("Bhardwaj2020", "Sandal2017", "Bhardwaj trích dẫn Sandal (cùng Chandigarh)"),
    ("Alharbi2019", "AlGelban2007", "Alharbi trích dẫn Al-Gelban (Saudi)"),

    # Shared instrument citations
    ("Alharbi2019", "Spitzer2006", "Sử dụng GAD-7"),
    ("Alharbi2019", "Kroenke2001", "Sử dụng PHQ-9"),
    ("Chen2023", "Spitzer2006", "Sử dụng GAD-7"),
    ("Chen2023", "Kroenke2001", "Sử dụng PHQ-9"),
    ("Xu2021", "Spitzer2006", "Sử dụng GAD-7"),
    ("Mandaknalli2021", "Spitzer2006", "Sử dụng GAD-7"),
    ("Saikia2023", "Lovibond1995", "Sử dụng DASS-21"),
    ("Nakie2022", "Lovibond1995", "Sử dụng DASS-21"),
    ("Bhardwaj2020", "Lovibond1995", "Sử dụng DASS-21"),

    # GBD citations
    ("Saikia2023", "GBD2020", "Trích dẫn GBD"),
    ("Nakie2022", "GBD2020", "Trích dẫn GBD"),
    ("Mandaknalli2021", "Polanczyk2015", "Trích dẫn meta-analysis"),

    # COVID-19 related
    ("Xu2021", "Racine2021", "Cùng chủ đề COVID-19 + lo âu"),
    ("Jenkins2023", "Racine2021", "Bối cảnh COVID-19"),

    # Papers citing our 11
    ("PLOSONE2025", "Jenkins2023", "Trích dẫn Jenkins 2023"),
    ("PLOSONE2025", "Nakie2022", "Trích dẫn Nakie 2022"),
    ("Girma2021", "Nakie2022", "Trích dẫn Nakie 2022"),
    ("Girma2021", "Alharbi2019", "Trích dẫn Alharbi 2019"),
    ("Karachi2022", "Alharbi2019", "Trích dẫn Alharbi 2019"),
    ("Sudan2025", "Alharbi2019", "Trích dẫn Alharbi 2019"),
    ("QassimCOVID2023", "Alharbi2019", "Trích dẫn Alharbi 2019"),

    # Thematic connections between our 11
    ("Chen2023", "Xu2021", "Cùng Trung Quốc, cùng GAD-7"),
    ("Chen2023", "Wen2020", "Cùng Trung Quốc, so sánh vùng"),
    ("Qiu2022", "Wen2020", "Cùng Trung Quốc, cùng LPA"),
    ("Xu2021", "Wen2020", "Cùng Trung Quốc, nam > nữ"),
    ("Saikia2023", "Bhardwaj2020", "Cùng Ấn Độ, cùng DASS-21"),
    ("Mandaknalli2021", "Saikia2023", "Cùng Ấn Độ"),
    ("Jenkins2023", "NSCH2020", "Cùng Hoa Kỳ"),
]

# ===== BUILD NETWORK =====
net = Network(
    height="800px", width="100%",
    bgcolor="#1a1a2e", font_color="white",
    directed=True, notebook=False,
    cdn_resources="remote",
)

net.barnes_hut(
    gravity=-3000, central_gravity=0.3,
    spring_length=200, spring_strength=0.01,
    damping=0.09
)

# Add core paper nodes
for pid, data in CORE_PAPERS.items():
    net.add_node(
        pid, label=data["label"], title=data["title"],
        color=data["color"], size=data["size"],
        borderWidth=3, borderWidthSelected=5,
        font={"size": 11, "color": "white", "face": "Arial"},
        shape="dot",
    )

# Add external nodes
for pid, data in EXTERNAL_PAPERS.items():
    net.add_node(
        pid, label=data["label"], title=data["title"],
        color=data["color"], size=data["size"],
        borderWidth=1,
        font={"size": 9, "color": "#ccc", "face": "Arial"},
        shape="dot",
    )

# Add edges
for src, dst, label in EDGES:
    net.add_edge(src, dst, title=label, color="#555", width=1.5, arrows="to")

# Add legend as HTML
net.set_options("""
{
  "interaction": {
    "hover": true,
    "tooltipDelay": 100,
    "navigationButtons": true
  },
  "physics": {
    "barnesHut": {
      "gravitationalConstant": -3000,
      "centralGravity": 0.3
    }
  }
}
""")

# Save
output_path = os.path.join(BASE, "DocFiles", "citation_network.html")
net.save_graph(output_path)

# Add custom header to HTML
with open(output_path, 'r', encoding='utf-8') as f:
    html = f.read()

legend_html = """
<div style="position:fixed;top:10px;left:10px;background:rgba(0,0,0,0.85);color:white;padding:15px;border-radius:10px;font-family:Arial;font-size:12px;z-index:1000;max-width:300px;">
<h3 style="margin:0 0 10px 0;color:#fff;">🔬 MẠNG LƯỚI TRÍCH DẪN</h3>
<p style="margin:2px 0;"><span style="color:#E91E63;">●</span> Hoa Kỳ (USA)</p>
<p style="margin:2px 0;"><span style="color:#FF9800;">●</span> Ấn Độ (India)</p>
<p style="margin:2px 0;"><span style="color:#4CAF50;">●</span> Ả Rập Saudi</p>
<p style="margin:2px 0;"><span style="color:#9C27B0;">●</span> Ethiopia</p>
<p style="margin:2px 0;"><span style="color:#FFC107;">●</span> Trung Quốc (China)</p>
<p style="margin:2px 0;"><span style="color:#90CAF9;">●</span> Bài tham chiếu chính</p>
<p style="margin:2px 0;"><span style="color:#B0BEC5;">●</span> Công cụ đo lường gốc</p>
<p style="margin:2px 0;"><span style="color:#C8E6C9;">●</span> Bài trích dẫn 11 bài</p>
<hr style="border-color:#555;">
<p style="margin:2px 0;font-size:10px;">⬤ Kích thước node = cỡ mẫu tương đối</p>
<p style="margin:2px 0;font-size:10px;">🖱 Hover để xem chi tiết</p>
<p style="margin:2px 0;font-size:10px;">🔍 Cuộn chuột để zoom</p>
<p style="margin:2px 0;font-size:10px;">✋ Kéo để di chuyển</p>
</div>
"""

html = html.replace("<body>", f"<body>{legend_html}")
html = html.replace("<title>", "<title>Mạng lưới trích dẫn — 11 nghiên cứu lo âu/trầm cảm học sinh | ")

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"DONE: {output_path}")
print(f"Nodes: {len(CORE_PAPERS) + len(EXTERNAL_PAPERS)}")
print(f"Edges: {len(EDGES)}")
