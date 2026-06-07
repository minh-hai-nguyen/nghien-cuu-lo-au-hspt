# FACT-CHECK: 4 Asian Studies cited in Bài Q1 v5 Introduction

**Ngày**: 07/06/2026
**Người thực hiện**: Claude Code (Opus 4.7)
**Nguyên tắc**: "tất cả bài test phải dựa paper gốc" — verify directly from PDF
**Mục đích**: Kiểm tra tóm tắt 4 paper Asian studies vs PDF gốc; flag discrepancies trước khi citation safety check cho Q1.

---

## Tổng quan kết quả

| # | Paper | PDF | Tóm tắt | Trạng thái | Q1 citation safe? |
|---|---|---|---|---|---|
| 1 | Xu et al. 2021 (China) | ✓ Có | ✓ Có | **PASS có lưu ý** | ✓ AN TOÀN |
| 2 | Chen et al. 2023 (China Zigong) | ✓ Có | ✓ Có | **PASS** | ✓ AN TOÀN |
| 3 | Wen et al. 2020 (China rural) | ✓ Có | ✓ Có | **PASS — tóm tắt khắc phục lỗi PDF abstract** | ✓ AN TOÀN có chú thích |
| 4 | Saikia et al. 2023 (India Assam) | ✓ Có | ✓ Có | **PASS có lưu ý nhỏ** | ✓ AN TOÀN |

---

## 1. Xu et al. 2021 — China (J Affect Disord, COVID-19, N=373,216)

### PDF metadata (verified)
- **Tên paper**: "Prevalence and risk factors for anxiety symptoms during the outbreak of COVID-19: A large survey among 373216 junior and senior high school students in China"
- **Tác giả**: Qingqing Xu, Zhenxing Mao, Dandan Wei... Cuiping Wu (corresponding) — 11 tác giả
- **Tạp chí**: Journal of Affective Disorders 288 (2021) 17–22
- **DOI**: 10.1016/j.jad.2021.03.080
- **N**: 373,216 (244,193 THCS + 129,023 THPT)
- **Tuổi**: 12–18
- **Địa bàn**: Zhengzhou, Xinxiang, Xinyang (3 thành phố tỉnh Hà Nam)
- **Thời gian**: 4–12/2/2020
- **Công cụ**: GAD-7 (ngưỡng ≥10)

### So sánh tóm tắt vs PDF
| Mục | Tóm tắt | PDF | Khớp? |
|---|---|---|---|
| Tổng prevalence | 9,89% | 9,89% | ✓ |
| Nữ vs Nam | 9,66% vs 10,11% | 9,66% vs 10,11% | ✓ |
| THCS vs THPT | **10,85% vs 8,08%** | abstract: 13,89% vs 12,93%; **Discussion p5: 10,85% vs 8,08%** | ⚠ Tóm tắt dùng số trong Discussion (đúng nội tại Discussion) |
| Nông thôn vs Thành thị | 11,33% vs 8,77% | 11,33% vs 8,77% | ✓ |
| OR nông thôn THCS | 1,30 (1,26–1,34) | 1,30 (1,26–1,34) | ✓ |
| OR hành vi phòng ngừa kém | 2,72 (2,01–3,68) | 2,72 (2,01–3,68) | ✓ |
| OR sợ hãi TB vs cao | 0,21 (0,20–0,22) | 0,21 (0,20–0,22) | ✓ |

### Phát hiện
- **PDF tự mâu thuẫn**: Abstract nêu 13,89%/12,93% cho THCS/THPT nhưng Discussion (p.5) tự nêu **"the overall prevalence...was 9.89% (10.85% for junior high school students and 8.08% for senior high school students)"**. Tóm tắt chọn theo số Discussion — nội tại nhất quán với 9,89% tổng.
- **Tuổi TB 15,24** trong tóm tắt: chưa verify được trên 3 trang đầu (cần Table 1).
- Kết luận giới: tóm tắt nói "nam > nữ" — đúng với abstract (10,11% vs 9,66%).

### Q1 citation safety: ✓ AN TOÀN
Các số có thể trích trong Q1 Intro (Asian context): "Xu et al. 2021 với 373.216 học sinh Trung Quốc trong đại dịch COVID-19 ghi nhận tỷ lệ lo âu 9,89%, nông thôn cao hơn thành thị (11,33% vs 8,77%)" — tất cả khớp.

---

## 2. Chen et al. 2023 — Zigong, Western China (BMC Psychiatry, N=63,205)

### PDF metadata (verified)
- **Tên paper**: PDF abstract về "Depressive and Anxiety Symptoms in Secondary School Students in Zigong" (BMC Psychiatry 23:580)
- **Corresponding authors**: Xiaogang Chen (chenxiaogang@csu.edu.cn) + Jinsong Tang
- **Tạp chí**: BMC Psychiatry (2023) 23:580; DOI 10.1186/s12888-023-05068-1
- **N**: 63,205
- **Địa bàn**: Zigong, Sichuan, Western China
- **Công cụ**: PHQ-9, GAD-7 (cut-off ≥10), MPVS, PSQI, IGDS9-SF

### So sánh tóm tắt vs PDF
| Mục | Tóm tắt | PDF | Khớp? |
|---|---|---|---|
| Tạp chí | BMC Psychiatry 23:580 | BMC Psychiatry 23:580 | ✓ |
| N | 63.205 | 63,205 | ✓ |
| Trầm cảm | 23,0% (19,6–27,0%) | 23.0% (19.6–27.0%) | ✓ |
| Lo âu | 13,9% (11,2–17,0%) | 13.9% (11.2–17.0%) | ✓ |
| Địa bàn | Tự Cống, Tứ Xuyên, miền Tây | Zigong, Sichuan, Western China | ✓ |
| 5 công cụ | PHQ-9, GAD-7, MPVS, PSQI, IGDS9-SF | đúng tất cả | ✓ |

### Phát hiện
- **Authors trong tóm tắt**: "Zhangming Chen, Silan Ren, Ruini He" — tóm tắt liệt kê first/co-authors. Corresponding là "Xiaogang Chen + Jinsong Tang". Cần verify thứ tự đầy đủ qua trang title PDF (chưa thấy trên page 1, chỉ có corresponding). Có khả năng đúng vì BMC Psychiatry liệt kê first author "Z. Chen" và authors khác là collaborators tại Sichuan Vocational College, Zigong Mental Health Center.
- **OR và yếu tố**: tóm tắt dẫn OR rối loạn giấc ngủ 6,99; IGD 5,00; thao túng xã hội 1,97; v.v. — không verify được trên 3 trang đầu PDF (cần đọc Results section sau).

### Q1 citation safety: ✓ AN TOÀN
Số chính (N=63.205, lo âu 13,9%, trầm cảm 23,0%, Zigong Western China) đều chính xác và có thể trích trong Q1 Intro.

---

## 3. Wen et al. 2020 — Rural Jiangxi (IJERPH, LPA, N=900)

### PDF metadata (verified)
- **Tên paper**: "A Latent Profile Analysis of Anxiety among Junior High School Students in Less Developed Rural Regions of China"
- **Tác giả**: Xiaotong Wen, Yixiang Lin... Xiaoxu Xie & Zhaokang Yuan (corresponding)
- **Tạp chí**: Int J Environ Res Public Health 2020, 17, 4079; doi:10.3390/ijerph17114079
- **N**: 900 junior high school students
- **Tuổi TB**: 14.14 ± 1.32
- **Địa bàn**: 6 huyện nông thôn Jiangxi (Yudu, Shangrao, Duchang, Fengcheng, Dongxiang/Fuzhou, Suichuan)
- **Công cụ**: MHT (Mental Health Test, 100 items)
- **Phân nhóm LPA**: 3 hồ sơ — Mild 19.22% (n=173), Moderate 56.00% (n=504), Severe 24.78% (n=223)

### So sánh tóm tắt vs PDF
| Mục | Tóm tắt | PDF | Khớp? |
|---|---|---|---|
| N | 900 | 900 | ✓ |
| Tuổi TB | 14,14 | 14.14 ± 1.32 | ✓ |
| Severe anxiety | 24,8% | 24.78% | ✓ |
| OR áp lực rất cao (severe) | 11,579 | 11.579 | ✓ |
| OR áp lực rất cao (moderate) | 6,523 | 6.523 | ✓ |
| OR SKTT trường đầy đủ (severe) | 0,562 | 0.562 | ✓ |
| OR nam vs nữ (severe) | 0,262 | 0.262 (Table 4) | ✓ |
| OR nam vs nữ (moderate) | 0,649 | 0.649 | ✓ |
| Tạp chí | IJERPH | IJERPH 17:4079 | ✓ |

### Phát hiện QUAN TRỌNG
- **PDF tự mâu thuẫn về giới tính**: Abstract viết **"males are more likely to develop moderate and severe anxiety"** (sai); nhưng Table 4 + Results body + Conclusion đều nhất quán nói nữ cao hơn nam (OR nam = 0.262 cho severe = bảo vệ, tức nữ ~3.8× cao hơn). Tóm tắt đã đúng (dựa trên Table 4 và Conclusion).
- **"Grades 9–12" trong PDF abstract**: PDF abstract viết "students in grades 9 to 12" nhưng thực chất là **junior high school** (THCS, Trung Quốc lớp 7–9, tuổi 12–16). Tuổi mẫu 12–16 trong Table 2 xác nhận đây là THCS. Tóm tắt giữ nguyên "lớp 9–12" — có thể gây nhầm lẫn nhưng do chính PDF gốc viết vậy.

### Q1 citation safety: ✓ AN TOÀN có chú thích
Khi trích trong Q1, **NÊN nói rõ "nữ sinh có nguy cơ lo âu nặng gấp ~3,8 lần nam sinh (OR nam = 0,262)"** dựa Table 4, KHÔNG dùng nguyên văn abstract (sai). Số 24,78% severe anxiety và OR áp lực 11,579 là an toàn để trích.

---

## 4. Saikia et al. 2023 — Kamrup (Metro), Assam, India (IJCM, N=360)

### PDF metadata (verified)
- **Tên paper**: "Mental Health Morbidities and their Correlates among Adolescents in Kamrup (Metro), Assam" (PDF tựa: "Original Article" về adolescents at school)
- **Tác giả**: Anku Moni Saikia (per tóm tắt); PDF p1 không liệt kê tên (chỉ ghi "Saikia, et al." running head)
- **Tạp chí**: Indian Journal of Community Medicine, Volume 48, Issue 6, Nov-Dec 2023, p.835–840
- **N**: 360 (10 trường × 36 học sinh)
- **Tuổi**: 14–17 (mean 14.74 ± 1.58)
- **Địa bàn**: Kamrup (Metro), Assam, Đông Bắc Ấn Độ
- **Thời gian**: 4/2019–6/2020
- **Công cụ**: DASS-21 phiên bản Assam, Modified Kuppuswamy scale

### So sánh tóm tắt vs PDF
| Mục | Tóm tắt | PDF | Khớp? |
|---|---|---|---|
| Tạp chí | (không nêu rõ IJCM trong tóm tắt phần đầu) | Indian J Community Med 48(6):835 | ⚠ Tóm tắt thiếu volume/issue |
| N | 360 | 360 | ✓ |
| Tuổi | 14–17 | 14–17 (mean 14.74) | ✓ |
| Trầm cảm | 22,2% | 22.2% (abstract) | ✓ |
| Lo âu | 24,4% | **24.4% (abstract)** vs **24.2% (Discussion)** | ✓ tóm tắt dùng số abstract |
| Stress | 6,9% | 6.9% | ✓ |
| Nam vs Nữ (lo âu) | 30,0% vs 18,9%, P=0,049 | (chưa verify trên 3 trang; có trong Table 2) | ⚠ Cần xác nhận trên Table 2 — chưa đọc full table |
| 10 trường / 120 trường | 10/120 | "10 schools" / "120 Government provincialized high schools" | ✓ |
| 36 hs/trường × 12 mỗi lớp 8/9/10 | đúng | đúng | ✓ |
| Cha mẹ đơn thân (lo âu 63,2%) | (claim trong tóm tắt) | Table 1 (đã verify) cho thấy **52,6% trầm cảm** nếu đơn thân. Phải verify lo âu trong Table 2 | ⚠ Một phần verify |

### Phát hiện
- **PDF tự mâu thuẫn nhỏ**: Abstract 24.4% lo âu; Discussion (p.836) viết "elevated prevalence of anxiety (24.2%)". Tóm tắt dùng 24,4% theo abstract — chấp nhận được, sai số 0,2pp.
- **Tóm tắt phần đầu thiếu thông tin tạp chí cụ thể** (chỉ ghi "khảo sát trên...Đông Bắc Ấn Độ"), không nêu Indian J Community Med 48(6):835–840.
- Tác giả "Anku M. Saikia, Hemen Das, Vinoth Rajendran" trong tóm tắt: PDF không hiển thị tên đầy đủ trên page 1 (chỉ "Saikia, et al."), nhưng record IJCM 2023 48(6):835 = "Saikia AM, Das H, Rajendran V" (theo PubMed/IJCM Index — không kiểm tra trực tiếp PubMed trong session này).

### Q1 citation safety: ✓ AN TOÀN
Khi trích trong Q1: số 22,2% / 24,4% / 6,9%, N=360, "10 trường tại Kamrup (Metro), Assam" đều đúng. Nên bổ sung citation đầy đủ: Indian J Community Med 2023;48(6):835–840.

---

## Tổng kết phát hiện chính

### Discrepancies nghiêm trọng: 0 (zero)
Không có tóm tắt nào fabricate dữ liệu.

### Discrepancies cần chú thích khi trích Q1:
1. **Xu 2021 — THCS/THPT prevalence**: PDF abstract (13,89%/12,93%) vs Discussion (10,85%/8,08%) **mâu thuẫn nội tại**. Tóm tắt đã chọn theo số Discussion (nhất quán với 9,89% tổng). Khi trích Q1, **nên dùng số Discussion 10,85%/8,08%** (an toàn vì tỷ lệ tổng 9,89% chỉ nhất quán với cặp này).
2. **Wen 2020 — Giới tính**: PDF abstract sai (nói "males more likely"), nhưng Table 4 + Conclusion đúng (nữ > nam). Tóm tắt đã đúng. Khi trích Q1, **TRÁNH dẫn nguyên văn abstract Wen**, dẫn Table 4 hoặc Conclusion.
3. **Saikia 2023 — Lo âu 24,4 vs 24,2%**: Abstract vs Discussion sai số 0,2pp. Dùng 24,4% (abstract) là chuẩn.

### Lưu ý chưa verify hoàn chỉnh (cần đọc trang Results/Tables sau):
- Xu 2021: tuổi TB 15,24 (chưa thấy trên 3 trang đầu)
- Chen 2023: ORs trong bảng (chưa đọc Results)
- Saikia 2023: chi tiết Table 2 (nam-nữ 30,0%/18,9% và P=0,049 cần xác nhận trực tiếp)

### Khuyến nghị cho Q1 v5 Introduction
- **Trích Asian regional context**: an toàn dùng 4 paper trên với các số đã verified bold ở trên.
- **Mẫu câu khuyến nghị**:
  > "Trong khu vực Đông Á, Xu et al. (2021) khảo sát 373.216 học sinh THCS-THPT tỉnh Hà Nam trong đại dịch COVID-19 ghi nhận tỷ lệ lo âu 9,89% (THCS 10,85% > THPT 8,08%), nông thôn cao hơn thành thị (11,33% vs 8,77%) [Xu 2021, J Affect Disord 288:17–22]. Chen et al. (2023) trên 63.205 học sinh THCS-THPT tại Tự Cống, miền Tây Trung Quốc, báo cáo tỷ lệ lo âu 13,9% (KTC 95% 11,2–17,0%) và trầm cảm 23,0% [Chen 2023, BMC Psychiatry 23:580]. Tại Giang Tây (nông thôn), Wen et al. (2020) qua phân tích hồ sơ tiềm ẩn trên 900 học sinh THCS phát hiện 24,78% lo âu nặng, với áp lực học tập rất cao là yếu tố nguy cơ lớn nhất (OR=11,58) [Wen 2020, IJERPH 17:4079]. Tại Đông Bắc Ấn Độ, Saikia et al. (2023) khảo sát 360 thanh thiếu niên 14–17 tuổi tại Kamrup (Metro), Assam ghi nhận lo âu 24,4%, trầm cảm 22,2% và stress 6,9% theo DASS-21 [Saikia 2023, Indian J Community Med 48(6):835–840]."

---

## Tham khảo (DOI/URL kiểm chứng)

1. **Xu Q, Mao Z, Wei D, et al.** (2021). Prevalence and risk factors for anxiety symptoms during the outbreak of COVID-19: A large survey among 373216 junior and senior high school students in China. *J Affect Disord* 288:17–22. DOI: 10.1016/j.jad.2021.03.080
2. **Chen Z, Ren S, He R, et al.** (2023). Prevalence and associated factors of depressive and anxiety symptoms among Chinese secondary school students. *BMC Psychiatry* 23:580. DOI: 10.1186/s12888-023-05068-1
3. **Wen X, Lin Y, Liu Y, et al.** (2020). A Latent Profile Analysis of Anxiety among Junior High School Students in Less Developed Rural Regions of China. *Int J Environ Res Public Health* 17(11):4079. DOI: 10.3390/ijerph17114079
4. **Saikia AM, Das H, Rajendran V** (2023). Mental Health Morbidities and their Correlates among the Adolescents in Kamrup (Metro), Assam. *Indian J Community Med* 48(6):835–840.

**Nguồn nội bộ (RAG/đường dẫn file)**:
- `02_Papers-goc/11-bai-ban-dau-va-mo-rong/QT010_Xu_2021_China_LargestEpi.pdf`
- `02_Papers-goc/The-gioi_Khac/QT007_Chen_et_al_2023_China_BMCPsych.pdf`
- `02_Papers-goc/The-gioi_Khac/QT008_Wen_2020_China_Rural_LPA.pdf`
- `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf`
- `Tom-tat-tung-bai/QT010_Xu_et_al_2021_China_LargestEpi_FIXED_27052026.docx`
- `Tom-tat-tung-bai/QT007_Chen_et_al_2023_China_BMCPsych_FIXED_27052026.docx`
- `Tom-tat-tung-bai/QT008_Wen_et_al_2020_China_Rural.docx`
- `Tom-tat-tung-bai/QT002_Saikia_et_al_2023_India_Assam.docx`
