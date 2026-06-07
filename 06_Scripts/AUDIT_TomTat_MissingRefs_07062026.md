# AUDIT — Mục Tài liệu tham khảo trong `Tom-tat-tung-bai/`

**Ngày audit:** 07/06/2026
**Phạm vi:** `Tom-tat-tung-bai/*.docx` (loại trừ `_Archive_Scripts/`, `BANG_*.docx`, `00_Mẫu...docx`, `bai-bao-khgdvn/`)
**Phương pháp:** scan READ-ONLY bằng `python-docx` + regex; phát hiện heading (`Tham khảo` / `References` / `TLTK` / `Tài liệu tham khảo` / `Bibliography` — bao gồm cả tiền tố số `7. Tham khảo` v.v.); fallback dò DOI / PMID / URL toàn văn.
**Script:** `06_Scripts/_audit_missing_refs_scan.py`
**Dữ liệu thô:** `06_Scripts/_audit_missing_refs_results.json`

---

## 1. Tổng quan

| Phân loại | Số lượng | Tỷ lệ |
|---|---:|---:|
| Tổng tóm tắt được scan | **108** | 100% |
| ✅ HAS_REF_SECTION (có mục Tham khảo + nội dung) | **11** | 10.2% |
| 🟡 HAS_INLINE_REFS (có DOI/PMID/URL nhưng không có mục riêng) | **50** | 46.3% |
| 🔴 NO_REFS (không có cả mục lẫn DOI/PMID/URL) | **47** | 43.5% |
| Lỗi mở file | 0 | 0% |

**Chia theo QT/VN:**
- QT: 31 🔴 + 42 🟡 + ~11 ✅
- VN: 16 🔴 + 8 🟡 + 1 ✅ (VN020 FIXED)

**Kết luận hệ thống:**
Xác nhận phát hiện của Audit Agent 1 (40% mẫu thiếu Tham khảo) — thực tế chính xác là **43.5% (47/108) hoàn toàn không có dấu vết tham khảo**, và thêm **46.3% (50/108) chỉ có 1 DOI lẻ** (thường là DOI bài gốc nêu trong câu mở đầu — không phải danh mục TLTK thật). Cộng dồn, **89.8% (97/108) tóm tắt KHÔNG có mục Tham khảo chính quy**. Chỉ 11 file đã được sửa hôm nay (07/06/2026, hậu tố `_FIXED_07062026`) đạt chuẩn.

---

## 2. Danh sách 🔴 NO_REFS (ƯU TIÊN CAO NHẤT — cần bổ sung TLTK từ đầu)

47 file — không có cả heading lẫn DOI/PMID/URL.

### QT (31)
1. `QT001_Jenkins_et_al_2023_USA.docx`
2. `QT002_Saikia_et_al_2023_India_Assam.docx`
3. `QT003_Mandaknalli_Malusare_2021_FIXED_27052026.docx`
4. `QT004_NSCH_2020_USA.docx`
5. `QT006_Nakie_et_al_2022_Ethiopia.docx`
6. `QT007_Chen_et_al_2023_China_BMCPsych_FIXED_27052026.docx`
7. `QT008_Wen_et_al_2020_China_Rural.docx`
8. `QT009_Qiu_et_al_2022_China_Parenting.docx`
9. `QT010_Xu_et_al_2021_China_LargestEpi_FIXED_27052026.docx`
10. `QT011_Bhardwaj_et_al_2020_India_Chandigarh.docx`
11. `QT018_Salari_SAD_Prevalence_Global_SR_MA_2024.docx`
12. `QT019_Shibuya_SchoolMH_Literacy_3Asian_2025.docx`
13. `QT020_Liu_CBT_Delivery_GAD_NMA_2025.docx` *(đã có bản FIXED_07062026 — file này có thể archive)*
14. `QT024_WHO_Europe_2025_LancetRegional.docx`
15. `QT026_UK_NHS_Parliament_2025.docx`
16. `QT027_NatureHumanBehav_Fassi_SocialMedia_2025.docx`
17. `QT028_AJP_Zugman_PediatricAnxiety_2024.docx`
18. `QT029_BMC_Li_CBT_NMA_2025.docx` *(đã có bản FIXED — archive bản này)*
19. `QT030_GBD_Trends_10-24y_2025.docx`
20. `QT031_Islam_59Countries_2025.docx`
21. `QT032_Ireland_Fitzgerald_MyWorld_2024.docx`
22. `QT033_JAMA_SchmidtPersson_ScreenMedia_2024.docx`
23. `QT034_Korea_Cho_MH_Trends_NatSciRep_2024.docx`
24. `QT052_Mindfulness_NatureMH_IPD_MA_2023.docx`
25. `QT053_Pharmacotherapy_Anxiety_Review_Frontiers_2020.docx`
26. `QT054_ChronicStress_Neuroinflammation_FrontPsych_2023.docx`
27. `QT055_Parade_Epigenetics_ChildhoodMaltreatment_TranslPsych_2021.docx`
28. `QT056_Jiang_Microbiota_GutBrain_Anxiety_FrontNeuro_2024.docx`
29. `QT057_NeuralCircuits_Anxiety_FrontNeuralCirc_2025.docx`
30. `QT058_Guo_Neuroinflammation_Neuromodulation_TranslPsych_2022.docx`
31. `QT073_He_2025_PowerUpCBTD_China_JAD.docx`

### VN (16)
1. `VN002_VNAMHS_2022_National.docx`
2. `VN003_Pham_2024_Hue_SocialSupport.docx`
3. `VN004_NguyenThiVan_2020_STAI_TPHCM.docx`
4. `VN005_TranNguyenNgoc_2018_LuanAn_ThuGian_GAD.docx`
5. `VN006_TranThiMyLuong_2020_DASS42_THPTChuyen.docx`
6. `VN014_HoangTrungHoc_2025_VN_COVID.docx`
7. `VN015_NgoAnhVinh_2024_LangSon_DTTS.docx`
8. `VN016_BaoQuyen_2025_YHCD.docx`
9. `VN017_NguyenDanhLam_2022_YHVN_FIXED_27052026.docx`
10. `VN019_HoThiTrucQuynh_2025_Hue_TLH.docx`
11. `VN022_UNICEF_VN_2022_SchoolFactors.docx`
12. `VN024_NguyenThanhTruyen_VinhLong_YHVN_2024.docx`
13. `VN025_PhamThiNgoc_HaiPhong_VinhBao_2024.docx`
14. `VN026_TranDucSi_LongAn_PNT_2025.docx`
15. `VN027_Dinh_SchoolFactors_VN_2021.docx`
16. `VN028_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.docx`

### Top 10 ưu tiên fix ngay (theo trọng số: VN gốc trong nước + nguồn cao IF)
1. `VN002_VNAMHS_2022_National.docx` — nguồn quốc gia nền tảng
2. `QT018_Salari_SAD_Prevalence_Global_SR_MA_2024.docx` — SR/MA toàn cầu
3. `QT024_WHO_Europe_2025_LancetRegional.docx` — Lancet
4. `QT030_GBD_Trends_10-24y_2025.docx` — GBD core
5. `QT052_Mindfulness_NatureMH_IPD_MA_2023.docx` — Nature MH IPD-MA
6. `QT028_AJP_Zugman_PediatricAnxiety_2024.docx` — Am J Psychiatry
7. `QT033_JAMA_SchmidtPersson_ScreenMedia_2024.docx` — JAMA
8. `QT001_Jenkins_et_al_2023_USA.docx` — bài đầu danh sách, có thể là template lỗi gốc
9. `QT004_NSCH_2020_USA.docx` — NSCH benchmark
10. `VN014_HoangTrungHoc_2025_VN_COVID.docx` — bài VN COVID quan trọng

---

## 3. Danh sách 🟡 HAS_INLINE_REFS (50) — cần bổ sung mục Tham khảo chính quy

Đa số chỉ có **1 DOI duy nhất** ở câu mở đầu (DOI bài gốc), không có danh mục TLTK. Cần thêm mục `Tham khảo` chính quy.

### QT (42)
QT005, QT012, QT013, QT014, QT015, QT016, QT017, QT021, QT022, QT023, QT025, QT035, QT036, QT037, QT038, QT039, QT040, QT041, QT042, QT043, QT044, QT045, QT046, QT047, QT048, QT049, QT050, QT051, QT059, QT060 (orig — đã có FIXED), QT061, QT062, QT063, QT064, QT065, QT066, QT067 (orig — đã có FIXED), QT068, QT069, QT070, QT071, QT072 (orig — đã có FIXED).

### VN (8)
VN001, VN007, VN018, VN020 (orig — đã có FIXED_07062026 trong nhánh HAS_REF_SECTION), VN021, VN023, VN029, VN030.

*Chi tiết đầy đủ tên file trong `_audit_missing_refs_results.json`.*

---

## 4. ✅ HAS_REF_SECTION (11 — đã đạt chuẩn)

Tất cả là file FIXED hôm nay (07/06/2026):
- QT020, QT029, QT060, QT067, QT072, QT074, QT075, QT076, QT077, QT078 (đều `_FIXED_07062026`)
- VN020 (`_FIXED_07062026`)

→ **Pattern xác nhận:** quy trình tóm tắt cũ (Apr–Jun 2026 trước hôm nay) **không yêu cầu mục Tham khảo**. Lỗi hệ thống ở template tóm tắt gốc.

---

## 5. Khuyến nghị

### Ưu tiên 1 — Sửa quy trình (nguồn gốc lỗi)
- Cập nhật `TEMPLATE_thiet_ke_bai_test_kiem_tra_tri_thuc.md` và `PROMPT_MAU_PHAN_TICH_BAI_BAO.md` **bắt buộc** mục `Tham khảo (APA 7)` cuối cùng kèm DOI/PMID/URL kiểm chứng — đúng theo `feedback_doc_phai_co_reference.md` đã lưu trong memory.
- Thêm hook hậu kiểm: script này có thể chạy ở pre-commit để chặn tóm tắt mới thiếu Tham khảo.

### Ưu tiên 2 — Bổ sung TLTK cho 47 file 🔴
- Dùng skeleton `06_Scripts/fix_missing_refs_helper.py` (đính kèm — chỉ là template, KHÔNG chạy).
- Cho mỗi file: đối chiếu PDF gốc trong `02_Papers-goc/`, kéo TLTK trong-bài đã trích (top 5–10 cite quan trọng) + DOI bài chính → append section "Tham khảo" cuối doc.
- Verify từng DOI bằng Crossref/PubMed TRƯỚC khi ghi (theo `feedback_verify_tung_entry_truoc_khi_gui.md`).

### Ưu tiên 3 — Nâng cấp 50 file 🟡
- Đa số chỉ có 1 DOI duy nhất (của bài gốc). Cần mở rộng thành danh mục có heading rõ ràng + ≥ 5 TLTK cốt lõi.

### Ưu tiên 4 — Dedup
- 6 file gốc đã có bản FIXED hôm nay (QT020, QT029, QT060, QT067, QT072 + VN020) — đề nghị move 6 file gốc vào `_Archive/Tom-tat-orig-pre-07062026/`.

---

## 6. Tham khảo nội bộ

- Audit Agent 1 báo cáo trước: `06_Scripts/AUDIT_TomTat_FactCheck_07062026.md`, `AUDIT_TomTat_FullFactCheck_MetaAnalysis_07062026.md`
- Memory: `feedback_doc_phai_co_reference.md`, `feedback_verify_tung_entry_truoc_khi_gui.md`
- Script tự động: `06_Scripts/_audit_missing_refs_scan.py`
- Helper skeleton (chưa chạy): `06_Scripts/fix_missing_refs_helper.py`
