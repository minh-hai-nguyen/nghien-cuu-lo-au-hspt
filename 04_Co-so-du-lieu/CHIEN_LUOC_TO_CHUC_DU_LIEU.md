# CHIẾN LƯỢC TỔ CHỨC DỮ LIỆU DỰ ÁN LO ÂU

**Cập nhật: 05/04/2026**
**Mục đích:** Quản lý dự án NC dài hạn với dữ liệu thêm/sửa/cập nhật liên tục qua nhiều phiên làm việc.

---

## I. PHÂN TÍCH HIỆN TRẠNG

### Quy mô dự án (05/04/2026)
- **35 bài NC** (7 VN + 28 QT), 572 MB tổng dung lượng
- **184 file DOCX**, **123 file PDF**, **112 file Python**, **56 file MD**, **210 file PNG**
- **6 thư mục chính**, nhiều thư mục con
- **Phiên làm việc:** kéo dài ≥ 12 ngày, nhiều phiên/ngày

### Vấn đề đã phát hiện

| Vấn đề | Ví dụ cụ thể | Tần suất |
|--------|--------------|----------|
| **File nhiều phiên bản theo ngày** | `Lo au - Bao cao + De cuong - 27032026.docx` + `05042026.docx` | Cao |
| **Script 1 lần dùng xếp chồng** | 42 file .py trong 03_Ban-dich, 31 trong Tom-tat-tung-bai | Rất cao |
| **PDF gốc không đầy đủ/lỗi** | NgoAnhVinh 2p metadata → 7p full, BaoQuyen 1p → 8p full | Trung bình |
| **Sửa lỗi lan tỏa nhiều file** | Wen 11,6 → 11,58 cần sửa ở 2+ file | Cao |
| **Trùng tên file** | 2 bản `2617-Văn bản của bài báo...` gần giống | Thấp |
| **Thư mục phân mảnh** | The-gioi, The-gioi-moi, Chua-phan-loai, Dong-Nam-A | Cao |
| **Memory vs Project State lệch** | Memory ghi "7 bài VN" nhưng thực tế có 10 | Cao |
| **Python scripts rải rác** | tao_Hoa2024.py, dich_QT32_full.py, fix_all_files.py | Cao |

---

## II. NGUYÊN TẮC TỔ CHỨC DỮ LIỆU

### Nguyên tắc 1: SINGLE SOURCE OF TRUTH (SSOT)
> Mỗi thông tin CHỈ có 1 nơi "chính thức" (authoritative source).

Ví dụ:
- **Số liệu bài báo** → PDF gốc (không phải bản dịch, không phải cross-study)
- **Tiến trình dự án** → `memory/project_lo_au_progress.md` (không phải memo text)
- **Phát hiện lỗi** → `04_Co-so-du-lieu/BANG_LOI_DA_SUA.md` (không phải notes phân tán)
- **Đề cương mới nhất** → file có timestamp + symlink `DECUONG_LATEST.docx`

### Nguyên tắc 2: APPEND-ONLY + SNAPSHOT
> KHÔNG xóa file cũ. Thêm phiên bản mới với timestamp. Chỉ xóa file tạm (`*.tmp`, `*_raw.txt`).

Cách làm:
- File chính: `Lo au - De cuong - YYYYMMDD.docx`
- Thư mục snapshot: `00_Archive/2026-Q1/` (lưu các bản cũ > 1 tuần)
- Latest pointer: `DECUONG_LATEST.md` (chứa tên file mới nhất)

### Nguyên tắc 3: IDEMPOTENT SCRIPTS
> Scripts phải chạy lại được nhiều lần mà không phá dữ liệu (idempotent).

- TRÁNH: `tao_Hoa2024.py` (chỉ chạy 1 lần → tạo file)
- NÊN: `rebuild_dich.py --paper VN01 --force` (chạy lại được)

### Nguyên tắc 4: DERIVED vs SOURCE
> Phân biệt rõ dữ liệu nguồn vs dữ liệu phái sinh.

| Loại | Ví dụ | Đặc tính |
|------|-------|----------|
| **SOURCE** (không thể tái tạo) | PDF gốc, file bạn tạo thủ công | Backup kỹ |
| **DERIVED** (tái tạo được) | Bản dịch DOCX, tóm tắt, RAG, cross-study | Có thể xóa và tạo lại |
| **TRANSIENT** (tạm) | `*_raw.txt`, `Charts/*.png`, `__pycache__` | Xóa an toàn |

### Nguyên tắc 5: NAMING CONVENTION
> Tên file tự giải thích + sortable.

Format chuẩn:
```
[Loai][SoId]_[TenNgan]_[Nam]_[Meta].[ext]
```

Ví dụ:
- `VN15_NgoAnhVinh_2024_JAffDisordRep.docx` ✓
- `QT21_Norway_2025_SocSciMed.docx` ✓
- `ban-dich-moi-nhat.docx` ✗ (không có ID, không có năm)

---

## III. CẤU TRÚC ĐỀ XUẤT

```
Lo-au/
│
├── 00_Meta/                          # Metadata, tracking, index
│   ├── PROJECT_STATE.md              # SSOT cho trạng thái dự án
│   ├── PAPERS_INDEX.md               # Index 35 bài (ID, tác giả, trạng thái)
│   ├── CHANGELOG.md                  # Nhật ký thay đổi theo phiên
│   ├── ERRORS_FIXED.md               # Bảng lỗi + sửa
│   └── NAMING_CONVENTIONS.md         # Quy tắc đặt tên
│
├── 01_Papers-goc/                    # SOURCE — PDF bài gốc
│   ├── VN/                           # 10 bài VN
│   │   └── VN15_NgoAnhVinh_2024.pdf
│   ├── QT/                           # 28 bài QT
│   │   └── QT21_Norway_2025.pdf
│   └── _Archive/                     # PDF cũ/lỗi (giữ để so sánh)
│       └── VN15_NgoAnhVinh_2024_METADATA_ONLY_2p.pdf
│
├── 02_Ban-dich/                      # DERIVED — bản dịch DOCX
│   ├── VN15_NgoAnhVinh_2024.docx
│   ├── QT21_Norway_2025.docx
│   └── _Charts/                      # Ảnh crop từ PDF
│       └── QT21_Fig1_trends.png
│
├── 03_Tom-tat/                       # DERIVED — tóm tắt CTH v5
│   ├── VN15_NgoAnhVinh_2024.docx
│   └── QT21_Norway_2025.docx
│
├── 04_Tong-hop/                      # DERIVED — bảng tổng hợp, cross-study
│   ├── Cross_Study_LATEST.docx       # Symlink → 04042026
│   ├── Cross_Study_2026-04-04.docx
│   ├── Tom_Tat_7VN_LATEST.docx
│   └── _Archive/
│
├── 05_De-cuong/                      # SOURCE + DERIVED
│   ├── DECUONG_LATEST.docx           # Symlink → 05042026
│   ├── DECUONG_2026-04-05.docx
│   └── _Archive/
│       └── DECUONG_2026-03-27.docx
│
├── 06_Scripts/                       # Tools tái tạo derived data
│   ├── lib/
│   │   ├── docx_helpers.py           # = tao_dich_template.py
│   │   └── pdf_helpers.py
│   ├── dich/                         # Scripts dịch
│   │   ├── dich_paper.py             # SCRIPT CHUNG (idempotent)
│   │   └── paper_configs.json        # Config từng bài
│   ├── tom_tat/
│   ├── kiem_tra/
│   │   ├── check_errors.py           # Kiểm tra lỗi toàn diện
│   │   └── verify_against_pdf.py     # So chéo với PDF gốc
│   └── update_all.py                 # Master script: build lại tất cả
│
├── 07_RAG/                           # DERIVED
│   ├── rag_db/                       # ChromaDB
│   └── rebuild_rag.py
│
└── _Sandbox/                         # Playground, file tạm
    ├── tmp/
    └── experiments/
```

---

## IV. WORKFLOW ĐỀ XUẤT

### Khi thêm bài mới
1. Đặt PDF vào `01_Papers-goc/VN/` hoặc `QT/` theo naming convention
2. Cập nhật `00_Meta/PAPERS_INDEX.md` (thêm 1 dòng)
3. Chạy `06_Scripts/dich/dich_paper.py --id VN21 --force`
4. Chạy `06_Scripts/kiem_tra/verify_against_pdf.py --id VN21`
5. Chạy `06_Scripts/update_all.py` (rebuild cross-study + RAG)
6. Cập nhật `00_Meta/CHANGELOG.md`

### Khi sửa lỗi phát hiện
1. Ghi lỗi vào `00_Meta/ERRORS_FIXED.md` NGAY (trước khi sửa)
2. Xác định phạm vi: chỉ 1 bài hay lan tỏa?
3. Nếu lan tỏa → sửa trong script chung, chạy lại tất cả
4. Chạy `verify_against_pdf.py` cho các bài liên quan
5. Commit vào CHANGELOG

### Khi bắt đầu phiên mới
1. Đọc `00_Meta/PROJECT_STATE.md` (1 file duy nhất)
2. Đọc `00_Meta/CHANGELOG.md` 3 ngày gần nhất
3. Check `00_Meta/ERRORS_FIXED.md` lỗi chưa giải quyết

### Khi kết thúc phiên
1. Cập nhật `PROJECT_STATE.md`
2. Thêm mục vào `CHANGELOG.md` với timestamp
3. Di chuyển file cũ > 1 tuần vào `_Archive/`

---

## V. HÀNH ĐỘNG CẦN LÀM NGAY

### Ưu tiên CAO (làm trước)

1. **Tạo `00_Meta/PROJECT_STATE.md`** — 1 file SSOT cho trạng thái
2. **Tạo `00_Meta/PAPERS_INDEX.md`** — bảng 35 bài với mọi thông tin
3. **Tạo `00_Meta/CHANGELOG.md`** — ghi lại lịch sử phiên
4. **Tạo `00_Meta/ERRORS_FIXED.md`** — bảng lỗi (đã phát hiện Wen 11,6→11,58, CAMS attribution)
5. **Gộp 42 script dịch → 1 script idempotent** (`dich_paper.py` + `paper_configs.json`)

### Ưu tiên TRUNG BÌNH

6. Archive file cũ: di chuyển file > 7 ngày vào `_Archive/`
7. Sắp xếp lại `02_Papers-goc/` (gộp VN/QT, bỏ thư mục phân mảnh)
8. Tạo symlink `DECUONG_LATEST.docx`, `CROSS_STUDY_LATEST.docx`
9. Viết `verify_against_pdf.py` chạy tự động kiểm tra tất cả

### Ưu tiên THẤP (có thể làm dần)

10. Dọn file `*_raw.txt` tạm
11. Dọn `Charts/` giữ lại các ảnh đang dùng
12. Compress `_Archive/` cũ

---

## VI. NGUYÊN TẮC CHO CLAUDE (tôi) TUÂN THỦ

Khi làm việc với dự án này, tôi phải:

1. **LUÔN** đọc `PROJECT_STATE.md` đầu phiên trước khi làm gì
2. **KHÔNG** tạo file mới có timestamp mà không cập nhật CHANGELOG
3. **LUÔN** verify số liệu với PDF gốc trước khi đưa vào derived files
4. **KHÔNG** viết script 1-lần (tao_XXX.py) — phải là script idempotent
5. **KHÔNG** để 2 nơi cùng lưu 1 thông tin (tránh drift)
6. **LUÔN** ghi lỗi phát hiện vào ERRORS_FIXED.md trước khi sửa
7. **LUÔN** verify chéo sau khi sửa lỗi (script + manual check)
8. **LUÔN** cập nhật PROJECT_STATE.md cuối phiên

---

## VII. LỢI ÍCH KỲ VỌNG

| Trước | Sau |
|-------|-----|
| Đọc memory + scan nhiều file để biết trạng thái | 1 file `PROJECT_STATE.md` |
| 42 script `tao_*.py` rải rác | 1 script `dich_paper.py` + config |
| Sửa lỗi phải tìm ở nhiều file | 1 script `update_all.py` rebuild |
| Không rõ phiên bản nào mới nhất | `*_LATEST.docx` symlink |
| Không có nhật ký thay đổi | CHANGELOG.md timestamped |
| Lỗi đã sửa có thể tái xuất hiện | ERRORS_FIXED.md + regression test |
| Memory có thể lệch với thực tế | Verify script chạy tự động |
