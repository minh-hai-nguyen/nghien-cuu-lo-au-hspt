# Lo âu ở học sinh THCS và THPT
## Dự án nghiên cứu tổng hợp — Tháng 3/2026

---

### Cấu trúc thư mục

```
Lo-au/
├── 01_Bao-cao/          Báo cáo tổng hợp, đề cương nghiên cứu
├── 02_Papers-goc/       PDF bài báo gốc (tiếng Anh)
│   ├── Viet-Nam/
│   ├── Dong-Nam-A/
│   └── The-gioi/
├── 03_Ban-dich/         Bản dịch tiếng Việt (MD + DOCX)
├── 04_Co-so-du-lieu/    Database bài báo, gợi ý thay thế
├── 05_Cong-cu/          Prompt mẫu, script tạo báo cáo, RAG
├── Tập-viết-theo-phong-cách/   Thuật toán mô phỏng phong cách CTH
├── madam Thao/          Biên bản, transcript, câu chuyện
├── rag_db/              Vector database cho hỏi đáp RAG
└── papers/              Thư mục gốc (giữ nguyên để tương thích)
```

### Sản phẩm chính

| File | Mô tả |
|------|-------|
| Lo au - Bao cao CTH v5 + De cuong | Báo cáo 5 bài + đề cương NC theo phong cách CTH v5 |
| Lo au - THCS THPT - 27032026 | Báo cáo 21 bài (10 VN + 5 ĐNA + 6 thế giới) |
| DICH_*.docx | 5 bản dịch toàn văn (Hoa, Anderson, GBD, Zhameden, V-NAMHS) |

### Hỏi đáp RAG

```bash
cd Lo-au
python rag_lo_au.py                    # Tương tác
python rag_lo_au.py query "câu hỏi"   # Hỏi trực tiếp
python rag_lo_au.py build              # Rebuild index
```

971 chunks từ 32 tài liệu. Hỗ trợ tiếng Việt và tiếng Anh.
