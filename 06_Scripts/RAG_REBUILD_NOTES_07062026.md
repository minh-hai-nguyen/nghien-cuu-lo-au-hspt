# RAG Rebuild Notes — 07/06/2026

## Mục đích
Ghi lại các tài liệu mới đã được thêm vào kho `02_Papers-goc/` để chạy lại
RAG embedding khi cần.

## Tài liệu cần index lại trong RAG

| File mới thêm | Đường dẫn | Mục đích |
|---|---|---|
| `karasu1986.pdf` | `02_Papers-goc/karasu1986.pdf` | Karasu 1986 — The psychotherapies: benefits and limitations. AJ Psychotherapy 40(3):324-342. **Verify "over 450 types" tại trang 325** |
| `biosketch_karasu_062112.pdf` | `02_Papers-goc/biosketch_karasu_062112.pdf` | Einstein bio sketch Karasu (verify credentials cho buổi bảo vệ LA) |

## Tài liệu cần download bổ sung để index

| Cần download | URL gợi ý | Mục đích |
|---|---|---|
| Rose 2002 PDF | https://pubmed.ncbi.nlm.nih.gov/12487497/ (paywall) | Co-rumination hypothesis (Section 4.4 Q1) |
| Stankov 2010 PDF | https://eric.ed.gov/?id=EJ905763 (ERIC) hoặc ScienceDirect | Confucian academic culture (Section 1.1 Q1) |
| Small & Blanc 2021 PDF | https://pmc.ncbi.nlm.nih.gov/articles/PMC7820702/ (PMC open access) | Tam giao Vietnam (Section 4.4 Q1) |

## Lệnh rebuild RAG

```bash
cd tro-ly-nghien-cuu-tam-ly
python build_rag.py        # rebuild collection 'lo_au_v3'
python build_full_rag.py   # rebuild collection 'lo_au_full'
```

Backup tự động tạo: `rag_db_backup/`

## Lưu ý
- RAG index hiện tại chưa bao gồm các paper mới (Karasu, Rose, Stankov, Small & Blanc).
- Sau khi rebuild, kiểm tra retrieve query "over 450 types Karasu" → phải trả về chunk từ
  trang 325 của `karasu1986.pdf`.
- Embedding model hiện tại: `all-MiniLM-L6-v2` (384-dim, multilingual hạn chế cho tiếng Việt).
- Cân nhắc chuyển sang `paraphrase-multilingual-MiniLM-L12-v2` (384-dim) hoặc
  `intfloat/multilingual-e5-base` (768-dim) nếu cần hỗ trợ tiếng Việt tốt hơn.

## Trạng thái CSDL + KG sau cập nhật 07/06/2026

| Hệ thống | Trạng thái |
|---|---|
| **CSDL** `04_Co-so-du-lieu/DATABASE_BAI_BAO_LO_AU.md` | ✓ Đã cập nhật mục 32 (Karasu + Rose + Stankov + Small & Blanc + DSM/ICD) |
| **KG** `06_Scripts/author_kg_data/author_kg_v2.json` | ✓ 260 nodes, 824 edges (thêm 19 nodes mới + 15 edges mới) |
| **RAG** `rag_db/` | ⏳ Chưa rebuild với PDF mới — cần chạy lại `build_rag.py` |
| **Đối chiếu Q1 v5** | ✓ 3 nguồn mới đã verify + cite đúng trong Section 1.1 và 4.4 |
