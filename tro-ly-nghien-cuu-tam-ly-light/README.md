# Trợ lý NC Tâm lý học — Bản LITE

Bản tối ưu cho **Render free tier (512 MB RAM)**.

Đây là phiên bản rút gọn của `tro-ly-nghien-cuu-tam-ly/`, đã thay thế phần embedding model + vector database bằng **TF-IDF keyword search pure Python** để tiết kiệm RAM.

## So sánh bản heavy vs lite

| Chỉ số | Heavy (gốc) | Lite (này) |
|---|---|---|
| Retrieval | sentence-transformers + ChromaDB | TF-IDF keyword pure Python |
| RAM tiêu thụ | ~500-700 MB | ~150-250 MB ✓ |
| Startup | 20-40 giây | 2-3 giây |
| Dependencies | 6 (gồm torch, transformers) | 3 (fastapi, uvicorn, python-multipart) |
| Chất lượng retrieval | Rất tốt (semantic) | Tốt (keyword-based) |
| Phù hợp | VPS ≥ 1 GB (Linode, Hetzner, DO) | Render free / máy yếu |

## Cấu trúc

```
tro-ly-nghien-cuu-tam-ly-light/
├── README.md
├── render.yaml              — cấu hình deploy Render (plan: free)
├── build_rag_index.py       — script build TF-IDF index (chạy khi cần re-build)
└── web/
    ├── app.py               — FastAPI backend LITE
    ├── requirements.txt     — 3 deps (không có torch/chromadb)
    ├── Dockerfile           — simple Python image
    ├── static/              — frontend HTML/CSS/JS
    ├── docs/                — 24 MB files PDF/DOCX cho download
    └── data/
        ├── glossary.json                — từ điển (213 abbrev + 47 thuật ngữ + 21 tổ chức)
        ├── rag_authors_index.json       — TF-IDF index 178 tác giả
        ├── author_kg_v1.json            — knowledge graph tác giả
        ├── authors_normalized.json      — tác giả đã chuẩn hoá
        ├── priority_authors.json        — 19 tác giả ưu tiên
        └── rag_main_index.json          — TF-IDF index 47 entries (36 papers + 11 insights) ← MỚI
```

## Các endpoint API

| Method | Path | Chức năng |
|---|---|---|
| GET | `/` | Serve index.html |
| POST | `/api/query` | Main RAG — tìm 47 entries (papers + insights) bằng TF-IDF |
| GET | `/api/documents` | List files trong /docs |
| POST | `/api/authors/query` | Tìm tác giả theo query |
| GET | `/api/authors/list` | List tác giả + filter (priority/child_focus/country) |
| GET | `/api/authors/{id}` | Detail tác giả |
| GET | `/api/authors/stats` | Stats tác giả |
| GET | `/api/glossary` | Full từ điển |
| GET | `/api/glossary/search?q=` | Tìm kiếm từ điển |
| GET | `/api/glossary/categories` | Danh mục từ điển |

## Chạy local

```bash
cd web
pip install -r requirements.txt
uvicorn app:app --port 8000
```

Mở http://localhost:8000/

## Deploy lên Render

### Cách 1 — Render tự detect render.yaml (khuyến nghị)
1. Tạo repo GitHub riêng cho folder `tro-ly-nghien-cuu-tam-ly-light/` (hoặc push folder này lên repo có sẵn, nhưng chỉ rootDir là `web/`)
2. Vào Render Dashboard → New Web Service → Connect GitHub repo
3. Render tự đọc `render.yaml` → deploy với plan free

### Cách 2 — Manual config
- Runtime: Python 3.11
- Root directory: `web`
- Build command: `pip install -r requirements.txt`
- Start command: `uvicorn app:app --host 0.0.0.0 --port $PORT`

## Re-build TF-IDF index

Khi thêm bài mới vào `../tro-ly-nghien-cuu-tam-ly/rebuild_rag_full.py`, chạy:

```bash
python build_rag_index.py
```

Script sẽ extract papers + insights từ heavy folder rồi regenerate `web/data/rag_main_index.json`.

## Khi nào nên switch sang bản heavy?

Nếu sau này:
- Thầy deploy lên VPS ≥ 1 GB RAM (Linode 2GB, Hetzner CX22, DigitalOcean Basic)
- Muốn chất lượng retrieval semantic tốt hơn (đặc biệt với query dài/phức tạp)
- Cần xử lý dataset lớn hơn (>100 entries)

→ Dùng folder `tro-ly-nghien-cuu-tam-ly/` (bản heavy) thay vì folder này.

## Ngày tạo

- 20/04/2026 — phiên bản lite v1
- Dữ liệu RAG: 47 entries extract từ heavy (36 papers + 11 insights)
- Dữ liệu tác giả: 178 authors (giữ nguyên từ heavy)
- Từ điển: 281 thuật ngữ (giữ nguyên từ heavy)
