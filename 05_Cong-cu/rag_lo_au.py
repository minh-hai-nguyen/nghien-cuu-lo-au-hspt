# -*- coding: utf-8 -*-
"""
RAG Hỏi đáp về Lo âu ở học sinh THCS/THPT
Dựa trên các bài báo nghiên cứu đã có trong thư mục papers/
"""
import sys, io, os, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import chromadb
from sentence_transformers import SentenceTransformer
import PyPDF2

# === CONFIG ===
PAPERS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'papers')
DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'rag_db')
COLLECTION_NAME = 'lo_au_hoc_sinh_v2'

def extract_text_from_pdf(pdf_path, max_pages=50):
    """Extract text from PDF using PyPDF2"""
    try:
        reader = PyPDF2.PdfReader(pdf_path)
        text = ''
        for i, page in enumerate(reader.pages[:max_pages]):
            text += (page.extract_text() or '') + '\n'
        return text
    except Exception as e:
        print(f'  ERROR reading {pdf_path}: {e}')
        return ''

def extract_text_from_md(md_path):
    """Extract text from markdown file"""
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f'  ERROR reading {md_path}: {e}')
        return ''

def chunk_text(text, chunk_size=500, overlap=100):
    """Split text into overlapping chunks"""
    words = text.split()
    chunks = []
    for i in range(0, len(words), chunk_size - overlap):
        chunk = ' '.join(words[i:i + chunk_size])
        if len(chunk.strip()) > 50:  # skip very short chunks
            chunks.append(chunk)
    return chunks

def build_index():
    """Build ChromaDB index from all papers"""
    print('=== Building RAG Index ===')
    print(f'Papers dir: {PAPERS_DIR}')

    # Initialize
    model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
    client = chromadb.PersistentClient(path=DB_DIR)

    # Delete existing collection if exists
    try:
        client.delete_collection(COLLECTION_NAME)
    except:
        pass

    collection = client.create_collection(
        name=COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"}
    )

    # Find all documents
    pdf_files = glob.glob(os.path.join(PAPERS_DIR, '*.pdf'))
    md_files = glob.glob(os.path.join(PAPERS_DIR, 'DICH_*.md'))
    # Also include Vietnamese reports
    md_files += glob.glob(os.path.join(PAPERS_DIR, '..', 'BAO_CAO_*.md'))
    md_files += glob.glob(os.path.join(PAPERS_DIR, '..', 'DATABASE_*.md'))

    all_chunks = []
    all_ids = []
    all_metadata = []

    doc_id = 0

    # Process PDFs
    for pdf_path in pdf_files:
        fname = os.path.basename(pdf_path)
        print(f'  Processing PDF: {fname}')
        text = extract_text_from_pdf(pdf_path)
        if not text:
            continue
        chunks = chunk_text(text)
        for i, chunk in enumerate(chunks):
            all_chunks.append(chunk)
            all_ids.append(f'doc{doc_id}_chunk{i}')
            all_metadata.append({
                'source': fname,
                'type': 'pdf',
                'chunk_index': i
            })
        doc_id += 1
        print(f'    -> {len(chunks)} chunks')

    # Process MDs (translations + reports)
    for md_path in md_files:
        fname = os.path.basename(md_path)
        print(f'  Processing MD: {fname}')
        text = extract_text_from_md(md_path)
        if not text:
            continue
        chunks = chunk_text(text)
        for i, chunk in enumerate(chunks):
            all_chunks.append(chunk)
            all_ids.append(f'doc{doc_id}_chunk{i}')
            all_metadata.append({
                'source': fname,
                'type': 'markdown',
                'chunk_index': i
            })
        doc_id += 1
        print(f'    -> {len(chunks)} chunks')

    # Embed and store
    print(f'\nTotal chunks: {len(all_chunks)}')
    print('Embedding chunks...')

    # Process in batches of 100
    batch_size = 100
    for i in range(0, len(all_chunks), batch_size):
        batch_chunks = all_chunks[i:i+batch_size]
        batch_ids = all_ids[i:i+batch_size]
        batch_meta = all_metadata[i:i+batch_size]

        embeddings = model.encode(batch_chunks).tolist()
        collection.add(
            documents=batch_chunks,
            embeddings=embeddings,
            ids=batch_ids,
            metadatas=batch_meta
        )
        print(f'  Batch {i//batch_size + 1}/{(len(all_chunks)-1)//batch_size + 1} done')

    print(f'\nIndex built: {len(all_chunks)} chunks from {doc_id} documents')
    print(f'Saved to: {DB_DIR}')
    return collection, model

def query_rag(question, collection, model, n_results=5):
    """Query the RAG system"""
    q_embedding = model.encode([question]).tolist()
    results = collection.query(
        query_embeddings=q_embedding,
        n_results=n_results,
        include=['documents', 'metadatas', 'distances']
    )
    return results

def format_answer(question, results):
    """Format RAG results into a readable answer"""
    output = f'\n{"="*60}\n'
    output += f'CÂU HỎI: {question}\n'
    output += f'{"="*60}\n\n'

    for i, (doc, meta, dist) in enumerate(zip(
        results['documents'][0],
        results['metadatas'][0],
        results['distances'][0]
    )):
        relevance = max(0, (1 - dist) * 100)
        output += f'--- Nguồn {i+1} (Độ liên quan: {relevance:.0f}%) ---\n'
        output += f'Tài liệu: {meta["source"]}\n'
        output += f'{doc[:500]}...\n\n'

    return output

def interactive_mode():
    """Run interactive Q&A"""
    print('\n' + '='*60)
    print('  RAG HỎI ĐÁP VỀ LO ÂU Ở HỌC SINH THCS/THPT')
    print('  Gõ câu hỏi bằng tiếng Việt hoặc tiếng Anh')
    print('  Gõ "quit" để thoát')
    print('='*60)

    # Load or build index
    client = chromadb.PersistentClient(path=DB_DIR)
    model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

    try:
        collection = client.get_collection(COLLECTION_NAME)
        count = collection.count()
        print(f'\nĐã tải index: {count} chunks')
    except:
        print('\nChưa có index. Đang xây dựng...')
        collection, model = build_index()

    while True:
        try:
            question = input('\n> Câu hỏi: ').strip()
        except (EOFError, KeyboardInterrupt):
            break

        if not question or question.lower() in ['quit', 'exit', 'q']:
            break

        results = query_rag(question, collection, model)
        print(format_answer(question, results))

if __name__ == '__main__':
    if len(sys.argv) > 1:
        if sys.argv[1] == 'build':
            build_index()
        elif sys.argv[1] == 'query':
            question = ' '.join(sys.argv[2:])
            client = chromadb.PersistentClient(path=DB_DIR)
            model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
            collection = client.get_collection(COLLECTION_NAME)
            results = query_rag(question, collection, model)
            print(format_answer(question, results))
        else:
            interactive_mode()
    else:
        interactive_mode()
