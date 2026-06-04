FROM python:3.11-slim

WORKDIR /app

# Install dependencies (only 3 light deps — no torch, no chromadb)
COPY requirements.txt .
RUN pip install --no-cache-dir --root-user-action=ignore -r requirements.txt

# Copy app + data
COPY . .

EXPOSE 10000
CMD uvicorn app:app --host 0.0.0.0 --port ${PORT:-10000}
