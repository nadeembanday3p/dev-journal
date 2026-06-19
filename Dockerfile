FROM python:3.11-slim

# Install Tesseract OCR + OpenCV system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (cached layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

# Copy app source — explicit folders so nothing is accidentally skipped
COPY app.py .
COPY templates/ ./templates/
COPY static/ ./static/

# Persistent data lives in /data (mount a volume here on Railway/Render)
ENV DATA_DIR=/data
RUN mkdir -p /data/uploads

EXPOSE 8080

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--timeout", "120", "app:app"]
