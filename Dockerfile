# Production Dockerfile for Khanna Travels Visa Document Automation Backend
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies for PDF / image processing & OCR
RUN apt-get update && apt-get install -y --no-install-recommends     tesseract-ocr     tesseract-ocr-eng     libgl1     libglib2.0-0     && rm -rf /var/lib/apt/lists/*

# Copy backend requirements and install
COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install --no-cache-dir -r backend/requirements.txt

# Copy backend source, templates, and reference template assets
COPY backend/ ./backend/
COPY templates/ ./templates/
COPY Europe_covering_letter_template_clean.* ./
COPY Japan_covering_letter_template_clean.* ./
COPY Singapore_covering_letter_template_clean.* ./
COPY khanna\ travels\ logo.png ./

# Create data, uploads, and temp storage directories
RUN mkdir -p data temp uploads internal_records

ENV HOST=0.0.0.0
ENV PORT=8000
ENV PYTHONUNBUFFERED=1

EXPOSE 8000

CMD ["python", "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
