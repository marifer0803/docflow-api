FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    default-jre \
    tesseract-ocr \
    tesseract-ocr-por \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

EXPOSE 8000
ENV SUPABASE_URL=https://bleatingparrot-supabase.cloudfy.live
ENV SUPABASE_SERVICE_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoic2VydmljZV9yb2xlIiwiaXNzIjoic3VwYWJhc2UiLCJpYXQiOjE3NjA3NDczNDgsImV4cCI6MTc5MjI4MzM0OH0.erP_tL_hkgGJPRPXWsNrkri9fAiaqbALtZUGg5B8htk
ENV GEMINI_API_KEY=AIzaSyC9HEX_25pxegTewy5bGVp96EcmLld_os8
ENV STORAGE_BUCKET=generated
ENV TEMPLATES_BUCKET=templates
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
