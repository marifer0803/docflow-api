import os
import io
import json
import re
import subprocess
import tempfile
import shutil
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import docx
from docx import Document as DocxDocument
from docxtpl import DocxTemplate
import pdfplumber
from PIL import Image
import httpx

app = FastAPI(title="DocFlow API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_SERVICE_KEY", "")
STORAGE_BUCKET = os.getenv("STORAGE_BUCKET", "generated")


# ─── HEALTH ───────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok"}


# ─── EXTRACT TEXT ─────────────────────────────────────────

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = docx.Document(io.BytesIO(file_bytes))
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                parts.append(row_text)
    return "\n".join(parts)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    import fitz
    from concurrent.futures import ThreadPoolExecutor
    parts = []
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        text = page.get_text()
        if text and text.strip():
            parts.append(text.strip())
    
    if len("\n".join(parts)) > 50:
        doc.close()
        return "\n".join(parts)
    
    try:
        import pytesseract
        
        # Renderiza imagens sequencialmente (thread-safe)
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        doc.close()
        
        # OCR em paralelo (Tesseract é thread-safe)
        def ocr_image(img):
            text = pytesseract.image_to_string(img, lang="por")
            clean = text.strip()
            if not clean or len(clean) < 20:
                return None
            if clean.count('X') / max(len(clean), 1) > 0.5:
                return None
            return clean
        
        with ThreadPoolExecutor(max_workers=4) as executor:
            results = list(executor.map(ocr_image, images))
        
        ocr_parts = [r for r in results if r]
        if ocr_parts:
            return "\n".join(ocr_parts)
    except Exception:
        pass
    
    doc.close()
    return "\n".join(parts)


def extract_text_from_image(file_bytes: bytes) -> str:
    try:
        import pytesseract
        image = Image.open(io.BytesIO(file_bytes))
        text = pytesseract.image_to_string(image, lang="por")
        return text.strip()
    except Exception as e:
        return f"[OCR falhou: {str(e)}]"


@app.post("/extract-text")
async def extract_text(file: UploadFile = File(...)):
    content = await file.read()
    filename = (file.filename or "").lower()

    if filename.endswith(".docx"):
        text = extract_text_from_docx(content)
    elif filename.endswith(".pdf"):
        text = extract_text_from_pdf(content)
    elif filename.endswith((".jpg", ".jpeg", ".png", ".webp", ".tiff", ".bmp")):
        text = extract_text_from_image(content)
    else:
        try:
            text = extract_text_from_pdf(content)
        except Exception:
            try:
                text = extract_text_from_docx(content)
            except Exception:
                raise HTTPException(400, "Formato não suportado. Use DOCX, PDF ou imagem.")

    return {"text": text, "char_count": len(text)}


# ─── PLACEHOLDER FIXES ────────────────────────────────────

def merge_fragmented_placeholders(docx_path: str):
    """Reconstrói placeholders quebrados em múltiplos runs pelo Word."""
    doc = DocxDocument(docx_path)

    def fix_paragraph(para):
        full_text = "".join(run.text for run in para.runs)
        if "{{" not in full_text:
            return
        if para.runs:
            para.runs[0].text = full_text
            for run in para.runs[1:]:
                run.text = ""

    for para in doc.paragraphs:
        fix_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    fix_paragraph(para)

    for section in doc.sections:
        for part in [section.header, section.footer,
                     section.first_page_header, section.first_page_footer]:
            if part:
                for para in part.paragraphs:
                    fix_paragraph(para)

    doc.save(docx_path)


def normalize_placeholders(docx_path: str):
    """Troca {{CAMPO}} por {{ CAMPO }} para compatibilidade com Jinja2/docxtpl."""
    doc = DocxDocument(docx_path)
    pattern = re.compile(r"\{\{(\w+)\}\}")

    def fix_runs(paragraphs):
        for para in paragraphs:
            for run in para.runs:
                if "{{" in run.text:
                    run.text = pattern.sub(r"{{ \1 }}", run.text)

    fix_runs(doc.paragraphs)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                fix_runs(cell.paragraphs)

    for section in doc.sections:
        for part in [section.header, section.footer,
                     section.first_page_header, section.first_page_footer]:
            if part:
                fix_runs(part.paragraphs)

    doc.save(docx_path)


# ─── SUPABASE STORAGE ─────────────────────────────────────

async def upload_to_supabase(file_bytes: bytes, path: str, content_type: str) -> str:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise HTTPException(500, "Supabase não configurado")

    url = f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{path}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": content_type,
        "x-upsert": "true",
    }
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(url, content=file_bytes, headers=headers)
        if resp.status_code not in (200, 201):
            raise HTTPException(500, f"Upload Supabase falhou: {resp.text}")

    public_url = f"{SUPABASE_URL}/storage/v1/object/public/{STORAGE_BUCKET}/{path}"
    return public_url


# ─── GENERATE DOCUMENT ────────────────────────────────────

def docx_to_pdf(docx_path: str, output_dir: str) -> str:
    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise HTTPException(500, f"Erro ao converter para PDF: {result.stderr}")

    pdf_name = Path(docx_path).stem + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    if not os.path.exists(pdf_path):
        raise HTTPException(500, "PDF não foi gerado")
    return pdf_path


@app.post("/generate")
async def generate(
    template: UploadFile = File(...),
    data: str = Form(...),
    formats: str = Form("pdf,docx"),
    generation_id: str = Form(""),
    file_prefix: str = Form("documento"),
):
    try:
        replacements = json.loads(data)
    except json.JSONDecodeError:
        raise HTTPException(400, "JSON inválido no campo 'data'")

    requested_formats = [f.strip() for f in formats.split(",")]
    template_bytes = await template.read()
    template_filename = template.filename or "template.docx"

    tmpdir = tempfile.mkdtemp()
    try:
        template_path = os.path.join(tmpdir, template_filename)
        with open(template_path, "wb") as f:
            f.write(template_bytes)

        # Fix placeholders fragmentados pelo Word
        merge_fragmented_placeholders(template_path)
        normalize_placeholders(template_path)

        # Render com docxtpl
        tpl = DocxTemplate(template_path)
        tpl.render(replacements)
        output_docx = os.path.join(tmpdir, f"output_{template_filename}")
        tpl.save(output_docx)

        result = {"file_name": file_prefix}

        # DOCX
        if "docx" in requested_formats:
            with open(output_docx, "rb") as f:
                docx_bytes = f.read()
            if SUPABASE_URL and SUPABASE_KEY:
                path = f"{generation_id}/{file_prefix}.docx"
                result["docx_url"] = await upload_to_supabase(
                    docx_bytes, path,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                import base64
                result["docx_base64"] = base64.b64encode(docx_bytes).decode("utf-8")

        # PDF
        if "pdf" in requested_formats:
            pdf_path = docx_to_pdf(output_docx, tmpdir)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            if SUPABASE_URL and SUPABASE_KEY:
                path = f"{generation_id}/{file_prefix}.pdf"
                result["pdf_url"] = await upload_to_supabase(
                    pdf_bytes, path, "application/pdf"
                )
            else:
                import base64
                result["pdf_base64"] = base64.b64encode(pdf_bytes).decode("utf-8")

        return result

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
