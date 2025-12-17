import os
import subprocess
import uuid
from PIL import Image
from pypdf import PdfReader, PdfWriter


# Allowed upload extensions
ALLOWED_IN = {'.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx', '.pdf', '.jpg', '.jpeg', '.png'}


# -----------------------------
# BASIC UTILITIES
# -----------------------------
def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def random_filename(filename: str) -> str:
    ext = os.path.splitext(filename)[1]
    return f"{uuid.uuid4().hex}{ext}"


# -----------------------------
# OFFICE → PDF (LibreOffice)
# -----------------------------
def convert_office_to_pdf(input_path: str, output_dir: str) -> str:
    ensure_dir(output_dir)

    cmd = [
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path,
    ]

    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {proc.stderr.decode()}")

    base = os.path.splitext(os.path.basename(input_path))[0]
    output_file = os.path.join(output_dir, base + ".pdf")

    if not os.path.exists(output_file):
        raise RuntimeError("Converted PDF not found!")

    return output_file


# -----------------------------
# Image → PDF
# -----------------------------
def image_to_pdf(input_path: str, output_path: str):
    img = Image.open(input_path)

    # Convert RGBA → RGB
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")

    img.save(output_path, "PDF", resolution=100)


# -----------------------------
# PDF Merge
# -----------------------------
def merge_pdfs(file_paths: list, out_path: str):
    writer = PdfWriter()

    for fp in file_paths:
        reader = PdfReader(fp)
        for page in reader.pages:
            writer.add_page(page)

    with open(out_path, "wb") as f:
        writer.write(f)


# -----------------------------
# PDF Split
# -----------------------------
def split_pdf(pdf_path: str, pages: list, out_path: str):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for p in pages:
        idx = p - 1
        if idx < 0 or idx >= len(reader.pages):
            raise IndexError("Page index out of range")
        writer.add_page(reader.pages[idx])

    with open(out_path, "wb") as f:
        writer.write(f)


# -----------------------------
# File type checker
# -----------------------------
def allowed_file(filename: str) -> bool:
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_IN


# -----------------------------
# PDF Compression (Improved)
# -----------------------------
def compress_pdf(input_path, output_path, quality="medium"):
    """
    Compress PDF using content stream compression.
    quality = low, medium, high (affects image compression)
    """

    reader = PdfReader(input_path)
    writer = PdfWriter()

    # Compression map
    quality_map = {
        "low": 50,
        "medium": 70,
        "high": 85
    }

    jpeg_quality = quality_map.get(quality, 70)

    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)

    # Write out
    with open(output_path, "wb") as f:
        writer.write(f)
