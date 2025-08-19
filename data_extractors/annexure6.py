"""
Invoice Data Extractor (Mixed PDFs: text + scanned) - Refined v6
---------------------------------------------------------------
- Improved name extraction (fallback: line before GSTIN)
- Expanded product description capture (multi-line, cleans trailing numbers)
"""

import os
import re
import sys
import argparse
from typing import List, Dict, Optional, Tuple

import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageFilter
import pandas as pd

POPPLER_PATH = r"C:\\tools\\poppler-24.08.0\\Library\\bin"
TESSERACT_CMD = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

TESSERACT_CONFIG = "--oem 3 --psm 6"
DPI = 300

GSTIN_RE = re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b")
HSN_LABELLED_RE = re.compile(r"(?:HSN|SAC|HSN/SAC|Service Accounting Code|SAC Code|Category of Service)\s*[:#-]?\s*(\d{4,8})", re.IGNORECASE)
HSN_ANY6_RE = re.compile(r"\b\d{6}\b")

INVOICE_NO_PATTERNS = [
    r"Invoice\s*(?:No|#|Number)\s*[:#\.-]*\s*([A-Z0-9\-/]+)",
    r"Document\s*No\s*[:#\.-]*\s*([A-Z0-9\-/]+)"
]
INVOICE_DATE_PATTERNS = [
    r"Invoice\s*Date\s*[:#\.-]*\s*([0-9]{1,2}[\-/][0-9]{1,2}[\-/][0-9]{2,4})",
    r"Created\s*[:#\.-]*\s*([0-9]{1,2}[\-/][0-9]{1,2}[\-/][0-9]{2,4})",
    r"Document\s*Date\s*[:#\.-]*\s*([0-9]{1,2}[\-/][0-9]{1,2}[\-/][0-9]{2,4})",
]

BLOCK_END_TOKENS = [
    r"GSTIN", r"From", r"Seller", r"Supplier", r"Place of Supply", r"Ship To",
    r"PAN", r"CIN", r"Tax Invoice", r"TAX INVOICE", r"Invoice", r"HSN", r"SAC",
]
BLOCK_END_PATTERN = re.compile(r"|".join([f"(?:{k})" for k in BLOCK_END_TOKENS]), re.IGNORECASE)


def add_paths_to_env():
    if POPPLER_PATH and os.path.isdir(POPPLER_PATH):
        os.environ["PATH"] += os.pathsep + POPPLER_PATH
    if TESSERACT_CMD and os.path.exists(TESSERACT_CMD):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


def read_pdf_text(pdf_path: str) -> Tuple[str, List[str]]:
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            texts.append(t)
    return ("\n".join(texts), texts)


def preprocess_image(img: Image.Image) -> Image.Image:
    scale = 1.5
    new_size = (int(img.width * scale), int(img.height * scale))
    img = img.resize(new_size, Image.LANCZOS)
    img = img.convert("L")
    img = img.filter(ImageFilter.SHARPEN)
    img = img.point(lambda x: 0 if x < 170 else 255, '1')
    return img


def ocr_pdf(pdf_path: str) -> str:
    images = convert_from_path(pdf_path, dpi=DPI, poppler_path=POPPLER_PATH if POPPLER_PATH else None)
    text_chunks = []
    for img in images:
        proc = preprocess_image(img)
        txt = pytesseract.image_to_string(proc, config=TESSERACT_CONFIG)
        text_chunks.append(txt)
    return "\n".join(text_chunks)


def is_text_pdf(page_texts: List[str]) -> bool:
    total_len = sum(len(t) for t in page_texts)
    return total_len >= 400 or any(len(t) >= 200 for t in page_texts)


def safe_search(pattern: str, text: str, flags=re.IGNORECASE | re.DOTALL) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else None


def find_first_match(patterns: List[str], text: str) -> Optional[str]:
    for pat in patterns:
        v = safe_search(pat, text)
        if v:
            return v
    return None


def extract_block(text: str, start_labels: List[str]) -> Optional[str]:
    for label in start_labels:
        m = re.search(label + r"\s*[:,-]*\s*(.*?)(?=\n(?:" + BLOCK_END_PATTERN.pattern + r")|$)", text, re.IGNORECASE | re.DOTALL)
        if m:
            block = m.group(1).strip()
            block = re.sub(r"\n{2,}", "\n", block)
            block = re.sub(r"[\t ]{2,}", " ", block)
            return block
    return None


def extract_name_only(block: str, norm: str) -> Optional[str]:
    if not block:
        return None
    lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
    for ln in lines:
        if re.match(r'^(From|To|Bill|Ship|Party|Customer|Buyer|Consignee|Invoice|Billed)', ln, re.IGNORECASE):
            continue
        if re.match(r'^(GSTIN|Address|CIN|PAN|Phone|Email)', ln, re.IGNORECASE):
            continue
        return ln
    # fallback: line before GSTIN
    gstin_match = GSTIN_RE.search(norm)
    if gstin_match:
        before = norm[:gstin_match.start()].splitlines()
        if before:
            return before[-1].strip()
    return None


def extract_hsn_codes(text: str) -> Optional[str]:
    label_hits = HSN_LABELLED_RE.findall(text)
    six_hits = HSN_ANY6_RE.findall(text)
    all_codes = []
    for c in label_hits + six_hits:
        if c and c not in all_codes and len(c) in (6, 7, 8):
            all_codes.append(c)
    return ", ".join(all_codes) if all_codes else None


def clean_description(desc: str) -> Optional[str]:
    if not desc:
        return None
    return re.split(r"\s+\d", desc.strip())[0].strip()


def parse_invoice_text(text: str) -> Dict[str, Optional[str]]:
    data: Dict[str, Optional[str]] = {
        "Invoice No": None,
        "Invoice Date": None,
        "Invoice Issued To": None,
        "GSTIN (Issued To)": None,
        "Name of Who Generated Invoice": None,
        "GSTIN of That Name": None,
        "Description": None,
        "HSN Code": None,
        "Value (Net of GST)": None,
        "Quantity": None,
        "Value Per Piece": None,
        "Detected Template": None,
    }

    norm = re.sub(r"\u00A0", " ", text)

    if re.search(r"Razorpay", norm, re.IGNORECASE):
        data["Detected Template"] = "Razorpay"
    elif re.search(r"Saraswati\s+and\s+Sons", norm, re.IGNORECASE):
        data["Detected Template"] = "Saraswati & Sons"
    else:
        data["Detected Template"] = "Generic"

    data["Invoice No"] = find_first_match(INVOICE_NO_PATTERNS, norm)
    data["Invoice Date"] = find_first_match(INVOICE_DATE_PATTERNS, norm)

    # --- Invoice Issued To ---
    issued_block = extract_block(
        norm,
        [r"Issued\s*To", r"Bill(?:ed)?\s*To", r"Invoiced\s*To", r"To\s*,?\s*Party\s*Name",
         r"To\s*,?", r"Buyer", r"Customer", r"Consignee", r"Invoice\s*For", r"Party"]
    )
    data["Invoice Issued To"] = extract_name_only(issued_block, norm)

    issued_scope = safe_search(r"(Issued\s*To.*?)(?:From|Seller|Supplier|Tax\s*Invoice|$)", norm)
    if issued_scope:
        issued_gstin = GSTIN_RE.search(issued_scope)
        if issued_gstin:
            data["GSTIN (Issued To)"] = issued_gstin.group(0)

    # --- Seller / From ---
    from_block = extract_block(
        norm,
        [r"From", r"Seller", r"Supplier", r"Proprietor", r"Billed\s*By", r"Issued\s*By", r"Generated\s*By", r"Company Name"]
    )
    data["Name of Who Generated Invoice"] = extract_name_only(from_block, norm)

    from_scope = safe_search(r"(From.*?)(?:Issued\s*To|Bill\s*To|To\s*,|Tax\s*Invoice|$)", norm)
    if from_scope:
        m2 = GSTIN_RE.search(from_scope)
        if m2:
            data["GSTIN of That Name"] = m2.group(0)

    if data["GSTIN of That Name"] == data["GSTIN (Issued To)"]:
        all_g = GSTIN_RE.findall(norm)
        if len(all_g) >= 2:
            data["GSTIN (Issued To)"] = all_g[0]
            data["GSTIN of That Name"] = all_g[1]

    # --- Description ---
    desc_match = re.search(
        r"(?:Particulars|Description|Item\s*Description|Name\s*of\s*Product\s*/\s*Service|Description\s*of\s*Service)\s*[:\n]+([A-Za-z].{3,200})",
        norm, re.IGNORECASE)
    if desc_match:
        raw_desc = desc_match.group(1)
        lines = []
        for ln in raw_desc.splitlines():
            if re.search(r"\b(Qty|Quantity|Rate|Amount|Total|HSN|SAC)\b", ln, re.IGNORECASE):
                break
            lines.append(ln.strip())
        desc = " ".join(lines).strip()
    else:
        desc = safe_search(r"\n\s*\d+\s+([A-Za-z].{3,100})", norm)

    data["Description"] = clean_description(desc)

    # --- HSN / SAC ---
    hsn_label = safe_search(r"(?:HSN|SAC|HSN/SAC|Service Accounting Code|SAC Code|Category of Service)[:\s]+(\d{4,8})", norm)
    if not hsn_label:
        hsn_label = extract_hsn_codes(norm)
    data["HSN Code"] = hsn_label

    # --- Quantity ---
    qty = safe_search(r"(?:Quantity|No\. of (?:days|hours)|Units|Nos\.?|Qty)[:\s\-]*([0-9]+(?:\.[0-9]+)?)", norm)
    data["Quantity"] = qty

    # --- Rate ---
    rate = safe_search(r"(?:Rate|Unit Price|Price|Value\s*per\s*piece)[:\s₹]*([0-9][0-9,\.]+)", norm)
    data["Value Per Piece"] = rate

    # --- Value (Net of GST / Total) ---
    value = safe_search(r"Gross\s*Amount\s*Payable.*?([0-9][0-9,\.]+)", norm)
    if not value:
        value = safe_search(r"(?:Grand\s*Total|Total|Net Value|Gross Value|Amount)[:\s₹]*([0-9][0-9,\.]+)", norm)
    data["Value (Net of GST)"] = value

    return data


def process_pdfs(input_dir: str) -> List[Dict[str, Optional[str]]]:
    rows: List[Dict[str, Optional[str]]] = []
    pdf_paths: List[str] = []

    if os.path.isdir(input_dir):
        for root, _, files in os.walk(input_dir):
            for f in files:
                if f.lower().endswith('.pdf'):
                    pdf_paths.append(os.path.join(root, f))
    else:
        if input_dir.lower().endswith('.pdf'):
            pdf_paths.append(input_dir)
        else:
            raise FileNotFoundError(f"No such file or directory: {input_dir}")

    pdf_paths.sort()
    add_paths_to_env()

    for path in pdf_paths:
        try:
            full_text, page_texts = read_pdf_text(path)
            if is_text_pdf(page_texts):
                text = full_text
                source = "text"
            else:
                text = ocr_pdf(path)
                source = "ocr"

            data = parse_invoice_text(text)
            data["File"] = os.path.basename(path)
            data["Parse Source"] = source
            rows.append(data)

        except Exception as e:
            rows.append({
                "File": os.path.basename(path),
                "Error": str(e)
            })

    return rows


def main():
    parser = argparse.ArgumentParser(description="Extract invoice fields from mixed PDFs (text + scanned)")
    parser.add_argument('--in', dest='input_path', required=True, help='Input folder (or single PDF file)')
    parser.add_argument('--out', dest='output_xlsx', default='Extracted_Invoices.xlsx', help='Output Excel file path')
    args = parser.parse_args()

    rows = process_pdfs(args.input_path)

    df = pd.DataFrame(rows, columns=[
        "File",
        "Invoice No",
        "Invoice Date",
        "Invoice Issued To",
        "GSTIN (Issued To)",
        "Name of Who Generated Invoice",
        "GSTIN of That Name",
        "Description",
        "HSN Code",
        "Value (Net of GST)",
        "Quantity",
        "Value Per Piece",
        "Detected Template",
        "Parse Source",
        "Error",
    ])

    df.to_excel(args.output_xlsx, index=False)
    print(f"✅ Extraction complete. Saved: {args.output_xlsx}")
    print(df.fillna("").head())


if __name__ == '__main__':
    main()
