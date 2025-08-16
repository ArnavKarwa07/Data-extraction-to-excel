import os, re, pandas as pd
from pdf2image import convert_from_path
import pytesseract
from pytesseract import Output
from PIL import Image
import cv2
import numpy as np

# ===================== CONFIG (EDIT THESE) =====================
PDF_PATH = "Instructions/Sample_Annex_VI.pdf"              # your input PDF
OUTPUT_XLSX = "Extracted_Invoice_6.xlsx"        # output file
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\tools\poppler-24.08.0\Library\bin"  # folder that contains pdfinfo.exe
# ===============================================================

GSTIN_REGEX_LENIENT = r"\b[0-9OI]{2}[A-Z0-9]{5}[0-9OI]{4}[A-Z][A-Z0-9]{1}[Z2][A-Z0-9]\b"
DATE_REGEX = r"\b(\d{2}[\/\-\.]\d{2}[\/\-\.]\d{4})\b"
HSN_REGEX  = r"\b\d{6,8}\b"
NUM        = r"[0-9][0-9,]*\.?[0-9]*"

# ----------------- Utility functions -----------------
def clean_text(t: str) -> str:
    return (t.replace('’', "'").replace('‘', "'").replace('“','"').replace('”','"')
             .replace('|','1').replace('I','1').replace('O','0').replace('o','0'))

def to_number(x):
    if x is None: return None
    x = x.replace(",", "")
    try: return float(x)
    except: return None

def preprocess_for_ocr(pil_img: Image.Image) -> Image.Image:
    img  = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thr  = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY, 35, 11)
    thr  = cv2.morphologyEx(thr, cv2.MORPH_OPEN, np.ones((1,1), np.uint8), iterations=1)
    return Image.fromarray(thr)

def ocr_page(pil_img: Image.Image):
    proc = preprocess_for_ocr(pil_img)
    cfg  = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(proc, config=cfg, lang='eng')
    data = pytesseract.image_to_data(proc, config=cfg, lang='eng', output_type=Output.DATAFRAME)
    data = data.dropna(subset=['text'])
    data['text'] = data['text'].astype(str).apply(clean_text)
    return clean_text(text), data

def group_lines(df):
    lines = []
    if df.empty: return lines
    for (_,_,ln), g in df.groupby(["block_num","par_num","line_num"]):
        g = g.sort_values("left")
        lines.append((int(min(g["top"])), " ".join(g["text"].tolist())))
    lines.sort(key=lambda x: x[0])
    return [t for _, t in lines]

# ----------------- Improved GSTIN Extraction -----------------
def find_gstins(txt):
    """
    Extract GSTINs even if OCR misreads characters.
    Debug print raw GSTIN-like strings found.
    """
    raw_matches = re.findall(GSTIN_REGEX_LENIENT, txt, flags=re.IGNORECASE)

    def fix_gstin(g):
        return (g.upper()
                  .replace("O","0")
                  .replace("I","1")
                  .replace("{","4")
                  .replace("!","1")
                  .replace("Z2","Z"))

    print("🔎 Raw GSTIN candidates found:", raw_matches)  # DEBUG PRINT

    gstins = [fix_gstin(g) for g in raw_matches if len(g) >= 15]

    g1 = gstins[0] if len(gstins) > 0 else ""
    g2 = gstins[1] if len(gstins) > 1 else ""
    return g1, g2

# ----------------- Header Extraction -----------------
def guess_invoice_no(txt):
    m = re.search(r"(Invoice\s*No[:\-]?\s*)(?P<num>\w+)", txt, flags=re.IGNORECASE)
    if m:
        cand = re.sub(r"[^\w]","", m.group("num"))
        if re.fullmatch(r"\d{6,12}", cand): return cand
    nums = re.findall(r"\b\d{6,10}\b", txt)
    return nums[0] if nums else ""

def extract_header(txt, file_name):
    h = {}
    m = re.search(r"\b\d{13}\b", txt);         h["PLI Request No"] = m.group(0) if m else ""
    h["IFCI No"] = "";                         h["File Name"] = file_name; h["IRN#"] = ""
    m = re.search(DATE_REGEX, txt);            h["Date"] = m.group(1) if m else ""
    h["Invoice#"] = guess_invoice_no(txt)
    sup, cust = find_gstins(txt)
    h["GSTIN of Local Supplier"] = sup
    h["invoice issued to GSTIN"] = cust
    return h

def guess_names(lines):
    supplier = ""; customer = ""
    for i, ln in enumerate(lines):
        if "TAX INVOICE" in ln.upper():
            supplier = " ".join(lines[max(0,i-3):i+1]).strip(); break
    for ln in lines:
        if any(k in ln.lower() for k in ["sold to","bill to","buyer","consignee","invoice issued to"]):
            customer = ln.split(":")[-1].strip(); break
    if not supplier:
        for ln in lines[:8]:
            if any(w in ln.title() for w in ["Pvt","Ltd","LLP","Plywood","Industries","Company","Traders"]):
                supplier = ln.strip(); break
    if not customer:
        for ln in lines[:20]:
            if "Plywood" in ln or "Ltd" in ln or "LLP" in ln:
                customer = ln.strip(); break
    return supplier or "Local Supplier", customer or "Customer"

# ----------------- Line Item Parsing -----------------
def parse_items(lines):
    items = []
    for ln in lines:
        s = " ".join(ln.split())
        hsn_m = re.search(HSN_REGEX, s)
        if not hsn_m: continue
        hsn = hsn_m.group(0)
        m_qty = re.search(rf"({NUM})\s*(Pcs|Nos|Qty|QTY)", s, flags=re.IGNORECASE)
        qty = to_number(m_qty.group(1)) if m_qty else None
        nums = [to_number(n) for n in re.findall(NUM, s)]
        nums = [n for n in nums if n is not None]
        line_total = nums[-1] if nums else None
        name = s.split(hsn,1)[0].strip(" -:,")
        per_piece = round(line_total/qty,2) if (qty and line_total and qty>0) else None
        if not (hsn and (qty or line_total)): continue
        items.append({
            "Name of Part/Component": name,
            "HSN Code": hsn,
            "Quantity": qty if qty is not None else "",
            "Value (net of GST) (Rs.)": line_total if line_total is not None else "",
            "Value per piece (net of GST) (Rs.)": per_piece if per_piece is not None else ""
        })
    return items

# ----------------- Main -----------------
def main():
    pages = convert_from_path(PDF_PATH, dpi=300, poppler_path=POPPLER_PATH)
    text_all, lines_all = [], []
    for p in pages:
        t, df = ocr_page(p)
        text_all.append(t)
        lines_all.extend(group_lines(df))
    full_text = "\n".join(text_all)
    header = extract_header(full_text, os.path.splitext(os.path.basename(PDF_PATH))[0])
    supplier, customer = guess_names(lines_all)
    items = parse_items(lines_all)

    rows, serial = [], 1
    for it in items:
        rows.append({
            "PLI Request No": header.get("PLI Request No",""),
            "IFCI No": header.get("IFCI No",""),
            "File Name": header.get("File Name",""),
            "invoice issued to": customer,
            "invoice issued to GSTIN": header.get("invoice issued to GSTIN",""),
            "#": serial,
            "IRN#": header.get("IRN#",""),
            "Invoice#": header.get("Invoice#",""),
            "Date": header.get("Date",""),
            "Name of Local Supplier": supplier,
            "GSTIN of Local Supplier": header.get("GSTIN of Local Supplier",""),
            "Name of Part/Component": it["Name of Part/Component"],
            "HSN Code": it["HSN Code"],
            "Value (net of GST) (Rs.)": it["Value (net of GST) (Rs.)"],
            "Quantity": it["Quantity"],
            "Value per piece (net of GST) (Rs.)": it["Value per piece (net of GST) (Rs.)"],
        })
        serial += 1

    if not rows:  # keep headers even if nothing matched
        rows = [dict.fromkeys([
            "PLI Request No","IFCI No","File Name","invoice issued to",
            "invoice issued to GSTIN","#","IRN#","Invoice#","Date",
            "Name of Local Supplier","GSTIN of Local Supplier",
            "Name of Part/Component","HSN Code",
            "Value (net of GST) (Rs.)","Quantity",
            "Value per piece (net of GST) (Rs.)"
        ], "")]

    pd.DataFrame(rows).to_excel(OUTPUT_XLSX, index=False)
    print(f"✓ Saved: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
