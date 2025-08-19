import os, re, pandas as pd, glob
from pdf2image import convert_from_path
import pytesseract
from pytesseract import Output
from PIL import Image
import cv2
import numpy as np
from datetime import datetime
import traceback


def check_dependencies():
    """Check if all required packages are installed"""
    required_packages = {
        "pdf2image": "pdf2image",
        "pytesseract": "pytesseract",
        "PIL": "Pillow",
        "cv2": "opencv-python",
        "numpy": "numpy",
        "pandas": "pandas",
    }

    missing_packages = []

    for package, pip_name in required_packages.items():
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(pip_name)

    if missing_packages:
        print("❌ Missing required packages:")
        for package in missing_packages:
            print(f"   - {package}")
        print("\nInstall missing packages with:")
        print(f"pip install {' '.join(missing_packages)}")
        return False

    return True


def validate_config():
    """Validate configuration paths and settings"""
    print("🔧 Validating configuration...")

    # Check if Tesseract path exists
    if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
        print(f"⚠️  Tesseract not found at: {pytesseract.pytesseract.tesseract_cmd}")
        print("   Please install Tesseract OCR or update the path in config")
        return False

    # Check if Poppler path exists
    if not os.path.exists(POPPLER_PATH):
        print(f"⚠️  Poppler not found at: {POPPLER_PATH}")
        print("   Please install Poppler or update the path in config")
        return False

    # Check if input folder exists
    if not os.path.exists(PDF_FOLDER):
        print(f"⚠️  PDF folder not found: {PDF_FOLDER}")
        print("   Please create the folder or update the path in config")
        return False

    print("✅ Configuration validated successfully")
    return True


# ===================== CONFIG (EDIT THESE) =====================
PDF_FOLDER = "fwdsamplesforannexvi"  # folder containing PDF files
OUTPUT_XLSX = "All_Extracted_Invoices.xlsx"  # output file for all invoices
OUTPUT_CSV = "All_Extracted_Invoices.csv"  # output CSV file
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = (
    r"C:\tools\poppler-24.08.0\Library\bin"  # folder that contains pdfinfo.exe
)
# ===============================================================

# Improved regex patterns
GSTIN_REGEX = r"\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9][Z][A-Z0-9]\b"
GSTIN_REGEX_LENIENT = r"\b[0-9OI]{2}[A-Z0-9]{5}[0-9OI]{4}[A-Z][A-Z0-9]{1}[Z2][A-Z0-9]\b"
DATE_PATTERNS = [
    r"\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})\b",
    r"\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2})\b",
    r"\b(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})\b",
]
INVOICE_NO_PATTERNS = [
    r"Invoice\s*(?:No|Number|#)[:\-\s]*([A-Z0-9\/\-]+)",
    r"Invoice[:\-\s]*([A-Z0-9\/\-]+)",
    r"Inv[:\-\s]*([A-Z0-9\/\-]+)",
    r"Bill\s*(?:No|Number|#)[:\-\s]*([A-Z0-9\/\-]+)",
    r"Receipt\s*(?:No|Number|#)[:\-\s]*([A-Z0-9\/\-]+)",
]
HSN_SAC_PATTERNS = [
    r"\b\d{4,8}\b",  # HSN codes
    r"\b9\d{5}\b",  # SAC codes starting with 9
    r"\b99\d{4}\b",  # Common SAC pattern
]
AMOUNT_PATTERNS = [
    r"[₹$]\s*([\d,]+\.?\d*)",
    r"([\d,]+\.?\d*)\s*[₹$]",
    r"\b([\d,]{1,}\.?\d*)\b",
]


# ----------------- Utility functions -----------------
def clean_text(t: str) -> str:
    """Clean text while preserving important characters"""
    if not t:
        return ""
    # Remove common OCR artifacts but keep alphanumeric and important punctuation
    t = re.sub(
        r"[^\w\s\.\-\/\(\)\[\]\{\}\@\#\$\%\^\&\*\+\=\:\;\<\>\,\?\!\'\"₹]", " ", t
    )
    t = re.sub(r"\s+", " ", t)  # Multiple spaces to single
    return t.strip()


def normalize_gstin_chars(text):
    """Fix common OCR errors in GSTIN"""
    return (
        text.upper()
        .replace("O", "0")
        .replace("I", "1")
        .replace("S", "5")
        .replace("Z", "Z")  # Keep Z as is
        .replace("l", "1")
        .replace("o", "0")
    )


def to_number(x):
    """Convert string to number, handling various formats"""
    if x is None or x == "":
        return None
    if isinstance(x, (int, float)):
        return float(x)

    # Remove currency symbols and clean
    x = str(x).strip()
    x = re.sub(r"[₹$,\s]", "", x)

    try:
        return float(x)
    except:
        return None


def preprocess_for_ocr(pil_img: Image.Image) -> Image.Image:
    """Enhanced image preprocessing for better OCR"""
    img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

    # Upscale for better OCR
    h, w = img.shape[:2]
    if max(h, w) < 2000:
        scale = 2000 / max(h, w)
        img = cv2.resize(
            img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC
        )

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Enhance contrast
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)

    # Denoise
    gray = cv2.fastNlMeansDenoising(gray)

    # Adaptive threshold
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
    )

    return Image.fromarray(binary)


def ocr_page_enhanced(pil_img: Image.Image):
    """Enhanced OCR with multiple configurations"""
    processed_img = preprocess_for_ocr(pil_img)

    # Try multiple OCR configurations
    configs = [
        r"--oem 3 --psm 6",  # Single uniform block
        r"--oem 3 --psm 4",  # Single column
        r"--oem 3 --psm 3",  # Fully automatic
        r"--oem 1 --psm 6",  # Legacy engine
    ]

    best_text = ""
    best_data = None
    max_confidence = 0

    for config in configs:
        try:
            text = pytesseract.image_to_string(processed_img, config=config, lang="eng")
            data = pytesseract.image_to_data(
                processed_img, config=config, lang="eng", output_type=Output.DATAFRAME
            )

            # Calculate average confidence
            valid_data = data[data["conf"] > 0]
            if not valid_data.empty:
                avg_conf = valid_data["conf"].mean()
                if avg_conf > max_confidence:
                    max_confidence = avg_conf
                    best_text = text
                    best_data = data
        except:
            continue

    if best_data is not None:
        best_data = best_data.dropna(subset=["text"])
        best_data["text"] = best_data["text"].astype(str).apply(clean_text)
    else:
        best_data = pd.DataFrame()

    return clean_text(best_text), best_data


# ----------------- Enhanced GSTIN Extraction -----------------
def extract_gstins(text):
    """Extract GSTINs with better error correction"""
    # First try strict pattern
    strict_matches = re.findall(GSTIN_REGEX, text.upper())

    # Then try lenient pattern
    lenient_matches = re.findall(GSTIN_REGEX_LENIENT, text.upper())

    # Normalize all matches
    all_matches = list(set(strict_matches + lenient_matches))
    normalized = [normalize_gstin_chars(g) for g in all_matches]

    # Filter valid GSTINs (should be 15 characters)
    valid_gstins = [g for g in normalized if len(g) == 15]

    print(f"🔍 Found GSTINs: {valid_gstins}")

    return valid_gstins


# ----------------- Enhanced Date Extraction -----------------
def extract_dates(text):
    """Extract dates with multiple patterns"""
    dates = []
    for pattern in DATE_PATTERNS:
        matches = re.findall(pattern, text)
        dates.extend(matches)

    # Clean and validate dates
    cleaned_dates = []
    for date_str in dates:
        # Try to parse to validate
        for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"]:
            try:
                parsed = datetime.strptime(date_str, fmt)
                # Only accept reasonable dates (not too old or too future)
                if 2020 <= parsed.year <= 2030:
                    cleaned_dates.append(date_str)
                    break
            except:
                continue

    return list(set(cleaned_dates))


# ----------------- Enhanced Invoice Number Extraction -----------------
def extract_invoice_number(text):
    """Extract invoice number with multiple patterns"""
    for pattern in INVOICE_NO_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            inv_no = match.group(1).strip()
            # Skip if it looks like a date or GSTIN
            if (
                not re.match(r"^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$", inv_no)
                and len(inv_no) != 15
            ):
                return inv_no

    return ""


# ----------------- Enhanced Company Name Extraction -----------------
def extract_company_names(text, lines):
    """Extract supplier and customer names using multiple strategies"""

    # Common company suffixes
    company_suffixes = [
        "LTD",
        "LIMITED",
        "PVT",
        "PRIVATE",
        "INC",
        "CORP",
        "CORPORATION",
        "LLC",
        "SERVICES",
        "SOLUTIONS",
        "TECHNOLOGIES",
        "SOFTWARE",
    ]

    # Find all potential company names
    potential_companies = []

    # Strategy 1: Lines with company suffixes
    for line in lines:
        line_upper = line.upper()
        if any(suffix in line_upper for suffix in company_suffixes):
            # Clean the line
            cleaned = re.sub(r"^[0-9\.\s\-\|]+", "", line).strip()
            if len(cleaned) > 5:
                potential_companies.append(cleaned)

    # Strategy 2: Text blocks around "FROM:", "TO:", "BILL TO:", etc.
    text_upper = text.upper()
    for keyword in [
        "FROM:",
        "TO:",
        "BILL TO:",
        "BILLED TO:",
        "SOLD TO:",
        "SUPPLIER:",
        "VENDOR:",
    ]:
        if keyword in text_upper:
            # Extract text after keyword
            parts = text_upper.split(keyword)
            if len(parts) > 1:
                after_keyword = parts[1][:200]  # First 200 chars after keyword
                # Look for company name in this section
                for line in after_keyword.split("\n")[:5]:  # First 5 lines
                    if any(suffix in line for suffix in company_suffixes):
                        potential_companies.append(line.strip())

    # Strategy 3: GSTIN context
    gstin_matches = extract_gstins(text)
    for gstin in gstin_matches:
        # Find text around GSTIN
        gstin_pos = text.upper().find(gstin)
        if gstin_pos > 0:
            # Look 200 chars before GSTIN for company name
            context = text[max(0, gstin_pos - 200) : gstin_pos]
            context_lines = context.split("\n")
            for line in reversed(context_lines[-5:]):  # Last 5 lines before GSTIN
                line_cleaned = line.strip()
                if len(line_cleaned) > 5 and any(
                    suffix in line_cleaned.upper() for suffix in company_suffixes
                ):
                    potential_companies.append(line_cleaned)
                    break

    # Deduplicate and clean potential companies
    unique_companies = []
    for company in potential_companies:
        company_clean = re.sub(r"[^a-zA-Z0-9\s\&\.\-]", " ", company)
        company_clean = re.sub(r"\s+", " ", company_clean).strip()
        if len(company_clean) > 5 and company_clean not in unique_companies:
            unique_companies.append(company_clean)

    # Return first two unique companies (supplier and customer)
    supplier = unique_companies[0] if len(unique_companies) > 0 else "Unknown Supplier"
    customer = (
        unique_companies[1]
        if len(unique_companies) > 1
        else unique_companies[0] if unique_companies else "Unknown Customer"
    )

    return supplier, customer


# ----------------- Enhanced Item Extraction -----------------
def extract_line_items(text, lines):
    """Extract line items with flexible parsing for different formats"""
    items = []

    # Find HSN/SAC codes in the text
    hsn_sac_codes = []
    for pattern in HSN_SAC_PATTERNS:
        matches = re.findall(pattern, text)
        hsn_sac_codes.extend(matches)

    # Remove duplicates and filter invalid codes
    hsn_sac_codes = list(set(hsn_sac_codes))

    # Filter out obvious non-codes (dates, postal codes, phone numbers)
    filtered_codes = []
    for code in hsn_sac_codes:
        code_int = int(code)
        # Skip years, postal codes, phone numbers
        if (
            (len(code) == 4 and 1900 <= code_int <= 2030)
            or (len(code) == 6 and 400000 <= code_int <= 799999)
            or (len(code) == 10)
        ):
            continue
        filtered_codes.append(code)

    print(f"🔍 Found HSN/SAC codes: {filtered_codes}")

    # For each code, try to extract associated item information
    for code in filtered_codes:
        # Find the line containing this code
        code_line = None
        code_line_index = -1

        for i, line in enumerate(lines):
            if code in line:
                code_line = line
                code_line_index = i
                break

        if not code_line:
            continue

        # Extract item name (text before the code)
        parts = code_line.split(code)
        if len(parts) > 0:
            name_part = parts[0].strip()

            # Clean up name
            name_part = re.sub(
                r"^[0-9\.\s\-\|]+", "", name_part
            )  # Remove leading numbers/symbols
            name_part = re.sub(r"\s+", " ", name_part).strip()

            # Skip if name is too short or just numbers
            if len(name_part) < 3 or re.match(r"^[\d\s\-\.\/]+$", name_part):
                # Try to get name from previous line
                if code_line_index > 0:
                    prev_line = lines[code_line_index - 1]
                    name_part = re.sub(r"^[0-9\.\s\-\|]+", "", prev_line).strip()
                    name_part = re.sub(r"\s+", " ", name_part).strip()

        # Extract numbers from the line and surrounding lines
        numbers = []
        search_lines = []

        # Include current line and next 2 lines for number extraction
        for i in range(max(0, code_line_index), min(len(lines), code_line_index + 3)):
            search_lines.append(lines[i])

        # Extract all numbers from search area
        for line in search_lines:
            line_numbers = []
            for pattern in AMOUNT_PATTERNS:
                matches = re.findall(pattern, line)
                for match in matches:
                    num_val = to_number(match)
                    if num_val and num_val > 0:
                        line_numbers.append(num_val)
            numbers.extend(line_numbers)

        # Remove duplicates and sort
        numbers = sorted(list(set(numbers)))

        # Try to identify quantity, rate, and amount
        quantity = None
        rate = None
        amount = None

        if numbers:
            # Heuristics for different number patterns
            if len(numbers) == 1:
                # Only one number - likely the total amount
                amount = numbers[0]
            elif len(numbers) == 2:
                # Two numbers - likely quantity and amount, or rate and amount
                if numbers[0] < 1000 and numbers[1] > numbers[0]:
                    quantity = numbers[0]
                    amount = numbers[1]
                else:
                    rate = numbers[0]
                    amount = numbers[1]
            elif len(numbers) >= 3:
                # Multiple numbers - try to identify pattern
                # Sort by value
                small_nums = [n for n in numbers if n < 1000]  # Likely quantities
                large_nums = [n for n in numbers if n >= 1000]  # Likely amounts

                if small_nums:
                    quantity = small_nums[0]
                if large_nums:
                    amount = large_nums[-1]  # Largest is usually total
                    if len(large_nums) > 1:
                        rate = large_nums[0]  # Smaller large number might be rate

        # Calculate missing values if possible
        if quantity and amount and not rate:
            rate = round(amount / quantity, 2)
        elif quantity and rate and not amount:
            amount = round(quantity * rate, 2)
        elif rate and amount and not quantity:
            quantity = round(amount / rate, 2)

        # Only add items with meaningful data
        if name_part and len(name_part) > 2 and (quantity or amount):
            items.append(
                {
                    "Name of Part/Component": name_part,
                    "HSN Code": code,
                    "Quantity": quantity if quantity is not None else "",
                    "Value (net of GST) (Rs.)": amount if amount is not None else "",
                    "Value per piece (net of GST) (Rs.)": (
                        rate if rate is not None else ""
                    ),
                }
            )

    # If no items found with codes, try alternative extraction
    if not items:
        items = extract_items_alternative(lines)

    return items


def extract_items_alternative(lines):
    """Alternative item extraction when HSN/SAC codes aren't clearly found"""
    items = []

    # Look for lines that seem like product/service descriptions
    for i, line in enumerate(lines):
        line_clean = line.strip()

        # Skip headers and footer lines
        if any(
            skip in line_clean.upper()
            for skip in [
                "INVOICE",
                "GSTIN",
                "ADDRESS",
                "PHONE",
                "EMAIL",
                "BANK",
                "IFSC",
                "ACCOUNT",
                "TOTAL",
                "TAX",
                "CGST",
                "SGST",
                "IGST",
                "AUTHORIZED",
                "SIGNATURE",
                "TERMS",
                "CONDITIONS",
                "AMOUNT IN WORDS",
            ]
        ):
            continue

        # Look for lines with meaningful product descriptions and numbers
        if (
            len(line_clean) > 10
            and re.search(r"[a-zA-Z]{3,}", line_clean)  # Has meaningful text
            and re.search(r"\d", line_clean)
        ):  # Has numbers

            # Extract numbers from this line
            numbers = []
            for pattern in AMOUNT_PATTERNS:
                matches = re.findall(pattern, line_clean)
                for match in matches:
                    num_val = to_number(match)
                    if num_val and num_val > 0:
                        numbers.append(num_val)

            if numbers:
                # Use the line as product description
                description = re.sub(r"[\d,₹$\.\-]+", "", line_clean)
                description = re.sub(r"\s+", " ", description).strip()

                if len(description) > 3:
                    # Try to find HSN/SAC in surrounding lines
                    hsn_code = ""
                    for j in range(max(0, i - 1), min(len(lines), i + 2)):
                        for pattern in HSN_SAC_PATTERNS:
                            matches = re.findall(pattern, lines[j])
                            if matches:
                                hsn_code = matches[0]
                                break
                        if hsn_code:
                            break

                    amount = max(numbers) if numbers else ""
                    quantity = (
                        min(numbers) if len(numbers) > 1 and min(numbers) < 100 else ""
                    )

                    items.append(
                        {
                            "Name of Part/Component": description,
                            "HSN Code": hsn_code,
                            "Quantity": quantity,
                            "Value (net of GST) (Rs.)": amount,
                            "Value per piece (net of GST) (Rs.)": "",
                        }
                    )

    return items


# ----------------- Enhanced Header Extraction -----------------
def extract_header_info(text, file_name):
    """Extract header information with multiple strategies"""
    header = {
        "PLI Request No": "",
        "IFCI No": "",
        "File Name": os.path.splitext(file_name)[0],
        "IRN#": "",
        "Invoice#": "",
        "Date": "",
    }

    # Extract IRN
    irn_match = re.search(r"IRN[:\-\s]*([a-f0-9]{64})", text, re.IGNORECASE)
    if irn_match:
        header["IRN#"] = irn_match.group(1)

    # Extract Invoice Number
    header["Invoice#"] = extract_invoice_number(text)

    # Extract Date
    dates = extract_dates(text)
    if dates:
        header["Date"] = dates[0]  # Take first valid date

    return header


# ----------------- Main Processing Function -----------------
def process_single_pdf(pdf_path):
    """Process a single PDF with enhanced extraction"""
    file_name = os.path.basename(pdf_path)
    print(f"📄 Processing: {file_name}")

    try:
        # Convert PDF to images
        pages = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)

        all_text = []
        all_lines = []

        # Process each page
        for page_num, page in enumerate(pages):
            print(f"  📖 Processing page {page_num + 1}/{len(pages)}")
            text, data = ocr_page_enhanced(page)
            all_text.append(text)

            # Group text into lines
            if not data.empty:
                page_lines = []
                for (_, _, _), group in data.groupby(
                    ["block_num", "par_num", "line_num"]
                ):
                    group = group.sort_values("left")
                    line_text = " ".join(group["text"].tolist())
                    if line_text.strip():
                        page_lines.append((int(group["top"].min()), line_text))

                # Sort by vertical position
                page_lines.sort(key=lambda x: x[0])
                all_lines.extend([line[1] for line in page_lines])

        # Combine all text
        full_text = "\n".join(all_text)

        # Extract information
        header = extract_header_info(full_text, file_name)
        gstins = extract_gstins(full_text)
        supplier, customer = extract_company_names(full_text, all_lines)
        items = extract_line_items(full_text, all_lines)

        # Assign GSTINs (first to supplier, second to customer if available)
        supplier_gstin = gstins[0] if len(gstins) > 0 else ""
        customer_gstin = gstins[1] if len(gstins) > 1 else ""

        # Create output rows
        rows = []
        serial = 1

        # If no items found, create at least one row with header info
        if not items:
            items = [
                {
                    "Name of Part/Component": "No items extracted",
                    "HSN Code": "",
                    "Quantity": "",
                    "Value (net of GST) (Rs.)": "",
                    "Value per piece (net of GST) (Rs.)": "",
                }
            ]

        for item in items:
            row = {
                "PLI Request No": header.get("PLI Request No", ""),
                "IFCI No": header.get("IFCI No", ""),
                "File Name": header.get("File Name", ""),
                "invoice issued to": customer,
                "invoice issued to GSTIN": customer_gstin,
                "#": serial,
                "IRN#": header.get("IRN#", ""),
                "Invoice#": header.get("Invoice#", ""),
                "Date": header.get("Date", ""),
                "Name of Local Supplier": supplier,
                "GSTIN of Local Supplier": supplier_gstin,
                "Name of Part/Component": item["Name of Part/Component"],
                "HSN Code": item["HSN Code"],
                "Value (net of GST) (Rs.)": item["Value (net of GST) (Rs.)"],
                "Quantity": item["Quantity"],
                "Value per piece (net of GST) (Rs.)": item[
                    "Value per piece (net of GST) (Rs.)"
                ],
            }
            rows.append(row)
            serial += 1

        print(f"✅ Extracted {len(rows)} item(s) from {file_name}")

        # Debug information
        print(f"   📋 Supplier: {supplier}")
        print(f"   👥 Customer: {customer}")
        print(f"   🔢 Invoice#: {header.get('Invoice#', 'Not found')}")
        print(f"   📅 Date: {header.get('Date', 'Not found')}")
        print(f"   🏢 GSTINs: {gstins}")

        return rows

    except Exception as e:
        print(f"❌ Error processing {file_name}: {str(e)}")
        print(f"   📋 Error details: {traceback.format_exc()}")

        # Return minimal row for error tracking
        return [
            {
                "PLI Request No": "",
                "IFCI No": "",
                "File Name": os.path.splitext(file_name)[0],
                "invoice issued to": f"Error: {str(e)[:50]}",
                "invoice issued to GSTIN": "",
                "#": 1,
                "IRN#": "",
                "Invoice#": "",
                "Date": "",
                "Name of Local Supplier": "",
                "GSTIN of Local Supplier": "",
                "Name of Part/Component": "Error during extraction",
                "HSN Code": "",
                "Value (net of GST) (Rs.)": "",
                "Quantity": "",
                "Value per piece (net of GST) (Rs.)": "",
            }
        ]


# ----------------- Main Function -----------------
def main():
    """Main processing function"""
    print("🚀 Starting Enhanced Multi-Format Invoice Extractor")
    print("=" * 60)

    # Check dependencies
    if not check_dependencies():
        return

    # Validate configuration
    if not validate_config():
        return

    # Get all PDF files
    pdf_pattern = os.path.join(PDF_FOLDER, "*.pdf")
    pdf_files = glob.glob(pdf_pattern)

    if not pdf_files:
        print(f"❌ No PDF files found in {PDF_FOLDER}")
        return

    print(f"📁 Found {len(pdf_files)} PDF files to process")
    print("-" * 60)

    all_rows = []
    processed_count = 0
    success_count = 0
    error_count = 0

    # Process each PDF
    for pdf_file in pdf_files:
        try:
            rows = process_single_pdf(pdf_file)
            all_rows.extend(rows)
            processed_count += 1

            # Check if extraction was successful
            if rows and any(
                row.get("Name of Part/Component", "") != "Error during extraction"
                for row in rows
            ):
                success_count += 1
            else:
                error_count += 1

            print(f"📊 Progress: {processed_count}/{len(pdf_files)} files processed")
            print("-" * 40)

        except Exception as e:
            print(f"💥 Critical error with {os.path.basename(pdf_file)}: {str(e)}")
            error_count += 1
            processed_count += 1

    # Save results
    if all_rows:
        df = pd.DataFrame(all_rows)

        # Save Excel file
        try:
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"✅ Excel file saved: {OUTPUT_XLSX}")
        except Exception as e:
            print(f"⚠️ Could not save Excel file: {e}")

        # Save CSV file
        try:
            df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8")
            print(f"✅ CSV file saved: {OUTPUT_CSV}")
        except Exception as e:
            print(f"⚠️ Could not save CSV file: {e}")

        # Print summary
        print("\n" + "=" * 60)
        print("📈 PROCESSING SUMMARY")
        print("=" * 60)
        print(f"📁 Total files found: {len(pdf_files)}")
        print(f"✅ Successfully processed: {success_count}")
        print(f"❌ Files with errors: {error_count}")
        print(f"📊 Total records extracted: {len(all_rows)}")

        # Summary by file
        if len(all_rows) > 1:
            print(f"\n📋 Items extracted per file:")
            summary = df.groupby("File Name").size().reset_index(name="Items")
            summary = summary.sort_values("Items", ascending=False)
            print(summary.to_string(index=False))

        # Show sample of extracted data
        print(f"\n🔍 Sample extracted data (first 3 rows):")
        sample_cols = [
            "File Name",
            "Invoice#",
            "Date",
            "Name of Local Supplier",
            "invoice issued to",
            "HSN Code",
            "Value (net of GST) (Rs.)",
        ]
        print(df[sample_cols].head(3).to_string(index=False))

    else:
        print("❌ No data was extracted from any files!")


# Execute the main function when script is run
if __name__ == "__main__":
    main()
