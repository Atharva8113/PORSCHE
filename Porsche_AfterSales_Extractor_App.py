"""
Porsche After Sales – Invoice Extractor
Extracts structured line-item data from Porsche AG After-Sales invoices (PDF).
Handles the portrait layout with the interleaved ORIGINAL watermark letters,
date stamps, and description lines between item rows.
Outputs CSV with INR/USD-style number formatting.
"""

import os
import sys
import re
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from typing import Optional

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


# ---------- RESOURCE PATH FUNCTION ----------
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


# ---------- EUR-STYLE → STANDARD NUMBER FORMATTING ----------
def convert_eur_to_standard_format(value_str: str) -> str:
    """
    Convert European number format to standard (INR/USD) format.
    European: 2.236,90 → Standard: 2,236.90
    European: 0,297   → Standard: 0.297
    European: 43.760,64 → Standard: 43,760.64
    If already standard-style (commas as thousands, period as decimal), pass through.
    """
    if not value_str or not isinstance(value_str, str):
        return value_str

    value_str = value_str.strip()

    # Detect EUR format: both '.' and ',' present, where comma comes AFTER the
    # last period → European style (e.g. "164.675,59" or "1.533,96")
    if '.' in value_str and ',' in value_str:
        last_dot = value_str.rfind('.')
        last_comma = value_str.rfind(',')

        if last_comma > last_dot:
            # European format: periods are thousands, comma is decimal
            converted = value_str.replace('.', '').replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            # Already standard format (comma is thousands, period is decimal)
            return value_str

    elif ',' in value_str and '.' not in value_str:
        # Ambiguous: could be EUR decimal or standard thousands.
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            # Likely European decimal (e.g. "0,297" or "517,80")
            converted = value_str.replace(',', '.')
            try:
                num = float(converted)
                decimal_places = len(parts[1])
                return f"{num:,.{decimal_places}f}"
            except ValueError:
                return value_str
        else:
            return value_str
    else:
        return value_str


def eur_str_to_float(value_str: str) -> float:
    """
    Parse a European-formatted number string to a Python float.
    e.g. '2.236,90' → 2236.90, '0,297' → 0.297
    Also handles standard-style strings.
    """
    if not value_str or not isinstance(value_str, str):
        return 0.0
    value_str = value_str.strip()
    if '.' in value_str and ',' in value_str:
        last_dot = value_str.rfind('.')
        last_comma = value_str.rfind(',')
        if last_comma > last_dot:
            return float(value_str.replace('.', '').replace(',', '.'))
        else:
            return float(value_str.replace(',', ''))
    elif ',' in value_str:
        parts = value_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 3:
            return float(value_str.replace(',', '.'))
        else:
            return float(value_str.replace(',', ''))
    else:
        try:
            return float(value_str)
        except ValueError:
            return 0.0


def smart_format_number(value_str: str) -> str:
    """
    Intelligently format a number string from EUR-style to standard.
    Always output standard-style (commas=thousands, period=decimal).
    """
    if not value_str or not isinstance(value_str, str):
        return value_str
    return convert_eur_to_standard_format(value_str.strip())


# ---------- CORE EXTRACTION LOGIC ----------
#
# Porsche After-Sales invoice format (per data page):
#
#   Header row (repeated on each page):
#     pos Customer Part number Code Quantity unit weight Goods- GRO COC Net
#     order                         price    in kg  No.            Value
#
#   Item block pattern (example):
#     0001 PUNBCKORDR PAF.001.987 6 1,94 0,204 73181660 DK 11,64
#     O 17122025                                          ← watermark letter + date
#     R hex. nut, self-locking with washer                ← watermark letter + description
#     I                                                   ← watermark letter (no data)
#
#   Fields in the item line (left to right):
#     POS  CUSTOMER_ORDER  PART_NUMBER  QTY  UNIT_PRICE  WEIGHT_KG  HS_CODE  COUNTRY  NET_VALUE
#
#   The ORIGINAL watermark cycles letters: O, R, I, G, I, N, A, L
#   across the left margin. Description text appears on lines after the
#   item line, mixed with the watermark & date lines.

# Regex to match item lines
# Example: 0001 PUNBCKORDR PAF.001.987 6 1,94 0,204 73181660 DK 11,64
# Example: 0067 PUNBCKORDR 9Y0.807.061.E .OK1 1 226,57 1,662 87081090 SK 226,57
# Example: 0081 PUNBCKORDR 971.831.721. Y 1 71,44 0,674 40169300 CZ 71,44
#
# Strategy: parse from the RIGHT (numeric fields are predictable), then extract part number.
# Right-to-left anchored fields:
#   NET_VALUE       = EUR-style number  (e.g. 11,64 or 1.533,96)
#   COUNTRY_CODE    = 2 uppercase letters
#   HS_CODE         = 8 digits
#   WEIGHT_KG       = EUR-style number
#   UNIT_PRICE      = EUR-style number
#   QTY             = integer (may have EUR thousands separator)
#   PART_NUMBER     = everything between PUNBCKORDR and QTY
#   POS             = 4-digit position
#   CUSTOMER_ORDER  = always PUNBCKORDR (or similar)

ITEM_LINE_RE = re.compile(
    r'^\s*(\d{4})\s+'                    # G1: Position (4 digits)
    r'(\S+)\s+'                          # G2: Customer order code (e.g. PUNBCKORDR)
    r'(.+?)\s+'                          # G3: Part number (greedy-minimal, can contain spaces)
    r'(\d+)\s+'                          # G4: Quantity (integer)
    r'([0-9.,]+)\s+'                     # G5: Unit price (EUR-style)
    r'([0-9.,]+)\s+'                     # G6: Weight in kg (EUR-style)
    r'(\d{8})\s+'                        # G7: HS Code (8 digits)
    r'([A-Z]{2})\s+'                     # G8: Country code (2 letters)
    r'([0-9.,]+)'                        # G9: Net value (EUR-style)
    r'\s*$'
)

# Lines that are just watermark letters (O, R, I, G, N, A, L) or contain
# the date stamp (e.g. "17122025"), possibly with a watermark letter prefix
WATERMARK_DATE_RE = re.compile(
    r'^\s*[ORIGINAL]?\s*\d{8}[A-Z]?\s*$'
)

# Single watermark letter on its own line
SINGLE_LETTER_RE = re.compile(r'^\s*[ORIGINAL]\s*$')

# Date stamp pattern: 8 digits optionally followed by a letter (e.g. 15122025A)
DATE_STAMP_RE = re.compile(r'^\d{8}[A-Z]?\s*$')

# Page header/footer patterns to skip
SKIP_PATTERNS = [
    re.compile(r'^SKODA AUTO Volkswagen', re.IGNORECASE),
    re.compile(r'^INVOICE\s+Private', re.IGNORECASE),
    re.compile(r'^15th Fl\.Commerz', re.IGNORECASE),
    re.compile(r'^Dr\.\s*Ing', re.IGNORECASE),
    re.compile(r'^400063\s+GOREGAON', re.IGNORECASE),
    re.compile(r'^pos\s+Customer\s+Part', re.IGNORECASE),
    re.compile(r'^order\s+price\s+in\s+kg', re.IGNORECASE),
    re.compile(r'^[A-Z]{4}\d+\s+Gross\s+KG', re.IGNORECASE),
    re.compile(r'^\d+[A-Z]?\s*$', re.IGNORECASE), # Skip internal codes like 12026, 12026B
    re.compile(r'^[A-Z]+\d+\s*$', re.IGNORECASE), # Skip internal codes like Kochi21
    re.compile(r'^Company\s+number', re.IGNORECASE),
    re.compile(r'^Customer\s+no', re.IGNORECASE),
    re.compile(r'^7740000\s+No\.', re.IGNORECASE),
    re.compile(r'^7740000\s+of\s+', re.IGNORECASE),
    re.compile(r'^tax-free\s+supply', re.IGNORECASE),
    re.compile(r'^Page\s+\d+\s+of\s+\d+', re.IGNORECASE),
    re.compile(r'^Service\s+date\s+equals', re.IGNORECASE),
    re.compile(r'Dr\.\s*Ing\.?\s*h\.c\.F', re.IGNORECASE),
    re.compile(r'^Place\s+of\s+the\s+company', re.IGNORECASE),
    re.compile(r'^Court\s+of\s+Registry', re.IGNORECASE),
    re.compile(r'^District\s+Court', re.IGNORECASE),
    re.compile(r'^VAT-No', re.IGNORECASE),
    re.compile(r'^Tax\s+No', re.IGNORECASE),
    re.compile(r'^WEEE-Reg', re.IGNORECASE),
    re.compile(r'^Terms\s+of\s+delivery', re.IGNORECASE),
    re.compile(r'^Terms\s+of\s+payment', re.IGNORECASE),
    re.compile(r'^CPT\s*;', re.IGNORECASE),
    re.compile(r'^Invoice\s+address', re.IGNORECASE),
    re.compile(r'^Shipping\s+address', re.IGNORECASE),
    re.compile(r'^Net\s+value\s+of\s+goods', re.IGNORECASE),
    re.compile(r'^Freight\s+charges', re.IGNORECASE),
    re.compile(r'^Total\s+Net\s+Value', re.IGNORECASE),
    re.compile(r'^Total\s+Value', re.IGNORECASE),
    re.compile(r'^Oberoi\s+Garden', re.IGNORECASE),
    re.compile(r'^Off\.\s+Western', re.IGNORECASE),
    re.compile(r'^INDIEN', re.IGNORECASE),
    re.compile(r'Markings:', re.IGNORECASE),
    re.compile(r'^Porscheplatz', re.IGNORECASE),
    re.compile(r'^D-70435', re.IGNORECASE),
    re.compile(r'Landesbank', re.IGNORECASE),
    re.compile(r'Commerzbank', re.IGNORECASE),
    re.compile(r'Seite\s+\d+', re.IGNORECASE),
    re.compile(r'IBAN\s+DE', re.IGNORECASE),
    re.compile(r'^\d{10}\s+\d+', re.IGNORECASE),  # Order numbers like 6309945042 02
    re.compile(r'^Private\s+Limited', re.IGNORECASE),
    re.compile(r'^Central\s+Logistics', re.IGNORECASE),
    re.compile(r'^410501\s+KHED', re.IGNORECASE),
    re.compile(r'^A\d{8}', re.IGNORECASE),  # Markings like A77400001
    re.compile(r'Carton', re.IGNORECASE),
    re.compile(r'Container/Loading', re.IGNORECASE),
    re.compile(r'Airfreight', re.IGNORECASE),
    re.compile(r'Packaging\s+type', re.IGNORECASE),
    re.compile(r'Weight\s+in\s+kg', re.IGNORECASE),
    re.compile(r'Gross\s+\d', re.IGNORECASE),
    re.compile(r'Net\s+\d{2,3},', re.IGNORECASE),  # Net 131,540
    re.compile(r'Order\s+number', re.IGNORECASE),
    re.compile(r'Shipping\s+type', re.IGNORECASE),
    re.compile(r'Chairman\s+of', re.IGNORECASE),
    re.compile(r'Supervisory\s+Board', re.IGNORECASE),
    re.compile(r'Board:', re.IGNORECASE),
    re.compile(r'Stuttgart', re.IGNORECASE),
    re.compile(r'BIC:', re.IGNORECASE),
    re.compile(r'For\s+USD\s+only', re.IGNORECASE),
    re.compile(r'J\.P\.\s+Morgan', re.IGNORECASE),
]

# The ORIGINAL watermark letters that appear as single characters on their own lines
WATERMARK_LETTERS = set('ORIGINAL')


def should_skip_line(line: str) -> bool:
    """Check if a line should be skipped (header/footer/non-data)."""
    if not line or not line.strip():
        return True
    stripped = line.strip()

    # Single watermark letter
    if len(stripped) == 1 and stripped in WATERMARK_LETTERS:
        return True

    # Date stamp (8 digits, possibly with trailing letter)
    if DATE_STAMP_RE.match(stripped):
        return True

    # Watermark letter + date stamp (e.g. "O 17122025")
    if len(stripped) >= 2 and stripped[0] in WATERMARK_LETTERS and stripped[1] == ' ':
        rest = stripped[2:].strip()
        if DATE_STAMP_RE.match(rest):
            return True

    # Header/footer patterns
    for pattern in SKIP_PATTERNS:
        if pattern.search(stripped):
            return True

    return False


def is_description_line(line: str) -> bool:
    """
    Check if a line is a description line (text that belongs to the previous item).
    Description lines are non-numeric text after the item line, often prefixed
    with a watermark letter (e.g. "R hex. nut, self-locking with washer").
    """
    stripped = line.strip()
    if not stripped:
        return False

    # Don't treat item lines as description
    if ITEM_LINE_RE.match(stripped):
        return False

    # Don't treat skippable lines as description
    # (but we need to check AFTER the watermark+description patterns)

    # Lines like "R hex. nut, self-locking with washer" or just "SCREW"
    # A watermark-prefixed description: single letter + space + text
    if len(stripped) > 2 and stripped[0] in WATERMARK_LETTERS and stripped[1] == ' ':
        rest = stripped[2:].strip()
        # If rest is NOT a date stamp, it's description text
        if rest and not DATE_STAMP_RE.match(rest):
            return True

    # Pure text description (e.g. "Hexagon collar nut", "SCREW")
    # Must contain at least one alphabetic character and NOT match known patterns
    if any(c.isalpha() for c in stripped):
        # Not a standalone watermark letter
        if len(stripped) > 1:
            return True

    return False


def extract_description(line: str) -> str:
    """Extract description text from a line, stripping watermark prefix if present."""
    stripped = line.strip()
    if len(stripped) > 2 and stripped[0] in WATERMARK_LETTERS and stripped[1] == ' ':
        rest = stripped[2:].strip()
        if rest and not DATE_STAMP_RE.match(rest):
            return rest
    return stripped


def extract_porsche_aftersales_invoice(pdf_path: str) -> dict:
    """
    Extract all line-item data from a Porsche AG After-Sales invoice PDF.
    Returns a dict with header info and a list of line items.
    """
    invoice_number: str = ""
    invoice_date: str = ""
    currency: str = "EUR"
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        # --- Extract header info from first page ---
        first_page_text = pdf.pages[0].extract_text() or ""

        # Invoice number: "INVOICE 7740000 No. 1394384215"
        inv_match = re.search(
            r'INVOICE\s+\d+\s+No\.\s+(\d+)', first_page_text
        )
        if inv_match:
            invoice_number = inv_match.group(1)

        # Invoice date: "7740000 of 17.12.2025"
        date_match = re.search(
            r'\d+\s+of\s+(\d{2}\.\d{2}\.\d{4})', first_page_text
        )
        if date_match:
            invoice_date = date_match.group(1)

        # Currency: detect from "Net value of goods EUR"
        curr_match = re.search(
            r'Net\s+value\s+of\s+goods\s+(EUR|USD|INR)',
            first_page_text,
            re.IGNORECASE
        )
        if curr_match:
            currency = curr_match.group(1).upper()

        # --- Process each data page ---
        for page_idx, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""

            # Skip the last page (summary-only, no line items)
            if 'tax-free supply' in page_text and 'Service date equals' in page_text:
                continue

            # Skip first page if it has no item lines (header-only page)
            if page_idx == 0:
                # Check if first page has item data (starts with "pos Customer")
                if 'pos Customer Part number' not in page_text:
                    continue

            lines = page_text.split('\n')

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # Try to parse as item line
                m = ITEM_LINE_RE.match(line)
                if m:
                    part_number_raw = m.group(3).strip()
                    qty_str = m.group(4)
                    unit_price_str = m.group(5)
                    weight_str = m.group(6)
                    hs_code = m.group(7)
                    country_code = m.group(8)
                    net_value_str = m.group(9)

                    # Clean part number: remove internal spaces and dots
                    part_number = part_number_raw.replace(' ', '').replace('.', '')

                    # Parse quantity as integer
                    try:
                        quantity = int(qty_str)
                    except ValueError:
                        quantity = 0

                    # Collect description from subsequent lines
                    description_parts: list[str] = []
                    j = i + 1
                    while j < len(lines):
                        next_line = lines[j].strip()

                        # Stop if we hit the next item line
                        if ITEM_LINE_RE.match(next_line):
                            break

                        # Stop if we hit a page marker
                        if re.match(r'^Page\s+\d+\s+of\s+\d+', next_line, re.IGNORECASE):
                            break

                        # Skip blank lines
                        if not next_line:
                            j += 1
                            continue

                        # Skip single watermark letters
                        if len(next_line) == 1 and next_line in WATERMARK_LETTERS:
                            j += 1
                            continue

                        # Skip date stamps
                        if DATE_STAMP_RE.match(next_line):
                            j += 1
                            continue

                        # Check for watermark letter + date stamp
                        if (len(next_line) >= 2 and
                                next_line[0] in WATERMARK_LETTERS and
                                next_line[1] == ' '):
                            rest = next_line[2:].strip()
                            if DATE_STAMP_RE.match(rest):
                                j += 1
                                continue
                            # Watermark letter + description text
                            if rest and not should_skip_line(rest):
                                description_parts.append(rest)
                                j += 1
                                continue

                        # Skip header/footer lines
                        if should_skip_line(next_line):
                            j += 1
                            continue

                        # This is a description line
                        description_parts.append(next_line)
                        j += 1

                    description = ' '.join(description_parts).strip()

                    # Format numbers from EUR to standard style
                    formatted_unit_price = smart_format_number(unit_price_str)
                    formatted_net_value = smart_format_number(net_value_str)
                    formatted_weight = smart_format_number(weight_str)

                    item = {
                        "Invoice Number": invoice_number,
                        "Invoice Date": invoice_date,
                        "Part No.": part_number,
                        "Description": description,
                        "Weight in Kg": formatted_weight,
                        "Country Code": country_code,
                        "HS Code": hs_code,
                        "Default": "(AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION)",
                        "Quantity": str(quantity),
                        "Unit Price": formatted_unit_price,
                        "Net Value": formatted_net_value,
                    }
                    line_items.append(item)

                i += 1

    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "currency": currency,
        "items": line_items,
    }


# ---------- CSV OUTPUT ----------
def write_csv(output_path: str, all_records: list[dict]) -> None:
    """Write all extracted records to a single CSV file."""
    if not all_records:
        return

    fieldnames = [
        "Invoice Number",
        "Invoice Date",
        "Part No.",
        "Description",
        "Weight in Kg",
        "Country Code",
        "HS Code",
        "Default",
        "Quantity",
        "Unit Price",
        "Net Value",
    ]

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in all_records:
            # Wrap Part No. so Excel preserves leading zeros
            safe_record = dict(record)
            part_no = safe_record.get("Part No.", "")
            if part_no and (part_no[0] == '0' or part_no.isdigit()):
                safe_record["Part No."] = f'="{part_no}"'
            writer.writerow(safe_record)


# ---------- NAGARKOT GUI IMPLEMENTATION ----------
class PorscheAfterSalesExtractorGUI:
    """Porsche AG After Sales Invoice Extractor – Nagarkot Branded GUI."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Porsche After Sales – Invoice Extractor")
        self.root.geometry("1200x750")
        self.root.state('zoomed')

        # Nagarkot brand palette
        self.bg_color = "#ffffff"
        self.brand_color = "#0056b3"
        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style()
        self.style.theme_use('clam')

        # --- Style configuration ---
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure(
            "TLabel", background=self.bg_color, font=("Segoe UI", 10)
        )
        self.style.configure(
            "Header.TLabel",
            font=("Helvetica", 18, "bold"),
            foreground=self.brand_color,
            background=self.bg_color,
        )
        self.style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 11),
            foreground="gray",
            background=self.bg_color,
        )
        self.style.configure(
            "Footer.TLabel",
            font=("Segoe UI", 9),
            foreground="#555555",
            background=self.bg_color,
        )
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background=self.brand_color,
            foreground="white",
            borderwidth=0,
            focuscolor=self.brand_color,
        )
        self.style.map("Primary.TButton", background=[('active', '#004494')])
        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            background="#f0f0f0",
            foreground="#333333",
            borderwidth=1,
        )
        self.style.map("Secondary.TButton", background=[('active', '#e0e0e0')])
        self.style.configure("TLabelframe", background=self.bg_color)
        self.style.configure(
            "TLabelframe.Label",
            background=self.bg_color,
            foreground=self.brand_color,
            font=("Segoe UI", 10, "bold"),
        )
        self.style.configure(
            "Treeview", font=("Segoe UI", 9), rowheight=25
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            foreground=self.brand_color,
        )

        self.setup_ui()
        self.selected_files: list[str] = []

    # ----- UI SETUP -----
    def setup_ui(self) -> None:
        # ---------- HEADER ----------
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", pady=20, padx=20)
        header_frame.columnconfigure(0, weight=0)
        header_frame.columnconfigure(1, weight=1)
        header_frame.columnconfigure(2, weight=0)

        # Logo (Left)
        try:
            if Image and ImageTk:
                logo_path = resource_path("Nagarkot Logo.png")
                if os.path.exists(logo_path):
                    pil_img = Image.open(logo_path)
                    h_pct = 20 / float(pil_img.size[1])
                    w_size = int(float(pil_img.size[0]) * h_pct)
                    pil_img = pil_img.resize(
                        (w_size, 20), Image.Resampling.LANCZOS
                    )
                    self.logo_img = ImageTk.PhotoImage(pil_img)
                    logo_lbl = ttk.Label(header_frame, image=self.logo_img)
                    logo_lbl.grid(
                        row=0, column=0, rowspan=2, sticky="w", padx=(0, 20)
                    )
                else:
                    print("Warning: Nagarkot Logo.png not found.")
                    ttk.Label(
                        header_frame, text="[LOGO]", foreground="gray"
                    ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
            else:
                ttk.Label(
                    header_frame, text="[PIL Missing]", foreground="red"
                ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
        except Exception as e:
            print(f"Error loading logo: {e}")
            ttk.Label(
                header_frame, text="[LOGO ERROR]", foreground="red"
            ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))

        # Title (Center)
        title_lbl = ttk.Label(
            header_frame,
            text="Porsche After Sales – Invoice Extractor",
            style="Header.TLabel",
        )
        title_lbl.grid(row=0, column=1, sticky="n")
        subtitle_lbl = ttk.Label(
            header_frame,
            text="Extract line-item data from Porsche AG After-Sales invoices",
            style="Subtitle.TLabel",
        )
        subtitle_lbl.grid(row=1, column=1, sticky="n")

        # ---------- FOOTER (Packed first to reserve bottom space) ----------
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")

        copyright_lbl = ttk.Label(
            footer_frame,
            text="© Nagarkot Forwarders Pvt Ltd",
            style="Footer.TLabel",
        )
        copyright_lbl.pack(side="left", anchor="s")

        self.btn_run = ttk.Button(
            footer_frame,
            text="     Extract & Generate CSV     ",
            command=self.run_extraction,
            style="Primary.TButton",
        )
        self.btn_run.pack(side="right", padx=10, pady=5)

        # ---------- MAIN CONTENT ----------
        content_frame = ttk.Frame(self.root, padding="20 10 20 10")
        content_frame.pack(fill="both", expand=True)

        # --- File Selection ---
        file_frame = ttk.LabelFrame(
            content_frame, text="File Selection", padding="15"
        )
        file_frame.pack(fill="x", pady=(0, 15))

        btn_container = ttk.Frame(file_frame)
        btn_container.pack(fill="x")

        self.btn_select = ttk.Button(
            btn_container,
            text="Select PDFs",
            command=self.select_files,
            style="Secondary.TButton",
        )
        self.btn_select.pack(side="left", padx=(0, 10))

        self.btn_clear = ttk.Button(
            btn_container,
            text="Clear List",
            command=self.clear_files,
            style="Secondary.TButton",
        )
        self.btn_clear.pack(side="left")

        self.lbl_count = ttk.Label(
            btn_container, text="No files selected", style="TLabel"
        )
        self.lbl_count.pack(side="left", padx=(20, 0))

        # --- Output Settings ---
        output_frame = ttk.LabelFrame(
            content_frame, text="Output Settings", padding="15"
        )
        output_frame.pack(fill="x", pady=(0, 15))

        # --- Processing Mode (Combined vs Individual) ---
        ttk.Label(output_frame, text="Processing Mode:").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=5
        )

        mode_frame = ttk.Frame(output_frame)
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="w")

        self.mode_var = tk.StringVar(value="combined")

        self.rb_combined = ttk.Radiobutton(
            mode_frame,
            text="Combined (All in one CSV)",
            variable=self.mode_var,
            value="combined",
            command=self.toggle_filename_state,
        )
        self.rb_combined.pack(side="left", padx=(0, 15))

        self.rb_individual = ttk.Radiobutton(
            mode_frame,
            text="Individual (Separate CSV per invoice)",
            variable=self.mode_var,
            value="individual",
            command=self.toggle_filename_state,
        )
        self.rb_individual.pack(side="left")

        # --- Output Folder ---
        ttk.Label(output_frame, text="Output Folder:").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_dir_var = tk.StringVar()
        self.entry_output_dir = ttk.Entry(
            output_frame, textvariable=self.output_dir_var, width=50
        )
        self.entry_output_dir.grid(row=1, column=1, sticky="ew", padx=(0, 10))

        self.btn_browse_out = ttk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_dir,
            style="Secondary.TButton",
        )
        self.btn_browse_out.grid(row=1, column=2, sticky="w")

        # --- Output Filename ---
        ttk.Label(output_frame, text="Output Filename:").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_name_var = tk.StringVar(
            value=f"Porsche_AfterSales_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.entry_output_name = ttk.Entry(
            output_frame, textvariable=self.output_name_var, width=50
        )
        self.entry_output_name.grid(row=2, column=1, sticky="ew", padx=(0, 10))

        self.lbl_filename_hint = ttk.Label(
            output_frame, text="(.csv added automatically)", foreground="gray"
        )
        self.lbl_filename_hint.grid(row=2, column=2, sticky="w")

        output_frame.columnconfigure(1, weight=1)

        # --- Data Preview ---
        preview_frame = ttk.LabelFrame(
            content_frame,
            text="Data Preview / Processing Queue",
            padding="15",
        )
        preview_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Invoice No.", "Status", "Items", "Details")
        self.tree = ttk.Treeview(
            preview_frame, columns=cols, show="headings", selectmode="extended"
        )
        self.tree.heading("File Name", text="File Name")
        self.tree.heading("Invoice No.", text="Invoice No.")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Items", text="Items Found")
        self.tree.heading("Details", text="Details")

        self.tree.column("File Name", width=250, anchor="w")
        self.tree.column("Invoice No.", width=120, anchor="center")
        self.tree.column("Status", width=90, anchor="center")
        self.tree.column("Items", width=90, anchor="center")
        self.tree.column("Details", width=500, anchor="w")

        scrollbar_y = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            content_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            foreground="#666666",
            background="#f5f5f5",
            anchor="w",
            padding="5 2",
        )
        status_bar.pack(fill="x", pady=(10, 0))

    # ----- FILE SELECTION -----
    def select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select Porsche After Sales Invoice PDFs",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if files:
            self.selected_files = list(files)
            self.lbl_count.config(
                text=f"{len(self.selected_files)} file(s) selected"
            )
            # Clear and populate treeview
            for row in self.tree.get_children():
                self.tree.delete(row)
            for fpath in self.selected_files:
                self.tree.insert(
                    "", "end",
                    values=(os.path.basename(fpath), "—", "Pending", "", ""),
                )
            self.status_var.set(
                f"{len(self.selected_files)} file(s) loaded. "
                "Click 'Extract & Generate CSV' to process."
            )
            # Auto-set output folder if empty
            if not self.output_dir_var.get():
                first_dir = os.path.dirname(self.selected_files[0])
                self.output_dir_var.set(first_dir)

    def browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def toggle_filename_state(self) -> None:
        """Enable/Disable filename entry based on mode."""
        if self.mode_var.get() == "individual":
            self.entry_output_name.config(state="disabled")
            self.lbl_filename_hint.config(text="(Auto-named by Invoice No.)")
        else:
            self.entry_output_name.config(state="normal")
            self.lbl_filename_hint.config(text="(.csv added automatically)")

    def clear_files(self) -> None:
        """Clear all selected files and reset output path/filename."""
        self.selected_files = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.lbl_count.config(text="No files selected")
        self.output_dir_var.set("")
        self.output_name_var.set(
            f"Porsche_AfterSales_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.status_var.set("File list cleared.")

    # ----- RUN EXTRACTION -----
    def run_extraction(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(
                "No Files", "Please select at least one PDF file."
            )
            return

        # Output setup
        out_dir = self.output_dir_var.get()
        if not out_dir:
            out_dir = os.path.dirname(self.selected_files[0])
            self.output_dir_var.set(out_dir)

        mode = self.mode_var.get()
        combined_records: list[dict] = []
        total_items = 0

        self.btn_run.config(state="disabled")
        self.btn_select.config(state="disabled")
        self.root.update_idletasks()

        tree_rows = self.tree.get_children()

        for idx, fpath in enumerate(self.selected_files):
            fname = os.path.basename(fpath)
            row_id = tree_rows[idx] if idx < len(tree_rows) else None

            try:
                self.status_var.set(f"Processing: {fname} ...")
                self.root.update_idletasks()

                result = extract_porsche_aftersales_invoice(fpath)
                items = result["items"]
                count = len(items)
                inv_no = result.get("invoice_number", "N/A")

                total_items += count

                # --- INDIVIDUAL MODE ---
                if mode == "individual" and items:
                    safe_inv = "".join(
                        c for c in inv_no if c.isalnum() or c in ('-', '_')
                    )
                    if safe_inv:
                        indiv_name = f"{safe_inv}.csv"
                    else:
                        base = os.path.splitext(fname)[0]
                        indiv_name = f"{base}_Extracted.csv"

                    indiv_path = os.path.join(out_dir, indiv_name)
                    write_csv(indiv_path, items)
                    detail_msg = f"Saved: {indiv_name} ({count} items)"

                # --- COMBINED MODE ---
                else:
                    combined_records.extend(items)
                    detail_msg = f"Invoice: {inv_no} | {count} items"

                if row_id:
                    self.tree.item(
                        row_id,
                        values=(
                            fname, inv_no, "✓ Done", str(count), detail_msg
                        ),
                    )

            except Exception as e:
                if row_id:
                    self.tree.item(
                        row_id,
                        values=(fname, "—", "✗ Error", "0", str(e)),
                    )
                self.status_var.set(f"Error processing {fname}: {e}")

            self.root.update_idletasks()

        # Finalize Combined Mode
        if mode == "combined":
            if combined_records:
                out_name = self.output_name_var.get().strip()
                if not out_name:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_name = f"Porsche_AfterSales_Extracted_{timestamp}"
                # Always ensure .csv extension (strip duplicates)
                if out_name.lower().endswith(".csv"):
                    out_name = out_name[:-4]
                out_name += ".csv"

                output_path = os.path.join(out_dir, out_name)
                try:
                    write_csv(output_path, combined_records)
                    messagebox.showinfo(
                        "Success",
                        f"Combined extraction complete!\n\n"
                        f"Total items: {total_items}\n"
                        f"Saved to: {output_path}",
                    )
                    self.status_var.set(f"Done. Saved to {out_name}")
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Could not write combined CSV:\n{e}"
                    )
            else:
                self.status_var.set("No data found to combine.")
                if total_items == 0:
                    messagebox.showwarning(
                        "No Data", "No items extracted from selected files."
                    )

        # Finalize Individual Mode
        else:
            messagebox.showinfo(
                "Success",
                f"Individual extraction complete!\n\n"
                f"Processed {len(self.selected_files)} files.\n"
                f"Total items found: {total_items}\n"
                f"Folder: {out_dir}",
            )
            self.status_var.set(f"Done. Files saved to {out_dir}")

        # Refresh timestamp for next run
        new_ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if self.mode_var.get() == "combined":
            self.output_name_var.set(f"Porsche_AfterSales_Extracted_{new_ts}")

        self._reset_buttons()

    def _reset_buttons(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_select.config(state="normal")

    def run(self) -> None:
        self.root.mainloop()


# ---------- ENTRY POINT ----------
if __name__ == "__main__":
    app = PorscheAfterSalesExtractorGUI()
    app.run()
