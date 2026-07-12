import os
import re
import pdfplumber
from pathlib import Path
from openpyxl import Workbook, load_workbook

BILLS_DIR = Path("/Users/marclivesey/Energy_Bills")
EXCEL_PATH = BILLS_DIR / "Energy_Tracker.xlsx"

# Regex patterns for EDF bill fields
PATTERNS = {
    "balance_last_bill": r"Balance on your last bill\s+£?([\d\.]+)",
    "charges_electricity": r"Charges for Electricity\s+£?([\d\.]+)",
    "charges_gas": r"Charges for Gas\s+£?([\d\.]+)",
    "payments": r"Payments\s+£?([\d\.]+)",
    "new_balance": r"Your new balance\s+£?([\d\.]+)",
    "electricity_used": r"Electricity used\s+([\d\.]+)\s*kWh",
    "electricity_total": r"Total electricity charges for this period\s+£?([\d\.]+)",
    "gas_used": r"Energy used\s+([\d\.]+)\s*kWh",
    "gas_total": r"Total gas charges for this period\s+£?([\d\.]+)"
}

def extract_text_from_pdf(pdf_path):
    """Extract full text from a PDF file."""
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)

def extract_fields(text):
    """Extract all required fields using regex."""
    data = {}
    for key, pattern in PATTERNS.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = float(match.group(1)) if match else None
    return data

def ensure_excel_exists():
    """Create Excel file with headers if it doesn't exist."""
    if not EXCEL_PATH.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Energy Data"
        ws.append([
            "File", "Balance Last Bill", "Electricity Charges", "Gas Charges",
            "Payments", "New Balance", "Electricity Used (kWh)",
            "Electricity Total (£)", "Gas Used (kWh)", "Gas Total (£)"
        ])
        wb.save(EXCEL_PATH)

def append_to_excel(filename, data):
    """Append extracted data to Excel workbook."""
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.append([
        filename,
        data["balance_last_bill"],
        data["charges_electricity"],
        data["charges_gas"],
        data["payments"],
        data["new_balance"],
        data["electricity_used"],
        data["electricity_total"],
        data["gas_used"],
        data["gas_total"]
    ])
    wb.save(EXCEL_PATH)

def process_bills():
    ensure_excel_exists()

    for pdf_file in BILLS_DIR.glob("EDF_*.pdf"):
        print(f"Processing: {pdf_file.name}")
        text = extract_text_from_pdf(pdf_file)
        data = extract_fields(text)
        append_to_excel(pdf_file.name, data)

    print("\n✔ All bills processed and added to Energy_Tracker.xlsx")

if __name__ == "__main__":
    process_bills()
