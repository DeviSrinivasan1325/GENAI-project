🧾 AI-Powered GST Invoice Document Processor
An end-to-end AI system that automates GST invoice data extraction, validation, and Excel reporting — eliminating manual entry and reducing filing errors for Indian businesses.

🚀 Overview
Processing GST invoices manually is time-consuming and error-prone. This project solves that by combining OCR, LLMs, and rule-based validation into a single automated pipeline.

👉 Upload an invoice → Extract data → Validate → Detect duplicates → Export to Excel

✨ Key Features
📸 Multi-Format Invoice Support
PDFs (digital + scanned)
Images: JPG, PNG, TIFF, BMP, WEBP
Handles low-quality mobile photos

🤖 Hybrid Extraction Engine
LLM (Google Gemini) extracts:
Vendor Name
Invoice Number
Invoice Date
Regex Rules extract:
Taxable Amount
CGST, SGST, IGST
Total Amount
Rule-based values override AI for higher accuracy

🔍 OCR Error Correction (GSTIN Intelligence)
Fixes common OCR mistakes:
O → 0, Z → 2, I → 1
Ensures GSTIN format compliance

✅ Automated Validation
GSTIN format validation
Tax calculation verification
Date parsing & normalization
Required field checks

🚫 Smart Duplicate Detection
Primary key: Invoice Number + GSTIN
Fuzzy vendor matching handles OCR inconsistencies

📊 Excel Report Generation
Structured invoice register
Color-coded rows based on confidence
Auto-calculated totals
Summary sheet with:
Total CGST / SGST / IGST
Final GST payable

🌐 Scalable Pipeline Design
Modular architecture
Cached LLM model (reduces API overhead)
Retry logic for API rate limits

🧠 Tech Stack
Python
Flask (Web Interface)
Google Gemini 2.5 Flash (LLM Extraction)
Tesseract OCR (Text Extraction)
pdfplumber (PDF Parsing)
Pandas (Data Processing)
OpenPyXL (Excel Generation)

🏗️ System Architecture
Invoice Input (PDF/Image)
        ↓
Text Extraction (OCR + pdfplumber)
        ↓
Rule-Based Extraction (Regex)
        ↓
LLM Extraction (Gemini)
        ↓
Merge Logic (Rules Override AI)
        ↓
Validation Layer
        ↓
Duplicate Detection
        ↓
Excel Output + Summary

⚙️ Installation
1. Clone the repository
git clone https://github.com/DeviSrinivasan1325/gst-invoice-processor.git
cd gst-invoice-processor
2. Install dependencies
pip install -r requirements.txt
3. Install Tesseract OCR

Download and install:
👉 https://github.com/tesseract-ocr/tesseract

Update path in code if needed:
pytesseract.pytesseract.tesseract_cmd = r"YOUR_PATH_TO_TESSERACT"

4. Set up Gemini API Key
api_key = "YOUR_GEMINI_API_KEY"

▶️ Usage
from your_module import process_invoice

result = process_invoice(
    file_path="invoice.pdf",
    api_key="YOUR_GEMINI_API_KEY",
    excel_file="output.xlsx"
)

print(result)

📤 Output
Excel File Includes:
Invoices Sheet
Structured data
Confidence scores
Color-coded validation
Summary Sheet
Total taxable amount
CGST / SGST / IGST
Final payable GST

📊 Confidence Scoring
Each invoice is assigned a confidence score based on:
GSTIN validity
Date parsing success
Tax calculation accuracy
Field completeness

⚠️ Limitations
Highly distorted images may reduce OCR accuracy
Complex invoice layouts may require prompt tuning
Depends on external API (Gemini) availability

🔮 Future Improvements
Bulk invoice processing
UI dashboard with analytics
Database integration
Vendor-wise reporting
API deployment (SaaS model)

💡 Real-World Impact
Reduces 2–3 hours of manual work → seconds per invoice
Minimizes GST filing errors
Improves compliance and audit readiness

🤝 Contributing
Contributions are welcome!
Feel free to open issues or submit pull requests.

📬 Contact
If you're working in GST, FinTech, or AI automation, let's connect!

⭐ If you found this useful
Give this repo a star ⭐ — it helps others discover it!
