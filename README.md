# 🚀 Incuxai AI: Autonomous Document-to-Excel Agent

An enterprise-grade autonomous document intelligence pipeline built for the Incuxai AI Agent Development Challenge. It processes up to 100-page complex PDFs and scanned images, automatically detects tables and headers, stitches multi-page datasets, and executes mathematical reconciliation with self-healing validation.

## 🚀 Key Features
* **Zero-Touch Autonomous Layout Profiler:** Automatically switches between Vector Grid Intersection, Spatial Stream Parsing, and Vision OCR.
* **100-Page Batch Engine:** Handles multi-page financial statements, invoices, and structured documents seamlessly.
* **Smart Table Stitching:** Automatically merges tables splitting across consecutive pages while stripping repeated headers.
* **Self-Healing Reconciliation:** Validates vertical column sums against reported totals and flags discrepancies.
* **Structured Excel Export:** Generates clean, styled workbook outputs with auto-fitted columns and zebra striping.

## 🛠️ Tech Stack
* Python
* Streamlit
* pdfplumber & PyMuPDF (fitz)
* pandas & numpy
* Pillow (PIL)
* openpyxl
* google-generativeai (Gemini 2.5 Flash Vision OCR)
* tabulate