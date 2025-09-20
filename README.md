# PDF → Excel Converter (Streamlit)

Convert one or more PDFs to Excel. Extracts tables when possible; falls back to page-wise text when not. If multiple PDFs are uploaded, all Excel outputs are packaged as a ZIP where each Excel matches the original PDF's filename.

## Features
- Multiple PDFs supported
- Table extraction via `pdfplumber`
- Fallback to `Text` sheet if no tables
- Flags scanned/image-only pages (OCR recommended)
- Single PDF → `.xlsx`
- Multiple PDFs → `.zip` (contains one `.xlsx` per PDF)

## Limitations
- OCR (for scanned PDFs) is **not enabled** by default in the cloud deployment (no system packages).
- Table detection is heuristic; results vary by PDF quality and structure.

## Running locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
