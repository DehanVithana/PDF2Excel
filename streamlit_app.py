# streamlit_app.py
import io
import os
import re
import zipfile
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st


# -----------------------------
# Utility helpers
# -----------------------------
def sanitize_sheet_name(name: str) -> str:
    """
    Excel sheet names must be <= 31 chars and cannot contain: : \ / ? * [ ]
    """
    name = re.sub(r'[:\\/\?\*\[\]]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:31] if name else "Sheet"

def infer_header_row(rows):
    """
    Very light heuristic: if first row is all strings (not numbers) and
    later rows have mixed/numeric content, treat first row as header.
    """
    if not rows or not rows[0]:
        return None
    first = rows[0]
    def is_number(x):
        try:
            float(str(x).replace(",", ""))  # tolerate comma thousands
            return True
        except:
            return False
    first_all_str = all((x is None) or (not is_number(x)) for x in first)
    if not first_all_str:
        return None
    # If most of remaining cells include numbers, it's likely a header
    remaining = [cell for row in rows[1:] for cell in row if cell not in (None, "")]
    numeric_ratio = sum(is_number(x) for x in remaining) / max(1, len(remaining))
    return 0 if numeric_ratio > 0.15 else None

def clean_table_data(rows):
    """
    Remove completely empty rows/cols and coerce to strings where appropriate.
    """
    if not rows:
        return rows
    # Drop empty rows
    rows = [row for row in rows if any((cell not in (None, "")) for cell in row)]
    if not rows:
        return rows
    # Normalize row lengths
    max_len = max(len(r) for r in rows)
    rows = [list(r) + [None] * (max_len - len(r)) for r in rows]
    # Drop empty columns
    cols_to_keep = []
    for c in range(max_len):
        if any((row[c] not in (None, "")) for row in rows):
            cols_to_keep.append(c)
    rows = [[row[c] for c in cols_to_keep] for row in rows]
    return rows

def looks_scanned(page) -> bool:
    """
    Heuristic: no extractable words but images exist => likely scanned.
    """
    try:
        words = page.extract_words() or []
        return (len(words) == 0) and (len(page.images or []) > 0)
    except Exception:
        # If extraction fails, don't falsely mark as scanned.
        return False


# -----------------------------
# Core PDF -> Excel conversion
# -----------------------------
def convert_pdf_to_excel_bytes(pdf_bytes: bytes, display_name: str) -> bytes:
    """
    Returns Excel bytes for a single PDF file.
    - Extracts tables per page; each table becomes a sheet "p{page}_tbl{n}"
    - If no tables found across the entire doc, creates a 'Text' sheet with page+lines
    - Adds 'Info' sheet for OCR recommendations if any scanned pages found
    """
    excel_buf = io.BytesIO()
    ocr_pages = []
    tables_found = 0
    text_rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf, \
         pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:

        for page_index, page in enumerate(pdf.pages, start=1):
            # Table extraction (try modern find_tables first, fall back to extract_tables)
            extracted_tables = []
            try:
                tbls = page.find_tables()
                extracted_tables = [t.extract() for t in tbls] if tbls else []
            except Exception:
                try:
                    extracted_tables = page.extract_tables() or []
                except Exception:
                    extracted_tables = []

            # Clean and write tables
            table_count_this_page = 0
            for raw_rows in extracted_tables:
                rows = clean_table_data(raw_rows)
                if not rows or len(rows) == 0 or (len(rows) == 1 and len(rows[0]) <= 1):
                    continue

                header_idx = infer_header_row(rows)
                if header_idx is not None:
                    header = rows[header_idx]
                    data = rows[header_idx + 1:]
                    df = pd.DataFrame(data, columns=[str(c) if c is not None else "" for c in header])
                else:
                    df = pd.DataFrame(rows)

                sheet = sanitize_sheet_name(f"p{page_index}_tbl{table_count_this_page+1}")
                # Avoid duplicate sheet names
                base = sheet
                suffix = 2
                while sheet in writer.book.sheetnames:
                    sheet = sanitize_sheet_name(f"{base}_{suffix}")
                    suffix += 1

                df.to_excel(writer, index=False, sheet_name=sheet)
                tables_found += 1
                table_count_this_page += 1

            # If no tables on this page, collect text lines
            if table_count_this_page == 0:
                try:
                    txt = page.extract_text() or ""
                except Exception:
                    txt = ""
                if txt.strip():
                    for line in txt.splitlines():
                        if line.strip():
                            text_rows.append({"Page": page_index, "Text": line.strip()})
                if looks_scanned(page):
                    ocr_pages.append(page_index)

        # If no tables in entire document, write text sheet (if any)
        if tables_found == 0:
            if text_rows:
                df_text = pd.DataFrame(text_rows)
            else:
                # completely empty extraction
                df_text = pd.DataFrame([{"Page": None, "Text": "No extractable content found."}])
            df_text.to_excel(writer, index=False, sheet_name="Text")

        # Add OCR info sheet if any pages likely scanned
        if ocr_pages:
            info = pd.DataFrame({
                "Note": ["Some pages appear scanned (image-only). OCR recommended."],
                "Pages": [", ".join(map(str, ocr_pages))]
            })
            info.to_excel(writer, index=False, sheet_name="Info")

    excel_buf.seek(0)
    return excel_buf.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="PDF → Excel Converter", page_icon="📄➡️📊", layout="centered")
st.title("📄 ➡️ 📊 PDF to Excel Converter")
st.caption("Upload one or more PDFs (tables or unstructured). I’ll extract tables when possible, or fall back to page-wise text.")

with st.expander("ℹ️ Notes & Limitations", expanded=False):
    st.markdown(
        """
- Works best on **digitally generated PDFs** (not scanned images).
- If a page looks **scanned**, it will be flagged in an **Info** sheet (no OCR by default).
- Tables appear as sheets `p{page}_tbl{n}`. If no tables are found, a **Text** sheet is created.
- When multiple PDFs are uploaded, you’ll get a **ZIP** containing one Excel per PDF with the **same base name**.
        """
    )

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True,
    help="Drop one or more PDFs here."
)

process_btn = st.button("Convert to Excel", type="primary", disabled=(not uploaded_files))

if process_btn and uploaded_files:
    # Single vs multiple handling
    if len(uploaded_files) == 1:
        f = uploaded_files[0]
        st.info(f"Processing **{f.name}** …")
        try:
            excel_bytes = convert_pdf_to_excel_bytes(f.read(), f.name)
            xlsx_name = os.path.splitext(f.name)[0] + ".xlsx"
            st.success("Done! Download your Excel below.")
            st.download_button(
                label=f"⬇️ Download {xlsx_name}",
                data=excel_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Conversion failed for {f.name}: {e}")

    else:
        st.info(f"Processing **{len(uploaded_files)} PDFs** …")
        zip_buf = io.BytesIO()
        failed = []

        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for f in uploaded_files:
                try:
                    excel_bytes = convert_pdf_to_excel_bytes(f.read(), f.name)
                    xlsx_name = os.path.splitext(f.name)[0] + ".xlsx"
                    zf.writestr(xlsx_name, excel_bytes)
                except Exception as e:
                    failed.append(f"{f.name} ({e})")

        zip_buf.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"converted_excels_{timestamp}.zip"

        if failed:
            st.warning("Some files failed to convert:\n\n- " + "\n- ".join(failed))

        st.success("Batch complete! Download your ZIP below.")
        st.download_button(
            label=f"⬇️ Download {zip_name}",
            data=zip_buf.getvalue(),
            file_name=zip_name,
            mime="application/zip"
        )

st.divider()
st.caption("Built with ❤️ using Streamlit + pdfplumber. For OCR support, see README.")
