import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from typing import Optional, List, Dict

st.set_page_config(page_title="Bank Statement Parser", page_icon="🏦", layout="wide")
st.title("🏦 Parse Bank Statement PDFs → Transactions")
st.write(
    "Upload one or more bank statement PDFs. The app extracts **Posted Date**, **Amount**, "
    "**Transaction Detail**, and tags each line as **Credit** or **Debit**. "
    "Anything after **“Daily ledger balance summary”** is excluded."
)

uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

# --- Regexes ---
TXN_LINE = re.compile(r"^\s*(\d{2}/\d{2})\s+([\d,]+\.\d{2})\s+(.*\S)\s*$")
CREDIT_KEYS = (
    "credits",
    "deposits",
    "electronic deposits",
    "bank credits",
    "electronic deposits/bank credits",
)
DEBIT_KEYS = (
    "debits",
    "electronic debits",
    "bank debits",
    "electronic debits/bank debits",
)
BALANCE_SUMMARY_MARK = "daily ledger balance summary"

def detect_section(line_lower, current_type):
    """Return 'Credit'/'Debit'/current_type based on header phrases in the line."""
    for key in CREDIT_KEYS:
        if key in line_lower:
            return "Credit"
    for key in DEBIT_KEYS:
        if key in line_lower:
            return "Debit"
    return current_type

def parse_pdf(file_like, filename):
    """Parse a single PDF into transaction records."""
    records = []
    with pdfplumber.open(file_like) as pdf:
        current_type = None
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.split("\n")

            # Trim anything after the balance summary marker
            trimmed = []
            cutoff = False
            for line in lines:
                if BALANCE_SUMMARY_MARK in line.lower():
                    cutoff = True
                    break
                trimmed.append(line)
            if cutoff:
                lines = trimmed

            for raw in lines:
                line = (raw or "").strip()
                if not line:
                    continue

                line_lower = line.lower()
                current_type = detect_section(line_lower, current_type)

                m = TXN_LINE.match(line)
                if m and current_type in ("Credit", "Debit"):
                    posted_date = m.group(1)
                    amount = float(m.group(2).replace(",", ""))
                    detail = m.group(3).strip()
                    records.append(
                        {
                            "Source File": filename,
                            "Posted Date": posted_date,
                            "Amount": amount,
                            "Transaction Detail": detail,
                            "Type": current_type,
                        }
                    )
    return records

# --- UI flow ---
if uploaded_files:
    all_rows = []
    for f in uploaded_files:
        try:
            all_rows.extend(parse_pdf(f, f.name))
        except Exception as e:
            st.error(f"❌ Error parsing **{f.name}**: {e}")

    if all_rows:
        df = pd.DataFrame(all_rows)

        credits = df[df["Type"] == "Credit"].reset_index(drop=True)
        debits = df[df["Type"] == "Debit"].reset_index(drop=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Transactions")
            credits.to_excel(writer, index=False, sheet_name="Credits")
            debits.to_excel(writer, index=False, sheet_name="Debits")

        st.success(f"✅ Parsed {len(uploaded_files)} file(s). Extracted {len(df)} transactions.")
        st.download_button(
            label="📥 Download Excel (Transactions, Credits, Debits)",
            data=output.getvalue(),
            file_name=f"bank_statement_export_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Preview (first 50 rows)")
        st.dataframe(df.head(50))
    else:
        st.warning("No transactions were extracted. Please check the PDF format.")
else:
    st.info("Upload PDFs to begin.")

