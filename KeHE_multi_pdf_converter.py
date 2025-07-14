import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime

st.title("🧾 Convert Multiple KeHE Chargeback PDFs to Excel")
st.write("Upload one or more KEHE invoice PDFs. The app extracts item/store details and adds a summary of Total Payable and Total Fee.")

uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_records = []
    summary_rows = []

    # Compile regex patterns
    store_pattern = re.compile(r"SOLD TO:\s+(.*)")
    store_id_city_pattern = re.compile(r"(\d+)\s+([A-Z\s\-']+)\s+([A-Z]{2})\s+(\d{5})")
    item_pattern = re.compile(
        r"(\d{12})\s+(\d+)\s+(.*?)\s+(\d+)\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+[\w/]+\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)"
    )
    total_payable_pattern = re.compile(r"TOTAL PAYABLE\s+\$?([\d,]+\.\d{2})")
    total_fee_pattern = re.compile(r"TOTAL FEE\s+\$?([\d,]+\.\d{2})")

    for uploaded_file in uploaded_files:
        records = []
        current_store = {}
        total_payable = None
        total_fee = None
        file_name = uploaded_file.name

        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                lines = page.extract_text().split('\n')

                if page_num == 0:
                    for line in lines:
                        if total_payable is None:
                            match_payable = total_payable_pattern.search(line)
                            if match_payable:
                                total_payable = float(match_payable.group(1).replace(",", ""))

                        if total_fee is None:
                            match_fee = total_fee_pattern.search(line)
                            if match_fee:
                                total_fee = float(match_fee.group(1).replace(",", ""))

                for i, line in enumerate(lines):
                    if "SOLD TO:" in line:
                        store_name = store_pattern.search(line).group(1).strip()
                        address = lines[i + 1].strip()
                        city_line = lines[i + 2].strip()
                        city_match = store_id_city_pattern.search(city_line)

                        if city_match:
                            store_id = city_match.group(1)
                            city = city_match.group(2).strip()
                            state = city_match.group(3)
                            zip_code = city_match.group(4)
                        else:
                            store_id = city = state = zip_code = ""

                        current_store = {
                            "Sold To": store_name,
                            "Address": address,
                            "Store ID": store_id,
                            "City": city,
                            "State": state,
                            "Zip": zip_code
                        }

                    item_match = item_pattern.match(line)
                    if item_match and current_store:
                        record = {
                            "Source File": file_name,
                            "Sold To": current_store.get("Sold To", ""),
                            "Address": current_store.get("Address", ""),
                            "Store ID": current_store.get("Store ID", ""),
                            "City": current_store.get("City", ""),
                            "State": current_store.get("State", ""),
                            "Zip": current_store.get("Zip", ""),
                            "UPC": item_match.group(1),
                            "Qty": int(item_match.group(2)),
                            "Description": item_match.group(3).strip(),
                            "Reference": item_match.group(4),
                            "Date": item_match.group(5),
                            "Cost": float(item_match.group(6)),
                            "Discount": float(item_match.group(7)),
                            "Extended Cost": float(item_match.group(8)),
                        }
                        records.append(record)

        # Add row to summary
        summary_rows.append({
            "File Name": file_name,
            "Total Payable": total_payable,
            "Total Fee": total_fee,
            "Total Ext Cost": total_payable - total_fee
        })

        all_records.extend(records)

    if all_records:
        df_items = pd.DataFrame(all_records)
        df_summary = pd.DataFrame(summary_rows)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_items.to_excel(writer, index=False, sheet_name="Parsed Invoices")
            df_summary.to_excel(writer, index=False, sheet_name="Invoice Summary")

        st.success(f"✅ Parsed {len(uploaded_files)} file(s) successfully!")

        st.download_button(
            label="📥 Download Excel File (with summary)",
            data=output.getvalue(),
            file_name=f"kehe_invoice_export_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(df_summary)
    else:
        st.warning("No records were extracted. Please check the format of your PDF(s).")
