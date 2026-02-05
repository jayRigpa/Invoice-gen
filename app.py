import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st

from invoice_core import generate_monthly_invoices, GenerateConfig

st.set_page_config(page_title="Timesheet → Monthly Invoices", page_icon="🧾")
st.title("🧾 Timesheet → Monthly Invoices")

st.markdown(
    "Upload your **timesheet** and the **invoice template**. "
    "This app will generate **one invoice per month** and give you a zip to download."
)

timesheet = st.file_uploader("Timesheet (.xls or .xlsx)", type=["xls", "xlsx"])
template = st.file_uploader("Invoice template (.xlsx)", type=["xlsx"])
rate = st.number_input("Hourly rate", min_value=0.0, value=100.0, step=10.0)
use_today = st.checkbox("Use today as the Invoice Date", value=True)

invoice_date = datetime.now() if use_today else None

if st.button("Generate invoices", type="primary", disabled=not (timesheet and template)):
    with st.spinner("Generating..."):
        with tempfile.TemporaryDirectory() as d:
            dpath = Path(d)
            ts_path = dpath / timesheet.name
            tpl_path = dpath / template.name
            ts_path.write_bytes(timesheet.getvalue())
            tpl_path.write_bytes(template.getvalue())

            out_dir = dpath / "out"
            cfg = GenerateConfig(rate_per_hour=float(rate), invoice_date=invoice_date)
            result = generate_monthly_invoices(str(ts_path), str(tpl_path), str(out_dir), cfg)

            zip_bytes = Path(result["zip"]).read_bytes()

    st.success(f"Created {result['count']} invoices.")
    st.download_button(
        "Download Monthly_Invoices.zip",
        data=zip_bytes,
        file_name="Monthly_Invoices.zip",
        mime="application/zip",
    )
