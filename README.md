# Timesheet → Monthly Invoices (Cloud App)

This is a small Streamlit web app that:
- takes a timesheet `.xls/.xlsx` (Timesheet1-style format where headers are on row 2)
- takes an invoice template `.xlsx` (SampleInvoice-style)
- generates **one invoice per month**
- returns a `Monthly_Invoices.zip`

## Run locally
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to a URL (Streamlit Community Cloud)
1. Create a GitHub repo and upload these files:
   - `app.py`
   - `invoice_core.py`
   - `requirements.txt`
2. Go to Streamlit Community Cloud and create a new app from your repo.
3. Set the entry point to `app.py`.

## Deploy to Render (alternative)
- Create a new **Web Service** from your repo.
- Build command: `pip install -r requirements.txt`
- Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
