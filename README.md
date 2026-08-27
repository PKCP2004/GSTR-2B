# Pushpak Kumar — GSTR-2B Consolidator

Professional GST Automation Toolkit by Pushpak Kumar.

Website: https://pushpakkumar.com

## Excel branding
The generated workbook includes a branded `Read me` landing page with:
- PKP logo
- Pushpak Kumar branding
- clickable pushpakkumar.com
- source file count
- period count
- transaction count
- consolidation method
- processing status
- summary/transaction calculation method

The actual GSTR-2B data sheets remain focused on the original portal structure.

## Calculation
Summary sheets are totalled into the same original portal cells. Transaction sheets are appended row-by-row with GSTR-2B Month.

## Run
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```
