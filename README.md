# Slow-Stock Review (Company X case study)

Upload the month's workbook (Item_Master, Stock_Aging, Sales_Lines, Customers) -> data checks, Push / Red List / No Action
classification, customer call list per Push item (mobile friendly), Excel export.

- `core.py`  all business rules (R1-R7) and data cleaning, no UI. Change rules here.
- `app.py`   Streamlit screen.

Run locally: `pip install -r requirements.txt && streamlit run app.py`
Host free: push to GitHub -> share.streamlit.io -> New app -> pick repo, main file `app.py`.
No keys or passwords are used. Uploaded files are processed in memory only.

## Uploading
Upload one workbook, or any of the four files (xlsx/csv), together or one by one. Sheets are recognised by their
columns, not their file names. Missing files give a partial view and a clear message. `sample_data.xlsx` powers the
"Try sample data" button.
