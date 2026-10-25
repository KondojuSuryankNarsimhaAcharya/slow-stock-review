# Slow-Stock Review (Company X case study)

Upload the month's workbook (Item_Master, Stock_Aging, Sales_Lines, Customers) -> data checks, Push / Red List / No Action
classification, customer call list per Push item (mobile friendly), Excel export.

- `core.py`  all business rules (R1-R7) and data cleaning, no UI. Change rules here.
- `app.py`   Streamlit screen.

Run locally: `pip install -r requirements.txt && streamlit run app.py`
Host free: push to GitHub -> share.streamlit.io -> New app -> pick repo, main file `app.py`.
No keys or passwords are used. Uploaded files are processed in memory only.
