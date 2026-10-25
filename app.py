import io
import urllib.parse
import pandas as pd
import streamlit as st
import core

st.set_page_config(page_title="Slow-Stock Review", page_icon="📦", layout="centered")
st.title("📦 Slow-Stock Review")
st.caption("Company X · upload the month's files → get the action list. Nothing is stored on the server.")

# ---------- 1. Upload ----------
with st.sidebar:
    st.header("1. Upload")
    mode = st.radio("Files", ["One workbook (4 sheets)", "Four separate files"])
    frames = None
    try:
        if mode.startswith("One"):
            f = st.file_uploader("Workbook (.xlsx)", type=["xlsx"])
            if f:
                frames = core.load_workbook_sheets(f)
        else:
            names = ["Item_Master", "Stock_Aging", "Sales_Lines", "Customers"]
            ups = [st.file_uploader(n, type=["xlsx", "csv"], key=n) for n in names]
            if all(ups):
                frames = [pd.read_csv(u) if u.name.lower().endswith(".csv") else pd.read_excel(u) for u in ups]
    except Exception as e:
        st.error(f"Could not read the files: {e}")
    st.caption("Sheets needed: Item_Master, Stock_Aging, Sales_Lines, Customers")

if frames is None:
    st.info("👈 Upload the workbook in the sidebar (on a phone, tap the **>** at top-left).")
    st.stop()

try:
    items, targets, issues, custs = core.run(*frames)
except Exception as e:
    st.error(f"The files do not look right: {e}. Check the sheet and column names.")
    st.stop()

tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Items", "📞 Call list", "Data checks"])

# ---------- Summary ----------
with tab1:
    c = items["Class"].value_counts()
    push = items[items["Class"] == "Push"]
    a, b, d = st.columns(3)
    a.metric("Push", int(c.get("Push", 0)))
    b.metric("Red List", int(c.get("Red List", 0)))
    d.metric("No Action", int(c.get("No Action", 0)))
    st.metric("Aged value in Push items", f"₹{push['Aged Value (Rs)'].sum():,.0f}")
    alerts = issues[issues["Severity"] == "Alert"]
    for _, r in alerts.iterrows():
        st.warning(r["Problem"])
    st.subheader("Top 10 Push items by aged value")
    top = push.nlargest(10, "Aged Value (Rs)")[["SKU", "Item Name", "Category", "Aged Qty", "Aged Value (Rs)", "Months to Clear"]]
    st.dataframe(top.style.format({"Aged Value (Rs)": "₹{:,.0f}", "Months to Clear": "{:.1f}"}), hide_index=True, width="stretch")
    st.bar_chart(top.set_index("SKU")["Aged Value (Rs)"])
    st.subheader("Red List (sold < 1 master pack in 6 months)")
    st.dataframe(items[items["Class"] == "Red List"][["SKU", "Item Name", "Category", "Total Stock", "Qty 6M", "Master Pack Qty"]], hide_index=True, width="stretch")

# ---------- Items ----------
with tab2:
    f1, f2 = st.columns(2)
    cl = f1.multiselect("Class", ["Push", "Red List", "No Action"], default=["Push", "Red List", "No Action"])
    ca = f2.multiselect("Category", sorted(items["Category"].unique()), default=sorted(items["Category"].unique()))
    view = items[items["Class"].isin(cl) & items["Category"].isin(ca)]
    cols = ["SKU", "Item Name", "Category", "Class", "Total Stock", "Aged Qty", "Aged Value (Rs)", "Qty 6M", "Avg Monthly Sales", "Months to Clear", "Merged From", "Reason"]
    st.dataframe(view[cols].style.format({"Aged Value (Rs)": "{:,.0f}", "Avg Monthly Sales": "{:.1f}", "Months to Clear": "{:.1f}"}), hide_index=True, width="stretch")
    st.caption(f"{len(view)} items")

# ---------- Call list (mobile friendly) ----------
with tab3:
    push = items[items["Class"] == "Push"].sort_values("Aged Value (Rs)", ascending=False)
    pick = st.selectbox("Pick a Push item", push["SKU"], format_func=lambda s: f"{s} · {push.set_index('SKU').loc[s,'Item Name']} · ₹{push.set_index('SKU').loc[s,'Aged Value (Rs)']:,.0f} aged")
    n = st.slider("Show top customers", 3, 15, 10)
    t = targets[targets["SKU"] == pick].head(n)
    row = push.set_index("SKU").loc[pick]
    st.write(f"**{row['Item Name']}** · {int(row['Total Stock'])} in stock · clears in {row['Months to Clear']:.1f} months")
    if t.empty:
        st.info("No eligible customers for this item.")
    for _, r in t.iterrows():
        with st.container(border=True):
            st.markdown(f"**{r['Rank']}. {r['Customer']}**  ·  {r['Industry'] if pd.notna(r['Industry']) else 'Industry n/a'}  ·  {r['Area']}")
            st.caption(f"{r['Bucket Name']}  ·  bought ₹{r['Value 6M']:,.0f} in 6 months")
            msg = (f"Hello {r['Customer']}, this is Company X. We have good stock of {row['Item Name']} "
                   f"available now at a special rate. Shall I send you a quote?")
            with st.expander("WhatsApp message"):
                msg = st.text_area("Edit before sending", msg, key=f"m{pick}{r['CustomerID']}")
                st.link_button("Open in WhatsApp", "https://wa.me/?text=" + urllib.parse.quote(msg))

# ---------- Data checks ----------
with tab4:
    st.write("Problems found in the uploaded data and how each was handled:")
    st.dataframe(issues, hide_index=True, width="stretch")

# ---------- Export ----------
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    items.to_excel(xw, sheet_name="Items", index=False)
    items[items["Class"] == "Push"].nlargest(10, "Aged Value (Rs)").to_excel(xw, sheet_name="Top10_Push", index=False)
    items[items["Class"] == "Red List"].to_excel(xw, sheet_name="Red_List", index=False)
    targets.to_excel(xw, sheet_name="Customer_Targets", index=False)
    issues.to_excel(xw, sheet_name="Data_Issues", index=False)
    for ws in xw.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(45, max(10, max(len(str(c.value or "")) for c in col[:50]) + 2))
st.sidebar.header("2. Export")
st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "slow_stock_review.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
