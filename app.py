import io
import os
import urllib.parse
import pandas as pd
import streamlit as st
import core

st.set_page_config(page_title="Slow-Stock Review · Company X", page_icon="📦", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1200px;}
.hero {background: linear-gradient(120deg,#0F766E,#115E59 60%,#134E4A); color:#fff; padding:22px 26px; border-radius:14px; margin-bottom:16px;}
.hero h1 {margin:0; font-size:1.7rem; color:#fff;}
.hero p {margin:4px 0 0 0; opacity:.9; font-size:.98rem;}
.kpi {background:#fff; border:1px solid #E2E8F0; border-left:6px solid var(--c); border-radius:12px; padding:14px 16px; box-shadow:0 1px 3px rgba(0,0,0,.05);}
.kpi .l {font-size:.78rem; color:#64748B; text-transform:uppercase; letter-spacing:.04em;}
.kpi .v {font-size:1.8rem; font-weight:700; color:#0F172A; line-height:1.2;}
.kpi .s {font-size:.8rem; color:#64748B;}
.step {background:#F1F5F9; border-radius:12px; padding:16px 18px; height:100%;}
.step b {color:#0F766E;}
.pill {display:inline-block; padding:2px 10px; border-radius:99px; font-size:.78rem; font-weight:600; color:#fff;}
@media (max-width:640px){ .hero h1{font-size:1.3rem;} .kpi .v{font-size:1.4rem;} }
</style>
<div class="hero"><h1>📦 Slow-Stock Review</h1>
<p>Upload the month's system exports → get the slow-stock action list and who to call. No installs, nothing stored on the server.</p></div>
""", unsafe_allow_html=True)


def kpi(col, label, value, sub="", color="#0F766E"):
    col.markdown(f'<div class="kpi" style="--c:{color}"><div class="l">{label}</div><div class="v">{value}</div><div class="s">{sub}</div></div>', unsafe_allow_html=True)


CLASS_COLOR = {"Push": "#EA580C", "Red List": "#DC2626", "No Action": "#16A34A"}

# ---------- Upload: any number of files, any mix ----------
if "data" not in st.session_state:
    st.session_state.data = {}
data = st.session_state.data

with st.sidebar:
    st.header("1 · Upload")
    st.caption("Upload a workbook, or any of the 4 files, one by one or together. Files are recognised by their columns, not their names.")
    ups = st.file_uploader("Excel / CSV files", type=["xlsx", "csv"], accept_multiple_files=True)
    for u in ups or []:
        try:
            found = core.read_any(u)
            if not found:
                st.error(f"{u.name}: no recognisable sheet (expected Item_Master, Stock_Aging, Sales_Lines or Customers columns).")
            data.update(found)
        except Exception as e:
            st.error(f"Could not read {u.name}: {e}")
    b1, b2 = st.columns(2)
    if b1.button("Try sample data", width="stretch"):
        data.update(core.read_any(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data.xlsx")))
    if b2.button("Clear all", width="stretch"):
        data.clear()
        st.rerun()
    st.markdown("**Files found**")
    for n, r in {"Item_Master": "required", "Stock_Aging": "required", "Sales_Lines": "required", "Customers": "optional"}.items():
        st.write(("✅ " if n in data else ("➖ " if r == "optional" else "❌ ")) + n.replace("_", " ") + ("" if n in data else f" ({r})"))

# ---------- What can we do with what was uploaded? ----------
have = set(data)
if not have:
    a, b, c = st.columns(3)
    a.markdown('<div class="step"><b>1 · Upload</b><br>Drop the workbook (or the four sheets) in the sidebar. On a phone tap the <b>&gt;</b> at top-left.</div>', unsafe_allow_html=True)
    b.markdown('<div class="step"><b>2 · Check</b><br>The tool lists every data problem it found and how it handled each one.</div>', unsafe_allow_html=True)
    c.markdown('<div class="step"><b>3 · Act</b><br>See Push / Red List items, then open the call list for who to contact.</div>', unsafe_allow_html=True)
    st.info("Nothing uploaded yet. Click **Try sample data** in the sidebar to see how it works.")
    st.stop()

if "Item_Master" not in have:
    st.error("The **Item Master** is needed for any analysis (pack size, cost, category). Please upload it.")
    st.stop()

if "Stock_Aging" in have and "Sales_Lines" not in have:
    st.warning("Sales lines not uploaded, so items cannot be classified yet. Showing aged stock only. Upload **Sales Lines** for Push / Red List and the call list.")
    sv = core.stock_view(data["Item_Master"], data["Stock_Aging"])
    k1, k2, k3 = st.columns(3)
    kpi(k1, "Active items", len(sv))
    kpi(k2, "Aged qty (91+ days)", f"{int(sv['Aged Qty'].sum()):,}", color="#EA580C")
    kpi(k3, "Aged value", f"₹{sv['Aged Value (Rs)'].sum():,.0f}", "at unit cost", "#DC2626")
    st.subheader("Top 10 items by aged value")
    st.dataframe(sv.head(10)[["SKU", "Item Name", "Category", "Total Stock", "Aged Qty", "Aged Value (Rs)"]], hide_index=True, width="stretch")
    st.bar_chart(sv.groupby("Category")["Aged Value (Rs)"].sum())
    st.stop()

if "Sales_Lines" in have and "Stock_Aging" not in have:
    st.warning("Stock aging not uploaded, so aged value and classes cannot be calculated. Showing sales only. Upload **Stock Aging** for the full review.")
    s = data["Sales_Lines"].copy()
    s["SKU"] = s["SKU"].astype(str).str.strip().str.upper()
    s = s.drop_duplicates()
    k1, k2, k3 = st.columns(3)
    kpi(k1, "Sales lines", f"{len(s):,}")
    kpi(k2, "Customers", s["CustomerID"].nunique())
    kpi(k3, "Sales value", f"₹{s['Value (Rs)'].sum():,.0f}")
    st.subheader("Top 10 customers by value")
    st.dataframe(s.groupby(["CustomerID", "CustomerName"], as_index=False)["Value (Rs)"].sum().nlargest(10, "Value (Rs)"), hide_index=True, width="stretch")
    st.stop()

if not {"Stock_Aging", "Sales_Lines"} <= have:
    st.warning("Upload at least **Stock Aging** or **Sales Lines** together with the Item Master.")
    st.stop()

# ---------- Full review ----------
try:
    items, targets, issues, custs = core.run(data["Item_Master"], data["Stock_Aging"], data["Sales_Lines"], data.get("Customers"))
except Exception as e:
    st.error(f"The files do not look right: {e}. Check the sheet and column names.")
    st.stop()

push = items[items["Class"] == "Push"]
cnt = items["Class"].value_counts()
nfix = int(issues["Severity"].isin(["Fixed", "Warning", "Alert"]).sum())
k = st.columns(4)
kpi(k[0], "Push items", int(cnt.get("Push", 0)), "offer to customers", CLASS_COLOR["Push"])
kpi(k[1], "Aged value in Push", f"₹{push['Aged Value (Rs)'].sum():,.0f}", "91+ days at cost", "#B45309")
kpi(k[2], "Red List", int(cnt.get("Red List", 0)), "sold < 1 pack in 6M", CLASS_COLOR["Red List"])
kpi(k[3], "Data issues found", nfix, "see Data checks tab", "#2563EB")
st.write("")
for _, r in issues[issues["Severity"] == "Alert"].iterrows():
    st.warning(r["Problem"] + " · " + r["How handled"])

tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📋 Items", "📞 Call list", "🔍 Data checks"])

with tab1:
    left, right = st.columns([3, 2])
    top = push.nlargest(10, "Aged Value (Rs)")[["SKU", "Item Name", "Category", "Aged Qty", "Aged Value (Rs)", "Months to Clear"]]
    with left:
        st.subheader("Top 10 Push items by aged value")
        st.dataframe(top.style.format({"Aged Value (Rs)": "₹{:,.0f}", "Months to Clear": "{:.1f}"}), hide_index=True, width="stretch")
    with right:
        st.subheader("Aged value by SKU")
        st.bar_chart(top.set_index("SKU")["Aged Value (Rs)"], color="#EA580C")
    l2, r2 = st.columns(2)
    with l2:
        st.subheader("Items by class")
        st.bar_chart(cnt.reindex(["Push", "Red List", "No Action"]).fillna(0), color="#0F766E")
    with r2:
        st.subheader("Push aged value by category")
        st.bar_chart(push.groupby("Category")["Aged Value (Rs)"].sum(), color="#B45309")
    st.subheader("Red List: sold less than 1 master pack in 6 months")
    st.dataframe(items[items["Class"] == "Red List"][["SKU", "Item Name", "Category", "Total Stock", "Aged Value (Rs)", "Qty 6M", "Master Pack Qty"]], hide_index=True, width="stretch")

with tab2:
    f1, f2 = st.columns(2)
    cats = sorted(items["Category"].unique())
    cl = f1.multiselect("Class", ["Push", "Red List", "No Action"], default=["Push", "Red List", "No Action"])
    ca = f2.multiselect("Category", cats, default=cats)
    view = items[items["Class"].isin(cl) & items["Category"].isin(ca)]
    cols = ["SKU", "Item Name", "Category", "Class", "Total Stock", "Aged Qty", "Aged Value (Rs)", "Qty 6M", "Avg Monthly Sales", "Months to Clear", "Merged From", "Reason"]
    colour = lambda v: f"background-color:{CLASS_COLOR.get(v, '#fff')}22; color:{CLASS_COLOR.get(v, '#000')}; font-weight:600"
    st.dataframe(view[cols].style.map(colour, subset=["Class"]).format({"Aged Value (Rs)": "{:,.0f}", "Avg Monthly Sales": "{:.1f}", "Months to Clear": "{:.1f}"}), hide_index=True, width="stretch", height=480)
    st.caption(f"{len(view)} of {len(items)} items")

with tab3:
    ps = push.sort_values("Aged Value (Rs)", ascending=False).set_index("SKU")
    p1, p2 = st.columns([3, 1])
    pick = p1.selectbox("Pick a Push item", ps.index, format_func=lambda s: f"{s} · {ps.loc[s, 'Item Name']} · ₹{ps.loc[s, 'Aged Value (Rs)']:,.0f} aged")
    n = p2.slider("Customers", 3, 15, 10)
    row = ps.loc[pick]
    st.markdown(f"**{row['Item Name']}** · {int(row['Total Stock'])} in stock · {int(row['Aged Qty'])} aged · clears in {row['Months to Clear']:.1f} months")
    t = targets[targets["SKU"] == pick].head(n)
    if t.empty:
        st.info("No eligible customers for this item.")
    bcol = {1: "#16A34A", 2: "#2563EB", 3: "#7C3AED"}
    for _, r in t.iterrows():
        with st.container(border=True):
            st.markdown(f"**{r['Rank']}. {r['Customer']}** &nbsp; <span class='pill' style='background:{bcol[r['Bucket']]}'>{r['Bucket Name']}</span>", unsafe_allow_html=True)
            st.caption(f"{r['Industry'] if pd.notna(r['Industry']) else 'Industry n/a'} · {r['Area']} · bought ₹{r['Value 6M']:,.0f} in 6 months")
            with st.expander("✉️ WhatsApp message"):
                msg = st.text_area("Edit before sending", f"Hello {r['Customer']}, this is Company X. We have good stock of {row['Item Name']} available now at a special rate. Shall I send you a quote?", key=f"m{pick}{r['CustomerID']}")
                st.link_button("Open in WhatsApp", "https://wa.me/?text=" + urllib.parse.quote(msg))

with tab4:
    st.write("Problems found in the uploaded data and how each was handled:")
    st.dataframe(issues, hide_index=True, width="stretch")

# ---------- Export ----------
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    items.to_excel(xw, sheet_name="Items", index=False)
    push.nlargest(10, "Aged Value (Rs)").to_excel(xw, sheet_name="Top10_Push", index=False)
    items[items["Class"] == "Red List"].to_excel(xw, sheet_name="Red_List", index=False)
    targets.to_excel(xw, sheet_name="Customer_Targets", index=False)
    issues.to_excel(xw, sheet_name="Data_Issues", index=False)
    for ws in xw.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(45, max(10, max(len(str(c.value or "")) for c in col[:50]) + 2))
st.sidebar.header("2 · Export")
st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "slow_stock_review.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
