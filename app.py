import io
import os
import urllib.parse
import pandas as pd
import streamlit as st
import core

st.set_page_config(page_title="Slow-Stock Review · Company X", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.block-container {padding-top: 3.2rem; max-width: 1180px;}
.hero {background: linear-gradient(120deg,#0F766E,#115E59 60%,#134E4A); color:#fff; padding:22px 26px; border-radius:14px; margin-bottom:16px;}
.hero h1 {margin:0; font-size:1.7rem; color:#fff;}
.hero p {margin:4px 0 0 0; opacity:.9; font-size:.98rem;}
.kpi {background:#fff; border:1px solid #E2E8F0; border-left:6px solid var(--c); border-radius:12px; padding:14px 16px; box-shadow:0 1px 3px rgba(0,0,0,.05);}
.kpi .l {font-size:.78rem; color:#64748B; text-transform:uppercase; letter-spacing:.04em;}
.kpi .v {font-size:1.8rem; font-weight:700; color:#0F172A; line-height:1.2;}
.kpi .s {font-size:.8rem; color:#64748B;}
.step {background:#F1F5F9; border-radius:12px; padding:16px 18px; min-height:128px;}
.step .n {font-size:1.5rem;}
.box {background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 16px;}
.sp {height:14px;}
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

# ---------- Upload (main page; what is shown always matches what is currently uploaded) ----------
ss = st.session_state
ss.setdefault("uk", 0)
ss.setdefault("sample", False)
SAMPLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data.xlsx")
LABEL = {"Item_Master": "Item Master", "Stock_Aging": "Stock Aging", "Sales_Lines": "Sales Lines", "Customers": "Customers"}

up_col, st_col = st.columns([3, 2], gap="large")
with up_col:
    ups = st.file_uploader("Upload the month's files: the workbook, or any of the four sheets (xlsx / csv)", type=["xlsx", "csv"],
                           accept_multiple_files=True, key=f"up{ss.uk}")
data = {}
if ups:
    ss.sample = False
    for u in ups:
        try:
            found = core.read_any(u)
            if not found:
                up_col.error(f"{u.name}: no recognised sheet (needs Item Master, Stock Aging, Sales Lines or Customers columns).")
            data.update(found)
        except Exception as e:
            up_col.error(f"Could not read {u.name}: {e}")
elif ss.sample:
    data.update(core.read_any(SAMPLE))
with st_col:
    st.markdown("**Files recognised**")
    for n, lab in LABEL.items():
        st.markdown(("✅ " if n in data else "⬜ ") + lab)
    b1, b2 = st.columns(2)
    if b1.button("Try sample data", width="stretch", disabled=bool(ups)):
        ss.sample = True
        st.rerun()
    if b2.button("Clear", width="stretch"):
        ss.sample = False
        ss.uk += 1
        st.rerun()
st.markdown('<div class="sp"></div>', unsafe_allow_html=True)

# ---------- Nothing uploaded: guide ----------
if not data:
    a, b, c = st.columns(3, gap="medium")
    a.markdown('<div class="step"><div class="n">📤</div><b>1 · Upload</b><br>Drop the workbook, or any of the four sheets, in the box above.</div>', unsafe_allow_html=True)
    b.markdown('<div class="step"><div class="n">🔍</div><b>2 · Check</b><br>The tool lists every data problem it found and how it handled each one.</div>', unsafe_allow_html=True)
    c.markdown('<div class="step"><div class="n">📞</div><b>3 · Act</b><br>See Push and Red List items, then open the call list to see who to contact.</div>', unsafe_allow_html=True)
    st.markdown('<div class="sp"></div>', unsafe_allow_html=True)
    st.subheader("What each file unlocks")
    st.dataframe(pd.DataFrame({
        "File": [LABEL[k] for k in LABEL],
        "Contains": ["SKU, name, category, pack size, cost, alternate item, status", "Stock per SKU by age bucket (0-30 ... 180+ days)",
                     "6 months of sales per customer and SKU", "Customer ID, name, industry, area"],
        "On its own": ["Catalogue view and master-data checks", "Aged stock view and stock checks", "Top customers and items, sales checks", "Customer mix and checks"],
        "Needed for the full review": ["Yes", "Yes", "Yes", "Optional (built from sales if missing)"]}), hide_index=True, width="stretch")
    st.stop()

have = set(data)
full = {"Item_Master", "Stock_Aging", "Sales_Lines"} <= have


def prep_sales(df):
    s = df.copy()
    s.columns = [str(c).strip() for c in s.columns]
    s["SKU"] = s["SKU"].astype(str).str.strip().str.upper()
    s["MonthDt"] = s["Month"].map(core.parse_month)
    return s.drop_duplicates(subset=["MonthDt", "CustomerID", "SKU", "Qty", "Value (Rs)"])


def export_button(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nm, d in sheets.items():
            d.to_excel(xw, sheet_name=nm, index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(45, max(10, max(len(str(c.value or "")) for c in col[:50]) + 2))
    st.download_button("⬇️ Download Excel", buf.getvalue(), "slow_stock_review.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============ FULL REVIEW ============
if full:
    try:
        items, targets, issues, custs = core.run(data["Item_Master"], data["Stock_Aging"], data["Sales_Lines"], data.get("Customers"))
    except Exception as e:
        st.error(f"The files do not look right: {e}. Check the sheet and column names.")
        st.stop()
    push = items[items["Class"] == "Push"]
    cnt = items["Class"].value_counts()
    nfix = int(issues["Severity"].isin(["Fixed", "Warning", "Alert"]).sum())
    k = st.columns(4, gap="medium")
    kpi(k[0], "Push items", int(cnt.get("Push", 0)), "offer to customers", CLASS_COLOR["Push"])
    kpi(k[1], "Aged value in Push", f"₹{push['Aged Value (Rs)'].sum():,.0f}", "91+ days at cost", "#B45309")
    kpi(k[2], "Red List", int(cnt.get("Red List", 0)), "sold < 1 pack in 6M", CLASS_COLOR["Red List"])
    kpi(k[3], "Data issues found", nfix, "see Data checks tab", "#2563EB")
    st.markdown('<div class="sp"></div>', unsafe_allow_html=True)
    for _, r in issues[issues["Severity"] == "Alert"].iterrows():
        st.warning(r["Problem"] + " · " + r["How handled"])
    export_button({"Items": items, "Top10_Push": push.nlargest(10, "Aged Value (Rs)"), "Red_List": items[items["Class"] == "Red List"],
                   "Customer_Targets": targets, "Data_Issues": issues})
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📋 Items", "📞 Call list", "🔍 Data checks"])
    with tab1:
        top = push.nlargest(10, "Aged Value (Rs)")[["SKU", "Item Name", "Category", "Aged Qty", "Aged Value (Rs)", "Months to Clear"]]
        left, right = st.columns([3, 2], gap="large")
        with left:
            st.subheader("Top 10 Push items by aged value")
            st.dataframe(top.style.format({"Aged Value (Rs)": "₹{:,.0f}", "Months to Clear": "{:.1f}"}), hide_index=True, width="stretch")
        with right:
            st.subheader("Aged value by SKU")
            st.bar_chart(top.set_index("SKU")["Aged Value (Rs)"], color="#EA580C")
        l2, r2 = st.columns(2, gap="large")
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
    st.stop()

# ============ PARTIAL: any one, two or three files ============
missing = [LABEL[k] for k in ("Item_Master", "Stock_Aging", "Sales_Lines") if k not in have]
st.info("**Partial view.** The full review (Push / Red List and the call list) also needs: **" + ", ".join(missing) + "**. Here is what your uploaded file(s) already show.")
checks = pd.DataFrame([r for k in have for r in core.file_checks(k, data[k])], columns=["Severity", "Sheet", "Problem", "How handled"])
k = st.columns(4, gap="medium")
ki = 0
im = sv = sa = None
if "Item_Master" in have:
    im = data["Item_Master"].copy()
    im.columns = [str(c).strip() for c in im.columns]
    act = im[im["Status"].astype(str).str.strip().str.title() == "Active"]
    kpi(k[ki], "Active items", len(act), f"{im['Category'].nunique()} categories", "#0F766E"); ki += 1
if "Stock_Aging" in have:
    sv = core.stock_view(data["Item_Master"], data["Stock_Aging"]) if "Item_Master" in have else None
    st0 = data["Stock_Aging"].copy()
    st0.columns = [str(c).strip() for c in st0.columns]
    tot, aged = st0[core.BUCKETS].clip(lower=0).sum().sum(), st0[core.AGED].clip(lower=0).sum().sum()
    kpi(k[ki], "Aged stock (91+ days)", f"{int(aged):,} units", f"{aged / tot:.0%} of {int(tot):,} units", "#EA580C"); ki += 1
    if sv is not None:
        kpi(k[ki], "Aged value", f"₹{sv['Aged Value (Rs)'].sum():,.0f}", "at unit cost", "#DC2626"); ki += 1
if "Sales_Lines" in have:
    sa = prep_sales(data["Sales_Lines"])
    cv = sa.groupby("CustomerID")["Value (Rs)"].sum()
    kpi(k[ki], "Sales value (6M)", f"₹{sa['Value (Rs)'].sum():,.0f}", f"{sa['CustomerID'].nunique()} customers", "#2563EB"); ki += 1
    if ki < 4:
        kpi(k[ki], "Customers over ₹1 L", int((cv > core.THRESHOLD).sum()), "eligible for targeting", "#7C3AED"); ki += 1
if "Customers" in have and ki < 4:
    cu0 = data["Customers"]
    kpi(k[ki], "Customers", len(cu0), f"{cu0['Industry'].nunique()} industries", "#0F766E"); ki += 1
st.markdown('<div class="sp"></div>', unsafe_allow_html=True)
sheets = {k2: data[k2] for k2 in have}
sheets["Data_Issues"] = checks
export_button(sheets)

names = [("Item_Master", "📘 Item Master"), ("Stock_Aging", "📦 Stock Aging"), ("Sales_Lines", "🧾 Sales Lines"), ("Customers", "👥 Customers")]
shown = [(n, t) for n, t in names if n in have]
tabs = st.tabs([t for _, t in shown] + ["🔍 Data checks"])
for (n, _), tb in zip(shown, tabs):
    with tb:
        if n == "Item_Master":
            c1, c2 = st.columns([3, 2], gap="large")
            c1.dataframe(im, hide_index=True, width="stretch", height=380)
            c2.subheader("Items by category")
            c2.bar_chart(im.groupby("Category").size(), color="#0F766E")
            sw = im[im.get("Alternate Flag", pd.Series(index=im.index, dtype=object)).astype(str).str.upper() == "SWITCH"]
            if len(sw):
                st.subheader("Pack-size switches (will be merged)")
                st.dataframe(sw[["SKU", "Item Name", "Alternate Item (name)"]], hide_index=True, width="stretch")
        elif n == "Stock_Aging":
            c1, c2 = st.columns([3, 2], gap="large")
            c2.subheader("Units by age")
            c2.bar_chart(st0[core.BUCKETS].clip(lower=0).sum().rename(lambda x: x.replace("Qty ", "")), color="#EA580C")
            st0["Aged Qty"] = st0[core.AGED].clip(lower=0).sum(axis=1)
            if sv is not None:
                c1.subheader("Top 10 by aged value")
                c1.dataframe(sv.head(10)[["SKU", "Item Name", "Category", "Total Stock", "Aged Qty", "Aged Value (Rs)"]], hide_index=True, width="stretch")
            else:
                c1.subheader("Top 10 by aged quantity")
                c1.dataframe(st0.nlargest(10, "Aged Qty")[["SKU", "Aged Qty"]], hide_index=True, width="stretch")
                c1.caption("Upload the Item Master to convert aged units into rupees.")
        elif n == "Sales_Lines":
            c1, c2 = st.columns(2, gap="large")
            c1.subheader("Top 10 customers by value")
            c1.dataframe(sa.groupby(["CustomerID", "CustomerName"], as_index=False)["Value (Rs)"].sum().nlargest(10, "Value (Rs)"), hide_index=True, width="stretch")
            c2.subheader("Top 10 SKUs by quantity")
            c2.dataframe(sa.groupby("SKU", as_index=False)["Qty"].sum().nlargest(10, "Qty"), hide_index=True, width="stretch")
            st.subheader("Sales value by month")
            st.bar_chart(sa.dropna(subset=["MonthDt"]).groupby("MonthDt")["Value (Rs)"].sum(), color="#2563EB")
        else:
            c1, c2 = st.columns([3, 2], gap="large")
            c1.dataframe(data["Customers"], hide_index=True, width="stretch", height=380)
            c2.subheader("Customers by industry")
            c2.bar_chart(data["Customers"].fillna({"Industry": "(missing)"}).groupby("Industry").size(), color="#0F766E")
with tabs[-1]:
    st.write("Problems found in the uploaded file(s) and how each is handled:")
    if checks.empty:
        st.success("No problems found in the uploaded file(s).")
    else:
        st.dataframe(checks, hide_index=True, width="stretch")
