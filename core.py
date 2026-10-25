"""Slow-stock review engine for Company X. Pure pandas, no UI. Rules R1-R7 from the case."""
import re
import pandas as pd

BUCKETS = ["Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days", "Qty 91-180 days", "Qty 180+ days"]
AGED = ["Qty 91-180 days", "Qty 180+ days"]
THRESHOLD = 100000  # R6: customer must have bought MORE than this over 6 months
MONTHS = 6


def norm(s):
    """Normalise a name for matching: collapse spaces, lower-case."""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def parse_month(v):
    """Accepts 'Apr-2026', '2026-06', or a real date (Excel often turns 'Jun-2026' into a date)."""
    if pd.isna(v):
        return pd.NaT
    if hasattr(v, "year"):
        return pd.Timestamp(v.year, v.month, 1)
    s = str(v).strip()
    for fmt in ("%b-%Y", "%Y-%m", "%Y-%m-%d", "%B %Y", "%b %Y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except Exception:
            pass
    return pd.NaT


def clean(items, stock, sales, custs):
    """Return cleaned frames and a list of issues: (severity, sheet, problem, how handled)."""
    log = []
    add = lambda sev, sh, p, h: log.append({"Severity": sev, "Sheet": sh, "Problem": p, "How handled": h})
    for df in (items, stock, sales, custs):
        df.columns = [str(c).strip() for c in df.columns]
        for c in df.select_dtypes(include="object").columns.union(df.select_dtypes(include="string").columns):
            df[c] = df[c].where(df[c].isna(), df[c].astype(str).str.strip())

    # SKU case/space
    for name, df in (("Item_Master", items), ("Stock_Aging", stock), ("Sales_Lines", sales)):
        bad = (df["SKU"] != df["SKU"].str.upper()).sum()
        if bad:
            add("Fixed", name, f"{bad} SKU values in lower case and/or with trailing spaces (e.g. 'sku1004 ') - would not match the master", "Trimmed spaces and converted to upper case")
        df["SKU"] = df["SKU"].str.upper()

    # Dates
    sales["MonthDt"] = sales["Month"].map(parse_month)
    badm = sales["MonthDt"].isna().sum()
    fmt_odd = sales["Month"].astype(str).str.match(r"^\d{4}-\d{2}").sum()
    if fmt_odd:
        add("Fixed", "Sales_Lines", f"{fmt_odd} rows have month as '2026-06' while the rest use 'Apr-2026' (Excel converted a text month)", "Parsed both formats into one date")
    if badm:
        add("Warning", "Sales_Lines", f"{badm} rows have an unreadable month", "Rows kept, month shown blank")
    nmonths = sales["MonthDt"].nunique()
    if nmonths != MONTHS:
        add("Warning", "Sales_Lines", f"Sales cover {nmonths} months, not {MONTHS}", f"Averages still divide by {MONTHS} as per R4 - check you uploaded the full period")

    # Duplicates
    dup = sales.duplicated(subset=["MonthDt", "CustomerID", "SKU", "Qty", "Value (Rs)"], keep="first")
    if dup.any():
        add("Fixed", "Sales_Lines", f"{dup.sum()} exact duplicate sales lines (same month, customer, SKU, qty, value), worth Rs {sales.loc[dup,'Value (Rs)'].sum():,.0f}",
            "Removed the repeat rows. Ask data owner to confirm they are not genuine repeat orders")
        sales = sales[~dup].copy()

    # Unknown SKUs / customers
    for name, df, key in (("Sales_Lines", sales, "SKU"), ("Stock_Aging", stock, "SKU")):
        miss = set(df[key]) - set(items["SKU"])
        if miss:
            add("Warning", name, f"{len(miss)} SKUs not in Item_Master: {sorted(miss)}", "Ignored (cannot classify without cost/pack)")
    miss = set(items["SKU"]) - set(stock["SKU"])
    if miss:
        add("Warning", "Stock_Aging", f"{len(miss)} master items have no stock row: {sorted(miss)}", "Treated as zero stock")
    missc = set(sales["CustomerID"]) - set(custs["CustomerID"])
    if missc:
        add("Warning", "Sales_Lines", f"Customers not in Customers sheet: {sorted(missc)}", "Industry taken from the sales line")

    # Duplicate SKUs in master / stock
    for name, df in (("Item_Master", items), ("Stock_Aging", stock)):
        d = df["SKU"].duplicated().sum()
        if d:
            add("Fixed", name, f"{d} duplicate SKU rows", "First row kept" if name == "Item_Master" else "Quantities summed")
    items = items.drop_duplicates("SKU").copy()
    stock = stock.groupby("SKU", as_index=False)[BUCKETS].sum()

    # Negative stock
    for b in BUCKETS:
        neg = stock[stock[b] < 0]
        for _, r in neg.iterrows():
            add("Fixed", "Stock_Aging", f"{r['SKU']} has negative stock ({int(r[b])}) in '{b}'", "Set to 0 (a negative bucket is a posting error). Ask warehouse to correct at source")
        stock[b] = stock[b].clip(lower=0)

    # Sales sanity
    badq = (sales["Qty"] <= 0).sum()
    if badq:
        add("Fixed", "Sales_Lines", f"{badq} lines with zero/negative quantity", "Removed (returns should be handled separately)")
        sales = sales[sales["Qty"] > 0].copy()

    # Master data: flags
    items["Alternate Flag"] = items["Alternate Flag"].fillna("").astype(str).str.upper().str.strip()
    items["Status"] = items["Status"].astype(str).str.strip().str.title()
    nf = items[(items["Alternate Item (name)"].notna()) & (items["Alternate Flag"] == "")]
    for _, r in nf.iterrows():
        add("Warning", "Item_Master", f"{r['SKU']} names an alternate item ('{norm(r['Alternate Item (name)'])}') but has no flag", "Not merged (R1 only applies to SWITCH). Ask product team whether this is a pack-size switch")
    ex = items[~items["Alternate Flag"].isin(["", "SWITCH"])]
    for _, r in ex.iterrows():
        add("Warning", "Item_Master", f"{r['SKU']} has flag '{r['Alternate Flag']}', which the rules do not define", "Not merged; analysed as a normal item. Ask supply chain what EXCLUDE should do")
    dn = items["Item Name"].str.contains(r"\s{2,}").sum() + items["Alternate Item (name)"].dropna().str.contains(r"\s{2,}").sum()
    if dn:
        add("Fixed", "Item_Master", f"{dn} item names contain double spaces (alternate item is stored by name, so matching breaks)", "Names normalised before matching")
    if custs["Industry"].isna().any():
        ids = list(custs.loc[custs["Industry"].isna(), "CustomerID"])
        add("Warning", "Customers", f"Industry missing for {ids}", "Kept; excluded from 'similar business' bucket (R6)")
    # Customer master vs sales industry mismatch
    m = sales.merge(custs[["CustomerID", "Industry"]], on="CustomerID", suffixes=("", "_m"))
    mm = m[(m["Industry"] != m["Industry_m"]) & m["Industry_m"].notna()]
    if len(mm):
        add("Warning", "Sales_Lines", f"{mm['CustomerID'].nunique()} customers have a different industry in sales vs Customers sheet", "Customers sheet used as the master")
    return items, stock, sales, custs, log


def run(items, stock, sales, custs=None):
    derived = custs is None
    if derived:  # Customers sheet not uploaded: build it from the sales lines
        custs = sales.copy()
        custs["CustomerID"] = custs["CustomerID"].astype(str).str.strip()
        custs = custs.drop_duplicates("CustomerID")[["CustomerID", "CustomerName", "Industry"]].assign(Area="-")
    items, stock, sales, custs = [d.copy() for d in (items, stock, sales, custs)]
    items, stock, sales, custs, log = clean(items, stock, sales, custs)
    add = lambda sev, sh, p, h: log.append({"Severity": sev, "Sheet": sh, "Problem": p, "How handled": h})
    if derived:
        add("Warning", "Customers", "Customers sheet not uploaded", "Customer name and industry taken from the sales lines; area is not available")
    # R1 map old SKU -> alternate SKU
    byname = {norm(n): s for s, n in zip(items["SKU"], items["Item Name"])}
    status = dict(zip(items["SKU"], items["Status"]))
    remap = {}
    for _, r in items[items["Alternate Flag"] == "SWITCH"].iterrows():
        new = byname.get(norm(r["Alternate Item (name)"]))
        if new is None:
            add("Warning", "Item_Master", f"{r['SKU']} is SWITCH but alternate '{r['Alternate Item (name)']}' not found", "Not merged")
        elif status.get(new) != "Active":
            add("Warning", "Item_Master", f"{r['SKU']} switches to inactive {new}", "Not merged")
        else:
            remap[r["SKU"]] = new
    if remap:
        add("Info", "Item_Master", f"{len(remap)} pack-size switches merged: " + ", ".join(f"{a}->{b}" for a, b in remap.items()),
            "Stock and sales of old code added to the new code; new code's pack and cost used (R1)")
    # R2 inactive
    inact = items[items["Status"] != "Active"]["SKU"]
    for s in inact:
        row = stock[stock["SKU"] == s]
        q = int(row[BUCKETS].sum(axis=1).iloc[0]) if len(row) else 0
        cost = items.loc[items["SKU"] == s, "Unit Cost (Rs)"].iloc[0]
        if q > 0:
            add("Alert", "Item_Master", f"Inactive item {s} still holds {q} units (Rs {q*cost:,.0f} at cost)", "Left out of the analysis (R2). Ask supply chain to decide: sell-off, return to vendor or write-off")
    sales["Item"] = sales["SKU"].replace(remap)
    stock["Item"] = stock["SKU"].replace(remap)
    # item table
    base = items[(items["Status"] == "Active") & (~items["SKU"].isin(remap))].set_index("SKU")
    st = stock.groupby("Item")[BUCKETS].sum()
    sq = sales.groupby("Item")["Qty"].sum()
    sv = sales.groupby("Item")["Value (Rs)"].sum()
    t = base[["Item Name", "Category", "Master Pack Qty", "Unit Cost (Rs)"]].join(st).join(sq.rename("Qty 6M")).join(sv.rename("Sales Value 6M"))
    t[BUCKETS + ["Qty 6M", "Sales Value 6M"]] = t[BUCKETS + ["Qty 6M", "Sales Value 6M"]].fillna(0)
    t["Merged From"] = pd.Series({v: k for k, v in remap.items()}).reindex(t.index).fillna("")
    t["Total Stock"] = t[BUCKETS].sum(axis=1)
    t["Aged Qty"] = t[AGED].sum(axis=1)
    t["Aged Value (Rs)"] = t["Aged Qty"] * t["Unit Cost (Rs)"]
    t["Avg Monthly Sales"] = t["Qty 6M"] / MONTHS
    t["Months to Clear"] = (t["Total Stock"] / t["Avg Monthly Sales"]).where(t["Avg Monthly Sales"] > 0)

    def cls(r):
        if r["Qty 6M"] < r["Master Pack Qty"]:
            return "Red List"
        if r["Aged Qty"] == 0 or r["Months to Clear"] < 3:
            return "No Action"
        return "Push"
    t["Class"] = t.apply(cls, axis=1)

    def why(r):
        if r["Class"] == "Red List":
            return f"Sold {int(r['Qty 6M'])} units in 6M, less than 1 pack ({int(r['Master Pack Qty'])})"
        if r["Aged Qty"] == 0:
            return "No aged stock"
        if r["Months to Clear"] < 3:
            return f"Has aged stock but clears in {r['Months to Clear']:.1f} months (<3)"
        return f"Aged stock {int(r['Aged Qty'])} units; {r['Months to Clear']:.1f} months to clear (>=3)"
    t["Reason"] = t.apply(why, axis=1)
    t = t.reset_index().rename(columns={"index": "SKU"})
    t = t.sort_values("Aged Value (Rs)", ascending=False)

    # customers & targets (R6, R7)
    ctot = sales.groupby("CustomerID")["Value (Rs)"].sum().rename("Value 6M")
    cu = custs.set_index("CustomerID")[["CustomerName", "Industry", "Area"]].join(ctot).fillna({"Value 6M": 0})
    eligible = cu[cu["Value 6M"] > THRESHOLD]
    cat = dict(zip(t["SKU"], t["Category"]))
    cat.update(dict(zip(items["SKU"], items["Category"])))
    sales["Cat"] = sales["Item"].map(cat)
    rows = []
    for _, it in t[t["Class"] == "Push"].iterrows():
        sk = it["SKU"]
        buyers = set(sales.loc[sales["Item"] == sk, "CustomerID"])
        inds = {cu.loc[c, "Industry"] for c in buyers if c in cu.index and pd.notna(cu.loc[c, "Industry"])}
        catbuy = set(sales.loc[(sales["Cat"] == it["Category"]) & (sales["Item"] != sk), "CustomerID"])
        for cid, c in eligible.iterrows():
            if cid in buyers:
                b = 1
            elif pd.notna(c["Industry"]) and c["Industry"] in inds:
                b = 2
            elif cid in catbuy:
                b = 3
            else:
                continue
            rows.append({"SKU": sk, "Item Name": it["Item Name"], "Bucket": b,
                         "Bucket Name": {1: "1 - Bought before", 2: "2 - Similar business", 3: "3 - Same category"}[b],
                         "CustomerID": cid, "Customer": c["CustomerName"], "Industry": c["Industry"], "Area": c["Area"],
                         "Value 6M": c["Value 6M"]})
    tg = pd.DataFrame(rows, columns=["SKU", "Item Name", "Bucket", "Bucket Name", "CustomerID", "Customer", "Industry", "Area", "Value 6M"])
    tg = tg.sort_values(["SKU", "Bucket", "Value 6M"], ascending=[True, True, False]).reset_index(drop=True)
    tg["Rank"] = tg.groupby("SKU").cumcount() + 1
    return t, tg, pd.DataFrame(log), cu


def load_workbook_sheets(src):
    x = pd.read_excel(src, sheet_name=None)
    need = ["Item_Master", "Stock_Aging", "Sales_Lines", "Customers"]
    miss = [n for n in need if n not in x]
    if miss:
        raise ValueError(f"Missing sheets: {miss}")
    return [x[n] for n in need]


SHEETS = ["Item_Master", "Stock_Aging", "Sales_Lines", "Customers"]


def detect(df):
    """Identify which of the 4 sheets a table is, from its column names (not the file name)."""
    cols = {norm(c) for c in df.columns}
    if {"master pack qty", "unit cost (rs)"} <= cols:
        return "Item_Master"
    if any(c.startswith("qty 0-30") for c in cols):
        return "Stock_Aging"
    if {"month", "sku", "qty", "value (rs)"} <= cols:
        return "Sales_Lines"
    if {"customerid", "area"} <= cols and "sku" not in cols:
        return "Customers"
    return None


def read_any(f):
    """Read an uploaded .xlsx (every sheet) or .csv and return {sheet_name: DataFrame} for the sheets recognised."""
    name = f if isinstance(f, str) else f.name  # a path (sample data) or an uploaded file
    d = {name: pd.read_csv(f)} if name.lower().endswith(".csv") else pd.read_excel(f, sheet_name=None)
    out = {}
    for df in d.values():
        t = detect(df)
        if t:
            out[t] = df
    return out


def stock_view(items, stock):
    """Partial result when no sales were uploaded: aged stock only (no classification possible)."""
    items, stock = items.copy(), stock.copy()
    for d in (items, stock):
        d.columns = [str(c).strip() for c in d.columns]
        d["SKU"] = d["SKU"].astype(str).str.strip().str.upper()
    stock = stock.groupby("SKU", as_index=False)[BUCKETS].sum()
    stock[BUCKETS] = stock[BUCKETS].clip(lower=0)
    t = items.merge(stock, on="SKU", how="left").fillna({b: 0 for b in BUCKETS})
    t = t[t["Status"].astype(str).str.strip().str.title() == "Active"].copy()
    t["Total Stock"] = t[BUCKETS].sum(axis=1)
    t["Aged Qty"] = t[AGED].sum(axis=1)
    t["Aged Value (Rs)"] = t["Aged Qty"] * t["Unit Cost (Rs)"]
    return t.sort_values("Aged Value (Rs)", ascending=False)


def file_checks(name, df):
    """Data checks that need only ONE sheet, so a single uploaded file still gets a data-quality report."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    out = []
    add = lambda sev, p, h: out.append({"Severity": sev, "Sheet": name, "Problem": p, "How handled": h})
    if "SKU" in df.columns:
        sku = df["SKU"].astype(str)
        bad = int((sku != sku.str.strip().str.upper()).sum())
        if bad:
            add("Fixed", f"{bad} SKU values in lower case and/or with trailing spaces", "Trimmed and upper-cased")
    if name == "Item_Master":
        d = int(df["SKU"].astype(str).str.strip().str.upper().duplicated().sum())
        if d:
            add("Fixed", f"{d} duplicate SKU rows", "First row kept")
        flag = df.get("Alternate Flag", pd.Series("", index=df.index)).fillna("").astype(str).str.upper().str.strip()
        alt = df.get("Alternate Item (name)", pd.Series(index=df.index, dtype=object))
        for s in df.loc[alt.notna() & (flag == ""), "SKU"]:
            add("Warning", f"{s} names an alternate item but has no flag", "Not merged (R1 needs SWITCH). Ask product team")
        for s, f in zip(df.loc[~flag.isin(["", "SWITCH"]), "SKU"], flag[~flag.isin(["", "SWITCH"])]):
            add("Warning", f"{s} has undefined flag '{f}'", "Treated as a normal item. Ask supply chain what it means")
        dn = int(df["Item Name"].astype(str).str.contains(r"\s{2,}").sum() + alt.dropna().astype(str).str.contains(r"\s{2,}").sum())
        if dn:
            add("Fixed", f"{dn} names contain double spaces (alternates are matched by name)", "Names normalised before matching")
        n_in = int((df["Status"].astype(str).str.strip().str.title() != "Active").sum())
        if n_in:
            add("Info", f"{n_in} inactive item(s): {', '.join(df.loc[df['Status'].astype(str).str.strip().str.title() != 'Active', 'SKU'])}", "Left out of the analysis (R2)")
        for c in ("Master Pack Qty", "Unit Cost (Rs)"):
            m = int((pd.to_numeric(df[c], errors="coerce").fillna(0) <= 0).sum())
            if m:
                add("Warning", f"{m} items with missing or zero {c}", "Check at source")
    elif name == "Stock_Aging":
        if df["SKU"].astype(str).str.strip().str.upper().duplicated().any():
            add("Fixed", "Duplicate SKU rows", "Quantities summed")
        for b in BUCKETS:
            for _, r in df[df[b] < 0].iterrows():
                add("Fixed", f"{r['SKU']} has negative stock ({int(r[b])}) in '{b}'", "Set to 0. Warehouse to correct at source")
    elif name == "Sales_Lines":
        m = df["Month"].map(parse_month)
        odd = int(df["Month"].astype(str).str.match(r"^\d{4}-\d{2}").sum())
        if odd:
            add("Fixed", f"{odd} rows have month as '2026-06' (Excel converted text to a date)", "Both formats parsed")
        if m.isna().any():
            add("Warning", f"{int(m.isna().sum())} rows have an unreadable month", "Check at source")
        key = pd.DataFrame({"m": m, "c": df["CustomerID"], "s": df["SKU"].astype(str).str.strip().str.upper(), "q": df["Qty"], "v": df["Value (Rs)"]})
        dup = key.duplicated()
        if dup.any():
            add("Fixed", f"{int(dup.sum())} exact duplicate lines worth Rs {df.loc[dup, 'Value (Rs)'].sum():,.0f}", "Removed. Confirm they are not real repeat orders")
        if (df["Qty"] <= 0).any():
            add("Fixed", f"{int((df['Qty'] <= 0).sum())} lines with zero/negative quantity", "Removed")
        if m.nunique() != MONTHS:
            add("Warning", f"Sales cover {m.nunique()} months, not {MONTHS}", "Check the full period was exported")
    elif name == "Customers":
        if df["CustomerID"].duplicated().any():
            add("Fixed", "Duplicate customer IDs", "First kept")
        if df["Industry"].isna().any():
            add("Warning", f"Industry missing for {list(df.loc[df['Industry'].isna(), 'CustomerID'])}", "Excluded from the 'similar business' bucket")
    return out
