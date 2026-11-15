import io
import re
from datetime import datetime

import openpyxl
import pandas as pd
import pypdf
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Packing List to Order QTY list", layout="wide")
st.title("📦 Glass Packing List ➔ Order QTY list")
st.write(
    "Upload factory **Packing List PDFs** to get a consolidated Delivery Advice "
    "plus a **per-order summary with Color and Quantity**."
)

# ---------------- Sidebar settings ----------------
st.sidebar.header("Settings")
weight_factor = st.sidebar.number_input("Weight (kg per m²)", value=35.0, step=0.5)
color_text = st.sidebar.text_area(
    "Color keywords (one per line)",
    "Clear\nGrey\nGray\nBronze\nBlue\nGreen\nEuro Grey\nReflective\nLow-E\nTinted\nBlack\nWhite\nGold\nSilver",
    height=200,
)
COLOR_KEYWORDS = [c.strip() for c in color_text.splitlines() if c.strip()]

st.sidebar.caption("Colors are detected from the item Reference text.")

orders_text = st.text_area(
    "Optional: paste Work Orders to include (duplicates and blank lines are fine). Leave empty for all.",
    height=150,
)
wanted_orders = sorted({int(x) for x in re.findall(r"\d{8}", orders_text)})

uploaded_files = st.file_uploader(
    "Upload Packing List PDFs", type=["pdf"], accept_multiple_files=True
)


def fmt_date(s):
    if not s:
        return "Unknown"
    for f in ("%m/%d/%Y", "%d/%m/%Y"):  # adjust order to match your PDF
        try:
            d = datetime.strptime(s, f)
            return f"{d.day}-{d.strftime('%b')}"
        except ValueError:
            continue
    return s


def detect_color(ref):
    found = [k for k in COLOR_KEYWORDS if re.search(re.escape(k), ref, re.I)]
    # drop shorter keywords contained in longer ones (e.g. "Grey" inside "Euro Grey")
    found = [k for k in found if not any(k != o and k.lower() in o.lower() for o in found)]
    return " + ".join(found) if found else ""


@st.cache_data(show_spinner="Reading PDFs...")
def parse_packing_pdfs(file_blobs, weight_factor, color_keywords):
    rows = []
    last_rack = "Unknown"
    for blob in file_blobs:
        reader = pypdf.PdfReader(io.BytesIO(blob))
        for page in reader.pages:
            text = page.extract_text() or ""
            m = re.search(r"Packaging\s+(\d+)\s+packed", text)
            rack = m.group(1) if m else last_rack
            last_rack = rack
            dm = re.search(r"Delivery date\s+([\d/]+)", text)
            date = fmt_date(dm.group(1)) if dm else "Unknown"

            lines = [l.strip() for l in text.split("\n") if l.strip()]
            i = 0
            while i < len(lines):
                m = re.match(r"^(\d{8})\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+DGU\s+(.*)$", lines[i])
                if m:
                    labels = set()
                    extra = []
                    j = i + 1
                    while j < len(lines):
                        l2 = lines[j]
                        if (
                            re.match(r"^\d{8}\s+\d+\s+[\d.]+\s+[\d.]+\s+DGU", l2)
                            or "Item quantity" in l2
                            or "Order quantity" in l2
                        ):
                            break
                        lm = re.search(r"^(029\d+)", l2)
                        if lm:
                            labels.add(lm.group(1))
                        else:
                            extra.append(l2)
                        j += 1
                    rows.append(
                        dict(
                            order_no=int(m.group(1)),
                            item_no=int(m.group(2)),
                            width=float(m.group(3)),
                            height=float(m.group(4)),
                            ref=m.group(5).strip(),
                            rack=rack,
                            date=date,
                            labels=labels,
                            desc=" ".join(extra),
                        )
                    )
                    i = j - 1
                i += 1

    if not rows:
        return pd.DataFrame()

    items = {}
    for r in rows:
        key = (r["order_no"], r["item_no"])
        if key not in items:
            sqm = (r["width"] * r["height"]) / 1_000_000.0
            items[key] = dict(
                order_no=r["order_no"], item_no=r["item_no"], ref=r["ref"],
                width=r["width"], height=r["height"], sqm=sqm,
                weight=sqm * weight_factor, labels=set(), boxes={}, desc=r["desc"],
            )
        it = items[key]
        new = r["labels"] - it["labels"]  # ignore duplicate labels from overlapping PDFs
        it["labels"] |= new
        if new:
            bk = f"{r['rack']} / {r['date']}"
            it["boxes"][bk] = it["boxes"].get(bk, 0) + len(new)

    out = []
    for key in sorted(items):
        it = items[key]
        qty = len(it["labels"])
        out.append({
            "Work Order": it["order_no"],
            "Item No.": it["item_no"],
            "Color": detect_color(it["ref"] + " " + it["desc"]),
            "Reference": it["ref"],
            "Glass Description": it["desc"],
            "Width (mm)": int(it["width"]),
            "Height (mm)": int(it["height"]),
            "SQM (each)": round(it["sqm"], 3),
            "Weight (kg each)": round(it["weight"], 2),
            "Qty (units)": qty,
            "Total SQM": round(it["sqm"] * qty, 2),
            "Total Weight (kg)": round(it["weight"] * qty, 2),
            "Boxes / Delivery Date (Qty)": ", ".join(f"{k} ({v})" for k, v in it["boxes"].items()),
        })
    return pd.DataFrame(out)


def raw_preview(blob, pages=2):
    reader = pypdf.PdfReader(io.BytesIO(blob))
    return "\n\n----- PAGE -----\n".join((p.extract_text() or "") for p in reader.pages[:pages])


def build_order_summary(df):
    g = df.groupby("Work Order")
    summary = g.agg(
        Colors=("Color", lambda s: " | ".join(sorted(set(s)))),
        Items=("Item No.", "count"),
        Qty=("Qty (units)", "sum"),
        SQM=("Total SQM", "sum"),
        Weight=("Total Weight (kg)", "sum"),
    ).reset_index()
    summary.columns = ["Work Order", "Color", "Line Items", "Total Qty", "Total SQM", "Total Weight (kg)"]
    summary["Total SQM"] = summary["Total SQM"].round(2)
    summary["Total Weight (kg)"] = summary["Total Weight (kg)"].round(2)

    # qty per color inside each order
    by_color = (
        df.groupby(["Work Order", "Color"])["Qty (units)"].sum().reset_index()
        .rename(columns={"Qty (units)": "Qty"})
    )
    return summary, by_color


def write_sheet(ws, df, widths, wrap_cols=()):
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(list(df.columns))
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = PatternFill("solid", start_color="1F497D", end_color="1F497D")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for r in df.itertuples(index=False):
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Calibri", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in wrap_cols)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def generate_excel(df_items, df_summary, df_color):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Order Summary"
    write_sheet(ws, df_summary, [14, 40, 12, 12, 12, 16], wrap_cols=(2,))

    ws2 = wb.create_sheet("Qty by Color")
    write_sheet(ws2, df_color, [14, 40, 10])

    ws3 = wb.create_sheet("Delivery Advice")
    write_sheet(ws3, df_items, [14, 10, 20, 16, 40, 12, 12, 12, 14, 10, 12, 16, 80], wrap_cols=(5, 13))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


if uploaded_files:
    blobs = tuple(f.getvalue() for f in uploaded_files)
    df = parse_packing_pdfs(blobs, weight_factor, tuple(COLOR_KEYWORDS))

    if df.empty:
        st.warning("No valid item records found in the uploaded PDFs.")
    else:
        if wanted_orders:
            missing = [o for o in wanted_orders if o not in set(df["Work Order"])]
            df = df[df["Work Order"].isin(wanted_orders)]
            if missing:
                st.warning(f"Work orders not found in the PDFs: {', '.join(map(str, missing))}")
        if df.empty:
            st.warning("None of the listed work orders were found.")
            st.stop()

        with st.expander("🔍 Raw PDF text (to see where the color is written)"):
            st.code(raw_preview(blobs[0]))

        # ---- Work Order -> Color mapping (auto-detected, editable) ----
        auto = (
            df[df["Color"] != ""].groupby("Work Order")["Color"]
            .agg(lambda s: " + ".join(sorted(set(s)))).to_dict()
        )
        map_df = pd.DataFrame({
            "Work Order": sorted(df["Work Order"].unique()),
        })
        map_df["Color"] = map_df["Work Order"].map(auto).fillna("")
        st.subheader("🎨 Order Colors (fill in / correct once per order)")
        map_df = st.data_editor(
            map_df, hide_index=True, use_container_width=True,
            disabled=["Work Order"], key="color_map",
        )
        df = df.copy()
        df["Color"] = df["Work Order"].map(dict(zip(map_df["Work Order"], map_df["Color"])))
        df["Color"] = df["Color"].fillna("").replace("", "Not set")

        df_summary, df_color = build_order_summary(df)

        c1, c2, c3 = st.columns(3)
        c1.metric("Work Orders", len(df_summary))
        c2.metric("Total Qty", int(df_summary["Total Qty"].sum()))
        c3.metric("Total SQM", round(df_summary["Total SQM"].sum(), 2))

        st.subheader("📋 Order Summary (Color & Qty per Work Order)")
        st.dataframe(df_summary, use_container_width=True, hide_index=True)

        st.subheader("Qty by Color per Work Order")
        st.dataframe(df_color, use_container_width=True, hide_index=True)

        st.subheader("📄 Detailed Delivery Advice")
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.download_button(
            "📥 Download Excel (.xlsx)",
            data=generate_excel(df, df_summary, df_color),
            file_name="Delivery_Advice_Consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
