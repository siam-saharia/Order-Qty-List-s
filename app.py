import base64
import io
import re
from collections import defaultdict

import openpyxl
import pandas as pd
import pypdf
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Packing List to Order QTY list", layout="wide")
st.title("📦 Glass Packing List ➔ Order QTY list")
st.caption(f"App version: v3 (xlsx download fix) | Streamlit {st.__version__}")
st.write(
    "Upload factory **Packing List PDFs** to get a consolidated Order QTY list "
    "plus a **per-order summary with Color and Quantity**. "
    "Duplicate labels are **counted and highlighted**, not removed."
)

# ---------------- Sidebar settings ----------------
st.sidebar.header("Settings")
weight_factor = st.sidebar.number_input("Weight (kg per m²)", value=45.0, step=0.5)
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

YELLOW = "FFFF00"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def fmt_date(s):
    return s if s else "Unknown"


def detect_color(ref, color_keywords):
    found = [k for k in color_keywords if re.search(re.escape(k), ref, re.I)]
    found = [k for k in found if not any(k != o and k.lower() in o.lower() for o in found)]
    return " + ".join(found) if found else ""


@st.cache_data(show_spinner="Reading PDFs...")
def parse_packing_pdfs(file_blobs, weight_factor, color_keywords):
    ITEM_RE = re.compile(r"^(\d{8})\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+DGU\s+(.*)$")
    LABEL_RE = re.compile(r"^(029\d+)\b")
    RACK_RE = re.compile(r"Packaging\s+(\d+)\s+packed")
    rows = []
    for blob in file_blobs:
        reader = pypdf.PdfReader(io.BytesIO(blob))
        stream = []
        for page in reader.pages:
            text = page.extract_text() or ""
            dm = re.search(r"Delivery date\s+([\d/]+)", text)
            date = fmt_date(dm.group(1)) if dm else "Unknown"
            for l in text.split("\n"):
                l = l.strip()
                if l:
                    stream.append((l, date))

        cur = None
        for line, date in stream:
            m = ITEM_RE.match(line)
            if m:
                cur = dict(
                    order_no=int(m.group(1)), item_no=int(m.group(2)),
                    width=float(m.group(3)), height=float(m.group(4)),
                    ref=m.group(5).strip(), desc="", packed=[],
                )
                rows.append(cur)
                continue
            lm = LABEL_RE.match(line)
            if lm and cur is not None:
                rm = RACK_RE.search(line)
                cur["packed"].append((lm.group(1), rm.group(1) if rm else "Unknown", date))

    if not rows:
        return pd.DataFrame(), pd.DataFrame()

    items = {}
    for r in rows:
        key = (r["order_no"], r["item_no"])
        if key not in items:
            sqm = (r["width"] * r["height"]) / 1_000_000.0
            items[key] = dict(
                order_no=r["order_no"], item_no=r["item_no"], ref=r["ref"],
                width=r["width"], height=r["height"], sqm=sqm,
                weight=sqm * weight_factor, entries={}, desc=r["desc"],
            )
        it = items[key]
        for label, rack, date in r["packed"]:
            k = (label, rack, date)
            if k in it["entries"]:
                # same label, same box, same date = same PDF uploaded twice -> skip
                continue
            it["entries"][k] = 1

    # labels that appear in MORE THAN ONE box (across all items/PDFs)
    label_locs = defaultdict(set)
    for it in items.values():
        for (label, rack, date) in it["entries"]:
            label_locs[label].add((rack, date))
    dup_labels = {l for l, s in label_locs.items() if len(s) > 1}

    out, dup_rows = [], []
    for key in sorted(items):
        it = items[key]
        boxes = {}
        for (label, rack, date) in it["entries"]:
            bk = f"{rack} / {date}"
            boxes[bk] = boxes.get(bk, 0) + 1
            if label in dup_labels:
                dup_rows.append({
                    "Label": label, "Work Order": it["order_no"],
                    "Item No.": it["item_no"], "Box": rack, "Delivery Date": date,
                })
        qty = len(it["entries"])  # EVERY label counts, duplicates included
        it_dups = sorted({l for (l, _, _) in it["entries"] if l in dup_labels})
        dup_text = ", ".join(
            f"{l} (Box {', '.join(sorted({r for r, _ in label_locs[l]}))})" for l in it_dups
        )
        out.append({
            "Work Order": it["order_no"],
            "Item No.": it["item_no"],
            "Color": detect_color(it["ref"] + " " + it["desc"], color_keywords),
            "Reference": it["ref"],
            "Glass Description": it["desc"],
            "Width (mm)": int(it["width"]),
            "Height (mm)": int(it["height"]),
            "SQM (each)": round(it["sqm"], 3),
            "Weight (kg each)": round(it["weight"], 2),
            "Qty (units)": qty,
            "Total SQM": round(it["sqm"] * qty, 2),
            "Total Weight (kg)": round(it["weight"] * qty, 2),
            "Boxes / Delivery Date (Qty)": ", ".join(f"{k} ({v})" for k, v in boxes.items()),
            "Duplicate Labels": dup_text,
        })
    dup_df = pd.DataFrame(dup_rows)
    if not dup_df.empty:
        dup_df = dup_df.sort_values(["Label", "Box"]).reset_index(drop=True)
    return pd.DataFrame(out), dup_df


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
    by_color = (
        df.groupby(["Work Order", "Color"])["Qty (units)"].sum().reset_index()
        .rename(columns={"Qty (units)": "Qty"})
    )
    return summary, by_color


def write_sheet(ws, df, widths, wrap_cols=(), highlight_col=None):
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
        ws.append([v.item() if hasattr(v, "item") else v for v in r])

    flags = None
    if highlight_col and highlight_col in df.columns:
        flags = (df[highlight_col].astype(str).str.strip() != "").tolist()
    for i, row in enumerate(ws.iter_rows(min_row=2)):
        hl = flags[i] if flags is not None and i < len(flags) else False
        for cell in row:
            cell.font = Font(name="Calibri", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column in wrap_cols)
            if hl:
                cell.fill = PatternFill("solid", start_color=YELLOW, end_color=YELLOW)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


PALETTE = ["F2DCDB", "EBF1DE", "E4DFEC", "FDE9D9", "D9D9D9", "DCE6F1", "FFF2CC", "DAEEF3"]


def build_advice_df(df):
    cols = ["Work Order", "Item No.", "Qty (units)", "Boxes / Delivery Date (Qty)", "Duplicate Labels"]
    return df[cols].copy()


def write_advice_sheet(ws, df):
    df = build_advice_df(df)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = list(df.columns)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = PatternFill("solid", start_color="1F497D", end_color="1F497D")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    row = 2
    hl_fill = PatternFill("solid", start_color=YELLOW, end_color=YELLOW)
    for n, (order, grp) in enumerate(df.groupby("Work Order", sort=True)):
        color = PALETTE[n % len(PALETTE)]
        fill = PatternFill("solid", start_color=color, end_color=color)
        for r in grp.itertuples(index=False):
            has_dup = bool(str(r[4]).strip())
            for c, v in enumerate(list(r), start=1):
                v = v.item() if hasattr(v, "item") else v
                cell = ws.cell(row=row, column=c, value=v)
                cell.fill = hl_fill if has_dup else fill
                cell.font = Font(name="Calibri", size=10, bold=(has_dup and c == 5),
                                 color="C00000" if (has_dup and c == 5) else "000000")
                cell.border = border
                cell.alignment = Alignment(horizontal="center" if c < 4 else "left",
                                           vertical="center", wrap_text=(c == 5))
            row += 1
        ws.cell(row=row, column=2, value="TOTAL")
        ws.cell(row=row, column=3, value=int(grp["Qty (units)"].sum()))
        for c in (2, 3):
            cell = ws.cell(row=row, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        row += 1

    ws.cell(row=row, column=2, value="GRAND TOTAL")
    ws.cell(row=row, column=3, value=int(df["Qty (units)"].sum()))
    for c in (2, 3):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for i, w in enumerate([14, 14, 12, 48, 45], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def generate_excel(df_items, df_summary, df_color, df_dups):
    """Returns the finished .xlsx as raw bytes."""
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "Delivery Advice"
    write_advice_sheet(ws0, df_items)

    ws = wb.create_sheet("Order Summary")
    write_sheet(ws, df_summary, [14, 40, 12, 12, 12, 16], wrap_cols=(2,))

    ws2 = wb.create_sheet("Qty by Color")
    write_sheet(ws2, df_color, [14, 40, 10])

    ws3 = wb.create_sheet("Details")
    write_sheet(ws3, df_items, [14, 10, 20, 16, 40, 12, 12, 12, 14, 10, 12, 16, 80, 45],
                wrap_cols=(5, 13, 14), highlight_col="Duplicate Labels")

    if df_dups is not None and not df_dups.empty:
        ws4 = wb.create_sheet("Duplicate Labels")
        d = df_dups.copy()
        d["Note"] = "Label appears in more than one box"
        write_sheet(ws4, d, [16, 14, 10, 8, 14, 36], highlight_col="Note")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()  # plain bytes: safest for st.download_button


def highlight_dups(df):
    def style_row(row):
        if str(row.get("Duplicate Labels", "")).strip():
            return ["background-color: #FFFF00; color: black"] * len(row)
        return [""] * len(row)
    return df.style.apply(style_row, axis=1)


if uploaded_files:
    blobs = tuple(f.getvalue() for f in uploaded_files)
    df, dup_df = parse_packing_pdfs(blobs, weight_factor, tuple(COLOR_KEYWORDS))

    if df.empty:
        st.warning("No valid item records found in the uploaded PDFs.")
    else:
        if wanted_orders:
            missing = [o for o in wanted_orders if o not in set(df["Work Order"])]
            df = df[df["Work Order"].isin(wanted_orders)]
            if not dup_df.empty:
                dup_df = dup_df[dup_df["Work Order"].isin(wanted_orders)]
            if missing:
                st.warning(f"Work orders not found in the PDFs: {', '.join(map(str, missing))}")
        if df.empty:
            st.warning("None of the listed work orders were found.")
            st.stop()

        with st.expander("🔍 Raw PDF text (to see where the color is written)"):
            st.code(raw_preview(blobs[0]))

        auto = (
            df[df["Color"] != ""].groupby("Work Order")["Color"]
            .agg(lambda s: " + ".join(sorted(set(s)))).to_dict()
        )
        map_df = pd.DataFrame({"Work Order": sorted(df["Work Order"].unique())})
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

        total_qty = int(df_summary["Total Qty"].sum())
        extra = 0 if dup_df.empty else len(dup_df) - dup_df["Label"].nunique()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Work Orders", len(df_summary))
        c2.metric("Total Qty (all boxes)", total_qty)
        c3.metric("Unique labels", total_qty - extra)
        c4.metric("Total SQM", round(df_summary["Total SQM"].sum(), 2))

        if not dup_df.empty:
            st.warning(
                f"⚠️ {dup_df['Label'].nunique()} label(s) appear in more than one box "
                f"({extra} extra count). They are INCLUDED in the Total Qty and highlighted in yellow."
            )
            st.subheader("🟡 Duplicate Labels")
            st.dataframe(dup_df, use_container_width=True, hide_index=True)

        st.subheader("📋 Order Summary (Color & Qty per Work Order)")
        st.dataframe(df_summary, use_container_width=True, hide_index=True)

        st.subheader("Qty by Color per Work Order")
        st.dataframe(df_color, use_container_width=True, hide_index=True)

        st.subheader("📑 Delivery Advice (Work Order / Item / Qty / Boxes)")
        st.dataframe(highlight_dups(build_advice_df(df)), use_container_width=True, hide_index=True)

        st.subheader("📄 Detailed table")
        st.dataframe(highlight_dups(df), use_container_width=True, hide_index=True)

        # Build the Excel once, then hand the bytes to the button
        excel_bytes = generate_excel(df, df_summary, df_color, dup_df)

        FILE_NAME = "Order_QTY_list.xlsx"  # no spaces, ends with .xlsx

        try:
            # on_click="ignore" stops Streamlit from rerunning the page on click
            st.download_button(
                "📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=FILE_NAME,
                mime=XLSX_MIME,
                key="download_xlsx",
                on_click="ignore",
            )
        except TypeError:  # older Streamlit without on_click="ignore"
            st.download_button(
                "📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=FILE_NAME,
                mime=XLSX_MIME,
                key="download_xlsx",
            )

        # Backup link: if the button still saves a file with no extension,
        # use this plain HTML link instead (the filename is set by the browser).
        b64 = base64.b64encode(excel_bytes).decode()
        st.markdown(
            f'<a href="data:{XLSX_MIME};base64,{b64}" download="{FILE_NAME}" '
            f'style="font-size:15px;">⬇️ Backup download link ({FILE_NAME})</a>',
            unsafe_allow_html=True,
        )
