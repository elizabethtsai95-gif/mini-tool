"""Alpro Pharmacy - Digital Prescription Recording.

Staff upload the extracted prescription file (sheet "PrescriptionData"),
enter the starting serial number, and copy the register rows as plain values.

Run:  streamlit run Doc2us.py
"""
import html
import re

import pandas as pd
import streamlit as st

REQUIRED_COLS = [
    "Dispensed_Date", "ERX_ID", "FullName", "Address", "Drug_Name", "Dispensed_Amount",
]

HEADERS = [
    "DATE OF SALE OR SUPPLY OF POISON/MEDICINE",
    "*SERIAL NUMBER OF THE ENTRY IN THE REGISTER OF THE PRESCRIPTION",
    "NAME OF PATIENT/RECIPIENT/PERSON",
    "ADDRESS OF PATIENT/RECIPIENT/PERSON",
    "NAME OF THE POISON/MEDICINE SUPPLIED",
    "QUANTITY OF POISON/MEDICINE SUPPLIED",
]


# ---------------------------------------------------------------- transform
def clean_address(value) -> str:
    """Flatten a multi-line address into one line."""
    if pd.isna(value):
        return ""
    parts = [p.strip() for p in str(value).splitlines()]
    parts = [p for p in parts if p and p != "-"]
    return re.sub(r"\s+", " ", " ".join(parts)).strip()


def to_register_rows(df: pd.DataFrame, start_serial: int) -> pd.DataFrame:
    """One serial per prescription (ERX_ID). Extra drugs go on following rows
    with date/serial/name/address left blank, as in the register layout."""
    df = df.copy()
    df["Dispensed_Date"] = pd.to_datetime(df["Dispensed_Date"], errors="coerce")
    if df["Dispensed_Date"].isna().any():
        raise ValueError("Some Dispensed_Date values could not be read as dates.")

    df = df.sort_values(["Dispensed_Date", "ERX_ID"], kind="stable")

    rows, serial = [], start_serial
    for _, grp in df.groupby(["Dispensed_Date", "ERX_ID"], sort=False):
        first = grp.iloc[0]
        for i, (_, r) in enumerate(grp.iterrows()):
            rows.append({
                "date": first["Dispensed_Date"].strftime("%Y %m %d") if i == 0 else "",
                "serial": serial if i == 0 else "",
                "name": str(first["FullName"]).strip() if i == 0 else "",
                "address": clean_address(first["Address"]) if i == 0 else "",
                "drug": str(r["Drug_Name"]).strip(),
                "qty": str(r["Dispensed_Amount"]).strip(),
            })
        serial += 1
    return pd.DataFrame(rows)


# ------------------------------------------------------------------- copy
def table_html(reg: pd.DataFrame, with_header: bool) -> str:
    td = "border:1px solid #bbb;padding:3px 8px;color:#111;background:#fff;font-size:13px;vertical-align:top;"
    th = td + "background:#c9daf8;font-weight:bold;"
    esc = html.escape
    out = ['<div style="overflow-x:auto"><table style="border-collapse:collapse">']
    if with_header:
        out.append("<tr>" + "".join(f'<th style="{th}">{esc(h)}</th>' for h in HEADERS) + "</tr>")
    for rec in reg.itertuples(index=False):
        out.append("<tr>" + "".join(f'<td style="{td}">{esc(str(v))}</td>' for v in rec) + "</tr>")
    out.append("</table></div>")
    return "".join(out)


def tsv_text(reg: pd.DataFrame, with_header: bool) -> str:
    def cell(v):
        return str(v).replace("\t", " ").replace("\n", " ").replace('"', "'")

    lines = ["\t".join(HEADERS)] if with_header else []
    lines += ["\t".join(cell(v) for v in rec) for rec in reg.itertuples(index=False)]
    return "\n".join(lines)


# ---------------------------------------------------------------------- UI
st.set_page_config(page_title="Doc2us", page_icon="💊", layout="wide")
st.title("💊 Digital Prescription Recording")
st.caption("Upload the extracted prescription file, set the starting serial number, then copy the rows into the register.")

col_a, col_b = st.columns([2, 1])
with col_a:
    upload = st.file_uploader("Extracted prescription file (.xlsx)", type=["xlsx"])
with col_b:
    start_serial = st.number_input(
        "Starting serial number", min_value=1, value=1, step=1,
        help="Next unused serial in the register, i.e. last entry's serial + 1.",
    )

if upload is None:
    st.info("Upload the extracted .xlsx file to begin.")
    st.stop()

try:
    src = pd.read_excel(upload, sheet_name=0)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read the file: {e}")
    st.stop()

missing = [c for c in REQUIRED_COLS if c not in src.columns]
if missing:
    st.error(f"Missing column(s): {', '.join(missing)}. Please upload the original extract without editing the headers.")
    st.stop()

src = src.dropna(subset=["ERX_ID"])
try:
    reg = to_register_rows(src, int(start_serial))
except ValueError as e:
    st.error(str(e))
    st.stop()

reg = reg.rename(columns=dict(zip(
    ["date", "serial", "name", "address", "drug", "qty"],
    ["Date", "Serial", "Name", "Address", "Medicine", "Quantity"],
)))

n_entries = reg["Serial"].replace("", pd.NA).dropna().shape[0]
last_serial = int(start_serial) + n_entries - 1
m1, m2, m3 = st.columns(3)
m1.metric("Prescriptions (serials)", n_entries)
m2.metric("Medicine lines", len(reg))
m3.metric("Serial range", f"{int(start_serial)} – {last_serial}")
st.success(f"Next starting serial for your next upload: **{last_serial + 1}**")

with_header = st.checkbox("Include the header row when copying", False)

st.subheader("Register rows")
st.caption("Drag from the first cell to the last cell, press Ctrl + C, then paste with Ctrl + V in your register. "
           "All cells are plain values. A second medicine of the same prescription has blank date, serial, name "
           "and address cells.")
st.markdown(table_html(reg, with_header), unsafe_allow_html=True)

with st.expander("Copy as text instead"):
    st.caption("Click inside the box, press Ctrl + A, then Ctrl + C. Pasting into Excel or Google Sheets splits it "
               "into columns.")
    st.text_area("Register rows", tsv_text(reg, with_header), height=220, label_visibility="collapsed")
