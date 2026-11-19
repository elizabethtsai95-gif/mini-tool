"""Poison C recording.

Staff upload the daily "OUTLET POISON C RECORDING" export and type the starting serial
number. The page shows the register rows as plain values to copy and paste.
Items on the exclusion list (the pseudoephedrine products) are left out of that register and
get their own recording instead, one tab per item, with IN / OUT / BALANCE.
"""
import hashlib
import html
import json
import re
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Poison C", page_icon="📒", layout="wide")

HEADERS = [
    "DATE OF SALE OR SUPPLY OF POISON/MEDICINE",
    "*SERIAL NUMBER OF THE ENTRY IN THE REGISTER OF THE PRESCRIPTION",
    "NAME OF PATIENT/RECIPIENT/PERSON",
    "ADDRESS OF PATIENT/RECIPIENT/PERSON",
    "NAME OF THE POISON/MEDICINE SUPPLIED",
    "QUANTITY OF POISON/MEDICINE SUPPLIED",
]
ITEM_HEADERS = ["DATE", "PATIENT/SUPPLIER NAME", "ADDRESS", "IC/INVOICE", "IN", "OUT", "BALANCE", "PIC OF RECORDING"]
DEFAULT_EXCLUDE = ["FEDAC", "CLARINASE", "ALLERPID", "CARINOX", "BEATAFED", "AERIUS-D"]
EXCLUDE_FILE = Path(__file__).resolve().parent.parent / "exclude_items.txt"
CONFIG_FILE = Path(__file__).resolve().parent.parent / "sheet_config.json"
REQUIRED = ["Date", "Client Name", "Client Bill Street", "Client Bill City", "Client Bill State",
            "Item Name", "Qty Total"]


# ---------------------------------------------------------------- helpers
def read_export(file) -> pd.DataFrame:
    """Read the export. The header row sits below a block of report settings."""
    raw = pd.read_excel(file, header=None, dtype=str)
    hdr = None
    for i in range(min(60, len(raw))):
        cells = [str(x).strip() for x in raw.iloc[i].fillna("")]
        if "Client Name" in cells and "Item Name" in cells:
            hdr = i
            break
    if hdr is None:
        raise ValueError("Could not find the header row with Client Name and Item Name.")
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(x).strip() for x in raw.iloc[hdr].fillna("")]
    return df.loc[:, [c for c in df.columns if c]]


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()
    return "" if s.upper() == "NAN" else s


def make_address(row) -> str:
    """Street + zipcode + city + state, without repeating anything already in the address."""
    out = ""
    for c in ("Client Bill Street", "Client Bill Zipcode", "Client Bill City", "Client Bill State"):
        part = clean_text(row.get(c))
        if not re.search(r"[A-Za-z0-9]", part):  # empty or just "-"
            continue
        if re.search(rf"(?<![A-Za-z0-9]){re.escape(part)}(?![A-Za-z0-9])", out, flags=re.IGNORECASE):
            continue
        out = f"{out} {part}".strip()
    return out


def fmt_qty(q: float) -> str:
    return str(int(q)) if float(q).is_integer() else f"{q:g}"


def split_keywords(text: str) -> list[str]:
    return [t.strip().upper() for t in re.split(r"[,\n]", text) if t.strip()]


def saved_exclude_text() -> str:
    try:
        if EXCLUDE_FILE.exists():
            return EXCLUDE_FILE.read_text(encoding="utf-8").strip()
    except OSError:
        pass
    return "\n".join(DEFAULT_EXCLUDE)


def prepare_lines(df: pd.DataFrame, exclude_keys: list[str]) -> pd.DataFrame:
    """Sales lines with a Reason: '' (goes in the register), 'ITEM:<key>' (own recording) or a problem."""
    df = df.copy()
    code = df["Item Code"].map(clean_text).str.upper() if "Item Code" in df.columns else pd.Series("", index=df.index)
    name = df["Client Name"].map(clean_text)
    item = df["Item Name"].map(clean_text)
    # per-customer and per-day total lines are not sales
    df = df[~(code.eq("TOTAL") | name.str.upper().eq("TOTAL") | item.eq(""))].copy()

    df["when"] = pd.to_datetime(df["Date"], errors="coerce")
    df["qty"] = pd.to_numeric(df["Qty Total"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    df["address"] = df.apply(make_address, axis=1)
    df["item"] = df["Item Name"].map(clean_text)
    df["name"] = df["Client Name"].map(clean_text)
    df["ic"] = df["Client IC"].map(clean_text) if "Client IC" in df.columns else ""

    def reason(r):
        if pd.isna(r["when"]) or pd.isna(r["qty"]):
            return "Date or quantity missing"
        if r["qty"] <= 0:
            return "Quantity is zero or below"
        up = r["item"].upper()
        for k in exclude_keys:
            if k in up:
                return f"ITEM:{k}"
        if not r["address"]:
            return "No address"
        return ""

    df["reason"] = df.apply(reason, axis=1)
    return df


def build_register(df: pd.DataFrame, start_serial: int):
    """Return (register rows, lines left out with the reason, register lines with no name)."""
    left_out = df[~df["reason"].isin([""]) & ~df["reason"].str.startswith("ITEM:")][
        ["Date", "Client Name", "Item Name", "Qty Total", "reason"]].rename(columns={"reason": "Reason"})
    main = df[df["reason"] == ""].sort_values("when", kind="stable").reset_index(drop=True)
    reg = pd.DataFrame({
        "Date": main["when"].dt.strftime("%d/%m/%Y"),
        "Serial": range(start_serial, start_serial + len(main)),
        "Name": main["name"],
        "Address": main["address"],
        "Medicine": main["item"],
        "Quantity": main["qty"].map(fmt_qty),
    })
    return reg, left_out, reg[reg["Name"] == ""]


def build_item_record(lines: pd.DataFrame, opening, pic: str) -> pd.DataFrame:
    """OUT lines of one item with a running balance (blank until the opening balance is known)."""
    lines = lines.sort_values("when", kind="stable")
    rows, bal = [], opening
    for r in lines.itertuples():
        if bal is not None:
            bal = bal - r.qty
        rows.append({
            "DATE": r.when.strftime("%d/%m/%Y"), "PATIENT/SUPPLIER NAME": r.name, "ADDRESS": r.address,
            "IC/INVOICE": r.ic, "IN": "", "OUT": fmt_qty(r.qty),
            "BALANCE": "" if bal is None else fmt_qty(bal), "PIC OF RECORDING": pic,
        })
    return pd.DataFrame(rows)


def table_html(df: pd.DataFrame, headers: list[str], with_header: bool) -> str:
    td = "border:1px solid #bbb;padding:3px 8px;color:#111;background:#fff;font-size:13px;vertical-align:top;"
    th = td + "background:#c9daf8;font-weight:bold;"
    esc = html.escape
    out = ['<div style="overflow-x:auto"><table style="border-collapse:collapse">']
    if with_header:
        out.append("<tr>" + "".join(f'<th style="{th}">{esc(h)}</th>' for h in headers) + "</tr>")
    for rec in df.itertuples(index=False):
        out.append("<tr>" + "".join(f'<td style="{td}">{esc(str(v))}</td>' for v in rec) + "</tr>")
    out.append("</table></div>")
    return "".join(out)


def tsv_text(df: pd.DataFrame, headers: list[str], with_header: bool) -> str:
    def cell(v):
        return str(v).replace("\t", " ").replace("\n", " ").replace('"', "'")

    lines = ["\t".join(headers)] if with_header else []
    lines += ["\t".join(cell(v) for v in rec) for rec in df.itertuples(index=False)]
    return "\n".join(lines)


# ------------------------------------------------------------- google sheet
def load_sheet_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            d = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if d.get("url") and d.get("secret"):
                return {"url": str(d["url"]).strip(), "secret": str(d["secret"]).strip()}
    except (OSError, ValueError):
        pass
    return {}


def save_sheet_config(url: str, secret: str) -> None:
    CONFIG_FILE.write_text(json.dumps({"url": url.strip(), "secret": secret.strip()}), encoding="utf-8")


def _decode(resp) -> dict:
    try:
        return resp.json()
    except ValueError:
        return {"ok": False, "error": "The Google web app did not send data back. Check the deployment: "
                                      "Who has access must be Anyone, and the URL must end with /exec."}


def sheet_get(cfg: dict, **params) -> dict:
    try:
        return _decode(requests.get(cfg["url"], params={"secret": cfg["secret"], **params}, timeout=30))
    except requests.RequestException:
        return {"ok": False, "error": "Could not reach Google. Check the internet connection and the Web app URL."}


def sheet_post(cfg: dict, payload: dict) -> dict:
    try:
        body = json.dumps({"secret": cfg["secret"], **payload})
        return _decode(requests.post(cfg["url"], data=body.encode("utf-8"),
                                     headers={"Content-Type": "application/json"}, timeout=60))
    except requests.RequestException:
        return {"ok": False, "error": "Could not reach Google. Check the internet connection and the Web app URL."}


@st.cache_data(ttl=120, show_spinner=False)
def fetch_last(url: str, secret: str, tab: str, nonce: int) -> dict:
    """Tab name and last balance of the item's tab. nonce changes to force a fresh read."""
    return sheet_get({"url": url, "secret": secret}, action="last", tab=tab)


def parse_balance(text: str):
    t = text.strip().replace(",", "")
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return "bad"


# --------------------------------------------------------------------- UI
st.title("📒 Poison C Recording")
st.caption("Upload today's Poison C export, type the starting serial number, then copy the rows into the register.")

c1, c2 = st.columns([2, 1])
with c1:
    upload = st.file_uploader("Outlet Poison C recording export (.xlsx)", type=["xlsx"])
with c2:
    start = st.number_input("Starting serial number", min_value=1, value=None, step=1, format="%d",
                            placeholder="Type the next serial, e.g. 54822",
                            help="The serial of the first line. Each line after it goes up by 1.")

st.session_state.setdefault("exclude_text", saved_exclude_text())
ex_col, btn_col = st.columns([3, 1])
with ex_col:
    exclude_text = st.text_area(
        "Items recorded separately (one per line, or separated by commas)", key="exclude_text", height=110,
        help="Any item whose name contains one of these words is left out of the register above and gets its own "
             "recording below, one tab for each item. Type a new word on a new line to add it.")
with btn_col:
    st.write("")
    st.write("")
    if st.button("Save this list for next time"):
        try:
            EXCLUDE_FILE.write_text("\n".join(split_keywords(exclude_text)), encoding="utf-8")
            st.success("Saved for everyone who uses this page.")
        except OSError as e:
            st.error(f"Could not save: {e}")
    st.caption("Without saving, your additions last until you refresh the page.")

if upload is None:
    st.info("Upload the export to begin.")
    st.stop()

try:
    raw = read_export(upload)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read the file: {e}")
    st.stop()
missing = [c for c in REQUIRED if c not in raw.columns]
if missing:
    st.error(f"Missing column(s): {', '.join(missing)}. Please upload the original export.")
    st.stop()

exclude_keys = split_keywords(exclude_text)
lines = prepare_lines(raw, exclude_keys)
with_header = st.checkbox("Include the header row when copying", False)

# ---------------------------------------------------------- main register
st.header("Poison C register")
if start is None:
    st.info("Type the starting serial number to see the register rows.")
else:
    reg, left_out, no_name = build_register(lines, int(start))
    if reg.empty:
        st.info("No lines for the register in this file.")
    else:
        total_qty = pd.to_numeric(reg["Quantity"]).sum()
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Lines", len(reg))
        m2.metric("Total quantity", fmt_qty(total_qty))
        m3.metric("Serial range", f"{int(start)} – {int(start) + len(reg) - 1}")
        m4.metric("Next starting serial", int(start) + len(reg))
        m5.metric("Left out", len(left_out))

        if not no_name.empty:
            st.warning(f"{len(no_name)} line(s) have no name. Please fill them in after pasting.")
            st.dataframe(no_name, hide_index=True, use_container_width=True)

        st.caption("Drag from the first cell to the last cell, press Ctrl + C, then paste with Ctrl + V in your "
                   "register. All cells are plain values.")
        st.markdown(table_html(reg, HEADERS, with_header), unsafe_allow_html=True)
        with st.expander("Copy as text instead"):
            st.caption("Click inside the box, press Ctrl + A, then Ctrl + C. Pasting into Excel or Google Sheets "
                       "splits it into columns.")
            st.text_area("Register rows", tsv_text(reg, HEADERS, with_header), height=220,
                         label_visibility="collapsed", key="tsv_main")
    if not left_out.empty:
        with st.expander(f"{len(left_out)} line(s) left out of the register", expanded=True):
            st.caption("No serial number is used for these lines.")
            st.dataframe(left_out, hide_index=True, use_container_width=True)

# --------------------------------------------- separate item recordings
st.header("Separate item recordings")
item_lines = lines[lines["reason"].str.startswith("ITEM:")].copy()
if item_lines.empty:
    st.info("No lines for the separately recorded items in this file.")
else:
    item_lines["key"] = item_lines["reason"].str.replace("ITEM:", "", regex=False)
    keys = [k for k in exclude_keys if k in set(item_lines["key"])]
    cfg = load_sheet_config()
    nonce = st.session_state.setdefault("sheet_nonce", 0)
    done = st.session_state.setdefault("sheet_done", {})

    with st.expander("Google Sheet connection", expanded=not cfg):
        st.caption("Paste the Web app URL and the secret from the Google Sheet script. They are saved on this PC.")
        url_in = st.text_input("Web app URL", cfg.get("url", ""), key="sheet_url_in")
        secret_in = st.text_input("Secret", cfg.get("secret", ""), type="password", key="sheet_secret_in")
        b1, b2, _ = st.columns([1, 1, 3])
        if b1.button("Save connection"):
            if url_in.strip() and secret_in.strip():
                try:
                    save_sheet_config(url_in, secret_in)
                    st.success("Saved.")
                    st.rerun()
                except OSError as e:
                    st.error(f"Could not save: {e}")
            else:
                st.error("Fill in both the URL and the secret.")
        if b2.button("Test connection", disabled=not (url_in.strip() and secret_in.strip())):
            res = sheet_get({"url": url_in.strip(), "secret": secret_in.strip()}, action="ping")
            if res.get("ok"):
                st.success("Connected. Tabs found: " + ", ".join(res.get("tabs", [])))
            else:
                st.error(res.get("error", "Not connected."))

    pic = st.text_input("PIC of recording", "", placeholder="Name of the person recording, e.g. IMAN",
                        help="Goes in the PIC OF RECORDING column of every line below.")
    if cfg:
        if st.button("Reload balances from the sheet"):
            st.session_state.sheet_nonce += 1
            st.rerun()
        st.caption("Each balance is read from the last row of the item's tab. You can change it if it is wrong. "
                   "IN (stock received) is left blank, so add it by hand.")
    else:
        st.caption("Type each item's balance from your sheet (the balance on its last row) so the new balances "
                   "follow on. IN (stock received) is left blank, so add it by hand.")

    tabs = st.tabs(keys)
    for tab, key in zip(tabs, keys):
        with tab:
            part = item_lines[item_lines["key"] == key]
            batch_id = hashlib.sha1("|".join(
                [key] + [f"{r.when:%Y%m%d}{r.name}{r.ic}{r.qty}" for r in part.itertuples()]).encode()).hexdigest()
            if batch_id in done:
                info = done[batch_id]
                st.success(f"Added {info['added']} line(s) to the {info['tab']} tab. "
                           f"The balance in the sheet is now {fmt_qty(info['balance'])}.")
                st.markdown(info["html"], unsafe_allow_html=True)
                continue

            sheet_info = fetch_last(cfg["url"], cfg["secret"], key, nonce) if cfg else {}
            default_open = ""
            if sheet_info.get("ok") and sheet_info.get("balance") is not None:
                default_open = fmt_qty(sheet_info["balance"])
            opening_txt = st.text_input(f"{key}: balance before these lines", default_open,
                                        key=f"open_{key}_{nonce}_{default_open}", placeholder="Balance on the last row of the sheet")
            opening = parse_balance(opening_txt)
            if opening == "bad":
                st.error("Type the balance as a number.")
                opening = None
            if cfg and sheet_info.get("ok"):
                st.caption(f"Sheet tab: {sheet_info.get('tab')}")
            elif cfg:
                st.warning(sheet_info.get("error", "Could not read the sheet.")
                           + (" Tabs in the sheet: " + ", ".join(sheet_info["tabs"]) if sheet_info.get("tabs") else ""))
            if opening is None:
                st.info("Type the balance to fill in the BALANCE column.")

            rec = build_item_record(part, opening, pic.strip().upper())
            gaps = part[(part["name"] == "") | (part["address"] == "") | (part["ic"] == "")]
            if not gaps.empty:
                st.warning(f"{len(gaps)} line(s) are missing a name, address or IC. Please complete them "
                           "after pasting.")
            c_a, c_b = st.columns(2)
            c_a.metric("Lines", len(rec))
            c_b.metric("Total OUT", fmt_qty(part["qty"].sum()))
            st.markdown(table_html(rec, ITEM_HEADERS, with_header), unsafe_allow_html=True)
            with st.expander("Copy as text instead"):
                st.text_area(f"{key} rows", tsv_text(rec, ITEM_HEADERS, with_header), height=180,
                             label_visibility="collapsed", key=f"tsv_{key}")

            if cfg:
                st.markdown("**Add to the Google Sheet**")
                ok_box = st.checkbox("I have checked these lines", key=f"ok_{key}_{batch_id[:8]}")
                force = st.checkbox("Add even if some lines look already added", key=f"force_{key}_{batch_id[:8]}")
                missing_bits = []
                if not sheet_info.get("ok"):
                    missing_bits.append("the sheet tab was not found")
                if opening is None:
                    missing_bits.append("type the balance")
                if not pic.strip():
                    missing_bits.append("type the PIC of recording")
                if missing_bits:
                    st.caption("Before adding: " + "; ".join(missing_bits) + ".")
                if st.button(f"Add {len(rec)} line(s) to the Google Sheet", key=f"add_{key}_{batch_id[:8]}",
                             type="primary", disabled=bool(missing_bits) or not ok_box):
                    rows = [{"date": r["DATE"], "name": r["PATIENT/SUPPLIER NAME"], "address": r["ADDRESS"],
                             "ic": r["IC/INVOICE"], "in": "", "out": float(r["OUT"]),
                             "balance": float(r["BALANCE"]), "pic": r["PIC OF RECORDING"]}
                            for _, r in rec.iterrows()]
                    res = sheet_post(cfg, {"tab": key, "rows": rows, "expectedBalance": opening,
                                           "force": bool(force)})
                    if res.get("ok"):
                        done[batch_id] = {"added": res.get("added", len(rows)), "tab": res.get("tab", key),
                                          "balance": res.get("balance", opening),
                                          "html": table_html(rec, ITEM_HEADERS, with_header)}
                        st.session_state.sheet_nonce += 1
                        st.rerun()
                    elif res.get("duplicates"):
                        st.warning("These lines look already added to the sheet: " + "; ".join(res["duplicates"])
                                   + ". Tick 'Add even if some lines look already added' if they are new.")
                    else:
                        st.error(res.get("error", "The lines were not added."))
            else:
                st.caption("Connect the Google Sheet above to add these lines to it automatically.")
