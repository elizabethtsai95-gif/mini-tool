"""MVP tracker.

Staff upload the MVP SKU list (item code, item name, profit per unit, monthly target)
and the sale invoice item export. The page splits each SKU's target equally between
the staff, counts each person's sales of the SKU, and produces the MVP sheet
(Target / Achieve per staff, Total, Balance) as an Excel file and as an on-screen table
that can be copied and pasted as values.
"""
import html
import io
import json
import math
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MVP", page_icon="🏆", layout="wide")

# ------------------------------------------------------------------ defaults
# Starting point for the staff list (full name in the sales file, name on the sheet), in sheet order.
# Anyone else found in the sales file is worked out by the page, and the list can be changed and saved.
SEED_STAFF = [
    ("ELIZABETH TSAI YEN PENG", "ELIZABETH"),
    ("LIM JUN MING", "LIM JUN MING"),
    ("ELIZZA ANAK VALENTINE", "ELIZZA"),
    ("NUR SYAZWANI BINTI REDI", "SYAZWANI"),
    ("ELLYTHIA NISHA ANAK SIGON", "ELLYTHIA"),
    ("NURSYUHADA BINTI RIZALMAN", "SYUHADA"),
    ("MELVIN CRIST ANAK ANOM", "MELL"),
]
STAFF_FILE = Path(__file__).resolve().parent.parent / "mvp_staff.json"
NOT_STAFF = ("PHARMACY", "ALPRO", "ADMIN", "SYSTEM", "CASH", "ONLINE")

DEFAULT_GROUPS = "100621259, 100621218, 100621234"
DEFAULT_COMPULSORY = "100634849, 100079672, 100276807, 100290576"

TEMPLATE_COLUMNS = ["Item Code", "Item Name", "Profit per Unit", "Monthly Target",
                    "Extra Qty", "Baseline", "Compulsory"]
# Taken from the MVP project calculator (QKCMC1). Please check against the original sheet.
TEMPLATE_ROWS = [
    ("100621259", "BIOTECTIVE GASYCON 18X10S", 132, 4.38889, 4, 0.388889, ""),
    ("100621218", "BIOTECTIVE GINEBEX 18X 10S", 132, 1, 1, 0, ""),
    ("100621234", "BIOTECTIVE NOPREXIA 18X10S", 132, 2.98148, 2, 0.981481, ""),
    ("57447", "GUMMY RASPBERRY EMERGENC 21S", 6.2, 36.1667, 12, 24.1667, ""),
    ("60051", "PROMUNE MBG 500MG 60S", 138, 5.67, 5, 0.67, ""),
    ("50074", "BIOTECTIVE EKINASE 300S", 240, 6.87778, 6, 0.877778, ""),
    ("79985", "DD EZ RELIEF CREAM 15ML", 6.6, 36.5, 12, 24.5, ""),
    ("100639889", "TONI'MASK STEAM EYE MASK 7PCS", 22.04, 2, 2, 0, ""),
    ("100004423", "POWERLIFE DIARESO5 12X10S", 66.04, 4.66667, 3, 1.66667, ""),
    ("100567387", "GATTOKOSO 20G 15S", 64, 7.37778, 6, 1.37778, ""),
    ("68154", "POWERLIFE IRON PYROPHOSPHATE PLUS 60S", 29.3, 9, 4, 5, ""),
    ("100657162", "PURO IKOSA-EPA 489MG SOFTGEL 60S", 73.95, 3, 2, 1, ""),
    ("100611540", "OMRON BLOOD PRESSURE MONITOR HEM-7383T1", 160.54, 2, 2, 0, ""),
    ("100611458", "OMRON BLOOD PRESSURE MONITOR HEM-7183T1-AP (EZ BT)", 71.6, 3, 3, 0, ""),
    ("100634849", "YNBY YUNNAN BAIYAO PROBIOTIC TOOHTPASTE 100G", 13.56, 4, 4, 0, "Y"),
    ("100079672", "TONIMER EYES OPHTHALMIC WIPES 16S", 14.8, 2, 2, 0, "Y"),
    ("100276807", "TONIMER LAB: EYE DROPS 10ML", 15.4, 12, 12, 0, "Y"),
    ("100290576", "VOZAN 20MG FILM COATED TABLETS 3X10S", 55.7, 6, 6, 0, "Y"),
]

REQUIRED_SALES = ["Sales Date", "Item Code", "Sales Quantity", "Sales Person"]


# ------------------------------------------------------------------ helpers
def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")


def norm_exact(code) -> str:
    """Upper-case, no spaces, and (L)123 written as L123."""
    s = str(code).strip().upper().replace(" ", "")
    return re.sub(r"^\(([A-Z]{1,2})\)", r"\1", s)


def base_code(code) -> str:
    """Code without a T / L style pack prefix: T50074 -> 50074."""
    return re.sub(r"^[A-Z]{1,2}(?=\d)", "", norm_exact(code))


def split_codes(text: str) -> list[str]:
    return [base_code(t) for t in re.split(r"[,\n/]", text) if t.strip()]


def default_label(full: str) -> str:
    """Short name for the sheet: NUR SYAZWANI BINTI REDI -> SYAZWANI, ELIZABETH TSAI YEN PENG -> ELIZABETH."""
    n = re.sub(r"\s+", " ", str(full).upper()).strip()
    parts = re.split(r"\s(?:ANAK|BINTI|BIN|BT)\s", n, maxsplit=1)
    base, had_particle = parts[0], len(parts) > 1
    if base.startswith("NUR "):
        base = base[4:]
    elif base.startswith("NUR") and len(base.split()[0]) >= 9:  # NURSYUHADA, but not NURUL or NURAINI
        base = base[3:]
    words = base.split()
    if not words:
        return n
    if had_particle:
        return words[0]
    return " ".join(words[:3]) if len(words[0]) <= 3 else words[0]


def load_staff_file() -> dict:
    try:
        if STAFF_FILE.exists():
            return json.loads(STAFF_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    return {}


def detect_staff(sales: pd.DataFrame) -> pd.DataFrame:
    """Everyone in the Sales Person column, with a guess of who is staff and the name used on the sheet."""
    names = sales["Sales Person"].dropna().astype(str).str.strip()
    counts = names[names != ""].value_counts()
    saved = load_staff_file()
    if saved:
        labels, known = saved.get("labels", {}), set(saved.get("known", []))
        selected, order = set(saved.get("selected", [])), saved.get("order", {})
    else:
        labels = {f.upper(): l for f, l in SEED_STAFF}
        known = set(labels)
        selected = set(labels)
        order = {f.upper(): i + 1 for i, (f, _) in enumerate(SEED_STAFF)}
    top = int(counts.max()) if len(counts) else 0
    rows = []
    for name, n in counts.items():
        up = name.upper()
        if up in known:
            use = up in selected
        else:  # new name: staff if they sell a fair share of what the busiest person sells
            use = n >= 0.25 * top and not any(w in up for w in NOT_STAFF)
        rows.append({"Use": use, "Order": order.get(up, 100 + len(rows)), "Name in sales file": name,
                     "Name on sheet": labels.get(up) or default_label(up), "Sales lines": int(n)})
    df = pd.DataFrame(rows)
    return df.sort_values(["Use", "Order", "Sales lines"], ascending=[False, True, False]).reset_index(drop=True)


def read_mvp(file) -> pd.DataFrame:
    """Read the MVP SKU list. The header row can be anywhere in the first 25 rows."""
    raw = pd.read_excel(file, header=None, dtype=str)
    hdr = None
    for i in range(min(25, len(raw))):
        txt = " ".join(raw.iloc[i].fillna("").astype(str)).lower()
        if "profit" in txt and "target" in txt:
            hdr = i
            break
    if hdr is None:
        raise ValueError("Could not find the header row. It needs columns for profit per unit and target.")
    cols = raw.iloc[hdr].fillna("").astype(str).str.strip().tolist()

    def find(patterns, exclude=None):
        for pat in patterns:
            for j, c in enumerate(cols):
                lc = c.lower()
                if re.search(pat, lc) and not (exclude and re.search(exclude, lc)):
                    return j
        return None

    code_j = find([r"item code", r"\bcode\b", r"sku"])
    if code_j is None:  # unlabelled first column, as in the sheet
        code_j = next((j for j, c in enumerate(cols) if c == "" or c.lower().startswith("unnamed")), 0)
    name_j = find([r"item name", r"^item\s*$", r"description", r"^name"])
    profit_j = find([r"profit per unit", r"unit profit", r"^profit$"], exclude=r"yield")
    target_j = find([r"final target", r"monthly target", r"^target"])
    extra_j = find([r"extra"])
    base_j = find([r"baseline"])
    comp_j = find([r"compuls"])
    missing = [n for n, j in (("profit per unit", profit_j), ("target", target_j)) if j is None]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(missing))

    body = raw.iloc[hdr + 1:]
    out = pd.DataFrame({
        "code": body.iloc[:, code_j].fillna("").astype(str).str.strip().str.replace(r"\.0$", "", regex=True),
        "name": body.iloc[:, name_j].fillna("").astype(str).str.strip() if name_j is not None else "",
        "profit": to_num(body.iloc[:, profit_j]),
        "target": to_num(body.iloc[:, target_j]),
        "extra": to_num(body.iloc[:, extra_j]) if extra_j is not None else float("nan"),
        "baseline": to_num(body.iloc[:, base_j]) if base_j is not None else float("nan"),
        "compulsory": (body.iloc[:, comp_j].fillna("").astype(str).str.strip().str.upper()
                       .isin(["Y", "YES", "1", "TRUE", "COMPULSORY"])) if comp_j is not None else False,
    })
    out = out[out["code"].str.match(r"^\(?[A-Za-z]{0,2}\)?\d{3,}")].reset_index(drop=True)
    out["profit"] = out["profit"].fillna(0.0)
    out["target"] = out["target"].fillna(0.0)
    return out


def combined_name(names: list[str]) -> str:
    toks = [n.split() for n in names]
    pre = 0
    while all(len(t) > pre for t in toks) and len({t[pre] for t in toks}) == 1:
        pre += 1
    if pre == 0 or any(len(t) <= pre for t in toks):
        return " / ".join(names)
    return " ".join(toks[0][:pre]) + " " + "/".join(t[pre] for t in toks)


def pack_size(name):
    """Units in a pack from the item name: 18X10S -> 180, 60S -> 60, 7PCS -> 7. None if not found."""
    t = str(name).upper()
    m = re.findall(r"(\d+)\s*X\s*(\d+)\s*(?:S|PCS)\b", t)
    if m:
        return int(m[-1][0]) * int(m[-1][1])
    m = re.findall(r"(\d+)\s*(?:S|PCS)\b", t)
    return int(m[-1]) if m else None


def auto_ratio(mother_name, loose_name):
    """How many loose packs make one mother pack, read from the two item names."""
    a, b = pack_size(mother_name), pack_size(loose_name)
    if a and b and a >= b:
        return a / b
    return None


def prepare_sales(sales, staff):
    s = sales.copy()
    s["qty"] = to_num(s["Sales Quantity"]).fillna(0.0)
    s["exact"] = s["Item Code"].map(norm_exact)
    s["base"] = s["Item Code"].map(base_code)

    def who(name):
        nm = str(name).upper()
        for label, keys in staff:
            if any(k in nm for k in keys):
                return label
        return "Others"

    s["staff"] = s["Sales Person"].fillna("").map(who)
    return s


def build_matches(mvp, s, include_variants):
    """One row per (MVP SKU, sales code) with the pack ratio read from the names."""
    rows = []
    has_name = "Item Name" in s.columns
    for r in mvp.itertuples():
        ex = norm_exact(r.code)
        m = s[s["base"] == base_code(r.code)] if include_variants else s[s["exact"] == ex]
        if m.empty:
            continue
        g = m.groupby("exact").agg(
            code=("Item Code", "first"), name=("Item Name", "first") if has_name else ("Item Code", "first"),
            qty=("qty", "sum"), lines=("qty", "size")).reset_index()
        for x in g.itertuples():
            if x.exact == ex:
                ratio, note = 1.0, "mother code"
            else:
                ratio = auto_ratio(r.name, x.name)
                note = "from pack size" if ratio else "CHECK: pack size not found"
                ratio = ratio or 1.0
            rows.append({"MVP SKU": r.code, "Sales code": x.code, "Sales item name": x.name,
                         "Qty sold": float(x.qty), "Lines": int(x.lines), "Loose packs per mother pack": float(ratio),
                         "Note": note, "exact": x.exact})
    return pd.DataFrame(rows)


def compute(mvp, s, staff, include_variants, groups_text, comp_text, round_up, extra_mode, ratios):
    n = len(staff)
    labels = [l for l, _ in staff] + ["Others"]
    comp_codes = set(split_codes(comp_text))
    skus = []
    p_target = {l: 0.0 for l, _ in staff}
    p_ach = {l: 0.0 for l, _ in staff}
    for r in mvp.itertuples():
        ex = norm_exact(r.code)
        m = s[s["base"] == base_code(r.code)] if include_variants else s[s["exact"] == ex]
        m = m.assign(equiv=[q / ratios.get((r.code, e), 1.0) for q, e in zip(m["qty"], m["exact"])])
        by = m.groupby("staff")["equiv"].sum()
        ach = {l: float(by.get(l, 0.0)) for l in labels}
        total = sum(ach.values())
        use_extra = extra_mode and not math.isnan(r.extra)
        if use_extra:
            target_rm = r.extra * r.profit
            base_q = 0.0 if math.isnan(r.baseline) else r.baseline
            profit_rm = max(total - base_q, 0.0) * r.profit
            # each person gets an equal share of the baseline and must beat it before extra profit counts
            for l, _ in staff:
                p_ach[l] += max(ach[l] - base_q / n, 0.0) * r.profit
        else:
            target_rm = r.target * r.profit
            profit_rm = total * r.profit
            for l, _ in staff:
                p_ach[l] += ach[l] * r.profit
        for l, _ in staff:
            p_target[l] += target_rm / n
        skus.append({
            "code": r.code, "name": r.name, "target": r.target, "ach": ach, "total": total,
            "target_rm": target_rm, "profit_rm": profit_rm,
            "compulsory": bool(r.compulsory) or base_code(r.code) in comp_codes,
        })

    # combine SKUs that are tracked in one column
    groups = [set(split_codes(line)) for line in groups_text.splitlines() if line.strip()]
    columns, used = [], set()
    for i, k in enumerate(skus):
        if i in used:
            continue
        members = [i]
        for g in groups:
            if base_code(k["code"]) in g:
                members = [j for j, o in enumerate(skus) if base_code(o["code"]) in g]
                break
        used.update(members)
        ms = [skus[j] for j in members]
        tot_target = sum(m["target"] for m in ms)
        per = tot_target / n if n else 0.0
        if round_up:
            per = math.ceil(per - 1e-9)
        ach = {l: sum(m["ach"][l] for m in ms) for l in labels}
        t_staff = {l: per for l, _ in staff}
        total_t = per * n
        total_a = sum(ach.values())
        columns.append({
            "code": "/".join(m["code"] for m in ms),
            "name": combined_name([m["name"] for m in ms]) if len(ms) > 1 else ms[0]["name"],
            "compulsory": any(m["compulsory"] for m in ms),
            "target": t_staff, "ach": ach, "total_t": total_t, "total_a": total_a,
            "balance": total_t - total_a,
        })

    summary = {
        "target_rm": sum(k["target_rm"] for k in skus),
        "profit_rm": sum(k["profit_rm"] for k in skus),
        "staff_profit": {l: (p_target[l], p_ach[l]) for l, _ in staff},
    }
    others = sorted(s.loc[s["staff"] == "Others", "Sales Person"].dropna().astype(str).unique())
    return columns, summary, labels, others


def staff_summary(columns, staff_labels, summary, measure="Profit (RM)") -> pd.DataFrame:
    """Per person achievement, by profit (RM) or by quantity (mother packs)."""
    by_profit = measure.startswith("Profit")
    t_col, a_col = ("Profit target (RM)", "Profit achieved (RM)") if by_profit else ("Target qty", "Sold qty")
    rows = []
    for lab in staff_labels:
        if by_profit:
            t, a = summary["staff_profit"][lab]
        else:
            t = sum(c["target"][lab] for c in columns)
            a = sum(c["ach"][lab] for c in columns)
        rows.append({"Staff": lab, t_col: t, a_col: a,
                     "Achievement %": (a / t * 100) if t else 0.0, "Balance": t - a})
    df = pd.DataFrame(rows)
    df.insert(0, "Rank", df["Achievement %"].rank(ascending=False, method="min").astype(int))
    tot_t, tot_a = df[t_col].sum(), df[a_col].sum()
    others = {"Rank": None, "Staff": "Others", t_col: None, a_col: None, "Achievement %": None, "Balance": None}
    if not by_profit:
        others[a_col] = sum(c["ach"]["Others"] for c in columns)
        tot_a += others[a_col]
    total = {"Rank": None, "Staff": "Total", t_col: tot_t, a_col: tot_a,
             "Achievement %": (tot_a / tot_t * 100) if tot_t else 0.0, "Balance": tot_t - tot_a}
    return pd.concat([df, pd.DataFrame([others, total])], ignore_index=True)


def fmt(v, blank_zero=False):
    if v is None or (blank_zero and abs(v) < 1e-9):
        return ""
    return f"{v:,.2f}"


# ------------------------------------------------------------------- excel
def build_workbook(columns, staff_labels, block_size, head, summary=None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "MVP"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    f = Font(name="Arial", size=10)
    fb = Font(name="Arial", size=10, bold=True)
    green = PatternFill("solid", fgColor="B6D7A8")
    grey = PatternFill("solid", fgColor="D9D9D9")
    yellow = PatternFill("solid", fgColor="FFFF00")
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)

    width_cols = 1 + 2 * min(block_size, len(columns))
    ws["A1"] = f"Target = RM{head['target_rm']:,.2f}"
    ws["A1"].font = fb
    for c in range(1, 3):
        ws.cell(row=1, column=c).fill = green
    ws.cell(row=1, column=3, value=f"Update date: {head['dates']}").font = fb
    mid = max(7, width_cols // 2 + 1)
    ws.cell(row=1, column=mid, value=f"EXTRA PROFIT = RM{head['profit_rm']:,.2f}").font = fb
    for c in range(mid, min(mid + 3, width_cols + 1)):
        ws.cell(row=1, column=c).fill = green

    n = len(staff_labels)
    row = 3
    for start in range(0, len(columns), block_size):
        block = columns[start:start + block_size]
        r_sku, r_name, r_hdr = row, row + 1, row + 2
        r_first = row + 3
        r_oth, r_tot, r_bal = r_first + n, r_first + n + 1, r_first + n + 2

        for r in (r_sku, r_name):
            ws.cell(row=r, column=1).border = box
        c = ws.cell(row=r_hdr, column=1, value="Staff Name")
        c.font, c.fill, c.border, c.alignment = fb, grey, box, ctr
        for i, lab in enumerate(staff_labels):
            c = ws.cell(row=r_first + i, column=1, value=lab)
            c.font, c.border = fb, box
        for r, lab in ((r_oth, "Others"), (r_tot, "Total"), (r_bal, "Balance")):
            c = ws.cell(row=r, column=1, value=lab)
            c.font, c.border = fb, box
            c.alignment = Alignment(horizontal="center" if r != r_oth else "left")
            if r == r_tot:
                c.fill = grey

        for i, col in enumerate(block):
            ct, ca = 2 + 2 * i, 3 + 2 * i
            fill = yellow if col["compulsory"] else None
            for r, val in ((r_sku, col["code"]), (r_name, col["name"])):
                ws.merge_cells(start_row=r, start_column=ct, end_row=r, end_column=ca)
                c = ws.cell(row=r, column=ct, value=val)
                c.font, c.alignment = fb, ctr
                for cc in (ct, ca):
                    ws.cell(row=r, column=cc).border = box
                    if fill:
                        ws.cell(row=r, column=cc).fill = fill
            for cc, txt in ((ct, "Target"), (ca, "Achieve")):
                c = ws.cell(row=r_hdr, column=cc, value=txt)
                c.font, c.fill, c.border, c.alignment = fb, grey, box, ctr
            for j, lab in enumerate(staff_labels):
                a = ws.cell(row=r_first + j, column=ct, value=col["target"][lab])
                b = ws.cell(row=r_first + j, column=ca, value=col["ach"][lab] or None)
                for c in (a, b):
                    c.font, c.border, c.number_format = f, box, "0.00"
                    c.alignment = Alignment(horizontal="center")
            a = ws.cell(row=r_oth, column=ct)
            b = ws.cell(row=r_oth, column=ca, value=col["ach"]["Others"] or None)
            for c in (a, b):
                c.font, c.border, c.number_format = f, box, "0.00"
                c.alignment = Alignment(horizontal="center")
            a = ws.cell(row=r_tot, column=ct, value=col["total_t"])
            b = ws.cell(row=r_tot, column=ca, value=col["total_a"])
            for c in (a, b):
                c.font, c.fill, c.border, c.number_format = fb, grey, box, "0.00"
                c.alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=r_bal, start_column=ct, end_row=r_bal, end_column=ca)
            c = ws.cell(row=r_bal, column=ct, value=col["balance"])
            c.font, c.number_format, c.alignment = fb, "0.00", Alignment(horizontal="right")
            for cc in (ct, ca):
                ws.cell(row=r_bal, column=cc).border = box
        row = r_bal + 2

    ws.column_dimensions["A"].width = 18
    for c in range(2, 2 + 2 * block_size):
        ws.column_dimensions[get_column_letter(c)].width = 11
    for r in range(3, row):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = None
    if summary is not None:
        ss = wb.create_sheet("Staff summary")
        heads = list(summary.columns)
        for j, h in enumerate(heads, start=1):
            c = ss.cell(row=1, column=j, value=h)
            c.font, c.fill, c.border, c.alignment = fb, grey, box, ctr
        for i, rec in enumerate(summary.itertuples(index=False), start=2):
            for j, v in enumerate(rec, start=1):
                v = None if (v is None or (isinstance(v, float) and math.isnan(v))) else v
                c = ss.cell(row=i, column=j, value=v)
                c.font = fb if rec[1] in ("Total",) else f
                c.border = box
                if heads[j - 1] == "Achievement %":
                    c.number_format = '0.0"%"'
                elif heads[j - 1] not in ("Rank", "Staff"):
                    c.number_format = "0.00"
        ss.column_dimensions["A"].width = 7
        ss.column_dimensions["B"].width = 18
        for col in "CDEF":
            ss.column_dimensions[col].width = 20
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# -------------------------------------------------------------------- html
def block_html(block, staff_labels) -> str:
    cell = "border:1px solid #999;padding:2px 8px;color:#111;background:#fff;white-space:nowrap;"
    hd = cell + "background:#d9d9d9;font-weight:bold;text-align:center;"
    esc = html.escape
    r = ['<div style="overflow-x:auto"><table style="border-collapse:collapse;font-size:13px">']
    r.append(f'<tr><td style="{cell}"></td>' + "".join(
        f'<td colspan="2" style="{cell}font-weight:bold;text-align:center;'
        f'{"background:#ffff00;" if c["compulsory"] else ""}">{esc(c["code"])}</td>' for c in block) + "</tr>")
    r.append(f'<tr><td style="{cell}"></td>' + "".join(
        f'<td colspan="2" style="{cell}font-weight:bold;text-align:center;white-space:normal;'
        f'{"background:#ffff00;" if c["compulsory"] else ""}">{esc(c["name"])}</td>' for c in block) + "</tr>")
    r.append(f'<tr><td style="{hd}">Staff Name</td>' + "".join(
        f'<td style="{hd}">Target</td><td style="{hd}">Achieve</td>' for _ in block) + "</tr>")
    for lab in staff_labels:
        r.append(f'<tr><td style="{cell}font-weight:bold">{esc(lab)}</td>' + "".join(
            f'<td style="{cell}text-align:center">{fmt(c["target"][lab])}</td>'
            f'<td style="{cell}text-align:center">{fmt(c["ach"][lab], True)}</td>' for c in block) + "</tr>")
    r.append(f'<tr><td style="{cell}font-weight:bold">Others</td>' + "".join(
        f'<td style="{cell}"></td><td style="{cell}text-align:center">{fmt(c["ach"]["Others"], True)}</td>'
        for c in block) + "</tr>")
    r.append(f'<tr><td style="{hd}">Total</td>' + "".join(
        f'<td style="{hd}">{fmt(c["total_t"])}</td><td style="{hd}">{fmt(c["total_a"])}</td>' for c in block) + "</tr>")
    r.append(f'<tr><td style="{cell}font-weight:bold;text-align:center">Balance</td>' + "".join(
        f'<td colspan="2" style="{cell}font-weight:bold;text-align:right">{fmt(c["balance"])}</td>' for c in block)
        + "</tr></table></div>")
    return "".join(r)


# --------------------------------------------------------------------- UI
st.title("🏆 MVP")
st.caption("Upload the MVP SKU list and the sale invoice item export. The page shares each target equally "
           "between the staff and fills in their sales.")

with st.sidebar:
    st.header("Settings")
    block_size = st.number_input("Items per block (across the sheet)", min_value=1, max_value=20, value=8, step=1)
    round_up = st.checkbox("Round each staff's target up to a whole unit", False,
                           help="Off: the target is shared exactly (for example 4 ÷ 7 = 0.57 each).")
    groups_text = st.text_area("Show these SKUs in one column (one group per line)", DEFAULT_GROUPS, height=70)
    comp_text = st.text_area("Compulsory SKU codes (header turns yellow)", DEFAULT_COMPULSORY, height=70,
                             help="A Compulsory column in the MVP file works too.")
    include_variants = st.checkbox("Count loose packs (T / L codes) as part of the mother SKU", True,
                                   help="Example: T50074 (1S) and L100290576 (10S) count for 50074 and 100290576. "
                                        "Loose quantities are converted to mother packs, and profit follows the "
                                        "mother pack profit.")
    measure = st.selectbox("Staff achievement by", ["Profit (RM)", "Quantity"],
                           help="Profit: each person's share of the profit target against the profit they earned. "
                                "Quantity: total target qty against total qty sold.")
    extra_mode = st.checkbox("Use Extra Qty and Baseline for the RM figures", True,
                             help="Same as your calculator: Target RM = extra qty × profit. Extra profit = "
                                  "(sales − baseline) × profit, never below zero. Needs the Extra Qty and Baseline "
                                  "columns. Off: Target RM = target × profit, Extra profit = sales × profit.")

tpl = io.BytesIO()
pd.DataFrame(TEMPLATE_ROWS, columns=TEMPLATE_COLUMNS).to_excel(tpl, index=False, sheet_name="MVP SKU")

c1, c2 = st.columns(2)
with c1:
    mvp_file = st.file_uploader("MVP SKU list (.xlsx)", type=["xlsx"],
                                help="Item code, item name, profit per unit and monthly target.")
    st.download_button("⬇️ MVP SKU template (filled from your calculator)", tpl.getvalue(),
                       file_name="MVP_SKU_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with c2:
    sales_file = st.file_uploader("Sale invoice item export (.xlsx)", type=["xlsx"])

if mvp_file is None or sales_file is None:
    st.info("Upload both files to begin.")
    st.stop()

try:
    mvp = read_mvp(mvp_file)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read the MVP file: {e}")
    st.stop()
if mvp.empty:
    st.error("No SKU rows found in the MVP file.")
    st.stop()

try:
    sales = pd.read_excel(sales_file, dtype={"Item Code": str})
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read the sales file: {e}")
    st.stop()
sales.columns = [str(c).strip() for c in sales.columns]
missing = [c for c in REQUIRED_SALES if c not in sales.columns]
if missing:
    st.error(f"Missing column(s) in the sales file: {', '.join(missing)}. Please upload the original export.")
    st.stop()

staff_df = detect_staff(sales)
with st.expander("Staff (found in the sales file)", expanded=True):
    st.caption("Tick the people who share the targets. Change Order to match your sheet, and edit Name on sheet. "
               "Everyone not ticked goes to the Others row.")
    edited = st.data_editor(
        staff_df, hide_index=True, use_container_width=True,
        disabled=["Name in sales file", "Sales lines"],
        column_config={
            "Use": st.column_config.CheckboxColumn("Use"),
            "Order": st.column_config.NumberColumn("Order", min_value=1, step=1, format="%d"),
            "Name on sheet": st.column_config.TextColumn("Name on sheet"),
        },
        key="staff_" + str(abs(hash(tuple(staff_df["Name in sales file"])))),
    )
    if st.button("Save this staff list for next time"):
        try:
            STAFF_FILE.write_text(json.dumps({
                "labels": {n.upper(): l.strip().upper() for n, l in zip(edited["Name in sales file"], edited["Name on sheet"])},
                "known": [n.upper() for n in edited["Name in sales file"]],
                "selected": [n.upper() for n, u in zip(edited["Name in sales file"], edited["Use"]) if u],
                "order": {n.upper(): int(o) for n, o in zip(edited["Name in sales file"], edited["Order"])},
            }, indent=1), encoding="utf-8")
            st.success("Saved for everyone who uses this page.")
        except OSError as e:
            st.error(f"Could not save: {e}")

chosen = edited[edited["Use"] & (edited["Name on sheet"].astype(str).str.strip() != "")].sort_values(
    "Order", kind="stable")
staff = [(str(lab).strip().upper(), [str(full).strip().upper()])
         for lab, full in zip(chosen["Name on sheet"], chosen["Name in sales file"])]
if not staff:
    st.error("Tick at least one staff in the list above.")
    st.stop()
if len({l for l, _ in staff}) < len(staff):
    st.error("Two staff have the same Name on sheet. Please give each a different name.")
    st.stop()
st.caption(f"{len(staff)} staff share the targets: " + ", ".join(l for l, _ in staff) + ".")

s_prep = prepare_sales(sales, staff)
matches = build_matches(mvp, s_prep, include_variants)

ratios = {}
with st.expander("Loose pack conversion: check this first", expanded=not matches.empty and
                 bool((matches["Note"] != "mother code").any())):
    if matches.empty:
        st.warning("No sales matched any SKU in the MVP file. Check the item codes.")
    else:
        st.caption("Your MVP profit is per mother pack. A loose pack (strip or single unit) is converted to mother "
                   "packs before it is counted, and its profit is the mother profit divided by the same number. "
                   "The number is read from the item names. Correct it here if it is wrong.")
        shown = matches.drop(columns=["exact"])
        edited = st.data_editor(
            shown, hide_index=True, use_container_width=True,
            disabled=[c for c in shown.columns if c != "Loose packs per mother pack"],
            column_config={"Loose packs per mother pack": st.column_config.NumberColumn(min_value=0.0001, format="%.2f")},
            key=f"ratio_{include_variants}_{len(matches)}_{hash(tuple(matches['Sales code']))}",
        )
        profit_of = dict(zip(mvp["code"], mvp["profit"]))
        def _ratio(v):
            return 1.0 if pd.isna(v) or v <= 0 else float(v)

        ratios = {(r["MVP SKU"], matches.loc[i, "exact"]): _ratio(r["Loose packs per mother pack"])
                  for i, r in edited.iterrows()}
        view = edited[["MVP SKU", "Sales code", "Sales item name", "Qty sold", "Loose packs per mother pack"]].copy()
        view["Mother packs counted"] = view["Qty sold"] / view["Loose packs per mother pack"]
        view["Profit per mother pack"] = view["MVP SKU"].map(profit_of)
        view["Profit per loose pack"] = view["Profit per mother pack"] / view["Loose packs per mother pack"]
        view["Profit counted"] = view["Mother packs counted"] * view["Profit per mother pack"]
        st.dataframe(view.round(2), hide_index=True, use_container_width=True)

columns, summary, labels, others = compute(
    mvp, s_prep, staff, include_variants, groups_text, comp_text, round_up, extra_mode, ratios)

dates = pd.to_datetime(sales["Sales Date"], errors="coerce").dropna()
date_txt = f"{dates.min():%d/%m/%Y} - {dates.max():%d/%m/%Y}" if len(dates) else "-"
head = {"target_rm": summary["target_rm"], "profit_rm": summary["profit_rm"], "dates": date_txt}
staff_labels = [l for l, _ in staff]

comp_cols = [c for c in columns if c["compulsory"]]
comp_met = sum(1 for c in comp_cols if c["total_a"] >= c["total_t"] - 1e-9)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Target", f"RM{head['target_rm']:,.2f}")
m2.metric("Extra profit", f"RM{head['profit_rm']:,.2f}")
m3.metric("Update date", date_txt)
m4.metric("Compulsory SKUs met", f"{comp_met} of {len(comp_cols)}" if comp_cols else "-")

st.subheader("MVP sheet")
st.caption("Target and Achieve are in mother packs. Select a table and copy it, or download the Excel file. "
           "All cells are plain values, so they can be pasted into your sheet. The staff achievement table is on a second sheet of the Excel file.")
for start in range(0, len(columns), int(block_size)):
    st.markdown(block_html(columns[start:start + int(block_size)], staff_labels), unsafe_allow_html=True)
    st.write("")

summary_df = staff_summary(columns, staff_labels, summary, measure)
st.subheader("Staff achievement")
if measure.startswith("Profit"):
    st.caption("The profit target is shared equally between the staff. Each person's profit is counted per SKU, "
               "using mother pack profit. "
               + ("Each person gets an equal share of the baseline, and only sales above it earn profit, "
                  "the same way the target is split. " if extra_mode else "")
               + "Sales under Others are not counted for anyone.")
else:
    st.caption("Each person's total target qty against the total qty they sold, whatever the item. "
               "Quantities are in mother packs. Sales above one item's target still count toward the total.")
num_cfg = {c: st.column_config.NumberColumn(format="%.2f") for c in summary_df.columns
           if c not in ("Rank", "Staff", "Achievement %")}
num_cfg["Achievement %"] = st.column_config.ProgressColumn(
    format="%.1f%%", min_value=0, max_value=max(100.0, float(summary_df["Achievement %"].max() or 0)))
st.dataframe(summary_df, hide_index=True, use_container_width=True, column_config=num_cfg)

xlsx = build_workbook(columns, staff_labels, int(block_size), head, summary_df)
st.download_button("⬇️ Download MVP sheet (.xlsx)", xlsx,
                   file_name=f"MVP_{datetime.now():%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

with st.expander("More checks"):
    no_sales = [c["code"] for c in columns if c["total_a"] == 0]
    if no_sales:
        st.caption("No sales found for: " + ", ".join(no_sales))
    if others:
        st.caption("Counted under Others: " + ", ".join(others))
    st.caption("Target RM = " + ("extra qty × profit per mother pack" if extra_mode else "target × profit per mother pack")
               + ". Extra profit = " + ("(sales − baseline) × profit per mother pack, never below zero."
                                        if extra_mode else "sales × profit per mother pack."))
