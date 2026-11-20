"""Sale Performance.

Staff upload the "Sales Report - By Store, Sales Person" export (all sales) and the same report for private
label items. The page keeps only the outlet (QKCMC1), puts each salesperson's numbers into the personal
performance table, ranks the staff for MVP, and draws charts to compare them.

Targets are set for the whole outlet in the sidebar (sale target and profit target) and shared between the
staff by role ratio, for example pharmacist 1.5 and healthcare assistant 1.0. Everything in the sidebar can be
saved for next time.
"""
import calendar
import html
import io
import json
import re
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="Sale Performance", page_icon="📈", layout="wide")

# Starting staff list (full name in the report, name on the sheet, role), in sheet order.
# The list of names is shared with the MVP page. Anyone else found in the report is worked out by the page.
SEED_STAFF = [
    ("ELIZABETH TSAI YEN PENG", "ELIZABETH", "Pharmacist"),
    ("LIM JUN MING", "LIM JUN MING", "Pharmacist"),
    ("ELIZZA ANAK VALENTINE", "ELIZZA", "HA"),
    ("NUR SYAZWANI BINTI REDI", "SYAZWANI", "HA"),
    ("ELLYTHIA NISHA ANAK SIGON", "ELLYTHIA", "HA"),
    ("NURSYUHADA BINTI RIZALMAN", "SYUHADA", "HA"),
    ("MELVIN CRIST ANAK ANOM", "MELL", "HA"),
]
ROLES = ["Pharmacist", "HA"]
APP_DIR = Path(__file__).resolve().parent.parent
STAFF_FILE = APP_DIR / "mvp_staff.json"
SETTINGS_FILE = APP_DIR / "performance_settings.json"
NOT_STAFF = ("PHARMACY", "ALPRO", "ADMIN", "SYSTEM", "CASH", "ONLINE")
REQUIRED = ["Store", "Sales Person", "No. Transaction", "Net Sales", "GP"]
CROWN = "\U0001F451"
COMPONENTS = ["Sales", "GP", "PL", "PP"]

# Starting values of the sidebar. Change them in the sidebar and press Save settings for next time.
DEFAULTS = {
    "outlet": "QKCMC1", "sale_target": 350000.0, "profit_target": 115000.0, "basis": "Profit (GP)",
    "ratio_pharmacist": 1.5, "ratio_ha": 1.0,
    "w_sales": 25, "w_gp": 25, "w_pl": 25, "w_pp": 25,
    "pl_target": 40.0, "pp_target": 50.0, "cap": False, "prorate": True,
}


# ---------------------------------------------------------------- reading
def read_report(file):
    """Return (rows, period) from a 'Sales Report - By Store, Sales Person' export."""
    raw = pd.read_excel(file, header=None, dtype=str)
    period = None
    for i in range(min(20, len(raw))):
        if str(raw.iloc[i, 0]).strip().lower().startswith("sales date"):
            found = re.findall(r"\d{4}-\d{2}-\d{2}", str(raw.iloc[i, 1]))
            if len(found) >= 2:
                period = (pd.Timestamp(found[0]), pd.Timestamp(found[1]))
    hdr = None
    for i in range(min(40, len(raw))):
        cells = [str(x).strip() for x in raw.iloc[i].fillna("")]
        if "Store" in cells and "Sales Person" in cells:
            hdr = i
            break
    if hdr is None:
        raise ValueError("Could not find the header row with Store and Sales Person.")
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(x).strip() for x in raw.iloc[hdr].fillna("")]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Missing column(s): " + ", ".join(missing))
    df = df[df["Store"].notna()].copy()  # the Total line has no store
    for c in ("No. Transaction", "Net Sales", "GP"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["person"] = df["Sales Person"].fillna("").astype(str).str.strip().str.upper()
    return df[["Store", "person", "No. Transaction", "Net Sales", "GP"]], period


def outlet_rows(df: pd.DataFrame, code: str) -> pd.DataFrame:
    return df[df["Store"].astype(str).str.upper().str.contains(code.strip().upper(), regex=False)].copy()


# ------------------------------------------------------------------ settings
def _load_json(path: Path) -> dict:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    return {}


def load_settings() -> dict:
    saved = _load_json(SETTINGS_FILE)
    return {**DEFAULTS, **{k: saved[k] for k in DEFAULTS if k in saved}}


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


def detect_staff(people: pd.DataFrame) -> pd.DataFrame:
    """people: columns person (full name, upper case) and tx. Guess who is staff, and fill names and roles."""
    people = people[people["person"] != ""].sort_values("tx", ascending=False)
    saved = _load_json(STAFF_FILE)
    roles = _load_json(SETTINGS_FILE).get("roles") or {f.upper(): r for f, _, r in SEED_STAFF}
    if saved:
        labels, known = saved.get("labels", {}), set(saved.get("known", []))
        selected, order = set(saved.get("selected", [])), saved.get("order", {})
    else:
        labels = {f.upper(): l for f, l, _ in SEED_STAFF}
        known = set(labels)
        selected = set(labels)
        order = {f.upper(): i + 1 for i, (f, _, _) in enumerate(SEED_STAFF)}
    top = float(people["tx"].max()) if len(people) else 0.0
    rows = []
    for r in people.itertuples():
        if r.person in known:
            use = r.person in selected
        else:  # new name: staff if they serve a fair share of what the busiest person serves
            use = r.tx >= 0.25 * top and not any(w in r.person for w in NOT_STAFF)
        role = roles.get(r.person, "HA")
        rows.append({"Use": use, "Order": order.get(r.person, 100 + len(rows)), "Name in report": r.person,
                     "Name on sheet": labels.get(r.person) or default_label(r.person),
                     "Role": role if role in ROLES else "HA", "Transactions": int(r.tx)})
    df = pd.DataFrame(rows)
    return df.sort_values(["Use", "Order", "Transactions"], ascending=[False, True, False]).reset_index(drop=True)


def save_everything(settings: dict, staff_table: pd.DataFrame) -> None:
    """Sidebar settings and the staff table (names, order, roles) for next time."""
    SETTINGS_FILE.write_text(json.dumps({
        **settings, "roles": {n: r for n, r in zip(staff_table["Name in report"], staff_table["Role"])},
    }, indent=1), encoding="utf-8")
    STAFF_FILE.write_text(json.dumps({
        "labels": {n: str(l).strip().upper() for n, l in zip(staff_table["Name in report"], staff_table["Name on sheet"])},
        "known": list(staff_table["Name in report"]),
        "selected": [n for n, u in zip(staff_table["Name in report"], staff_table["Use"]) if u],
        "order": {n: int(o) for n, o in zip(staff_table["Name in report"], staff_table["Order"])},
    }, indent=1), encoding="utf-8")


def share_targets(roles: list, ratios: dict, total: float) -> list:
    """Each person's part of a whole-outlet target: total x ratio / sum of ratios."""
    r = [ratios.get(x, 1.0) for x in roles]
    return [total * x / sum(r) for x in r] if sum(r) > 0 else [0.0] * len(r)


# ------------------------------------------------------------ calculation
def build_table(sales: pd.DataFrame, pl: pd.DataFrame, staff: list, basis_gp: bool) -> pd.DataFrame:
    """staff: list of (label, full name, target). Returns one row per person, then Others and Total."""
    s_by = sales.groupby("person")[["No. Transaction", "Net Sales", "GP"]].sum()
    p_by = pl.groupby("person")["Net Sales"].sum()

    def row(label, names, target):
        sub = s_by[s_by.index.isin(names)]
        sales_amt, gp, tx = sub["Net Sales"].sum(), sub["GP"].sum(), sub["No. Transaction"].sum()
        pl_amt = p_by[p_by.index.isin(names)].sum()
        return {"Staff": label, "Target": target, "Sales": sales_amt, "GP": gp, "Tx": tx, "PL": pl_amt}

    chosen = {full for _, full, _ in staff}
    rows = [row(label, [full], target) for label, full, target in staff]
    others = [p for p in set(s_by.index) | set(p_by.index) if p not in chosen]
    rows.append(row("Others", others, None))
    df = pd.DataFrame(rows)
    total = {"Staff": "Total", "Target": df["Target"].sum(min_count=1), "Sales": df["Sales"].sum(),
             "GP": df["GP"].sum(), "Tx": df["Tx"].sum(), "PL": df["PL"].sum()}
    df = pd.concat([df, pd.DataFrame([total])], ignore_index=True)
    df["Achieve"] = df["GP"] if basis_gp else df["Sales"]
    df["Balance"] = df["Target"] - df["Achieve"]
    df["Pct"] = df["Achieve"] / df["Target"]
    df["GPpct"] = df["GP"] / df["Sales"].where(df["Sales"] != 0)
    df["PP"] = df["Sales"] / df["Tx"].where(df["Tx"] != 0)
    df["PLpct"] = df["PL"] / df["Sales"].where(df["Sales"] != 0)
    df.loc[df["Staff"] == "Others", ["Balance", "Pct"]] = float("nan")
    return df


def pace_of_month(period):
    """Share of the month that has passed at the end of the report, when the report starts on the 1st."""
    if not period or period[0].day != 1:
        return None
    return period[1].day / calendar.monthrange(period[1].year, period[1].month)[1]


# ------------------------------------------------------------------ output
def rm(v, decimals=2):
    return "" if pd.isna(v) else f"RM{v:,.{decimals}f}"


def pct(v):
    return "" if pd.isna(v) else f"{v * 100:.2f}%"


def num(v, d=2):
    return "" if pd.isna(v) else f"{v:,.{d}f}"


def table_html(df: pd.DataFrame, head: dict) -> str:
    esc = html.escape
    base = "border:1px solid #bbb;padding:3px 8px;color:#111;background:#fff;font-size:13px;white-space:nowrap;"
    hd = base + "background:#d9d9d9;font-weight:bold;text-align:center;"
    right = base + "text-align:right;"
    out = ['<div style="overflow-x:auto"><table style="border-collapse:collapse">']
    out.append(f'<tr><td colspan="2" style="{base}font-weight:bold">{esc(head["target_text"])}</td>'
               f'<td colspan="7" style="{hd}">Update date: {esc(head["dates"])}</td>'
               f'<td colspan="2" style="{hd}">Private Label</td></tr>')
    cols = ["Staff Name", "Target", "Achieve", "Balance", "% Achieve", "Sales", "GP%", "Tx", "PP", "PL", "PL%"]
    out.append("<tr>" + "".join(f'<td style="{hd}">{c}</td>' for c in cols) + "</tr>")
    for r in df.itertuples():
        bold = "font-weight:bold;" if r.Staff == "Total" else ""
        bg = "background:#e6e6e6;" if r.Staff == "Total" else ""
        cells = [r.Staff, rm(r.Target), rm(r.Achieve), rm(r.Balance), pct(r.Pct), rm(r.Sales), pct(r.GPpct),
                 num(r.Tx, 0), num(r.PP), num(r.PL), pct(r.PLpct)]
        out.append(f'<tr><td style="{base}font-weight:bold;{bg}">{esc(str(cells[0]))}</td>' + "".join(
            f'<td style="{right}{bold}{bg}">{esc(c)}</td>' for c in cells[1:]) + "</tr>")
    out.append("</table></div>")
    return "".join(out)


def build_workbook(df: pd.DataFrame, head: dict, ranking: pd.DataFrame = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Personal performance"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    grey = PatternFill("solid", fgColor="D9D9D9")
    bold = Font(name="Arial", size=10, bold=True)
    norm = Font(name="Arial", size=10)
    ctr = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:B1")
    ws["A1"] = head["target_text"]
    ws["A1"].font = bold
    ws.merge_cells("C1:I1")
    ws["C1"] = f"Update date: {head['dates']}"
    ws["C1"].font, ws["C1"].alignment = bold, ctr
    ws.merge_cells("J1:K1")
    ws["J1"] = "Private Label"
    ws["J1"].font, ws["J1"].alignment = bold, ctr
    for c in range(1, 12):
        ws.cell(row=1, column=c).fill = grey
        ws.cell(row=1, column=c).border = box
    cols = ["Staff Name", "Target", "Achieve", "Balance", "% Achieve", "Sales", "GP%", "Tx", "PP", "PL", "PL%"]
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=2, column=j, value=c)
        cell.font, cell.fill, cell.border, cell.alignment = bold, grey, box, ctr
    fmts = {2: '"RM"#,##0.00', 3: '"RM"#,##0.00', 4: '"RM"#,##0.00', 5: "0.00%", 6: '"RM"#,##0.00',
            7: "0.00%", 8: "0", 9: "0.00", 10: "#,##0.00", 11: "0.00%"}
    for i, r in enumerate(df.itertuples(), start=3):
        vals = [r.Staff, r.Target, r.Achieve, r.Balance, r.Pct, r.Sales, r.GPpct, r.Tx, r.PP, r.PL, r.PLpct]
        for j, v in enumerate(vals, start=1):
            v = None if (isinstance(v, float) and pd.isna(v)) else v
            cell = ws.cell(row=i, column=j, value=v)
            cell.border = box
            cell.font = bold if (j == 1 or r.Staff == "Total") else norm
            if j in fmts:
                cell.number_format = fmts[j]
                cell.alignment = Alignment(horizontal="right")
            if r.Staff == "Total":
                cell.fill = grey
    ws.column_dimensions["A"].width = 16
    for j in range(2, 12):
        ws.column_dimensions[chr(64 + j)].width = 15
    if ranking is not None:
        add_ranking_sheet(wb, ranking)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------ MVP ranking
def build_ranking(viz: pd.DataFrame, targets: dict, weights: dict, pl_target: float, pp_target: float,
                  cap: bool, factor: float) -> pd.DataFrame:
    """Points = achievement against target x weight, for sales, GP, private label % and PP. Rank 1 is the MVP.
    targets: label -> (sale target, profit target). factor scales the RM targets (share of the month so far)."""
    total_w = sum(weights.values())
    rows = []
    for r in viz.itertuples():
        t_sale, t_profit = targets[r.Staff]
        actual = {"Sales": r.Sales, "GP": r.GP, "PL": (0.0 if pd.isna(r.PLpct) else r.PLpct * 100),
                  "PP": (0.0 if pd.isna(r.PP) else r.PP)}
        target = {"Sales": t_sale * factor, "GP": t_profit * factor, "PL": pl_target, "PP": pp_target}
        rec = {"Staff": r.Staff}
        total = 0.0
        for k in COMPONENTS:
            ach = actual[k] / target[k] if target[k] > 0 else 0.0
            ach = max(ach, 0.0)
            ach = min(ach, 1.0) if cap else ach
            pts = ach * weights[k] / total_w * 100
            rec[k], rec[k + " target"], rec[k + " ach"], rec[k + " points"] = actual[k], target[k], ach * 100, pts
            total += pts
        rec["Total"] = total
        rows.append(rec)
    out = pd.DataFrame(rows)
    out["Rank"] = out["Total"].rank(ascending=False, method="min").astype(int)
    out["MVP"] = out["Rank"] == 1
    out["Name"] = [f"{CROWN} {n}" if m else n for n, m in zip(out["Staff"], out["MVP"])]
    return out.sort_values(["Rank", "Staff"]).reset_index(drop=True)


MEASURE_NAME = {"Sales": "Sales (RM)", "GP": "GP (RM)", "PL": "Private label %", "PP": "PP (RM)"}


def _fmt_measure(k: str, v: float) -> str:
    return f"{v:,.2f}%" if k == "PL" else f"{v:,.2f}"


def ranking_html(rk: pd.DataFrame, weights: dict) -> str:
    esc = html.escape
    base = "border:1px solid #bbb;padding:3px 8px;color:#111;background:#fff;font-size:13px;white-space:nowrap;"
    hd = base + "background:#d9d9d9;font-weight:bold;text-align:center;"
    total_w = sum(weights.values())
    out = [f'<div style="overflow-x:auto"><table style="border-collapse:collapse"><tr>'
           f'<td style="{hd}" rowspan="2">Rank</td><td style="{hd}" rowspan="2">Staff</td>']
    for k in COMPONENTS:
        out.append(f'<td style="{hd}" colspan="4">{MEASURE_NAME[k]} ({weights[k] / total_w * 100:.0f}%)</td>')
    out.append(f'<td style="{hd}" rowspan="2">Total points</td></tr><tr>')
    out.append("".join(f'<td style="{hd}">Actual</td><td style="{hd}">Target</td><td style="{hd}">Achieve %</td>'
                       f'<td style="{hd}">Points</td>' for _ in COMPONENTS) + "</tr>")
    for _, r in rk.iterrows():
        bg = "background:#fff4c2;" if r["MVP"] else ""
        cells = "".join(
            f'<td style="{base}text-align:right;{bg}">{esc(_fmt_measure(k, r[k]))}</td>'
            f'<td style="{base}text-align:right;{bg}">{esc(_fmt_measure(k, r[k + " target"]))}</td>'
            f'<td style="{base}text-align:right;{bg}">{r[k + " ach"]:.1f}%</td>'
            f'<td style="{base}text-align:right;{bg}">{r[k + " points"]:.1f}</td>' for k in COMPONENTS)
        out.append(f'<tr><td style="{base}text-align:center;font-weight:bold;{bg}">{r["Rank"]}</td>'
                   f'<td style="{base}font-weight:bold;{bg}">{esc(r["Name"])}</td>{cells}'
                   f'<td style="{base}text-align:right;font-weight:bold;{bg}">{r["Total"]:.1f}</td></tr>')
    out.append("</table></div>")
    return "".join(out)


def add_ranking_sheet(wb, rk: pd.DataFrame) -> None:
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    grey = PatternFill("solid", fgColor="D9D9D9")
    bold = Font(name="Arial", size=10, bold=True)
    norm = Font(name="Arial", size=10)
    ws = wb.create_sheet("MVP ranking")
    cols = ["Rank", "Staff"] + [f"{MEASURE_NAME[k]} {x}" for k in COMPONENTS
                                for x in ("actual", "target", "achieve %", "points")] + ["Total points"]
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=j, value=c)
        cell.font, cell.fill, cell.border = bold, grey, box
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, (_, r) in enumerate(rk.iterrows(), start=2):
        vals = [int(r["Rank"]), f"{CROWN} {r['Staff']}" if r["MVP"] else r["Staff"]]
        for k in COMPONENTS:
            vals += [float(r[k]), float(r[k + " target"]), float(r[k + " ach"]), float(r[k + " points"])]
        vals.append(float(r["Total"]))
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.border, cell.font = box, bold if j <= 2 else norm
            if j > 2:
                cell.number_format = "#,##0.00"
    ws.column_dimensions["B"].width = 18
    ws.row_dimensions[1].height = 45


def ranking_charts(rk: pd.DataFrame, weights: dict):
    total_w = sum(weights.values())
    comps = [f"{MEASURE_NAME[k]} ({weights[k] / total_w * 100:.0f}%)" for k in COMPONENTS]
    order = rk["Name"].tolist()
    long = pd.concat([pd.DataFrame({"Name": rk["Name"], "Component": c, "Points": rk[k + " points"]})
                      for k, c in zip(COMPONENTS, comps)])
    stack = alt.Chart(long).mark_bar().encode(
        y=alt.Y("Name:N", sort=order, title=None),
        x=alt.X("Points:Q", title="Points (100 = every target reached)", stack="zero"),
        color=alt.Color("Component:N", scale=alt.Scale(domain=comps, range=["#3b82c4", "#2e9e5b", "#7a6fc4", "#e08a2e"]),
                        legend=alt.Legend(title=None, orient="bottom", columns=2)),
        tooltip=["Name", "Component", alt.Tooltip("Points:Q", format=".1f")])
    tot = alt.Chart(rk).mark_text(align="left", dx=5, fontSize=13, fontWeight="bold").encode(
        y=alt.Y("Name:N", sort=order), x="Total:Q", text=alt.Text("Total:Q", format=".1f"))
    c1 = (stack + tot).properties(height=320, title="MVP ranking: points from achievement against target")

    heat_long = pd.concat([pd.DataFrame({"Name": rk["Name"], "Measure": c, "Ach": rk[k + " ach"]})
                           for k, c in zip(COMPONENTS, comps)])
    heat = alt.Chart(heat_long).mark_rect().encode(
        y=alt.Y("Name:N", sort=order, title=None),
        x=alt.X("Measure:N", sort=comps, title=None, axis=alt.Axis(labelAngle=-20, labelLimit=170)),
        color=alt.Color("Ach:Q", scale=alt.Scale(domain=[0, 100], range=["#fbe3c8", "#2e9e5b"], clamp=True), legend=None),
        tooltip=["Name", "Measure", alt.Tooltip("Ach:Q", format=".1f", title="Achieve %")])
    txt = alt.Chart(heat_long).mark_text(fontSize=12).encode(
        y=alt.Y("Name:N", sort=order), x=alt.X("Measure:N", sort=comps), text=alt.Text("Ach:Q", format=".0f"))
    c2 = (heat + txt).properties(height=320, title="Achievement % against target for each measure")
    return c1, c2


# ------------------------------------------------------------------ charts
def charts(viz: pd.DataFrame, pace, basis_name: str):
    order = viz.sort_values("Pct", ascending=False)["Staff"].tolist()
    v = viz.assign(PctNum=viz["Pct"] * 100)

    def status(p):
        if p >= 100:
            return "Target reached"
        if pace is not None and p >= pace * 100:
            return "On pace"
        return "Behind"

    v["Status"] = v["PctNum"].map(status)
    colors = alt.Scale(domain=["Target reached", "On pace", "Behind"], range=["#2e9e5b", "#3b82c4", "#e08a2e"])
    bars = alt.Chart(v).mark_bar().encode(
        y=alt.Y("Staff:N", sort=order, title=None),
        x=alt.X("PctNum:Q", title=f"% of {basis_name} target achieved"),
        color=alt.Color("Status:N", scale=colors, legend=alt.Legend(title=None, orient="bottom")),
        tooltip=["Staff", alt.Tooltip("PctNum:Q", format=".1f", title="% achieved"),
                 alt.Tooltip("Achieve:Q", format=",.2f"), alt.Tooltip("Target:Q", format=",.2f")])
    labels = alt.Chart(v).mark_text(align="left", dx=4).encode(
        y=alt.Y("Staff:N", sort=order), x="PctNum:Q", text=alt.Text("PctNum:Q", format=".1f"))
    rules = [alt.Chart(pd.DataFrame({"x": [100]})).mark_rule(color="#888", strokeDash=[4, 3]).encode(x="x:Q")]
    if pace is not None:
        rules.append(alt.Chart(pd.DataFrame({"x": [pace * 100]})).mark_rule(color="#3b82c4", strokeDash=[2, 2])
                     .encode(x="x:Q"))
    c1 = alt.layer(bars, labels, *rules).properties(height=300, title="Progress to target (dashed blue line = share of the month passed)"
                                                    if pace is not None else "Progress to target")

    long = pd.concat([
        v[["Staff", "Achieve"]].rename(columns={"Achieve": "RM"}).assign(Part="Achieved"),
        v.assign(RM=v["Balance"].clip(lower=0))[["Staff", "RM"]].assign(Part="Still to go"),
    ])
    c2 = alt.Chart(long).mark_bar().encode(
        y=alt.Y("Staff:N", sort=order, title=None),
        x=alt.X("RM:Q", title="RM", stack="zero"),
        color=alt.Color("Part:N", scale=alt.Scale(domain=["Achieved", "Still to go"], range=["#2e9e5b", "#d9d9d9"]),
                        legend=alt.Legend(title=None, orient="bottom")),
        tooltip=["Staff", "Part", alt.Tooltip("RM:Q", format=",.2f")],
    ).properties(height=300, title=f"{basis_name.capitalize()} achieved and still to go (RM)")

    c3 = alt.Chart(v.assign(GP=v["GPpct"] * 100, PLs=v["PLpct"] * 100)).mark_circle(opacity=0.85).encode(
        x=alt.X("GP:Q", title="GP %", scale=alt.Scale(zero=False)),
        y=alt.Y("PLs:Q", title="Private label share of sales %", scale=alt.Scale(zero=False)),
        size=alt.Size("Sales:Q", legend=None, scale=alt.Scale(range=[200, 1400])),
        color=alt.Color("Staff:N", legend=None),
        tooltip=["Staff", alt.Tooltip("GP:Q", format=".1f"), alt.Tooltip("PLs:Q", format=".1f", title="PL %"),
                 alt.Tooltip("Sales:Q", format=",.2f")])
    c3 = (c3 + alt.Chart(v.assign(GP=v["GPpct"] * 100, PLs=v["PLpct"] * 100)).mark_text(dy=-18).encode(
        x="GP:Q", y="PLs:Q", text="Staff:N")).properties(height=300, title="Margin and private label (bubble size = sales)")

    c4 = alt.Chart(v).mark_bar().encode(
        y=alt.Y("Staff:N", sort=order, title=None),
        x=alt.X("PP:Q", title="Average sale per transaction (RM)"),
        color=alt.value("#7a6fc4"),
        tooltip=["Staff", alt.Tooltip("PP:Q", format=",.2f"), alt.Tooltip("Tx:Q", format=",.0f", title="Transactions"),
                 alt.Tooltip("Sales:Q", format=",.2f")])
    t4 = alt.Chart(v).mark_text(align="left", dx=4).encode(
        y=alt.Y("Staff:N", sort=order), x="PP:Q", text=alt.Text("PP:Q", format=".2f"))
    c4 = (c4 + t4).properties(height=300, title="Average sale per transaction")
    return c1, c2, c3, c4


# --------------------------------------------------------------------- UI
st.title("📈 Sale Performance")
st.caption("Upload the sales report and the private label report. The page keeps only your outlet and fills in the "
           "personal performance table and the MVP ranking.")

S0 = load_settings()
for _k, _v in S0.items():
    st.session_state.setdefault(f"set_{_k}", _v)

with st.sidebar:
    st.header("Settings")
    outlet = st.text_input("Outlet code", key="set_outlet", help="Only rows of this store are used.")

    st.subheader("Targets for the whole outlet")
    sale_target = st.number_input("Sale target (RM)", min_value=0.0, step=1000.0, format="%.2f", key="set_sale_target")
    profit_target = st.number_input("Profit target (RM)", min_value=0.0, step=1000.0, format="%.2f",
                                    key="set_profit_target", help="The profit is the GP.")
    basis = st.radio("The table shows the", ["Profit (GP)", "Sales"], key="set_basis", horizontal=True,
                     help="Which target and which achievement go in the Target and Achieve columns.")

    st.subheader("Share by role (ratio)")
    ratio_pharm = st.number_input("Pharmacist", min_value=0.0, step=0.1, format="%.2f", key="set_ratio_pharmacist")
    ratio_ha = st.number_input("HA", min_value=0.0, step=0.1, format="%.2f", key="set_ratio_ha")
    st.caption("A person's target = total target x their ratio / the ratios of everyone shown. "
               "Set each person's role in the staff table.")

    st.subheader("MVP ranking")
    st.caption("Points = achievement against target x weight.")
    w_sales = st.number_input("Weight: sales (%)", min_value=0, max_value=100, step=5, key="set_w_sales")
    w_gp = st.number_input("Weight: GP (%)", min_value=0, max_value=100, step=5, key="set_w_gp")
    w_pl = st.number_input("Weight: private label (%)", min_value=0, max_value=100, step=5, key="set_w_pl")
    w_pp = st.number_input("Weight: PP (%)", min_value=0, max_value=100, step=5, key="set_w_pp")
    pl_target = st.number_input("Private label target (% of sales)", min_value=0.0, step=1.0, format="%.1f",
                                key="set_pl_target", help="The same for everyone.")
    pp_target = st.number_input("PP target (RM per transaction)", min_value=0.0, step=1.0, format="%.2f",
                                key="set_pp_target", help="The same for everyone.")
    cap = st.checkbox("Count at most 100% of a target", key="set_cap",
                      help="Off: doing better than a target earns more points.")
    prorate = st.checkbox("Compare with the target for the days so far", key="set_prorate",
                          help="Scales the sale and profit targets by the share of the month passed, when the report "
                               "starts on the 1st. Private label % and PP have no time element, so without this, "
                               "sales and GP look low early in the month and the points lean on those two.")

c1, c2 = st.columns(2)
with c1:
    sales_file = st.file_uploader("Sales person report (.xlsx)", type=["xlsx"])
with c2:
    pl_file = st.file_uploader("Private label report (.xlsx)", type=["xlsx"])
if sales_file is None or pl_file is None:
    st.info("Upload both reports to begin.")
    st.stop()

try:
    sales_all, period_s = read_report(sales_file)
    pl_all, period_p = read_report(pl_file)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read a file: {e}")
    st.stop()

sales = outlet_rows(sales_all, outlet)
pl = outlet_rows(pl_all, outlet)
if sales.empty:
    st.error(f"No rows for {outlet} in the sales person report. Stores found: "
             + ", ".join(sorted(sales_all["Store"].astype(str).unique())))
    st.stop()
if period_s and period_p and period_s != period_p:
    st.warning(f"The two reports cover different dates: {period_s[0]:%d/%m/%Y} - {period_s[1]:%d/%m/%Y} and "
               f"{period_p[0]:%d/%m/%Y} - {period_p[1]:%d/%m/%Y}.")

people = sales.groupby("person", as_index=False)["No. Transaction"].sum().rename(columns={"No. Transaction": "tx"})
staff_df = detect_staff(people)
with st.expander("Staff and roles", expanded=True):
    st.caption("Tick the people to show, set each person's role, change Order to match your sheet, and edit Name on "
               "sheet. Everyone not ticked goes to Others. The names are shared with the MVP page. "
               "Press Save settings for next time in the sidebar to keep your changes.")
    edited = st.data_editor(
        staff_df, hide_index=True, use_container_width=True,
        disabled=["Name in report", "Transactions"],
        column_config={
            "Use": st.column_config.CheckboxColumn("Use"),
            "Order": st.column_config.NumberColumn("Order", min_value=1, step=1, format="%d"),
            "Role": st.column_config.SelectboxColumn("Role", options=ROLES, required=True),
        },
        key="perf_staff_" + str(abs(hash(tuple(staff_df["Name in report"])))),
    )

chosen = edited[edited["Use"] & (edited["Name on sheet"].astype(str).str.strip() != "")].sort_values(
    "Order", kind="stable")
if chosen.empty:
    st.error("Tick at least one staff in the table above.")
    st.stop()
labels = [str(x).strip().upper() for x in chosen["Name on sheet"]]
if len(set(labels)) < len(labels):
    st.error("Two staff have the same Name on sheet. Please give each a different name.")
    st.stop()
ratios = {"Pharmacist": ratio_pharm, "HA": ratio_ha}
roles = list(chosen["Role"])
if sum(ratios.get(r, 1.0) for r in roles) <= 0:
    st.error("The ratios add up to 0. Give at least one role a ratio above 0.")
    st.stop()
t_sale = share_targets(roles, ratios, sale_target)
t_profit = share_targets(roles, ratios, profit_target)

basis_gp = basis.startswith("Profit")
basis_name = "profit" if basis_gp else "sales"
staff = list(zip(labels, chosen["Name in report"], t_profit if basis_gp else t_sale))

with st.expander("Personal targets", expanded=False):
    tdf = pd.DataFrame({"Staff": labels, "Role": roles, "Ratio": [ratios.get(r, 1.0) for r in roles],
                        "Sale target (RM)": t_sale, "Profit target (RM)": t_profit})
    st.dataframe(tdf, hide_index=True, use_container_width=True, column_config={
        "Ratio": st.column_config.NumberColumn(format="%.2f"),
        "Sale target (RM)": st.column_config.NumberColumn(format="%,.2f"),
        "Profit target (RM)": st.column_config.NumberColumn(format="%,.2f")})

df = build_table(sales, pl, staff, basis_gp)
period = period_s or period_p
dates = f"{period[0]:%d/%m/%Y} - {period[1]:%d/%m/%Y}" if period else "-"
total = df[df["Staff"] == "Total"].iloc[0]
head = {"target_text": f"{basis_name.capitalize()} Target = RM{total['Target']:,.0f}", "dates": dates}
pace = pace_of_month(period)

m1, m2, m3, m4 = st.columns(4)
m1.metric(f"{basis_name.capitalize()} target", rm(total["Target"], 0))
m2.metric("Achieved", rm(total["Achieve"]), delta=pct(total["Pct"]) + " of target", delta_color="off")
m3.metric("Still to go", rm(total["Balance"]))
m4.metric("Month so far", "-" if pace is None else f"{pace * 100:.0f}%",
          help="Share of the month passed at the end of the report. Only shown when the report starts on the 1st.")

st.subheader("Personal performance")
st.caption("Drag from the first cell to the last cell, press Ctrl + C, then paste into your sheet. "
           "Achieve is the " + ("GP" if basis_gp else "net sales") + " of each person. "
           "PP is the average sale per transaction, and PL% is the private label share of sales.")
st.markdown(table_html(df, head), unsafe_allow_html=True)

viz = df[~df["Staff"].isin(["Others", "Total"])].copy()
weights = {"Sales": w_sales, "GP": w_gp, "PL": w_pl, "PP": w_pp}
ranking = None
st.subheader(f"{CROWN} MVP ranking")
if sum(weights.values()) == 0:
    st.warning("All the weights are 0. Give at least one measure a weight in the sidebar.")
elif len(viz) < 2:
    st.info("Tick at least two staff to rank them.")
else:
    factor = pace if (prorate and pace is not None) else 1.0
    targets = {lab: (ts, tp) for lab, ts, tp in zip(labels, t_sale, t_profit)}
    ranking = build_ranking(viz, targets, weights, pl_target, pp_target, cap, factor)
    mvp = ranking.iloc[0]
    tie = (ranking["Total"].round(6) == round(float(mvp["Total"]), 6)).sum() > 1
    st.success(f"{CROWN} **MVP: {mvp['Staff']}** with {mvp['Total']:.1f} points" + (" (tied)" if tie else ""))
    podium = st.columns(min(3, len(ranking)))
    for col, (_, r) in zip(podium, ranking.head(3).iterrows()):
        col.metric(f"#{r['Rank']}{' ' + CROWN if r['MVP'] else ''}", r["Staff"], delta=f"{r['Total']:.1f} points",
                   delta_color="off")
    notes = [f"Each measure's points = the person's achievement against their own target x the weight. "
             f"100 points means every target was reached.",
             "The sale and profit targets follow the role ratios. The private label and PP targets are the same "
             "for everyone."]
    if sum(weights.values()) != 100:
        notes.append(f"The weights add up to {sum(weights.values())}, so they are scaled to 100.")
    if prorate and pace is not None:
        notes.append(f"Sale and profit targets are scaled to {pace * 100:.0f}% of the month.")
    elif prorate:
        notes.append("The report does not start on the 1st, so the targets were not scaled.")
    st.caption(" ".join(notes))
    st.markdown(ranking_html(ranking, weights), unsafe_allow_html=True)
    rc1, rc2 = st.columns(2)
    k1, k2 = ranking_charts(ranking, weights)
    rc1.altair_chart(k1, use_container_width=True)
    rc2.altair_chart(k2, use_container_width=True)

st.download_button("⬇️ Download as Excel (table and MVP ranking)", build_workbook(df, head, ranking),
                   file_name=f"Sale_Performance_{datetime.now():%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.subheader("Who is doing what")
ch1, ch2, ch3, ch4 = charts(viz, pace, basis_name)
left, right = st.columns(2)
left.altair_chart(ch1, use_container_width=True)
right.altair_chart(ch2, use_container_width=True)
left.altair_chart(ch3, use_container_width=True)
right.altair_chart(ch4, use_container_width=True)

with st.sidebar:
    st.divider()
    if st.button("Save settings for next time", type="primary"):
        try:
            save_everything({
                "outlet": outlet, "sale_target": float(sale_target), "profit_target": float(profit_target),
                "basis": basis, "ratio_pharmacist": float(ratio_pharm), "ratio_ha": float(ratio_ha),
                "w_sales": int(w_sales), "w_gp": int(w_gp), "w_pl": int(w_pl), "w_pp": int(w_pp),
                "pl_target": float(pl_target), "pp_target": float(pp_target), "cap": bool(cap),
                "prorate": bool(prorate)}, edited)
            st.success("Saved for everyone who uses this page. The staff names are shared with the MVP page.")
        except OSError as e:
            st.error(f"Could not save: {e}")
