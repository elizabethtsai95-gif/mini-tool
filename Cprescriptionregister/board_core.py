"""Core logic for the QKCMC1 Display Board.
Pure functions (no Streamlit) so they can be tested on their own.
"""
import os
import re

import numpy as np
import pandas as pd

OUTLET = "QKCMC1"
WEIGHTS = {"Sales": 25, "GP": 25, "PL": 25, "PP": 25}  # points out of 100

ROSTER_COLS = ["POS Name", "Display", "Role", "Ratio", "Ranked"]
TARGET_COLS = ["Month", "Sale Target", "Profit Target", "PL% Target", "PP Target"]

DEFAULT_ROSTER = pd.DataFrame(
    [
        ["ELIZABETH TSAI YEN PENG", "ELIZABETH", "Pharmacist", 1.5, "Yes"],
        ["LIM JUN MING", "LIM JUN MING", "Pharmacist", 1.5, "Yes"],
        ["CHONG FUNG CHEN", "CHONG FUNG CHEN", "HA", 1.0, "Yes"],
        ["ELIZZA ANAK VALENTINE", "ELIZZA", "HA", 1.0, "Yes"],
        ["NUR SYAZWANI BINTI REDI", "SYAZWANI", "HA", 1.0, "Yes"],
        ["ELLYTHIA NISHA ANAK SIGON", "ELLYTHIA", "HA", 1.0, "Yes"],
        ["NURSYUHADA BINTI RIZALMAN", "SYUHADA", "HA", 1.0, "Yes"],
        ["MELVIN CRIST ANAK ANOM", "MELL", "HA", 1.0, "Yes"],
    ],
    columns=ROSTER_COLS,
)

# Placeholders - edit in the Targets tab / Update page.
DEFAULT_TARGET = {"Sale Target": 330000, "Profit Target": 100000, "PL% Target": 40.0, "PP Target": 50.0}

MONTHLY_COLS = [
    "Month", "Staff", "Role", "Sales", "GP", "GP%", "Tx", "PP", "PL", "PL%",
    "Sale Target", "Profit Target", "PL% Target", "PP Target",
    "Sales Pts", "GP Pts", "PL Pts", "PP Pts", "Points", "Rank", "MVP",
]


# ----------------------------------------------------------------- reading
def read_report(file, outlet=OUTLET):
    """Read a POS 'Sales Report - By Store, Sales Person' export.
    Returns (month 'YYYY-MM', DataFrame[Sales Person, Tx, Sales, GP]) for one outlet."""
    raw = pd.read_excel(file, header=None)
    hdr = raw.index[raw[0].astype(str).str.strip() == "Company"]
    if len(hdr) == 0:
        raise ValueError("Cannot find the table header row ('Company') in this file.")
    h = hdr[0]

    month = None
    for i in range(h):
        if str(raw.iat[i, 0]).strip().startswith("Sales Date"):
            m = re.search(r"(\d{4})-(\d{2})-\d{2}", str(raw.iat[i, 1]))
            if m:
                month = f"{m.group(1)}-{m.group(2)}"
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h]]
    df = df[df["Store"].astype(str).str.contains(outlet, na=False)]
    df["Sales Person"] = df["Sales Person"].fillna("OTHERS").astype(str).str.strip()
    for c in ["No. Transaction", "Net Sales", "GP"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    out = (
        df.groupby("Sales Person", as_index=False)[["No. Transaction", "Net Sales", "GP"]]
        .sum()
        .rename(columns={"No. Transaction": "Tx", "Net Sales": "Sales"})
    )
    return month, out


# ----------------------------------------------------------------- scoring
def build_month(month, sales_df, pl_df, roster, target):
    """Turn the two reports into one row per staff with points and rank.

    Scoring (same idea as the Sale Performance page - achievement vs target x weight):
      Sales pts = 25 x (Sales / personal sale target)
      GP pts    = 25 x (GP / personal profit target)
      PL pts    = 25 x (PL% / PL% target)
      PP pts    = 25 x (PP / PP target)        PP = Sales / Tx
    Personal targets = outlet target shared by role ratio among ranked staff who
    actually sold that month. 100 points = exactly on target everywhere.
    """
    r = roster.copy()
    r["POS Name"] = r["POS Name"].astype(str).str.strip().str.upper()
    name_map = dict(zip(r["POS Name"], r["Display"]))
    role_map = dict(zip(r["Display"], r["Role"]))
    ratio_map = dict(zip(r["Display"], pd.to_numeric(r["Ratio"], errors="coerce").fillna(1.0)))
    ranked_set = set(r.loc[r["Ranked"].astype(str).str.lower().eq("yes"), "Display"])

    s = sales_df.copy()
    s["Staff"] = s["Sales Person"].str.upper().map(name_map).fillna("Others")
    s = s.groupby("Staff", as_index=False)[["Tx", "Sales", "GP"]].sum()

    p = pl_df.copy()
    p["Staff"] = p["Sales Person"].str.upper().map(name_map).fillna("Others")
    p = p.groupby("Staff", as_index=False)[["Sales"]].sum().rename(columns={"Sales": "PL"})

    t = s.merge(p, on="Staff", how="left").fillna({"PL": 0})
    t["Role"] = t["Staff"].map(role_map).fillna("")
    t["GP%"] = np.where(t["Sales"] > 0, t["GP"] / t["Sales"] * 100, 0)
    t["PP"] = np.where(t["Tx"] > 0, t["Sales"] / t["Tx"], 0)
    t["PL%"] = np.where(t["Sales"] > 0, t["PL"] / t["Sales"] * 100, 0)

    t["_ranked"] = t["Staff"].isin(ranked_set) & (t["Tx"] > 0)
    ratios = t["Staff"].map(ratio_map).fillna(1.0)
    share = ratios.where(t["_ranked"], 0)
    share = share / share.sum() if share.sum() > 0 else share

    t["Sale Target"] = share * float(target["Sale Target"])
    t["Profit Target"] = share * float(target["Profit Target"])
    t["PL% Target"] = float(target["PL% Target"])
    t["PP Target"] = float(target["PP Target"])

    def pts(actual, tgt, w):
        return np.where(t["_ranked"] & (tgt > 0), w * actual / tgt.replace(0, np.nan), np.nan)

    t["Sales Pts"] = pts(t["Sales"], t["Sale Target"], WEIGHTS["Sales"])
    t["GP Pts"] = pts(t["GP"], t["Profit Target"], WEIGHTS["GP"])
    t["PL Pts"] = pts(t["PL%"], t["PL% Target"], WEIGHTS["PL"])
    t["PP Pts"] = pts(t["PP"], t["PP Target"], WEIGHTS["PP"])
    pt_cols = ["Sales Pts", "GP Pts", "PL Pts", "PP Pts"]
    t["Points"] = t[pt_cols].sum(axis=1, min_count=1)
    t["Rank"] = t["Points"].rank(ascending=False, method="min")
    t["MVP"] = np.where(t["Rank"] == 1, "Yes", "")
    t["Month"] = month
    t.loc[~t["_ranked"], "Sale Target"] = np.nan
    t.loc[~t["_ranked"], "Profit Target"] = np.nan
    t.loc[~t["_ranked"], ["PL% Target", "PP Target"]] = np.nan

    t = t.sort_values(["Rank", "Sales"], ascending=[True, False], na_position="last")
    return t[MONTHLY_COLS].round(2).reset_index(drop=True)


# ----------------------------------------------------------------- storage
class LocalStore:
    """CSV files in ./board_data - used when no Google Sheet is configured (testing)."""

    def __init__(self, folder="board_data"):
        self.folder = folder
        os.makedirs(folder, exist_ok=True)

    def load(self, tab):
        path = os.path.join(self.folder, f"{tab}.csv")
        return pd.read_csv(path) if os.path.exists(path) else pd.DataFrame()

    def save(self, tab, df):
        df.to_csv(os.path.join(self.folder, f"{tab}.csv"), index=False)


class GSheetStore:
    """Google Sheet storage via a service account. One tab per table."""

    def __init__(self, info, sheet_id):
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            dict(info),
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        self.book = gspread.authorize(creds).open_by_key(sheet_id)

    def load(self, tab):
        try:
            ws = self.book.worksheet(tab)
        except Exception:
            return pd.DataFrame()
        rows = ws.get_all_records()
        return pd.DataFrame(rows)

    def save(self, tab, df):
        try:
            ws = self.book.worksheet(tab)
        except Exception:
            ws = self.book.add_worksheet(tab, rows=1000, cols=max(26, len(df.columns)))
        clean = df.astype(object).where(df.notna(), "")
        ws.clear()
        ws.update([list(df.columns)] + clean.values.tolist(), value_input_option="USER_ENTERED")


def upsert_month(store, month_df):
    """Replace this month's rows in the Monthly tab (safe to re-upload a month)."""
    cur = store.load("Monthly")
    if not cur.empty:
        cur = cur[cur["Month"].astype(str) != str(month_df["Month"].iloc[0])]
    out = pd.concat([cur, month_df], ignore_index=True)
    out = out.sort_values(["Month", "Rank"], na_position="last").reset_index(drop=True)
    store.save("Monthly", out[MONTHLY_COLS])
    return out


def get_targets(store, month):
    """Targets for a month; falls back to the latest earlier month, then defaults."""
    tg = store.load("Targets")
    if tg.empty:
        return dict(DEFAULT_TARGET)
    tg["Month"] = tg["Month"].astype(str)
    prev = tg[tg["Month"] <= month].sort_values("Month")
    row = (prev.iloc[-1] if len(prev) else tg.sort_values("Month").iloc[-1]).to_dict()
    return {k: float(row[k]) for k in DEFAULT_TARGET}


def month_label(m):
    return pd.Timestamp(f"{m}-01").strftime("%b %Y")
