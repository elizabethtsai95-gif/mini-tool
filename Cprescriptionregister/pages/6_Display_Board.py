import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from board_core import (
    DEFAULT_ROSTER, DEFAULT_TARGET, MONTHLY_COLS, ROSTER_COLS, TARGET_COLS, WEIGHTS,
    GSheetStore, LocalStore, build_month, get_targets, month_label, read_report, upsert_month,
)

st.set_page_config(page_title="QKCMC1 Display Board", page_icon="🏆", layout="wide")


def secret(key, default=None):
    """st.secrets raises if no secrets file exists (e.g. local testing) - make it safe."""
    try:
        return st.secrets[key] if key in st.secrets else default
    except Exception:
        return default


@st.cache_resource
def get_store():
    if secret("gcp_service_account") and secret("sheet_id"):
        return GSheetStore(secret("gcp_service_account"), secret("sheet_id")), "Google Sheet"
    return LocalStore(), "local files (no Google Sheet configured)"


store, store_name = get_store()


@st.cache_data(ttl=60)
def load_monthly(_store, nonce):
    df = _store.load("Monthly")
    if df.empty:
        return df
    df["Month"] = df["Month"].astype(str)
    for c in df.columns:
        if c not in ("Month", "Staff", "Role", "MVP"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


if "nonce" not in st.session_state:
    st.session_state.nonce = 0
monthly = load_monthly(store, st.session_state.nonce)

st.title("🏆 QKCMC1 Staff Display Board")
tab_board, tab_year, tab_admin = st.tabs(["🏆 Monthly Board", "📈 Year Tracker", "⬆️ Update (admin)"])

# ------------------------------------------------------------------ Monthly board
with tab_board:
    if monthly.empty:
        st.info("No data yet. Go to the Update tab and upload the first month.")
    else:
        months = sorted(monthly["Month"].unique())
        month = st.selectbox("Month", months[::-1], format_func=month_label)
        m = monthly[monthly["Month"] == month].copy()
        ranked = m[m["Rank"].notna()].sort_values("Rank")
        others = m[m["Rank"].isna()]

        if len(ranked):
            top = ranked.iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("👑 MVP", top["Staff"], f'{top["Points"]:.1f} pts')
            c2.metric("Outlet Sales", f'RM{m["Sales"].sum():,.0f}')
            c3.metric("Outlet GP", f'RM{m["GP"].sum():,.0f}', f'{m["GP"].sum() / m["Sales"].sum() * 100:.1f}%')
            c4.metric("Profit Target", f'RM{ranked["Profit Target"].sum():,.0f}')

            table = ranked.copy()
            table["Staff"] = [("👑 " if r == 1 else "") + s for s, r in zip(table["Staff"], table["Rank"])]
            table["Rank"] = table["Rank"].astype(int)
            show = table[["Rank", "Staff", "Role", "Points", "Sales", "GP", "GP%", "Tx", "PP", "PL", "PL%"]]
            st.dataframe(
                show, hide_index=True, width="stretch",
                column_config={
                    "Points": st.column_config.ProgressColumn("Points", min_value=0, max_value=120, format="%.1f"),
                    "Sales": st.column_config.NumberColumn(format="RM %.0f"),
                    "GP": st.column_config.NumberColumn(format="RM %.0f"),
                    "PL": st.column_config.NumberColumn(format="RM %.0f"),
                    "GP%": st.column_config.NumberColumn(format="%.1f%%"),
                    "PL%": st.column_config.NumberColumn(format="%.1f%%"),
                    "PP": st.column_config.NumberColumn(format="RM %.2f"),
                },
            )

            parts = ranked.melt(
                id_vars=["Staff"], value_vars=["Sales Pts", "GP Pts", "PL Pts", "PP Pts"],
                var_name="Component", value_name="Pts",
            )
            order = ranked["Staff"].tolist()[::-1]
            fig = px.bar(
                parts, y="Staff", x="Pts", color="Component", orientation="h",
                category_orders={"Staff": order}, title=f"Points breakdown - {month_label(month)}",
            )
            fig.add_vline(x=100, line_dash="dash", annotation_text="100 = on target")
            fig.update_layout(height=420, legend_title_text="", xaxis_title="Points")
            st.plotly_chart(fig, width="stretch")

        if len(others):
            st.caption(
                f'Others (not ranked): Sales RM{others["Sales"].sum():,.0f}, GP RM{others["GP"].sum():,.0f}, '
                f'{int(others["Tx"].sum())} Tx.'
            )
        st.caption(
            "Points = Sales 25% + GP 25% + Private Label 25% + PP 25%, each as achievement vs personal target. "
            "100 pts = on target."
        )

# ------------------------------------------------------------------ Year tracker
with tab_year:
    if monthly.empty:
        st.info("No data yet.")
    else:
        years = sorted({m[:4] for m in monthly["Month"]}, reverse=True)
        year = st.selectbox("Year", years)
        y = monthly[monthly["Month"].str.startswith(year)].copy()
        all_months = [f"{year}-{i:02d}" for i in range(1, 13)]
        labels = [month_label(m) for m in all_months]
        y["Label"] = y["Month"].map(month_label)

        st.subheader("👑 MVP of each month")
        mvps = y[y["MVP"] == "Yes"].sort_values("Month")
        if len(mvps):
            hall = mvps[["Label", "Staff", "Points", "Sales", "GP"]].rename(columns={"Label": "Month"})
            cA, cB = st.columns([3, 2])
            cA.dataframe(
                hall, hide_index=True, width="stretch",
                column_config={
                    "Points": st.column_config.NumberColumn(format="%.1f"),
                    "Sales": st.column_config.NumberColumn(format="RM %.0f"),
                    "GP": st.column_config.NumberColumn(format="RM %.0f"),
                },
            )
            wins = mvps["Staff"].value_counts().reset_index()
            wins.columns = ["Staff", "MVP wins"]
            f = px.bar(wins, x="Staff", y="MVP wins", title="MVP wins this year")
            f.update_yaxes(dtick=1)
            cB.plotly_chart(f, width="stretch")

        st.subheader("Performance through the year")
        metric = st.selectbox("Measure", ["Points", "Sales", "GP", "GP%", "PL%", "PP", "Tx", "PL"])
        staff_all = sorted(y.loc[y["Rank"].notna(), "Staff"].unique())
        picked = st.multiselect("Staff", staff_all, default=staff_all)
        d = y[y["Staff"].isin(picked)]
        fig = px.line(
            d, x="Label", y=metric, color="Staff", markers=True,
            category_orders={"Label": labels},
        )
        fig.update_xaxes(categoryorder="array", categoryarray=labels)
        fig.update_layout(height=460, xaxis_title="", legend_title_text="")
        if metric == "Points":
            fig.add_hline(y=100, line_dash="dash", annotation_text="on target")
        st.plotly_chart(fig, width="stretch")

        st.subheader("Points heatmap")
        heat = y[y["Rank"].notna()].pivot_table(index="Staff", columns="Label", values="Points")
        heat = heat.reindex(columns=[l for l in labels if l in heat.columns])
        hf = px.imshow(heat, text_auto=".0f", aspect="auto", color_continuous_scale="RdYlGn", color_continuous_midpoint=100)
        hf.update_layout(height=120 + 40 * len(heat), xaxis_title="", yaxis_title="")
        st.plotly_chart(hf, width="stretch")

# ------------------------------------------------------------------ Admin update
with tab_admin:
    pw = secret("admin_password")
    ok = True
    if pw:
        ok = st.text_input("Admin password", type="password") == pw
    if not ok:
        st.info("Enter the admin password to update the board.")
        st.stop()

    st.caption(f"Data is stored in: {store_name}")
    c1, c2 = st.columns(2)
    f_sales = c1.file_uploader("1) Salesperson report (month end)", type=["xlsx"], key="up_sales")
    f_pl = c2.file_uploader("2) Private label report (month end)", type=["xlsx"], key="up_pl")

    with st.expander("Roster (POS name → display name, role, ratio, ranked?)"):
        roster_saved = store.load("Roster")
        roster_now = roster_saved if not roster_saved.empty else DEFAULT_ROSTER
        roster_now = roster_now[ROSTER_COLS]
        roster_edit = st.data_editor(roster_now, num_rows="dynamic", width="stretch", key="roster_ed")
        st.caption("Anyone not listed goes to 'Others' and is not ranked. Set Ranked = No to leave someone out.")
        if st.button("Save roster"):
            store.save("Roster", roster_edit)
            st.success("Roster saved.")

    if f_sales and f_pl:
        try:
            m1, s_df = read_report(f_sales)
            m2, p_df = read_report(f_pl)
        except Exception as e:
            st.error(f"Could not read the files: {e}")
            st.stop()
        if m1 != m2:
            st.error(f"The two files are for different months ({m1} vs {m2}).")
            st.stop()
        st.success(f"Detected month: {month_label(m1)} - outlet QKCMC1 only")

        tg = get_targets(store, m1)
        st.markdown("**Targets for this month** (edit if needed)")
        tg_df = pd.DataFrame([tg])
        tg_edit = st.data_editor(tg_df, hide_index=True, width="stretch", key=f"tg_{m1}")
        target = {k: float(tg_edit.iloc[0][k]) for k in DEFAULT_TARGET}

        result = build_month(m1, s_df, p_df, roster_edit, target)
        st.markdown("**Preview**")
        st.dataframe(result, hide_index=True, width="stretch")

        if st.button("✅ Save to board", type="primary"):
            upsert_month(store, result)
            tgs = store.load("Targets")
            row = pd.DataFrame([{"Month": m1, **target}])[TARGET_COLS]
            if not tgs.empty:
                tgs = tgs[tgs["Month"].astype(str) != m1]
                row = pd.concat([tgs[TARGET_COLS], row], ignore_index=True)
            store.save("Targets", row.sort_values("Month"))
            st.session_state.nonce += 1
            load_monthly.clear()
            st.success(f"{month_label(m1)} saved. Open the Monthly Board tab to see it.")
