"""
ZwiftPower Results Dashboard
Run with:  streamlit run dashboard.py
"""

import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIG  (keep in sync with main.py)
# ---------------------------------------------------------------------------
OUTPUT_EXCEL = "zwift_results.xlsx"

RACE_NAMES = [
    "Stage 1: Zwift Games: Kaze Kicker",
    "Stage 2: Zwift Games: Hudson Hustle",
    "Stage 3: Zwift Games: Cobbled Crown",
    "Stage 4: Zwift Games: Peaky Pave",
    "Stage 5: Zwift Games: Three Step Sisters",
    "Stage 6a: Zwift Games: Epiloch",
]

SHORT_LABELS = {
    "Stage 1: Zwift Games: Kaze Kicker":          "S1 Kaze Kicker",
    "Stage 2: Zwift Games: Hudson Hustle":         "S2 Hudson Hustle",
    "Stage 3: Zwift Games: Cobbled Crown":         "S3 Cobbled Crown",
    "Stage 4: Zwift Games: Peaky Pave":            "S4 Peaky Pave",
    "Stage 5: Zwift Games: Three Step Sisters":    "S5 Three Step Sisters",
    "Stage 6a: Zwift Games: Epiloch":              "S6a Epiloch",
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def _to_serial(val) -> float:
    """Normalise a cell value to an Excel serial-time float.

    openpyxl returns time-formatted cells as datetime.time objects.
    We convert those back to the same serial fraction main.py stored.
    """
    if isinstance(val, datetime.time):
        return (val.hour * 3600 + val.minute * 60 + val.second) / 86400
    return float(val)


def fmt_time(val) -> str:
    """Return HH:MM:SS from a serial float or datetime.time, or 'DNS'."""
    if val is None:
        return "DNS"
    if isinstance(val, str):
        return val
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M:%S")
    try:
        if pd.isna(val):
            return "DNS"
    except (TypeError, ValueError):
        pass
    try:
        total_s = round(_to_serial(val) * 86400)
        h, rem  = divmod(total_s, 3600)
        m, s    = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    except (ValueError, TypeError):
        return "DNS"


def is_numeric_time(val) -> bool:
    if val is None or isinstance(val, str):
        return False
    if isinstance(val, datetime.time):
        return True
    try:
        return not pd.isna(val)
    except (TypeError, ValueError):
        return True


def total_time(row) -> float:
    """Sum all stage times as serial fractions; returns NaN if any stage is DNS."""
    total = 0.0
    for rn in RACE_NAMES:
        val = row.get(rn)
        if not is_numeric_time(val):
            return float("nan")
        total += _to_serial(val)
    return total


def cumulative_time(row, stages: list) -> float:
    """Sum serial times for *stages*; returns NaN if any stage is DNS."""
    total = 0.0
    for rn in stages:
        val = row.get(rn)
        if not is_numeric_time(val):
            return float("nan")
        total += _to_serial(val)
    return total


@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    df["total_time"] = df.apply(total_time, axis=1)
    return df


def highlight_dns(val):
    return "color: #b0b0b0; font-style: italic;" if val == "DNS" else ""


def _ranking_table(finishers: pd.DataFrame) -> pd.DataFrame:
    """Return a display-ready ranking table with Rank, Name, Cat, Time columns."""
    out = finishers[["rider_name", "pace_group", "Time"]].copy()
    out.insert(0, "Rank", range(1, len(out) + 1))
    return out.rename(columns={"rider_name": "Name", "pace_group": "Cat"})


def fmt_gap(gap_serial: float) -> str:
    """Format a time gap as '+H:MM:SS'. Zero (the leader) returns 'Leader'."""
    if gap_serial <= 0:
        return "Leader"
    total_s = round(gap_serial * 86400)
    h, rem  = divmod(total_s, 3600)
    m, s    = divmod(rem, 60)
    return f"+{h}:{m:02d}:{s:02d}"


def _cat_ranking_tables(finishers: pd.DataFrame) -> None:
    """Render per-category ranking tables side by side."""
    cats = sorted(finishers["pace_group"].dropna().unique())
    if not cats:
        st.info("No category data available.")
        return

    cols = st.columns(len(cats))
    for col, cat in zip(cols, cats):
        cat_df = finishers[finishers["pace_group"] == cat].reset_index(drop=True)
        cat_df = cat_df[["rider_name", "Time"]].copy()
        cat_df.insert(0, "#", range(1, len(cat_df) + 1))
        cat_df = cat_df.rename(columns={"rider_name": "Name"})
        with col:
            st.markdown(f"**Cat {cat}** — {len(cat_df)} riders")
            st.dataframe(
                cat_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "#":    st.column_config.NumberColumn(width="small"),
                    "Name": st.column_config.TextColumn(width="medium"),
                    "Time": st.column_config.TextColumn(width="medium"),
                },
            )


# ---------------------------------------------------------------------------
# PAGE SETUP
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="ZwiftPower Results",
    page_icon="🚴",
    layout="wide",
)

st.title("🚴 ZwiftPower Race Results")

if not Path(OUTPUT_EXCEL).exists():
    st.error(f"`{OUTPUT_EXCEL}` not found — run `main.py` first.")
    st.stop()

df = load_data(OUTPUT_EXCEL)
present_races = [rn for rn in RACE_NAMES if rn in df.columns]

# ---------------------------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------------------------

pace_options = sorted(df["pace_group"].dropna().unique().tolist())

with st.sidebar:
    st.header("Filters")
    search = st.text_input("Search rider name")

    st.divider()
    st.caption(f"Source: `{OUTPUT_EXCEL}`")
    if st.button("Reload data"):
        st.cache_data.clear()
        st.rerun()

# Apply search filter only; category filtering is done per-tab
mask = pd.Series([True] * len(df), index=df.index)
if search:
    mask &= df["rider_name"].str.contains(search, case=False, na=False)
filtered = df[mask].copy()

# ---------------------------------------------------------------------------
# SUMMARY METRICS
# ---------------------------------------------------------------------------

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Riders", len(filtered))
c2.metric("Categories", len(pace_options))
c3.metric("Completed all stages", int(filtered["total_time"].notna().sum()))

any_dns = filtered[present_races].apply(
    lambda col: col.apply(lambda v: isinstance(v, str) and v == "DNS")
).any(axis=1).sum()
c4.metric("Riders with any DNS", int(any_dns))

all_dns = filtered[present_races].apply(
    lambda col: col.apply(lambda v: isinstance(v, str) and v == "DNS")
).all(axis=1).sum()
c5.metric("Riders all DNS", int(all_dns))

st.divider()

# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------

tab_labels = ["Overview"] + [SHORT_LABELS.get(rn, rn) for rn in present_races] + ["Leaderboard"]
tabs       = st.tabs(tab_labels)

# ── OVERVIEW ────────────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("All Riders")

    # Controls row: category filter + two toggles
    ov_col1, ov_col2, ov_col3 = st.columns([2, 1, 1])
    with ov_col1:
        ov_cats = st.multiselect("Category", pace_options, default=pace_options, key="cat_overview")
    with ov_col2:
        ov_cumulative = st.toggle("Cumulative time", key="ov_cum")
    with ov_col3:
        ov_gap = st.toggle("Gap to leader", key="ov_gap")

    filtered_ov = filtered[filtered["pace_group"].isin(ov_cats)].copy()

    # Always compute raw values from the original numeric data, then build disp
    show_cols = ["rider_id", "rider_name", "pace_group", "profile"] + present_races
    disp = filtered_ov[show_cols].copy()

    for stage_idx, rn in enumerate(present_races):
        stages_so_far = present_races[: stage_idx + 1]

        # Read raw serial times from filtered_ov (never from disp, which gets overwritten)
        if ov_cumulative:
            raw = filtered_ov.apply(lambda r: cumulative_time(r, stages_so_far), axis=1)
        else:
            raw = filtered_ov[rn].apply(lambda v: _to_serial(v) if is_numeric_time(v) else float("nan"))

        # Apply gap or time formatting
        if ov_gap:
            leader = raw.min()
            disp[rn] = (raw - leader).apply(lambda v: fmt_gap(v) if not pd.isna(v) else "DNS")
        else:
            disp[rn] = raw.apply(fmt_time)

    col_label_suffix = " (Gap)" if ov_gap else (" (Cum)" if ov_cumulative else "")
    col_cfg = {
        "rider_id":   st.column_config.NumberColumn("ID",      width="small"),
        "rider_name": st.column_config.TextColumn(  "Name",    width="medium"),
        "pace_group": st.column_config.TextColumn(  "Cat",     width="small"),
        "profile":    st.column_config.LinkColumn(  "Profile", display_text="View", width="small"),
        **{rn: st.column_config.TextColumn(SHORT_LABELS.get(rn, rn) + col_label_suffix, width="medium")
           for rn in present_races},
    }

    st.dataframe(
        disp.style.applymap(highlight_dns, subset=present_races),
        use_container_width=True,
        hide_index=True,
        column_config=col_cfg,
    )

    csv = disp.to_csv(index=False).encode()
    st.download_button("Download CSV", csv, "zwift_results.csv", "text/csv")

# ── PER-STAGE TABS ──────────────────────────────────────────────────────────
for stage_idx, (tab, rn) in enumerate(zip(tabs[1:-1], present_races)):
    with tab:
        st.subheader(rn)

        # Controls row: two toggles + category filter
        tcol1, tcol2, tcol3 = st.columns([1, 1, 2])
        with tcol1:
            cumulative = st.toggle("Cumulative time", key=f"cum_{stage_idx}")
        with tcol2:
            show_gap = st.toggle("Gap to leader", key=f"gap_{stage_idx}")
        with tcol3:
            tab_cats = st.multiselect(
                "Category", pace_options, default=pace_options,
                key=f"cat_{stage_idx}", label_visibility="collapsed",
            )

        stages_so_far = present_races[: stage_idx + 1]
        tab_filtered  = filtered[filtered["pace_group"].isin(tab_cats)]

        # Build finishers with a raw numeric _raw column for gap calculation
        if cumulative:
            stage_df         = tab_filtered[["rider_name", "pace_group"] + stages_so_far].copy()
            stage_df["_raw"] = stage_df.apply(lambda r: cumulative_time(r, stages_so_far), axis=1)
        else:
            stage_df         = tab_filtered[["rider_name", "pace_group", rn]].copy()
            stage_df["_raw"] = stage_df[rn].apply(
                lambda v: _to_serial(v) if is_numeric_time(v) else float("nan")
            )

        finishers = stage_df[stage_df["_raw"].notna()].copy().sort_values("_raw")
        dns_df    = stage_df[stage_df["_raw"].isna()][["rider_name", "pace_group"]]

        # Format the time column based on the gap toggle
        if show_gap and len(finishers) > 0:
            leader_raw        = finishers["_raw"].iloc[0]
            finishers["Time"] = (finishers["_raw"] - leader_raw).apply(fmt_gap)
        else:
            finishers["Time"] = finishers["_raw"].apply(fmt_time)

        time_col_label = "Gap" if show_gap else ("Cumulative" if cumulative else "Time")

        # ── Overall ranking ──────────────────────────────────────────────
        left, right = st.columns([3, 1])
        with left:
            st.markdown(f"**{len(finishers)} finishers**")
            tbl = _ranking_table(finishers).rename(columns={"Time": time_col_label})
            st.dataframe(
                tbl,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Rank":          st.column_config.NumberColumn(width="small"),
                    "Name":          st.column_config.TextColumn(  width="medium"),
                    "Cat":           st.column_config.TextColumn(  width="small"),
                    time_col_label:  st.column_config.TextColumn(  width="medium"),
                },
            )
        with right:
            st.markdown(f"**{len(dns_df)} DNS**")
            st.dataframe(
                dns_df.rename(columns={"rider_name": "Name", "pace_group": "Cat"}),
                use_container_width=True,
                hide_index=True,
            )

        # ── Ranking by category ──────────────────────────────────────────
        with st.expander("Ranking by Category"):
            _cat_ranking_tables(finishers)

# ── OVERALL LEADERBOARD ─────────────────────────────────────────────────────
with tabs[-1]:
    st.subheader("Overall Leaderboard — all stages completed")

    lb_cats      = st.multiselect("Category", pace_options, default=pace_options, key="cat_lb")
    filtered_lb  = filtered[filtered["pace_group"].isin(lb_cats)]
    lb = filtered_lb[filtered_lb["total_time"].notna()].copy()
    lb = lb.sort_values("total_time").reset_index(drop=True)
    lb.insert(0, "Rank", lb.index + 1)

    # Compute gap to overall leader (raw serial fractions, before formatting)
    leader_time  = lb["total_time"].iloc[0] if len(lb) > 0 else 0.0
    lb["_gap"]   = lb["total_time"] - leader_time

    show_gap = st.toggle("Show gap to leader", key="show_gap")

    lb["Total"] = lb["_gap"].apply(fmt_gap) if show_gap else lb["total_time"].apply(fmt_time)
    for rn in present_races:
        lb[rn] = lb[rn].apply(fmt_time)

    show       = ["Rank", "rider_name", "pace_group"] + present_races + ["Total"]
    rename_map = {"rider_name": "Name", "pace_group": "Cat", **SHORT_LABELS}
    lb_disp    = lb[show].rename(columns=rename_map)

    total_col_label = "Gap" if show_gap else "Total"
    lb_disp = lb_disp.rename(columns={"Total": total_col_label})

    st.dataframe(
        lb_disp,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Rank":            st.column_config.NumberColumn(width="small"),
            "Name":            st.column_config.TextColumn(  width="medium"),
            "Cat":             st.column_config.TextColumn(  width="small"),
            total_col_label:   st.column_config.TextColumn(  width="medium"),
            **{SHORT_LABELS.get(rn, rn): st.column_config.TextColumn(width="medium") for rn in present_races},
        },
    )

    if len(lb_disp) > 0:
        st.markdown("---")
        st.markdown("**Top 3 Overall**")
        for _, row in lb_disp.head(3).iterrows():
            medal = {1: "🥇", 2: "🥈", 3: "🥉"}.get(int(row["Rank"]), "")
            st.markdown(f"{medal} **{row['Name']}** ({row['Cat']}) — {row[total_col_label]}")

    # ── Leaderboard by category ──────────────────────────────────────────
    st.divider()
    st.subheader("By Category")

    cats = sorted(lb["pace_group"].dropna().unique())
    if cats:
        cat_cols = st.columns(len(cats))
        for col, cat in zip(cat_cols, cats):
            cat_lb = lb[lb["pace_group"] == cat].copy().reset_index(drop=True)

            # Gap relative to the category leader
            cat_leader_time   = cat_lb["total_time"].iloc[0] if len(cat_lb) > 0 else 0.0
            cat_lb["_cat_gap"] = cat_lb["total_time"] - cat_leader_time

            display_col = cat_lb["_cat_gap"].apply(fmt_gap) if show_gap else cat_lb["total_time"].apply(fmt_time)
            cat_lb["Display"] = display_col
            cat_lb = cat_lb[["rider_name", "Display"]].copy()
            cat_lb.insert(0, "#", range(1, len(cat_lb) + 1))
            cat_lb = cat_lb.rename(columns={"rider_name": "Name", "Display": total_col_label})

            with col:
                st.markdown(f"**Cat {cat}** — {len(cat_lb)} finishers")
                if len(cat_lb) > 0:
                    top = cat_lb.iloc[0]
                    st.markdown(f"🥇 {top['Name']} — {top[total_col_label]}")
                st.dataframe(
                    cat_lb,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "#":              st.column_config.NumberColumn(width="small"),
                        "Name":           st.column_config.TextColumn(  width="medium"),
                        total_col_label:  st.column_config.TextColumn(  width="medium"),
                    },
                )
