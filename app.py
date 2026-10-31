import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# PAGE NAVIGATION STATE & LIQUID GLASS / POWER-BI ANALYTICS CSS
# -----------------------------------------------------------------------------
if "pc_view_page" not in st.session_state:
    st.session_state["pc_view_page"] = "dashboard"

st.markdown(
    """
    <style>
        /* Liquid Glass Theme for Top-Right Analytics Navigation Button */
        div[data-testid="stHorizontalBlock"]:has(.liquid-glass-Marker) div[data-testid="stButton"] > button {
            background: linear-gradient(
                135deg,
                rgba(255, 255, 255, 0.20) 0%,
                rgba(59, 130, 246, 0.28) 45%,
                rgba(14, 165, 233, 0.18) 100%
            ) !important;
            backdrop-filter: blur(18px) saturate(190%) !important;
            -webkit-backdrop-filter: blur(18px) saturate(190%) !important;
            border: 1px solid rgba(255, 255, 255, 0.34) !important;
            border-top: 1px solid rgba(255, 255, 255, 0.55) !important;
            border-left: 1px solid rgba(255, 255, 255, 0.45) !important;
            border-radius: 14px !important;
            color: #F8FAFC !important;
            font-weight: 800 !important;
            font-size: 0.96rem !important;
            letter-spacing: 0.3px !important;
            padding: 0.55rem 1.25rem !important;
            min-height: 44px !important;
            box-shadow:
                0 8px 28px rgba(2, 132, 199, 0.32),
                inset 0 1px 2px rgba(255, 255, 255, 0.45),
                inset 0 -2px 6px rgba(15, 23, 42, 0.45) !important;
            transition: all 0.25s cubic-bezier(0.22, 1, 0.36, 1) !important;
        }
        div[data-testid="stHorizontalBlock"]:has(.liquid-glass-Marker) div[data-testid="stButton"] > button:hover {
            transform: translateY(-2px) scale(1.02) !important;
            background: linear-gradient(
                135deg,
                rgba(255, 255, 255, 0.30) 0%,
                rgba(59, 130, 246, 0.42) 50%,
                rgba(56, 189, 248, 0.30) 100%
            ) !important;
            border-color: rgba(255, 255, 255, 0.68) !important;
            box-shadow:
                0 12px 34px rgba(56, 189, 248, 0.48),
                0 0 18px rgba(59, 130, 246, 0.55),
                inset 0 1px 3px rgba(255, 255, 255, 0.65) !important;
            color: #FFFFFF !important;
        }
        div[data-testid="stHorizontalBlock"]:has(.liquid-glass-Marker) div[data-testid="stButton"] > button:active {
            transform: translateY(0px) scale(0.99) !important;
        }

        /* Power BI Style Cards & Leaderboard Tables */
        .pbi-kpi-grid {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 12px;
            margin-top: 8px;
            margin-bottom: 16px;
        }
        .pbi-kpi-card {
            background: linear-gradient(145deg, rgba(30, 41, 59, 0.68), rgba(15, 23, 42, 0.92));
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 12px;
            padding: 12px 14px;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
        }
        .pbi-kpi-label {
            font-size: 0.74rem;
            font-weight: 600;
            color: #94A3B8;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            margin-bottom: 4px;
        }
        .pbi-kpi-value {
            font-size: 1.45rem;
            font-weight: 800;
            color: #F8FAFC;
            line-height: 1.15;
        }
        .pbi-kpi-sub {
            font-size: 0.73rem;
            font-weight: 600;
            margin-top: 4px;
            color: #CBD5E1;
        }

        /* Top-5 Leaderboard Card Styling */
        .pbi-lb-card {
            background: linear-gradient(160deg, rgba(30, 41, 59, 0.58), rgba(15, 23, 42, 0.90));
            border: 1px solid rgba(148, 163, 184, 0.20);
            border-radius: 12px;
            padding: 12px 14px;
            margin-bottom: 12px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
        }
        .pbi-lb-card-danger {
            border-top: 3px solid #EF4444;
        }
        .pbi-lb-card-success {
            border-top: 3px solid #10B981;
        }
        .pbi-lb-title {
            font-size: 0.88rem;
            font-weight: 800;
            color: #F8FAFC;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        table.pbi-mini-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.79rem;
            color: #E2E8F0;
        }
        table.pbi-mini-table th {
            text-align: left;
            padding: 6px 8px;
            font-size: 0.70rem;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            color: #94A3B8;
            border-bottom: 1px solid rgba(148, 163, 184, 0.22);
            white-space: nowrap;
        }
        table.pbi-mini-table td {
            padding: 7px 8px;
            border-bottom: 1px solid rgba(148, 163, 184, 0.10);
            vertical-align: middle;
        }
        table.pbi-mini-table tr:last-child td {
            border-bottom: none;
        }
        .pbi-rank-pill {
            display: inline-block;
            padding: 2px 7px;
            border-radius: 6px;
            font-weight: 800;
            font-size: 0.71rem;
            background: rgba(59, 130, 246, 0.18);
            color: #93C5FD;
            border: 1px solid rgba(59, 130, 246, 0.35);
        }
        .pbi-badge-red {
            background: rgba(239, 68, 68, 0.18);
            color: #FCA5A5;
            border: 1px solid rgba(239, 68, 68, 0.40);
            padding: 2px 8px;
            border-radius: 999px;
            font-weight: 800;
            font-size: 0.75rem;
            white-space: nowrap;
        }
        .pbi-badge-green {
            background: rgba(16, 185, 129, 0.18);
            color: #6EE7B7;
            border: 1px solid rgba(16, 185, 129, 0.40);
            padding: 2px 8px;
            border-radius: 999px;
            font-weight: 800;
            font-size: 0.75rem;
            white-space: nowrap;
        }
        .pbi-badge-amber {
            background: rgba(245, 158, 11, 0.18);
            color: #FCD34D;
            border: 1px solid rgba(245, 158, 11, 0.40);
            padding: 2px 8px;
            border-radius: 999px;
            font-weight: 700;
            font-size: 0.73rem;
            white-space: nowrap;
        }

        /* CE Truth Detector Verdict Box */
        .truth-audit-box {
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.75), rgba(15, 23, 42, 0.95));
            border: 1px solid rgba(148, 163, 184, 0.25);
            border-left: 5px solid #EF4444;
            border-radius: 12px;
            padding: 14px 18px;
            margin-top: 8px;
            margin-bottom: 12px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# TOP HEADER WITH CENTERED TITLE & RIGHT-ALIGNED LIQUID GLASS 'ANALYTICS' BUTTON
# -----------------------------------------------------------------------------
hdr_left_spacer, hdr_center_title, hdr_right_btn = st.columns(
    [0.16, 0.68, 0.16], gap="small", vertical_alignment="center"
)

with hdr_left_spacer:
    st.markdown('<div class="liquid-glass-Marker"></div>', unsafe_allow_html=True)

with hdr_center_title:
    if st.session_state["pc_view_page"] == "dashboard":
        st.markdown(
            """
            <div class="premium-compact-header" style="justify-content: center;">
                <div class="premium-compact-title">Minor Irrigation - Overdue Works Dashboard</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="premium-compact-header" style="justify-content: center;">
                <div class="premium-compact-title">Minor Irrigation — CE Executive Analytics & ETS Truth Audit</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

with hdr_right_btn:
    if st.session_state["pc_view_page"] == "dashboard":
        if st.button("✨ Analytics", use_container_width=True, key="btn_go_analytics"):
            st.session_state["pc_view_page"] = "analytics"
            st.rerun()
    else:
        if st.button("⬅️ Dashboard", use_container_width=True, key="btn_go_dashboard"):
            st.session_state["pc_view_page"] = "dashboard"
            st.rerun()

# -----------------------------------------------------------------------------
# DATA LOAD (AUTOMATICALLY CONNECTS TO YOUR EXISTING SHARED STATE FUNCTION)
# -----------------------------------------------------------------------------
_shared_state_fn = next(
    (
        v
        for k, v in globals().items()
        if k.startswith("get_shared_state") and callable(v)
    ),
    None,
)
if _shared_state_fn is None:
    st.error("Could not find `get_shared_state` function in the top section of `app.py`.")
    st.stop()

shared_state = _shared_state_fn()

needs_fresh_load = (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or sync_clicked
)

if not needs_fresh_load and shared_state["data"] is not None:
    cached_tuple = shared_state["data"]
    cached_od_df = cached_tuple[1]
    if not cached_od_df.empty and (
        "Overdue By" not in cached_od_df.columns
        or COL_WORK_NAME_KN not in cached_od_df.columns
    ):
        needs_fresh_load = True

if uploaded_xlsx is not None:
    file_id = f"{uploaded_xlsx.name}_{uploaded_xlsx.size}_{SCHEMA_VERSION}"
    if st.session_state.get("_last_uploaded_id") != file_id:
        with shared_state["lock"]:
            parsed = parse_xlsx_workbook(uploaded_xlsx.getvalue())
            save_shared_state(
                shared_state, parsed, f"Uploaded File ({uploaded_xlsx.name})"
            )
            st.session_state["_last_uploaded_id"] = file_id

elif needs_fresh_load:
    with shared_state["lock"]:
        with st.spinner("Loading 23 worksheets from Google Sheets..."):
            try:
                xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                parsed = parse_xlsx_workbook(xlsx_bytes)
                save_shared_state(
                    shared_state, parsed, "Google Sheets Live Snapshot"
                )
            except Exception as e:
                if shared_state["data"] is not None:
                    st.warning(
                        f"Google Sheets is busy ({e}). Showing snapshot from {shared_state['last_synced']}."
                    )
                else:
                    st.error(f"Unable to load data right now. ({e})")
                    st.stop()

_loaded_tuple = shared_state["data"]
master_df, overdue_full_df, diag_df, today_str = (
    _loaded_tuple[0],
    _loaded_tuple[1],
    _loaded_tuple[2],
    _loaded_tuple[3],
)

st.markdown(
    f"""
    <div class="mi-status-bar">
        <span class="mi-status-pill">⚡ {shared_state['source_label']}</span>
        <span>🕒 <b>Last Synced:</b> {shared_state['last_synced']} IST</span>
        <span>📌 <b>Row Rule:</b> Valid text in <b>Division</b> column (empty <b>Sub-division</b> rows included)</span>
    </div>
    """,
    unsafe_allow_html=True,
)

if master_df.empty:
    st.warning("No matching data rows found.")
    st.stop()


# =============================================================================
# HELPER FUNCTIONS FOR THE CE ANALYTICS PAGE
# =============================================================================
def build_officer_analytics_df(df_scope: pd.DataFrame, col_key: str, label_name: str) -> pd.DataFrame:
    """Builds complete officer-level metrics including Overdue, ETS, On-Track P, and Reality Gap %."""
    sub = df_scope[
        (df_scope[col_key] != "")
        & (df_scope[col_key] != UNASSIGNED_LABEL)
        & (df_scope[col_key] != UNSPECIFIED_LABEL)
    ].copy()
    if sub.empty:
        return pd.DataFrame()

    grp = (
        sub.groupby(col_key, as_index=False)
        .agg(
            total_works=("is_overdue", "count"),
            ongoing_p_works=("is_status_p", "sum"),
            ets_works=("is_status_ets", "sum"),
            overdue_works=("is_overdue", "sum"),
            primary_div=("division_name", lambda s: s.mode().iloc[0] if not s.mode().empty else ""),
        )
        .rename(columns={col_key: label_name})
    )

    grp["on_track_p"] = (grp["ongoing_p_works"] - grp["overdue_works"]).clip(lower=0)
    grp["rest_works"] = (
        grp["total_works"] - grp["overdue_works"] - grp["ets_works"]
    ).clip(lower=0)
    grp["overdue_pct"] = (
        (grp["overdue_works"] / grp["total_works"]) * 100
    ).round(1).fillna(0.0)
    grp["ets_pct"] = (
        (grp["ets_works"] / grp["total_works"]) * 100
    ).round(1).fillna(0.0)
    grp["delayed_exposure"] = grp["overdue_works"] + grp["ets_works"]
    grp["unexcused_ratio_pct"] = (
        (grp["overdue_works"] / grp["delayed_exposure"].replace(0, pd.NA)) * 100
    ).fillna(0.0).round(1)
    return grp


def render_top5_leaderboard_html(
    df_top5: pd.DataFrame,
    officer_col: str,
    title_text: str,
    subtitle_text: str,
    is_danger: bool = True,
    primary_metric: str = "count",
) -> str:
    card_border_cls = "pbi-lb-card-danger" if is_danger else "pbi-lb-card-success"
    badge_cls = "pbi-badge-red" if is_danger else "pbi-badge-green"

    if df_top5.empty:
        return (
            f'<div class="pbi-lb-card {card_border_cls}">'
            f'<div class="pbi-lb-title"><span>{html.escape(title_text)}</span>'
            f'<span style="font-size:0.72rem;color:#94A3B8;">{html.escape(subtitle_text)}</span></div>'
            f'<div style="color:#94A3B8;font-size:0.8rem;padding:8px 0;">No officers match the current criteria.</div>'
            f"</div>"
        )

    rows_html = []
    for rank_idx, (_, r) in enumerate(df_top5.iterrows(), start=1):
        off_name = html.escape(str(r[officer_col]))
        div_name = html.escape(str(r.get("primary_div", "")))
        od_cnt = int(r["overdue_works"])
        od_pct = float(r["overdue_pct"])
        ets_cnt = int(r["ets_works"])
        tot_cnt = int(r["total_works"])

        if primary_metric == "pct":
            main_badge = f'<span class="{badge_cls}">{od_pct:.1f}% ({od_cnt})</span>'
        else:
            main_badge = f'<span class="{badge_cls}">{od_cnt} ({od_pct:.1f}%)</span>'

        ets_badge = (
            f'<span class="pbi-badge-amber">{ets_cnt} ETS</span>'
            if ets_cnt > 0
            else '<span style="color:#64748B;font-size:0.74rem;">0 ETS</span>'
        )

        rows_html.append(
            f"<tr>"
            f'<td><span class="pbi-rank-pill">#{rank_idx}</span></td>'
            f'<td><div style="font-weight:700;color:#F8FAFC;">{off_name}</div>'
            f'<div style="font-size:0.68rem;color:#94A3B8;">🏛️ {div_name}</div></td>'
            f'<td style="text-align:right;">{main_badge}</td>'
            f'<td style="text-align:right;">{ets_badge}</td>'
            f'<td style="text-align:right;font-weight:700;color:#38BDF8;">{tot_cnt}</td>'
            f"</tr>"
        )

    return (
        f'<div class="pbi-lb-card {card_border_cls}">'
        f'<div class="pbi-lb-title">'
        f"<span>{html.escape(title_text)}</span>"
        f'<span style="font-size:0.71rem;color:#94A3B8;font-weight:600;">{html.escape(subtitle_text)}</span>'
        f"</div>"
        f'<table class="pbi-mini-table">'
        f"<thead><tr>"
        f"<th>#</th><th>{html.escape(officer_col)}</th>"
        f'<th style="text-align:right;">Overdue</th>'
        f'<th style="text-align:right;">In ETS</th>'
        f'<th style="text-align:right;">Total</th>'
        f"</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        f"</table></div>"
    )


def classify_overdue_age_bucket(target_str: str, today_dt: date) -> str:
    parsed_list = parse_indian_date_list(pd.Series([target_str]))
    t_dt = parsed_list[0] if parsed_list else None
    if not isinstance(t_dt, date) or t_dt >= today_dt:
        return "Unknown"
    days_od = (today_dt - t_dt).days
    if days_od <= 180:
        return "1. < 6 Months"
    if days_od <= 365:
        return "2. 6 – 12 Months"
    if days_od <= 730:
        return "3. 1 – 2 Years"
    if days_od <= 1825:
        return "4. 2 – 5 Years"
    return "5. > 5 Years (Chronic)"


# =============================================================================
# PAGE 2: CHIEF ENGINEER POWER-BI ANALYTICS & ETS TRUTH AUDIT VIEW
# =============================================================================
if st.session_state["pc_view_page"] == "analytics":
    # --- Top Scope Slicers for Analytics ---
    with st.container(border=True):
        s_col1, s_col2, s_col3, s_col4 = st.columns([1.2, 1.2, 1.2, 1.4], gap="small")

        div_opts_an = [ALL_OPTION] + sorted(
            [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
        )
        with s_col1:
            an_div = st.selectbox("🏛️ Filter Analytics by Division:", options=div_opts_an, index=0)

        an_df1 = (
            master_df[master_df["division_name"] == an_div].copy()
            if an_div != ALL_OPTION
            else master_df.copy()
        )

        subdiv_opts_an = build_dropdown_options(an_df1, "subdivision_name")
        with s_col2:
            an_subdiv = st.selectbox("🏢 Filter by Sub-division:", options=subdiv_opts_an, index=0)

        an_df2 = (
            an_df1[an_df1["subdivision_name"] == an_subdiv].copy()
            if an_subdiv != ALL_OPTION
            else an_df1.copy()
        )

        ee_opts_an = build_dropdown_options(an_df2, "ee_name")
        with s_col3:
            an_ee = st.selectbox("👔 Filter by Executive Engineer:", options=ee_opts_an, index=0)

        an_scope_df = (
            an_df2[an_df2["ee_name"] == an_ee].copy()
            if an_ee != ALL_OPTION
            else an_df2.copy()
        )

        with s_col4:
            min_works_threshold = st.slider(
                "⚖️ Min. Works Assigned for 'Least Overdue' Ranking:",
                min_value=1,
                max_value=50,
                value=5,
                help="Filters out officers with very few works (e.g., 1 or 2 works) when ranking Top 5 Least Overdue.",
            )

    if an_scope_df.empty:
        st.warning("No works found for the selected analytics filter scope.")
        st.stop()

    # Build Officer DataFrames for AE and AEE
    ae_stats_df = build_officer_analytics_df(an_scope_df, "so_name", LABEL_AE_EN)
    aee_stats_df = build_officer_analytics_df(an_scope_df, "aee_name", LABEL_AEE_EN)

    # Calculate Portfolio-Wide KPIs
    tot_works_an = len(an_scope_df)
    tot_od_an = int(an_scope_df["is_overdue"].sum())
    tot_ets_an = int(an_scope_df["is_status_ets"].sum())
    tot_ongoing_p_an = int(an_scope_df["is_status_p"].sum())
    tot_ontrack_p_an = max(0, tot_ongoing_p_an - tot_od_an)
    tot_rest_other_an = max(0, tot_works_an - tot_od_an - tot_ets_an - tot_ontrack_p_an)

    od_pct_an = round((tot_od_an / tot_works_an) * 100, 1) if tot_works_an > 0 else 0.0
    ets_pct_an = round((tot_ets_an / tot_works_an) * 100, 1) if tot_works_an > 0 else 0.0
    tot_delayed_an = tot_od_an + tot_ets_an
    unexcused_share_an = (
        round((tot_od_an / tot_delayed_an) * 100, 1) if tot_delayed_an > 0 else 0.0
    )

    # Compute Ageing on Overdue Works in Scope
    today_ist_obj = datetime.now(ZoneInfo("Asia/Kolkata")).date()
    od_rows_in_scope = an_scope_df[an_scope_df["is_overdue"]].copy()
    if not od_rows_in_scope.empty:
        od_rows_in_scope["age_bucket"] = od_rows_in_scope["target_date_raw"].map(
            lambda x: classify_overdue_age_bucket(str(x), today_ist_obj)
        )
        chronic_2yr_plus = int(
            od_rows_in_scope["age_bucket"]
            .isin(["4. 2 – 5 Years", "5. > 5 Years (Chronic)"])
            .sum()
        )
    else:
        chronic_2yr_plus = 0

    # --- 6-Card Power BI Executive KPI Ribbon ---
    st.markdown(
        f"""
        <div class="pbi-kpi-grid">
            <div class="pbi-kpi-card" style="border-top: 3px solid #06B6D4;">
                <div class="pbi-kpi-label">📋 Total Works in Scope</div>
                <div class="pbi-kpi-value">{tot_works_an:,}</div>
                <div class="pbi-kpi-sub">{len(ae_stats_df)} AEs &bull; {len(aee_stats_df)} AEEs</div>
            </div>
            <div class="pbi-kpi-card" style="border-top: 3px solid #EF4444;">
                <div class="pbi-kpi-label">🔴 True Overdue ('P')</div>
                <div class="pbi-kpi-value" style="color:#FCA5A5;">{tot_od_an:,} ({od_pct_an}%)</div>
                <div class="pbi-kpi-sub">Past Target Date &amp; Not in ETS</div>
            </div>
            <div class="pbi-kpi-card" style="border-top: 3px solid #10B981;">
                <div class="pbi-kpi-label">🟢 Claimed 'ETS' Works</div>
                <div class="pbi-kpi-value" style="color:#6EE7B7;">{tot_ets_an:,} ({ets_pct_an}%)</div>
                <div class="pbi-kpi-sub">Extension of Time Status</div>
            </div>
            <div class="pbi-kpi-card" style="border-top: 3px solid #F59E0B;">
                <div class="pbi-kpi-label">⚠️ Total Delayed Exposure</div>
                <div class="pbi-kpi-value" style="color:#FCD34D;">{tot_delayed_an:,}</div>
                <div class="pbi-kpi-sub">Overdue ('P') + 'ETS' Combined</div>
            </div>
            <div class="pbi-kpi-card" style="border-top: 3px solid #EC4899;">
                <div class="pbi-kpi-label">🕵️ Unexcused Overdue Ratio</div>
                <div class="pbi-kpi-value" style="color:#F472B6;">{unexcused_share_an}%</div>
                <div class="pbi-kpi-sub">Of delayed works are NOT in ETS</div>
            </div>
            <div class="pbi-kpi-card" style="border-top: 3px solid #8B5CF6;">
                <div class="pbi-kpi-label">⏳ Chronic Overdue (&gt;2 Yrs)</div>
                <div class="pbi-kpi-value" style="color:#C4B5FD;">{chronic_2yr_plus:,}</div>
                <div class="pbi-kpi-sub">Target elapsed by 2+ years</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # =========================================================================
    # SECTION 1: CE's "TRUTH DETECTOR" — OVERDUE ('P') VS. 'ETS' EXCUSE AUDIT
    # =========================================================================
    st.markdown(
        """
        <div class="exec-section-banner">
            <div class="exec-section-title">🕵️ CE Truth Detector: Actual Overdue ('P') vs. Claimed 'ETS' Verification</div>
            <div class="exec-scope-badge">Exposes whether delayed works are genuinely in 'ETS' or sitting unexcused in Overdue ('Progress')</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        td_col_left, td_col_right = st.columns([0.58, 0.42], gap="medium")

        with td_col_left:
            audit_role = st.radio(
                "Select Officer Cadre for ETS vs. Overdue Audit:",
                options=[LABEL_AE_EN, LABEL_AEE_EN],
                horizontal=True,
                key="an_truth_role",
            )
            active_audit_df = (
                ae_stats_df if audit_role == LABEL_AE_EN else aee_stats_df
            )

            if not active_audit_df.empty:
                top_delayed_15 = (
                    active_audit_df[active_audit_df["delayed_exposure"] > 0]
                    .sort_values(
                        by=["overdue_works", "ets_works", "total_works"],
                        ascending=[False, False, False],
                    )
                    .head(15)
                    .iloc[::-1]
                )

                fig_truth = go.Figure()
                fig_truth.add_trace(
                    go.Bar(
                        y=top_delayed_15[audit_role],
                        x=top_delayed_15["overdue_works"],
                        name="🔴 Actual Overdue ('P' — Not in ETS)",
                        orientation="h",
                        marker_color="#EF4444",
                        text=top_delayed_15["overdue_works"],
                        textposition="inside",
                        insidetextanchor="middle",
                        textfont=dict(color="#FFFFFF", weight="bold", size=11),
                    )
                )
                fig_truth.add_trace(
                    go.Bar(
                        y=top_delayed_15[audit_role],
                        x=top_delayed_15["ets_works"],
                        name="🟢 Claimed in 'ETS'",
                        orientation="h",
                        marker_color="#10B981",
                        text=top_delayed_15["ets_works"],
                        textposition="outside",
                        textfont=dict(color="#6EE7B7", weight="bold", size=11),
                    )
                )
                fig_truth.update_layout(
                    barmode="stack",
                    title=dict(
                        text=f"Top 15 {audit_role}s: Actual Overdue ('P') vs. Claimed 'ETS'",
                        font=dict(size=14, color="#F8FAFC"),
                    ),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(15,23,42,0.45)",
                    font=dict(color="#E2E8F0", size=12),
                    height=460,
                    margin=dict(l=10, r=40, t=42, b=20),
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.01,
                        xanchor="left",
                        x=0,
                    ),
                    xaxis=dict(
                        title="Number of Works (Overdue + ETS)",
                        gridcolor="rgba(148,163,184,0.12)",
                    ),
                    yaxis=dict(title=""),
                )
                st.plotly_chart(fig_truth, use_container_width=True)

        with td_col_right:
            st.markdown(
                f"##### 🔍 Instant {audit_role} Fact-Checker (During Review Meeting)"
            )
            if not active_audit_df.empty:
                audit_sorted_opts = active_audit_df.sort_values(
                    by=["overdue_works", "ets_works"], ascending=[False, False]
                )[audit_role].tolist()

                picked_off = st.selectbox(
                    f"Select {audit_role} to verify their ETS claim:",
                    options=audit_sorted_opts,
                    index=0,
                    key="an_fact_check_select",
                )

                off_row = active_audit_df[
                    active_audit_df[audit_role] == picked_off
                ].iloc[0]
                o_tot = int(off_row["total_works"])
                o_od = int(off_row["overdue_works"])
                o_ets = int(off_row["ets_works"])
                o_ontrack = int(off_row["on_track_p"])
                o_rest = int(off_row["rest_works"])
                o_od_pct = float(off_row["overdue_pct"])
                o_unexcused_pct = float(off_row["unexcused_ratio_pct"])

                if o_od == 0 and o_ets == 0:
                    verdict_color = "#10B981"
                    verdict_title = "✅ CLEAN RECORD (0 Overdue, 0 ETS)"
                    verdict_msg = f"<b>{html.escape(picked_off)}</b> has zero overdue works and zero ETS works across {o_tot} assigned works."
                elif o_od > 0 and o_ets == 0:
                    verdict_color = "#EF4444"
                    verdict_title = "🚨 100% UNEXCUSED OVERDUE (ZERO ETS!)"
                    verdict_msg = (
                        f"If <b>{html.escape(picked_off)}</b> claims their delayed works are in ETS, "
                        f"<b>that is false</b>: <b>{o_od} work(s)</b> are sitting in raw Overdue ('P') and <b>0 works</b> are in ETS."
                    )
                elif o_od > o_ets:
                    verdict_color = "#F97316"
                    verdict_title = f"⚠️ MORE OVERDUE THAN ETS ({o_unexcused_pct:.1f}% Unexcused)"
                    verdict_msg = (
                        f"Out of <b>{o_od + o_ets}</b> delayed works under <b>{html.escape(picked_off)}</b>, "
                        f"only <b>{o_ets}</b> are in ETS while <b>{o_od} works ({o_unexcused_pct:.1f}%)</b> are still unexcused Overdue ('P')."
                    )
                else:
                    verdict_color = "#F59E0B"
                    verdict_title = f"📌 ETS HEAVY ({o_ets} in ETS vs. {o_od} Overdue)"
                    verdict_msg = (
                        f"<b>{html.escape(picked_off)}</b> has placed <b>{o_ets}</b> works in ETS, "
                        f"but still has <b>{o_od}</b> active Overdue ('P') work(s) ({o_od_pct:.1f}% of total)."
                    )

                st.markdown(
                    f"""
                    <div class="truth-audit-box" style="border-left-color: {verdict_color};">
                        <div style="font-size:0.86rem;font-weight:800;color:{verdict_color};margin-bottom:6px;">
                            {verdict_title}
                        </div>
                        <div style="font-size:0.82rem;color:#E2E8F0;line-height:1.45;margin-bottom:10px;">
                            {verdict_msg}
                        </div>
                        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px;text-align:center;">
                            <div style="background:rgba(239,68,68,0.16);border:1px solid rgba(239,68,68,0.4);border-radius:8px;padding:6px;">
                                <div style="font-size:0.65rem;color:#FCA5A5;">OVERDUE ('P')</div>
                                <div style="font-size:1.1rem;font-weight:800;color:#FCA5A5;">{o_od}</div>
                            </div>
                            <div style="background:rgba(16,185,129,0.16);border:1px solid rgba(16,185,129,0.4);border-radius:8px;padding:6px;">
                                <div style="font-size:0.65rem;color:#6EE7B7;">IN 'ETS'</div>
                                <div style="font-size:1.1rem;font-weight:800;color:#6EE7B7;">{o_ets}</div>
                            </div>
                            <div style="background:rgba(59,130,246,0.16);border:1px solid rgba(59,130,246,0.4);border-radius:8px;padding:6px;">
                                <div style="font-size:0.65rem;color:#93C5FD;">ON-TRACK ('P')</div>
                                <div style="font-size:1.1rem;font-weight:800;color:#93C5FD;">{o_ontrack}</div>
                            </div>
                            <div style="background:rgba(148,163,184,0.16);border:1px solid rgba(148,163,184,0.3);border-radius:8px;padding:6px;">
                                <div style="font-size:0.65rem;color:#CBD5E1;">TOTAL WORKS</div>
                                <div style="font-size:1.1rem;font-weight:800;color:#F8FAFC;">{o_tot}</div>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # Mini Donut of Selected Officer's Portfolio
                fig_off_donut = go.Figure(
                    data=[
                        go.Pie(
                            labels=[
                                "Overdue ('P')",
                                "Claimed 'ETS'",
                                "On-Track ('P')",
                                "Completed / Rest",
                            ],
                            values=[o_od, o_ets, o_ontrack, o_rest],
                            hole=0.62,
                            marker=dict(
                                colors=["#EF4444", "#10B981", "#3B82F6", "#475569"]
                            ),
                            textinfo="label+value",
                            textfont=dict(size=11, color="#F8FAFC"),
                        )
                    ]
                )
                fig_off_donut.update_layout(
                    height=225,
                    margin=dict(l=10, r=10, t=10, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    showlegend=False,
                    annotations=[
                        dict(
                            text=f"<b>{o_tot}</b><br>Works",
                            x=0.5,
                            y=0.5,
                            font_size=13,
                            font_color="#F8FAFC",
                            showarrow=False,
                        )
                    ],
                )
                st.plotly_chart(fig_off_donut, use_container_width=True)

    # =========================================================================
    # SECTION 2: THE 8 TOP-5 LEADERBOARDS (AE & AEE — HIGHEST & LEAST OVERDUE)
    # =========================================================================
    st.markdown(
        f"""
        <div class="exec-section-banner">
            <div class="exec-section-title">🏆 Top 5 AE &amp; AEE Leaderboards — Highest &amp; Least Overdue (By Count &amp; Percentage)</div>
            <div class="exec-scope-badge">Least Overdue rankings require at least {min_works_threshold} assigned works (adjustable via top slider)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Compute the 4 AE Top-5 Slices
    ae_qualified = ae_stats_df[ae_stats_df["total_works"] >= min_works_threshold].copy()
    if ae_qualified.empty:
        ae_qualified = ae_stats_df.copy()

    ae_high_cnt = ae_stats_df.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"], ascending=[False, False, False]
    ).head(5)
    ae_high_pct = ae_qualified.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"], ascending=[False, False, False]
    ).head(5)
    ae_least_cnt = ae_qualified.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"], ascending=[True, True, False]
    ).head(5)
    ae_least_pct = ae_qualified.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"], ascending=[True, True, False]
    ).head(5)

    # Compute the 4 AEE Top-5 Slices
    aee_qualified = aee_stats_df[aee_stats_df["total_works"] >= min_works_threshold].copy()
    if aee_qualified.empty:
        aee_qualified = aee_stats_df.copy()

    aee_high_cnt = aee_stats_df.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"], ascending=[False, False, False]
    ).head(5)
    aee_high_pct = aee_qualified.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"], ascending=[False, False, False]
    ).head(5)
    aee_least_cnt = aee_qualified.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"], ascending=[True, True, False]
    ).head(5)
    aee_least_pct = aee_qualified.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"], ascending=[True, True, False]
    ).head(5)

    # Row 1 of Leaderboards: HIGHEST OVERDUE (4 Cards: AE Count, AE %, AEE Count, AEE %)
    lb_r1_c1, lb_r1_c2, lb_r1_c3, lb_r1_c4 = st.columns(4, gap="small")
    with lb_r1_c1:
        st.markdown(
            render_top5_leaderboard_html(
                ae_high_cnt,
                LABEL_AE_EN,
                "🔴 Top 5 AEs — Highest Overdue (Count)",
                "By Overdue Count",
                is_danger=True,
                primary_metric="count",
            ),
            unsafe_allow_html=True,
        )
    with lb_r1_c2:
        st.markdown(
            render_top5_leaderboard_html(
                ae_high_pct,
                LABEL_AE_EN,
                "🔴 Top 5 AEs — Highest Overdue (%)",
                f"Min {min_works_threshold} Works",
                is_danger=True,
                primary_metric="pct",
            ),
            unsafe_allow_html=True,
        )
    with lb_r1_c3:
        st.markdown(
            render_top5_leaderboard_html(
                aee_high_cnt,
                LABEL_AEE_EN,
                "🔴 Top 5 AEEs — Highest Overdue (Count)",
                "By Overdue Count",
                is_danger=True,
                primary_metric="count",
            ),
            unsafe_allow_html=True,
        )
    with lb_r1_c4:
        st.markdown(
            render_top5_leaderboard_html(
                aee_high_pct,
                LABEL_AEE_EN,
                "🔴 Top 5 AEEs — Highest Overdue (%)",
                f"Min {min_works_threshold} Works",
                is_danger=True,
                primary_metric="pct",
            ),
            unsafe_allow_html=True,
        )

    # Row 2 of Leaderboards: LEAST OVERDUE (4 Cards: AE Count, AE %, AEE Count, AEE %)
    lb_r2_c1, lb_r2_c2, lb_r2_c3, lb_r2_c4 = st.columns(4, gap="small")
    with lb_r2_c1:
        st.markdown(
            render_top5_leaderboard_html(
                ae_least_cnt,
                LABEL_AE_EN,
                "🟢 Top 5 AEs — Least Overdue (Count)",
                f"Min {min_works_threshold} Works",
                is_danger=False,
                primary_metric="count",
            ),
            unsafe_allow_html=True,
        )
    with lb_r2_c2:
        st.markdown(
            render_top5_leaderboard_html(
                ae_least_pct,
                LABEL_AE_EN,
                "🟢 Top 5 AEs — Least Overdue (%)",
                f"Min {min_works_threshold} Works",
                is_danger=False,
                primary_metric="pct",
            ),
            unsafe_allow_html=True,
        )
    with lb_r2_c3:
        st.markdown(
            render_top5_leaderboard_html(
                aee_least_cnt,
                LABEL_AEE_EN,
                "🟢 Top 5 AEEs — Least Overdue (Count)",
                f"Min {min_works_threshold} Works",
                is_danger=False,
                primary_metric="count",
            ),
            unsafe_allow_html=True,
        )
    with lb_r2_c4:
        st.markdown(
            render_top5_leaderboard_html(
                aee_least_pct,
                LABEL_AEE_EN,
                "🟢 Top 5 AEEs — Least Overdue (%)",
                f"Min {min_works_threshold} Works",
                is_danger=False,
                primary_metric="pct",
            ),
            unsafe_allow_html=True,
        )

    # =========================================================================
    # SECTION 3: POWER-BI DONUT CHARTS, AGEING SEVERITY & DIVISION/HOA BREAKDOWN
    # =========================================================================
    st.markdown(
        """
        <div class="exec-section-banner">
            <div class="exec-section-title">📊 Portfolio Breakdown: Work Status, Overdue Ageing, Division &amp; Scheme (H.O.A.)</div>
            <div class="exec-scope-badge">Power BI Visuals for Executive Review</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pbi_c1, pbi_c2, pbi_c3 = st.columns([1, 1, 1.35], gap="small")

    with pbi_c1:
        with st.container(border=True):
            fig_status_donut = go.Figure(
                data=[
                    go.Pie(
                        labels=[
                            "Overdue ('P')",
                            "Claimed 'ETS'",
                            "On-Track Ongoing ('P')",
                            "Completed / Other Statuses",
                        ],
                        values=[
                            tot_od_an,
                            tot_ets_an,
                            tot_ontrack_p_an,
                            tot_rest_other_an,
                        ],
                        hole=0.58,
                        marker=dict(
                            colors=["#EF4444", "#10B981", "#3B82F6", "#475569"]
                        ),
                        textinfo="percent+value",
                        textfont=dict(size=11, color="#F8FAFC"),
                    )
                ]
            )
            fig_status_donut.update_layout(
                title=dict(
                    text="🍩 Overall Works Status Breakdown",
                    font=dict(size=14, color="#F8FAFC"),
                ),
                height=340,
                margin=dict(l=10, r=10, t=40, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.18,
                    xanchor="center",
                    x=0.5,
                    font=dict(color="#CBD5E1", size=10),
                ),
                annotations=[
                    dict(
                        text=f"<b>{tot_works_an:,}</b><br>Total",
                        x=0.5,
                        y=0.5,
                        font_size=14,
                        font_color="#F8FAFC",
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_status_donut, use_container_width=True)

    with pbi_c2:
        with st.container(border=True):
            if od_rows_in_scope.empty:
                st.info("No overdue works to compute ageing buckets.")
            else:
                age_counts = (
                    od_rows_in_scope["age_bucket"]
                    .value_counts()
                    .reindex(
                        [
                            "1. < 6 Months",
                            "2. 6 – 12 Months",
                            "3. 1 – 2 Years",
                            "4. 2 – 5 Years",
                            "5. > 5 Years (Chronic)",
                        ],
                        fill_value=0,
                    )
                )
                fig_age_donut = go.Figure(
                    data=[
                        go.Pie(
                            labels=age_counts.index.tolist(),
                            values=age_counts.values.tolist(),
                            hole=0.58,
                            marker=dict(
                                colors=[
                                    "#38BDF8",
                                    "#FBBF24",
                                    "#F97316",
                                    "#EF4444",
                                    "#991B1B",
                                ]
                            ),
                            textinfo="percent+value",
                            textfont=dict(size=11, color="#F8FAFC"),
                        )
                    ]
                )
                fig_age_donut.update_layout(
                    title=dict(
                        text="⏳ Overdue Works by Delay Duration",
                        font=dict(size=14, color="#F8FAFC"),
                    ),
                    height=340,
                    margin=dict(l=10, r=10, t=40, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.22,
                        xanchor="center",
                        x=0.5,
                        font=dict(color="#CBD5E1", size=10),
                    ),
                    annotations=[
                        dict(
                            text=f"<b>{tot_od_an:,}</b><br>Overdue",
                            x=0.5,
                            y=0.5,
                            font_size=14,
                            font_color="#FCA5A5",
                            showarrow=False,
                        )
                    ],
                )
                st.plotly_chart(fig_age_donut, use_container_width=True)

    with pbi_c3:
        with st.container(border=True):
            div_grp = (
                an_scope_df.groupby("division_name", as_index=False)
                .agg(
                    total_works=("is_overdue", "count"),
                    overdue_works=("is_overdue", "sum"),
                    ets_works=("is_status_ets", "sum"),
                )
                .sort_values(by="overdue_works", ascending=True)
            )
            fig_div_bar = go.Figure()
            fig_div_bar.add_trace(
                go.Bar(
                    y=div_grp["division_name"],
                    x=div_grp["overdue_works"],
                    name="Overdue ('P')",
                    orientation="h",
                    marker_color="#EF4444",
                    text=div_grp["overdue_works"],
                    textposition="outside",
                )
            )
            fig_div_bar.add_trace(
                go.Bar(
                    y=div_grp["division_name"],
                    x=div_grp["ets_works"],
                    name="In 'ETS'",
                    orientation="h",
                    marker_color="#10B981",
                    text=div_grp["ets_works"],
                    textposition="outside",
                )
            )
            fig_div_bar.update_layout(
                barmode="group",
                title=dict(
                    text="🏛️ Division-Wise Overdue ('P') vs. 'ETS' Works",
                    font=dict(size=14, color="#F8FAFC"),
                ),
                height=340,
                margin=dict(l=10, r=35, t=40, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(15,23,42,0.45)",
                font=dict(color="#E2E8F0", size=11),
                legend=dict(orientation="h", y=1.02, x=0),
                xaxis=dict(gridcolor="rgba(148,163,184,0.12)"),
            )
            st.plotly_chart(fig_div_bar, use_container_width=True)

    # --- Row 2 of Visuals: Scheme (H.O.A.) Bottlenecks & Risk Quadrant Scatter ---
    pbi_r2_c1, pbi_r2_c2 = st.columns([1, 1.15], gap="small")

    with pbi_r2_c1:
        with st.container(border=True):
            hoa_grp = (
                an_scope_df.groupby("sheet_name", as_index=False)
                .agg(
                    total_works=("is_overdue", "count"),
                    overdue_works=("is_overdue", "sum"),
                    ets_works=("is_status_ets", "sum"),
                )
            )
            hoa_grp = hoa_grp[
                (hoa_grp["overdue_works"] > 0) | (hoa_grp["ets_works"] > 0)
            ].sort_values(by="overdue_works", ascending=True)

            fig_hoa = go.Figure()
            fig_hoa.add_trace(
                go.Bar(
                    y=hoa_grp["sheet_name"],
                    x=hoa_grp["overdue_works"],
                    name="Overdue ('P')",
                    orientation="h",
                    marker_color="#EF4444",
                )
            )
            fig_hoa.add_trace(
                go.Bar(
                    y=hoa_grp["sheet_name"],
                    x=hoa_grp["ets_works"],
                    name="In 'ETS'",
                    orientation="h",
                    marker_color="#10B981",
                )
            )
            fig_hoa.update_layout(
                barmode="stack",
                title=dict(
                    text="📑 Scheme / Worksheet (H.O.A.) Delay Bottlenecks",
                    font=dict(size=14, color="#F8FAFC"),
                ),
                height=420,
                margin=dict(l=10, r=25, t=40, b=20),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(15,23,42,0.45)",
                font=dict(color="#E2E8F0", size=11),
                legend=dict(orientation="h", y=1.02, x=0),
                xaxis=dict(
                    title="Delayed Works Count",
                    gridcolor="rgba(148,163,184,0.12)",
                ),
            )
            st.plotly_chart(fig_hoa, use_container_width=True)

    with pbi_r2_c2:
        with st.container(border=True):
            scatter_df = ae_stats_df[ae_stats_df["total_works"] >= min_works_threshold].copy()
            if not scatter_df.empty:
                med_x = float(scatter_df["total_works"].median())
                med_y = float(scatter_df["overdue_pct"].median())

                fig_quad = go.Figure()
                fig_quad.add_trace(
                    go.Scatter(
                        x=scatter_df["total_works"],
                        y=scatter_df["overdue_pct"],
                        mode="markers+text",
                        text=scatter_df.apply(
                            lambda r: r[LABEL_AE_EN]
                            if r["overdue_works"] >= scatter_df["overdue_works"].nlargest(6).min()
                            else "",
                            axis=1,
                        ),
                        textposition="top center",
                        textfont=dict(size=10, color="#FCA5A5"),
                        marker=dict(
                            size=scatter_df["overdue_works"].clip(lower=4) * 1.4 + 6,
                            color=scatter_df["overdue_works"],
                            colorscale="YlOrRd",
                            showscale=True,
                            colorbar=dict(title="Overdue"),
                            line=dict(width=1, color="#F8FAFC"),
                        ),
                        customdata=scatter_df[
                            [LABEL_AE_EN, "overdue_works", "ets_works", "primary_div"]
                        ].values,
                        hovertemplate=(
                            "<b>%{customdata[0]}</b> (%{customdata[3]})<br>"
                            "Total Works: %{x}<br>"
                            "Overdue %: %{y:.1f}%<br>"
                            "Overdue Count: %{customdata[1]}<br>"
                            "ETS Count: %{customdata[2]}<extra></extra>"
                        ),
                    )
                )
                fig_quad.add_vline(
                    x=med_x,
                    line_dash="dash",
                    line_color="rgba(148,163,184,0.45)",
                )
                fig_quad.add_hline(
                    y=med_y,
                    line_dash="dash",
                    line_color="rgba(148,163,184,0.45)",
                )
                fig_quad.update_layout(
                    title=dict(
                        text="🎯 AE Risk Matrix: Total Workload vs. Overdue % (Bubble Size = Overdue Count)",
                        font=dict(size=14, color="#F8FAFC"),
                    ),
                    height=420,
                    margin=dict(l=10, r=10, t=40, b=20),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(15,23,42,0.45)",
                    font=dict(color="#E2E8F0", size=11),
                    xaxis=dict(
                        title="Total Works Assigned",
                        gridcolor="rgba(148,163,184,0.12)",
                    ),
                    yaxis=dict(
                        title="Overdue Percentage (%)",
                        gridcolor="rgba(148,163,184,0.12)",
                    ),
                )
                st.plotly_chart(fig_quad, use_container_width=True)

    st.stop()


# =============================================================================
# PAGE 1: EXISTING MAIN OVERDUE WORKS DASHBOARD (PC VERSION)
# =============================================================================
unique_divisions = sorted(
    [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
)
division_options = [ALL_OPTION] + unique_divisions

f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5, gap="small")

with f_col1:
    with st.container(border=True):
        selected_division = st.selectbox(
            f"🏛️ Select {LABEL_DIV_EN}:",
            options=division_options,
            index=0,
        )

df_step1 = (
    master_df[master_df["division_name"] == selected_division].copy()
    if selected_division != ALL_OPTION
    else master_df.copy()
)

subdivision_options = build_dropdown_options(df_step1, "subdivision_name")
with f_col2:
    with st.container(border=True):
        selected_subdivision = st.selectbox(
            f"🏢 Select {LABEL_SUBDIV_EN}:",
            options=subdivision_options,
            index=0,
        )

df_step2 = (
    df_step1[df_step1["subdivision_name"] == selected_subdivision].copy()
    if selected_subdivision != ALL_OPTION
    else df_step1.copy()
)

ee_options = build_dropdown_options(df_step2, "ee_name")
with f_col3:
    with st.container(border=True):
        selected_ee = st.selectbox(
            f"👔 Select {LABEL_EE_EN}:",
            options=ee_options,
            index=0,
        )

df_step3 = (
    df_step2[df_step2["ee_name"] == selected_ee].copy()
    if selected_ee != ALL_OPTION
    else df_step2.copy()
)

ta_options = build_dropdown_options(df_step3, "ta_name")
with f_col4:
    with st.container(border=True):
        selected_ta = st.selectbox(
            f"🛠️ Select {LABEL_TA_EN}:",
            options=ta_options,
            index=0,
        )

df_step4 = (
    df_step3[df_step3["ta_name"] == selected_ta].copy()
    if selected_ta != ALL_OPTION
    else df_step3.copy()
)

as_options = build_dropdown_options(df_step4, "as_name")
with f_col5:
    with st.container(border=True):
        selected_as = st.selectbox(
            f"📑 Select {LABEL_AS_EN}:",
            options=as_options,
            index=0,
        )

filtered_master_df = (
    df_step4[df_step4["as_name"] == selected_as].copy()
    if selected_as != ALL_OPTION
    else df_step4.copy()
)

ctrl_col1, ctrl_col2 = st.columns(2, gap="small")
with ctrl_col1:
    with st.container(border=True):
        view_choice = st.radio(
            "👤 Select Officer View:",
            options=[LABEL_AE_EN, LABEL_AEE_EN],
            horizontal=True,
        )
with ctrl_col2:
    with st.container(border=True):
        sort_choice = st.radio(
            "⬇️ Sort Chart & Table By:",
            options=[
                "Highest Overdue Count at Top",
                "Highest Overdue Percentage (%) at Top",
            ],
            horizontal=True,
        )

group_col = "so_name" if view_choice == LABEL_AE_EN else "aee_name"
full_filter_col = "_so_name" if view_choice == LABEL_AE_EN else "_aee_name"

valid_df = filtered_master_df[filtered_master_df[group_col] != ""].copy()

if not show_unassigned_bar:
    valid_df = valid_df[valid_df[group_col] != UNASSIGNED_LABEL].copy()

if denominator_mode.startswith("Only Ongoing Works"):
    valid_df = valid_df[valid_df["is_status_p"]].copy()

if valid_df.empty:
    st.warning("No works found matching the selected filters.")
    st.stop()

summary_df = (
    valid_df.groupby(group_col, as_index=False)
    .agg(
        total_works=("is_overdue", "count"),
        ongoing_p_works=("is_status_p", "sum"),
        ets_works=("is_status_ets", "sum"),
        overdue_works=("is_overdue", "sum"),
    )
    .rename(columns={group_col: view_choice})
)

summary_df["rest_works"] = (
    summary_df["total_works"] - summary_df["overdue_works"] - summary_df["ets_works"]
).clip(lower=0)

summary_df["overdue_pct"] = (
    (summary_df["overdue_works"] / summary_df["total_works"]) * 100
).round(1).fillna(0.0)

summary_df["_is_unassigned"] = summary_df[view_choice] == UNASSIGNED_LABEL

if sort_choice == "Highest Overdue Percentage (%) at Top":
    summary_df = summary_df.sort_values(
        by=["_is_unassigned", "overdue_pct", "overdue_works", "total_works"],
        ascending=[True, False, False, False],
    ).reset_index(drop=True)
else:
    summary_df = summary_df.sort_values(
        by=["_is_unassigned", "overdue_works", "overdue_pct", "total_works"],
        ascending=[True, False, False, False],
    ).reset_index(drop=True)

total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
total_ets_all = int(summary_df["ets_works"].sum())

if hide_zero_overdue:
    named_officers_count = int(
        ((~summary_df["_is_unassigned"]) & (summary_df["overdue_works"] > 0)).sum()
    )
else:
    named_officers_count = int((~summary_df["_is_unassigned"]).sum())

overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

active_scope_parts = [
    f"{LABEL_DIV_EN}: {selected_division}",
    f"{LABEL_SUBDIV_EN}: {selected_subdivision}",
    f"EE: {selected_ee}",
    f"TA: {selected_ta}",
    f"AS: {selected_as}",
]
filter_scope_label = " | ".join(active_scope_parts)

k1, k2, k3, k4, k5 = st.columns(5, gap="small")
with k1:
    st.metric("📅 Date (Today)", today_str)
with k2:
    st.metric(
        f"👥 {view_choice}s",
        f"{named_officers_count:,}",
        help="Active named officers in current view",
    )
with k3:
    st.metric("📋 Total Works", f"{total_assigned_all:,}")
with k4:
    st.metric(
        "🔴 Overdue Works ('Progress')",
        f"{total_overdue_all:,} ({overall_overdue_pct}%)",
    )
with k5:
    st.metric("🟢 'ETS' Works", f"{total_ets_all:,}")

# =============================================================================
# SECTION 1: HORIZONTAL BAR CHART (PC VIEW)
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📈 Overdue &amp; ETS Works by {html.escape(view_choice)} ({html.escape(sort_choice)})</div>
        <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if hide_zero_overdue:
    chart_source_df = summary_df[
        (summary_df["overdue_works"] > 0) | summary_df["_is_unassigned"]
    ].copy()
else:
    chart_source_df = summary_df.copy()

search_officer_query = st.text_input(
    f"🔍 Quick Search {view_choice} Name (Optional):",
    value="",
    placeholder=f"Type {view_choice} name to filter the bar chart...",
)

if search_officer_query.strip():
    q_str = search_officer_query.strip().lower()
    chart_filtered_df = chart_source_df[
        chart_source_df[view_choice]
        .astype(str)
        .str.lower()
        .str.contains(q_str, na=False)
    ].copy()
else:
    chart_filtered_df = chart_source_df.copy()

if chart_filtered_df.empty:
    st.info(f"No {view_choice} found matching '{search_officer_query}'.")
else:
    plot_df = chart_filtered_df.iloc[::-1].copy()

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["overdue_works"],
            name="Overdue Works ('P')",
            orientation="h",
            marker_color="#EF4444",
        )
    )
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["ets_works"],
            name="ETS Works ('ETS')",
            orientation="h",
            marker_color="#4ADE80",
        )
    )
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["rest_works"],
            name="Rest of Works (Other Statuses)",
            orientation="h",
            marker_color="rgba(148, 163, 184, 0.28)",
            text=[
                f"  <b>{int(od)} Overdue</b> | <span style='color:#4ADE80'>{int(ets)} ETS</span> / {int(tot)} Total ({pct:.1f}%)"
                for od, ets, tot, pct in zip(
                    plot_df["overdue_works"],
                    plot_df["ets_works"],
                    plot_df["total_works"],
                    plot_df["overdue_pct"],
                )
            ],
            textposition="outside",
            cliponaxis=False,
        )
    )

    chart_height = max(360, len(plot_df) * 34 + 90)
    fig.update_layout(
        barmode="stack",
        height=chart_height,
        margin=dict(l=10, r=240, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15, 23, 42, 0.45)",
        font=dict(color="#E2E8F0", size=12),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0
        ),
        xaxis=dict(gridcolor="rgba(148, 163, 184, 0.12)"),
    )
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# =============================================================================
# SECTION 2: COMBINED OVERDUE WORKS BY SELECTED OFFICER
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📂 Combined Overdue Works by {html.escape(view_choice)}</div>
        <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

officer_options = (
    chart_source_df[view_choice].tolist()
    if not chart_source_df.empty
    else summary_df[view_choice].tolist()
)

officer_stats_map = {
    row[view_choice]: (
        int(row["overdue_works"]),
        int(row["ets_works"]),
        int(row["total_works"]),
        float(row["overdue_pct"]),
    )
    for _, row in summary_df.iterrows()
}

selected_officer = st.selectbox(
    f"Select {view_choice} to view and download their Overdue Works:",
    options=officer_options,
    index=0 if officer_options else None,
    format_func=lambda name: (
        f"{name} ({officer_stats_map[name][0]} Overdue | "
        f"{officer_stats_map[name][1]} ETS / {officer_stats_map[name][2]} Total — "
        f"{officer_stats_map[name][3]:.1f}%)"
    ),
)

if selected_officer:
    officer_overdue_df = overdue_full_df[
        overdue_full_df[full_filter_col] == selected_officer
    ].copy()

    if selected_division != ALL_OPTION:
        officer_overdue_df = officer_overdue_df[
            officer_overdue_df["_division_name"] == selected_division
        ]
    if selected_subdivision != ALL_OPTION:
        officer_overdue_df = officer_overdue_df[
            officer_overdue_df["_subdivision_name"] == selected_subdivision
        ]
    if selected_ee != ALL_OPTION:
        officer_overdue_df = officer_overdue_df[
            officer_overdue_df["_ee_name"] == selected_ee
        ]
    if selected_ta != ALL_OPTION:
        officer_overdue_df = officer_overdue_df[
            officer_overdue_df["_ta_name"] == selected_ta
        ]
    if selected_as != ALL_OPTION:
        officer_overdue_df = officer_overdue_df[
            officer_overdue_df["_as_name"] == selected_as
        ]

    opposite_officer_col = (
        COL_AEE_KN if view_choice == LABEL_AE_EN else COL_SO_KN
    )
    opposite_officer_en = (
        LABEL_AEE_EN if view_choice == LABEL_AE_EN else LABEL_AE_EN
    )

    curated_cols = [
        COL_WORK_NAME_KN,
        COL_DIV_KN,
        COL_SUBDIV_KN,
        opposite_officer_col,
        LABEL_TA_EN,
        LABEL_AS_EN,
        LABEL_EE_EN,
        COL_START_DATE_KN,
        COL_TARGET_DATE_KN,
        "Overdue By",
        COL_HOA_KN,
    ]

    officer_overdue_df = officer_overdue_df[curated_cols].reset_index(drop=True)
    officer_overdue_df.rename(
        columns={
            COL_WORK_NAME_KN: "Work Name",
            COL_DIV_KN: LABEL_DIV_EN,
            COL_SUBDIV_KN: LABEL_SUBDIV_EN,
            opposite_officer_col: opposite_officer_en,
            COL_START_DATE_KN: "Start Date",
            COL_TARGET_DATE_KN: "Target Date",
            COL_HOA_KN: "H.O.A.",
        },
        inplace=True,
    )
    officer_overdue_df.insert(0, "Sl. No.", range(1, len(officer_overdue_df) + 1))

    if officer_overdue_df.empty:
        st.success(
            f"🎉 **{selected_officer}** has **0** overdue works in **{filter_scope_label}**!"
        )
    else:
        st.markdown(
            render_ce_overdue_html_table(officer_overdue_df),
            unsafe_allow_html=True,
        )
