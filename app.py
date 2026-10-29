import calendar
import html
import io
import os
import pickle
import re
import threading
import time
from datetime import date, datetime
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
if "active_page" not in st.session_state:
    st.session_state["active_page"] = "overdue"

st.set_page_config(
    page_title="MI Dashboard & Division Abstract",
    page_icon="📊",
    layout="wide" if st.session_state["active_page"] == "abstract" else "centered",
    initial_sidebar_state="collapsed",
)

MAX_CONTAINER_WIDTH = (
    "100%" if st.session_state["active_page"] == "abstract" else "680px"
)

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Montserrat:wght@700;800&display=swap');

        header[data-testid="stHeader"] {{
            background: transparent !important;
        }}
        .stApp {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }}
        .block-container {{
            max-width: {MAX_CONTAINER_WIDTH} !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 2.1rem !important;
            padding-bottom: 2.2rem !important;
        }}

        /* 1. Compact Sidebar & Always-Visible Toggle Button */
        section[data-testid="stSidebar"][aria-expanded="true"] {{
            width: 220px !important;
            min-width: 220px !important;
            max-width: 220px !important;
            background: linear-gradient(180deg, rgba(15, 23, 42, 0.98) 0%, rgba(15, 23, 42, 0.94) 100%) !important;
            border-right: 1px solid rgba(148, 163, 184, 0.14) !important;
        }}
        [data-testid="stSidebarCollapseButton"],
        [data-testid="collapsedControl"],
        [data-testid="stExpandSidebarButton"] {{
            opacity: 1 !important;
            visibility: visible !important;
            display: flex !important;
            z-index: 999995 !important;
        }}
        [data-testid="stSidebarCollapseButton"] button,
        [data-testid="collapsedControl"] button,
        [data-testid="stExpandSidebarButton"] {{
            opacity: 1 !important;
            visibility: visible !important;
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.95), rgba(15, 23, 42, 0.98)) !important;
            border: 1.5px solid rgba(59, 130, 246, 0.75) !important;
            border-radius: 10px !important;
            color: #93C5FD !important;
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.35), 0 2px 6px rgba(0, 0, 0, 0.5) !important;
            width: 32px !important;
            height: 32px !important;
            min-height: 32px !important;
        }}

        /* 2. Top Header (Centered & Stylish Montserrat Font) */
        .premium-compact-header {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.55rem;
            padding: 0.25rem 0;
            margin-top: 0.05rem;
            margin-bottom: 0.35rem;
            background: transparent;
            border: none;
            width: 100%;
            overflow: hidden;
        }}
        .premium-icon-wrapper {{
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(59, 130, 246, 0.15);
            border: 1px solid rgba(59, 130, 246, 0.35);
            border-radius: 7px;
            width: 30px;
            height: 30px;
            font-size: 1.05rem;
            flex-shrink: 0;
            box-shadow: 0 0 8px rgba(59, 130, 246, 0.25);
        }}
        .premium-compact-title {{
            font-family: 'Montserrat', sans-serif !important;
            font-size: clamp(0.95rem, 4.1vw, 1.32rem) !important;
            font-weight: 800 !important;
            background: linear-gradient(90deg, #FFFFFF 0%, #CBD5E1 100%) !important;
            -webkit-background-clip: text !important;
            -webkit-text-fill-color: transparent !important;
            margin: 0 !important;
            line-height: 1.25 !important;
            letter-spacing: -0.2px !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
        }}

        /* Status Bar */
        .mi-status-bar {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 0.35rem 0.6rem;
            padding: 0.38rem 0.65rem;
            margin-bottom: 0.6rem;
            border-radius: 10px;
            background: rgba(30, 41, 59, 0.45);
            border: 1px solid rgba(148, 163, 184, 0.14);
            font-size: 0.73rem;
            color: #CBD5E1;
        }}
        .mi-status-pill {{
            display: inline-flex;
            align-items: center;
            gap: 0.25rem;
            padding: 0.1rem 0.45rem;
            border-radius: 999px;
            background: rgba(59, 130, 246, 0.15);
            border: 1px solid rgba(59, 130, 246, 0.35);
            color: #93C5FD;
            font-weight: 600;
            font-size: 0.71rem;
        }}

        /* 3. Control Cards */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            border-radius: 12px !important;
            padding: 0.45rem 0.6rem !important;
            margin-bottom: 0.45rem !important;
            background: linear-gradient(145deg, rgba(30, 41, 59, 0.48), rgba(15, 23, 42, 0.68)) !important;
            border: 1px solid rgba(148, 163, 184, 0.18) !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2) !important;
        }}

        /* Prevent virtual keyboard from popping up on mobile dropdowns */
        div[data-baseweb="select"] input {{
            pointer-events: none !important;
            caret-color: transparent !important;
            user-select: none !important;
            -webkit-user-select: none !important;
        }}

        /* 4. 5-Box Mobile KPI Grid */
        .kpi-mobile-top-row {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 7px;
            margin-top: 6px;
            margin-bottom: 7px;
        }}
        .kpi-mobile-bottom-row {{
            display: grid;
            grid-template-columns: 1.15fr 0.85fr;
            gap: 7px;
            margin-bottom: 10px;
        }}
        .kpi-m-card {{
            border-radius: 11px;
            padding: 8px 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.22);
        }}
        .kpi-blue {{
            border: 1px solid rgba(59, 130, 246, 0.25);
            border-top: 3px solid #3B82F6;
            background: linear-gradient(145deg, rgba(30, 58, 138, 0.40), rgba(15, 23, 42, 0.88));
        }}
        .kpi-violet {{
            border: 1px solid rgba(139, 92, 246, 0.25);
            border-top: 3px solid #8B5CF6;
            background: linear-gradient(145deg, rgba(88, 28, 135, 0.38), rgba(15, 23, 42, 0.88));
        }}
        .kpi-cyan {{
            border: 1px solid rgba(6, 182, 212, 0.25);
            border-top: 3px solid #06B6D4;
            background: linear-gradient(145deg, rgba(22, 78, 99, 0.42), rgba(15, 23, 42, 0.88));
        }}
        .kpi-red {{
            border: 1px solid rgba(239, 68, 68, 0.30);
            border-top: 3px solid #EF4444;
            background: linear-gradient(145deg, rgba(127, 29, 29, 0.42), rgba(15, 23, 42, 0.90));
        }}
        .kpi-green {{
            border: 1px solid rgba(74, 222, 128, 0.28);
            border-top: 3px solid #4ADE80;
            background: linear-gradient(145deg, rgba(6, 78, 59, 0.40), rgba(15, 23, 42, 0.90));
        }}
        .kpi-m-label {{
            font-size: 0.69rem;
            font-weight: 600;
            color: #CBD5E1;
            margin-bottom: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .kpi-m-val {{
            font-size: 1.08rem;
            font-weight: 800;
            line-height: 1.2;
            color: #F8FAFC;
        }}

        /* 5. Section Header Banners */
        .exec-section-banner {{
            display: flex;
            flex-direction: column;
            gap: 0.35rem;
            padding: 0.55rem 0.75rem;
            margin-top: 0.25rem;
            margin-bottom: 0.6rem;
            border-radius: 10px;
            background: rgba(15, 23, 42, 0.65);
            border: 1px solid rgba(148, 163, 184, 0.16);
            border-left: 4px solid #3B82F6;
        }}
        .exec-section-title {{
            font-size: 1.05rem;
            font-weight: 700;
            color: #F8FAFC;
            line-height: 1.25;
        }}
        .exec-scope-badge {{
            font-size: 0.70rem;
            font-weight: 600;
            padding: 0.2rem 0.55rem;
            border-radius: 6px;
            background: rgba(59, 130, 246, 0.12);
            border: 1px solid rgba(96, 165, 250, 0.28);
            color: #93C5FD;
            line-height: 1.35;
            word-break: break-word;
        }}

        /* 6. SECTION 1: Mobile Officer Scorecards */
        .scorecard-list-wrapper {{
            margin-bottom: 0.5rem;
        }}
        .officer-scorecard {{
            background: linear-gradient(145deg, rgba(30, 41, 59, 0.52), rgba(15, 23, 42, 0.85));
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 11px;
            padding: 10px 12px;
            margin-bottom: 8px;
            box-shadow: 0 3px 10px rgba(0, 0, 0, 0.2);
        }}
        .officer-scorecard-unassigned {{
            background: linear-gradient(145deg, rgba(69, 26, 3, 0.35), rgba(15, 23, 42, 0.88));
            border: 1px solid rgba(251, 191, 36, 0.32);
        }}
        .sc-top-row {{
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 8px;
            margin-bottom: 7px;
        }}
        .sc-name-group {{
            display: flex;
            align-items: center;
            gap: 7px;
            min-width: 0;
        }}
        .sc-rank-badge {{
            background: rgba(59, 130, 246, 0.18);
            border: 1px solid rgba(59, 130, 246, 0.4);
            color: #93C5FD;
            font-size: 0.72rem;
            font-weight: 800;
            padding: 2px 6px;
            border-radius: 6px;
            flex-shrink: 0;
        }}
        .sc-officer-name {{
            font-size: 0.90rem;
            font-weight: 700;
            color: #F8FAFC;
            line-height: 1.3;
            word-break: break-word;
        }}
        .sc-overdue-main-pill {{
            background: rgba(239, 68, 68, 0.18);
            border: 1px solid rgba(239, 68, 68, 0.45);
            color: #FCA5A5;
            font-size: 0.78rem;
            font-weight: 800;
            padding: 3px 8px;
            border-radius: 999px;
            white-space: nowrap;
            flex-shrink: 0;
        }}
        .sc-metrics-row {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 6px 10px;
            font-size: 0.75rem;
            color: #CBD5E1;
            margin-bottom: 7px;
        }}
        .sc-pill-ets {{ color: #4ADE80; font-weight: 700; }}
        .sc-pill-rest {{ color: #94A3B8; font-weight: 600; }}
        .sc-pill-total {{ color: #38BDF8; font-weight: 700; }}
        .sc-bar-track {{
            width: 100%;
            height: 6px;
            background: rgba(148, 163, 184, 0.20);
            border-radius: 999px;
            overflow: hidden;
            display: flex;
        }}
        .sc-bar-od {{ background: #EF4444; height: 100%; }}
        .sc-bar-ets {{ background: #4ADE80; height: 100%; }}

        /* 7. SECTION 2: Mobile Work Dossier Cards */
        .work-cards-wrapper {{
            margin-top: 0.35rem;
            margin-bottom: 0.6rem;
        }}
        .mobile-work-card {{
            background: linear-gradient(155deg, rgba(30, 41, 59, 0.58), rgba(15, 23, 42, 0.90));
            border: 1px solid rgba(148, 163, 184, 0.22);
            border-top: 3px solid #EF4444;
            border-radius: 12px;
            padding: 11px 12px;
            margin-bottom: 10px;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.25);
        }}
        .mwc-top-bar {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 6px;
            margin-bottom: 8px;
        }}
        .mwc-sl-badge {{
            background: rgba(148, 163, 184, 0.16);
            color: #E2E8F0;
            font-size: 0.74rem;
            font-weight: 800;
            padding: 2px 8px;
            border-radius: 6px;
        }}
        .mwc-hoa-badge {{
            background: rgba(59, 130, 246, 0.16);
            border: 1px solid rgba(59, 130, 246, 0.38);
            color: #93C5FD;
            font-size: 0.73rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
        }}
        .mwc-overdue-by-badge {{
            background: rgba(239, 68, 68, 0.16);
            border: 1px solid rgba(239, 68, 68, 0.42);
            color: #FCA5A5;
            font-size: 0.74rem;
            font-weight: 800;
            padding: 2px 8px;
            border-radius: 6px;
        }}
        .mwc-work-title {{
            font-size: 0.88rem;
            font-weight: 700;
            color: #F8FAFC;
            line-height: 1.42;
            margin-bottom: 10px;
            padding-bottom: 8px;
            border-bottom: 1px solid rgba(148, 163, 184, 0.14);
            white-space: normal;
            word-wrap: break-word;
        }}
        .mwc-meta-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 7px 10px;
            font-size: 0.75rem;
        }}
        .mwc-meta-item {{
            background: rgba(15, 23, 42, 0.55);
            border: 1px solid rgba(148, 163, 184, 0.12);
            border-radius: 8px;
            padding: 5px 8px;
            min-width: 0;
        }}
        .mwc-meta-label {{
            font-size: 0.67rem;
            font-weight: 600;
            color: #94A3B8;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            margin-bottom: 1px;
        }}
        .mwc-meta-val {{
            font-size: 0.78rem;
            font-weight: 700;
            color: #E2E8F0;
            word-break: break-word;
        }}
        .mwc-meta-val-red {{
            font-size: 0.79rem;
            font-weight: 800;
            color: #FCA5A5;
        }}

        /* 8. PAGE 2: 32-Column Division-Wise Abstract Table Styling */
        .abs-filter-header {{
            background: linear-gradient(90deg, #075985 0%, #0369A1 100%);
            color: #FFFFFF;
            padding: 7px 12px;
            border-radius: 8px 8px 0 0;
            font-weight: 700;
            font-size: 0.82rem;
            letter-spacing: 0.3px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        .abs-table-container {{
            width: 100%;
            max-height: 74vh;
            overflow: auto;
            -webkit-overflow-scrolling: touch;
            border: 2px solid #334155;
            border-radius: 10px;
            background: #FFFFFF;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
            margin-top: 0.4rem;
            margin-bottom: 1rem;
        }}
        table.abs-matrix-table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            font-family: 'Inter', Arial, sans-serif;
            font-size: 0.76rem;
            color: #0F172A;
            background: #FFFFFF;
        }}
        table.abs-matrix-table th,
        table.abs-matrix-table td {{
            border-right: 1px solid #94A3B8;
            border-bottom: 1px solid #94A3B8;
            padding: 5px 7px;
            text-align: center;
            vertical-align: middle;
        }}
        table.abs-matrix-table thead tr.abs-hdr-row-1 th {{
            position: sticky;
            top: 0;
            z-index: 30;
            background: #CFE2F3;
            color: #000000;
            font-weight: 800;
            font-size: 0.73rem;
            line-height: 1.25;
            min-width: 68px;
            max-width: 135px;
            white-space: normal;
            border-bottom: 1px solid #475569;
        }}
        table.abs-matrix-table thead tr.abs-hdr-row-2 th {{
            position: sticky;
            top: 54px;
            z-index: 29;
            background: #9FC5E8;
            color: #000000;
            font-weight: 800;
            font-size: 0.71rem;
            padding: 3px 6px;
            border-bottom: 2px solid #1E293B;
        }}
        /* Sticky First 3 Columns on Desktop/Tablet for Easy Reading */
        @media (min-width: 768px) {{
            table.abs-matrix-table th:nth-child(1),
            table.abs-matrix-table td:nth-child(1) {{
                position: sticky;
                left: 0;
                z-index: 20;
                min-width: 44px;
                max-width: 44px;
            }}
            table.abs-matrix-table th:nth-child(2),
            table.abs-matrix-table td:nth-child(2) {{
                position: sticky;
                left: 44px;
                z-index: 20;
                min-width: 110px;
                max-width: 110px;
            }}
            table.abs-matrix-table th:nth-child(3),
            table.abs-matrix-table td:nth-child(3) {{
                position: sticky;
                left: 154px;
                z-index: 20;
                min-width: 135px;
                max-width: 155px;
                border-right: 2px solid #334155;
            }}
            table.abs-matrix-table thead tr.abs-hdr-row-1 th:nth-child(1),
            table.abs-matrix-table thead tr.abs-hdr-row-1 th:nth-child(2),
            table.abs-matrix-table thead tr.abs-hdr-row-1 th:nth-child(3),
            table.abs-matrix-table thead tr.abs-hdr-row-2 th:nth-child(1),
            table.abs-matrix-table thead tr.abs-hdr-row-2 th:nth-child(2),
            table.abs-matrix-table thead tr.abs-hdr-row-2 th:nth-child(3) {{
                z-index: 40;
            }}
        }}
        tr.abs-detail-row td {{
            background: #FFFFFF;
            color: #0F172A;
            font-weight: 600;
            white-space: nowrap;
        }}
        tr.abs-detail-row:nth-child(even) td {{
            background: #F8FAFC;
        }}
        tr.abs-detail-row:hover td {{
            background: #EFF6FF !important;
        }}
        tr.abs-total-row td {{
            background: #FFF2CC !important;
            color: #000000 !important;
            font-weight: 800 !important;
            border-top: 1.5px solid #334155 !important;
            border-bottom: 2px solid #334155 !important;
            white-space: nowrap;
        }}
        tr.abs-spacer-row td {{
            background: #D9EAD3 !important;
            color: #000000 !important;
            font-weight: 800 !important;
            font-size: 0.80rem !important;
            text-align: left !important;
            padding: 6px 10px !important;
            border-top: 2px solid #1E293B !important;
            border-bottom: 2px solid #1E293B !important;
        }}

        hr {{
            margin-top: 0.5rem !important;
            margin-bottom: 0.55rem !important;
            border-color: rgba(148, 163, 184, 0.15) !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# PREVENT MOBILE VIRTUAL KEYBOARD FROM POPPING UP ON DROPDOWN TAP
# -----------------------------------------------------------------------------
def disable_mobile_select_keyboard():
    components.html(
        """
        <script>
        (function() {
            const parentDoc = window.parent.document;

            function lockSelectInputs() {
                const inputs = parentDoc.querySelectorAll('div[data-baseweb="select"] input');
                inputs.forEach(input => {
                    input.readOnly = true;
                    input.setAttribute('readonly', 'readonly');
                    input.setAttribute('inputmode', 'none');
                    input.setAttribute('autocomplete', 'off');
                    input.setAttribute('autocorrect', 'off');
                    input.setAttribute('autocapitalize', 'off');
                    input.setAttribute('spellcheck', 'false');
                });
            }

            lockSelectInputs();

            if (!window.parent.__miSelectTouchLockAdded) {
                parentDoc.addEventListener('touchstart', lockSelectInputs, { capture: true, passive: true });
                parentDoc.addEventListener('pointerdown', lockSelectInputs, { capture: true, passive: true });
                parentDoc.addEventListener('mousedown', lockSelectInputs, { capture: true, passive: true });
                window.parent.__miSelectTouchLockAdded = true;
            }

            if (!window.parent.__miSelectKeyboardObserver) {
                const observer = new MutationObserver(lockSelectInputs);
                observer.observe(parentDoc.body, { childList: true, subtree: true });
                window.parent.__miSelectKeyboardObserver = observer;
            }
        })();
        </script>
        """,
        height=0,
        width=0,
    )


# -----------------------------------------------------------------------------
# CONSTANTS & TARGET SHEETS
# -----------------------------------------------------------------------------
TARGET_SHEETS = [
    "4702_MOT",
    "4702-New Tanks",
    "4702-LIS",
    "4702_A&P",
    "PMKSY-SMI-RRR",
    "4702_Paschima Vahini",
    "4711__Other Flood Control",
    "4702-Kharland",
    "NABARD - New Tank",
    "SDMF",
    "KMERC",
    "NABARD - A&P",
    "ATAL BHUJAL",
    "NABARD_MOT",
    "NABARD - LIS",
    "other dept funds",
    "SCP (LIS)",
    "SCP (Borewell)",
    "SCP (Checkdam)",
    "TSP (Checkdam)",
    "TSP (Borewell)",
    "TSP (LIS)",
    "NABARD - Other Flood Control",
]

# Exact 19-sheet configuration from the GSQL Division-Wise Abstract formula
# Tuple format: (sheet_name, start_row_1based, end_row_1based, mode_str)
GSQL_ABSTRACT_SPECS = [
    ("4702_MOT", 4, 4004, "AG_AH"),
    ("4702-New Tanks", 2, 4000, "AG_AH"),
    ("4702_A&P", 2, 4002, "AG_AH"),
    ("4702-LIS", 1, 3996, "AG_AH"),
    ("4702_Paschima Vahini", 2, 4000, "AG_AH"),
    ("4711__Other Flood Control", 2, 4001, "AG_AH"),
    ("4702-Kharland", 2, 3999, "AG_AH"),
    ("NABARD - A&P", 4, 4002, "AG_AH"),
    ("TSP (LIS)", 2, 3999, "AH_AI"),
    ("SCP (Checkdam)", 2, 3999, "AG_AH"),
    ("SCP (LIS)", 2, 3999, "AI_AJ"),
    ("TSP (Borewell)", 2, 3998, "AM_AN"),
    ("SCP (Borewell)", 4, 4015, "SCP_BW"),
    ("TSP (Checkdam)", 2, 3999, "AG_AH"),
    ("NABARD - Other Flood Control", 2, 4000, "AG_AH"),
    ("NABARD - LIS", 2, 4000, "AG_AH"),
    ("NABARD_MOT", 4, 4005, "AG_AH"),
    ("PMKSY-SMI-RRR", 2, 4000, "AG_AH"),
    ("KMERC", 4, 4003, "AG_AH"),
]

ABSTRACT_HEADERS_32 = [
    "ಕ್ರ.ಸಂ",
    "ವಲಯ / ವಿಭಾಗ",
    "ಲೆಕ್ಕಶೀರ್ಷಿಕೆ",
    "ಮುಂದುವರೆದ ಕಾಮಗಾರಿಗಳ ಸಂಖ್ಯೆ",
    "ಅಂದಾಜು ಮೊತ್ತ",
    "ಹೊಸ ಕಾಮಗಾರಿಗಳ ಸಂಖ್ಯೆ",
    "ಅಂದಾಜು ಮೊತ್ತ",
    "ಒಟ್ಟು ಕಾಮಗಾರಿಗಳ ಸಂಖ್ಯೆ",
    "ಒಟ್ಟು ಅಂದಾಜು ಮೊತ್ತ",
    "ಗುತ್ತಿಗೆ ಮೊತ್ತ (ಜಿ.ಎಸ್.ಟಿ ಸೇರಿದಂತೆ)",
    "ಪರಿಷ್ಕೃತ ಅಂದಾಜು ಮೊತ್ತ (ಜಿ.ಎಸ್.ಟಿ ಸೇರಿದಂತೆ)",
    "2026-27ನೇ ಸಾಲಿಗೆ ಹಂಚಿಕೆಯಾದ ಅನುದಾನ",
    "ಮಾರ್ಚ್ 2026 ವರೆಗಿನ ವೆಚ್ಚ",
    "2026-27ನೇ ಸಾಲಿನ ವೆಚ್ಚ",
    "ಒಟ್ಟಾರೆ ವೆಚ್ಚ",
    "ಬಾಕಿ ಬಿಲ್ಲಿನ ಮೊತ್ತ",
    "ಕಾಮಗಾರಿಯನ್ನು ಪೂರ್ಣಗೊಳಿಸಲು ಬೇಕಾಗುವ ಬಾಕಿ ಮೊತ್ತ",
    "CP",
    "CF",
    "P",
    "TBS",
    "TP",
    "ETS",
    "OTH",
    "CBW",
    "CEF",
    "PBW",
    "CBWPM",
    "RRR",
    "BWF",
    "Overdue %",
    "% of (TBS, OTH, TP, ETS)",
]

COL_WORK_NAME_KN = "Work Name"
COL_DIV_KN = "ವಿಭಾಗ"
COL_SUBDIV_KN = "ಉಪವಿಭಾಗ"
COL_SO_KN = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE_KN = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_START_DATE_KN = "ಕಾಮಗಾರಿ ಪ್ರಾರಂಭಿಸಿದ ದಿನಾಂಕ"
COL_TARGET_DATE_KN = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE_KN = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"
COL_HOA_KN = "ಲೆಕ್ಕಶೀರ್ಷಿಕೆ"

LABEL_DIV_EN = "Division"
LABEL_SUBDIV_EN = "Sub-division"
LABEL_AE_EN = "Assistant Engineer"
LABEL_AEE_EN = "Assistant Executive Engineer"
LABEL_EE_EN = "Executive Engineer"
LABEL_TA_EN = "Technical Assistant"
LABEL_AS_EN = "Account Superintendent"

ALL_OPTION = "ALL"
UNSPECIFIED_LABEL = "Not Specified (Empty)"
UNASSIGNED_LABEL = "Unassigned (Empty)"

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

SCHEMA_VERSION = "v18_ce_dual_page_abstract"
DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v18_mob.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state_v18_mob() -> dict:
    state = {
        "version": SCHEMA_VERSION,
        "lock": threading.Lock(),
        "data": None,
        "last_synced": None,
        "source_label": None,
    }
    if os.path.exists(DISK_BACKUP_FILE):
        try:
            with open(DISK_BACKUP_FILE, "rb") as f:
                saved = pickle.load(f)
                if saved.get("version") == SCHEMA_VERSION:
                    state["data"] = saved.get("data")
                    state["last_synced"] = saved.get("last_synced")
                    state["source_label"] = saved.get("source_label")
        except Exception:
            pass
    return state


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def shift_months(d: date, months: int) -> date:
    """Equivalent to Excel/Google Sheets EDATE(d, months)."""
    total_m = (d.year * 12 + (d.month - 1)) + months
    ny = total_m // 12
    nm = (total_m % 12) + 1
    max_d = calendar.monthrange(ny, nm)[1]
    return date(ny, nm, min(d.day, max_d))


def calc_overdue_duration(target_date_obj, current_date_obj) -> str:
    if not isinstance(target_date_obj, date) or target_date_obj >= current_date_obj:
        return "-"

    y = current_date_obj.year - target_date_obj.year
    m = current_date_obj.month - target_date_obj.month
    d = current_date_obj.day - target_date_obj.day + 1

    if d <= 0:
        m -= 1
        prev_month = current_date_obj.month - 1 if current_date_obj.month > 1 else 12
        prev_year = (
            current_date_obj.year
            if current_date_obj.month > 1
            else current_date_obj.year - 1
        )
        d += calendar.monthrange(prev_year, prev_month)[1]

    if m < 0:
        y -= 1
        m += 12

    days_in_current_month = calendar.monthrange(
        current_date_obj.year, current_date_obj.month
    )[1]
    if d > days_in_current_month:
        d -= days_in_current_month
        m += 1
        if m >= 12:
            m -= 12
            y += 1

    parts = []
    if y > 0:
        parts.append(f"{y} year{'s' if y > 1 else ''}")
    if m > 0:
        parts.append(f"{m} month{'s' if m > 1 else ''}")
    if d > 0:
        parts.append(f"{d} day{'s' if d > 1 else ''}")

    return "(" + ", ".join(parts) + ")" if parts else "(0 days)"


def to_safe_python_date(val):
    if val is None:
        return None
    if isinstance(val, pd.Timestamp):
        return date(val.year, val.month, val.day)
    if isinstance(val, datetime):
        return date(val.year, val.month, val.day)
    if isinstance(val, date):
        return val
    return val


def sanitize_dataframe_cells(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_rows = [
        [to_safe_python_date(cell) for cell in row]
        for row in df.to_numpy().tolist()
    ]
    return pd.DataFrame(cleaned_rows, dtype=object)


def read_sheet_safely(
    excel_file: pd.ExcelFile, workbook_bytes: bytes, sheet_name: str
) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=None, dtype=object
        )
        return sanitize_dataframe_cells(raw_df)
    except Exception:
        wb = openpyxl.load_workbook(
            io.BytesIO(workbook_bytes), data_only=True, read_only=True
        )
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            cleaned_row = [to_safe_python_date(cell) for cell in row]
            if any(c is not None and str(c).strip() != "" for c in cleaned_row):
                rows.append(cleaned_row)
            elif len(rows) < 35:
                rows.append(cleaned_row)
        wb.close()
        return pd.DataFrame(rows, dtype=object)


def normalize_text(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(val)).strip()


def to_safe_float(val) -> float:
    if val is None or isinstance(val, (date, datetime)):
        return 0.0
    if isinstance(val, (int, float)):
        try:
            return 0.0 if pd.isna(val) else float(val)
        except Exception:
            return 0.0
    s = str(val).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_single_date(raw_val):
    val = to_safe_python_date(raw_val)
    if val is None:
        return None
    if isinstance(val, date):
        return val
    s = normalize_text(val)
    if not s or s.upper() in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", ""}:
        return None
    s_clean = re.sub(r"[./]", "-", s).strip(" -")
    match = re.search(r"(\d{1,4})-(\d{1,2})-(\d{1,4})", s_clean)
    if match:
        p1, p2, p3 = int(match.group(1)), int(match.group(2)), int(match.group(3))
        try:
            if p1 > 31:
                return date(max(1, min(9999, p1)), p2, p3)
            yr = p3 + 2000 if p3 < 100 else p3
            return date(max(1, min(9999, yr)), p2, p1)
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, dayfirst=True, utc=True, errors="coerce")
        if not pd.isna(dt):
            return date(dt.year, dt.month, dt.day)
    except Exception:
        pass
    return None


def get_fin_year_label(d_obj) -> str:
    if not isinstance(d_obj, date):
        return ""
    if d_obj.month >= 4:
        return f"{d_obj.year}-{str(d_obj.year + 1)[-2:]}"
    return f"{d_obj.year - 1}-{str(d_obj.year)[-2:]}"


def format_cell_for_display(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return normalize_text(val)


def extract_spreadsheet_key(sheet_url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL.")
    return match.group(1)


def download_xlsx_bytes(sheet_url: str) -> bytes:
    key = extract_spreadsheet_key(sheet_url)
    urls = [
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx",
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx&id={key}",
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 14; Pixel 8) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0.0.0 Mobile Safari/537.36"
        ),
    }

    last_err = None
    for attempt in range(3):
        url = urls[attempt % len(urls)]
        if attempt > 0:
            url = f"{url}&t={int(time.time())}"
        try:
            resp = requests.get(url, headers=headers, timeout=120)
            resp.raise_for_status()
            if len(resp.content) > 1000:
                return resp.content
        except Exception as e:
            last_err = e
            time.sleep(2 * (attempt + 1))

    raise RuntimeError(f"Could not download workbook after 3 attempts: {last_err}")


def find_matching_column_indices(headers: list[str]) -> dict:
    col_idx_map = {}

    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_col = re.sub(r"[\s\-_]+", "", norm_col)

        if norm_col in {"ವಿಭಾಗ", "ವಿಭಾಗದ ಹೆಸರು"} and "division" not in col_idx_map:
            col_idx_map["division"] = idx
        elif (
            compact_col in {"ಉಪವಿಭಾಗ", "ಉಪವಿಭಾಗದಹೆಸರು"}
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx
        elif upper_col == "EXECUTIVE ENGINEER" and "ee" not in col_idx_map:
            col_idx_map["ee"] = idx
        elif upper_col == "TECHNICAL ASSISTANT" and "ta" not in col_idx_map:
            col_idx_map["ta"] = idx
        elif (
            upper_col in {"ACCOUNT SUPERINTENDENT", "ACCOUNTS SUPERINTENDENT"}
            and "as_off" not in col_idx_map
        ):
            col_idx_map["as_off"] = idx
        elif (
            norm_col
            in {
                "ಕಾಮಗಾರಿ ಹೆಸರು",
                "ಕಾಮಗಾರಿಯ ಹೆಸರು",
                "ಕಾಮಗಾರಿಗಳ ಹೆಸರು",
                "WORK NAME",
                "NAME OF WORK",
            }
            and "work_name" not in col_idx_map
        ):
            col_idx_map["work_name"] = idx
        elif (
            compact_col in {"ಲೆಕ್ಕಶೀರ್ಷಿಕೆ", "ಲೆಕ್ಕಶೀರ್ಷಿಕೆಹೆಸರು"}
            and "hoa" not in col_idx_map
        ):
            col_idx_map["hoa"] = idx

    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_upper = re.sub(r"\s+", "", upper_col)

        if (
            ("ಉಪವಿಭಾಗ" in norm_col or "ಉಪ ವಿಭಾಗ" in norm_col or "ಉಪ-ವಿಭಾಗ" in norm_col)
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx
        elif (
            "ವಿಭಾಗ" in norm_col
            and "ಉಪ" not in norm_col
            and "SUB" not in upper_col
            and "division" not in col_idx_map
        ):
            col_idx_map["division"] = idx
        elif (COL_SO_KN in norm_col or "ಶಾಖಾಧಿಕಾರಿ" in norm_col) and "so" not in col_idx_map:
            col_idx_map["so"] = idx
        elif (
            COL_AEE_KN in norm_col or "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col
        ) and "aee" not in col_idx_map:
            col_idx_map["aee"] = idx
        elif (
            ("EXECUTIVE ENGINEER" in upper_col and "ASSISTANT" not in upper_col)
            or ("ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col and "ಸಹಾಯಕ" not in norm_col)
        ) and "ee" not in col_idx_map:
            col_idx_map["ee"] = idx
        elif (
            "TECHNICAL ASSISTANT" in upper_col or "ತಾಂತ್ರಿಕ ಸಹಾಯಕ" in norm_col
        ) and "ta" not in col_idx_map:
            col_idx_map["ta"] = idx
        elif (
            "ACCOUNT SUPERINTENDENT" in upper_col
            or "ACCOUNTS SUPERINTENDENT" in upper_col
            or "ಲೆಕ್ಕ ಅಧೀಕ್ಷಕ" in norm_col
        ) and "as_off" not in col_idx_map:
            col_idx_map["as_off"] = idx
        elif (
            COL_TARGET_DATE_KN in norm_col or "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ" in norm_col
        ) and "target_date" not in col_idx_map:
            col_idx_map["target_date"] = idx
        elif (
            COL_ACTUAL_DATE_KN in norm_col
            and "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ" not in norm_col
            and "actual_date" not in col_idx_map
        ):
            col_idx_map["actual_date"] = idx
        elif (
            "ಪ್ರಾರಂಭಿಸಿದ ದಿನಾಂಕ" in norm_col
            or "ಪ್ರಾರಂಭವಾದ ದಿನಾಂಕ" in norm_col
            or "ಪ್ರಾರಂಭ ದಿನಾಂಕ" in norm_col
        ) and "start_date" not in col_idx_map:
            col_idx_map["start_date"] = idx
        elif (
            (
                "ಕಾಮಗಾರಿ ಹೆಸರು" in norm_col
                or "ಕಾಮಗಾರಿಯ ಹೆಸರು" in norm_col
                or "ಕಾಮಗಾರಿಗಳ ಹೆಸರು" in norm_col
            )
            and "ದಿನಾಂಕ" not in norm_col
            and "work_name" not in col_idx_map
        ):
            col_idx_map["work_name"] = idx
        elif (
            ("ಲೆಕ್ಕಶೀರ್ಷಿಕೆ" in norm_col or "ಲೆಕ್ಕ ಶೀರ್ಷಿಕೆ" in norm_col)
            and "hoa" not in col_idx_map
        ):
            col_idx_map["hoa"] = idx
        elif "CP/CF" in compact_upper and "status" not in col_idx_map:
            col_idx_map["status"] = idx

    if "work_name" not in col_idx_map:
        for idx, col in enumerate(headers):
            norm_col = normalize_text(col)
            upper_col = norm_col.upper()
            if (
                (
                    "ಕಾಮಗಾರಿ" in norm_col
                    or "WORK NAME" in upper_col
                    or "NAME OF WORK" in upper_col
                )
                and "ದಿನಾಂಕ" not in norm_col
                and "ಮೊತ್ತ" not in norm_col
                and "ಸಂಖ್ಯೆ" not in norm_col
                and "ಹಂತ" not in norm_col
                and idx not in col_idx_map.values()
            ):
                col_idx_map["work_name"] = idx
                break

    return col_idx_map


def parse_indian_date_list(series: pd.Series) -> list:
    return [parse_single_date(v) for v in series.tolist()]


def is_empty_date(series: pd.Series) -> pd.Series:
    results = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            results.append(True)
        elif isinstance(val, date):
            results.append(False)
        else:
            s = normalize_text(val)
            results.append(
                s.upper() in {"", "-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAT"}
            )
    return pd.Series(results, index=series.index, dtype=bool)


def is_valid_division_text(val: str) -> bool:
    if not val:
        return False
    upper = val.upper().strip()
    if upper in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAN", "NAT", "NULL", "0"}:
        return False
    if re.fullmatch(r"[\d\W_]+", val):
        return False
    if val in {COL_DIV_KN, "ವಿಭಾಗದ ಹೆಸರು", "ಒಟ್ಟು", "ಒಟ್ಟು ಮೊತ್ತ", "ಮಹಾ ಒಟ್ಟು"}:
        return False
    if upper in {"DIVISION", "DIVISION NAME", "TOTAL", "GRAND TOTAL", "SUB TOTAL"}:
        return False
    if val.startswith("ಒಟ್ಟು ") or val.endswith(" ಒಟ್ಟು") or upper.startswith("TOTAL "):
        return False
    return True


def map_optional_field(name: str) -> str:
    if not name:
        return UNSPECIFIED_LABEL
    upper = name.upper().strip()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or re.fullmatch(r"[\d\W_]+", name):
        return UNSPECIFIED_LABEL
    if (
        len(name) > 75
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or name in {COL_SUBDIV_KN, "ಉಪ ವಿಭಾಗ", "ಉಪವಿಭಾಗದ ಹೆಸರು"}
        or upper
        in {
            "EXECUTIVE ENGINEER",
            "TECHNICAL ASSISTANT",
            "ACCOUNT SUPERINTENDENT",
            "ACCOUNTS SUPERINTENDENT",
        }
    ):
        return UNSPECIFIED_LABEL
    return name


def map_officer_or_unassigned(name: str) -> str:
    if not name:
        return UNASSIGNED_LABEL
    upper = name.upper()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or name.isdigit():
        return UNASSIGNED_LABEL
    if (
        len(name) > 65
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or COL_SO_KN in name
        or COL_AEE_KN in name
    ):
        return UNASSIGNED_LABEL
    return name


def display_clean_cell(val: str) -> str:
    if not val or val in {UNSPECIFIED_LABEL, UNASSIGNED_LABEL}:
        return "-"
    return val


# -----------------------------------------------------------------------------
# MOBILE RENDERERS: 1) SCORECARDS & 2) DOSSIER CARDS
# -----------------------------------------------------------------------------
def render_mobile_officer_scorecards(df_source: pd.DataFrame, officer_col: str) -> str:
    cards_html = []
    rank_counter = 1

    for _, row in df_source.iterrows():
        raw_name = str(row[officer_col])
        is_unassigned = bool(row["_is_unassigned"])
        od_count = int(row["overdue_works"])
        ets_count = int(row["ets_works"])
        rest_count = int(row["rest_works"])
        tot_count = max(int(row["total_works"]), 1)
        od_pct = float(row["overdue_pct"])

        od_bar_w = min(100.0, round((od_count / tot_count) * 100, 1))
        ets_bar_w = min(100.0 - od_bar_w, round((ets_count / tot_count) * 100, 1))

        if is_unassigned:
            rank_badge = '<span class="sc-rank-badge" style="background:rgba(251,191,36,0.2);border-color:rgba(251,191,36,0.5);color:#FDE68A;">⚠️</span>'
            card_cls = "officer-scorecard officer-scorecard-unassigned"
        else:
            rank_badge = f'<span class="sc-rank-badge">#{rank_counter}</span>'
            card_cls = "officer-scorecard"
            rank_counter += 1

        esc_name = html.escape(raw_name)

        card_item = (
            f'<div class="{card_cls}">'
            f'<div class="sc-top-row">'
            f'<div class="sc-name-group">{rank_badge}<span class="sc-officer-name">{esc_name}</span></div>'
            f'<span class="sc-overdue-main-pill">🔴 {od_count} Overdue ({od_pct:.1f}%)</span>'
            f"</div>"
            f'<div class="sc-metrics-row">'
            f'<span class="sc-pill-ets">🟢 {ets_count} ETS</span>'
            f"<span>•</span>"
            f'<span class="sc-pill-rest">⚪ {rest_count} Rest</span>'
            f"<span>•</span>"
            f'<span class="sc-pill-total">📋 {int(row["total_works"])} Total Works</span>'
            f"</div>"
            f'<div class="sc-bar-track">'
            f'<div class="sc-bar-od" style="width:{od_bar_w}%;"></div>'
            f'<div class="sc-bar-ets" style="width:{ets_bar_w}%;"></div>'
            f"</div>"
            f"</div>"
        )
        cards_html.append(card_item)

    return f'<div class="scorecard-list-wrapper">{"".join(cards_html)}</div>'


def render_mobile_work_dossier_cards(
    df_display: pd.DataFrame, opposite_officer_en: str
) -> str:
    cards_html = []

    for _, row in df_display.iterrows():
        sl_no = html.escape(str(row.get("Sl. No.", "-")))
        work_name = html.escape(str(row.get("Work Name", "-") or "-"))
        div_val = html.escape(str(row.get(LABEL_DIV_EN, "-") or "-"))
        subdiv_val = html.escape(str(row.get(LABEL_SUBDIV_EN, "-") or "-"))
        opp_val = html.escape(str(row.get(opposite_officer_en, "-") or "-"))
        ee_val = html.escape(str(row.get(LABEL_EE_EN, "-") or "-"))
        ta_val = html.escape(str(row.get(LABEL_TA_EN, "-") or "-"))
        as_val = html.escape(str(row.get(LABEL_AS_EN, "-") or "-"))
        start_val = html.escape(str(row.get("Start Date", "-") or "-"))
        target_val = html.escape(str(row.get("Target Date", "-") or "-"))
        overdue_by_val = html.escape(str(row.get("Overdue By", "-") or "-"))
        hoa_val = html.escape(str(row.get("H.O.A.", "-") or "-"))

        opp_short_label = (
            "A.E.E."
            if opposite_officer_en == LABEL_AEE_EN
            else "A.E. / S.O."
        )

        card_html = (
            f'<div class="mobile-work-card">'
            f'<div class="mwc-top-bar">'
            f'<div><span class="mwc-sl-badge">Sl. #{sl_no}</span> <span class="mwc-hoa-badge">📑 {hoa_val}</span></div>'
            f'<span class="mwc-overdue-by-badge">⏳ {overdue_by_val}</span>'
            f"</div>"
            f'<div class="mwc-work-title">{work_name}</div>'
            f'<div class="mwc-meta-grid">'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">🚀 Start Date</div><div class="mwc-meta-val">{start_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">🎯 Target Date</div><div class="mwc-meta-val-red">{target_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">🏛️ Division</div><div class="mwc-meta-val">{div_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">🏢 Sub-division</div><div class="mwc-meta-val">{subdiv_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">👤 {html.escape(opp_short_label)}</div><div class="mwc-meta-val">{opp_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">👔 Exec. Engineer</div><div class="mwc-meta-val">{ee_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">🛠️ Tech. Assistant</div><div class="mwc-meta-val">{ta_val}</div></div>'
            f'<div class="mwc-meta-item"><div class="mwc-meta-label">📑 Account Supt.</div><div class="mwc-meta-val">{as_val}</div></div>'
            f"</div>"
            f"</div>"
        )
        cards_html.append(card_html)

    return f'<div class="work-cards-wrapper">{"".join(cards_html)}</div>'


# -----------------------------------------------------------------------------
# PAGE 2 ENGINE: GSQL DIVISION-WISE ABSTRACT PARSER & CALCULATOR
# -----------------------------------------------------------------------------
def extract_gsql_30col_slice(raw_df: pd.DataFrame, start_1b: int, end_1b: int, mode: str) -> pd.DataFrame:
    """Slices the exact rows & 30 columns specified in the Google Sheets LET() formula."""
    if raw_df.empty:
        return pd.DataFrame()

    s_idx = max(0, start_1b - 1)
    e_idx = min(len(raw_df), end_1b)
    if s_idx >= e_idx:
        return pd.DataFrame()

    sub = raw_df.iloc[s_idx:e_idx].copy()
    max_needed = 40
    if sub.shape[1] < max_needed:
        for c in range(sub.shape[1], max_needed):
            sub[c] = None

    if mode == "AG_AH":
        col_indices = list(range(0, 28)) + [32, 33]
    elif mode == "AH_AI":
        col_indices = list(range(0, 28)) + [33, 34]
    elif mode == "AI_AJ":
        col_indices = list(range(0, 28)) + [34, 35]
    elif mode == "AM_AN":
        col_indices = list(range(0, 28)) + [38, 39]
    elif mode == "SCP_BW":
        # {'SCP (Borewell)'!A4:AA4015, 'SCP (Borewell)'!AD4:AD4015, 'SCP (Borewell)'!AM4:AN4015}
        col_indices = list(range(0, 27)) + [29, 38, 39]
    else:
        col_indices = list(range(0, 28)) + [32, 33]

    sliced = sub.iloc[:, col_indices].copy()
    sliced.columns = [f"Col{i}" for i in range(1, 31)]
    return sliced


def build_abstract_dataset_from_frames(
    raw_sheets_map: dict, available_sheets: dict, today_ist_date: date
):
    """Builds the 19-sheet raw_dataset and reads 'Division_wise ABSTRACT' allocations."""
    frames = []
    for sh_name, start_r, end_r, mode in GSQL_ABSTRACT_SPECS:
        actual_name = available_sheets.get(normalize_text(sh_name).lower())
        if actual_name and actual_name in raw_sheets_map:
            sliced = extract_gsql_30col_slice(
                raw_sheets_map[actual_name], start_r, end_r, mode
            )
            if not sliced.empty:
                frames.append(sliced)

    if not frames:
        return pd.DataFrame(), pd.DataFrame()

    full_raw = pd.concat(frames, ignore_index=True)

    # Filter out invalid Col2 rows right away per GSQL WHERE clause:
    # WHERE Col2 IS NOT NULL AND Col2 != 'ವಿಭಾಗದ ಹೆಸರು' AND Col2 != 'ವಿಭಾಗ' AND Col2 != '2' AND Col2 != 'F' AND Col2 != 'S'
    col2_series = full_raw["Col2"].map(lambda v: str(v).strip() if v is not None and not pd.isna(v) else "")
    invalid_col2 = {"", "ವಿಭಾಗದ ಹೆಸರು", "ವಿಭಾಗ", "2", "F", "S", "NAN", "NONE"}
    valid_mask = ~col2_series.str.upper().isin(invalid_col2) & ~col2_series.isin(invalid_col2)

    full_raw = full_raw[valid_mask].reset_index(drop=True)
    col2_series = col2_series[valid_mask].reset_index(drop=True)

    def _clean_str_col(s: pd.Series) -> pd.Series:
        return s.map(
            lambda v: ""
            if v is None or (isinstance(v, float) and pd.isna(v))
            else re.sub(r"\s+", " ", str(v)).strip()
        )

    col3 = _clean_str_col(full_raw["Col3"])   # District
    col4 = _clean_str_col(full_raw["Col4"])   # Taluk
    col5 = _clean_str_col(full_raw["Col5"])   # Constituency
    col6 = _clean_str_col(full_raw["Col6"])   # Scheme / HOA
    col8 = _clean_str_col(full_raw["Col8"]).str.upper()  # S/F
    col28 = _clean_str_col(full_raw["Col28"]).str.upper()  # Status
    col29 = _clean_str_col(full_raw["Col29"])  # AE
    col30 = _clean_str_col(full_raw["Col30"])  # AEE

    col9 = full_raw["Col9"].map(to_safe_float)     # Estimated Amount
    col15 = full_raw["Col15"].map(to_safe_float)   # Tender Amount
    col17 = full_raw["Col17"].map(to_safe_float)   # Revised Amount
    col23 = full_raw["Col23"].map(to_safe_float)   # Exp up to Mar 26
    col24 = full_raw["Col24"].map(to_safe_float)   # Exp 26-27
    col25 = full_raw["Col25"].map(to_safe_float)   # Cum Exp
    col26 = full_raw["Col26"].map(to_safe_float)   # Bal Bill
    col27 = full_raw["Col27"].map(to_safe_float)   # Bal Amt

    col12_dates = [parse_single_date(v) for v in full_raw["Col12"].tolist()]
    col19_dates = [parse_single_date(v) for v in full_raw["Col19"].tolist()]
    col21_dates = [parse_single_date(v) for v in full_raw["Col21"].tolist()]
    col22_raw = _clean_str_col(full_raw["Col22"])
    col21_raw = _clean_str_col(full_raw["Col21"])

    col12_fy = [get_fin_year_label(d) for d in col12_dates]
    col19_fy = [get_fin_year_label(d) for d in col19_dates]

    # Exact GSQL is_overdue (> today) AND standard overdue (< today)
    is_od_gsql = []
    is_od_std = []
    for t_raw, a_raw, stat_val, t_dt in zip(
        col21_raw.tolist(), col22_raw.tolist(), col28.tolist(), col21_dates
    ):
        if a_raw == "" and t_raw != "" and stat_val == "P" and isinstance(t_dt, date):
            is_od_gsql.append(1 if t_dt > today_ist_date else 0)
            is_od_std.append(1 if t_dt < today_ist_date else 0)
        else:
            is_od_gsql.append(0)
            is_od_std.append(0)

    abs_df = pd.DataFrame(
        {
            "Col2": col2_series,
            "Col3": col3,
            "Col4": col4,
            "Col5": col5,
            "Col6": col6,
            "Col8": col8,
            "Col9": col9,
            "Col12_dt": col12_dates,
            "Col12_fy": col12_fy,
            "Col15": col15,
            "Col17": col17,
            "Col19_dt": col19_dates,
            "Col19_fy": col19_fy,
            "Col21_dt": col21_dates,
            "Col22_empty": col22_raw == "",
            "Col23": col23,
            "Col24": col24,
            "Col25": col25,
            "Col26": col26,
            "Col27": col27,
            "Col28": col28,
            "Col29": col29,
            "Col30": col30,
            "Col32_gsql": is_od_gsql,
            "Col32_std": is_od_std,
        }
    )

    # Parse 'Division_wise ABSTRACT' sheet for 2026-27 Grant Allocation (Col B, Col C, Col L)
    alloc_df = pd.DataFrame(columns=["abs_div", "abs_scheme", "abs_alloc"])
    abs_sheet_actual = available_sheets.get("division_wise abstract")
    if abs_sheet_actual and abs_sheet_actual in raw_sheets_map:
        raw_abs = raw_sheets_map[abs_sheet_actual]
        if len(raw_abs) > 1 and raw_abs.shape[1] >= 12:
            sub_abs = raw_abs.iloc[1:].copy()
            alloc_df = pd.DataFrame(
                {
                    "abs_div": sub_abs.iloc[:, 1].map(
                        lambda v: "" if v is None or pd.isna(v) else str(v).strip()
                    ),
                    "abs_scheme": sub_abs.iloc[:, 2].map(
                        lambda v: "" if v is None or pd.isna(v) else str(v).strip()
                    ),
                    "abs_alloc": sub_abs.iloc[:, 11].map(to_safe_float),
                }
            )
            alloc_df = alloc_df[alloc_df["abs_div"] != ""].reset_index(drop=True)

    return abs_df, alloc_df


def get_scheme_regex(v_scheme: str) -> str:
    """Exact replica of GetRegex LAMBDA in GSQL."""
    if not v_scheme or v_scheme == "-":
        return "NO_MATCH"
    if v_scheme.startswith("SCP"):
        return r"ವಿಶೇಷ ಘಟಕ ಯೋಜನೆ"
    if v_scheme.startswith("TSP"):
        return r"ಗಿರಿಜನ ಉಪಯೋಜನೆ"
    mapping = {
        "KMERC": r"ಕೆ\.ಎಂ\.ಇ\.ಆರ್",
        "PMKSY": r"PMKSY",
        "FMP": r"FPW",
        "KHARLAND": r"KHARLAND",
        "NT": r"\-NT\s*$",
        "MOT": r"\-\s*MOT\s*$",
        "NMOT": r"NMOT\s*$",
        "APB": r"\-APB\s*$",
        "NAPB": r"NAPB\s*$",
        "LIS": r"\-\s*LIS\s*$",
        "NLIS": r"NLIS\s*$",
        "PV": r"\-\s*PV\s*$",
    }
    return mapping.get(v_scheme, re.escape(v_scheme))


def sum_grant_allocation(alloc_df: pd.DataFrame, div_name: str, regex_pattern: str) -> float:
    if alloc_df.empty or regex_pattern == "NO_MATCH":
        return 0.0
    sub = alloc_df[alloc_df["abs_div"] == div_name]
    if sub.empty:
        return 0.0
    try:
        matched = sub["abs_scheme"].str.contains(
            regex_pattern, case=False, regex=True, na=False
        )
        return float(sub.loc[matched, "abs_alloc"].sum())
    except Exception:
        return 0.0


def format_alloc_val(val: float, v_scheme: str = "", is_total: bool = False):
    rounded = round(val, 2)
    num_str = f"{rounded:.2f}".rstrip("0").rstrip(".") if rounded != 0 else "0.00"
    if not is_total:
        if v_scheme.startswith("SCP"):
            return f"{rounded:.2f} (Combined SCP)"
        if v_scheme.startswith("TSP"):
            return f"{rounded:.2f} (Combined TSP)"
        return f"{rounded:.2f}"
    return f"{rounded:.2f}" if rounded != 0 else num_str


def compute_gsql_abstract_table(
    base_df: pd.DataFrame,
    alloc_df: pd.DataFrame,
    currently_showing: str,
    use_gsql_overdue_rule: bool = True,
) -> pd.DataFrame:
    """Executes the exact zoneDetails, zoneTotal, spacerRow, detailsData, totalsData, and SORT/SCAN logic."""
    if base_df.empty:
        return pd.DataFrame(columns=ABSTRACT_HEADERS_32 + ["_row_kind"])

    od_col = "Col32_gsql" if use_gsql_overdue_rule else "Col32_std"
    is_overdue_mode = currently_showing.strip().upper() == "OVERDUE WORKS"

    status_codes = [
        "CP", "CF", "P", "TBS", "TP", "ETS", "OTH",
        "CBW", "CEF", "PBW", "CBWPM", "RRR", "BWF",
    ]
    special_set = {"TBS", "OTH", "TP", "ETS"}

    def _agg_subset(df_true: pd.DataFrame, target_div_for_alloc: str, v_scheme: str, is_total_row: bool, unique_schemes_in_group: list[str]):
        true_cS = int((df_true["Col8"] == "S").sum())
        true_cF = int((df_true["Col8"] == "F").sum())
        true_tot = true_cS + true_cF
        true_od = int(df_true[od_col].sum())
        pct_overdue = f"{(true_od / true_tot) * 100:.2f}%" if true_tot > 0 else "0.00%"

        true_spec = int(df_true["Col28"].isin(special_set).sum())
        pct_special = f"{(true_spec / true_tot) * 100:.2f}%" if true_tot > 0 else "0.00%"

        df_h3 = df_true[df_true[od_col] == 1] if is_overdue_mode else df_true

        mask_s = df_h3["Col8"] == "S"
        mask_f = df_h3["Col8"] == "F"

        cS = int(mask_s.sum())
        sS = round(float(df_h3.loc[mask_s, "Col9"].sum()), 2)
        cF = int(mask_f.sum())
        sF = round(float(df_h3.loc[mask_f, "Col9"].sum()), 2)
        tot_works = cS + cF
        tot_est = round(sS + sF, 2)

        sTender = round(float(df_h3["Col15"].sum()), 2)
        sRev = round(float(df_h3["Col17"].sum()), 2)

        if not is_total_row:
            rgx = get_scheme_regex(v_scheme)
            sAlloc = sum_grant_allocation(alloc_df, target_div_for_alloc, rgx)
            sAlloc_disp = format_alloc_val(sAlloc, v_scheme=v_scheme, is_total=False)
        else:
            unique_rgx = list(dict.fromkeys(get_scheme_regex(s) for s in unique_schemes_in_group))
            sAlloc = sum(
                sum_grant_allocation(alloc_df, target_div_for_alloc, r)
                for r in unique_rgx
            )
            sAlloc_disp = format_alloc_val(sAlloc, is_total=True)

        sExpMar26 = round(float(df_h3["Col23"].sum()), 2)
        sExp26_27 = round(float(df_h3["Col24"].sum()), 2)
        sCumExp = round(float(df_h3["Col25"].sum()), 2)
        sBalBill = round(float(df_h3["Col26"].sum()), 2)
        sBalAmt = round(float(df_h3["Col27"].sum()), 2)

        st_counts = [int((df_h3["Col28"] == sc).sum()) for sc in status_codes]

        return (
            [
                cS, sS, cF, sF, tot_works, tot_est,
                sTender, sRev, sAlloc_disp,
                sExpMar26, sExp26_27, sCumExp, sBalBill, sBalAmt,
            ]
            + st_counts
            + [pct_overdue, pct_special]
        )

    unique_schemes = sorted([s for s in base_df["Col6"].unique().tolist() if s != ""])
    unique_divs = sorted([d for d in base_df["Col2"].unique().tolist() if d != ""])

    combined_rows = []

    # 1. zoneDetails
    for v_scheme in unique_schemes:
        sub_s = base_df[base_df["Col6"] == v_scheme]
        metrics = _agg_subset(sub_s, "ದಕ್ಷಿಣ ವಲಯ", v_scheme, False, [v_scheme])
        if metrics[4] > 0:  # tot_works > 0
            combined_rows.append(
                [1, "Zone", 1, v_scheme, "ದಕ್ಷಿಣ ವಲಯ", v_scheme] + metrics + ["detail"]
            )

    # 2. zoneTotal
    z_tot_metrics = _agg_subset(base_df, "ದಕ್ಷಿಣ ವಲಯ", "ZZZ", True, unique_schemes)
    if z_tot_metrics[4] > 0:
        combined_rows.append(
            [1, "Zone", 2, "ZZZ", "", "👉 ಸಣ್ಣ ನೀರಾವರಿ (ದ) ಒಟ್ಟು"]
            + z_tot_metrics
            + ["total"]
        )

    # 3. spacerRow
    combined_rows.append(
        [2, "Spacer", 2, "AAA", "ವಿಭಾಗವಾರು", ""] + ([""] * 29) + ["spacer"]
    )

    # 4. detailsData & 5. totalsData per Division
    for v_div in unique_divs:
        sub_div = base_df[base_df["Col2"] == v_div]
        div_schemes = sorted([s for s in sub_div["Col6"].unique().tolist() if s != ""])
        for v_scheme in div_schemes:
            sub_ds = sub_div[sub_div["Col6"] == v_scheme]
            metrics = _agg_subset(sub_ds, v_div, v_scheme, False, [v_scheme])
            if metrics[4] > 0:
                combined_rows.append(
                    [3, v_div, 1, v_scheme, v_div, v_scheme] + metrics + ["detail"]
                )

        d_tot_metrics = _agg_subset(sub_div, v_div, "ZZZ", True, div_schemes)
        if d_tot_metrics[4] > 0:
            combined_rows.append(
                [3, v_div, 2, "ZZZ", "", f"👉 {v_div} ಒಟ್ಟು"]
                + d_tot_metrics
                + ["total"]
            )

    # Sort by (col1, col2, col3, col4) just like SORT(filtered_combined, 1, TRUE, 2, TRUE, 3, TRUE, 4, TRUE)
    combined_rows.sort(key=lambda r: (r[0], str(r[1]), r[2], str(r[3])))

    # Generate Sl. No. using SCAN logic (resets to 1 per group for type==1 rows, blank for type==2)
    final_records = []
    acc = 0
    prev_grp = ""
    for r in combined_rows:
        cur_grp = r[1]
        cur_type = r[2]
        row_kind = r[-1]
        payload_31 = r[4:-1]

        if cur_type == 1:
            acc = 1 if cur_grp != prev_grp else acc + 1
            sl_str = str(acc)
        else:
            sl_str = ""
        prev_grp = cur_grp

        final_records.append([sl_str] + payload_31 + [row_kind])

    return pd.DataFrame(final_records, columns=ABSTRACT_HEADERS_32 + ["_row_kind"])


def render_abstract_32col_html(df_abs: pd.DataFrame) -> str:
    """Renders the 32-column abstract table matching screenshot 1 with sticky headers and row highlights."""
    hdr1_cells = "".join(f"<th>{html.escape(c)}</th>" for c in ABSTRACT_HEADERS_32)
    hdr2_cells = "".join(f"<th>{i}</th>" for i in range(1, 33))

    float_col_indices = {4, 6, 8, 9, 10, 12, 13, 14, 15, 16}  # 0-based in 32 cols

    body_rows = []
    for _, row in df_abs.iterrows():
        row_kind = row["_row_kind"]
        if row_kind == "spacer":
            body_rows.append(
                '<tr class="abs-spacer-row">'
                '<td colspan="32">ವಿಭಾಗವಾರು (Division-Wise Breakdown)</td>'
                "</tr>"
            )
            continue

        tr_cls = "abs-total-row" if row_kind == "total" else "abs-detail-row"
        tds = []
        for idx in range(32):
            val = row.iloc[idx]
            if val == "" or val is None:
                tds.append("<td></td>")
            elif idx in float_col_indices and isinstance(val, (int, float)):
                tds.append(f"<td>{val:.2f}</td>")
            else:
                tds.append(f"<td>{html.escape(str(val))}</td>")
        body_rows.append(f'<tr class="{tr_cls}">{"".join(tds)}</tr>')

    return (
        '<div class="abs-table-container">'
        '<table class="abs-matrix-table">'
        f'<thead><tr class="abs-hdr-row-1">{hdr1_cells}</tr>'
        f'<tr class="abs-hdr-row-2">{hdr2_cells}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        "</table></div>"
    )


def export_abstract_to_excel_bytes(df_abs: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Division_Wise_Abstract"

    fill_hdr1 = PatternFill("solid", fgColor="CFE2F3")
    fill_hdr2 = PatternFill("solid", fgColor="9FC5E8")
    fill_total = PatternFill("solid", fgColor="FFF2CC")
    fill_spacer = PatternFill("solid", fgColor="D9EAD3")

    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_reg = Font(name="Calibri", size=10, bold=False, color="000000")
    thin = Side(style="thin", color="888888")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c_idx, h_text in enumerate(ABSTRACT_HEADERS_32, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h_text)
        cell.fill = fill_hdr1
        cell.font = font_bold
        cell.alignment = align_c
        cell.border = border_all
    ws.row_dimensions[1].height = 36

    for c_idx in range(1, 33):
        cell = ws.cell(row=2, column=c_idx, value=c_idx)
        cell.fill = fill_hdr2
        cell.font = font_bold
        cell.alignment = align_c
        cell.border = border_all
    ws.row_dimensions[2].height = 19

    curr_r = 3
    for _, row in df_abs.iterrows():
        r_kind = row["_row_kind"]
        for c_idx in range(1, 33):
            val = row.iloc[c_idx - 1]
            cell = ws.cell(row=curr_r, column=c_idx, value=val)
            cell.border = border_all
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r_kind == "total":
                cell.fill = fill_total
                cell.font = font_bold
            elif r_kind == "spacer":
                cell.fill = fill_spacer
                cell.font = font_bold
            else:
                cell.font = font_reg
        curr_r += 1

    for c_idx in range(1, 33):
        ws.column_dimensions[get_column_letter(c_idx)].width = 14
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# WORKBOOK PARSER (READS ALL SHEETS FOR BOTH PAGE 1 & PAGE 2)
# -----------------------------------------------------------------------------
def parse_xlsx_workbook(workbook_bytes: bytes):
    progress = st.progress(0, text="Opening workbook...")

    try:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="calamine")
    except Exception:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="openpyxl")

    available_sheets = {
        normalize_text(name).lower(): name for name in excel_file.sheet_names
    }

    today_ist_date = datetime.now(ZoneInfo("Asia/Kolkata")).date()

    compiled_rows = []
    full_overdue_rows = []
    sheet_diagnostics = []
    raw_sheets_cache = {}

    total_targets = len(TARGET_SHEETS)
    for idx_sheet, target in enumerate(TARGET_SHEETS, start=1):
        progress.progress(
            idx_sheet / (total_targets + 1),
            text=f"Reading {idx_sheet}/{total_targets}: {target}...",
        )

        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Missing",
                    "Sub-division Col": "N/A",
                    "EE / TA / AS Cols": "N/A",
                    "Rows Loaded": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )
            continue

        try:
            raw_df = read_sheet_safely(excel_file, workbook_bytes, actual_sheet_name)
            raw_sheets_cache[actual_sheet_name] = raw_df
            if raw_df.empty:
                continue

            header_idx = None
            for idx in range(min(30, len(raw_df))):
                row_text = " | ".join(
                    normalize_text(v) for v in raw_df.iloc[idx].tolist()
                )
                if (
                    COL_SO_KN in row_text
                    or COL_AEE_KN in row_text
                    or "ಶಾಖಾಧಿಕಾರಿ" in row_text
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                continue

            raw_headers = [
                normalize_text(col) for col in raw_df.iloc[header_idx].tolist()
            ]
            col_idx_map = find_matching_column_indices(raw_headers)

            required_keys = {
                "division",
                "so",
                "aee",
                "target_date",
                "actual_date",
                "status",
            }
            optional_keys = {
                "subdivision",
                "ee",
                "ta",
                "as_off",
                "work_name",
                "start_date",
                "hoa",
            }

            if (
                (required_keys - set(col_idx_map.keys()))
                or (optional_keys - set(col_idx_map.keys()))
            ) and (header_idx + 1 < len(raw_df)):
                next_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx + 1].tolist()
                ]
                combined_headers = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(raw_headers, next_row)
                ]
                alt_map = find_matching_column_indices(combined_headers)
                for k, v in alt_map.items():
                    if k not in col_idx_map:
                        col_idx_map[k] = v

            if header_idx > 0:
                prev_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx - 1].tolist()
                ]
                combined_prev = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(prev_row, raw_headers)
                ]
                alt_prev_map = find_matching_column_indices(combined_prev)
                for k, v in alt_prev_map.items():
                    if k not in col_idx_map:
                        col_idx_map[k] = v

            if required_keys - set(col_idx_map.keys()):
                continue

            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            def _extract_col_series(key: str, default_val: str = "") -> pd.Series:
                if key in col_idx_map:
                    return pd.Series(
                        [
                            format_cell_for_display(v)
                            for v in df.iloc[:, col_idx_map[key]].tolist()
                        ],
                        index=df.index,
                        dtype=str,
                    )
                return pd.Series([default_val] * len(df), index=df.index, dtype=str)

            s_div_raw = _extract_col_series("division")
            s_subdiv_raw = _extract_col_series("subdivision")
            s_so_raw = _extract_col_series("so")
            s_aee_raw = _extract_col_series("aee")
            s_ee_raw = _extract_col_series("ee")
            s_ta_raw = _extract_col_series("ta")
            s_as_raw = _extract_col_series("as_off")
            s_work_name = _extract_col_series("work_name")
            s_start_date = _extract_col_series("start_date")
            s_hoa_raw = _extract_col_series("hoa", default_val=target)
            s_hoa = s_hoa_raw.map(lambda x: x if x != "" else target)

            s_status = _extract_col_series("status")
            s_status_upper = s_status.str.upper().str.strip()
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            target_date_list = parse_indian_date_list(s_target_raw)
            actual_empty = is_empty_date(s_actual_raw)

            duration_list = [
                calc_overdue_duration(td, today_ist_date) for td in target_date_list
            ]
            s_overdue_duration = pd.Series(duration_list, index=df.index, dtype=str)

            is_status_p = s_status_upper == "P"
            is_status_ets = s_status_upper == "ETS"

            target_is_before_today = pd.Series(
                [
                    (d is not None) and (d < today_ist_date)
                    for d in target_date_list
                ],
                index=df.index,
                dtype=bool,
            )

            is_overdue = is_status_p & target_is_before_today & actual_empty

            valid_row_mask = s_div_raw.map(is_valid_division_text)

            s_subdiv_mapped = s_subdiv_raw.map(map_optional_field)
            s_ee_mapped = s_ee_raw.map(map_optional_field)
            s_ta_mapped = s_ta_raw.map(map_optional_field)
            s_as_mapped = s_as_raw.map(map_optional_field)
            s_so_mapped = s_so_raw.map(map_officer_or_unassigned)
            s_aee_mapped = s_aee_raw.map(map_officer_or_unassigned)
            s_target_Formatted = pd.Series(
                [format_cell_for_display(v) for v in s_target_raw.tolist()],
                index=df.index,
                dtype=str,
            )

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "ಕಾಮಗಾರಿ ಹೆಸರು": s_work_name,
                    "division_name": s_div_raw,
                    "subdivision_name": s_subdiv_mapped,
                    "ee_name": s_ee_mapped,
                    "ta_name": s_ta_mapped,
                    "as_name": s_as_mapped,
                    "so_name": s_so_mapped,
                    "aee_name": s_aee_mapped,
                    "status_raw": s_status,
                    "target_date_raw": s_target_Formatted,
                    "actual_empty": actual_empty,
                    "is_status_p": is_status_p,
                    "is_status_ets": is_status_ets,
                    "is_overdue": is_overdue,
                }
            )
            sub = sub[valid_row_mask].reset_index(drop=True)
            compiled_rows.append(sub)

            overdue_mask = valid_row_mask & is_overdue
            if overdue_mask.any():
                overdue_curated = pd.DataFrame(
                    {
                        COL_WORK_NAME_KN: s_work_name.loc[overdue_mask].values,
                        COL_DIV_KN: s_div_raw.loc[overdue_mask].values,
                        COL_SUBDIV_KN: [
                            display_clean_cell(v)
                            for v in s_subdiv_mapped.loc[overdue_mask].values
                        ],
                        COL_SO_KN: [
                            display_clean_cell(v)
                            for v in s_so_mapped.loc[overdue_mask].values
                        ],
                        COL_AEE_KN: [
                            display_clean_cell(v)
                            for v in s_aee_mapped.loc[overdue_mask].values
                        ],
                        LABEL_TA_EN: [
                            display_clean_cell(v)
                            for v in s_ta_mapped.loc[overdue_mask].values
                        ],
                        LABEL_AS_EN: [
                            display_clean_cell(v)
                            for v in s_as_mapped.loc[overdue_mask].values
                        ],
                        LABEL_EE_EN: [
                            display_clean_cell(v)
                            for v in s_ee_mapped.loc[overdue_mask].values
                        ],
                        COL_START_DATE_KN: s_start_date.loc[overdue_mask].values,
                        COL_TARGET_DATE_KN: s_target_Formatted.loc[overdue_mask].values,
                        "Overdue By": s_overdue_duration.loc[overdue_mask].values,
                        COL_HOA_KN: s_hoa.loc[overdue_mask].values,
                        "_division_name": s_div_raw.loc[overdue_mask].values,
                        "_subdivision_name": s_subdiv_mapped.loc[overdue_mask].values,
                        "_ee_name": s_ee_mapped.loc[overdue_mask].values,
                        "_ta_name": s_ta_mapped.loc[overdue_mask].values,
                        "_as_name": s_as_mapped.loc[overdue_mask].values,
                        "_so_name": s_so_mapped.loc[overdue_mask].values,
                        "_aee_name": s_aee_mapped.loc[overdue_mask].values,
                    }
                )
                full_overdue_rows.append(overdue_curated.reset_index(drop=True))

            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "Sub-division Col": "Found" if "subdivision" in col_idx_map else "Missing",
                    "EE / TA / AS Cols": (
                        f"EE:{'Y' if 'ee' in col_idx_map else 'N'} | "
                        f"TA:{'Y' if 'ta' in col_idx_map else 'N'} | "
                        f"AS:{'Y' if 'as_off' in col_idx_map else 'N'}"
                    ),
                    "Rows Loaded": len(sub),
                    "Overdue ('P')": int(overdue_mask.sum()),
                    "ETS": int((valid_row_mask & is_status_ets).sum()),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "Sub-division Col": "Error",
                    "EE / TA / AS Cols": "Error",
                    "Rows Loaded": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )

    # Also read 'Division_wise ABSTRACT' for Page 2 Grant Allocations
    progress.progress(
        0.98, text="Building Page 2 Division-Wise Abstract dataset..."
    )
    abs_sheet_actual = available_sheets.get("division_wise abstract")
    if abs_sheet_actual and abs_sheet_actual not in raw_sheets_cache:
        try:
            raw_sheets_cache[abs_sheet_actual] = read_sheet_safely(
                excel_file, workbook_bytes, abs_sheet_actual
            )
        except Exception:
            pass

    abstract_raw_df, abstract_alloc_df = build_abstract_dataset_from_frames(
        raw_sheets_cache, available_sheets, today_ist_date
    )

    progress.empty()

    master_df = (
        pd.concat(compiled_rows, ignore_index=True)
        if compiled_rows
        else pd.DataFrame()
    )
    overdue_full_df = (
        pd.concat(full_overdue_rows, ignore_index=True).fillna("")
        if full_overdue_rows
        else pd.DataFrame()
    )
    diag_df = pd.DataFrame(sheet_diagnostics)
    return (
        master_df,
        overdue_full_df,
        diag_df,
        today_ist_date.strftime("%d-%m-%Y"),
        abstract_raw_df,
        abstract_alloc_df,
    )


def save_shared_state(state: dict, data_tuple, source_label: str):
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M:%S %p")
    state["version"] = SCHEMA_VERSION
    state["data"] = data_tuple
    state["last_synced"] = now_str
    state["source_label"] = source_label
    try:
        with open(DISK_BACKUP_FILE, "wb") as f:
            pickle.dump(
                {
                    "version": SCHEMA_VERSION,
                    "data": data_tuple,
                    "last_synced": now_str,
                    "source_label": source_label,
                },
                f,
            )
    except Exception:
        pass


def build_dropdown_options(df: pd.DataFrame, col_name: str) -> list[str]:
    raw_vals = [
        v
        for v in df[col_name].dropna().unique().tolist()
        if v != "" and v != UNSPECIFIED_LABEL
    ]
    options = [ALL_OPTION] + sorted(raw_vals)
    if (df[col_name] == UNSPECIFIED_LABEL).any():
        options.append(UNSPECIFIED_LABEL)
    return options


# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Dashboard Controls")

sheet_url_input = st.sidebar.text_input(
    "Google Spreadsheet Link",
    value=DEFAULT_SHEET_URL,
)

uploaded_xlsx = st.sidebar.file_uploader(
    "Or Upload Workbook (.xlsx) Directly",
    type=["xlsx"],
)

sync_clicked = st.sidebar.button(
    "🔄 Sync Latest Data from Google Sheets",
    use_container_width=True,
    type="primary",
)

denominator_mode = st.sidebar.radio(
    "Total Works Assigned Count Basis:",
    options=[
        "All Works Assigned (All Statuses)",
        "Only Ongoing Works (Status = 'Progress')",
    ],
    index=0,
)

hide_zero_overdue = st.sidebar.checkbox(
    "Show only officers with Overdue > 0",
    value=True,
)

show_unassigned_bar = st.sidebar.checkbox(
    "Include Empty / Unassigned Officer card at the bottom",
    value=True,
)

st.sidebar.divider()
abstract_od_rule = st.sidebar.radio(
    "Page 2 Abstract 'Overdue %' Rule:",
    options=[
        "Match GSQL Formula (Target > Today)",
        "Standard Overdue (Target < Today)",
    ],
    index=0,
)

# -----------------------------------------------------------------------------
# TOP HEADER WITH RIGHT-SIDE NEXT/BACK PAGE BUTTON
# -----------------------------------------------------------------------------
hdr_col_left, hdr_col_right = st.columns([0.73, 0.27], gap="small")

with hdr_col_left:
    if st.session_state["active_page"] == "overdue":
        st.markdown(
            """
            <div class="premium-compact-header">
                <div class="premium-icon-wrapper">📊</div>
                <div class="premium-compact-title">Minor Irrigation — Overdue Works Dashboard</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="premium-compact-header">
                <div class="premium-icon-wrapper">📑</div>
                <div class="premium-compact-title">Minor Irrigation — Division-Wise Abstract</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

with hdr_col_right:
    if st.session_state["active_page"] == "overdue":
        if st.button(
            "📑 Division Abstract ➔",
            type="primary",
            use_container_width=True,
            key="nav_to_abstract_btn",
        ):
            st.session_state["active_page"] = "abstract"
            st.rerun()
    else:
        if st.button(
            "⬅️ Overdue Dashboard",
            type="primary",
            use_container_width=True,
            key="nav_to_overdue_btn",
        ):
            st.session_state["active_page"] = "overdue"
            st.rerun()

# -----------------------------------------------------------------------------
# DATA LOAD & SHARED CACHE
# -----------------------------------------------------------------------------
shared_state = get_shared_state_v18_mob()

needs_fresh_load = (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or len(shared_state["data"]) != 6
    or sync_clicked
)

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
        with st.spinner("Loading worksheets & Division Abstract from Google Sheets..."):
            try:
                xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                parsed = parse_xlsx_workbook(xlsx_bytes)
                save_shared_state(
                    shared_state, parsed, "Google Sheets Live Snapshot"
                )
            except Exception as e:
                if shared_state["data"] is not None and len(shared_state["data"]) == 6:
                    st.warning(
                        f"Google Sheets is busy ({e}). Showing snapshot from {shared_state['last_synced']}."
                    )
                else:
                    st.error(f"Unable to load data right now. ({e})")
                    st.stop()

(
    master_df,
    overdue_full_df,
    diag_df,
    today_str,
    abstract_raw_df,
    abstract_alloc_df,
) = shared_state["data"]

st.markdown(
    f"""
    <div class="mi-status-bar">
        <span class="mi-status-pill">⚡ {shared_state['source_label']}</span>
        <span>🕒 <b>Synced:</b> {shared_state['last_synced']} IST</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# PAGE 2 VIEW: DIVISION-WISE ABSTRACT (32-COLUMN MATRIX + 15 GSQL FILTERS)
# =============================================================================
if st.session_state["active_page"] == "abstract":
    if abstract_raw_df.empty:
        st.warning("No Division-Wise Abstract rows found in the workbook.")
        st.stop()

    st.markdown(
        '<div class="abs-filter-header">'
        "<span>🎛️ DIVISION-WISE ABSTRACT FILTERS (GSQL LIVE ENGINE)</span>"
        "<span>Select filters below to dynamically recalculate Zone & Division totals</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    def _opts(series: pd.Series) -> list[str]:
        vals = sorted([v for v in series.dropna().unique().tolist() if str(v).strip() != ""])
        return [ALL_OPTION] + vals

    with st.container(border=True):
        # Row 1 of Filters: A3 (S/F), B3 (PROGRESS STATUS), C3 (ಲೆಕ್ಕಶೀರ್ಷಿಕೆ), D3 (ವಲಯ / ವಿಭಾಗ), E3 (ಜಿಲ್ಲೆ)
        f_r1 = st.columns(5, gap="small")
        with f_r1[0]:
            f_a3_sf = st.selectbox("S/F", options=_opts(abstract_raw_df["Col8"]), index=0)
        with f_r1[1]:
            f_b3_status = st.selectbox("PROGRESS STATUS", options=_opts(abstract_raw_df["Col28"]), index=0)
        with f_r1[2]:
            f_c3_scheme = st.selectbox("ಲೆಕ್ಕಶೀರ್ಷಿಕೆ", options=_opts(abstract_raw_df["Col6"]), index=0)
        with f_r1[3]:
            f_d3_div = st.selectbox("ವಲಯ / ವಿಭಾಗ", options=_opts(abstract_raw_df["Col2"]), index=0)
        with f_r1[4]:
            f_e3_dist = st.selectbox("ಜಿಲ್ಲೆ", options=_opts(abstract_raw_df["Col3"]), index=0)

        # Row 2 of Filters: F3 (ತಾಲ್ಲೂಕು), G3 (ಮತಕ್ಷೇತ್ರ), H3 (CURRENTLY SHOWING), I3 (ಮೂಲ ಅಂದಾಜು ಮೊತ್ತ), J3 (Overdue)
        f_r2 = st.columns(5, gap="small")
        with f_r2[0]:
            f_f3_taluk = st.selectbox("ತಾಲ್ಲೂಕು", options=_opts(abstract_raw_df["Col4"]), index=0)
        with f_r2[1]:
            f_g3_const = st.selectbox("ಮತಕ್ಷೇತ್ರ", options=_opts(abstract_raw_df["Col5"]), index=0)
        with f_r2[2]:
            f_h3_showing = st.selectbox(
                "CURRENTLY SHOWING",
                options=["All Works", "Overdue Works"],
                index=0,
            )
        with f_r2[3]:
            f_i3_amt = st.selectbox(
                "ಮೂಲ ಅಂದಾಜು ಮೊತ್ತ (ರೂ.ಲಕ್ಷಗಳಲ್ಲಿ)",
                options=[
                    ALL_OPTION,
                    "Less than 100",
                    "0 - 200",
                    "100 - 500",
                    "200 - 500",
                    "500 - 1000",
                    "1000 - 2500",
                    "2500 - 5000",
                    "5000 - 10000",
                    "More than 10000",
                ],
                index=0,
            )
        with f_r2[4]:
            f_j3_od_age = st.selectbox(
                "Overdue",
                options=[
                    ALL_OPTION,
                    "<2 years",
                    "2-5 years",
                    "5-10 years",
                    "More than 10 years",
                ],
                index=0,
            )

        # Row 3 of Filters: K3 (ಸರ್ಕಾರದ ಆದೇಶ ದಿನಾಂಕ), L3 (DATA PULL/PAUSE), M3 (A.E.), N3 (A.E.E.), O3 (ಕಾರ್ಯಾದೇಶ ನೀಡಿದ ದಿನಾಂಕ)
        f_r3 = st.columns(5, gap="small")
        with f_r3[0]:
            f_k3_go_fy = st.selectbox(
                "ಸರ್ಕಾರದ ಆದೇಶ ದಿನಾಂಕ",
                options=_opts(abstract_raw_df["Col12_fy"]),
                index=0,
            )
        with f_r3[1]:
            f_l3_pull = st.checkbox("DATA PULL/PAUSE (Active)", value=True)
        with f_r3[2]:
            f_m3_ae = st.selectbox("A.E.", options=_opts(abstract_raw_df["Col29"]), index=0)
        with f_r3[3]:
            f_n3_aee = st.selectbox("A.E.E.", options=_opts(abstract_raw_df["Col30"]), index=0)
        with f_r3[4]:
            f_o3_wo_fy = st.selectbox(
                "ಕಾರ್ಯಾದೇಶ ನೀಡಿದ ದಿನಾಂಕ",
                options=_opts(abstract_raw_df["Col19_fy"]),
                index=0,
            )

    if not f_l3_pull:
        st.error("🛑 DATA PAUSED. Please check 'DATA PULL/PAUSE (Active)' to display the Division-Wise Abstract.")
        disable_mobile_select_keyboard()
        st.stop()

    # Apply the exact GSQL filterStr rules
    f_df = abstract_raw_df.copy()

    if f_a3_sf != ALL_OPTION:
        f_df = f_df[f_df["Col8"] == f_a3_sf]
    if f_b3_status != ALL_OPTION:
        f_df = f_df[f_df["Col28"] == f_b3_status]
    if f_c3_scheme != ALL_OPTION:
        f_df = f_df[f_df["Col6"] == f_c3_scheme]
    if f_d3_div != ALL_OPTION:
        f_df = f_df[f_df["Col2"] == f_d3_div]
    if f_e3_dist != ALL_OPTION:
        f_df = f_df[f_df["Col3"] == f_e3_dist]
    if f_f3_taluk != ALL_OPTION:
        f_df = f_df[f_df["Col4"] == f_f3_taluk]
    if f_g3_const != ALL_OPTION:
        f_df = f_df[f_df["Col5"] == f_g3_const]

    if f_i3_amt != ALL_OPTION:
        if f_i3_amt == "More than 10000":
            f_df = f_df[f_df["Col9"] > 10000]
        elif f_i3_amt == "5000 - 10000":
            f_df = f_df[(f_df["Col9"] >= 5000) & (f_df["Col9"] <= 10000)]
        elif f_i3_amt == "2500 - 5000":
            f_df = f_df[(f_df["Col9"] >= 2500) & (f_df["Col9"] <= 5000)]
        elif f_i3_amt == "1000 - 2500":
            f_df = f_df[(f_df["Col9"] >= 1000) & (f_df["Col9"] <= 2500)]
        elif f_i3_amt == "500 - 1000":
            f_df = f_df[(f_df["Col9"] >= 500) & (f_df["Col9"] <= 1000)]
        elif f_i3_amt == "200 - 500":
            f_df = f_df[(f_df["Col9"] >= 200) & (f_df["Col9"] <= 500)]
        elif f_i3_amt == "0 - 200":
            f_df = f_df[(f_df["Col9"] >= 0) & (f_df["Col9"] <= 200)]
        elif f_i3_amt == "100 - 500":
            f_df = f_df[(f_df["Col9"] >= 100) & (f_df["Col9"] <= 500)]
        elif f_i3_amt == "Less than 100":
            f_df = f_df[f_df["Col9"] < 100]

    if f_j3_od_age != ALL_OPTION:
        today_d = datetime.now(ZoneInfo("Asia/Kolkata")).date()
        d_24 = shift_months(today_d, -24)
        d_60 = shift_months(today_d, -60)
        d_120 = shift_months(today_d, -120)

        f_df = f_df[f_df["Col21_dt"].notna() & f_df["Col22_empty"]]
        if f_j3_od_age == "<2 years":
            f_df = f_df[(f_df["Col21_dt"] > d_24) & (f_df["Col21_dt"] <= today_d)]
        elif f_j3_od_age == "2-5 years":
            f_df = f_df[(f_df["Col21_dt"] <= d_24) & (f_df["Col21_dt"] > d_60)]
        elif f_j3_od_age == "5-10 years":
            f_df = f_df[(f_df["Col21_dt"] <= d_60) & (f_df["Col21_dt"] > d_120)]
        elif f_j3_od_age == "More than 10 years":
            f_df = f_df[f_df["Col21_dt"] <= d_120]

    if f_k3_go_fy != ALL_OPTION:
        f_df = f_df[f_df["Col12_fy"] == f_k3_go_fy]
    if f_o3_wo_fy != ALL_OPTION:
        f_df = f_df[f_df["Col19_fy"] == f_o3_wo_fy]
    if f_m3_ae != ALL_OPTION:
        f_df = f_df[f_df["Col29"] == f_m3_ae]
    if f_n3_aee != ALL_OPTION:
        f_df = f_df[f_df["Col30"] == f_n3_aee]

    if f_df.empty:
        st.warning("No works match the selected abstract filters.")
    else:
        abs_matrix_df = compute_gsql_abstract_table(
            f_df,
            abstract_alloc_df,
            currently_showing=f_h3_showing,
            use_gsql_overdue_rule=abstract_od_rule.startswith("Match GSQL"),
        )

        # Quick Download Bar & Row Count
        dl_c1, dl_c2, dl_c3 = st.columns([0.5, 0.25, 0.25], gap="small")
        with dl_c1:
            st.caption(
                f"Showing **{len(abs_matrix_df)}** abstract rows (`ದಕ್ಷಿಣ ವಲಯ` Zone Summary + `ವಿಭಾಗವಾರು` Division Breakdown) | Showing Mode: **{f_h3_showing}**"
            )
        with dl_c2:
            excel_abs_bytes = export_abstract_to_excel_bytes(abs_matrix_df)
            st.download_button(
                "📥 Download Abstract (.xlsx)",
                data=excel_abs_bytes,
                file_name=f"MI_Division_Wise_Abstract_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_c3:
            csv_abs_bytes = (
                abs_matrix_df.drop(columns=["_row_kind"])
                .to_csv(index=False)
                .encode("utf-8-sig")
            )
            st.download_button(
                "📥 Download CSV (.csv)",
                data=csv_abs_bytes,
                file_name=f"MI_Division_Wise_Abstract_{today_str}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.markdown(render_abstract_32col_html(abs_matrix_df), unsafe_allow_html=True)

    disable_mobile_select_keyboard()
    st.stop()


# =============================================================================
# PAGE 1 VIEW: OVERDUE WORKS DASHBOARD (EXISTING MOBILE/CE PORTRAIT VIEW)
# =============================================================================
if master_df.empty:
    st.warning("No matching data rows found.")
    st.stop()

unique_divisions = sorted(
    [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
)
division_options = [ALL_OPTION] + unique_divisions

with st.container(border=True):
    d_col1, d_col2 = st.columns(2, gap="small")
    with d_col1:
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
    with d_col2:
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

    o_col1, o_col2 = st.columns(2, gap="small")
    ta_options = build_dropdown_options(df_step3, "ta_name")
    with o_col1:
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
    with o_col2:
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

with st.container(border=True):
    view_choice = st.radio(
        "👤 Select Officer View:",
        options=[LABEL_AE_EN, LABEL_AEE_EN],
        horizontal=True,
    )

    sort_choice = st.radio(
        "⬇️ Sort Officers By:",
        options=[
            "Highest Overdue Count at Top",
            "Highest Overdue Percentage (%) at Top",
        ],
        horizontal=True,
    )

group_col = "so_name" if view_choice == LABEL_AE_EN else "aee_name"
full_filter_col = "_so_name" if view_choice == LABEL_AE_EN else "_aee_name"

# -----------------------------------------------------------------------------
# APPLY FILTERS TO BUILD SUMMARY DATAFRAME
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# 5-CARD COLOR-CODED MOBILE KPI GRID
# -----------------------------------------------------------------------------
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

st.markdown(
    f"""
    <div class="kpi-mobile-top-row">
        <div class="kpi-m-card kpi-blue">
            <div class="kpi-m-label">📅 Date (Today)</div>
            <div class="kpi-m-val">{today_str}</div>
        </div>
        <div class="kpi-m-card kpi-violet">
            <div class="kpi-m-label">👥 {html.escape(view_choice)}s</div>
            <div class="kpi-m-val">{named_officers_count:,}</div>
        </div>
        <div class="kpi-m-card kpi-cyan">
            <div class="kpi-m-label">📋 Total Works</div>
            <div class="kpi-m-val">{total_assigned_all:,}</div>
        </div>
    </div>
    <div class="kpi-mobile-bottom-row">
        <div class="kpi-m-card kpi-red">
            <div class="kpi-m-label">🔴 Overdue Works ('Progress')</div>
            <div class="kpi-m-val">{total_overdue_all:,} ({overall_overdue_pct}%)</div>
        </div>
        <div class="kpi-m-card kpi-green">
            <div class="kpi-m-label">🟢 'ETS' Works</div>
            <div class="kpi-m-val">{total_ets_all:,}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# SECTION 1: OFFICER OVERDUE SCORECARDS
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📈 Overdue & ETS Works by {html.escape(view_choice)} ({html.escape(sort_choice)})</div>
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

if chart_source_df.empty:
    st.info(
        "No officers match the current filter (try unchecking 'Show only officers with Overdue > 0' in the sidebar)."
    )
else:
    st.caption(f"Showing **{len(chart_source_df)}** {view_choice}(s):")
    st.markdown(
        render_mobile_officer_scorecards(chart_source_df, view_choice),
        unsafe_allow_html=True,
    )

st.divider()

# =============================================================================
# SECTION 2: COMBINED OVERDUE WORKS BY SELECTED AE / AEE
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
    f"Tap to select {view_choice} to inspect their Overdue Works:",
    options=officer_options,
    index=0 if officer_options else None,
    format_func=lambda name: (
        f"{name} ({officer_stats_map[name][0]} Overdue | "
        f"{officer_stats_map[name][1]} ETS / {officer_stats_map[name][2]} Total — "
        f"{officer_stats_map[name][3]:.1f}%)"
    ),
)

if selected_officer:
    if overdue_full_df.empty:
        st.info("No overdue works found across the 23 worksheets.")
    else:
        officer_overdue_df = overdue_full_df[
            overdue_full_df[full_filter_col] == selected_officer
        ].copy()

        if selected_division != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_division_name"] == selected_division
            ].copy()

        if selected_subdivision != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_subdivision_name"] == selected_subdivision
            ].copy()

        if selected_ee != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_ee_name"] == selected_ee
            ].copy()

        if selected_ta != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_ta_name"] == selected_ta
            ].copy()

        if selected_as != ALL_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_as_name"] == selected_as
            ].copy()

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

        rename_map = {
            COL_WORK_NAME_KN: "Work Name",
            COL_DIV_KN: LABEL_DIV_EN,
            COL_SUBDIV_KN: LABEL_SUBDIV_EN,
            opposite_officer_col: opposite_officer_en,
            COL_START_DATE_KN: "Start Date",
            COL_TARGET_DATE_KN: "Target Date",
            COL_HOA_KN: "H.O.A.",
        }
        officer_overdue_df.rename(columns=rename_map, inplace=True)
        officer_overdue_df.insert(0, "Sl. No.", range(1, len(officer_overdue_df) + 1))

        if officer_overdue_df.empty:
            st.success(
                f"🎉 **{selected_officer}** has **0** overdue works in **{filter_scope_label}**!"
            )
        else:
            st.markdown(
                render_mobile_work_dossier_cards(
                    officer_overdue_df, opposite_officer_en
                ),
                unsafe_allow_html=True,
            )

with st.expander("🔍 Inspect Worksheet Loading Status & Quick Sync (23 Target Sheets)"):
    if st.button(
        "🔄 Sync Latest Data Now",
        use_container_width=True,
        type="primary",
        key="mobile_bottom_sync",
    ):
        with shared_state["lock"]:
            with st.spinner("Syncing latest changes from Google Sheets..."):
                xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                parsed = parse_xlsx_workbook(xlsx_bytes)
                save_shared_state(
                    shared_state, parsed, "Google Sheets Live Snapshot"
                )
                st.rerun()
    st.dataframe(diag_df, use_container_width=True)


# =============================================================================
# BOTTOM SECTION: PENDING DATA ENTRY LIST GENERATOR (WITH MAIL ID & TIMESTAMP)
# =============================================================================
def build_pending_entry_summary(df_all: pd.DataFrame) -> pd.DataFrame:
    divisions = sorted(
        [d for d in df_all["division_name"].dropna().unique().tolist() if d != ""]
    )
    rows = []
    for div in divisions:
        sub_df = df_all[df_all["division_name"] == div]
        rows.append(
            {
                "Division": div,
                "Sub-division:": int((sub_df["subdivision_name"] == UNSPECIFIED_LABEL).sum()),
                "AE": int((sub_df["so_name"] == UNASSIGNED_LABEL).sum()),
                "AEE": int((sub_df["aee_name"] == UNASSIGNED_LABEL).sum()),
                "EE": int((sub_df["ee_name"] == UNSPECIFIED_LABEL).sum()),
                "TA": int((sub_df["ta_name"] == UNSPECIFIED_LABEL).sum()),
                "AS": int((sub_df["as_name"] == UNSPECIFIED_LABEL).sum()),
            }
        )
    return pd.DataFrame(rows)


def render_pending_entry_html(
    pending_df: pd.DataFrame, gen_date: str, gen_time: str, email_id: str
) -> str:
    cols = ["Sub-division:", "AE", "AEE", "EE", "TA", "AS"]

    body_rows_html = []
    for _, row in pending_df.iterrows():
        div_name = html.escape(str(row["Division"]))
        counts = [int(row[c]) for c in cols]

        if sum(counts) == 0:
            body_rows_html.append(
                f"<tr>"
                f'<td class="pe-div-cell">{div_name}</td>'
                f'<td class="pe-all-done" colspan="6">ALL COMPLETED 👏</td>'
                f"</tr>"
            )
        else:
            cells = [f'<td class="pe-div-cell">{div_name}</td>']
            for c_val in counts:
                if c_val == 0:
                    cells.append('<td class="pe-done-cell">Completed</td>')
                else:
                    cells.append(f'<td class="pe-num-cell">{c_val:,}</td>')
            body_rows_html.append(f"<tr>{''.join(cells)}</tr>")

    esc_email = html.escape(email_id)
    return f"""
    <style>
        .pe-card-outer {{
            background: #FFFFFF;
            padding: 14px 12px 10px 12px;
            border-radius: 8px;
            border: 1px solid #CBD5E1;
            overflow-x: auto;
            margin-top: 8px;
            margin-bottom: 10px;
            color: #111827;
            font-family: 'Inter', Arial, sans-serif;
        }}
        table.pe-table {{
            width: 100%;
            border-collapse: collapse;
            border: 2px solid #374151;
            font-size: 0.84rem;
            background: #FFFFFF;
        }}
        table.pe-table th, table.pe-table td {{
            border: 1px solid #9CA3AF;
            padding: 6px 10px;
            text-align: center;
            vertical-align: middle;
            white-space: nowrap;
        }}
        .pe-title-row {{
            background-color: #F7DC9F !important;
            color: #000000 !important;
            font-weight: 800 !important;
            font-size: 0.92rem !important;
        }}
        .pe-header-row th {{
            background-color: #D5A6BD !important;
            color: #000000 !important;
            font-weight: 700 !important;
        }}
        .pe-div-cell {{
            background-color: #EAD1DC !important;
            color: #000000 !important;
            font-weight: 700 !important;
        }}
        .pe-all-done {{
            background-color: #B7E1CD !important;
            color: #1E6B3A !important;
            font-weight: 800 !important;
        }}
        .pe-done-cell {{
            background-color: #B7E1CD !important;
            color: #1E6B3A !important;
            font-weight: 600 !important;
        }}
        .pe-num-cell {{
            background-color: #FFFFFF !important;
            color: #111827 !important;
            font-weight: 600 !important;
        }}
        .pe-footer-site {{
            text-align: center;
            font-style: italic;
            color: #1F2937;
            font-size: 0.82rem;
            margin-top: 8px;
            font-weight: 500;
        }}
        .pe-footer-meta {{
            text-align: center;
            color: #374151;
            font-size: 0.74rem;
            margin-top: 3px;
            font-weight: 600;
        }}
    </style>

    <div class="pe-card-outer" id="pe-capture-area">
        <table class="pe-table">
            <thead>
                <tr>
                    <th class="pe-title-row" colspan="7">PENDING DATA ENTRY LIST (Count of works)</th>
                </tr>
                <tr class="pe-header-row">
                    <th></th>
                    <th>Sub-division:</th>
                    <th>AE</th>
                    <th>AEE</th>
                    <th>EE</th>
                    <th>TA</th>
                    <th>AS</th>
                </tr>
            </thead>
            <tbody>
                {''.join(body_rows_html)}
            </tbody>
        </table>
        <div class="pe-footer-site">Analysis by ae-aee-overdue.streamlit.website</div>
        <div class="pe-footer-meta">
            Generated Date: {gen_date} &nbsp;|&nbsp; Time: {gen_time} IST &nbsp;|&nbsp; Mail ID: {esc_email}
        </div>
    </div>
    """


def generate_pending_entry_excel(
    pending_df: pd.DataFrame, gen_date: str, gen_time: str, email_id: str
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pending_Data_Entry"

    fill_title = PatternFill("solid", fgColor="F7DC9F")
    fill_header = PatternFill("solid", fgColor="D5A6BD")
    fill_div = PatternFill("solid", fgColor="EAD1DC")
    fill_done = PatternFill("solid", fgColor="B7E1CD")
    fill_white = PatternFill("solid", fgColor="FFFFFF")

    font_title = Font(name="Calibri", size=12, bold=True, color="000000")
    font_header = Font(name="Calibri", size=11, bold=True, color="000000")
    font_div = Font(name="Calibri", size=11, bold=True, color="000000")
    font_all_done = Font(name="Calibri", size=11, bold=True, color="1E6B3A")
    font_done = Font(name="Calibri", size=11, bold=False, color="1E6B3A")
    font_num = Font(name="Calibri", size=11, bold=False, color="000000")
    font_footer = Font(name="Calibri", size=10, italic=True, color="1F2937")
    font_meta = Font(name="Calibri", size=9, bold=True, color="374151")

    thin_side = Side(style="thin", color="888888")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:G1")
    ws["A1"] = "PENDING DATA ENTRY LIST (Count of works)"
    for col_idx in range(1, 8):
        c = ws.cell(row=1, column=col_idx)
        c.fill = fill_title
        c.font = font_title
        c.alignment = align_center
        c.border = border_cell
    ws.row_dimensions[1].height = 24

    headers = ["", "Sub-division:", "AE", "AEE", "EE", "TA", "AS"]
    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx, value=h_text)
        c.fill = fill_header
        c.font = font_header
        c.alignment = align_center
        c.border = border_cell
    ws.row_dimensions[2].height = 21

    cols = ["Sub-division:", "AE", "AEE", "EE", "TA", "AS"]
    curr_row = 3
    for _, row in pending_df.iterrows():
        ws.row_dimensions[curr_row].height = 20
        div_cell = ws.cell(row=curr_row, column=1, value=str(row["Division"]))
        div_cell.fill = fill_div
        div_cell.font = font_div
        div_cell.alignment = align_center
        div_cell.border = border_cell

        counts = [int(row[c]) for c in cols]
        if sum(counts) == 0:
            ws.merge_cells(
                start_row=curr_row, start_column=2, end_row=curr_row, end_column=7
            )
            for col_idx in range(2, 8):
                c = ws.cell(row=curr_row, column=col_idx)
                if col_idx == 2:
                    c.value = "ALL COMPLETED 👏"
                c.fill = fill_done
                c.font = font_all_done
                c.alignment = align_center
                c.border = border_cell
        else:
            for offset, c_val in enumerate(counts, start=2):
                c = ws.cell(row=curr_row, column=offset)
                if c_val == 0:
                    c.value = "Completed"
                    c.fill = fill_done
                    c.font = font_done
                else:
                    c.value = c_val
                    c.number_format = "#,##0"
                    c.fill = fill_white
                    c.font = font_num
                c.alignment = align_center
                c.border = border_cell
        curr_row += 1

    footer_row_1 = curr_row + 1
    ws.merge_cells(
        start_row=footer_row_1, start_column=1, end_row=footer_row_1, end_column=7
    )
    f1 = ws.cell(
        row=footer_row_1, column=1, value="Analysis by ae-aee-overdue.streamlit.website"
    )
    f1.font = font_footer
    f1.alignment = align_center

    footer_row_2 = curr_row + 2
    ws.merge_cells(
        start_row=footer_row_2, start_column=1, end_row=footer_row_2, end_column=7
    )
    f2 = ws.cell(
        row=footer_row_2,
        column=1,
        value=f"Generated Date: {gen_date} | Time: {gen_time} IST | Mail ID: {email_id}",
    )
    f2.font = font_meta
    f2.alignment = align_center

    col_widths = [20, 15, 14, 14, 14, 14, 14]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


st.divider()
st.markdown(
    """
    <div class="exec-section-banner">
        <div class="exec-section-title">📊 Pending Data Entry List (Count of Works)</div>
        <div class="exec-scope-badge">Enter your Mail ID below to generate & download the official status sheet</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.container(border=True):
    user_email_input = st.text_input(
        "📧 Enter Your Mail ID (Required to Generate & Download):",
        value=st.session_state.get("_pe_email", ""),
        placeholder="e.g. name@karnataka.gov.in / example@gmail.com",
    )

    gen_btn_clicked = st.button(
        "⚡ Generate Pending Data Entry List",
        type="primary",
        use_container_width=True,
    )

    if gen_btn_clicked:
        cleaned_mail = user_email_input.strip()
        if not cleaned_mail or "@" not in cleaned_mail:
            st.warning("⚠️ Please enter a valid Mail ID (containing '@') to generate the report.")
        else:
            now_ist = datetime.now(ZoneInfo("Asia/Kolkata"))
            st.session_state["_pe_email"] = cleaned_mail
            st.session_state["_pe_date"] = now_ist.strftime("%d-%m-%Y")
            st.session_state["_pe_time"] = now_ist.strftime("%I:%M:%S %p")
            st.session_state["_pe_generated"] = True

if st.session_state.get("_pe_generated"):
    saved_mail = st.session_state["_pe_email"]
    saved_date = st.session_state["_pe_date"]
    saved_time = st.session_state["_pe_time"]

    pending_summary_df = build_pending_entry_summary(master_df)
    table_html_str = render_pending_entry_html(
        pending_summary_df, saved_date, saved_time, saved_mail
    )
    excel_file_bytes = generate_pending_entry_excel(
        pending_summary_df, saved_date, saved_time, saved_mail
    )

    st.markdown(table_html_str, unsafe_allow_html=True)

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=excel_file_bytes,
            file_name=f"Pending_Data_Entry_List_{saved_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl_col2:
        components.html(
            f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <button id="pe-png-btn" style="
                width: 100%;
                height: 38px;
                background: #1E293B;
                color: #F8FAFC;
                border: 1px solid rgba(148, 163, 184, 0.35);
                border-radius: 8px;
                font-family: 'Inter', sans-serif;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
            ">📸 Download Image (.png)</button>
            <script>
            document.getElementById("pe-png-btn").addEventListener("click", function() {{
                const target = window.parent.document.getElementById("pe-capture-area");
                if (!target) return;
                html2canvas(target, {{ scale: 2, backgroundColor: "#FFFFFF" }}).then(canvas => {{
                    const link = document.createElement("a");
                    link.download = "Pending_Data_Entry_List_{saved_date}.png";
                    link.href = canvas.toDataURL("image/png");
                    link.click();
                }});
            }});
            </script>
            """,
            height=45,
        )

# =============================================================================
# DEBUG: FIND UNASSIGNED / MISSING DATA WORKS (ACROSS ALL 6 COLUMNS)
# =============================================================================
st.error("🔍 DEBUG: Find Unassigned / Missing Data Works (Across All Columns)")

empty_markers = {UNASSIGNED_LABEL, UNSPECIFIED_LABEL, "", "-"}

missing_any_mask = (
    filtered_master_df["subdivision_name"].isin(empty_markers)
    | filtered_master_df["so_name"].isin(empty_markers)
    | filtered_master_df["aee_name"].isin(empty_markers)
    | filtered_master_df["ee_name"].isin(empty_markers)
    | filtered_master_df["ta_name"].isin(empty_markers)
    | filtered_master_df["as_name"].isin(empty_markers)
)

debug_df = filtered_master_df[missing_any_mask].copy()

if debug_df.empty:
    st.success("🎉 No unassigned or missing fields found with current filters!")
else:
    def _list_missing_cols(row):
        missing = []
        if row["subdivision_name"] in empty_markers:
            missing.append("Sub-division")
        if row["so_name"] in empty_markers:
            missing.append("AE")
        if row["aee_name"] in empty_markers:
            missing.append("AEE")
        if row["ee_name"] in empty_markers:
            missing.append("EE")
        if row["ta_name"] in empty_markers:
            missing.append("TA")
        if row["as_name"] in empty_markers:
            missing.append("AS")
        return ", ".join(missing)

    debug_df["⚠️ Missing Column(s)"] = debug_df.apply(_list_missing_cols, axis=1)

    desired_debug_cols = [
        "sheet_name",
        "⚠️ Missing Column(s)",
        "ಕಾಮಗಾರಿ ಹೆಸರು",
        "division_name",
        "subdivision_name",
        "so_name",
        "aee_name",
        "ee_name",
        "ta_name",
        "as_name",
        "status_raw",
    ]
    available_debug_cols = [c for c in desired_debug_cols if c in debug_df.columns]

    st.caption(
        f"Showing **{len(debug_df):,}** work(s) where at least one column (Sub-division, AE, AEE, EE, TA, or AS) is missing:"
    )
    st.dataframe(
        debug_df[available_debug_cols],
        use_container_width=True,
    )

# Lock selectbox inputs so tapping a dropdown on mobile never opens the keyboard
disable_mobile_select_keyboard()
