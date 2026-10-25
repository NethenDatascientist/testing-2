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
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Minor Irrigation — Overdue Works Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@600;700&display=swap');

        /* Global App Ambient Background & Typography */
        .stApp {
            background:
                radial-gradient(circle at 12% 8%, rgba(37, 99, 235, 0.10), transparent 32%),
                radial-gradient(circle at 88% 14%, rgba(239, 68, 68, 0.07), transparent 30%),
                radial-gradient(circle at 50% 95%, rgba(16, 185, 129, 0.05), transparent 35%),
                #0B0F19 !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        .block-container {
            max-width: 100% !important;
            padding-left: 1.2rem !important;
            padding-right: 1.3rem !important;
            padding-top: 2.1rem !important;
            padding-bottom: 2.2rem !important;
        }
        [data-testid="stPlotlyChart"] {
            width: 100% !important;
        }

        /* Custom Sleek Dark Scrollbars */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: rgba(15, 23, 42, 0.6);
            border-radius: 6px;
        }
        ::-webkit-scrollbar-thumb {
            background: rgba(71, 85, 105, 0.7);
            border-radius: 6px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: rgba(59, 130, 246, 0.85);
        }

        /* 1. Sleek 200px Left Sidebar */
        section[data-testid="stSidebar"][aria-expanded="true"] {
            width: 200px !important;
            min-width: 200px !important;
            max-width: 200px !important;
            background: linear-gradient(180deg, rgba(15, 23, 42, 0.98) 0%, rgba(11, 15, 25, 0.98) 100%) !important;
            border-right: 1px solid rgba(148, 163, 184, 0.14) !important;
            box-shadow: 4px 0 24px rgba(0, 0, 0, 0.35) !important;
        }
        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 0.85rem !important;
            padding-left: 0.55rem !important;
            padding-right: 0.55rem !important;
        }
        section[data-testid="stSidebar"] .stMarkdown p,
        section[data-testid="stSidebar"] label {
            font-size: 0.79rem !important;
        }

        /* 2. Always-Visible Glowing Sidebar Collapse (<<) & Expand (>>) Button */
        [data-testid="stSidebarCollapseButton"],
        [data-testid="collapsedControl"],
        [data-testid="stExpandSidebarButton"] {
            opacity: 1 !important;
            visibility: visible !important;
            display: flex !important;
            z-index: 999995 !important;
        }
        [data-testid="stSidebarCollapseButton"] button,
        [data-testid="collapsedControl"] button,
        [data-testid="stExpandSidebarButton"] {
            opacity: 1 !important;
            visibility: visible !important;
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.95), rgba(15, 23, 42, 0.98)) !important;
            border: 1.5px solid rgba(59, 130, 246, 0.8) !important;
            border-radius: 10px !important;
            color: #93C5FD !important;
            box-shadow: 0 0 14px rgba(59, 130, 246, 0.4), 0 2px 6px rgba(0, 0, 0, 0.6) !important;
            width: 34px !important;
            height: 34px !important;
            min-height: 34px !important;
            transition: all 0.2s cubic-bezier(0.22, 1, 0.36, 1) !important;
        }
        [data-testid="stSidebarCollapseButton"] button:hover,
        [data-testid="collapsedControl"] button:hover,
        [data-testid="stExpandSidebarButton"]:hover {
            background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
            border-color: #BFDBFE !important;
            color: #FFFFFF !important;
            box-shadow: 0 0 20px rgba(59, 130, 246, 0.9) !important;
            transform: scale(1.07) !important;
        }
        [data-testid="stSidebarCollapseButton"] svg,
        [data-testid="collapsedControl"] svg,
        [data-testid="stExpandSidebarButton"] svg {
            fill: currentColor !important;
            color: inherit !important;
            width: 1.2rem !important;
            height: 1.2rem !important;
        }

        /* 3. Executive Command Header Banner */
        .exec-header-card {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem 1.2rem;
            padding: 0.85rem 1.15rem;
            margin-bottom: 0.85rem;
            border-radius: 14px;
            background:
                radial-gradient(circle at 0% 0%, rgba(59, 130, 246, 0.16), transparent 50%),
                linear-gradient(135deg, rgba(30, 41, 59, 0.65) 0%, rgba(15, 23, 42, 0.85) 100%);
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-left: 4px solid #3B82F6;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.08);
        }
        .exec-title-block {
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }
        .exec-icon-box {
            width: 46px;
            height: 46px;
            border-radius: 12px;
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.25), rgba(37, 99, 235, 0.1));
            border: 1px solid rgba(96, 165, 250, 0.4);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.55rem;
            box-shadow: 0 0 16px rgba(59, 130, 246, 0.25);
        }
        .exec-kicker {
            font-size: 0.68rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1.4px;
            color: #60A5FA;
            margin-bottom: 2px;
        }
        .exec-main-title {
            font-size: 1.65rem;
            font-weight: 800;
            letter-spacing: -0.4px;
            color: #F8FAFC;
            line-height: 1.15;
            margin: 0;
        }
        .exec-meta-cluster {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 0.5rem 0.75rem;
            font-size: 0.79rem;
            color: #CBD5E1;
        }
        .exec-live-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            padding: 0.22rem 0.65rem;
            border-radius: 999px;
            background: rgba(59, 130, 246, 0.16);
            border: 1px solid rgba(59, 130, 246, 0.45);
            color: #93C5FD;
            font-weight: 600;
            font-size: 0.76rem;
            box-shadow: 0 0 12px rgba(59, 130, 246, 0.2);
        }
        .pulse-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #38BDF8;
            box-shadow: 0 0 8px #38BDF8;
            animation: execPulse 2s infinite ease-in-out;
        }
        @keyframes execPulse {
            0%, 100% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.45; transform: scale(0.85); }
        }
        .exec-meta-tag {
            padding: 0.2rem 0.6rem;
            border-radius: 8px;
            background: rgba(15, 23, 42, 0.65);
            border: 1px solid rgba(148, 163, 184, 0.16);
            color: #E2E8F0;
            font-size: 0.76rem;
        }

        /* 4. Glassmorphic Top Filter & Radio Control Cards */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 13px !important;
            padding: 0.45rem 0.65rem !important;
            background: linear-gradient(155deg, rgba(30, 41, 59, 0.55) 0%, rgba(15, 23, 42, 0.78) 100%) !important;
            border: 1px solid rgba(148, 163, 184, 0.18) !important;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.24), inset 0 1px 0 rgba(255, 255, 255, 0.08) !important;
            transition: all 0.22s ease !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:hover {
            border-color: rgba(96, 165, 250, 0.48) !important;
            box-shadow: 0 8px 22px rgba(37, 99, 235, 0.15), inset 0 1px 0 rgba(255, 255, 255, 0.14) !important;
            transform: translateY(-1px) !important;
        }
        div[role="radiogroup"] {
            gap: 0.9rem !important;
            padding-top: 0.18rem !important;
        }

        /* 5. Bespoke Multi-Layered KPI Metric Cards */
        div[data-testid="stMetric"] {
            background:
                radial-gradient(circle at 85% 0%, rgba(59, 130, 246, 0.18), transparent 55%),
                linear-gradient(155deg, rgba(30, 41, 59, 0.62) 0%, rgba(15, 23, 42, 0.88) 100%) !important;
            border: 1px solid rgba(148, 163, 184, 0.18) !important;
            border-top: 3px solid #3B82F6 !important;
            border-radius: 14px !important;
            padding: 0.7rem 0.95rem !important;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.08) !important;
            transition: transform 0.2s ease, box-shadow 0.2s ease !important;
        }
        div[data-testid="stMetric"]:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 12px 26px rgba(0, 0, 0, 0.36) !important;
        }

        /* 2nd KPI Card: Violet Accent (Officers) */
        div[data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetric"],
        div[data-testid="column"]:nth-child(2) div[data-testid="stMetric"] {
            border-top: 3px solid #8B5CF6 !important;
            background:
                radial-gradient(circle at 85% 0%, rgba(139, 92, 246, 0.20), transparent 55%),
                linear-gradient(155deg, rgba(30, 27, 75, 0.45) 0%, rgba(15, 23, 42, 0.88) 100%) !important;
        }

        /* 3rd KPI Card: Cyan Accent (Total Works) */
        div[data-testid="stHorizontalBlock"] > div:nth-child(3) div[data-testid="stMetric"],
        div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"] {
            border-top: 3px solid #06B6D4 !important;
            background:
                radial-gradient(circle at 85% 0%, rgba(6, 182, 212, 0.18), transparent 55%),
                linear-gradient(155deg, rgba(22, 78, 99, 0.35) 0%, rgba(15, 23, 42, 0.88) 100%) !important;
        }

        /* 4th KPI Card: Crimson Red Glow (Overdue Works) */
        div[data-testid="stHorizontalBlock"] > div:nth-child(4) div[data-testid="stMetric"],
        div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] {
            border-top: 3px solid #EF4444 !important;
            border-color: rgba(239, 68, 68, 0.28) !important;
            background:
                radial-gradient(circle at 85% 0%, rgba(239, 68, 68, 0.26), transparent 60%),
                linear-gradient(145deg, rgba(127, 29, 29, 0.42) 0%, rgba(15, 23, 42, 0.90) 100%) !important;
            box-shadow: 0 8px 22px rgba(239, 68, 68, 0.12), inset 0 1px 0 rgba(255, 255, 255, 0.08) !important;
        }

        /* 5th KPI Card: Emerald Green Glow (ETS Works) */
        div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetric"],
        div[data-testid="column"]:nth-child(5) div[data-testid="stMetric"] {
            border-top: 3px solid #10B981 !important;
            border-color: rgba(16, 185, 129, 0.28) !important;
            background:
                radial-gradient(circle at 85% 0%, rgba(16, 185, 129, 0.24), transparent 60%),
                linear-gradient(145deg, rgba(6, 78, 59, 0.42) 0%, rgba(15, 23, 42, 0.90) 100%) !important;
            box-shadow: 0 8px 22px rgba(16, 185, 129, 0.12), inset 0 1px 0 rgba(255, 255, 255, 0.08) !important;
        }

        div[data-testid="stMetricValue"] {
            font-size: 1.72rem !important;
            font-weight: 800 !important;
            letter-spacing: -0.5px !important;
            color: #F8FAFC !important;
        }
        div[data-testid="stMetricLabel"] {
            font-size: 0.82rem !important;
            font-weight: 600 !important;
            color: #CBD5E1 !important;
        }

        /* 6. Section Header Banners */
        .exec-section-banner {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            gap: 0.5rem 1rem;
            padding: 0.55rem 0.95rem;
            margin-top: 0.25rem;
            margin-bottom: 0.65rem;
            border-radius: 11px;
            background: linear-gradient(90deg, rgba(30, 41, 59, 0.72) 0%, rgba(15, 23, 42, 0.45) 100%);
            border: 1px solid rgba(148, 163, 184, 0.16);
            border-left: 4px solid #3B82F6;
        }
        .exec-section-title {
            font-size: 1.15rem;
            font-weight: 700;
            color: #F8FAFC;
            letter-spacing: -0.2px;
        }
        .exec-scope-badge {
            font-size: 0.76rem;
            font-weight: 600;
            padding: 0.2rem 0.65rem;
            border-radius: 999px;
            background: rgba(59, 130, 246, 0.14);
            border: 1px solid rgba(96, 165, 250, 0.35);
            color: #93C5FD;
        }

        /* Selected Officer Summary Strip */
        .officer-summary-strip {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 0.55rem 0.9rem;
            padding: 0.5rem 0.85rem;
            margin-top: 0.15rem;
            margin-bottom: 0.65rem;
            border-radius: 10px;
            background: rgba(30, 41, 59, 0.50);
            border: 1px solid rgba(148, 163, 184, 0.16);
            font-size: 0.83rem;
        }
        .strip-badge-officer {
            background: rgba(59, 130, 246, 0.18);
            border: 1px solid rgba(96, 165, 250, 0.4);
            color: #BFDBFE;
            padding: 0.15rem 0.6rem;
            border-radius: 6px;
            font-weight: 700;
        }
        .strip-badge-overdue {
            background: rgba(239, 68, 68, 0.18);
            border: 1px solid rgba(248, 113, 113, 0.45);
            color: #FCA5A5;
            padding: 0.15rem 0.6rem;
            border-radius: 6px;
            font-weight: 700;
        }
        .strip-badge-hoa {
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid rgba(52, 211, 153, 0.38);
            color: #6EE7B7;
            padding: 0.15rem 0.6rem;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.79rem;
        }

        /* 7. Flagship CE Combined Overdue Works Table */
        .ce-table-wrapper {
            max-height: 590px;
            overflow-y: auto;
            overflow-x: auto;
            border: 1px solid rgba(148, 163, 184, 0.24);
            border-radius: 14px;
            background: rgba(15, 23, 42, 0.82);
            box-shadow: 0 10px 28px rgba(0, 0, 0, 0.35), inset 0 1px 0 rgba(255, 255, 255, 0.05);
            margin-top: 0.45rem;
            margin-bottom: 0.85rem;
        }
        table.ce-overdue-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.84rem;
            color: #E2E8F0;
        }
        table.ce-overdue-table thead th {
            position: sticky;
            top: 0;
            z-index: 10;
            background: linear-gradient(180deg, #2B394F 0%, #1E293B 100%);
            color: #F8FAFC;
            font-weight: 700;
            text-align: left;
            padding: 12px 10px;
            border-bottom: 2px solid #3B82F6;
            white-space: nowrap;
            font-size: 0.81rem;
            letter-spacing: 0.2px;
            box-shadow: 0 4px 8px -1px rgba(0, 0, 0, 0.45);
        }
        table.ce-overdue-table tbody tr {
            border-bottom: 1px solid rgba(148, 163, 184, 0.11);
            transition: background-color 0.15s ease;
        }
        table.ce-overdue-table tbody tr:nth-child(even) {
            background: rgba(30, 41, 59, 0.34);
        }
        table.ce-overdue-table tbody tr:hover {
            background: rgba(59, 130, 246, 0.16);
        }
        table.ce-overdue-table td {
            padding: 9px 10px;
            vertical-align: middle;
            line-height: 1.38;
        }

        /* 25% Compact Width + 4-Line Clamp for Work Name */
        .ce-work-name-cell {
            width: 25%;
            min-width: 285px;
            max-width: 375px;
            white-space: normal;
        }
        .ce-work-name-clamp {
            display: -webkit-box;
            -webkit-line-clamp: 4;
            -webkit-box-orient: vertical;
            overflow: hidden;
            text-overflow: ellipsis;
            line-height: 1.35;
            max-height: 5.4em;
            font-weight: 600;
            color: #F8FAFC;
        }
        .ce-compact-cell {
            white-space: nowrap;
            font-size: 0.82rem;
            color: #E2E8F0;
        }
        .ce-date-overdue {
            white-space: nowrap;
        }
        .ce-overdue-pill {
            display: inline-block;
            color: #FCA5A5;
            background: rgba(239, 68, 68, 0.14);
            border: 1px solid rgba(239, 68, 68, 0.32);
            border-radius: 6px;
            padding: 2px 7px;
            font-weight: 700;
            font-size: 0.80rem;
        }
        .ce-hoa-pill {
            display: inline-block;
            color: #93C5FD;
            background: rgba(59, 130, 246, 0.14);
            border: 1px solid rgba(59, 130, 246, 0.30);
            border-radius: 6px;
            padding: 2px 7px;
            font-weight: 600;
            font-size: 0.78rem;
            white-space: nowrap;
        }
        .ce-sl-cell {
            text-align: center;
            font-weight: 700;
            color: #94A3B8;
            width: 44px;
        }

        hr {
            margin-top: 0.65rem !important;
            margin-bottom: 0.75rem !important;
            border-color: rgba(148, 163, 184, 0.15) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# ULTRA-SLIM (14px), TEXT-FREE RIGHT-EDGE SCROLLER (WITH RELOAD FIX)
# -----------------------------------------------------------------------------
def inject_scroll_hud_graphics(enabled: bool = True):
    components.html(
        f"""
        <script>
        (function() {{
            const parentDoc = window.parent.document;
            const parentWin = window.parent;
            const isEnabled = {str(enabled).lower()};

            if (parentWin.__miScrollHudRAF) {{
                parentWin.cancelAnimationFrame(parentWin.__miScrollHudRAF);
            }}
            if (parentWin.__miScrollHandler) {{
                parentDoc.removeEventListener("scroll", parentWin.__miScrollHandler, {{ capture: true }});
            }}

            const oldHud = parentDoc.getElementById("mi-scroll-hud-root");
            if (oldHud) oldHud.remove();

            const oldStyle = parentDoc.getElementById("mi-scroll-hud-style");
            if (oldStyle) oldStyle.remove();

            parentDoc.querySelectorAll(
                "#mi-scroll-hud-active-topic, .mi-hud-section-Node, #mi-scroll-hud-crosshair"
            ).forEach(el => el.remove());

            if (!isEnabled) return;

            const style = parentDoc.createElement("style");
            style.id = "mi-scroll-hud-style";
            style.innerHTML = `
                @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@600;700&display=swap');

                #mi-scroll-hud-root {{
                    position: fixed !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 14px !important;
                    height: 100vh !important;
                    z-index: 999990 !important;
                    pointer-events: none !important;
                    user-select: none !important;
                    overflow: visible !important;
                    background: transparent !important;
                    border: none !important;
                    box-shadow: none !important;
                    font-family: 'JetBrains Mono', monospace !important;
                }}

                #mi-scroll-hud-track {{
                    position: absolute !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 14px !important;
                    height: 100vh !important;
                    overflow: hidden !important;
                    pointer-events: none !important;
                    background: transparent !important;
                }}

                #mi-scroll-hud-ticks-inner {{
                    position: absolute !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 100% !important;
                    will-change: transform;
                }}

                .mi-hud-tick {{
                    position: absolute;
                    right: 0;
                    display: flex;
                    align-items: center;
                    justify-content: flex-end;
                    transform: translateY(-50%);
                }}

                .mi-hud-tick-label {{
                    font-size: 7.5px;
                    font-weight: 600;
                    color: rgba(226, 232, 240, 0.42);
                    margin-right: 1.5px;
                    letter-spacing: -0.2px;
                    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.9);
                }}

                .mi-hud-tick-line-major {{
                    width: 4px;
                    height: 1.2px;
                    background: rgba(226, 232, 240, 0.50);
                }}

                .mi-hud-tick-line-mid {{
                    width: 2.5px;
                    height: 1px;
                    background: rgba(226, 232, 240, 0.24);
                }}

                #mi-scroll-hud-indicator {{
                    position: absolute;
                    right: 0;
                    width: 14px;
                    top: 50%;
                    transform: translateY(-1px);
                    pointer-events: none;
                    z-index: 6;
                }}

                #mi-scroll-hud-blue-line {{
                    width: 14px;
                    height: 2px;
                    background: linear-gradient(90deg, rgba(59, 130, 246, 0.2) 0%, #3B82F6 45%, #93C5FD 100%);
                    box-shadow: 0 0 6px rgba(59, 130, 246, 0.95);
                }}

                #mi-scroll-hud-pct {{
                    margin-top: 2px;
                    padding-right: 1px;
                    text-align: right;
                    font-size: 7.5px;
                    font-weight: 700;
                    color: #93C5FD;
                    white-space: nowrap;
                    text-shadow: 0 0 5px rgba(59, 130, 246, 0.95), 0 1px 2px rgba(0, 0, 0, 0.95);
                }}
            `;
            parentDoc.head.appendChild(style);

            const PX_PER_UNIT = 40;

            const hudRoot = parentDoc.createElement("div");
            hudRoot.id = "mi-scroll-hud-root";

            let ticksHtml = "";
            for (let step = -20; step <= 220; step++) {{
                const unitVal = step * 0.5;
                const yPos = unitVal * PX_PER_UNIT;

                if (step % 10 === 0 && unitVal >= 0 && unitVal <= 100) {{
                    ticksHtml += `
                        <div class="mi-hud-tick" style="top: ${{yPos}}px;">
                            <span class="mi-hud-tick-label">${{unitVal}}</span>
                            <div class="mi-hud-tick-line-major"></div>
                        </div>
                    `;
                }} else if (step % 5 === 0) {{
                    ticksHtml += `
                        <div class="mi-hud-tick" style="top: ${{yPos}}px;">
                            <div class="mi-hud-tick-line-mid"></div>
                        </div>
                    `;
                }}
            }}

            hudRoot.innerHTML = `
                <div id="mi-scroll-hud-track">
                    <div id="mi-scroll-hud-ticks-inner">${{ticksHtml}}</div>
                </div>
                <div id="mi-scroll-hud-indicator">
                    <div id="mi-scroll-hud-blue-line"></div>
                    <div id="mi-scroll-hud-pct">0%</div>
                </div>
            `;
            parentDoc.body.appendChild(hudRoot);

            parentWin.__miActiveScroller = null;
            parentWin.__miScrollHandler = function(e) {{
                const t = e.target;
                if (
                    t &&
                    t !== parentDoc &&
                    t.scrollHeight > t.clientHeight + 20 &&
                    t.clientHeight > parentWin.innerHeight * 0.45
                ) {{
                    parentWin.__miActiveScroller = t;
                }}
            }};
            parentDoc.addEventListener("scroll", parentWin.__miScrollHandler, {{ passive: true, capture: true }});

            function findScrollElement() {{
                const activeScroller = parentWin.__miActiveScroller;
                if (
                    activeScroller &&
                    parentDoc.body.contains(activeScroller) &&
                    activeScroller.scrollHeight > activeScroller.clientHeight + 10
                ) {{
                    return activeScroller;
                }}

                const selectors = [
                    '[data-testid="stMain"]',
                    '[data-testid="stAppViewContainer"]',
                    'section.main',
                    '.main',
                    '[data-testid="stApp"]'
                ];

                let bestCandidate = null;
                let bestRange = 0;

                for (const sel of selectors) {{
                    const el = parentDoc.querySelector(sel);
                    if (!el) continue;
                    const range = el.scrollHeight - el.clientHeight;
                    if (range > 10) {{
                        if (el.scrollTop > 0) {{
                            parentWin.__miActiveScroller = el;
                            return el;
                        }}
                        const st = parentWin.getComputedStyle(el);
                        if (
                            (st.overflowY === "auto" || st.overflowY === "scroll" || st.overflow === "auto") &&
                            range > bestRange
                        ) {{
                            bestRange = range;
                            bestCandidate = el;
                        }}
                    }}
                }}

                if (bestCandidate) return bestCandidate;
                return parentDoc.scrollingElement || parentDoc.documentElement;
            }}

            function getScrollPercentage() {{
                const winY = parentWin.scrollY || parentWin.pageYOffset || 0;
                const docEl = parentDoc.scrollingElement || parentDoc.documentElement;
                if (winY > 0 && docEl && docEl.scrollHeight > docEl.clientHeight) {{
                    return Math.max(0, Math.min(100, (winY / (docEl.scrollHeight - docEl.clientHeight)) * 100));
                }}

                const el = findScrollElement();
                if (!el) return 0;
                const maxScroll = Math.max(1, el.scrollHeight - el.clientHeight);
                const pct = (el.scrollTop / maxScroll) * 100;
                return Math.max(0, Math.min(100, pct));
            }}

            parentWin.__miCurrentPct = parentWin.__miCurrentPct || 0;

            parentWin.__miHudStepFn = function() {{
                const root = parentDoc.getElementById("mi-scroll-hud-root");
                const inner = parentDoc.getElementById("mi-scroll-hud-ticks-inner");
                const indicator = parentDoc.getElementById("mi-scroll-hud-indicator");
                const pctLabel = parentDoc.getElementById("mi-scroll-hud-pct");

                if (root && inner && indicator && pctLabel) {{
                    const targetPct = getScrollPercentage();
                    parentWin.__miCurrentPct += (targetPct - parentWin.__miCurrentPct) * 0.18;

                    const vh = root.clientHeight;
                    const indicatorY = vh * (0.24 + 0.52 * (parentWin.__miCurrentPct / 100.0));
                    indicator.style.top = `${{indicatorY.toFixed(2)}}px`;

                    const trackY = indicatorY - (parentWin.__miCurrentPct * PX_PER_UNIT);
                    inner.style.transform = `translate3d(0, ${{trackY.toFixed(2)}}px, 0)`;

                    const displayNum = Math.round(parentWin.__miCurrentPct);
                    pctLabel.textContent = `${{displayNum}}%`;
                }}
            }};

            function loop() {{
                if (parentWin.__miHudStepFn) {{
                    parentWin.__miHudStepFn();
                }}
                parentWin.__miScrollHudRAF = parentWin.requestAnimationFrame(loop);
            }}

            parentWin.__miScrollHudRAF = parentWin.requestAnimationFrame(loop);
        }})();
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

COL_WORK_NAME_KN = "Work Name"
COL_DIV_KN = "ವಿಭಾಗ"
COL_SUBDIV_KN = "ಉಪವಿಭಾಗ"
COL_SO_KN = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE_KN = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_START_DATE_KN = "ಕಾಮಗಾರಿ ಪ್ರಾರಂಭಿಸಿದ ದಿನಾಂಕ"
COL_TARGET_DATE_KN = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE_KN = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"
COL_HOA_KN = "ಲೆಕ್ಕಶೀರ್ಷಿಕೆ"
COL_STATUS_STD = "CP/CF/P/TBS/TP/ETS/OTH"

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

BAR_SHRINK_THRESHOLD = 250
BAR_SHRINK_MAX_EXTRA = 18.0

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

SCHEMA_VERSION = "v12_ce_flagship_dashboard"
DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v12.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state_v12() -> dict:
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
def calc_overdue_duration(target_date_obj, current_date_obj) -> str:
    """Calculates the exact elapsed time (Years, Months, Days) including today."""
    if not isinstance(target_date_obj, date) or target_date_obj >= current_date_obj:
        return "-"

    y = current_date_obj.year - target_date_obj.year
    m = current_date_obj.month - target_date_obj.month
    d = current_date_obj.day - target_date_obj.day + 1  # Include today

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


def compress_bar_length(val: float, threshold: int = BAR_SHRINK_THRESHOLD) -> float:
    v = float(val)
    t = float(threshold)
    if v <= t:
        return v
    excess = v - t
    return t + BAR_SHRINK_MAX_EXTRA * (excess / (excess + 120.0))


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
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0.0.0 Safari/537.36"
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
    parsed_dates = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            parsed_dates.append(None)
            continue

        if isinstance(val, date):
            parsed_dates.append(val)
            continue

        s = normalize_text(val)
        if not s or s.upper() in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", ""}:
            parsed_dates.append(None)
            continue

        s = re.sub(r"[./]", "-", s).strip(" -")
        match = re.search(r"(\d{1,4})-(\d{1,2})-(\d{1,4})", s)
        if match:
            p1, p2, p3 = int(match.group(1)), int(match.group(2)), int(match.group(3))
            try:
                if p1 > 31:
                    parsed_dates.append(date(max(1, min(9999, p1)), p2, p3))
                else:
                    yr = p3 + 2000 if p3 < 100 else p3
                    parsed_dates.append(date(max(1, min(9999, yr)), p2, p1))
                continue
            except Exception:
                pass

        try:
            dt = pd.to_datetime(s, dayfirst=True, utc=True, errors="coerce")
            if pd.isna(dt):
                parsed_dates.append(None)
            else:
                parsed_dates.append(date(dt.year, dt.month, dt.day))
        except Exception:
            parsed_dates.append(None)

    return parsed_dates


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


def shorten_label(name: str, max_len: int = 28) -> str:
    if name == UNASSIGNED_LABEL:
        return "⚠️ Unassigned (Empty)"
    return name if len(name) <= max_len else name[: max_len - 1] + "…"


def display_clean_cell(val: str) -> str:
    if not val or val in {UNSPECIFIED_LABEL, UNASSIGNED_LABEL}:
        return "-"
    return val


def render_ce_overdue_html_table(df_display: pd.DataFrame) -> str:
    headers_html = "".join(
        f"<th>{html.escape(str(col))}</th>" for col in df_display.columns
    )

    rows_html = []
    for idx_row, (_, row) in enumerate(df_display.iterrows(), start=1):
        cells_html = [f'<td class="ce-sl-cell">{idx_row}</td>']
        for col in df_display.columns[1:]:
            raw_text = str(row[col]) if pd.notna(row[col]) else "-"
            escaped_text = html.escape(raw_text if raw_text else "-")

            if col == "Work Name":
                cells_html.append(
                    f'<td class="ce-work-name-cell" title="{escaped_text}">'
                    f'<div class="ce-work-name-clamp">{escaped_text}</div>'
                    f"</td>"
                )
            elif col in ["Target Date", "Overdue By"]:
                cells_html.append(
                    f'<td class="ce-date-overdue">'
                    f'<span class="ce-overdue-pill">{escaped_text}</span>'
                    f"</td>"
                )
            elif col == "H.O.A.":
                cells_html.append(
                    f'<td class="ce-compact-cell">'
                    f'<span class="ce-hoa-pill">{escaped_text}</span>'
                    f"</td>"
                )
            else:
                cells_html.append(
                    f'<td class="ce-compact-cell">{escaped_text}</td>'
                )
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    return (
        f'<div class="ce-table-wrapper">'
        f'<table class="ce-overdue-table">'
        f"<thead><tr>{headers_html}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        f"</table>"
        f"</div>"
    )


def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Overdue_Works")
    return output.getvalue()


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

    total_targets = len(TARGET_SHEETS)
    for idx_sheet, target in enumerate(TARGET_SHEETS, start=1):
        progress.progress(
            idx_sheet / total_targets,
            text=f"Processing worksheet {idx_sheet}/{total_targets}: {target}...",
        )

        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Missing in Workbook",
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
            if raw_df.empty:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Empty Sheet",
                        "Sub-division Col": "N/A",
                        "EE / TA / AS Cols": "N/A",
                        "Rows Loaded": 0,
                        "Overdue ('P')": 0,
                        "ETS": 0,
                    }
                )
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
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Header Row Not Found",
                        "Sub-division Col": "N/A",
                        "EE / TA / AS Cols": "N/A",
                        "Rows Loaded": 0,
                        "Overdue ('P')": 0,
                        "ETS": 0,
                    }
                )
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

            missing_keys = required_keys - set(col_idx_map.keys())
            if missing_keys:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": f"Missing Columns: {', '.join(sorted(missing_keys))}",
                        "Sub-division Col": "Found" if "subdivision" in col_idx_map else "Missing",
                        "EE / TA / AS Cols": (
                            f"EE:{'Y' if 'ee' in col_idx_map else 'N'} "
                            f"TA:{'Y' if 'ta' in col_idx_map else 'N'} "
                            f"AS:{'Y' if 'as_off' in col_idx_map else 'N'}"
                        ),
                        "Rows Loaded": 0,
                        "Overdue ('P')": 0,
                        "ETS": 0,
                    }
                )
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
                        f"EE:{'Y' if 'ee' in col_idx_map else 'N'} "
                        f"TA:{'Y' if 'ta' in col_idx_map else 'N'} "
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
    return master_df, overdue_full_df, diag_df, today_ist_date.strftime("%d-%m-%Y")


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
    help="Ensure the Google Sheet is set to 'Anyone with the link can view'.",
)

uploaded_xlsx = st.sidebar.file_uploader(
    "Or Upload Workbook (.xlsx) Directly",
    type=["xlsx"],
    help="Optional: Upload a downloaded .xlsx copy to update all screens immediately.",
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

chart_style = st.sidebar.radio(
    "Bar Width Mode:",
    options=[
        "Overlay (Overdue + ETS + Rest of Works)",
        "Scale by Overdue Works (Full-Width Bars)",
        "Scale by Overdue Percentage (%)",
    ],
    index=0,
    help="Choose how the horizontal bars are scaled across the page.",
)

shrink_outliers = st.sidebar.checkbox(
    f"Shrink long bars above {BAR_SHRINK_THRESHOLD} works",
    value=True,
    help=f"Compresses bars above {BAR_SHRINK_THRESHOLD} works so all other bars look long and readable.",
)

hide_zero_overdue = st.sidebar.checkbox(
    "Show only officers with Overdue > 0",
    value=True,
)

show_unassigned_bar = st.sidebar.checkbox(
    "Include Empty / Unassigned Officer bar at the bottom",
    value=True,
    help="Groups all works where the officer name is blank into a single bar pinned at the bottom.",
)

show_scroll_hud = st.sidebar.checkbox(
    "Show Right-Edge Scroll HUD Ruler",
    value=True,
    help="Displays the ultra-slim 14px transparent scrolling ruler on the right edge.",
)

# -----------------------------------------------------------------------------
# THREAD-SAFE DATA LOADING (WITH SELF-HEALING CACHE CHECK)
# -----------------------------------------------------------------------------
shared_state = get_shared_state_v12()

needs_fresh_load = (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or sync_clicked
)

# Self-healing check if an older cached dataframe lacks the new columns
if not needs_fresh_load and shared_state["data"] is not None:
    _, cached_od_df, _, _ = shared_state["data"]
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
        with st.spinner(
            "Downloading workbook snapshot from Google Sheets (takes ~15-20s)..."
        ):
            try:
                xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                parsed = parse_xlsx_workbook(xlsx_bytes)
                save_shared_state(
                    shared_state, parsed, "Google Sheets Live Snapshot"
                )
            except Exception as e:
                if shared_state["data"] is not None:
                    st.warning(
                        f"Google Sheets is busy right now ({e}). Showing last synced snapshot from {shared_state['last_synced']}."
                    )
                else:
                    st.error(
                        f"Could not download from Google Sheets right now.\n\n"
                        f"**Tip:** Wait 10 seconds and click **Sync Latest Data**, or upload the `.xlsx` file in the sidebar.\n\n"
                        f"Details: {e}"
                    )
                    st.stop()

master_df, overdue_full_df, diag_df, today_str = shared_state["data"]

# -----------------------------------------------------------------------------
# EXECUTIVE COMMAND HEADER BANNER
# -----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="exec-header-card">
        <div class="exec-title-block">
            <div class="exec-icon-box">📊</div>
            <div>
                <div class="exec-kicker">Minor Irrigation Department • Executive Monitor</div>
                <h1 class="exec-main-title">Minor Irrigation — Overdue Works Dashboard</h1>
            </div>
        </div>
        <div class="exec-meta-cluster">
            <span class="exec-live-pill">
                <span class="pulse-dot"></span>
                ⚡ {shared_state['source_label']}
            </span>
            <span class="exec-meta-tag">🕒 <b>Last Synced:</b> {shared_state['last_synced']} IST</span>
            <span class="exec-meta-tag">📌 <b>Row Rule:</b> Valid text in <b>Division</b> column (empty <b>Sub-division</b> included)</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if master_df.empty:
    st.warning("No matching data rows were found across the target worksheets.")
    st.dataframe(diag_df, use_container_width=True)
    st.stop()

# -----------------------------------------------------------------------------
# TOP CONTROLS ROW 1: 5 CASCADING DROPDOWNS (Division, Sub-division, EE, TA, AS)
# -----------------------------------------------------------------------------
f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5, gap="small")

unique_divisions = sorted(
    [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
)
division_options = [ALL_OPTION] + unique_divisions

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

# -----------------------------------------------------------------------------
# TOP CONTROLS ROW 2: OFFICER VIEW & SORT ORDER
# -----------------------------------------------------------------------------
r_col1, r_col2 = st.columns(2, gap="medium")

with r_col1:
    with st.container(border=True):
        view_choice = st.radio(
            "👤 Select Officer View:",
            options=[LABEL_AE_EN, LABEL_AEE_EN],
            horizontal=True,
        )

with r_col2:
    with st.container(border=True):
        sort_choice = st.radio(
            "⬇️ Sort Chart & Table By:",
            options=[
                "Highest Overdue Count at Top",
                "Highest Overdue Percentage (%) at Top",
            ],
            horizontal=True,
        )

st.markdown("<div style='margin-top: 0.45rem;'></div>", unsafe_allow_html=True)

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
# KPI SUMMARY METRIC CARDS (5 COLUMNS WITH CUSTOM ACCENTS & RED/GREEN GLOWS)
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

unassigned_works_count = int(
    summary_df.loc[summary_df["_is_unassigned"], "total_works"].sum()
)

overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

active_scope_parts = [f"{LABEL_DIV_EN}: {selected_division}"]
if selected_subdivision != ALL_OPTION:
    active_scope_parts.append(f"{LABEL_SUBDIV_EN}: {selected_subdivision}")
if selected_ee != ALL_OPTION:
    active_scope_parts.append(f"EE: {selected_ee}")
if selected_ta != ALL_OPTION:
    active_scope_parts.append(f"TA: {selected_ta}")
if selected_as != ALL_OPTION:
    active_scope_parts.append(f"AS: {selected_as}")
filter_scope_label = " | ".join(active_scope_parts)

k1, k2, k3, k4, k5 = st.columns([1.1, 1.0, 1.05, 1.35, 1.0], gap="medium")
k1.metric("📅 Date (Today)", today_str)

officer_help_text = (
    f"Officers with >0 Overdue works (plus {unassigned_works_count:,} unassigned works)."
    if hide_zero_overdue
    else f"All named {view_choice}s (plus {unassigned_works_count:,} unassigned works)."
)

k2.metric(
    f"👥 {view_choice}s",
    f"{named_officers_count:,}",
    help=officer_help_text,
)
k3.metric("📋 Total Works", f"{total_assigned_all:,}")
k4.metric(
    "🔴 Overdue Works ('Progress')",
    f"{total_overdue_all:,} ({overall_overdue_pct}%)",
)
k5.metric("🟢 'ETS' Works", f"{total_ets_all:,}")

st.divider()

# =============================================================================
# SECTION 1: HORIZONTAL BAR CHART
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📈 Overdue & ETS Works by {view_choice} ({sort_choice})</div>
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
    plot_df = chart_source_df.iloc[::-1].copy()
    plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_label(s, 28))
    plot_df["is_shrunk"] = shrink_outliers & (
        plot_df["total_works"] > BAR_SHRINK_THRESHOLD
    )

    def _build_bar_label(r) -> str:
        shrink_tag = (
            f" <span style='color:#FBBF24'>[>{BAR_SHRINK_THRESHOLD} Shrunk]</span>"
            if r["is_shrunk"]
            and chart_style == "Overlay (Overdue + ETS + Rest of Works)"
            else ""
        )
        if sort_choice == "Highest Overdue Percentage (%) at Top":
            return (
                f"  <b>{r['overdue_pct']:.1f}% Overdue</b> "
                f"({int(r['overdue_works'])} Overdue | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])} Total){shrink_tag}"
            )
        return (
            f"  <b>{int(r['overdue_works'])} Overdue</b> | "
            f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
            f"{int(r['total_works'])} Total ({r['overdue_pct']:.1f}%){shrink_tag}"
        )

    plot_df["bar_label"] = plot_df.apply(_build_bar_label, axis=1)

    fig = go.Figure()

    if chart_style == "Scale by Overdue Works (Full-Width Bars)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_works"],
                name="Overdue Works ('P')",
                orientation="h",
                marker=dict(
                    color="#EF4444",
                    line=dict(color="#DC2626", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13, color="#F8FAFC"),
                cliponaxis=False,
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "ongoing_p_works",
                        "ets_works",
                        "rest_works",
                        "overdue_pct",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    "<b>Overdue Works ('P'):</b> %{x}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Total Works Assigned:</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )
        max_x = max(int(plot_df["overdue_works"].max()), 1)
        x_range_max = max_x * 1.25
        x_axis_title = "Number of Overdue Works"

    elif chart_style == "Scale by Overdue Percentage (%)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_pct"],
                name="Overdue %",
                orientation="h",
                marker=dict(
                    color="#EF4444",
                    line=dict(color="#DC2626", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13, color="#F8FAFC"),
                cliponaxis=False,
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "ongoing_p_works",
                        "overdue_works",
                        "ets_works",
                        "rest_works",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    "<b>Overdue Percentage:</b> %{x:.1f}%<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[3]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[4]}<br>"
                    "<b>Rest of Works:</b> %{customdata[5]}<br>"
                    "<b>Total Works Assigned:</b> %{customdata[1]}<extra></extra>"
                ),
            )
        )
        max_x = max(float(plot_df["overdue_pct"].max()), 1.0)
        x_range_max = max_x * 1.25
        x_axis_title = "Overdue Percentage (%)"

    else:
        plot_df["overdue_plus_ets"] = plot_df["overdue_works"] + plot_df["ets_works"]

        if shrink_outliers:
            plot_df["vis_total"] = plot_df["total_works"].map(compress_bar_length)
            plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].map(
                compress_bar_length
            )
            plot_df["vis_overdue"] = plot_df["overdue_works"].map(compress_bar_length)
        else:
            plot_df["vis_total"] = plot_df["total_works"].astype(float)
            plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].astype(float)
            plot_df["vis_overdue"] = plot_df["overdue_works"].astype(float)

        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["vis_total"],
                name="Rest of Works (Other Statuses)",
                orientation="h",
                marker=dict(
                    color="rgba(148, 163, 184, 0.28)",
                    line=dict(color="rgba(148, 163, 184, 0.55)", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13, color="#F8FAFC"),
                cliponaxis=False,
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "overdue_works",
                        "ets_works",
                        "rest_works",
                        "overdue_pct",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[2]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )

        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["vis_overdue_plus_ets"],
                name="ETS Works ('ETS')",
                orientation="h",
                marker=dict(color="#10B981"),
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "overdue_works",
                        "ets_works",
                        "rest_works",
                        "overdue_pct",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[2]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )

        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["vis_overdue"],
                name="Overdue Works ('P')",
                orientation="h",
                marker=dict(color="#EF4444"),
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "ongoing_p_works",
                        "overdue_pct",
                        "ets_works",
                        "rest_works",
                        "overdue_works",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[6]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[4]}<br>"
                    "<b>Rest of Works:</b> %{customdata[5]}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[3]:.1f}%<extra></extra>"
                ),
            )
        )

        max_x = max(float(plot_df["vis_total"].max()), 1.0)
        x_range_max = max_x * 1.24
        x_axis_title = "Number of Works"

    chart_height = max(480, len(plot_df) * 36 + 90)

    fig.update_layout(
        barmode="overlay",
        height=chart_height,
        bargap=0.22,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15, 23, 42, 0.42)",
        xaxis=dict(
            title=x_axis_title,
            range=[0, x_range_max],
            gridcolor="rgba(148, 163, 184, 0.12)",
            zeroline=True,
            zerolinecolor="rgba(148, 163, 184, 0.28)",
        ),
        yaxis=dict(
            title="",
            automargin=False,
            tickmode="array",
            tickvals=plot_df[view_choice],
            ticktext=plot_df["y_short"],
            tickfont=dict(size=13, color="#E2E8F0"),
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="left",
            x=0,
            traceorder="reversed",
            font=dict(size=12, color="#CBD5E1"),
        ),
        margin=dict(l=210, r=14, t=35, b=40),
    )

    st.plotly_chart(fig, use_container_width=True)

st.divider()

# =============================================================================
# SECTION 2: COMBINED OVERDUE WORKS TABLE (ENGLISH HEADERS + STATUS BADGES)
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📂 Combined Overdue Works by {view_choice}</div>
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
        f"{name}  —  ({officer_stats_map[name][0]} Overdue | "
        f"{officer_stats_map[name][1]} ETS / "
        f"{officer_stats_map[name][2]} Total | {officer_stats_map[name][3]:.1f}%)"
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
            hoa_involved = officer_overdue_df["H.O.A."].unique().tolist()
            st.markdown(
                f"""
                <div class="officer-summary-strip">
                    <span>👤 <b>Selected {html.escape(view_choice)}:</b> <span class="strip-badge-officer">{html.escape(selected_officer)}</span></span>
                    <span>🔴 <b>Total Overdue Works:</b> <span class="strip-badge-overdue">{len(officer_overdue_df)}</span></span>
                    <span>📑 <b>Account Heads (H.O.A.):</b> <span class="strip-badge-hoa">{html.escape(', '.join(hoa_involved))}</span></span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Unassigned"
            )
            safe_div = (
                re.sub(r"[^\w\-]+", "_", selected_division).strip("_") or "All_Divisions"
            )
            excel_bytes = convert_df_to_excel_bytes(officer_overdue_df)
            csv_bytes = officer_overdue_df.to_csv(index=False).encode("utf-8-sig")

            dl_col1, dl_col2, _ = st.columns([1.5, 1.5, 5])
            with dl_col1:
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Overdue_Works_{safe_div}_{safe_name}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl_col2:
                st.download_button(
                    label="📥 Download CSV (.csv)",
                    data=csv_bytes,
                    file_name=f"Overdue_Works_{safe_div}_{safe_name}_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            # Render the Flagship CE HTML table
            st.markdown(
                render_ce_overdue_html_table(officer_overdue_df),
                unsafe_allow_html=True,
            )

st.divider()

# =============================================================================
# SECTION 3: OFFICER-WISE SUMMARY TABLE (ENGLISH)
# =============================================================================
st.markdown(
    f"""
    <div class="exec-section-banner">
        <div class="exec-section-title">📋 Officer-Wise Summary Table</div>
        <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

display_table = summary_df[
    [
        view_choice,
        "total_works",
        "overdue_works",
        "ets_works",
        "rest_works",
        "ongoing_p_works",
        "overdue_pct",
    ]
].rename(
    columns={
        "total_works": "Total Works Assigned",
        "overdue_works": "Overdue Works ('P')",
        "ets_works": "ETS Works ('ETS')",
        "rest_works": "Rest of Works",
        "ongoing_p_works": "Ongoing Works ('P')",
        "overdue_pct": "Overdue %",
    }
)
display_table.index = display_table.index + 1

st.dataframe(
    display_table.style.format({"Overdue %": "{:.1f}%"}),
    use_container_width=True,
)

with st.expander("🔍 Inspect Worksheet Loading Status (23 Target Sheets)"):
    st.dataframe(diag_df, use_container_width=True)

# Mount the ultra-slim 14px, writing-free right-edge scrolling HUD overlay
inject_scroll_hud_graphics(enabled=show_scroll_hud)
