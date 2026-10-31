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
# PAGE CONFIGURATION & NAVIGATION STATE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Overdue Works Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "pc_active_page" not in st.session_state:
    st.session_state["pc_active_page"] = "dashboard"


def clean_html(raw_html: str) -> str:
    """Strips leading line indentation so Markdown never renders HTML as a code block."""
    return "\n".join(
        line.strip() for line in raw_html.splitlines() if line.strip()
    )


st.markdown(
    clean_html(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Montserrat:wght@700;800&display=swap');

            .block-container {
                max-width: 100% !important;
                padding-left: 1.1rem !important;
                padding-right: 1.2rem !important;
                padding-top: 2.3rem !important;
                padding-bottom: 2rem !important;
            }
            [data-testid="stPlotlyChart"] {
                width: 100% !important;
            }

            /* 3D Depth & Hover Growth on Plotly Donut Slices */
            .js-plotly-plot .pie path.surface {
                transition: transform 0.24s cubic-bezier(0.22, 1, 0.36, 1), filter 0.24s ease !important;
                transform-box: fill-box !important;
                transform-origin: center center !important;
                filter: drop-shadow(0px 6px 6px rgba(0, 0, 0, 0.55));
            }
            .js-plotly-plot .pie g.slice:hover path.surface {
                transform: scale(1.055) translateY(-4px) !important;
                filter: drop-shadow(0px 12px 14px rgba(0, 0, 0, 0.78)) brightness(1.08) !important;
            }

            /* 1. Narrower, Sleek Left Sidebar (200px) */
            section[data-testid="stSidebar"][aria-expanded="true"] {
                width: 200px !important;
                min-width: 200px !important;
                max-width: 200px !important;
                background: linear-gradient(180deg, rgba(15, 23, 42, 0.98) 0%, rgba(15, 23, 42, 0.92) 100%) !important;
                border-right: 1px solid rgba(148, 163, 184, 0.14) !important;
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

            /* 2. Always-Visible Sidebar Collapse (<<) & Expand (>>) Button */
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
                background: #1E293B !important;
                border: 1px solid rgba(148, 163, 184, 0.45) !important;
                border-radius: 10px !important;
                color: #E2E8F0 !important;
                box-shadow: none !important;
                width: 34px !important;
                height: 34px !important;
                min-height: 34px !important;
                transition: all 0.18s ease-in-out !important;
            }
            [data-testid="stSidebarCollapseButton"] button:hover,
            [data-testid="collapsedControl"] button:hover,
            [data-testid="stExpandSidebarButton"]:hover {
                background: #334155 !important;
                border-color: #94A3B8 !important;
                color: #FFFFFF !important;
                box-shadow: none !important;
            }
            [data-testid="stSidebarCollapseButton"] svg,
            [data-testid="collapsedControl"] svg,
            [data-testid="stExpandSidebarButton"] svg {
                fill: currentColor !important;
                color: inherit !important;
                width: 1.2rem !important;
                height: 1.2rem !important;
            }

            /* 3. Top Filter Control Cards */
            div[data-testid="stVerticalBlockBorderWrapper"] {
                border-radius: 12px !important;
                padding: 0.4rem 0.6rem !important;
                background: linear-gradient(145deg, rgba(30, 41, 59, 0.48), rgba(15, 23, 42, 0.68)) !important;
                border: 1px solid rgba(148, 163, 184, 0.18) !important;
                box-shadow: 0 4px 14px rgba(0, 0, 0, 0.22) !important;
            }
            div[role="radiogroup"] {
                gap: 0.8rem !important;
                padding-top: 0.15rem !important;
            }

            /* 4. KPI Metric Cards */
            div[data-testid="stMetric"] {
                border: 1px solid rgba(148, 163, 184, 0.18) !important;
                border-radius: 12px !important;
                padding: 0.6rem 0.85rem !important;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.22) !important;
                transition: transform 0.22s cubic-bezier(0.22, 1, 0.36, 1),
                            box-shadow 0.22s ease,
                            border-color 0.22s ease !important;
                cursor: pointer;
            }
            div[data-testid="stMetric"]:hover {
                transform: translateY(-3px) scale(1.02) !important;
                box-shadow: 0 8px 20px rgba(0, 0, 0, 0.35) !important;
            }

            div[data-testid="stHorizontalBlock"] > div:nth-child(1) div[data-testid="stMetric"],
            div[data-testid="column"]:nth-child(1) div[data-testid="stMetric"] {
                border-top: 3px solid #3B82F6 !important;
                background: linear-gradient(145deg, rgba(30, 58, 138, 0.40), rgba(15, 23, 42, 0.85)) !important;
            }
            div[data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetric"],
            div[data-testid="column"]:nth-child(2) div[data-testid="stMetric"] {
                border-top: 3px solid #8B5CF6 !important;
                background: linear-gradient(145deg, rgba(88, 28, 135, 0.38), rgba(15, 23, 42, 0.85)) !important;
            }
            div[data-testid="stHorizontalBlock"] > div:nth-child(3) div[data-testid="stMetric"],
            div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"] {
                border-top: 3px solid #06B6D4 !important;
                background: linear-gradient(145deg, rgba(22, 78, 99, 0.42), rgba(15, 23, 42, 0.85)) !important;
            }
            div[data-testid="stHorizontalBlock"] > div:nth-child(4) div[data-testid="stMetric"],
            div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] {
                border-top: 3px solid #EF4444 !important;
                background: linear-gradient(145deg, rgba(127, 29, 29, 0.38), rgba(15, 23, 42, 0.85)) !important;
            }
            div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetric"],
            div[data-testid="column"]:nth-child(5) div[data-testid="stMetric"] {
                border-top: 3px solid #4ADE80 !important;
                background: linear-gradient(145deg, rgba(6, 78, 59, 0.38), rgba(15, 23, 42, 0.85)) !important;
            }

            div[data-testid="stMetricValue"] {
                font-size: 1.6rem !important;
                font-weight: 700 !important;
            }
            div[data-testid="stMetricLabel"] {
                font-size: 0.82rem !important;
                color: #CBD5E1 !important;
            }

            /* Section Header Banners */
            .exec-section-banner {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                justify-content: space-between;
                gap: 0.5rem 1rem;
                padding: 0.6rem 1rem;
                margin-top: 0.25rem;
                margin-bottom: 0.65rem;
                border-radius: 10px;
                background: rgba(15, 23, 42, 0.6);
                border: 1px solid rgba(148, 163, 184, 0.16);
                border-left: 4px solid #3B82F6;
            }
            .exec-section-title {
                font-size: 1.35rem;
                font-weight: 700;
                color: #F8FAFC;
            }
            .exec-scope-badge {
                font-size: 0.78rem;
                font-weight: 600;
                padding: 0.25rem 0.75rem;
                border-radius: 999px;
                background: rgba(59, 130, 246, 0.12);
                border: 1px solid rgba(96, 165, 250, 0.3);
                color: #93C5FD;
            }

            /* Main Header */
            .premium-compact-header {
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 0.5rem;
                padding: 0.2rem 0;
                margin-top: 0.15rem;
                margin-bottom: 0.5rem;
                background: transparent;
                border: none;
                width: 100%;
                overflow: hidden;
            }
            .premium-compact-title {
                font-family: 'Montserrat', sans-serif !important;
                font-size: clamp(0.92rem, 4.2vw, 1.75rem) !important;
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
            }

            /* Clean Matte-Glass Navigation Button (NO Outside Glow) */
            div[data-testid="stHorizontalBlock"]:has(.clean-nav-marker) div[data-testid="stButton"] > button {
                background: linear-gradient(180deg, rgba(30, 41, 59, 0.92) 0%, rgba(15, 23, 42, 0.98) 100%) !important;
                backdrop-filter: blur(10px) !important;
                -webkit-backdrop-filter: blur(10px) !important;
                border: 1px solid rgba(148, 163, 184, 0.38) !important;
                border-top: 1px solid rgba(226, 232, 240, 0.45) !important;
                border-radius: 10px !important;
                color: #F8FAFC !important;
                font-family: 'Montserrat', sans-serif !important;
                font-weight: 700 !important;
                font-size: 0.90rem !important;
                letter-spacing: 0.3px !important;
                padding: 0.48rem 1.15rem !important;
                min-height: 42px !important;
                box-shadow: none !important;
                transition: background 0.18s ease, border-color 0.18s ease, transform 0.15s ease !important;
            }
            div[data-testid="stHorizontalBlock"]:has(.clean-nav-marker) div[data-testid="stButton"] > button:hover {
                background: linear-gradient(180deg, rgba(51, 65, 85, 0.95) 0%, rgba(30, 41, 59, 1) 100%) !important;
                border-color: #93C5FD !important;
                color: #FFFFFF !important;
                box-shadow: none !important;
                transform: translateY(-1px) !important;
            }
            div[data-testid="stHorizontalBlock"]:has(.clean-nav-marker) div[data-testid="stButton"] > button:active {
                transform: translateY(0px) !important;
                box-shadow: none !important;
            }

            /* Status Bar Banner */
            .mi-status-bar {
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 0.6rem 1.1rem;
                padding: 0.42rem 0.85rem;
                margin-top: -0.35rem;
                margin-bottom: 0.75rem;
                border-radius: 10px;
                background: rgba(30, 41, 59, 0.45);
                border: 1px solid rgba(148, 163, 184, 0.14);
                font-size: 0.81rem;
                color: #CBD5E1;
            }
            .mi-status-pill {
                display: inline-flex;
                align-items: center;
                gap: 0.35rem;
                padding: 0.12rem 0.55rem;
                border-radius: 999px;
                background: rgba(59, 130, 246, 0.15);
                border: 1px solid rgba(59, 130, 246, 0.35);
                color: #93C5FD;
                font-weight: 600;
                font-size: 0.77rem;
            }

            /* 5. CE-Friendly Combined Overdue Works Table */
            .ce-table-wrapper {
                max-height: 580px;
                overflow-y: auto;
                overflow-x: auto;
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 12px;
                background: rgba(15, 23, 42, 0.75);
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.28);
                margin-top: 0.4rem;
                margin-bottom: 0.8rem;
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
                background: #273346;
                color: #F1F5F9;
                font-weight: 700;
                text-align: left;
                padding: 11px 10px;
                border-bottom: 2px solid #3B82F6;
                white-space: nowrap;
                font-size: 0.82rem;
            }
            table.ce-overdue-table tbody tr {
                border-bottom: 1px solid rgba(148, 163, 184, 0.12);
            }
            table.ce-overdue-table tbody tr:nth-child(even) {
                background: rgba(30, 41, 59, 0.32);
            }
            table.ce-overdue-table tbody tr:hover {
                background: rgba(59, 130, 246, 0.14);
            }
            table.ce-overdue-table td {
                padding: 8px 10px;
                vertical-align: middle;
                line-height: 1.35;
            }
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
            }
            .ce-date-overdue {
                color: #FCA5A5;
                font-weight: 700;
                white-space: nowrap;
            }
            .ce-sl-cell {
                text-align: center;
                font-weight: 600;
                color: #94A3B8;
                width: 44px;
            }

            /* 6. Power BI Executive Analytics Page Styling */
            .pbi-kpi-grid {
                display: grid;
                grid-template-columns: repeat(6, 1fr);
                gap: 12px;
                margin-top: 6px;
                margin-bottom: 16px;
            }
            .pbi-kpi-card {
                background: linear-gradient(145deg, rgba(30, 41, 59, 0.72), rgba(15, 23, 42, 0.92));
                border: 1px solid rgba(148, 163, 184, 0.18);
                border-radius: 12px;
                padding: 12px 14px;
                box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
            }
            .pbi-kpi-label {
                font-size: 0.73rem;
                font-weight: 700;
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
                font-size: 0.72rem;
                font-weight: 600;
                margin-top: 4px;
                color: #CBD5E1;
            }

            .truth-audit-box {
                background: linear-gradient(135deg, rgba(30, 41, 59, 0.78), rgba(15, 23, 42, 0.95));
                border: 1px solid rgba(148, 163, 184, 0.25);
                border-left: 5px solid #EF4444;
                border-radius: 12px;
                padding: 16px 18px;
                margin-top: 6px;
                margin-bottom: 14px;
            }

            /* Redesigned Spacious Power BI Leaderboard Cards */
            .exec-lb-panel {
                background: linear-gradient(160deg, rgba(30, 41, 59, 0.62), rgba(15, 23, 42, 0.94));
                border: 1px solid rgba(148, 163, 184, 0.20);
                border-radius: 14px;
                padding: 16px 18px;
                margin-bottom: 14px;
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25);
            }
            .exec-lb-panel-danger {
                border-top: 4px solid #EF4444;
            }
            .exec-lb-panel-success {
                border-top: 4px solid #10B981;
            }
            .exec-lb-header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                margin-bottom: 12px;
                padding-bottom: 8px;
                border-bottom: 1px solid rgba(148, 163, 184, 0.16);
            }
            .exec-lb-title {
                font-size: 1.0rem;
                font-weight: 800;
                color: #F8FAFC;
            }
            .exec-lb-sub {
                font-size: 0.74rem;
                font-weight: 600;
                color: #94A3B8;
                background: rgba(15, 23, 42, 0.7);
                padding: 3px 9px;
                border-radius: 999px;
                border: 1px solid rgba(148, 163, 184, 0.2);
            }
            .exec-lb-row {
                background: rgba(15, 23, 42, 0.65);
                border: 1px solid rgba(148, 163, 184, 0.14);
                border-radius: 10px;
                padding: 10px 12px;
                margin-bottom: 8px;
                transition: background 0.15s ease, border-color 0.15s ease;
            }
            .exec-lb-row:hover {
                background: rgba(30, 41, 59, 0.85);
                border-color: rgba(148, 163, 184, 0.32);
            }
            .exec-lb-row-top {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 10px;
                margin-bottom: 6px;
            }
            .exec-lb-officer-wrap {
                display: flex;
                align-items: center;
                gap: 8px;
                min-width: 0;
            }
            .exec-lb-rank {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                min-width: 28px;
                height: 24px;
                border-radius: 6px;
                font-size: 0.75rem;
                font-weight: 800;
                background: rgba(59, 130, 246, 0.18);
                color: #93C5FD;
                border: 1px solid rgba(59, 130, 246, 0.35);
                flex-shrink: 0;
            }
            .exec-lb-name {
                font-size: 0.90rem;
                font-weight: 700;
                color: #F8FAFC;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .exec-lb-div-tag {
                font-size: 0.71rem;
                font-weight: 600;
                color: #94A3B8;
                background: rgba(148, 163, 184, 0.12);
                padding: 2px 7px;
                border-radius: 6px;
                white-space: nowrap;
                flex-shrink: 0;
            }
            .exec-lb-badges {
                display: flex;
                align-items: center;
                gap: 6px;
                flex-shrink: 0;
            }
            .exec-pill-danger {
                background: rgba(239, 68, 68, 0.18);
                color: #FCA5A5;
                border: 1px solid rgba(239, 68, 68, 0.42);
                padding: 2px 9px;
                border-radius: 999px;
                font-weight: 800;
                font-size: 0.76rem;
                white-space: nowrap;
            }
            .exec-pill-success {
                background: rgba(16, 185, 129, 0.18);
                color: #6EE7B7;
                border: 1px solid rgba(16, 185, 129, 0.42);
                padding: 2px 9px;
                border-radius: 999px;
                font-weight: 800;
                font-size: 0.76rem;
                white-space: nowrap;
            }
            .exec-pill-ets {
                background: rgba(245, 158, 11, 0.16);
                color: #FCD34D;
                border: 1px solid rgba(245, 158, 11, 0.35);
                padding: 2px 8px;
                border-radius: 999px;
                font-weight: 700;
                font-size: 0.73rem;
                white-space: nowrap;
            }
            .exec-pill-total {
                color: #38BDF8;
                font-weight: 700;
                font-size: 0.75rem;
                white-space: nowrap;
            }
            .exec-lb-bar-bg {
                width: 100%;
                height: 6px;
                background: rgba(148, 163, 184, 0.16);
                border-radius: 999px;
                overflow: hidden;
                display: flex;
            }
            .exec-lb-bar-od {
                background: linear-gradient(90deg, #EF4444, #F87171);
                height: 100%;
            }
            .exec-lb-bar-ets {
                background: linear-gradient(90deg, #10B981, #34D399);
                height: 100%;
            }

            hr {
                margin-top: 0.55rem !important;
                margin-bottom: 0.65rem !important;
                border-color: rgba(148, 163, 184, 0.15) !important;
            }
        </style>
        """
    ),
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# ULTRA-SLIM (14px), TEXT-FREE RIGHT-EDGE SCROLLER
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

            if (!isEnabled) return;

            const style = parentDoc.createElement("style");
            style.id = "mi-scroll-hud-style";
            style.innerHTML = `
                #mi-scroll-hud-root {{
                    position: fixed !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 14px !important;
                    height: 100vh !important;
                    z-index: 999990 !important;
                    pointer-events: none !important;
                    user-select: none !important;
                    background: transparent !important;
                    font-family: monospace !important;
                }}
                #mi-scroll-hud-track {{
                    position: absolute !important;
                    top: 0 !important;
                    right: 0 !important;
                    width: 14px !important;
                    height: 100vh !important;
                    overflow: hidden !important;
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
                    z-index: 6;
                }}
                #mi-scroll-hud-blue-line {{
                    width: 14px;
                    height: 2px;
                    background: linear-gradient(90deg, rgba(59, 130, 246, 0.2) 0%, #3B82F6 45%, #93C5FD 100%);
                }}
                #mi-scroll-hud-pct {{
                    margin-top: 2px;
                    padding-right: 1px;
                    text-align: right;
                    font-size: 7.5px;
                    font-weight: 700;
                    color: #93C5FD;
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
                    ticksHtml += `<div class="mi-hud-tick" style="top: ${{yPos}}px;"><span class="mi-hud-tick-label">${{unitVal}}</span><div class="mi-hud-tick-line-major"></div></div>`;
                }} else if (step % 5 === 0) {{
                    ticksHtml += `<div class="mi-hud-tick" style="top: ${{yPos}}px;"><div class="mi-hud-tick-line-mid"></div></div>`;
                }}
            }}

            hudRoot.innerHTML = `
                <div id="mi-scroll-hud-track"><div id="mi-scroll-hud-ticks-inner">${{ticksHtml}}</div></div>
                <div id="mi-scroll-hud-indicator"><div id="mi-scroll-hud-blue-line"></div><div id="mi-scroll-hud-pct">0%</div></div>
            `;
            parentDoc.body.appendChild(hudRoot);

            parentWin.__miActiveScroller = null;
            parentWin.__miScrollHandler = function(e) {{
                const t = e.target;
                if (t && t !== parentDoc && t.scrollHeight > t.clientHeight + 20 && t.clientHeight > parentWin.innerHeight * 0.45) {{
                    parentWin.__miActiveScroller = t;
                }}
            }};
            parentDoc.addEventListener("scroll", parentWin.__miScrollHandler, {{ passive: true, capture: true }});

            function findScrollElement() {{
                const activeScroller = parentWin.__miActiveScroller;
                if (activeScroller && parentDoc.body.contains(activeScroller) && activeScroller.scrollHeight > activeScroller.clientHeight + 10) {{
                    return activeScroller;
                }}
                const selectors = ['[data-testid="stMain"]', '[data-testid="stAppViewContainer"]', 'section.main', '.main', '[data-testid="stApp"]'];
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
                        if ((st.overflowY === "auto" || st.overflowY === "scroll" || st.overflow === "auto") && range > bestRange) {{
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
                return Math.max(0, Math.min(100, (el.scrollTop / maxScroll) * 100));
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
                    pctLabel.textContent = `${{Math.round(parentWin.__miCurrentPct)}}%`;
                }}
            }};

            function loop() {{
                if (parentWin.__miHudStepFn) parentWin.__miHudStepFn();
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

SCHEMA_VERSION = "v12_ce_overdue_analytics"
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
                    f'<td class="ce-date-overdue">{escaped_text}</td>'
                )
            else:
                cells_html.append(
                    f'<td class="ce-compact-cell">{escaped_text}</td>'
                )
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    return clean_html(
        f"""
        <div class="ce-table-wrapper">
            <table class="ce-overdue-table">
                <thead><tr>{headers_html}</tr></thead>
                <tbody>{''.join(rows_html)}</tbody>
            </table>
        </div>
        """
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
                    "ಕಾಮಗಾರಿ ಹೆಸರು": s_work_name,
                    "Work Name": s_work_name,
                    "division_name": s_div_raw,
                    "subdivision_name": s_subdiv_mapped,
                    "ee_name": s_ee_mapped,
                    "ta_name": s_ta_mapped,
                    "as_name": s_as_mapped,
                    "so_name": s_so_mapped,
                    "aee_name": s_aee_mapped,
                    "status_raw": s_status,
                    "target_date_raw": s_target_Formatted,
                    "Target Date": s_target_Formatted,
                    "Overdue By": s_overdue_duration,
                    "H.O.A.": s_hoa,
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
# TOP HEADER BAR WITH CENTERED TITLE & MATTE NAVIGATION BUTTON (NO GLOW)
# -----------------------------------------------------------------------------
hdr_col_left, hdr_col_center, hdr_col_right = st.columns(
    [0.15, 0.70, 0.15], gap="small", vertical_alignment="center"
)

with hdr_col_left:
    st.markdown('<div class="clean-nav-marker"></div>', unsafe_allow_html=True)

with hdr_col_center:
    if st.session_state["pc_active_page"] == "dashboard":
        st.markdown(
            clean_html(
                """
                <div class="premium-compact-header">
                    <h1 class="premium-compact-title">Minor Irrigation - Overdue Works Dashboard</h1>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            clean_html(
                """
                <div class="premium-compact-header">
                    <h1 class="premium-compact-title">Minor Irrigation — CE Analytics &amp; ETS Truth Audit</h1>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

with hdr_col_right:
    if st.session_state["pc_active_page"] == "dashboard":
        if st.button("📊 Analytics ➔", use_container_width=True, key="nav_to_analytics"):
            st.session_state["pc_active_page"] = "analytics"
            st.rerun()
    else:
        if st.button("⬅️ Dashboard", use_container_width=True, key="nav_to_dashboard"):
            st.session_state["pc_active_page"] = "dashboard"
            st.rerun()

# -----------------------------------------------------------------------------
# THREAD-SAFE DATA LOADING
# -----------------------------------------------------------------------------
shared_state = get_shared_state_v12()

if uploaded_xlsx is not None:
    file_id = f"{uploaded_xlsx.name}_{uploaded_xlsx.size}_{SCHEMA_VERSION}"
    if st.session_state.get("_last_uploaded_id") != file_id:
        with shared_state["lock"]:
            parsed = parse_xlsx_workbook(uploaded_xlsx.getvalue())
            save_shared_state(
                shared_state, parsed, f"Uploaded File ({uploaded_xlsx.name})"
            )
            st.session_state["_last_uploaded_id"] = file_id

elif (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or sync_clicked
):
    with shared_state["lock"]:
        if (
            shared_state["data"] is None
            or shared_state.get("version") != SCHEMA_VERSION
            or sync_clicked
        ):
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

st.markdown(
    clean_html(
        f"""
        <div class="mi-status-bar">
            <span class="mi-status-pill">⚡ {shared_state['source_label']}</span>
            <span>🕒 <b>Last Synced:</b> {shared_state['last_synced']} IST</span>
            <span>📌 <b>Row Rule:</b> Valid text in <b>Division</b> column (empty <b>Sub-division</b> rows included)</span>
        </div>
        """
    ),
    unsafe_allow_html=True,
)

if master_df.empty:
    st.warning("No matching data rows were found across the target worksheets.")
    st.dataframe(diag_df, use_container_width=True)
    st.stop()


# =============================================================================
# ANALYTICS HELPER FUNCTIONS (USED BY PAGE 2: CE ANALYTICS)
# =============================================================================
def build_officer_analytics_df(
    df_scope: pd.DataFrame, col_key: str, label_name: str
) -> pd.DataFrame:
    """Aggregates named officer performance including Overdue ('P'), Claimed 'ETS', and Unexcused Ratio."""
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
            primary_div=(
                "division_name",
                lambda s: s.mode().iloc[0] if not s.mode().empty else "-",
            ),
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


def render_executive_leaderboard_panel(
    df_slice: pd.DataFrame,
    officer_col: str,
    title_text: str,
    subtitle_text: str,
    is_danger: bool = True,
    primary_metric: str = "count",
) -> str:
    """Renders a Power BI-style progress-card leaderboard panel."""
    panel_cls = "exec-lb-panel-danger" if is_danger else "exec-lb-panel-success"
    pill_cls = "exec-pill-danger" if is_danger else "exec-pill-success"

    if df_slice.empty:
        return clean_html(
            f"""
            <div class="exec-lb-panel {panel_cls}">
                <div class="exec-lb-header">
                    <span class="exec-lb-title">{html.escape(title_text)}</span>
                    <span class="exec-lb-sub">{html.escape(subtitle_text)}</span>
                </div>
                <div style="color:#94A3B8;font-size:0.84rem;padding:12px 0;">
                    No officers match the current filter criteria.
                </div>
            </div>
            """
        )

    max_od = max(int(df_slice["overdue_works"].max()), 1)
    rows_html = []

    for rank_idx, (_, r) in enumerate(df_slice.iterrows(), start=1):
        off_name = html.escape(str(r[officer_col]))
        div_name = html.escape(str(r.get("primary_div", "-")))
        od_cnt = int(r["overdue_works"])
        od_pct = float(r["overdue_pct"])
        ets_cnt = int(r["ets_works"])
        tot_cnt = max(int(r["total_works"]), 1)

        if primary_metric == "pct":
            main_badge = f'<span class="{pill_cls}">{od_pct:.1f}% ({od_cnt} OD)</span>'
            od_bar_w = min(100.0, max(2.0 if od_cnt > 0 else 0.0, od_pct))
        else:
            main_badge = f'<span class="{pill_cls}">{od_cnt} OD ({od_pct:.1f}%)</span>'
            od_bar_w = (
                min(100.0, max(4.0, (od_cnt / max_od) * 100.0))
                if is_danger and od_cnt > 0
                else min(100.0, (od_cnt / tot_cnt) * 100.0)
            )

        ets_bar_w = min(100.0 - od_bar_w, (ets_cnt / tot_cnt) * 100.0)
        ets_badge = (
            f'<span class="exec-pill-ets">{ets_cnt} ETS</span>'
            if ets_cnt > 0
            else '<span style="color:#64748B;font-size:0.74rem;">0 ETS</span>'
        )

        rows_html.append(
            f"""
            <div class="exec-lb-row">
                <div class="exec-lb-row-top">
                    <div class="exec-lb-officer-wrap">
                        <span class="exec-lb-rank">#{rank_idx}</span>
                        <span class="exec-lb-name" title="{off_name}">{off_name}</span>
                        <span class="exec-lb-div-tag">🏛️ {div_name}</span>
                    </div>
                    <div class="exec-lb-badges">
                        {main_badge}
                        {ets_badge}
                        <span class="exec-pill-total">{tot_cnt} Total</span>
                    </div>
                </div>
                <div class="exec-lb-bar-bg">
                    <div class="exec-lb-bar-od" style="width:{od_bar_w:.1f}%;"></div>
                    <div class="exec-lb-bar-ets" style="width:{ets_bar_w:.1f}%;"></div>
                </div>
            </div>
            """
        )

    return clean_html(
        f"""
        <div class="exec-lb-panel {panel_cls}">
            <div class="exec-lb-header">
                <span class="exec-lb-title">{html.escape(title_text)}</span>
                <span class="exec-lb-sub">{html.escape(subtitle_text)}</span>
            </div>
            {''.join(rows_html)}
        </div>
        """
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
# PAGE 2 VIEW: CE EXECUTIVE ANALYTICS & ETS TRUTH AUDIT SUITE
# =============================================================================
if st.session_state["pc_active_page"] == "analytics":
    af_c1, af_c2, af_c3, af_c4, af_c5, af_c6 = st.columns(
        [1.1, 1.1, 1.1, 1.0, 1.0, 1.3], gap="small"
    )

    div_opts_an = [ALL_OPTION] + sorted(
        [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
    )
    with af_c1:
        with st.container(border=True):
            an_div = st.selectbox(
                f"🏛️ {LABEL_DIV_EN}:", options=div_opts_an, index=0, key="an_f_div"
            )

    an_s1 = (
        master_df[master_df["division_name"] == an_div].copy()
        if an_div != ALL_OPTION
        else master_df.copy()
    )

    subdiv_opts_an = build_dropdown_options(an_s1, "subdivision_name")
    with af_c2:
        with st.container(border=True):
            an_subdiv = st.selectbox(
                f"🏢 {LABEL_SUBDIV_EN}:",
                options=subdiv_opts_an,
                index=0,
                key="an_f_subdiv",
            )

    an_s2 = (
        an_s1[an_s1["subdivision_name"] == an_subdiv].copy()
        if an_subdiv != ALL_OPTION
        else an_s1.copy()
    )

    ee_opts_an = build_dropdown_options(an_s2, "ee_name")
    with af_c3:
        with st.container(border=True):
            an_ee = st.selectbox(
                f"👔 {LABEL_EE_EN}:", options=ee_opts_an, index=0, key="an_f_ee"
            )

    an_s3 = (
        an_s2[an_s2["ee_name"] == an_ee].copy()
        if an_ee != ALL_OPTION
        else an_s2.copy()
    )

    ta_opts_an = build_dropdown_options(an_s3, "ta_name")
    with af_c4:
        with st.container(border=True):
            an_ta = st.selectbox(
                f"🛠️ {LABEL_TA_EN}:", options=ta_opts_an, index=0, key="an_f_ta"
            )

    an_s4 = (
        an_s3[an_s3["ta_name"] == an_ta].copy()
        if an_ta != ALL_OPTION
        else an_s3.copy()
    )

    as_opts_an = build_dropdown_options(an_s4, "as_name")
    with af_c5:
        with st.container(border=True):
            an_as = st.selectbox(
                f"📑 {LABEL_AS_EN}:", options=as_opts_an, index=0, key="an_f_as"
            )

    an_scope_df = (
        an_s4[an_s4["as_name"] == an_as].copy()
        if an_as != ALL_OPTION
        else an_s4.copy()
    )

    with af_c6:
        with st.container(border=True):
            min_works_threshold = st.slider(
                "⚖️ Min. Works for Ranking:",
                min_value=1,
                max_value=50,
                value=5,
                help="Ensures officers with only 1 or 2 works do not skew the Percentage & Least Overdue rankings.",
            )

    if denominator_mode.startswith("Only Ongoing Works"):
        an_scope_df = an_scope_df[an_scope_df["is_status_p"]].copy()

    if an_scope_df.empty:
        st.warning("No works found matching the selected Analytics filter scope.")
        inject_scroll_hud_graphics(enabled=show_scroll_hud)
        st.stop()

    an_scope_label = (
        f"Division: {an_div} | Sub-division: {an_subdiv} | "
        f"EE: {an_ee} | TA: {an_ta} | AS: {an_as}"
    )

    ae_stats_df = build_officer_analytics_df(an_scope_df, "so_name", LABEL_AE_EN)
    aee_stats_df = build_officer_analytics_df(an_scope_df, "aee_name", LABEL_AEE_EN)

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
        clean_html(
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
            """
        ),
        unsafe_allow_html=True,
    )

    # =========================================================================
    # ANALYTICS SECTION 1: AE / AEE WISE ANALYTICS & CONNECTED AGEING TREE
    # =========================================================================
    with st.container(border=True):
        audit_role = st.radio(
            "👤 Select Officer Cadre:",
            options=[LABEL_AE_EN, LABEL_AEE_EN],
            horizontal=True,
            key="an_truth_role_radio",
        )
        active_audit_df = (
            ae_stats_df if audit_role == LABEL_AE_EN else aee_stats_df
        )

        st.markdown(
            clean_html(
                f"<h4 style='color:#F8FAFC; margin-bottom:0.5rem; margin-top:0.2rem;'>🔍 {html.escape(audit_role)} Wise Analytics</h4>"
            ),
            unsafe_allow_html=True,
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
                verdict_msg = (
                    f"<b>{html.escape(picked_off)}</b> has zero overdue works and "
                    f"zero ETS works across <b>{o_tot}</b> assigned works."
                )
            elif o_od > 0 and o_ets == 0:
                verdict_color = "#EF4444"
                verdict_title = "🚨 100% UNEXCUSED OVERDUE (ZERO ETS!)"
                verdict_msg = (
                    f"If <b>{html.escape(picked_off)}</b> claims their delayed works are in ETS, "
                    f"<b>that is false</b>: <b>{o_od} work(s)</b> are sitting in Overdue ('P') and <b>0 works</b> are in ETS."
                )
            elif o_od > o_ets:
                verdict_color = "#F97316"
                verdict_title = (
                    f"⚠️ MORE OVERDUE THAN ETS ({o_unexcused_pct:.1f}% Unexcused)"
                )
                verdict_msg = (
                    f"Out of <b>{o_od + o_ets}</b> delayed works under <b>{html.escape(picked_off)}</b>, "
                    f"only <b>{o_ets}</b> are in ETS while <b>{o_od} works ({o_unexcused_pct:.1f}%)</b> are sitting in Overdue ('P')."
                )
            else:
                verdict_color = "#F59E0B"
                verdict_title = f"📌 ETS HEAVY ({o_ets} in ETS vs. {o_od} Overdue)"
                verdict_msg = (
                    f"<b>{html.escape(picked_off)}</b> has placed <b>{o_ets}</b> works in ETS, "
                    f"and still has <b>{o_od}</b> active Overdue ('P') work(s) ({o_od_pct:.1f}% of total)."
                )

            # Calculate 6 Overdue Age Buckets for the Connected Root Tree
            target_col = "so_name" if audit_role == LABEL_AE_EN else "aee_name"
            off_od_df = od_rows_in_scope[
                od_rows_in_scope[target_col] == picked_off
            ].copy()

            od_u6 = od_6_12 = od_1_2 = od_2_3 = od_3_4 = od_4_plus = 0
            if not off_od_df.empty:
                def _custom_age_bucket(target_str):
                    parsed = parse_indian_date_list(pd.Series([target_str]))
                    t_dt = parsed[0] if parsed else None
                    if not isinstance(t_dt, date) or t_dt >= today_ist_obj:
                        return "Unknown"
                    days = (today_ist_obj - t_dt).days
                    if days <= 180:
                        return "Under 6 Months"
                    if days <= 365:
                        return "6 Months to 1 Year"
                    if days <= 730:
                        return "1-2 Years"
                    if days <= 1095:
                        return "2-3 Years"
                    if days <= 1460:
                        return "3-4 Years"
                    return "4+ Years"

                off_od_df["tree_bucket"] = off_od_df["target_date_raw"].map(
                    _custom_age_bucket
                )
                tree_counts = off_od_df["tree_bucket"].value_counts()

                od_u6 = int(tree_counts.get("Under 6 Months", 0))
                od_6_12 = int(tree_counts.get("6 Months to 1 Year", 0))
                od_1_2 = int(tree_counts.get("1-2 Years", 0))
                od_2_3 = int(tree_counts.get("2-3 Years", 0))
                od_3_4 = int(tree_counts.get("3-4 Years", 0))
                od_4_plus = int(tree_counts.get("4+ Years", 0))

            tree_html = f"""
            <div style="position:relative; width:100%; margin-top:0px; padding-top:26px; z-index:1;">
                <div style="position:absolute; top:0; left:12.2%; width:2.5px; height:14px; background:#EF4444;"></div>
                <div style="position:absolute; top:14px; left:7.8%; right:7.8%; height:2.5px; background:#EF4444;"></div>
                <div style="display:grid; grid-template-columns:repeat(6, 1fr); gap:10px; width:100%;">
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">Under 6 Months</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_u6}</div>
                        </div>
                    </div>
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">6 Months – 1 Yr</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_6_12}</div>
                        </div>
                    </div>
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">1 – 2 Years</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_1_2}</div>
                        </div>
                    </div>
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">2 – 3 Years</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_2_3}</div>
                        </div>
                    </div>
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">3 – 4 Years</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_3_4}</div>
                        </div>
                    </div>
                    <div style="text-align:center; position:relative;">
                        <div style="position:absolute; top:-12px; left:50%; transform:translateX(-50%); width:2.5px; height:12px; background:#EF4444;"></div>
                        <div style="background:rgba(127,29,29,0.28); border:1.5px solid rgba(239,68,68,0.55); border-radius:8px; padding:9px 6px;">
                            <div style="font-size:0.68rem; color:#FCA5A5; font-weight:700; text-transform:uppercase;">4+ Years</div>
                            <div style="font-size:1.3rem; font-weight:800; color:#FFFFFF; margin-top:2px;">{od_4_plus}</div>
                        </div>
                    </div>
                </div>
            </div>
            """

            st.markdown(
                clean_html(
                    f"""
                    <div class="truth-audit-box" style="border-left-color: {verdict_color};">
                        <div style="font-size:0.95rem;font-weight:800;color:{verdict_color};margin-bottom:6px;">
                            {verdict_title}
                        </div>
                        <div style="font-size:0.86rem;color:#E2E8F0;line-height:1.45;margin-bottom:14px;">
                            {verdict_msg}
                        </div>
                        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;text-align:center;position:relative;z-index:2;">
                            <div style="background:rgba(239,68,68,0.18);border:1.5px solid rgba(239,68,68,0.65);border-radius:8px;padding:12px;">
                                <div style="font-size:0.75rem;color:#FCA5A5;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">Overdue ('P')</div>
                                <div style="font-size:1.55rem;font-weight:800;color:#FCA5A5;">{o_od}</div>
                            </div>
                            <div style="background:rgba(16,185,129,0.16);border:1px solid rgba(16,185,129,0.4);border-radius:8px;padding:12px;">
                                <div style="font-size:0.75rem;color:#6EE7B7;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">In 'ETS'</div>
                                <div style="font-size:1.55rem;font-weight:800;color:#6EE7B7;">{o_ets}</div>
                            </div>
                            <div style="background:rgba(59,130,246,0.16);border:1px solid rgba(59,130,246,0.4);border-radius:8px;padding:12px;">
                                <div style="font-size:0.75rem;color:#93C5FD;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">On-Track ('P')</div>
                                <div style="font-size:1.55rem;font-weight:800;color:#93C5FD;">{o_ontrack}</div>
                            </div>
                            <div style="background:rgba(148,163,184,0.16);border:1px solid rgba(148,163,184,0.3);border-radius:8px;padding:12px;">
                                <div style="font-size:0.75rem;color:#CBD5E1;text-transform:uppercase;font-weight:700;letter-spacing:0.5px;">Total Works</div>
                                <div style="font-size:1.55rem;font-weight:800;color:#F8FAFC;">{o_tot}</div>
                            </div>
                        </div>
                        {tree_html}
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

            # --- Below the Tree: 3D Hover-Grow Donut (Overdue Combined) & Single Table ---
            fc_c1, fc_c2 = st.columns([1.05, 1.45], gap="large")

            with fc_c1:
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
                            hole=0.56,
                            pull=[0.06 if o_od > 0 else 0, 0, 0, 0],
                            hoverinfo="label+percent+value",
                            marker=dict(
                                colors=["#EF4444", "#10B981", "#3B82F6", "#475569"],
                                line=dict(color="#0F172A", width=3),
                            ),
                            textinfo="label+value",
                            textfont=dict(size=12, color="#F8FAFC"),
                        )
                    ]
                )
                fig_off_donut.update_traces(
                    hoverlabel=dict(
                        bgcolor="#1E293B", font_size=13, font_color="#F8FAFC"
                    )
                )
                fig_off_donut.update_layout(
                    height=300,
                    margin=dict(l=10, r=10, t=12, b=12),
                    paper_bgcolor="rgba(0,0,0,0)",
                    showlegend=False,
                    annotations=[
                        dict(
                            text=f"<b>{o_tot}</b><br>Works",
                            x=0.5,
                            y=0.5,
                            font_size=16,
                            font_color="#F8FAFC",
                            showarrow=False,
                        )
                    ],
                )
                st.plotly_chart(fig_off_donut, use_container_width=True)

            with fc_c2:
                off_od_col = "_so_name" if audit_role == LABEL_AE_EN else "_aee_name"
                proof_df = pd.DataFrame()
                if not overdue_full_df.empty:
                    proof_df = overdue_full_df[
                        overdue_full_df[off_od_col] == picked_off
                    ].copy()
                    if an_div != ALL_OPTION:
                        proof_df = proof_df[proof_df["_division_name"] == an_div]
                    if an_subdiv != ALL_OPTION:
                        proof_df = proof_df[proof_df["_subdivision_name"] == an_subdiv]
                    if an_ee != ALL_OPTION:
                        proof_df = proof_df[proof_df["_ee_name"] == an_ee]
                    if an_ta != ALL_OPTION:
                        proof_df = proof_df[proof_df["_ta_name"] == an_ta]
                    if an_as != ALL_OPTION:
                        proof_df = proof_df[proof_df["_as_name"] == an_as]

                if not proof_df.empty:
                    st.markdown(
                        f"**📂 View {len(proof_df)} Actual Overdue Work(s) of {html.escape(picked_off)}**"
                    )
                    st.dataframe(
                        proof_df[
                            [
                                COL_WORK_NAME_KN,
                                COL_TARGET_DATE_KN,
                                "Overdue By",
                                COL_HOA_KN,
                            ]
                        ].rename(
                            columns={
                                COL_WORK_NAME_KN: "Work Name",
                                COL_TARGET_DATE_KN: "Target Date",
                                COL_HOA_KN: "H.O.A.",
                            }
                        ),
                        use_container_width=True,
                        hide_index=True,
                        height=265,
                    )
                else:
                    st.info(
                        f"🎉 {picked_off} has 0 actual Overdue ('P') works in the selected scope."
                    )

    # =========================================================================
    # ANALYTICS SECTION 2: REDESIGNED SPACIOUS AE & AEE LEADERBOARDS
    # =========================================================================
    lb_top_ctrl1, lb_top_ctrl2 = st.columns([0.75, 0.25], vertical_alignment="center")
    with lb_top_ctrl1:
        st.markdown(
            clean_html(
                f"""
                <div class="exec-section-banner">
                    <div class="exec-section-title">🏆 Executive AE &amp; AEE Leaderboards — Highest &amp; Least Overdue</div>
                    <div class="exec-scope-badge">Min. {min_works_threshold} assigned works required for Percentage &amp; Least-Overdue rankings</div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
    with lb_top_ctrl2:
        top_n_count = st.radio(
            "Leaderboard Depth:",
            options=[5, 10],
            format_func=lambda x: f"Top {x}",
            horizontal=True,
            index=0,
            key="lb_depth_radio",
        )

    ae_qual = ae_stats_df[ae_stats_df["total_works"] >= min_works_threshold].copy()
    if ae_qual.empty:
        ae_qual = ae_stats_df.copy()

    ae_high_cnt = ae_stats_df.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"],
        ascending=[False, False, False],
    ).head(top_n_count)
    ae_high_pct = ae_qual.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"],
        ascending=[False, False, False],
    ).head(top_n_count)
    ae_least_cnt = ae_qual.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"],
        ascending=[True, True, False],
    ).head(top_n_count)
    ae_least_pct = ae_qual.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"],
        ascending=[True, True, False],
    ).head(top_n_count)

    aee_qual = aee_stats_df[aee_stats_df["total_works"] >= min_works_threshold].copy()
    if aee_qual.empty:
        aee_qual = aee_stats_df.copy()

    aee_high_cnt = aee_stats_df.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"],
        ascending=[False, False, False],
    ).head(top_n_count)
    aee_high_pct = aee_qual.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"],
        ascending=[False, False, False],
    ).head(top_n_count)
    aee_least_cnt = aee_qual.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"],
        ascending=[True, True, False],
    ).head(top_n_count)
    aee_least_pct = aee_qual.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"],
        ascending=[True, True, False],
    ).head(top_n_count)

    tab_ae_lb, tab_aee_lb = st.tabs(
        [
            f"👷 Assistant Engineer (AE) — Top {top_n_count} Leaderboards",
            f"👔 Assistant Executive Engineer (AEE) — Top {top_n_count} Leaderboards",
        ]
    )

    with tab_ae_lb:
        ae_r1_c1, ae_r1_c2 = st.columns(2, gap="medium")
        with ae_r1_c1:
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_high_cnt,
                    LABEL_AE_EN,
                    f"🔴 Top {top_n_count} AEs — Highest Overdue (By Count)",
                    "Ranked by Overdue Works Count",
                    is_danger=True,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
        with ae_r1_c2:
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_high_pct,
                    LABEL_AE_EN,
                    f"🔴 Top {top_n_count} AEs — Highest Overdue (By Percentage %)",
                    f"Min. {min_works_threshold} Works Assigned",
                    is_danger=True,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

        ae_r2_c1, ae_r2_c2 = st.columns(2, gap="medium")
        with ae_r2_c1:
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_least_cnt,
                    LABEL_AE_EN,
                    f"🟢 Top {top_n_count} AEs — Least Overdue (By Count)",
                    f"Min. {min_works_threshold} Works • Ties broken by highest workload",
                    is_danger=False,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
        with ae_r2_c2:
            st.markdown(
                render_executive_leaderboard_panel(
                    ae_least_pct,
                    LABEL_AE_EN,
                    f"🟢 Top {top_n_count} AEs — Least Overdue (By Percentage %)",
                    f"Min. {min_works_threshold} Works • Ties broken by highest workload",
                    is_danger=False,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

    with tab_aee_lb:
        aee_r1_c1, aee_r1_c2 = st.columns(2, gap="medium")
        with aee_r1_c1:
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_high_cnt,
                    LABEL_AEE_EN,
                    f"🔴 Top {top_n_count} AEEs — Highest Overdue (By Count)",
                    "Ranked by Overdue Works Count",
                    is_danger=True,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
        with aee_r1_c2:
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_high_pct,
                    LABEL_AEE_EN,
                    f"🔴 Top {top_n_count} AEEs — Highest Overdue (By Percentage %)",
                    f"Min. {min_works_threshold} Works Assigned",
                    is_danger=True,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

        aee_r2_c1, aee_r2_c2 = st.columns(2, gap="medium")
        with aee_r2_c1:
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_least_cnt,
                    LABEL_AEE_EN,
                    f"🟢 Top {top_n_count} AEEs — Least Overdue (By Count)",
                    f"Min. {min_works_threshold} Works • Ties broken by highest workload",
                    is_danger=False,
                    primary_metric="count",
                ),
                unsafe_allow_html=True,
            )
        with aee_r2_c2:
            st.markdown(
                render_executive_leaderboard_panel(
                    aee_least_pct,
                    LABEL_AEE_EN,
                    f"🟢 Top {top_n_count} AEEs — Least Overdue (By Percentage %)",
                    f"Min. {min_works_threshold} Works • Ties broken by highest workload",
                    is_danger=False,
                    primary_metric="pct",
                ),
                unsafe_allow_html=True,
            )

    # =========================================================================
    # ANALYTICS SECTION 3: POWER BI DONUT CHARTS, AGEING & DIVISION BREAKDOWN
    # =========================================================================
    st.markdown(
        clean_html(
            """
            <div class="exec-section-banner">
                <div class="exec-section-title">📊 Portfolio Visuals: Work Status Donut, Delay Ageing Donut &amp; Division Comparison</div>
                <div class="exec-scope-badge">Power BI Interactive Breakdown</div>
            </div>
            """
        ),
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
                        pull=[0.05 if tot_od_an > 0 else 0, 0, 0, 0],
                        marker=dict(
                            colors=["#EF4444", "#10B981", "#3B82F6", "#475569"],
                            line=dict(color="#0F172A", width=2.5),
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
                height=350,
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
                st.info("No overdue works in the current scope to compute ageing.")
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
                                ],
                                line=dict(color="#0F172A", width=2.5),
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
                    height=350,
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
                    name="Claimed 'ETS'",
                    orientation="h",
                    marker_color="#10B981",
                    text=div_grp["ets_works"],
                    textposition="outside",
                )
            )
            fig_div_bar.update_layout(
                barmode="group",
                title=dict(
                    text="🏛️ Division-Wise Overdue ('P') vs. Claimed 'ETS'",
                    font=dict(size=14, color="#F8FAFC"),
                ),
                height=350,
                margin=dict(l=10, r=35, t=40, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(15,23,42,0.45)",
                font=dict(color="#E2E8F0", size=11),
                legend=dict(orientation="h", y=1.02, x=0),
                xaxis=dict(gridcolor="rgba(148,163,184,0.12)"),
            )
            st.plotly_chart(fig_div_bar, use_container_width=True)

    # --- Row 2 of Visuals: Scheme (H.O.A.) Bottlenecks & Officer Risk Matrix ---
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
                    name="Claimed 'ETS'",
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
            scatter_df = ae_qual.copy()
            if not scatter_df.empty:
                med_x = float(scatter_df["total_works"].median())
                med_y = float(scatter_df["overdue_pct"].median())
                top6_cut = (
                    scatter_df["overdue_works"].nlargest(6).min()
                    if len(scatter_df) >= 6
                    else 0
                )

                fig_quad = go.Figure()
                fig_quad.add_trace(
                    go.Scatter(
                        x=scatter_df["total_works"],
                        y=scatter_df["overdue_pct"],
                        mode="markers+text",
                        text=[
                            r[LABEL_AE_EN]
                            if r["overdue_works"] >= top6_cut
                            and r["overdue_works"] > 0
                            else ""
                            for _, r in scatter_df.iterrows()
                        ],
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

    # =========================================================================
    # ANALYTICS SECTION 4: COMPLETE OFFICER TRUTH AUDIT TABLE & EXCEL DOWNLOAD
    # =========================================================================
    st.markdown(
        clean_html(
            """
            <div class="exec-section-banner">
                <div class="exec-section-title">📋 Complete Officer Overdue vs. ETS Truth Audit Register</div>
                <div class="exec-scope-badge">Flags officers with unexcused Overdue ('P') works vs. Claimed 'ETS'</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    audit_table_df = (
        ae_stats_df.copy() if audit_role == LABEL_AE_EN else aee_stats_df.copy()
    )
    if not audit_table_df.empty:
        def _build_audit_status(r):
            od = int(r["overdue_works"])
            ets = int(r["ets_works"])
            if od == 0 and ets == 0:
                return "✅ Clean Record (0 Overdue, 0 ETS)"
            if od > 0 and ets == 0:
                return "🚨 100% Unexcused Overdue (0 in ETS)"
            if od > ets:
                return "⚠️ High Unexcused Overdue (Overdue > ETS)"
            if od > 0 and ets >= od:
                return "📌 Mixed (Both Overdue & ETS Active)"
            return "🟢 Only ETS Active (0 Overdue 'P')"

        audit_table_df["CE Audit Verdict"] = audit_table_df.apply(
            _build_audit_status, axis=1
        )
        audit_table_df = audit_table_df.sort_values(
            by=["overdue_works", "unexcused_ratio_pct", "ets_works"],
            ascending=[False, False, False],
        ).reset_index(drop=True)
        audit_table_df.index = audit_table_df.index + 1

        display_audit_cols = audit_table_df[
            [
                audit_role,
                "primary_div",
                "total_works",
                "overdue_works",
                "overdue_pct",
                "ets_works",
                "delayed_exposure",
                "unexcused_ratio_pct",
                "CE Audit Verdict",
            ]
        ].rename(
            columns={
                "primary_div": "Primary Division",
                "total_works": "Total Works",
                "overdue_works": "Actual Overdue ('P')",
                "overdue_pct": "Overdue %",
                "ets_works": "Claimed 'ETS'",
                "delayed_exposure": "Total Delayed (Overdue+ETS)",
                "unexcused_ratio_pct": "Unexcused Ratio %",
            }
        )

        dl_a1, dl_a2, _ = st.columns([1.6, 1.6, 4.8])
        with dl_a1:
            st.download_button(
                "📥 Download Truth Audit (.xlsx)",
                data=convert_df_to_excel_bytes(display_audit_cols),
                file_name=f"CE_Truth_Audit_{audit_role.replace(' ', '_')}_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_a2:
            st.download_button(
                "📥 Download Truth Audit (.csv)",
                data=display_audit_cols.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"CE_Truth_Audit_{audit_role.replace(' ', '_')}_{today_str}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.dataframe(
            display_audit_cols.style.format(
                {
                    "Overdue %": "{:.1f}%",
                    "Unexcused Ratio %": "{:.1f}%",
                }
            ),
            use_container_width=True,
        )

    inject_scroll_hud_graphics(enabled=show_scroll_hud)
    st.stop()


# =============================================================================
# PAGE 1 VIEW: MAIN OVERDUE WORKS DASHBOARD (EXISTING PC VERSION)
# =============================================================================
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

st.markdown("<div style='margin-top: 0.4rem;'></div>", unsafe_allow_html=True)

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
# KPI SUMMARY METRIC CARDS (5 COLUMNS WITH DISTINCT COLOR TINTS)
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

active_scope_parts = [
    f"{LABEL_DIV_EN}: {selected_division}",
    f"{LABEL_SUBDIV_EN}: {selected_subdivision}",
    f"EE: {selected_ee}",
    f"TA: {selected_ta}",
    f"AS: {selected_as}",
]
filter_scope_label = " | ".join(active_scope_parts)

k1, k2, k3, k4, k5 = st.columns(5, gap="medium")
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
    clean_html(
        f"""
        <div class="exec-section-banner">
            <div class="exec-section-title">📈 Overdue &amp; ETS Works by {html.escape(view_choice)} ({html.escape(sort_choice)})</div>
            <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
        </div>
        """
    ),
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
                textfont=dict(size=13),
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
                textfont=dict(size=13),
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
                    color="rgba(148, 163, 184, 0.30)",
                    line=dict(color="#64748B", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13),
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
                marker=dict(color="#4ADE80"),
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
        plot_bgcolor="rgba(15, 23, 42, 0.35)",
        xaxis=dict(
            title=x_axis_title,
            range=[0, x_range_max],
            gridcolor="rgba(148, 163, 184, 0.14)",
            zeroline=True,
            zerolinecolor="rgba(148, 163, 184, 0.3)",
        ),
        yaxis=dict(
            title="",
            automargin=False,
            tickmode="array",
            tickvals=plot_df[view_choice],
            ticktext=plot_df["y_short"],
            tickfont=dict(size=13),
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="left",
            x=0,
            traceorder="reversed",
        ),
        margin=dict(l=210, r=14, t=35, b=40),
    )

    st.plotly_chart(fig, use_container_width=True)

st.divider()

# =============================================================================
# SECTION 2: COMBINED OVERDUE WORKS TABLE (ENGLISH HEADERS)
# =============================================================================
st.markdown(
    clean_html(
        f"""
        <div class="exec-section-banner">
            <div class="exec-section-title">📂 Combined Overdue Works by {html.escape(view_choice)}</div>
            <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
        </div>
        """
    ),
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
                f"**Selected {view_choice}:** `{selected_officer}` &nbsp;|&nbsp; "
                f"**Total Overdue Works:** `{len(officer_overdue_df)}` &nbsp;|&nbsp; "
                f"**Account Heads (H.O.A.):** `{', '.join(hoa_involved)}`"
            )

            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Unassigned"
            )
            safe_div = (
                re.sub(r"[^\w\-]+", "_", selected_division).strip("_")
                or "All_Divisions"
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

            st.markdown(
                render_ce_overdue_html_table(officer_overdue_df),
                unsafe_allow_html=True,
            )

st.divider()

# =============================================================================
# SECTION 3: OFFICER-WISE SUMMARY TABLE (ENGLISH)
# =============================================================================
st.markdown(
    clean_html(
        f"""
        <div class="exec-section-banner">
            <div class="exec-section-title">📋 Officer-Wise Summary Table</div>
            <div class="exec-scope-badge">{html.escape(filter_scope_label)}</div>
        </div>
        """
    ),
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

inject_scroll_hud_graphics(enabled=show_scroll_hud)

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
