import sys
import io as io_module

# Windows-Konsolencodierung auf UTF-8 setzen
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except (AttributeError, io_module.UnsupportedOperation):
        # Fallback für ältere Python-Versionen oder nicht unterstützte Streams
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'ignore')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'ignore')

import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
import io
import os
import json
from pathlib import Path

# Debug-Ausgaben verbergen (durch Entfernung der print-Statements und detaillierten UI-Ausgaben)

# Importieren der Module
from data_import import (
    preprocess_data,
    scale_consumption_by_persons,
    scale_pv_generation,
    preprocess_data_with_standard_profile,
    load_pv_generation_from_excel,
    load_consumption_from_excel,
    load_pv_generation_from_csv,
    load_battery_cost_curve,
    load_battery_tech_params,
    get_battery_cost,
)
from model import simulate_one_year
from scenarios import find_optimal_size, run_variable_tariff_scenario
from analysis import (
    calculate_financial_kpis,
    plot_energy_flows_for_period,
    plot_sankey_diagram,
    plot_sankey_diagram_no_battery,
    plot_cost_comparison_with_without_battery,
    plot_data_control,
    plot_technical_optimization_curve,
    plot_economic_optimization_curve,
    plot_energy_flows_optimization,
    plot_economic_optimization_extended,
    compute_energy_axis_range,
)
from config import *
from settings_manager import SettingsManager
from excel_export import excel_export_section

PV_LINK_FILE = Path("pv_link.json")

def load_pv_link_from_file() -> str:
    if PV_LINK_FILE.exists():
        try:
            with PV_LINK_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("pv_link", "")
        except Exception:
            return ""
    return ""

def save_pv_link_to_file(link: str):
    try:
        with PV_LINK_FILE.open("w", encoding="utf-8") as f:
            json.dump({"pv_link": link}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

if 'pv_data_source_link_input' not in st.session_state:
    st.session_state['pv_data_source_link_input'] = load_pv_link_from_file()

DEFAULT_CONSUMPTION_FILE = Path("Verbrauchsdaten.xlsx")

# Excel Export wird über excel_export.py bereitgestellt

# Custom CSS für Orange-Schwarz Farbschema
st.markdown("""
<style>
    /* Hauptfarben */
    :root {
        --primary-orange: #FF6B35;
        --secondary-orange: #FF8C42;
        --light-orange: #FFB366;
        --dark-orange: #E55A2B;
        --primary-black: #1A1A1A;
        --secondary-black: #2D2D2D;
        --light-black: #404040;
        --white: #FFFFFF;
        --light-gray: #F5F5F5;
        --medium-gray: #E0E0E0;
        --dark-gray: #666666;
    }

    /* Hauptcontainer */
    .stApp {
        background-color: var(--primary-black) !important;
    }
    
    .main {
        background-color: var(--primary-black);
        color: var(--white);
    }

    /* Horizontalen Overflow verhindern, damit Inhalte (inkl. Menü) nie aus dem Sichtfeld rutschen */
    html, body,
    [data-testid="stAppViewContainer"],
    .main,
    .block-container {
        overflow-x: hidden !important;
    }

    /* Sidebar Styling */
    .css-1d391kg {
        background-color: var(--primary-black) !important;
    }
    .css-1d391kg .sidebar-content {
        background-color: var(--primary-black) !important;
        color: var(--white);
    }
    [data-testid="stSidebar"] {
        background-color: var(--primary-black) !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        background-color: var(--primary-black) !important;
    }
    section[data-testid="stSidebar"] {
        background-color: var(--primary-black) !important;
    }
    section[data-testid="stSidebar"] > div {
        background-color: var(--primary-black) !important;
    }

    /* Plotly-Charts auf Containerbreite begrenzen, auch nach Fullscreen/Resize */
    .js-plotly-plot,
    .js-plotly-plot .main-svg {
        max-width: 100% !important;
    }


    /* Headers */
    h1, h2, h3, h4, h5, h6 {
        color: var(--primary-orange) !important;
        font-weight: 600;
    }

    /* Buttons */
    .stButton > button {
        background-color: #FF8C42 !important;
        color: var(--white) !important;
        border: 2px solid #FF6B35 !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 2px 4px rgba(255, 107, 53, 0.2) !important;
    }

    .stButton > button:hover {
        background-color: #FF6B35 !important;
        border-color: #E55A2B !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 12px rgba(255, 107, 53, 0.4) !important;
    }

    /* Success Buttons */
    .stButton > button[data-baseweb="button"]:has-text("✅") {
        background-color: #28a745 !important;
    }

    .stButton > button[data-baseweb="button"]:has-text("✅"):hover {
        background-color: #218838 !important;
    }

    /* Danger Buttons */
    .stButton > button[data-baseweb="button"]:has-text("🗑️") {
        background-color: #dc3545 !important;
    }

    .stButton > button[data-baseweb="button"]:has-text("🗑️"):hover {
        background-color: #c82333 !important;
    }

    /* Input Fields */
    .stTextInput > div > div > input {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stTextInput"] input {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stTextInput"] input:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }
    [data-testid="stTextInput"] input::placeholder {
        color: var(--dark-gray) !important;
    }

    .stTextInput > div > div > input:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }

    /* Number Inputs */
    .stNumberInput > div > div > input {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stNumberInput"] input {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stNumberInput"] input:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }

    .stNumberInput > div > div > input:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }

    /* Sliders */
    .stSlider > div > div > div > div {
        background-color: var(--primary-orange) !important;
    }

    .stSlider > div > div > div > div > div {
        background-color: var(--primary-orange) !important;
    }

    /* Selectboxes */
    .stSelectbox > div > div > div {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stSelectbox"] > div > div > div {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stSelectbox"] select {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
        border-radius: 6px !important;
    }
    [data-testid="stSelectbox"] > div > div > div:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }

    .stSelectbox > div > div > div:focus {
        border-color: var(--primary-orange) !important;
        box-shadow: 0 0 0 2px rgba(255, 107, 53, 0.2) !important;
    }

    /* File Uploader */
    .stFileUploader > div > div {
        background-color: var(--light-black) !important;
        border: 2px dashed var(--primary-orange) !important;
        border-radius: 8px !important;
        color: var(--white) !important;
    }
    [data-testid="stFileUploader"] {
        background-color: var(--light-black) !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: var(--light-black) !important;
        border: 2px dashed var(--primary-orange) !important;
        color: var(--white) !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
    }
    [data-testid="stFileUploader"] .uploadedFile {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
        border: 1px solid var(--primary-orange) !important;
    }

    /* Progress Bar */
    .stProgress > div > div > div > div {
        background-color: var(--primary-orange) !important;
    }

    /* Metrics */
    .metric-container {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
        border-radius: 8px !important;
        padding: 16px !important;
        margin: 8px 0 !important;
    }

    .metric-label {
        color: var(--light-orange) !important;
        font-size: 0.9em !important;
        font-weight: 500 !important;
    }

    .metric-value {
        color: var(--white) !important;
        font-size: 1.5em !important;
        font-weight: 700 !important;
    }

    .metric-delta {
        color: var(--primary-orange) !important;
        font-size: 0.8em !important;
    }

    /* Success/Error Messages */
    .stSuccess {
        background-color: rgba(40, 167, 69, 0.1) !important;
        border: 1px solid #28a745 !important;
        border-radius: 6px !important;
        color: #28a745 !important;
    }

    .stError {
        background-color: rgba(220, 53, 69, 0.1) !important;
        border: 1px solid #dc3545 !important;
        border-radius: 6px !important;
        color: #dc3545 !important;
    }

    .stWarning {
        background-color: rgba(255, 193, 7, 0.1) !important;
        border: 1px solid #ffc107 !important;
        border-radius: 6px !important;
        color: #ffc107 !important;
    }

    .stInfo {
        background-color: rgba(23, 162, 184, 0.1) !important;
        border: 1px solid #17a2b8 !important;
        border-radius: 6px !important;
        color: #17a2b8 !important;
    }

    /* DataFrames */
    .dataframe {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
        border: 1px solid var(--light-black) !important;
    }

    .dataframe th {
        background-color: var(--primary-orange) !important;
        color: var(--white) !important;
    }

    .dataframe td {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
        border: 1px solid var(--light-black) !important;
    }

    /* Cards */
    .card {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
        border-radius: 8px !important;
        padding: 16px !important;
        margin: 8px 0 !important;
    }

    /* Custom Divider */
    .custom-divider {
        border-top: 2px solid var(--primary-orange) !important;
        margin: 20px 0 !important;
        opacity: 0.7 !important;
    }

    /* Sidebar Headers */
    .sidebar .sidebar-content h3 {
        color: var(--primary-orange) !important;
        border-bottom: 2px solid var(--primary-orange) !important;
        padding-bottom: 8px !important;
        margin-bottom: 16px !important;
    }

    /* Sidebar Buttons - Extra sichtbar - Höchste Spezifität */
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] .stButton > button:not(:hover),
    [data-testid="stSidebar"] .stButton > button:not(:focus),
    .sidebar .stButton > button,
    .sidebar .stButton > button:not(:hover),
    .sidebar .stButton > button:not(:focus) {
        background-color: #FF8C42 !important;
        color: #FFFFFF !important;
        border: 2px solid #FF6B35 !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 3px 6px rgba(255, 107, 53, 0.3) !important;
        text-shadow: none !important;
    }

    .sidebar .stButton > button:hover {
        background-color: #FF6B35 !important;
        color: #FFFFFF !important;
        border-color: #E55A2B !important;
        transform: translateY(-3px) !important;
        box-shadow: 0 8px 16px rgba(255, 107, 53, 0.5) !important;
    }

    /* Main Content Background */
    .block-container {
        background-color: var(--primary-black) !important;
        color: var(--white) !important;
        padding-top: 0.8rem !important;
    }
    [data-testid="stAppViewContainer"] {
        background-color: var(--primary-black) !important;
    }
    [data-testid="stHeader"] {
        background-color: var(--primary-black) !important;
    }
    .main .block-container {
        background-color: var(--primary-black) !important;
    }

    /* Text Colors */
    p, div, span {
        color: var(--white) !important;
    }

    /* Links */
    a {
        color: var(--primary-orange) !important;
    }

    a:hover {
        color: var(--light-orange) !important;
    }

    /* Code Blocks */
    code {
        background-color: var(--secondary-black) !important;
        color: var(--light-orange) !important;
        border: 1px solid var(--light-black) !important;
        border-radius: 4px !important;
        padding: 2px 4px !important;
    }

    /* JSON Display */
    .json {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
        border: 1px solid var(--light-black) !important;
        border-radius: 6px !important;
        padding: 12px !important;
    }

    /* Widget Labels */
    label {
        color: var(--white) !important;
    }
    [data-testid="stTextInput"] > label,
    [data-testid="stNumberInput"] > label,
    [data-testid="stSelectbox"] > label,
    [data-testid="stFileUploader"] > label {
        color: var(--white) !important;
    }
    
    /* Input-Container */
    [data-testid="stTextInput"] > label,
    [data-testid="stNumberInput"] > label,
    [data-testid="stSelectbox"] > label {
        color: var(--white) !important;
    }
    
    /* Dropdown-Menüs (aufgeklappt) */
    [data-baseweb="select"] > div {
        background-color: var(--light-black) !important;
        color: var(--white) !important;
        border: 2px solid var(--medium-gray) !important;
    }
    [data-baseweb="select"] > div:hover {
        border-color: var(--primary-orange) !important;
    }
    [data-baseweb="popover"] {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
    }
    [data-baseweb="menu"] {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
    }
    [role="listbox"] {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
    }
    [role="option"] {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
    }
    [role="option"]:hover {
        background-color: var(--light-black) !important;
        color: var(--primary-orange) !important;
    }
    ul[role="listbox"] {
        background-color: var(--secondary-black) !important;
        border: 1px solid var(--light-black) !important;
    }
    li[role="option"] {
        background-color: var(--secondary-black) !important;
        color: var(--white) !important;
    }
    li[role="option"]:hover {
        background-color: var(--light-black) !important;
        color: var(--primary-orange) !important;
    }
    
    /* Plotly Container - Text sichtbar machen */
    .stPlotlyChart {
        background-color: var(--secondary-black) !important;
    }
    
    .stPlotlyChart * {
        color: var(--white) !important;
    }
    
    /* Plotly Toolbar Buttons - Nur Symbole einfärben, kein Hintergrund */
    .plotly .modebar-btn {
        background-color: transparent !important;
        border: none !important;
        color: #FFFFFF !important;
    }
    
    .plotly .modebar-btn:hover {
        background-color: rgba(255, 255, 255, 0.1) !important;
        color: #FF8C42 !important;
    }
    
    .plotly .modebar-btn svg {
        fill: #FFFFFF !important;
    }
    
    .plotly .modebar-btn:hover svg {
        fill: #FF8C42 !important;
    }
    
    /* Plotly Toolbar Container */
    .plotly .modebar {
        background-color: transparent !important;
        border: none !important;
    }
    
    /* Plotly Fullscreen-Button ausblenden */
    .plotly .modebar-btn[data-title*="fullscreen"],
    .plotly .modebar-btn[data-title*="Full screen"],
    .plotly .modebar-btn[data-title*="Vollbild"],
    .plotly .modebar-btn[aria-label*="fullscreen"],
    .plotly .modebar-btn[aria-label*="Full screen"],
    .plotly .modebar-btn[aria-label*="Vollbild"],
    .plotly .modebar-btn[data-title*="Fullscreen"],
    .plotly .modebar-btn[aria-label*="Fullscreen"],
    .plotly .modebar .modebar-btn:has(svg[data-title*="fullscreen"]),
    .plotly .modebar .modebar-btn:has(svg[data-title*="Full screen"]),
    .plotly .modebar .modebar-btn:has(svg[data-title*="Fullscreen"]) {
        display: none !important;
    }
    
    /* Allgemeine Text-Elemente - Sichtbarkeit sicherstellen */
    .stMarkdown, .stText, .stCaption, .stAlert {
        color: var(--white) !important;
        background-color: transparent !important;
    }
    
    /* Main Content Buttons - Spezifisch für Hauptprogramm */
    .main .stButton > button,
    .main .stButton > button:not(:hover),
    .main .stButton > button:not(:focus),
    .block-container .stButton > button,
    .block-container .stButton > button:not(:hover),
    .block-container .stButton > button:not(:focus),
    div[data-testid="stAppViewContainer"] .stButton > button,
    div[data-testid="stAppViewContainer"] .stButton > button:not(:hover),
    div[data-testid="stAppViewContainer"] .stButton > button:not(:focus) {
        background-color: #FF8C42 !important;
        color: #FFFFFF !important;
        border: 2px solid #FF6B35 !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 2px 4px rgba(255, 107, 53, 0.2) !important;
    }
    
    /* Main Content Text Elements - Spezifisch für Hauptprogramm */
    .main .stMarkdown,
    .main .stText,
    .main .stCaption,
    .main .stAlert,
    .block-container .stMarkdown,
    .block-container .stText,
    .block-container .stCaption,
    .block-container .stAlert,
    div[data-testid="stAppViewContainer"] .stMarkdown,
    div[data-testid="stAppViewContainer"] .stText,
    div[data-testid="stAppViewContainer"] .stCaption,
    div[data-testid="stAppViewContainer"] .stAlert {
        color: #FFFFFF !important;
        background-color: transparent !important;
    }
    
    /* Main Content Selectboxes - Spezifisch für Hauptprogramm */
    .main [data-testid="stSelectbox"] > div > div > div,
    .main [data-testid="stSelectbox"] select,
    .main [data-testid="stTextInput"] input,
    .main [data-testid="stNumberInput"] input,
    .block-container [data-testid="stSelectbox"] > div > div > div,
    .block-container [data-testid="stSelectbox"] select,
    .block-container [data-testid="stTextInput"] input,
    .block-container [data-testid="stNumberInput"] input,
    div[data-testid="stAppViewContainer"] [data-testid="stSelectbox"] > div > div > div,
    div[data-testid="stAppViewContainer"] [data-testid="stSelectbox"] select,
    div[data-testid="stAppViewContainer"] [data-testid="stTextInput"] input,
    div[data-testid="stAppViewContainer"] [data-testid="stNumberInput"] input {
        color: #FFFFFF !important;
        background-color: #2A2A2A !important;
        border-color: #555555 !important;
    }
    
    /* ULTIMATIVE CSS-ÜBERSCHREIBUNG - Höchste Priorität */
    * {
        --streamlit-button-color: #FFFFFF !important;
        --streamlit-button-background: #FF8C42 !important;
    }
    
    /* Alle möglichen Button-Selektoren überschreiben */
    button, 
    .stButton button,
    [data-testid="stButton"] button,
    div[data-testid="stAppViewContainer"] button,
    .block-container button,
    .main button {
        color: #FFFFFF !important;
        background-color: #FF8C42 !important;
        border-color: #FF6B35 !important;
    }
    
    /* Alle möglichen Text-Selektoren überschreiben */
    p, span, div, label,
    .stMarkdown p,
    .stMarkdown span,
    .stMarkdown div,
    .stText p,
    .stText span,
    .stText div,
    div[data-testid="stAppViewContainer"] p,
    div[data-testid="stAppViewContainer"] span,
    div[data-testid="stAppViewContainer"] div,
    .block-container p,
    .block-container span,
    .block-container div {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] button, [data-testid="stSidebar"] .stButton button {
        color: #FFFFFF !important;
        background-color: #FF8C42 !important;
        border-color: #FF6B35 !important;
    }
    button[data-testid="baseButton-primary"] {
        background-color: #FF8C42 !important;
        border: 2px solid #FF6B35 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
    }
    button[data-testid="baseButton-secondary"] {
        background-color: #FF8C42 !important;
        border: 2px solid #FF6B35 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
    }

    /* Spezielles Styling für den ⚙️ Einstellungen-Popover-Button,
       damit der orange Hintergrund die komplette Fläche abdeckt */
    [data-testid="stPopoverButton"] {
        background-color: #FF8C42 !important;
        border: 2px solid #FF6B35 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        padding: 0.35rem 1.4rem !important;  /* mehr horizontaler Abstand */
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-sizing: border-box !important;
        white-space: nowrap !important;      /* Text bleibt in einer Zeile */
        min-width: max-content !important;   /* Button mindestens so breit wie Inhalt */
        overflow: visible !important;        /* nichts wird abgeschnitten */
    }
    [data-testid="stToolbar"] .stCheckbox label {
        color: #000000 !important;
    }

    /* Sidebar-Menü-Ein-/Ausblende-Button (keyboard_double_arrow_right) - immer sichtbar und fixed positioniert */
    [data-testid="collapsedControl"] {
        position: fixed !important;
        top: 1rem !important;
        right: 1rem !important;
        left: auto !important;
        z-index: 1000 !important;
        display: block !important;
        visibility: visible !important;
    }
    
    [data-testid="collapsedControl"] button {
        background-color: #FF8C42 !important;
        border-radius: 999px !important;
        border: 2px solid #FF6B35 !important;
        color: #FFFFFF !important;
        display: block !important;
        visibility: visible !important;
    }
    
    /* Alle möglichen Input-Selektoren */
    input, select, textarea,
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] select,
    [data-testid="stSelectbox"] > div > div > div {
        color: #FFFFFF !important;
        background-color: #2A2A2A !important;
        border-color: #555555 !important;
    }
    
    /* Scrollbar Styling - Orange für bessere Sichtbarkeit */
    /* Webkit Browser (Chrome, Edge, Safari) */
    ::-webkit-scrollbar {
        width: 12px !important;
        height: 12px !important;
    }
    
    ::-webkit-scrollbar-track {
        background: var(--secondary-black) !important;
        border-radius: 6px !important;
    }
    
    ::-webkit-scrollbar-thumb {
        background: var(--primary-orange) !important;
        border-radius: 6px !important;
        border: 2px solid var(--secondary-black) !important;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: var(--secondary-orange) !important;
    }
    
    ::-webkit-scrollbar-thumb:active {
        background: var(--dark-orange) !important;
    }
    
    ::-webkit-scrollbar-corner {
        background: var(--secondary-black) !important;
    }
    
    /* Firefox */
    * {
        scrollbar-width: thin !important;
        scrollbar-color: var(--primary-orange) var(--secondary-black) !important;
    }

    /* Tab-Aktionsbuttons (⚙️ Einstellungen, ✏️ Umbenennen, 🗑️ Löschen) immer in einer Zeile halten */
    /* Container mit diesen Buttons auf no-wrap setzen */
    div:has(> [data-testid="stPopoverButton"]),
    div:has(> button[data-testid="stBaseButton-secondary"]) {
        display: flex !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        gap: 0.25rem !important;
        white-space: nowrap !important;
    }

    /* Buttons sollen NICHT schrumpfen, sondern immer so breit wie ihr Inhalt bleiben */
    [data-testid="stPopoverButton"],
    button[data-testid="stBaseButton-secondary"] {
        flex: 0 0 auto !important;          /* nicht schrumpfen */
        min-width: max-content !important;  /* mindestens so breit wie Text + Icon */
        white-space: nowrap !important;
    }

    /* Doppelten Text neben dem ⚙️-Button ausblenden, damit nichts über den Button-Hintergrund hinausragt */
    div:has(> [data-testid="stPopoverButton"]) > p {
        display: none !important;
    }
    
    /* Standard-Streamlit-Header-Buttons ausblenden */
    [data-testid="stAppDeployButton"],
    [data-testid="stMainMenu"] {
        display: none !important;
    }
    
    /* Plotly-Graphen sollen sich immer an Container-Größe anpassen */
    .js-plotly-plot {
        max-width: 100% !important;
        width: 100% !important;
    }
    
    .js-plotly-plot .plotly {
        width: 100% !important;
        max-width: 100% !important;
    }
</style>
<script>
(function() {
    // Warte bis DOM geladen ist
    function initPlotlyResize() {
        // Funktion zum Entfernen des Fullscreen-Buttons
        function removeFullscreenButtons() {
            const modebars = document.querySelectorAll('.plotly .modebar');
            modebars.forEach(function(modebar) {
                const buttons = modebar.querySelectorAll('.modebar-btn');
                buttons.forEach(function(btn) {
                    const title = btn.getAttribute('data-title') || btn.getAttribute('aria-label') || '';
                    const svg = btn.querySelector('svg');
                    const svgTitle = svg ? (svg.getAttribute('data-title') || svg.getAttribute('aria-label') || '') : '';
                    const combinedTitle = (title + ' ' + svgTitle).toLowerCase();
                    if (combinedTitle.includes('fullscreen') || combinedTitle.includes('full screen') || combinedTitle.includes('vollbild')) {
                        btn.style.display = 'none';
                        btn.remove();
                    }
                });
            });
        }
        
        // Funktion zum Neuanpassen aller Plotly-Graphen
        function resizeAllPlots() {
            const plots = document.querySelectorAll('.js-plotly-plot');
            plots.forEach(function(plot) {
                if (plot && window.Plotly) {
                    try {
                        // Hole die Plotly-Figure-ID
                        const plotId = plot.id || plot.getAttribute('id');
                        if (plotId) {
                            // Warte kurz, damit das Layout sich anpassen kann
                            setTimeout(function() {
                                // Trigger Plotly resize
                                window.Plotly.Plots.resize(plotId);
                            }, 100);
                        }
                    } catch (e) {
                        console.log('Plot resize error:', e);
                    }
                }
            });
        }
        
        // Überwache Fullscreen-Änderungen
        function handleFullscreenChange() {
            const isFullscreen = document.fullscreenElement || 
                                document.webkitFullscreenElement || 
                                document.mozFullScreenElement || 
                                document.msFullscreenElement;
            
            if (!isFullscreen) {
                // Fullscreen wurde geschlossen - warte kurz und passe Graphen an
                setTimeout(function() {
                    resizeAllPlots();
                    // Zweiter Versuch nach etwas längerer Wartezeit
                    setTimeout(resizeAllPlots, 300);
                }, 150);
            }
        }
        
        // Event-Listener für Fullscreen-Änderungen
        document.addEventListener('fullscreenchange', handleFullscreenChange);
        document.addEventListener('webkitfullscreenchange', handleFullscreenChange);
        document.addEventListener('mozfullscreenchange', handleFullscreenChange);
        document.addEventListener('MSFullscreenChange', handleFullscreenChange);
        
        // Überwache auch Resize-Events
        let resizeTimeout;
        window.addEventListener('resize', function() {
            clearTimeout(resizeTimeout);
            resizeTimeout = setTimeout(function() {
                resizeAllPlots();
            }, 200);
        });
        
        // Initial resize nach kurzer Wartezeit
        setTimeout(resizeAllPlots, 500);
        
        // Initial Fullscreen-Button-Entfernung
        setTimeout(removeFullscreenButtons, 500);
        setTimeout(removeFullscreenButtons, 1000);
        setTimeout(removeFullscreenButtons, 2000);
        
        // Observer für dynamisch hinzugefügte Graphen
        const observer = new MutationObserver(function(mutations) {
            let shouldResize = false;
            let shouldRemoveButtons = false;
            mutations.forEach(function(mutation) {
                if (mutation.addedNodes.length) {
                    mutation.addedNodes.forEach(function(node) {
                        if (node.nodeType === 1) {
                            if (node.classList && node.classList.contains('js-plotly-plot')) {
                                shouldResize = true;
                                shouldRemoveButtons = true;
                            }
                            if (node.querySelector && node.querySelector('.js-plotly-plot')) {
                                shouldResize = true;
                                shouldRemoveButtons = true;
                            }
                            if (node.querySelector && node.querySelector('.modebar')) {
                                shouldRemoveButtons = true;
                            }
                        }
                    });
                }
            });
            if (shouldResize) {
                setTimeout(resizeAllPlots, 300);
            }
            if (shouldRemoveButtons) {
                setTimeout(removeFullscreenButtons, 300);
                setTimeout(removeFullscreenButtons, 600);
            }
        });
        
        // Starte Beobachtung
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
        
        // Regelmäßige Überprüfung und Entfernung der Fullscreen-Buttons (falls sie wieder erscheinen)
        setInterval(removeFullscreenButtons, 2000);
    }
    
    // Initialisiere wenn DOM bereit ist
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initPlotlyResize);
    } else {
        initPlotlyResize();
    }
    
    // Auch nach Streamlit-Rerun (wenn window.location sich ändert)
    if (window.addEventListener) {
        window.addEventListener('load', initPlotlyResize);
    }
    
        // Button mit keyboard_double_arrow_left Icon komplett entfernen
    function hideLeftArrowButton() {
        // Suche nach allen Icons
        const icons = document.querySelectorAll('[data-testid="stIconMaterial"]');
        icons.forEach(function(icon) {
            // Prüfe ob der Text keyboard_double_arrow_left enthält
            if (icon.textContent && icon.textContent.trim() === 'keyboard_double_arrow_left') {
                // Verstecke das Icon selbst
                icon.style.display = 'none';
                icon.style.visibility = 'hidden';
                // Finde alle möglichen Parent-Elemente und verstecke sie
                let element = icon;
                for (let i = 0; i < 10; i++) { // Maximal 10 Ebenen nach oben
                    element = element.parentElement;
                    if (!element) break;
                    // Verstecke Button, stButton Container, etc.
                    if (element.tagName === 'BUTTON' || 
                        element.classList.contains('stButton') ||
                        element.classList.contains('stBaseButton') ||
                        element.getAttribute('data-testid') === 'stBaseButton') {
                        element.style.display = 'none';
                        element.style.visibility = 'hidden';
                    }
                }
            }
        });
    }
    
    // keyboard_double_arrow_right Button immer sichtbar halten
    function ensureRightArrowVisible() {
        const collapsedControl = document.querySelector('[data-testid="collapsedControl"]');
        if (collapsedControl) {
            collapsedControl.style.position = 'fixed';
            collapsedControl.style.top = '1rem';
            collapsedControl.style.right = '1rem';
            collapsedControl.style.left = 'auto';
            collapsedControl.style.zIndex = '1000';
            collapsedControl.style.display = 'block';
            collapsedControl.style.visibility = 'visible';
            const button = collapsedControl.querySelector('button');
            if (button) {
                button.style.display = 'block';
                button.style.visibility = 'visible';
            }
        }
    }
    
    // Initial hide/show
    setTimeout(function() {
        hideLeftArrowButton();
        ensureRightArrowVisible();
    }, 100);
    setTimeout(function() {
        hideLeftArrowButton();
        ensureRightArrowVisible();
    }, 500);
    setTimeout(function() {
        hideLeftArrowButton();
        ensureRightArrowVisible();
    }, 1000);
    setTimeout(function() {
        hideLeftArrowButton();
        ensureRightArrowVisible();
    }, 2000);
    
    // Observer für dynamisch hinzugefügte Buttons
    const buttonObserver = new MutationObserver(function(mutations) {
        let shouldUpdate = false;
        mutations.forEach(function(mutation) {
            if (mutation.addedNodes.length) {
                mutation.addedNodes.forEach(function(node) {
                    if (node.nodeType === 1) {
                        if (node.querySelector && node.querySelector('[data-testid="stIconMaterial"]')) {
                            shouldUpdate = true;
                        }
                        if (node.getAttribute && node.getAttribute('data-testid') === 'stIconMaterial') {
                            shouldUpdate = true;
                        }
                        if (node.getAttribute && node.getAttribute('data-testid') === 'collapsedControl') {
                            shouldUpdate = true;
                        }
                    }
                });
            }
        });
        if (shouldUpdate) {
            setTimeout(function() {
                hideLeftArrowButton();
                ensureRightArrowVisible();
            }, 100);
        }
    });
    
    buttonObserver.observe(document.body, {
        childList: true,
        subtree: true
    });
    
    // Regelmäßige Überprüfung - beide Funktionen
    setInterval(function() {
        hideLeftArrowButton();
        ensureRightArrowVisible();
    }, 1000);
})();
</script>
""", unsafe_allow_html=True)

st.set_page_config(
    layout="wide", 
    page_title="Batteriespeicher Analyse Tool",
    page_icon="lava_energy_icon.ico"
)

# Haupttitel mit Icon und Styling
# Benutzerdefinierte Header-Buttons
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
header_actions = st.columns([0.85, 0.15], gap="small")
st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
with header_actions[1]:
    with st.popover("⚙️ Einstellungen", use_container_width=False):
        st.subheader("Eigenes Einstellungsmenü")
        st.markdown("**PV-Datenquelle**")
        settings_link_value = st.text_input(
            "Link (z.B. PVGIS)",
            value=st.session_state.get('pv_data_source_link_input', ''),
            key="settings_pv_link_input"
        )
        if st.button("💾 Link speichern", use_container_width=True, key="save_pv_link_button"):
            cleaned_link = settings_link_value.strip()
            st.session_state['pv_data_source_link_input'] = cleaned_link
            save_pv_link_to_file(cleaned_link)
            st.success("Link gespeichert.")
        st.caption("Der gespeicherte Link erscheint als Schnellzugriff unter den PV-Erzeugungsdaten.")

st.markdown("""
<div style="text-align: center; padding: 12px 0; border-bottom: 2px solid #FF6B35; margin-top: 10px; margin-bottom: 20px;">
    <h1 style="color: #FF6B35; font-size: 2.2em; margin: 10px 0; font-weight: 600;">🔋 Batteriespeicher Analyse Tool</h1>
</div>
""", unsafe_allow_html=True)

# --- Einstellungsverwaltung ---
settings_manager = SettingsManager()

# ---------------------------
# Analyse-Status verwalten
# ---------------------------
ANALYSIS_RESULT_KEYS = [
    'optimization_results',
    'variable_tariff_result',
    'current_settings',
    'calculation_details',
    'cost_comparison_details',
    'original_consumption_series',
    'scaled_pv_generation_series',
    'battery_cost_curve',
    'optimal_sim_result',
    'optimal_capacity',
    'results_summary',
    'time_series_timerange',
    'energy_flow_fig_full',
]
ANALYSIS_PAYLOAD_KEY = 'analysis_payload'

def init_analysis_state():
    if 'analysis_state' not in st.session_state:
        st.session_state['analysis_state'] = 'idle'  # idle | running | ready | error
    if 'analysis_error' not in st.session_state:
        st.session_state['analysis_error'] = None
    if 'analysis_completed' not in st.session_state:
        st.session_state['analysis_completed'] = False

def reset_analysis_results():
    for key in ANALYSIS_RESULT_KEYS:
        if key in st.session_state:
            del st.session_state[key]

def clear_analysis_payload():
    if ANALYSIS_PAYLOAD_KEY in st.session_state:
        del st.session_state[ANALYSIS_PAYLOAD_KEY]

def enqueue_analysis(payload: dict):
    reset_analysis_results()
    st.session_state[ANALYSIS_PAYLOAD_KEY] = payload
    st.session_state['analysis_state'] = 'running'
    st.session_state['analysis_error'] = None
    st.session_state['analysis_completed'] = False

def fail_analysis(message: str):
    reset_analysis_results()
    clear_analysis_payload()
    st.session_state['analysis_state'] = 'error'
    st.session_state['analysis_error'] = message
    st.session_state['analysis_completed'] = False

def finalize_analysis():
    st.session_state['analysis_state'] = 'ready'
    st.session_state['analysis_error'] = None
    st.session_state['analysis_completed'] = True
    clear_analysis_payload()

def _update_status(status_placeholder, progress_bar, message: str, value: float | None = None):
    if status_placeholder is not None:
        status_placeholder.info(message)
    if progress_bar is not None and value is not None:
        try:
            progress_bar.progress(max(0, min(100, int(value))))
        except Exception:
            pass

def run_analysis_job(payload: dict, status_placeholder=None, progress_bar=None):
    params = payload.get('params', {})
    _update_status(status_placeholder, progress_bar, "Initialisiere Analyse...", 1)

    if not payload.get('consumption_file') or not payload.get('pv_file'):
        raise ValueError("Eingabedateien konnten nicht geladen werden.")

    consumption_file = io.BytesIO(payload['consumption_file'])
    consumption_file.name = payload.get('consumption_name', 'consumption.xlsx')
    pv_file = io.BytesIO(payload['pv_file'])
    pv_file.name = payload.get('pv_name', 'pv_data')
    pv_extension = payload.get('pv_extension', 'xlsx').lower()

    selected_year = params.get('selected_year')
    annual_consumption_kwh = params.get('annual_consumption_kwh')
    number_of_persons = params.get('number_of_persons')
    bundesland_code = params.get('bundesland_code')
    profile_type = params.get('load_profile_type', 'H25')

    # 1) Daten laden
    _update_status(status_placeholder, progress_bar, "Lade Verbrauchsdaten...", 5)
    consumption_series = load_consumption_from_excel(
        consumption_file,
        selected_year,
        annual_consumption_kwh,
        number_of_persons,
        bundesland_code,
        profile_type
    )
    if consumption_series is None:
        raise ValueError("Verbrauchsdaten konnten nicht geladen werden.")

    _update_status(status_placeholder, progress_bar, "Lade PV-Daten...", 15)
    if pv_extension == 'csv':
        pv_generation_series = load_pv_generation_from_csv(pv_file, selected_year)
    else:
        pv_generation_series = load_pv_generation_from_excel(pv_file, selected_year)
    if pv_generation_series is None:
        raise ValueError("PV-Daten konnten nicht geladen werden.")

    _update_status(status_placeholder, progress_bar, "Lade technische Parameter...", 25)
    battery_cost_curve = load_battery_cost_curve()
    if battery_cost_curve is None:
        raise ValueError("Batteriekostenkurven-Datei konnte nicht geladen werden.")

    # Versuche technische Parameter zu laden, aber mache es optional
    try:
        battery_tech_params = load_battery_tech_params("Batteriespeicherkosten.xlsm")
        if battery_tech_params is None:
            st.warning("⚠️ Technische Batterieparameter konnten nicht geladen werden. Verwende Standardwerte.")
            battery_tech_params = {}
    except Exception as e:
        st.warning(f"⚠️ Technische Batterieparameter konnten nicht geladen werden: {str(e)}. Verwende Standardwerte.")
        battery_tech_params = {}

    scaled_consumption_series = consumption_series
    scaled_pv_generation_series = pv_generation_series
    original_consumption_series = consumption_series.copy()

    # 2) Optimierung
    _update_status(status_placeholder, progress_bar, "Optimiere Batteriespeichergröße...", 45)
    optimization_results = find_optimal_size(
        consumption_series=scaled_consumption_series,
        pv_generation_series=scaled_pv_generation_series,
        min_capacity_kwh=params.get('min_battery_capacity'),
        max_capacity_kwh=params.get('max_battery_capacity'),
        step_kwh=params.get('battery_step_size'),
        battery_efficiency_charge=params.get('battery_efficiency_charge'),
        battery_efficiency_discharge=params.get('battery_efficiency_discharge'),
        battery_max_charge_kw=DEFAULT_BATTERY_MAX_CHARGE_KW,
        battery_max_discharge_kw=DEFAULT_BATTERY_MAX_DISCHARGE_KW,
        price_grid_per_kwh=params.get('price_grid_per_kwh'),
        price_feed_in_per_kwh=params.get('price_feed_in_per_kwh'),
        battery_cost_curve=battery_cost_curve,
        project_lifetime_years=params.get('project_lifetime_years'),
        discount_rate=params.get('discount_rate'),
        initial_soc_percent=params.get('initial_soc_percent'),
        min_soc_percent=params.get('min_soc_percent'),
        max_soc_percent=params.get('max_soc_percent'),
        annual_capacity_loss_percent=params.get('annual_capacity_loss_percent'),
        battery_tech_params=battery_tech_params,
        project_interest_rate_db=params.get('project_interest_rate_db')
    )
    if not optimization_results:
        raise ValueError("Keine Ergebnisse bei der Optimierung erhalten.")

    df_results = pd.DataFrame(optimization_results)
    if df_results.empty:
        raise ValueError("Optimierungsliste ist leer.")

    criterion = params.get('optimization_criterion', 'Deckungsbeitrag III gesamt (Barwert)')
    # Nur DB Barwert und DB Nominal unterstützen
    if criterion == "Deckungsbeitrag III gesamt (Nominal)":
        best_idx = df_results['total_db3_nominal'].idxmax()
        criterion_value = df_results.loc[best_idx, 'total_db3_nominal']
        criterion_name = "DB III gesamt (Nominal)"
    else:  # Default: Deckungsbeitrag III gesamt (Barwert)
        best_idx = df_results['total_db3_present_value'].idxmax()
        criterion_value = df_results.loc[best_idx, 'total_db3_present_value']
        criterion_name = "DB III gesamt (Barwert)"

    best_option = df_results.loc[best_idx]
    optimal_capacity = float(best_option['battery_capacity_kwh'])
    investment_cost = float(best_option.get('investment_cost', 0) or 0)

    # 3) Variable Tarife (Beispiel)
    _update_status(status_placeholder, progress_bar, "Berechne variable Tarife...", 60)
    price_series_length = len(scaled_consumption_series)
    hourly_prices_grid = pd.Series(
        params.get('price_grid_per_kwh', 0.3) * (1 + 0.2 * np.sin(np.linspace(0, 2 * np.pi * 365, price_series_length))),
        index=scaled_consumption_series.index
    )
    hourly_prices_feed_in = pd.Series(
        params.get('price_feed_in_per_kwh', 0.08) * (1 - 0.1 * np.sin(np.linspace(0, 2 * np.pi * 365, price_series_length))),
        index=scaled_consumption_series.index
    )
    variable_tariff_result = run_variable_tariff_scenario(
        consumption_series=scaled_consumption_series,
        pv_generation_series=scaled_pv_generation_series,
        battery_capacity_kwh=params.get('min_battery_capacity'),
        battery_efficiency_charge=params.get('battery_efficiency_charge'),
        battery_efficiency_discharge=params.get('battery_efficiency_discharge'),
        battery_max_charge_kw=DEFAULT_BATTERY_MAX_CHARGE_KW,
        battery_max_discharge_kw=DEFAULT_BATTERY_MAX_DISCHARGE_KW,
        price_grid_series=hourly_prices_grid,
        price_feed_in_series=hourly_prices_feed_in,
        battery_cost_curve=battery_cost_curve,
        project_lifetime_years=params.get('project_lifetime_years'),
        discount_rate=params.get('discount_rate'),
        initial_soc_percent=params.get('initial_soc_percent'),
        min_soc_percent=params.get('min_soc_percent'),
        max_soc_percent=params.get('max_soc_percent'),
        annual_capacity_loss_percent=params.get('annual_capacity_loss_percent'),
        battery_tech_params=battery_tech_params
    )

    # 4) Simulation mit optimaler Kapazität
    _update_status(status_placeholder, progress_bar, "Simuliere optimales Szenario...", 75)
    def resolve_power_for_capacity(cap_kwh: float):
        cap_int = int(round(cap_kwh))
        if battery_tech_params and cap_int in battery_tech_params:
            data = battery_tech_params[cap_kwh if cap_kwh in battery_tech_params else cap_int]
        else:
            data = None
        if not data and battery_tech_params:
            keys = sorted(battery_tech_params.keys())
            if keys:
                nearest = min(keys, key=lambda k: abs(k - cap_int))
                data = battery_tech_params.get(nearest)
        if not data:
            return DEFAULT_BATTERY_MAX_CHARGE_KW, DEFAULT_BATTERY_MAX_DISCHARGE_KW
        return (
            float(data.get("max_charge_kw", DEFAULT_BATTERY_MAX_CHARGE_KW)),
            float(data.get("max_discharge_kw", DEFAULT_BATTERY_MAX_DISCHARGE_KW)),
        )

    opt_charge_kw, opt_discharge_kw = resolve_power_for_capacity(optimal_capacity)
    optimal_sim_result = simulate_one_year(
        consumption_series=scaled_consumption_series,
        pv_generation_series=scaled_pv_generation_series,
        battery_capacity_kwh=optimal_capacity,
        battery_efficiency_charge=params.get('battery_efficiency_charge'),
        battery_efficiency_discharge=params.get('battery_efficiency_discharge'),
        battery_max_charge_kw=opt_charge_kw,
        battery_max_discharge_kw=opt_discharge_kw,
        price_grid_per_kwh=params.get('price_grid_per_kwh'),
        price_feed_in_per_kwh=params.get('price_feed_in_per_kwh'),
        initial_soc_percent=params.get('initial_soc_percent'),
        min_soc_percent=params.get('min_soc_percent'),
        max_soc_percent=params.get('max_soc_percent'),
        annual_capacity_loss_percent=params.get('annual_capacity_loss_percent'),
        simulation_year=1
    )

    _update_status(status_placeholder, progress_bar, "Berechne finanzielle Kennzahlen...", 88)
    financials = calculate_financial_kpis(
        annual_energy_cost_with_battery=optimal_sim_result["kpis"]["annual_energy_cost"],
        total_consumption=optimal_sim_result["kpis"]["total_consumption_kwh"],
        total_pv_generation=optimal_sim_result["kpis"]["total_pv_generation_kwh"],
        battery_capacity_kwh=optimal_capacity,
        battery_cost_curve=battery_cost_curve,
        project_lifetime_years=params.get('project_lifetime_years'),
        discount_rate=params.get('discount_rate'),
        price_grid_per_kwh=params.get('price_grid_per_kwh'),
        price_feed_in_per_kwh=params.get('price_feed_in_per_kwh'),
        annual_capacity_loss_percent=params.get('annual_capacity_loss_percent'),
        consumption_series=scaled_consumption_series,
        pv_generation_series=scaled_pv_generation_series,
        battery_efficiency_charge=params.get('battery_efficiency_charge'),
        battery_efficiency_discharge=params.get('battery_efficiency_discharge'),
        battery_max_charge_kw=opt_charge_kw,
        battery_max_discharge_kw=opt_discharge_kw,
        initial_soc_percent=params.get('initial_soc_percent'),
        min_soc_percent=params.get('min_soc_percent'),
        max_soc_percent=params.get('max_soc_percent'),
        grid_import_with_battery=optimal_sim_result["kpis"]["total_grid_import_kwh"],
        grid_export_with_battery=optimal_sim_result["kpis"]["total_grid_export_kwh"]
    )

    annual_savings = financials.get("annual_savings", 0.0) or 0.0
    total_savings = annual_savings * params.get('project_lifetime_years')
    roi_percentage = (total_savings / investment_cost - 1) * 100 if investment_cost else 0.0

    results_summary = {
        'optimal_capacity': float(optimal_capacity),
        'criterion_name': criterion_name,
        'criterion_value': float(criterion_value) if pd.notna(criterion_value) else None,
        'investment_cost': float(investment_cost),
        'annual_savings': float(annual_savings),
        'roi_percentage': float(roi_percentage),
        'payback_period_years': float(financials.get("payback_period_years")) if financials.get("payback_period_years") is not None else None,
        'irr_percentage': float(financials.get("irr_percentage")) if financials.get("irr_percentage") is not None else None,
        'project_lifetime_years': int(params.get('project_lifetime_years'))
    }

    _update_status(status_placeholder, progress_bar, "Bereite Ausgabedaten vor...", 96)
    sankey_with_battery = plot_sankey_diagram(optimal_sim_result["kpis"])
    calculation_details = sankey_with_battery.get('calculation_details')

    cost_comp = plot_cost_comparison_with_without_battery(
        annual_energy_cost_with_battery=optimal_sim_result["kpis"]["annual_energy_cost"],
        total_consumption=optimal_sim_result["kpis"]["total_consumption_kwh"],
        total_pv_generation=optimal_sim_result["kpis"]["total_pv_generation_kwh"],
        consumption_series=scaled_consumption_series,
        pv_generation_series=scaled_pv_generation_series,
        battery_capacity_kwh=optimal_capacity,
        battery_efficiency_charge=params.get('battery_efficiency_charge'),
        battery_efficiency_discharge=params.get('battery_efficiency_discharge'),
        battery_max_charge_kw=opt_charge_kw,
        battery_max_discharge_kw=opt_discharge_kw,
        price_grid_per_kwh=params.get('price_grid_per_kwh'),
        price_feed_in_per_kwh=params.get('price_feed_in_per_kwh'),
        battery_cost_curve=battery_cost_curve,
        project_lifetime_years=params.get('project_lifetime_years'),
        discount_rate=params.get('discount_rate'),
        initial_soc_percent=params.get('initial_soc_percent'),
        min_soc_percent=params.get('min_soc_percent'),
        max_soc_percent=params.get('max_soc_percent'),
        annual_capacity_loss_percent=params.get('annual_capacity_loss_percent'),
        grid_export_with_battery=optimal_sim_result["kpis"]["total_grid_export_kwh"],
        grid_import_with_battery=optimal_sim_result["kpis"]["total_grid_import_kwh"]
    )
    cost_comparison_details = cost_comp.get('comparison_details')
    if cost_comparison_details is not None:
        cost_comparison_details['irr_percentage'] = financials.get('irr_percentage')
        cost_comparison_details['payback_period_years'] = financials.get('payback_period_years')
        cost_comparison_details['annual_savings'] = annual_savings

    energy_flow_fig = plot_energy_flows_for_period(
        optimal_sim_result["time_series_data"],
        optimal_sim_result["time_series_data"].index.min().strftime('%Y-%m-%d'),
        optimal_sim_result["time_series_data"].index.max().strftime('%Y-%m-%d'),
        optimal_capacity
    )

    # Ergebnisse speichern
    st.session_state['current_settings'] = payload.get('current_settings', {})
    st.session_state['optimization_results'] = optimization_results
    st.session_state['variable_tariff_result'] = variable_tariff_result
    st.session_state['optimal_sim_result'] = optimal_sim_result
    st.session_state['optimal_capacity'] = optimal_capacity
    st.session_state['results_summary'] = results_summary
    st.session_state['calculation_details'] = calculation_details
    st.session_state['cost_comparison_details'] = cost_comparison_details
    st.session_state['original_consumption_series'] = original_consumption_series
    st.session_state['scaled_pv_generation_series'] = scaled_pv_generation_series
    st.session_state['battery_cost_curve'] = battery_cost_curve
    st.session_state['time_series_timerange'] = {
        'start': optimal_sim_result["time_series_data"].index.min().strftime('%Y-%m-%d'),
        'end': optimal_sim_result["time_series_data"].index.max().strftime('%Y-%m-%d')
    }
    st.session_state['energy_flow_fig_full'] = energy_flow_fig

    _update_status(status_placeholder, progress_bar, "Analyse abgeschlossen.", 100)

init_analysis_state()
analysis_state = st.session_state.get('analysis_state', 'idle')
analysis_ready = analysis_state == 'ready'
st.session_state['analysis_completed'] = analysis_ready

if analysis_state == 'running':
    status_placeholder = st.empty()
    progress_bar = st.progress(0)
    payload = st.session_state.get(ANALYSIS_PAYLOAD_KEY)
    if not payload:
        fail_analysis("Es wurden keine Analysedaten gefunden. Bitte Analyse erneut starten.")
        st.rerun()
    try:
        run_analysis_job(payload, status_placeholder=status_placeholder, progress_bar=progress_bar)
        finalize_analysis()
    except Exception as analysis_error:
        fail_analysis(str(analysis_error))
    st.rerun()

if analysis_state == 'error' and st.session_state.get('analysis_error'):
    st.error(f"Analyse fehlgeschlagen: {st.session_state['analysis_error']}")
    if st.button("🔁 Analyse zurücksetzen"):
        reset_analysis_results()
        clear_analysis_payload()
        st.session_state['analysis_state'] = 'idle'
        st.session_state['analysis_error'] = None
        st.rerun()

# --- Seitenleiste für Parameter-Eingabe ---
st.sidebar.markdown("""
<div style="text-align: center; padding: 10px 0; border-bottom: 2px solid #FF6B35; margin-bottom: 20px;">
    <h3 style="color: #FF6B35; margin: 0;">⚙️ Simulationseinstellungen</h3>
</div>
""", unsafe_allow_html=True)

# Einstellungsverwaltung
st.sidebar.markdown("""
<div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
    <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">💾 Einstellungen speichern/laden</h4>
</div>
""", unsafe_allow_html=True)

# Aktuelle Einstellungen sammeln (wird später gefüllt)
current_settings = {}

# Gespeicherte Konfigurationen anzeigen
saved_configs = settings_manager.get_all_configurations()

if saved_configs:
    selected_config = st.sidebar.selectbox(
        "Gespeicherte Konfigurationen",
        ["-- Neue Konfiguration --"] + saved_configs,
        key="config_selector"
    )
    
    # Prüfe ob eine neue Konfiguration ausgewählt wurde
    if 'last_loaded_config' not in st.session_state:
        st.session_state['last_loaded_config'] = None
    
    # Wenn eine neue Konfiguration ausgewählt wurde, lade sie
    if selected_config != "-- Neue Konfiguration --":
        if st.session_state['last_loaded_config'] != selected_config:
            # Neue Konfiguration wurde ausgewählt - lade sie
            loaded_settings = settings_manager.load_configuration(selected_config)
            if loaded_settings:
                st.session_state['last_loaded_config'] = selected_config
                st.session_state['loaded_config_settings'] = loaded_settings.copy()
                # Zurücksetzen des "angewendet"-Flags, damit neue Werte verwendet werden
                if 'config_applied' in st.session_state:
                    del st.session_state['config_applied']
                if 'applied_config_name' in st.session_state:
                    del st.session_state['applied_config_name']
                st.sidebar.success(f"✅ Konfiguration '{selected_config}' geladen. Sie können die Werte jetzt anpassen.")
                config_link = loaded_settings.get('pv_data_source_link')
                if config_link is not None:
                    st.session_state['pv_data_source_link_input'] = config_link
                    save_pv_link_to_file(config_link)
    else:
        # Keine Konfiguration ausgewählt - zurücksetzen
        if 'last_loaded_config' in st.session_state:
            st.session_state['last_loaded_config'] = None
        if 'loaded_config_settings' in st.session_state:
            del st.session_state['loaded_config_settings']
        if 'config_applied' in st.session_state:
            del st.session_state['config_applied']
        if 'applied_config_name' in st.session_state:
            del st.session_state['applied_config_name']
    
    # Konfiguration verwalten
    col1, col2 = st.sidebar.columns(2)
    
    with col1:
        if st.button("🗑️ Löschen", key="delete_config"):
            if selected_config != "-- Neue Konfiguration --":
                if settings_manager.delete_configuration(selected_config):
                    st.sidebar.success(f"Konfiguration '{selected_config}' gelöscht")
                    st.rerun()
    
    with col2:
        if st.button("✏️ Umbenennen", key="rename_config"):
            if selected_config != "-- Neue Konfiguration --":
                st.session_state.show_rename = True
    
current_pv_link = st.session_state.get('pv_data_source_link_input', '')

# Link zur PV-Datenquelle verwalten
st.sidebar.markdown("""
<div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
    <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">📁 Daten hochladen</h4>
</div>
""", unsafe_allow_html=True)

# Separate Dateien für Verbrauch und Erzeugung
uploaded_consumption_file = st.sidebar.file_uploader(
    "📊 Verbrauchsdaten (Excel)", 
    type="xlsx", 
    key="consumption_upload",
    help="Excel-Datei mit Ihren Verbrauchsdaten. Erwartete Spalten: Zeitstempel und Verbrauch_kWh"
)
uploaded_pv_file = st.sidebar.file_uploader(
    "☀️ PV-Erzeugungsdaten (Excel/CSV)", 
    type=["xlsx", "csv"], 
    key="pv_upload",
    help="Excel- oder CSV-Datei mit PV-Erzeugungsdaten. CSV: Spalten 'time' und 'P' (PVGIS-Format). Excel: Zeitstempel und PV_Erzeugung_kWh"
)
pv_link_display = current_pv_link.strip()
if pv_link_display:
    st.sidebar.markdown(
        f"""
        <div style="background-color: #2D2D2D; padding: 10px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <a href="{pv_link_display}" target="_blank" style="color: #FFB366; text-decoration: underline;">🔗 PV-Datenquelle öffnen</a>
        </div>
        """,
        unsafe_allow_html=True
)
uploaded_file = None

# Standard-Verbrauchsdatei automatisch verwenden, falls verfügbar und keine eigene Datei gewählt wurde
if uploaded_consumption_file is None:
    if DEFAULT_CONSUMPTION_FILE.exists():
        try:
            default_bytes = DEFAULT_CONSUMPTION_FILE.read_bytes()
            default_buffer = io.BytesIO(default_bytes)
            default_buffer.name = DEFAULT_CONSUMPTION_FILE.name
            uploaded_consumption_file = default_buffer
            st.sidebar.info(f"📁 Standard-Verbrauchsdatei '{DEFAULT_CONSUMPTION_FILE.name}' wird verwendet. Lade eine eigene Datei hoch, um sie zu ersetzen.")
        except Exception as default_err:
            st.sidebar.error(f"Standard-Verbrauchsdatei konnte nicht geladen werden: {default_err}")
    else:
        st.sidebar.warning(f"⚠️ Standard-Verbrauchsdatei '{DEFAULT_CONSUMPTION_FILE.name}' wurde nicht gefunden.")

# Datenverarbeitung
consumption_series = None
pv_generation_series = None
parameters = {}

# Prüfe ob beide Dateien hochgeladen wurden
data_available = uploaded_consumption_file is not None and uploaded_pv_file is not None

if not analysis_ready:
    if not data_available:
        st.markdown("""
        <div style="background-color: #2D2D2D; padding: 30px; border-radius: 12px; border: 2px solid #FF6B35; text-align: center; margin: 50px 0;">
            <h3 style="color: #FF6B35; margin: 0 0 15px 0; font-size: 1.5em;">📁 Separate Dateien hochladen</h3>
            <p style="color: #FFFFFF; font-size: 1.1em; margin: 0;">Bitte laden Sie separate Excel-Dateien für Ihre Verbrauchs- und PV-Erzeugungsdaten hoch.</p>
            <p style="color: #FFFFFF; opacity: 0.8; margin: 10px 0 0 0; font-size: 0.9em;">Beide Dateien sind erforderlich, um die Simulation zu starten.</p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()
    else:
        st.sidebar.success("Separate Verbrauchs- und Erzeugungsdaten erfolgreich geladen.")
        
        # Geladene Einstellungen verwenden - nur wenn eine Konfiguration geladen wurde
        # und noch nicht angewendet wurde (nur beim ersten Laden)
        loaded_settings = None
        if 'config_selector' in st.session_state:
            selected_config = st.session_state.config_selector
            # Verwende gespeicherte Einstellungen nur wenn eine Konfiguration aktiv ist
            if selected_config != "-- Neue Konfiguration --" and 'loaded_config_settings' in st.session_state:
                # Verwende die gespeicherten Einstellungen nur beim ersten Laden
                # Danach werden die aktuellen UI-Eingaben verwendet
                if 'config_applied' not in st.session_state or st.session_state.get('applied_config_name') != selected_config:
                    loaded_settings = st.session_state['loaded_config_settings']
                    st.session_state['config_applied'] = True
                    st.session_state['applied_config_name'] = selected_config
                # Nach dem ersten Laden: loaded_settings bleibt None, damit aktuelle UI-Werte als Default verwendet werden
                # Die UI-Eingabefelder behalten ihre Werte durch Streamlit's State-Management

    # --- Allgemeine Parameter ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">🏠 Allgemeine Parameter</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Standardwerte oder geladene Werte verwenden
        default_persons = loaded_settings.get('number_of_persons', 1) if loaded_settings else 1
        default_pv_size = loaded_settings.get('pv_system_size_kwp', 10.0) if loaded_settings else 10.0
        
        # Bundesland-Auswahl für Feiertagsberechnung
        bundesland_options = {
            'Baden-Württemberg': 'BW',
            'Bayern': 'BY', 
            'Berlin': 'BE',
            'Brandenburg': 'BB',
            'Bremen': 'HB',
            'Hamburg': 'HH',
            'Hessen': 'HE',
            'Mecklenburg-Vorpommern': 'MV',
            'Niedersachsen': 'NI',
            'Nordrhein-Westfalen': 'NW',
            'Rheinland-Pfalz': 'RP',
            'Saarland': 'SL',
            'Sachsen': 'SN',
            'Sachsen-Anhalt': 'ST',
            'Schleswig-Holstein': 'SH',
            'Thüringen': 'TH'
        }
        
        default_bundesland = loaded_settings.get('bundesland', 'Baden-Württemberg') if loaded_settings else 'Baden-Württemberg'
        
        selected_bundesland = st.sidebar.selectbox(
            "Bundesland für Feiertagsberechnung",
            options=list(bundesland_options.keys()),
            index=list(bundesland_options.keys()).index(default_bundesland),
            help="Wird für die korrekte Berechnung der Feiertage und damit der Lastprofile benötigt."
        )
        
        bundesland_code = bundesland_options[selected_bundesland]
        
        # Jahresauswahl für Wochentag-Sheets
        year_options = [2024, 2025, 2026, 2027, 2028]
        default_year = loaded_settings.get('selected_year', 2024) if loaded_settings else 2024
        
        selected_year = st.sidebar.selectbox(
            "Jahr für Wochentag-Berechnung",
            options=year_options,
            index=year_options.index(default_year),
            help="Wählt das Jahr für die Wochentag- und Feiertagsberechnung."
        )
        
        # Profiltyp-Auswahl
        profile_type_options = {
            'H25': 'H25 - Haushalte',
            'G25': 'G25 - Gewerbe',
            'L25': 'L25 - Landwirtschaft',
            'P25': 'P25 - Pumpen',
            'S25': 'S25 - Speicherheizung'
        }
        default_profile_type = loaded_settings.get('load_profile_type', 'H25') if loaded_settings else 'H25'
        # Stelle sicher, dass der Standardwert gültig ist
        if default_profile_type not in profile_type_options:
            default_profile_type = 'H25'
        
        selected_profile_type = st.sidebar.selectbox(
            "Lastprofil-Typ",
            options=list(profile_type_options.keys()),
            index=list(profile_type_options.keys()).index(default_profile_type),
            format_func=lambda x: profile_type_options[x],
            help="Wählt das Standard-Lastprofil für die Verbrauchsdaten. H25 ist für Haushalte, G25 für Gewerbe, etc."
        )
        
        # Daten werden direkt aus den hochgeladenen Excel-Dateien verwendet
        st.sidebar.info("📊 Verbrauch und PV-Erzeugung werden direkt aus Ihren hochgeladenen Daten verwendet")
        
        number_of_persons = st.sidebar.number_input(
            "Anzahl Haushalte", min_value=1, max_value=50, value=default_persons, step=1,
            key="ui_number_of_persons",
            help="Anzahl der Haushalte (nicht Personen). Jeder Haushalt wird mit dem unten angegebenen jährlichen Durchschnittsverbrauch pro Haushalt berechnet."
        )
        pv_system_size_kwp = st.sidebar.number_input(
            "PV-Anlagengröße (kWp)", min_value=0.1, max_value=100.0, value=default_pv_size, step=0.1,
            key="ui_pv_system_size_kwp",
            help="Die Größe Ihrer PV-Anlage in kWp. Wird für den Bericht verwendet. Ihre PV-Daten sollten bereits für diese Anlagengröße berechnet sein."
        )
        
        # Jährlicher Durchschnittsverbrauch pro Haushalt
        default_annual_consumption = loaded_settings.get('annual_consumption_kwh', 3500) if loaded_settings else 3500
        # Stelle sicher, dass der Standardwert innerhalb der Grenzen liegt
        if default_annual_consumption > 20000:
            default_annual_consumption = 3500  # Fallback auf Standardwert
        annual_consumption_per_household = st.sidebar.number_input(
            "Jährlicher Durchschnittsverbrauch pro Haushalt (kWh)", 
            min_value=1000, max_value=20000, value=default_annual_consumption, step=100,
            key="ui_annual_consumption_per_household",
            help="Der durchschnittliche jährliche Stromverbrauch eines Haushalts. Wird für die Verbrauchsaufschlüsselung und Berechnungen benötigt."
        )
        
        # Berechne den Gesamtverbrauch für alle Haushalte
        annual_consumption_kwh = annual_consumption_per_household * number_of_persons

        # --- Batterieparameter ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">🔋 Batterieparameter</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Hinweis zur Kostenkurve
        st.sidebar.info(f"💰 Kosten basieren auf Kostenkurve (1-200 kWh)")
        
        # Standardwerte oder geladene Werte verwenden
        default_charge_eff = loaded_settings.get('battery_efficiency_charge', 0.95) if loaded_settings else 0.95
        default_discharge_eff = loaded_settings.get('battery_efficiency_discharge', 0.95) if loaded_settings else 0.95
        battery_efficiency_charge = st.sidebar.number_input(
            "Ladeeffizienz Batterie (%)", 
            min_value=80.0, 
            max_value=100.0, 
            value=float(default_charge_eff * 100), 
            step=0.1,
            format="%.1f",
            key="ui_battery_efficiency_charge",
            help="Wirkungsgrad beim Laden der Batterie. Typische Werte: 92-98%"
        ) / 100
        battery_efficiency_discharge = st.sidebar.number_input(
            "Entladeeffizienz Batterie (%)", 
            min_value=80.0, 
            max_value=100.0, 
            value=float(default_discharge_eff * 100), 
            step=0.1,
            format="%.1f",
            key="ui_battery_efficiency_discharge",
            help="Wirkungsgrad beim Entladen der Batterie. Typische Werte: 92-98%"
        ) / 100
        # Lade-/Entladeleistung: keine manuelle Eingabe – aus Excel je Kapazität
        st.sidebar.info("🔧 Max. Lade-/Entladeleistung werden je Größe automatisch aus 'Batteriespeicherkosten.xlsm' übernommen.")
        
        # Initialer Ladezustand der Batterie
        default_initial_soc = loaded_settings.get('initial_soc_percent', 50.0) if loaded_settings else 50.0
        initial_soc_percent = st.sidebar.slider(
            "Initialer Ladezustand Batterie (%)", min_value=0, max_value=100, value=int(default_initial_soc), step=1,
            key="ui_initial_soc_percent",
            help="Ladezustand der Batterie zu Beginn der Simulation. 50% ist ein realistischer Standardwert."
        )
        
        # SOC-Grenzen der Batterie
        default_min_soc = loaded_settings.get('min_soc_percent', 10.0) if loaded_settings else 10.0
        default_max_soc = loaded_settings.get('max_soc_percent', 90.0) if loaded_settings else 90.0
        
        col1, col2 = st.sidebar.columns(2)
        with col1:
            min_soc_percent = st.sidebar.slider(
                "Min. Ladezustand (%)", min_value=0, max_value=50, value=int(default_min_soc), step=1,
                key="ui_min_soc_percent",
                help="Minimaler Ladezustand der Batterie. 0% = Keine Untergrenze (nicht empfohlen)."
            )
        with col2:
            max_soc_percent = st.sidebar.slider(
                "Max. Ladezustand (%)", min_value=50, max_value=100, value=int(default_max_soc), step=1,
                key="ui_max_soc_percent",
                help="Maximaler Ladezustand der Batterie. 100% = Keine Obergrenze (nicht empfohlen)."
            )
        
        # KRITISCHE VALIDIERUNG: Min. Ladezustand muss kleiner als Max. Ladezustand sein
        # Verhindert unrealistische Batterie-Konfigurationen
        if min_soc_percent >= max_soc_percent:
            st.error("❌ Minimaler Ladezustand muss kleiner als maximaler Ladezustand sein!")
            st.stop()
        
        # Jährlicher Kapazitätsverlust
        default_capacity_loss = loaded_settings.get('annual_capacity_loss_percent', 1.0) if loaded_settings else 1.0
        annual_capacity_loss_percent = st.sidebar.slider(
            "Jährlicher Kapazitätsverlust (%)", min_value=0.0, max_value=10.0, value=default_capacity_loss, step=0.1,
            key="ui_annual_capacity_loss_percent",
            help="Jährlicher Kapazitätsverlust der Batterie durch Alterung. Der Wert ist frei wählbar und entspricht dem eingegebenen Wert. Typische Werte für moderne Lithium-Ionen-Batterien liegen zwischen 0.5% und 2% pro Jahr."
        )

        # --- Optimierungsparameter ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">📊 Optimierungsparameter</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Standardwerte oder geladene Werte verwenden
        default_min_capacity = loaded_settings.get('min_battery_capacity', 0.0) if loaded_settings else 0.0
        default_max_capacity = loaded_settings.get('max_battery_capacity', 30.0) if loaded_settings else 30.0
        default_step_size = loaded_settings.get('battery_step_size', 1.0) if loaded_settings else 1.0
        
        min_battery_capacity = st.sidebar.number_input(
            "Min. Batteriekapazität (kWh)", 
            min_value=0.0, 
            max_value=200.0, 
            value=default_min_capacity, 
            step=1.0,
            key="ui_min_battery_capacity",
            help="Minimale zu untersuchende Batteriekapazität. Kann gleich der maximalen sein für eine einzelne Größe."
        )
        max_battery_capacity = st.sidebar.number_input(
            "Max. Batteriekapazität (kWh)", 
            min_value=0.0, 
            max_value=200.0, 
            value=default_max_capacity, 
            step=1.0,
            key="ui_max_battery_capacity",
            help="Maximale zu untersuchende Batteriekapazität. Kann gleich der minimalen sein für eine einzelne Größe."
        )
        battery_step_size = st.sidebar.number_input(
            "Schrittweite (kWh)", 
            min_value=0.5, 
            max_value=5.0, 
            value=default_step_size, 
            step=0.5,
            key="ui_battery_step_size",
            help="Schrittweite für die Optimierung. Bei min=max wird die Schrittweite ignoriert."
        )
        
        # KRITISCHE VALIDIERUNG: Min. Batteriekapazität darf nicht größer als Max. sein (gleich ist OK)
        if min_battery_capacity > max_battery_capacity:
            st.error("❌ Minimale Batteriekapazität darf nicht größer als maximale Batteriekapazität sein!")
            st.stop()
        
        # Optimierungskriterium
        default_optimization_criterion = loaded_settings.get('optimization_criterion', 'Deckungsbeitrag III gesamt (Barwert)') if loaded_settings else 'Deckungsbeitrag III gesamt (Barwert)'
        # Stelle sicher, dass nur gültige Optionen verwendet werden
        if default_optimization_criterion not in ['Deckungsbeitrag III gesamt (Barwert)', 'Deckungsbeitrag III gesamt (Nominal)']:
            default_optimization_criterion = 'Deckungsbeitrag III gesamt (Barwert)'
        optimization_criterion = st.sidebar.selectbox(
            "Optimierungskriterium",
            options=[
                "Deckungsbeitrag III gesamt (Barwert)",
                "Deckungsbeitrag III gesamt (Nominal)"
            ],
            index=0 if default_optimization_criterion == 'Deckungsbeitrag III gesamt (Barwert)' else 1,
            help="Wählen Sie das Kriterium für die optimale Batteriespeichergröße. DB III gesamt maximiert den Gesamtgewinn über die Projektlaufzeit."
        )

        # --- Strompreise ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">💰 Strompreise</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Standardwerte oder geladene Werte verwenden
        default_grid_price = loaded_settings.get('price_grid_per_kwh', 0.30) if loaded_settings else 0.30
        default_feed_in_price = loaded_settings.get('price_feed_in_per_kwh', 0.08) if loaded_settings else 0.08
        
        price_grid_per_kwh = st.sidebar.number_input(
            "Strombezugspreis (ct/kWh)", min_value=10.0, max_value=50.0, value=default_grid_price * 100, step=0.5,
            key="ui_price_grid_per_kwh"
        ) / 100
        price_feed_in_per_kwh = st.sidebar.number_input(
            "Einspeisevergütung (ct/kWh)", min_value=5.0, max_value=20.0, value=default_feed_in_price * 100, step=0.5,
            key="ui_price_feed_in_per_kwh"
        ) / 100

        # --- Finanzielle Parameter ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">📈 Finanzielle Parameter</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Standardwerte oder geladene Werte verwenden
        default_lifetime = loaded_settings.get('project_lifetime_years', 20) if loaded_settings else 20
        default_discount = loaded_settings.get('discount_rate', 0.05) if loaded_settings else 0.05
        default_db_interest = loaded_settings.get('project_interest_rate_db', 0.03) if loaded_settings else 0.03
        
        project_lifetime_years = st.sidebar.slider(
            "Projektlaufzeit (Jahre)", min_value=5, max_value=30, value=default_lifetime, step=1,
            key="ui_project_lifetime_years"
        )
        discount_rate = st.sidebar.slider(
            "Diskontierungsrate (%)", min_value=0.0, max_value=10.0, value=default_discount * 100, step=0.1,
            key="ui_discount_rate"
        ) / 100
        project_interest_rate_db = st.sidebar.slider(
            "Projektzins für DB Kalkulation (%)", min_value=0.0, max_value=10.0, value=default_db_interest * 100, step=0.1,
            key="ui_project_interest_rate_db"
        ) / 100

        # --- Einstellungen speichern ---
        st.sidebar.markdown("""
        <div style="background-color: #2D2D2D; padding: 12px; border-radius: 8px; border-left: 4px solid #FF6B35; margin: 10px 0;">
            <h4 style="color: #FF6B35; margin: 0; font-size: 1.1em;">💾 Aktuelle Einstellungen speichern</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Aktuelle Einstellungen sammeln
        current_settings = {
            'load_profile_type': selected_profile_type,
            'bundesland': selected_bundesland,
            'bundesland_code': bundesland_code, # Hinzugefügt für Excel-Export
            'selected_year': selected_year,
            'annual_consumption_kwh': annual_consumption_kwh,  # Verwende den eingegebenen Wert
            'number_of_persons': number_of_persons,
            'pv_system_size_kwp': pv_system_size_kwp,
            'pv_data_source_link': st.session_state.get('pv_data_source_link_input', '').strip(),
            'battery_efficiency_charge': battery_efficiency_charge,
            'battery_efficiency_discharge': battery_efficiency_discharge,
            'initial_soc_percent': initial_soc_percent,
            'min_soc_percent': min_soc_percent,
            'max_soc_percent': max_soc_percent,
            'annual_capacity_loss_percent': annual_capacity_loss_percent,
            'min_battery_capacity': min_battery_capacity,
            'max_battery_capacity': max_battery_capacity,
            'battery_step_size': battery_step_size,
            'price_grid_per_kwh': price_grid_per_kwh,
            'price_feed_in_per_kwh': price_feed_in_per_kwh,
            'project_lifetime_years': project_lifetime_years,
            'discount_rate': discount_rate,
            'project_interest_rate_db': project_interest_rate_db,
            'optimization_criterion': optimization_criterion,
            'consumption_file': uploaded_consumption_file.name if uploaded_consumption_file else None,
            'pv_file': uploaded_pv_file.name if uploaded_pv_file else None
        }
        
        # Speichern Dialog
        config_name = st.sidebar.text_input("Konfigurationsname:", placeholder="z.B. Standard-Einstellungen")
        col1, col2 = st.sidebar.columns(2)
        
        with col1:
            if st.button("💾 Speichern"):
                if config_name.strip():
                    config_to_save = current_settings.copy()
                    config_to_save['pv_data_source_link'] = current_pv_link.strip()
                    if settings_manager.save_configuration(config_name, config_to_save):
                        st.sidebar.success(f"Konfiguration '{config_name}' gespeichert!")
                        st.rerun()
                    else:
                        st.sidebar.error("Speichern fehlgeschlagen")
                else:
                    st.sidebar.error("Bitte geben Sie einen Namen ein")
        
        with col2:
            if st.button("🔄 Aktualisieren"):
                if 'config_selector' in st.session_state:
                    selected_config = st.session_state.config_selector
                    if selected_config != "-- Neue Konfiguration --":
                        config_to_save = current_settings.copy()
                        config_to_save['pv_data_source_link'] = current_pv_link.strip()
                        if settings_manager.update_configuration(selected_config, config_to_save):
                            st.sidebar.success(f"Konfiguration '{selected_config}' aktualisiert!")
                            st.rerun()
                        else:
                            st.sidebar.error("Aktualisierung fehlgeschlagen")
                    else:
                        st.sidebar.error("Bitte wählen Sie eine Konfiguration aus")

        
        if st.sidebar.button("🚀 Vollständige Analyse starten", use_container_width=True):
            if not data_available or uploaded_consumption_file is None or uploaded_pv_file is None:
                st.sidebar.error("Bitte laden Sie zuerst Verbrauchs- und PV-Daten hoch.")
                st.stop()
                    
            cleanup_keys = [
                'persist_energy_month_select',
                'persist_energy_view_mode',
                'time_series_timerange'
            ]
            for key in cleanup_keys:
                if key in st.session_state:
                    del st.session_state[key]

            day_keys = [key for key in st.session_state.keys() if key.startswith('persist_day_select_')]
            for key in day_keys:
                del st.session_state[key]

            legacy_plot_keys = [key for key in st.session_state.keys() if key.startswith('plot_cache_')]
            for key in legacy_plot_keys:
                del st.session_state[key]

            pv_name = uploaded_pv_file.name or "pv_data"
            pv_extension = pv_name.split('.')[-1].lower() if '.' in pv_name else 'xlsx'

            analysis_params = {
                    'selected_year': selected_year,
                'annual_consumption_kwh': annual_consumption_kwh,
                'number_of_persons': number_of_persons,
                'bundesland_code': bundesland_code,
                'load_profile_type': selected_profile_type,
                'min_battery_capacity': min_battery_capacity,
                'max_battery_capacity': max_battery_capacity,
                'battery_step_size': battery_step_size,
                'battery_efficiency_charge': battery_efficiency_charge,
                'battery_efficiency_discharge': battery_efficiency_discharge,
                'initial_soc_percent': initial_soc_percent,
                'min_soc_percent': min_soc_percent,
                'max_soc_percent': max_soc_percent,
                'annual_capacity_loss_percent': annual_capacity_loss_percent,
                'price_grid_per_kwh': price_grid_per_kwh,
                'price_feed_in_per_kwh': price_feed_in_per_kwh,
                'project_lifetime_years': project_lifetime_years,
                'discount_rate': discount_rate,
                'project_interest_rate_db': project_interest_rate_db,
                'optimization_criterion': optimization_criterion,
                    'pv_system_size_kwp': pv_system_size_kwp
                }

            payload = {
                'params': analysis_params,
                'current_settings': current_settings,
                'consumption_file': uploaded_consumption_file.getvalue(),
                'consumption_name': uploaded_consumption_file.name,
                'pv_file': uploaded_pv_file.getvalue(),
                'pv_name': pv_name,
                'pv_extension': pv_extension
            }

            enqueue_analysis(payload)
            st.info("Analyse gestartet – während der Berechnung werden Status und Fortschritt angezeigt.")
            st.rerun()

else:
    st.sidebar.info("✅ Analyse abgeschlossen – nutze die Auswertung unten. Für eine neue Analyse bitte auf 'Analyse zurücksetzen' klicken.")
    if st.sidebar.button("🔄 Analyse zurücksetzen", use_container_width=True):
        reset_analysis_results()
        st.session_state['analysis_state'] = 'idle'
        st.session_state['analysis_error'] = None
        st.session_state['analysis_completed'] = False
        # Zurücksetzen der Konfiguration, damit aktuelle UI-Eingaben verwendet werden
        if 'config_applied' in st.session_state:
            del st.session_state['config_applied']
        if 'applied_config_name' in st.session_state:
            del st.session_state['applied_config_name']
        st.rerun()

# ==== Persistente Energiefluss-Ansicht (Plotly Range Controls) ====
if analysis_ready and 'optimal_sim_result' in st.session_state:
    try:
        st.markdown("""
        <div style="background-color: #2D2D2D; padding: 20px; border-radius: 12px; border-left: 6px solid #FF6B35; margin: 20px 0;">
            <h3 style="color: #FF6B35; margin: 0; font-size: 1.5em;">⏰ Energieflüsse – Interaktive Zeitfenster-Auswahl</h3>
        </div>
        """, unsafe_allow_html=True)
    
        ts_df = st.session_state['optimal_sim_result']["time_series_data"]
        opt_capacity = st.session_state.get('optimal_capacity', 0.0)
        min_timestamp = ts_df.index.min()
        max_timestamp = ts_df.index.max()
        min_date = min_timestamp.date()
        max_date = max_timestamp.date()

        if 'energy_flow_day_selector' not in st.session_state:
            st.session_state['energy_flow_day_selector'] = min_date
        if 'energy_flow_day_prev' not in st.session_state:
            st.session_state['energy_flow_day_prev'] = st.session_state['energy_flow_day_selector']
        if 'energy_flow_force_day' not in st.session_state:
            st.session_state['energy_flow_force_day'] = False

        date_col = st.columns([1])[0]
        with date_col:
            st.date_input(
                "Tag auswählen",
                key="energy_flow_day_selector",
                min_value=min_date,
                max_value=max_date
            )
            if st.session_state['energy_flow_day_selector'] != st.session_state['energy_flow_day_prev']:
                st.session_state['energy_flow_force_day'] = True

        st.session_state['energy_flow_day_prev'] = st.session_state['energy_flow_day_selector']
        selected_day = st.session_state['energy_flow_day_selector']

        global_start_date = min_timestamp.strftime('%Y-%m-%d')
        global_end_date = max_timestamp.strftime('%Y-%m-%d')
        st.session_state['time_series_timerange'] = {'start': global_start_date, 'end': global_end_date}
        energy_flow_fig = plot_energy_flows_for_period(ts_df, global_start_date, global_end_date, opt_capacity)
        st.session_state['energy_flow_fig_full'] = energy_flow_fig

        if st.session_state['energy_flow_force_day']:
            view_start = datetime.combine(selected_day, datetime.min.time())
            view_end = view_start + timedelta(days=1)
        else:
            view_start = datetime(min_timestamp.year, min_timestamp.month, 1)
            next_month = view_start + pd.DateOffset(months=1)
            view_end = min(next_month - pd.DateOffset(days=1), max_timestamp)

        view_data = ts_df.loc[view_start:view_end]
        if not view_data.empty:
            y_min, y_max = compute_energy_axis_range(view_data)
            energy_flow_fig.update_yaxes(range=[y_min, y_max], row=1, col=1)
        energy_flow_fig.update_yaxes(range=[0, 100], row=2, col=1)

        energy_flow_fig.update_layout(
            xaxis=dict(range=[view_start, view_end], rangeslider=dict(range=[min_timestamp, max_timestamp])),
            xaxis2=dict(range=[view_start, view_end])
        )

        st.plotly_chart(energy_flow_fig, use_container_width=True, key="energy_flow_fig_full_persist")
    except Exception as e:
        # Fehler abfangen - verhindert Reset der Seite
        st.error(f"⚠️ Fehler in der Energiefluss-Ansicht: {str(e)}")

# ==== Persistente Ergebnis-Abschnitte (KPIs, Tabellen, Kurven) ====
# Nur rendern wenn Analyse abgeschlossen ist (verhindert Rendering bei Parameteränderungen)
if (analysis_ready 
    and 'optimization_results' in st.session_state 
    and st.session_state['optimization_results']):
    try:
        df_results = pd.DataFrame(st.session_state['optimization_results'])
        # Zusammenfassung aus Session
        summary = st.session_state.get('results_summary', {})
        if summary:
            st.markdown("""
            <div style="background-color: #2D2D2D; padding: 20px; border-radius: 12px; border-left: 6px solid #FF6B35; margin: 20px 0;">
                <h3 style="color: #FF6B35; margin: 0; font-size: 1.5em;">🏆 Konsolidierte Ergebnisse (persistente Ansicht)</h3>
            </div>
            """, unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Optimale Kapazität", f"{summary.get('optimal_capacity', float('nan')):.1f} kWh")
                if summary.get('investment_cost') is not None:
                    st.metric("Anfangsinvestition", f"{summary['investment_cost']:.0f} €")
            with col2:
                st.metric("Optimierungskriterium", summary.get('criterion_name', '—'))
                if summary.get('criterion_value') is not None:
                    st.metric("Optimaler Wert", f"{summary['criterion_value']:.2f}")
            with col3:
                if summary.get('payback_period_years') is not None:
                    st.metric("Amortisationszeit", f"{summary['payback_period_years']:.1f} Jahre")
                if summary.get('roi_percentage') is not None:
                    st.metric(f"ROI über {summary.get('project_lifetime_years', 0):.0f} Jahre", f"{summary['roi_percentage']:.1f}%")
            with col4:
                irr = summary.get('irr_percentage')
                st.metric("IRR (Interner Zinsfuß)", f"{irr:.2f}%" if irr is not None else "—")

        # Detaillierte Optimierungsergebnisse (Tabelle)
        st.markdown("#### 📋 Detaillierte Optimierungsergebnisse")
        display_columns = [
            'battery_capacity_kwh',
            'investment_cost',
            'autarky_rate',
            'self_consumption_rate',
            'self_consumption_kwh',
            'pv_generation_kwh',
            'total_consumption_kwh',
            'effective_annual_energy_cost',
            'grid_import_cost',
            'grid_export_revenue',
            'contribution_margin_3',
            'payback_period_years',
            'roi_percentage',
            'npv'
        ]
        available_columns = [c for c in display_columns if c in df_results.columns]
        if available_columns:
            df_display = df_results[available_columns].copy()
            rename_map = {
                'battery_capacity_kwh': 'Kapazität (kWh)',
                'investment_cost': 'Investition (€)',
                'autarky_rate': 'Autarkiegrad',
                'self_consumption_rate': 'Eigenverbrauchsquote',
                'self_consumption_kwh': 'Eigenverbrauch (kWh)',
                'pv_generation_kwh': 'PV-Erzeugung (kWh)',
                'total_consumption_kwh': 'Gesamtverbrauch (kWh)',
                'effective_annual_energy_cost': 'Effektive jährl. Energiekosten (€)',
                'grid_import_cost': 'Strombezugskosten (€)',
                'grid_export_revenue': 'Einspeisevergütung (€)',
                'contribution_margin_3': 'DB III (jährlich €)',
                'payback_period_years': 'Amortisation (Jahre)',
                'roi_percentage': 'ROI (%)',
                'npv': 'NPV (€)'
            }
            df_display.columns = [rename_map.get(c, c) for c in available_columns]
            st.dataframe(df_display, use_container_width=True)
        # Einzelne Kurvenanzeigen aus analysis.py
        st.markdown("#### 📈 Technische Kennzahlen")
        st.plotly_chart(
            plot_technical_optimization_curve(st.session_state['optimization_results']),
            use_container_width=True,
            key="persist_technical_optimization_fig"
        )
        st.markdown("#### 💶 Wirtschaftliche Kennzahlen")
        st.plotly_chart(
            plot_economic_optimization_curve(st.session_state['optimization_results']),
            use_container_width=True,
            key="persist_economic_optimization_fig"
        )
        st.markdown("#### 🔄 Energieflüsse (Optimierung)")
        st.plotly_chart(
            plot_energy_flows_optimization(st.session_state['optimization_results']),
            use_container_width=True,
            key="persist_energy_flows_optimization_fig"
        )
        
        # Sankey-Diagramme: Mit und ohne Batterie - nebeneinander
        if ('optimal_sim_result' in st.session_state and 
            'cost_comparison_details' in st.session_state and
            st.session_state['optimal_sim_result'] and
            st.session_state['cost_comparison_details']):
            try:
                st.markdown("#### 🔄 Energieflüsse - Sankey-Diagramme")
                
                # Erstelle zwei Spalten für nebeneinander Anzeige
                col1, col2 = st.columns(2)
                
                with col1:
                    # Diagramm ohne Batterie
                    cost_details = st.session_state['cost_comparison_details']
                    kpis_no_battery = {
                        "total_grid_import_kwh": cost_details.get('grid_import_no_battery'),
                        "total_grid_export_kwh": cost_details.get('grid_export_no_battery'),
                        "total_direct_self_consumption_kwh": cost_details.get('self_consumption_no_battery')
                    }
                    total_consumption = st.session_state['optimal_sim_result']["kpis"].get("total_consumption_kwh", 0)
                    total_pv_generation = st.session_state['optimal_sim_result']["kpis"].get("total_pv_generation_kwh", 0)
                    
                    sankey_without = plot_sankey_diagram_no_battery(kpis_no_battery, total_consumption, total_pv_generation)
                    if sankey_without:
                        st.markdown("##### Ohne Batteriespeicher")
                        st.plotly_chart(
                            sankey_without,
                            use_container_width=True,
                            key="sankey_without_battery"
                        )
                
                with col2:
                    # Diagramm mit Batterie
                    sankey_with = plot_sankey_diagram(st.session_state['optimal_sim_result']["kpis"])
                    if sankey_with and 'figure' in sankey_with:
                        st.markdown("##### Mit Batteriespeicher")
                        st.plotly_chart(
                            sankey_with['figure'],
                            use_container_width=True,
                            key="sankey_with_battery"
                        )
                        
            except Exception as e:
                st.warning(f"Sankey-Diagramme konnten nicht angezeigt werden: {str(e)}")
    except Exception:
        pass

excel_keys = [
    'optimization_results',
    'optimal_sim_result',
    'variable_tariff_result',
    'current_settings',
    'calculation_details',
    'cost_comparison_details',
    'original_consumption_series',
    'scaled_pv_generation_series',
    'battery_cost_curve',
]

if analysis_ready and all(k in st.session_state for k in excel_keys):
    st.markdown("""
    <div style="background-color: #2D2D2D; padding: 20px; border-radius: 12px; border-left: 6px solid #FF6B35; margin: 20px 0;">
        <h3 style="color: #FF6B35; margin: 0; font-size: 1.5em;">📤 Excel-Export</h3>
    </div>
    """, unsafe_allow_html=True)

    excel_export_section(
        st.session_state['optimization_results'],
        st.session_state['optimal_sim_result'],
        st.session_state['variable_tariff_result'],
        st.session_state['current_settings'],
        st.session_state['calculation_details'],
        st.session_state['cost_comparison_details'],
        st.session_state['original_consumption_series'],
        st.session_state['scaled_pv_generation_series'],
        st.session_state['battery_cost_curve'],
        st.session_state.get('results_summary'),
    )