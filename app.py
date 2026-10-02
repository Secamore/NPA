"""
NPA Depot Material Balance Portal — Professional Edition (v3.4)
===============================================================
Multi-role platform: Managers · Companies · Staff
v3.4:
  • PDF rebuilt: embedded Unicode fonts (crisp, searchable text), 300-dpi charts, logo
    re-sampled at high resolution (or vector SVG), wrapped tables, page numbers
  • Reports now include Daily, Weekly, Monthly and Yearly breakdowns (screen, PDF, Excel)
    and can be scoped to one specific day / week / month / year
  • Facility auto-detection: depot, refinery, import terminal, jetty, retail or transport is
    inferred from file names, sheet names, column headers, node names and products; the
    report wording, KPIs, tolerance and written analysis adapt to what was detected
  • Public-internet ready: NPA_DB_PATH / NPA_ASSETS_DIR / NPA_MANAGER_* env vars, no default
    credentials shown or accepted off-localhost, Docker + config files (see DEPLOY.md)
v3.3:
  • Universal ingestion + cleaning engine: any xlsx/xlsm/xls/csv/tsv/txt, auto header
    detection, titles/totals/footers removed, text & accounting numbers, European
    decimals, Excel serial dates, per-row / per-header units (L, KL, m3, gal, bbl),
    product/depot standardisation, de-duplication, full cleaning report + cleaned export
  • Fixes: blank closing stock was counted as a 100% loss (now "Unverified"); values were
    converted to litres but labelled in the source unit; exceptions counted unverifiable
    rows; "Loss Volume" actually netted gains; company sign-up could leave orphan records;
    login revealed which accounts exist
v3.2:
  • Official NPA logo pipeline (assets/npa_logo.png) used in the login scene,
    app header, sidebar, watermark background and PDF reports
  • Redesigned login + in-app backgrounds (layered scene, grid, logo watermark)
  • New sections: Trends, My Profile, Help & Template, Submissions Log,
    company approval/edit, user approve/delete, staff approval by company
  • Admin pages (users, companies, audit) no longer require a data upload
  • Login lock-out after repeated failures, env-configurable invite code
  • Fixes: PDF ellipsis crash, short-alias column mis-detection, % pivot totals
v3.1:
  • Auto DB migration, defensive column guards, PDF export (fpdf2)

Run:  streamlit run app.py
Deps: streamlit pandas numpy plotly openpyxl xlrd fpdf2 matplotlib pillow

Logo:  put the official file at  ./assets/npa_logo.png  (or .svg/.jpg/.webp),
       or set env NPA_LOGO_URL, or (Manager) upload it from the sidebar ▸ Branding.
"""

from __future__ import annotations

import base64
import csv
import datetime as dt
import hashlib
import html
import io
import json
import math
import os
import re
import secrets
import sqlite3
import unicodedata
import warnings
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# PDF (optional but recommended)
try:
    from fpdf import FPDF
    _HAS_FPDF = True
except Exception:
    FPDF = None
    _HAS_FPDF = False

try:
    pd.set_option("mode.string_storage", "python")
except Exception:
    pass
try:
    pd.set_option("future.infer_string", False)
except Exception:
    pass


# ==========================================================================
# Config
# ==========================================================================
st.set_page_config(
    page_title="NPA Depot Material Balance Portal",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "NPA Depot Material Balance Portal · Enterprise Edition v3.4"},
)

APP_DIR = Path(__file__).parent.resolve()
# On a public host point these at a PERSISTENT disk/volume (see DEPLOY.md), otherwise the
# user database and the saved logo are wiped every time the container restarts.
DB_PATH = Path(os.environ.get("NPA_DB_PATH") or (APP_DIR / "npa_portal.db"))
ASSETS_DIR = Path(os.environ.get("NPA_ASSETS_DIR") or (APP_DIR / "assets"))
try:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
except Exception:
    pass
DEFAULT_MANAGER_EMAIL = os.environ.get("NPA_MANAGER_EMAIL", "").strip() or "manager@npa.gov.gh"
DEFAULT_MANAGER_PASSWORD = "ChangeMe123!"      # local first-run only — never used on a public host
_DEV_INVITE_CODE = "NPA-MGR-2026"              # local first-run only


def _is_local_request() -> bool:
    """True only when the browser reached the app via localhost (developer machine)."""
    try:
        host = (st.context.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]").lower()
        return host in ("localhost", "127.0.0.1", "::1")
    except Exception:
        return False


def manager_invite_code() -> str:
    """Manager self-registration code. Public hosts must set NPA_MANAGER_INVITE; without it
    manager sign-up is disabled (the dev default is public in the source code)."""
    env = os.environ.get("NPA_MANAGER_INVITE", "").strip()
    if env:
        return env
    return _DEV_INVITE_CODE if _is_local_request() else ''
MAX_FAILED_LOGINS = 5
LOCKOUT_MINUTES = 15
LOGO_NAMES = ("npa_logo", "logo", "npa")
LOGO_EXTS = ("png", "svg", "webp", "jpg", "jpeg")
_MIME = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
         "webp": "image/webp", "svg": "image/svg+xml"}

NAVY, NAVY2, GOLD = "#123b63", "#0a2945", "#d7a72f"
GREEN, AMBER, RED, BLUE, GREY = "#157347", "#b4690e", "#b42318", "#2f6fb0", "#667085"
PURPLE, TEAL, CYAN, ORANGE, PINK = "#8b5cf6", "#14b8a6", "#0ea5e9", "#f97316", "#f43f5e"
LINE, BG = "#dde4ec", "#f4f7fb"

ROLES = {
    "manager": {"label": "Manager", "icon": "🛡️", "color": GOLD},
    "company": {"label": "Company", "icon": "🏢", "color": BLUE},
    "staff":   {"label": "Staff",   "icon": "👤", "color": GREEN},
}
PLOTLY_TEMPLATE = "plotly_white"
PRODUCT_COLORS = px.colors.qualitative.Bold


# ==========================================================================
# DEPARTMENT REGISTRY
# ==========================================================================
DEPARTMENTS: Dict[str, Dict] = {
    "DEPOT": {
        "label": "Depot & Terminal Operations", "short": "Depot Ops", "icon": "🏭",
        "tol": 0.35, "color": BLUE,
        "scope": "Bulk storage depots — receipts, issues, dips, ullage, gains/losses",
        "kpis": ["Depot throughput", "Gain/Loss %", "Stock cover days", "Dip compliance"],
        "kind": "material",
    },
    "IMPORT": {
        "label": "Imports / Vessel Discharge", "short": "Imports", "icon": "🚢",
        "tol": 0.25, "color": CYAN,
        "scope": "Cargo discharge — out-turn vs Bill of Lading, shore-tank receipt",
        "kpis": ["Cargo out-turn %", "Discharge variance", "Demurrage", "Cargo count"],
        "kind": "material",
    },
    "MARINE": {
        "label": "Marine & Jetty Operations", "short": "Marine", "icon": "⚓",
        "tol": 0.30, "color": "#06b6d4",
        "scope": "Jetty loading arms, bunkering, coastal movements, ship-shore reconciliation",
        "kpis": ["Jetty utilisation", "Bunkering volume", "Line-fill variance", "Berth hours"],
        "kind": "material",
    },
    "REFINERY": {
        "label": "Refining & Local Production", "short": "Refinery", "icon": "🔥",
        "tol": 0.60, "color": ORANGE,
        "scope": "Local refinery yield, feedstock intake, product slate",
        "kpis": ["Refinery yield %", "Feedstock vs output", "Product slate mix", "Capacity use"],
        "kind": "material",
    },
    "RETAIL": {
        "label": "Retail & OMC / Consumer Services", "short": "Retail", "icon": "⛽",
        "tol": 0.45, "color": GREEN,
        "scope": "Retail outlets, OMC allocations, pump sales, under-delivery",
        "kpis": ["Outlet sales", "Allocation vs lift", "Under-delivery %", "Outlets reporting"],
        "kind": "material",
    },
    "TRANSPORT": {
        "label": "Transport & Pipeline Logistics", "short": "Transport", "icon": "🚛",
        "tol": 0.40, "color": AMBER,
        "scope": "Bulk road transport fleet, pipeline throughput, transit losses",
        "kpis": ["Transit loss %", "Truck turnaround", "Pipeline uptime", "Volume hauled"],
        "kind": "material",
    },
    "QUALITY": {
        "label": "Quality Assurance & Laboratory", "short": "Quality", "icon": "🧪",
        "tol": 0.50, "color": PURPLE,
        "scope": "Product specs, density/temperature correction, off-spec quarantine",
        "kpis": ["Off-spec rate", "Tests performed", "Volume correction factor", "Quarantine volume"],
        "kind": "cross",
    },
    "FINANCE": {
        "label": "Finance & Revenue Assurance", "short": "Finance", "icon": "💰",
        "tol": 0.50, "color": TEAL,
        "scope": "Levies, margins, UPPF, revenue-at-risk from volume losses",
        "kpis": ["Levy collected", "Revenue at risk", "Effective levy/L", "Collection rate"],
        "kind": "cross",
    },
    "COMPLIANCE": {
        "label": "Compliance & Enforcement", "short": "Compliance", "icon": "🛡️",
        "tol": 0.50, "color": RED,
        "scope": "Inspections, licences, penalties, regulatory breaches",
        "kpis": ["Compliance rate", "Open findings", "Penalties", "Licences valid"],
        "kind": "cross",
    },
    "HSE": {
        "label": "Health, Safety & Environment", "short": "HSE", "icon": "🚨",
        "tol": 0.50, "color": PINK,
        "scope": "Incidents, spills, product loss to environment, downtime",
        "kpis": ["Incidents", "Volume spilled", "Downtime hrs", "LTIFR"],
        "kind": "cross",
    },
    "IT": {
        "label": "ICT / Data Integrity", "short": "ICT", "icon": "💻",
        "tol": 0.50, "color": GREY,
        "scope": "Dip-capture completeness, late submissions, system uptime, data quality",
        "kpis": ["Data completeness", "Late submissions", "Negative stock flags", "Uptime"],
        "kind": "cross",
    },
}

DEPOT_KEYWORD_TO_DEPT = {
    "vessel":   "IMPORT", "mt ":       "IMPORT", "mv ":    "IMPORT",
    "cargo":    "IMPORT", "discharge": "IMPORT", "b/l":    "IMPORT",
    "jetty":    "MARINE", "berth":     "MARINE", "bunker": "MARINE",
    "refinery": "REFINERY", "ref ":    "REFINERY", "crude":  "REFINERY",
    "retail":   "RETAIL", "outlet":    "RETAIL", "station":"RETAIL",
    "pump":     "RETAIL", "omc":       "RETAIL",
    "pipeline": "TRANSPORT", "truck":  "TRANSPORT", "fleet":  "TRANSPORT",
    "road":     "TRANSPORT", "tanker": "TRANSPORT",
    "bost":     "DEPOT", "tema":       "DEPOT", "takoradi":"DEPOT",
    "kumasi":   "DEPOT", "tamale":     "DEPOT", "buipe":   "DEPOT",
    "depot":    "DEPOT", "terminal":   "DEPOT",
}

DEPARTMENT_ORDER = ["DEPOT", "IMPORT", "MARINE", "REFINERY", "RETAIL",
                    "TRANSPORT", "QUALITY", "FINANCE", "COMPLIANCE", "HSE", "IT"]


def infer_department(name: str) -> str:
    n = str(name).lower()
    for key, dept in DEPOT_KEYWORD_TO_DEPT.items():
        if key in n:
            return dept
    return "DEPOT"


def dept_meta(code: str) -> Dict:
    return DEPARTMENTS.get(code, {"label": code, "short": code, "icon": "•",
                                   "tol": 0.50, "color": GREY, "kind": "material",
                                   "scope": "", "kpis": []})


# ==========================================================================
# Fallback background SVG
# ==========================================================================
_FALLBACK_SVG = """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1920 1080' preserveAspectRatio='xMidYMid slice'>
<defs>
<linearGradient id='sky' x1='0' y1='0' x2='0' y2='1'>
<stop offset='0%' stop-color='#041428'/><stop offset='40%' stop-color='#0e3a5c'/>
<stop offset='75%' stop-color='#4a7ba8'/><stop offset='100%' stop-color='#c8a06a'/></linearGradient>
<linearGradient id='sea' x1='0' y1='0' x2='0' y2='1'>
<stop offset='0%' stop-color='#0a2945'/><stop offset='100%' stop-color='#031628'/></linearGradient>
<linearGradient id='structure' x1='0' y1='0' x2='0' y2='1'>
<stop offset='0%' stop-color='#0a0f18'/><stop offset='100%' stop-color='#000000'/></linearGradient>
<radialGradient id='flare'>
<stop offset='0%' stop-color='#ffee88'/><stop offset='30%' stop-color='#ffaa33' stop-opacity='0.9'/>
<stop offset='100%' stop-color='#ff3300' stop-opacity='0'/></radialGradient>
<radialGradient id='sun'>
<stop offset='0%' stop-color='#ffb347' stop-opacity='0.9'/>
<stop offset='100%' stop-color='#ffb347' stop-opacity='0'/></radialGradient>
<linearGradient id='seaReflect' x1='0' y1='0' x2='0' y2='1'>
<stop offset='0%' stop-color='#ff8c42' stop-opacity='0.35'/>
<stop offset='100%' stop-color='#ff8c42' stop-opacity='0'/></linearGradient>
<filter id='glow'><feGaussianBlur stdDeviation='4'/></filter>
</defs>
<rect width='1920' height='1080' fill='url(#sky)'/>
<circle cx='1380' cy='720' r='180' fill='url(#sun)'/>
<circle cx='1380' cy='720' r='42' fill='#ffcc66' opacity='0.85'/>
<g fill='#fff' opacity='0.5'>
<circle cx='180' cy='80' r='1.2'/><circle cx='420' cy='140' r='0.9'/>
<circle cx='780' cy='60' r='1.4'/><circle cx='1100' cy='180' r='1.1'/>
<circle cx='1650' cy='150' r='0.9'/><circle cx='300' cy='220' r='1'/>
<circle cx='950' cy='240' r='0.7'/><circle cx='1250' cy='300' r='1'/>
<circle cx='1800' cy='200' r='0.9'/></g>
<g fill='#1a3a5a' opacity='0.35'>
<ellipse cx='400' cy='320' rx='280' ry='40'/>
<ellipse cx='1000' cy='280' rx='340' ry='45'/>
<ellipse cx='1550' cy='360' rx='300' ry='38'/></g>
<rect y='800' width='1920' height='280' fill='url(#sea)'/>
<ellipse cx='1380' cy='860' rx='420' ry='60' fill='url(#seaReflect)'/>
<g fill='url(#structure)'>
<rect x='500' y='620' width='900' height='30'/>
<rect x='700' y='480' width='380' height='140'/>
<rect x='720' y='440' width='340' height='40'/>
<polygon points='1000,300 950,620 1050,620'/>
<polygon points='1000,300 970,620 1030,620' fill='#0f1622'/>
<rect x='980' y='280' width='40' height='30'/>
<rect x='1380' y='330' width='18' height='320'/>
<rect x='1370' y='315' width='38' height='20'/>
<circle cx='620' cy='590' r='55' fill='#0a0f18'/>
<circle cx='620' cy='590' r='42' fill='none' stroke='#1a2536' stroke-width='3'/>
<text x='620' y='600' text-anchor='middle' fill='#1a2536' font-family='sans-serif' font-size='28' font-weight='bold'>H</text>
<rect x='560' y='650' width='14' height='160'/><rect x='1320' y='650' width='14' height='160'/>
<rect x='900' y='650' width='14' height='160'/><rect x='1050' y='650' width='14' height='160'/>
<line x1='560' y1='680' x2='900' y2='780' stroke='#0a0f18' stroke-width='5'/>
<line x1='900' y1='680' x2='560' y2='780' stroke='#0a0f18' stroke-width='5'/>
<line x1='1050' y1='680' x2='1320' y2='780' stroke='#0a0f18' stroke-width='5'/>
<line x1='1320' y1='680' x2='1050' y2='780' stroke='#0a0f18' stroke-width='5'/>
<rect x='500' y='600' width='900' height='20' fill='#0a0f18'/>
<line x1='720' y1='480' x2='640' y2='380' stroke='#0a0f18' stroke-width='6'/>
<line x1='640' y1='380' x2='720' y2='420' stroke='#0a0f18' stroke-width='4'/></g>
<g fill='#ffbb44' opacity='0.9'>
<rect x='730' y='540' width='12' height='8'/><rect x='760' y='540' width='12' height='8'/>
<rect x='790' y='540' width='12' height='8'/><rect x='820' y='540' width='12' height='8'/>
<rect x='880' y='540' width='12' height='8'/><rect x='910' y='540' width='12' height='8'/>
<rect x='940' y='540' width='12' height='8'/><rect x='1000' y='540' width='12' height='8'/>
<rect x='1030' y='540' width='12' height='8'/><rect x='1060' y='540' width='12' height='8'/></g>
<g filter='url(#glow)'>
<ellipse cx='1389' cy='310' rx='22' ry='40' fill='url(#flare)'/>
<ellipse cx='1389' cy='290' rx='10' ry='22' fill='#ffdd66'/></g>
<g stroke='#1a4a6a' stroke-width='1.5' opacity='0.4' fill='none'>
<path d='M0,850 Q80,840 160,850 T320,850 T480,850 T640,850 T800,850 T960,850 T1120,850 T1280,850 T1440,850 T1600,850 T1760,850 T1920,850'/></g>
</svg>"""


@st.cache_data(show_spinner=False)
def _file_data_url(path_str: str, mtime: float) -> str:
    """Base64 data-URL for a file; cached on (path, mtime) so edits are picked up."""
    path = Path(path_str)
    ext = path.suffix.lower().lstrip(".")
    return f"data:{_MIME.get(ext, 'image/png')};base64," + base64.b64encode(path.read_bytes()).decode()


def _find_logo_path() -> Optional[Path]:
    for folder in dict.fromkeys((ASSETS_DIR, APP_DIR / "assets", APP_DIR, APP_DIR / ".streamlit")):
        for stem in LOGO_NAMES:
            for ext in LOGO_EXTS:
                cand = folder / f"{stem}.{ext}"
                if cand.exists():
                    return cand
    return None


def get_logo_url() -> Optional[str]:
    """Official NPA logo as a data-URL (session override > local file > NPA_LOGO_URL)."""
    if st.session_state.get("logo_url_override"):
        return st.session_state["logo_url_override"]
    path = _find_logo_path()
    if path is not None:
        return _file_data_url(str(path), path.stat().st_mtime)
    remote = os.environ.get("NPA_LOGO_URL", "").strip()
    if remote.startswith(("http://", "https://")):
        return remote
    return None


def logo_html(size: int = 54, plate: bool = True, extra_style: str = "") -> str:
    """The real logo on a white plate when available, else the built-in NPA badge."""
    url = get_logo_url()
    if url:
        pad = max(3, size // 12)
        bg = "background:#fff;" if plate else ""
        return (f'<img src="{url}" alt="National Petroleum Authority logo" '
                f'style="width:{size}px;height:{size}px;object-fit:contain;'
                f'border-radius:50%;{bg}padding:{pad}px;box-sizing:border-box;'
                f'box-shadow:0 2px 10px rgba(0,0,0,.28);flex-shrink:0;{extra_style}"/>')
    return (f'<div class="npa-logo" style="width:{size}px;height:{size}px;'
            f'font-size:{max(11, size // 3.6):.0f}px;{extra_style}">NPA</div>')


def get_background_url() -> str:
    candidates = [
        APP_DIR / "background.png", APP_DIR / "background.jpg",
        APP_DIR / "background.jpeg", APP_DIR / "background.webp",
        ASSETS_DIR / "background.png", ASSETS_DIR / "background.jpg",
        ASSETS_DIR / "background.webp", APP_DIR / ".streamlit" / "background.png",
    ]
    for path in candidates:
        if path.exists():
            return _file_data_url(str(path), path.stat().st_mtime)
    return "data:image/svg+xml;base64," + base64.b64encode(_FALLBACK_SVG.encode()).decode()


def current_bg() -> str:
    return st.session_state.get("bg_url_override") or get_background_url()


def _watermark_layer(opacity: float, size: str = "min(46vw, 560px)",
                     position: str = "calc(100% - 4vw) 62%",
                     filter_css: str = "grayscale(.15)") -> str:
    """CSS rule for a fixed, faint logo watermark that sits behind all content."""
    url = get_logo_url()
    if not url:
        return ""
    return f"""
      .stApp::after {{
          content:''; position:fixed; inset:0; pointer-events:none; z-index:1;
          background: url('{url}') {position} / {size} no-repeat;
          opacity:{opacity};
          filter: {filter_css};
      }}"""


def save_logo_permanently(data: bytes, filename: str) -> Path:
    ext = filename.lower().rsplit(".", 1)[-1]
    if ext not in LOGO_EXTS:
        raise ValueError("Logo must be PNG, SVG, WEBP or JPG.")
    ASSETS_DIR.mkdir(exist_ok=True)
    for old_ext in LOGO_EXTS:                      # keep exactly one active logo
        (ASSETS_DIR / f"npa_logo.{old_ext}").unlink(missing_ok=True)
    target = ASSETS_DIR / f"npa_logo.{ext}"
    target.write_bytes(data)
    _file_data_url.clear()
    return target


# ==========================================================================
# Database layer + AUTO-MIGRATION
# ==========================================================================
def _connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def hash_password(password: str, salt: Optional[bytes] = None) -> str:
    if salt is None:
        salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 200_000)
    return f"{salt.hex()}:{dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        s_hex, h_hex = stored.split(":")
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(s_hex), 200_000)
        return secrets.compare_digest(dk.hex(), h_hex)
    except Exception:
        return False


def _first_run_password() -> str:
    """Seed password for the first manager account.
    NPA_MANAGER_PASSWORD wins. On a localhost run the documented default is used; on any
    other host a random password is generated and written ONCE to the server log."""
    env = os.environ.get("NPA_MANAGER_PASSWORD", "")
    if env:
        return env
    if _is_local_request():
        return DEFAULT_MANAGER_PASSWORD
    pw = secrets.token_urlsafe(12)
    print(f"[NPA] First-run manager account created: {DEFAULT_MANAGER_EMAIL} / password: {pw} "
          f"(set NPA_MANAGER_PASSWORD to choose your own; change it after signing in)", flush=True)
    return pw


def init_db() -> None:
    con = _connect()
    cur = con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL,
        code TEXT UNIQUE, contact_email TEXT, contact_phone TEXT, address TEXT,
        status TEXT DEFAULT 'active', created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT UNIQUE NOT NULL,
        staff_id TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('manager','company','staff')),
        company_id INTEGER, depot TEXT, department TEXT, phone TEXT,
        status TEXT DEFAULT 'pending', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        last_login TEXT, failed_attempts INTEGER DEFAULT 0, locked_until TEXT,
        FOREIGN KEY(company_id) REFERENCES companies(id))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_email TEXT, full_name TEXT,
        company_id INTEGER, depot TEXT, department TEXT, filename TEXT,
        row_count INTEGER, throughput REAL, variance REAL, exceptions INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_email TEXT, action TEXT,
        details TEXT, timestamp TEXT DEFAULT CURRENT_TIMESTAMP)""")
    con.commit()

    # ---- AUTO-MIGRATION: add missing columns on older DBs -----------------
    def _column_exists(table: str, column: str) -> bool:
        cur.execute(f"PRAGMA table_info({table})")
        return any(row[1] == column for row in cur.fetchall())

    for col_name, col_type in (("department", "TEXT"),
                               ("depot", "TEXT"),
                               ("phone", "TEXT"),
                               ("failed_attempts", "INTEGER DEFAULT 0"),
                               ("locked_until", "TEXT")):
        if not _column_exists("users", col_name):
            try:
                cur.execute(f"ALTER TABLE users ADD COLUMN {col_name} {col_type}")
                con.commit()
            except sqlite3.OperationalError:
                pass
    # ------------------------------------------------------------------------

    cur.execute("SELECT COUNT(*) FROM companies")
    if cur.fetchone()[0] == 0:
        cur.execute("INSERT INTO companies (name, code, contact_email) VALUES (?,?,?)",
                    ("NPA Head Office", "NPA-HQ", "hq@npa.gov.gh"))
    cur.execute("SELECT COUNT(*) FROM users WHERE role='manager'")
    if cur.fetchone()[0] == 0:
        cur.execute("""INSERT INTO users (email, staff_id, password_hash, full_name,
                                          role, department, status)
                       VALUES (?,?,?,?,?,?,?)""",
                    (DEFAULT_MANAGER_EMAIL, "NPA-0001", hash_password(_first_run_password()),
                     "System Manager", "manager", "DEPOT", "active"))
    con.commit()
    con.close()


def db_query(sql: str, params: Tuple = ()) -> List[sqlite3.Row]:
    con = _connect()
    try:
        return con.execute(sql, params).fetchall()
    finally:
        con.close()


def db_execute(sql: str, params: Tuple = ()) -> int:
    con = _connect()
    try:
        cur = con.execute(sql, params)
        con.commit()
        return cur.lastrowid or cur.rowcount
    finally:
        con.close()


def log_action(email: str, action: str, details: str = "") -> None:
    try:
        db_execute("INSERT INTO audit_log (user_email, action, details) VALUES (?,?,?)",
                   (email, action, details))
    except Exception:
        pass


def get_companies() -> List[sqlite3.Row]:
    return db_query("SELECT * FROM companies ORDER BY name")


def default_manager_unchanged() -> bool:
    """True while the seeded manager account still uses the factory password."""
    rows = db_query("SELECT password_hash FROM users WHERE lower(email)=lower(?)",
                    (DEFAULT_MANAGER_EMAIL,))
    return bool(rows) and verify_password(DEFAULT_MANAGER_PASSWORD, rows[0]["password_hash"])


def record_submission(user: dict, filename: str, rows: int, throughput: float,
                      variance: float, exceptions: int) -> None:
    try:
        db_execute("""INSERT INTO submissions (user_email, full_name, company_id, depot,
                      department, filename, row_count, throughput, variance, exceptions)
                      VALUES (?,?,?,?,?,?,?,?,?,?)""",
                   (user["email"], user["full_name"], user.get("company_id"),
                    user.get("depot") or "", user.get("department") or "", filename,
                    int(rows), float(throughput), float(variance), int(exceptions)))
    except Exception:
        pass


def get_users(company_id: Optional[int] = None) -> List[sqlite3.Row]:
    if company_id is None:
        return db_query("""SELECT u.*, c.name AS company_name FROM users u
                           LEFT JOIN companies c ON c.id=u.company_id
                           ORDER BY u.created_at DESC""")
    return db_query("""SELECT u.*, c.name AS company_name FROM users u
                       LEFT JOIN companies c ON c.id=u.company_id
                       WHERE u.company_id=? ORDER BY u.created_at DESC""", (company_id,))


# ==========================================================================
# Numeric helpers
# ==========================================================================
def _clean_numeric_cell(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(x).strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    return re.sub(r"[^0-9eE\.\-\+]", "", text)


def to_number(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").astype(float)
    cleaned = pd.Series([_clean_numeric_cell(x) for x in s],
                        index=s.index, dtype="object")
    return pd.to_numeric(cleaned, errors="coerce").astype(float)


def detect_units(series: pd.Series) -> str:
    if series is None or len(series) == 0:
        return "L"
    from collections import Counter
    tally = Counter()
    for x in series:
        try:
            if pd.isna(x):
                continue
        except (TypeError, ValueError):
            pass
        u = str(x).strip().upper()
        if u in UNIT_CONVERSIONS:
            tally[u] += 1
    if tally:
        return tally.most_common(1)[0][0]
    return "L"


def _normalize_string_dtypes(frame: pd.DataFrame) -> pd.DataFrame:
    for col in frame.columns:
        try:
            if (not pd.api.types.is_numeric_dtype(frame[col])
                    and not pd.api.types.is_datetime64_any_dtype(frame[col])
                    and not pd.api.types.is_bool_dtype(frame[col])):
                frame[col] = frame[col].astype(object)
        except Exception:
            frame[col] = frame[col].astype(object)
    return frame


# ==========================================================================
# Style
# ==========================================================================
def _grid_url(stroke: str, step: int = 72) -> str:
    """Chart-paper grid (like a bathymetric / depot layout sheet) as a data-URL."""
    svg = (f"<svg xmlns='http://www.w3.org/2000/svg' width='{step}' height='{step}'>"
           f"<path d='M{step} 0H0V{step}' fill='none' stroke='{stroke}' stroke-width='1'/>"
           f"<circle cx='0' cy='0' r='1.6' fill='{stroke}'/></svg>")
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()


def base_css(overlay: float = 0.92) -> str:
    bg = current_bg()
    grid = _grid_url("rgba(18,59,99,0.07)")
    grid_light = _grid_url("rgba(255,255,255,0.06)")
    wm = _watermark_layer(0.055, size="min(52vw, 640px)", position="calc(100% - 3vw) calc(100% - 6vh)",
                          filter_css="grayscale(.25)")
    return f"""
    <style>
      html, body, .stApp {{
          background:
              radial-gradient(900px 520px at 6% -4%, rgba(215,167,47,.16), transparent 62%),
              radial-gradient(1100px 640px at 100% 108%, rgba(18,59,99,.16), transparent 60%),
              linear-gradient(rgba(244,247,251,{overlay}), rgba(244,247,251,{overlay})),
              url('{bg}') center/cover fixed no-repeat;
      }}
      .stApp::before {{
          content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
          background: url('{grid}') 0 0 / 72px 72px repeat;
          -webkit-mask-image: linear-gradient(180deg, rgba(0,0,0,.9), rgba(0,0,0,.25));
                  mask-image: linear-gradient(180deg, rgba(0,0,0,.9), rgba(0,0,0,.25));
      }}
      {wm}
      header[data-testid="stHeader"] {{ background: transparent; }}
      #MainMenu, footer {{ visibility: hidden; }}
      .block-container {{ position:relative; z-index:2; padding-top: 1.2rem;
                          padding-bottom: 2rem; max-width: 1500px; }}
      section[data-testid="stSidebar"] > div:first-child {{
          background: linear-gradient(180deg, #f9fbfe 0%, #e9eff6 100%);
          border-right: 1px solid {LINE};
      }}
      .sb-brand {{ display:flex; align-items:center; gap:12px; padding:10px 4px 14px 4px;
                   border-bottom:2px solid {GOLD}; margin-bottom:14px; }}
      .sb-brand b {{ color:{NAVY2}; font-size:13px; letter-spacing:.3px; display:block; }}
      .sb-brand span {{ color:{GREY}; font-size:11px; }}
      .app-footer {{ margin-top:28px; padding-top:12px; border-top:1px solid {LINE};
                     color:{GREY}; font-size:11.5px; display:flex; align-items:center;
                     gap:10px; justify-content:center; }}

      .npa-header {{
          background: linear-gradient(120deg, {NAVY2} 0%, {NAVY} 70%);
          color: #fff; padding: 18px 26px; border-radius: 14px;
          display: flex; align-items: center; gap: 18px;
          box-shadow: 0 4px 18px rgba(10,41,69,.22); margin-bottom: 20px;
          border-bottom: 3px solid {GOLD};
          background-image:
              url('{grid_light}'),
              linear-gradient(120deg, {NAVY2} 0%, {NAVY} 70%);
          background-size: 72px 72px, auto;
      }}
      .npa-logo {{
          width: 54px; height: 54px; border-radius: 50%;
          border: 2.5px solid {GOLD}; background: rgba(215,167,47,.14);
          display: grid; place-items: center; font-weight: 900;
          font-size: 15px; color: {GOLD}; letter-spacing: .5px; flex-shrink: 0;
      }}
      .npa-header h1 {{ margin: 0; font-size: 19px; color: #fff; font-weight: 800; }}
      .npa-header small {{ opacity: .82; font-size: 12.5px; display: block; margin-top: 2px; }}
      .npa-header .who {{ margin-left: auto; text-align: right; font-size: 12px; line-height: 1.5; }}
      .npa-header .who b {{ color: {GOLD}; }}

      .role-badge {{
          display:inline-flex; align-items:center; gap:6px;
          padding:5px 12px; border-radius:999px; font-size:11.5px;
          font-weight:800; letter-spacing:.4px; text-transform:uppercase;
      }}
      .dept-badge {{
          display:inline-flex; align-items:center; gap:6px;
          padding:4px 11px; border-radius:6px; font-size:11px;
          font-weight:800; letter-spacing:.35px; text-transform:uppercase;
          border:1px solid currentColor;
      }}
      .kpi {{
          background:#fff; border:1px solid #e5eaf0; border-left:4px solid {NAVY};
          border-radius:10px; padding:14px 16px;
          box-shadow:0 1px 4px rgba(16,24,40,.05); height:100%;
      }}
      .kpi.gold {{ border-left-color:{GOLD}; }} .kpi.green {{ border-left-color:{GREEN}; }}
      .kpi.red  {{ border-left-color:{RED};  }} .kpi.amber {{ border-left-color:{AMBER}; }}
      .kpi.purple {{ border-left-color:{PURPLE}; }}
      .kpi.teal {{ border-left-color:{TEAL}; }}
      .kpi .l {{ font-size:11px; color:{GREY}; font-weight:800;
                 text-transform:uppercase; letter-spacing:.55px; }}
      .kpi .v {{ font-size:22px; font-weight:850; margin-top:6px; color:{NAVY2};
                 line-height:1.15; font-variant-numeric:tabular-nums; }}
      .kpi .d {{ font-size:12px; color:{GREY}; margin-top:4px; }}

      .section-title {{ font-size:15px; font-weight:800; color:{NAVY2};
                        margin:8px 0 12px 0; padding-bottom:6px;
                        border-bottom:2px solid {LINE}; }}

      .dept-banner {{
          border-radius:12px; padding:18px 22px; margin-bottom:16px;
          color:#fff; display:flex; align-items:center; gap:16px;
      }}
      .dept-banner h2 {{ margin:0; font-size:18px; font-weight:800; color:#fff; }}
      .dept-banner p {{ margin:4px 0 0 0; font-size:12.5px; opacity:.9; }}
      .dept-banner .icon {{ font-size:34px; }}

      .note {{ padding:12px 14px; background:#eef5fb; border-left:4px solid {NAVY};
               border-radius:8px; font-size:13px; color:{NAVY2}; margin-bottom:10px; }}
      .note.warn {{ background:#fff8ec; border-left-color:{AMBER}; }}
      .note.bad  {{ background:#fdecec; border-left-color:{RED}; }}
      .note.ok   {{ background:#e8f5ee; border-left-color:{GREEN}; }}

      .pill {{ display:inline-block; padding:3px 10px; border-radius:999px;
               font-size:11px; font-weight:800; }}
      .pill.ok {{ background:#e8f5ee; color:{GREEN}; }}
      .pill.warn {{ background:#fff2dd; color:{AMBER}; }}
      .pill.bad {{ background:#fdecec; color:{RED}; }}
      .pill.info {{ background:#eef5fb; color:{NAVY}; }}

      div[data-testid="stMetricValue"] {{ font-size:22px; font-weight:800; }}
      div[data-testid="stDataFrame"] {{ border-radius:10px; overflow:hidden; }}

      .report-shell {{ background:#fff; border:1px solid {LINE};
                       border-radius:12px; padding:28px 32px;
                       box-shadow:0 1px 4px rgba(16,24,40,.05); }}
      .report-shell h2 {{ color:{NAVY2}; margin-top:0;
                          border-bottom:2px solid {GOLD}; padding-bottom:8px; font-size:18px; }}
      .report-shell h3 {{ color:{NAVY}; margin-top:22px; font-size:14px;
                          text-transform:uppercase; letter-spacing:.6px; }}
      .report-shell table {{ border-collapse:collapse; width:100%; margin-top:8px; font-size:12.5px; }}
      .report-shell th {{ background:{NAVY}; color:#fff; padding:8px 10px;
                          text-align:right; font-weight:700; }}
      .report-shell th:first-child {{ text-align:left; }}
      .report-shell td {{ padding:7px 10px; border-bottom:1px solid {LINE};
                          text-align:right; font-variant-numeric:tabular-nums; }}
      .report-shell td:first-child {{ text-align:left; font-weight:600; }}
      .report-shell tr:nth-child(even) td {{ background:#fafbfd; }}
    </style>
    """


def login_css() -> str:
    bg = current_bg()
    grid = _grid_url("rgba(215,167,47,0.08)", 96)
    wm = _watermark_layer(0.07, size="min(60vw, 720px)", position="14% 58%",
                          filter_css="grayscale(1) brightness(1.6) contrast(.9)")
    return f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
      html, body, .stApp {{ background: url('{bg}') center/cover fixed no-repeat !important;
                           font-family:'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif; }}
      .stApp::before {{
          content:''; position:fixed; inset:0; z-index:0; pointer-events:none;
          background:
              radial-gradient(900px 600px at 85% 15%, rgba(215,167,47,.20), transparent 60%),
              radial-gradient(800px 600px at 5% 95%, rgba(14,165,233,.18), transparent 60%),
              url('{grid}') 0 0 / 96px 96px repeat,
              linear-gradient(135deg, rgba(3,16,34,.82) 0%, rgba(8,34,64,.74) 55%, rgba(3,16,34,.88) 100%);
      }}
      {wm}
      .stApp::after {{ mix-blend-mode: screen; }}
      header[data-testid="stHeader"] {{ background: transparent; height:0; }}
      #MainMenu, footer {{ visibility:hidden; }}
      .block-container {{ position:relative; z-index:2; max-width:1240px; padding-top:6vh; padding-bottom:4vh; }}

      /* ---------- floating glow orbs ---------- */
      .orb {{ position:fixed; border-radius:50%; filter:blur(70px); pointer-events:none; z-index:1;
              animation: drift 18s ease-in-out infinite alternate; }}
      .orb.o1 {{ width:340px; height:340px; background:rgba(215,167,47,.22); top:8%; left:42%; }}
      .orb.o2 {{ width:420px; height:420px; background:rgba(14,165,233,.18); bottom:-8%; left:-6%;
                 animation-duration:24s; }}
      @keyframes drift {{ from {{ transform:translate(0,0) scale(1); }}
                          to   {{ transform:translate(40px,-30px) scale(1.12); }} }}
      @keyframes rise {{ from {{ opacity:0; transform:translateY(18px); }}
                         to   {{ opacity:1; transform:translateY(0); }} }}
      @keyframes shimmer {{ 0% {{ background-position:0% 50%; }} 100% {{ background-position:200% 50%; }} }}
      @keyframes pulse {{ 0%,100% {{ box-shadow:0 0 0 0 rgba(52,211,153,.55); }}
                          50% {{ box-shadow:0 0 0 7px rgba(52,211,153,0); }} }}

      /* ---------- hero (left) ---------- */
      .hero {{ padding:18px 8px 0 8px; animation: rise .8s ease both; }}
      .login-plate {{
          display:flex; align-items:center; gap:16px; margin-bottom:30px;
          padding:10px 22px 10px 10px; width:fit-content; border-radius:999px;
          background: linear-gradient(135deg, rgba(255,255,255,.12), rgba(255,255,255,.04));
          border:1px solid rgba(215,167,47,.5);
          backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
          box-shadow: 0 8px 28px rgba(0,0,0,.28);
      }}
      .eyebrow {{ display:inline-flex; align-items:center; gap:8px; color:{GOLD}; font-size:11.5px;
                  font-weight:800; letter-spacing:2.2px; text-transform:uppercase; margin-bottom:14px; }}
      .eyebrow i {{ width:8px; height:8px; border-radius:50%; background:#34d399; display:inline-block;
                    animation:pulse 2s infinite; }}
      .hero-title {{ color:#fff; font-size:clamp(34px, 4.4vw, 56px); font-weight:900; line-height:1.04;
                     margin:0 0 16px 0; letter-spacing:-1.2px; }}
      .hero-title .accent {{
          background: linear-gradient(90deg, {GOLD}, #ffe9a8, {GOLD}, #f5c24b);
          background-size:200% auto; -webkit-background-clip:text; background-clip:text;
          -webkit-text-fill-color:transparent; animation: shimmer 6s linear infinite; }}
      .hero-sub {{ color:rgba(255,255,255,.78); font-size:15.5px; line-height:1.65;
                   max-width:560px; margin:0 0 26px 0; }}

      .stat-row {{ display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; max-width:560px;
                   margin-bottom:22px; }}
      .stat {{ padding:14px 14px 12px; border-radius:14px; background:rgba(255,255,255,.07);
               border:1px solid rgba(255,255,255,.14); backdrop-filter: blur(10px); }}
      .stat b {{ display:block; color:#fff; font-size:25px; font-weight:900; letter-spacing:-.5px; }}
      .stat span {{ color:rgba(255,255,255,.66); font-size:11.5px; font-weight:600; }}

      .preview {{ max-width:560px; border-radius:16px; padding:16px 18px; margin-bottom:22px;
                  background:linear-gradient(160deg, rgba(255,255,255,.12), rgba(255,255,255,.04));
                  border:1px solid rgba(255,255,255,.16); backdrop-filter: blur(14px);
                  box-shadow:0 18px 40px rgba(0,0,0,.30); }}
      .pv-head {{ display:flex; justify-content:space-between; align-items:center; color:#fff;
                  font-size:12px; font-weight:700; letter-spacing:.4px; margin-bottom:12px; }}
      .pv-head em {{ font-style:normal; font-size:10.5px; padding:3px 9px; border-radius:999px;
                     background:rgba(52,211,153,.18); color:#6ee7b7; border:1px solid rgba(52,211,153,.4); }}
      .pv-eq {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; }}
      .pv-eq div {{ text-align:center; padding:9px 4px; border-radius:10px; background:rgba(3,16,34,.45);
                    border:1px solid rgba(255,255,255,.08); }}
      .pv-eq small {{ display:block; color:rgba(255,255,255,.55); font-size:10px; font-weight:700;
                      text-transform:uppercase; letter-spacing:.6px; }}
      .pv-eq b {{ color:#fff; font-size:14px; font-weight:800; }}
      .pv-bar {{ margin-top:12px; height:7px; border-radius:99px; background:rgba(255,255,255,.1); overflow:hidden; }}
      .pv-bar i {{ display:block; height:100%; width:99.6%; border-radius:99px;
                   background:linear-gradient(90deg,#34d399,{GOLD}); }}
      .pv-foot {{ display:flex; justify-content:space-between; margin-top:7px; color:rgba(255,255,255,.6);
                  font-size:10.5px; font-weight:600; }}

      .hero-feat {{ display:flex; flex-wrap:wrap; gap:7px; max-width:600px; }}
      .hero-feat span {{
          display:inline-flex; align-items:center; gap:6px; padding:5px 11px; border-radius:999px;
          background: rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.16);
          color:rgba(255,255,255,.9); font-size:11px; font-weight:600; transition:all .2s; }}
      .hero-feat span:hover {{ background:rgba(215,167,47,.2); border-color:{GOLD}; transform:translateY(-2px); }}
      .hero-trust {{ margin-top:26px; color:rgba(255,255,255,.5); font-size:11.5px; letter-spacing:.3px; }}

      /* ---------- sign-in card (right) ---------- */
      .st-key-login_card {{
          position:relative; overflow:hidden; animation: rise .9s .1s ease both;
          background: linear-gradient(165deg, rgba(14,34,62,.78), rgba(8,22,44,.72));
          backdrop-filter: blur(26px) saturate(150%); -webkit-backdrop-filter: blur(26px) saturate(150%);
          border:1px solid rgba(255,255,255,.16); border-radius:22px; padding:30px 32px 26px;
          box-shadow:0 30px 70px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.12);
      }}
      .st-key-login_card::before {{
          content:''; position:absolute; left:0; right:0; top:0; height:4px;
          background:linear-gradient(90deg, {GOLD}, #ffe9a8, {GOLD}); background-size:200% auto;
          animation: shimmer 5s linear infinite; }}
      .card-head {{ display:flex; align-items:center; gap:12px; margin-bottom:4px; }}
      .card-head .ico {{ width:42px; height:42px; border-radius:12px; display:grid; place-items:center;
                         font-size:20px; background:linear-gradient(135deg,{GOLD},#f5c24b); color:{NAVY2};
                         box-shadow:0 6px 16px rgba(215,167,47,.35); }}
      .st-key-login_card h2 {{ color:#fff; margin:0; font-size:22px; font-weight:800; letter-spacing:-.3px; }}
      .st-key-login_card p.sub {{ color:rgba(255,255,255,.66); font-size:12.8px; margin:2px 0 14px 0; }}
      .st-key-login_card label, .st-key-login_card .stMarkdown p, .st-key-login_card [data-testid="stRadio"] label p,
      .st-key-login_card [data-testid="stCheckbox"] label p {{ color:rgba(255,255,255,.92) !important; }}
      .st-key-login_card label p {{ font-weight:600; font-size:12.8px; }}
      .st-key-login_card [data-baseweb="input"], .st-key-login_card [data-baseweb="select"] > div,
      .st-key-login_card [data-baseweb="base-input"] {{
          background: rgba(255,255,255,.07) !important; border:1px solid rgba(255,255,255,.2) !important;
          border-radius:11px !important; transition: border-color .2s, box-shadow .2s, background .2s; }}
      .st-key-login_card [data-baseweb="input"]:focus-within, .st-key-login_card [data-baseweb="select"] > div:focus-within {{
          border-color:{GOLD} !important; box-shadow:0 0 0 3px rgba(215,167,47,.25) !important;
          background: rgba(255,255,255,.1) !important; }}
      .st-key-login_card input {{ color:#fff !important; background:transparent !important; }}
      .st-key-login_card input::placeholder {{ color:rgba(255,255,255,.4) !important; }}
      .st-key-login_card [data-baseweb="select"] * {{ color:#fff !important; }}
      .st-key-login_card button[kind="icon"], .st-key-login_card [data-baseweb="input"] button {{ color:#fff !important; }}
      .st-key-login_card .stTabs [data-baseweb="tab-list"] {{
          background:rgba(255,255,255,.06); padding:5px; border-radius:13px; gap:4px; border:0; margin-bottom:12px; }}
      .st-key-login_card .stTabs [data-baseweb="tab"] {{
          flex:1; justify-content:center; color:rgba(255,255,255,.7); background:transparent;
          border-radius:9px; padding:8px 10px; font-weight:700; font-size:13px; height:auto; }}
      .st-key-login_card .stTabs [aria-selected="true"] {{
          background:linear-gradient(135deg,{GOLD},#f5c24b) !important; color:{NAVY2} !important;
          box-shadow:0 4px 12px rgba(215,167,47,.35); }}
      .st-key-login_card .stTabs [data-baseweb="tab-highlight"], .st-key-login_card .stTabs [data-baseweb="tab-border"] {{ display:none; }}
      .st-key-login_card [data-testid="stForm"] {{ border:0 !important; padding:0 !important; background:transparent !important; }}
      .st-key-login_card .stButton>button, .st-key-login_card [data-testid="stFormSubmitButton"]>button {{
          width:100%; border:0; border-radius:12px; padding:12px 18px; font-weight:800; font-size:15px;
          letter-spacing:.3px; color:{NAVY2}; background:linear-gradient(135deg,{GOLD},#f5c24b);
          box-shadow:0 10px 24px rgba(215,167,47,.32); transition:transform .15s, box-shadow .15s, filter .15s; }}
      .st-key-login_card .stButton>button:hover, .st-key-login_card [data-testid="stFormSubmitButton"]>button:hover {{
          transform:translateY(-2px); box-shadow:0 14px 30px rgba(215,167,47,.45); filter:brightness(1.05); color:{NAVY2}; }}
      .st-key-login_card [data-testid="stAlert"] {{ border-radius:11px; }}
      .st-key-login_card [data-testid="stRadio"] > div {{ gap:6px; flex-wrap:wrap; }}
      .st-key-login_card [data-testid="stRadio"] label {{
          background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.16); padding:6px 12px;
          border-radius:999px; }}
      .card-foot {{ text-align:center; color:rgba(255,255,255,.45); font-size:11px; margin-top:14px;
                    letter-spacing:.3px; }}

      @media (max-width: 900px) {{
          .preview, .stat-row, .hero-feat, .hero-trust {{ display:none; }}
          .hero {{ padding-top:0; }} .login-plate {{ margin-bottom:16px; }}
          .block-container {{ padding-top:3vh; }}
          .st-key-login_card {{ padding:24px 20px 20px; }}
      }}
      @media (prefers-reduced-motion: reduce) {{ * {{ animation:none !important; }} }}
    </style>
    <div class="orb o1"></div><div class="orb o2"></div>
    """


def kpi(label: str, value: str, delta: str = "", tone: str = "") -> str:
    tone_cls = f" {tone}" if tone else ""
    return (f'<div class="kpi{tone_cls}"><div class="l">{html.escape(label)}</div>'
            f'<div class="v">{html.escape(value)}</div>'
            f'<div class="d">{html.escape(delta)}</div></div>')


def role_badge(role: str) -> str:
    meta = ROLES.get(role, ROLES["staff"])
    return (f'<span class="role-badge" '
            f'style="background:{meta["color"]}22;color:{meta["color"]}">'
            f'{meta["icon"]} {meta["label"]}</span>')


def dept_badge(code: str) -> str:
    m = dept_meta(code)
    return (f'<span class="dept-badge" style="color:{m["color"]}">'
            f'{m["icon"]} {html.escape(m["short"])}</span>')


def dept_banner(code: str, extra: str = "") -> str:
    m = dept_meta(code)
    return f"""
    <div class="dept-banner" style="background:linear-gradient(120deg,{m['color']},{NAVY2})">
      <div class="icon">{m['icon']}</div>
      <div style="flex:1">
        <h2>{html.escape(m['label'])}</h2>
        <p>{html.escape(m['scope'])} {html.escape(extra)}</p>
      </div>
      <div style="text-align:right;font-size:11.5px;opacity:.92">
        <div>Tolerance</div>
        <div style="font-size:18px;font-weight:800">±{m['tol']}%</div>
      </div>
    </div>"""


# ==========================================================================
# Balance engine
# ==========================================================================
FIELD_ALIASES = {
    "date": ["date", "transaction date", "txn date", "report date", "posting date",
             "day", "period", "dated", "business date", "entry date"],
    "depot": ["depot", "depot name", "depotname", "terminal", "plant", "location",
              "site", "station", "facility", "installations", "tank farm", "node",
              "node name", "asset", "outlet", "outlet name", "jetty", "berth", "process unit",
              "unit name", "plant unit"],
    "department": ["department", "dept", "section", "business unit", "bu", "function",
                   "division", "operating unit"],
    "node_type": ["node type", "facility type", "asset type", "category", "type",
                  "operations type"],
    "product": ["product", "product name", "productname", "product type", "fuel",
                "fuel type", "commodity", "grade", "item", "description",
                "petroleum product"],
    "opening_stock": ["opening stock", "opening", "open stock", "opening balance",
                      "beginning stock", "opening inventory", "bop", "opening dip", "open", "open qty"],
    "receipts": ["receipts", "receipt", "discharge", "discharges", "inflow", "received",
                 "imports", "supply", "deliveries in", "in", "inbound", "vessel discharge",
                 "feedstock", "crude intake", "crude receipts", "intake", "feedstock charge",
                 "unloaded", "quantity received", "qty received"],
    "transfers_in": ["transfer in", "transfers in", "inter depot in", "received from"],
    "sales": ["sales", "deliveries", "delivery", "outflow", "dispatched", "dispatch",
              "ofd", "out", "issues", "issued", "sold", "loadings", "loading",
              "truck out", "gantry", "product output", "output", "production", "liftings",
              "lifting", "offtake", "off take", "dispatches", "outlet sales", "pump sales",
              "litres sold", "volume sold", "loaded out"],
    "transfers_out": ["transfer out", "transfers out", "inter depot out", "issued to"],
    "closing_stock": ["closing stock", "closing", "close stock", "closing balance",
                      "ending stock", "end stock", "physical stock", "actual stock",
                      "dip", "closing dip", "eop", "close", "close qty", "dip reading", "stock close"],
    "gains_losses": ["gain", "gains", "loss", "losses", "gain loss", "gain/loss",
                     "variance", "variance litres", "g/l", "stock variance",
                     "overage", "shortage", "process loss", "transit loss", "line loss", "yield loss"],
    "throughput": ["throughput", "volume", "quantity", "qty", "total volume", "turnover"],
    "units": ["unit", "units", "uom", "unit of measure"],
    "reference": ["reference", "ref", "ticket", "ticket no", "document", "waybill",
                  "vessel", "vessel name", "voyage", "cargo", "bl number", "truck no",
                  "trip no", "truck"],
}
PRIORITY = ["date", "depot", "department", "product", "node_type",
            "opening_stock", "receipts", "transfers_in", "sales", "transfers_out",
            "closing_stock", "gains_losses", "throughput", "units", "reference"]
FIELD_LABELS = {
    "date": "Transaction Date", "depot": "Depot / Terminal",
    "department": "Department (NPA demarcation)",
    "node_type": "Node / Facility Type", "product": "Product",
    "opening_stock": "Opening Stock", "receipts": "Receipts",
    "transfers_in": "Transfers In", "sales": "Sales / Deliveries",
    "transfers_out": "Transfers Out", "closing_stock": "Closing Stock (Actual)",
    "gains_losses": "Gain / (Loss) as Reported", "throughput": "Throughput",
    "units": "Units of Measure", "reference": "Reference / Ticket",
}
REQUIRED_FIELDS = ["depot", "product", "opening_stock", "receipts", "sales"]
BALANCE_REQUIRES = ["closing_stock", "gains_losses"]
UNIT_CONVERSIONS = {
    "L": 1.0, "LTR": 1.0, "LITRE": 1.0, "LITRES": 1.0, "LITER": 1.0, "LITERS": 1.0,
    "M3": 1000.0, "CBM": 1000.0, "CUM": 1000.0,
    "GAL": 3.78541, "GALLON": 3.78541, "GALLONS": 3.78541,
    "BBL": 158.987, "BARREL": 158.987, "BARRELS": 158.987,
}
SEVERITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3, "Info": 4}


def normalize(s: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def fmt(x, dec: int = 0) -> str:
    if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
        return "—"
    return f"{x:,.{dec}f}"


def compliance_pct(df: pd.DataFrame) -> float:
    """% of verifiable rows within tolerance (Unverified rows are excluded)."""
    ver = df[df["Status"] != "Unverified"]
    return float((ver["Status"] == "Within Tolerance").mean() * 100) if len(ver) else float("nan")


def fmt_comp(x) -> str:
    return "n/a" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.1f}%"


def fmt_pct(x, dec: int = 2) -> str:
    if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
        return "—"
    return f"{x:+.{dec}f}%"


def convert_to_litres(values: pd.Series, unit: str) -> pd.Series:
    return values * UNIT_CONVERSIONS.get(unit.upper(), 1.0)


def detect_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    mapping: Dict[str, Optional[str]] = {}
    used: set = set()
    for field in PRIORITY:
        aliases = [normalize(a) for a in FIELD_ALIASES[field]]
        best, best_score = None, 0
        for col in columns:
            if col in used:
                continue
            nc = normalize(re.sub(r"[\(\[\{][^\)\]\}]*[\)\]\}]", " ", str(col))) or normalize(col)
            for a in aliases:
                if not a:
                    continue
                if nc == a:
                    score = 3
                elif len(a) >= 4 and a in nc:          # 'in', 'out', 'dip' … must match exactly
                    score = 2
                elif len(nc) >= 4 and nc in a:
                    score = 2
                else:
                    score = 0
                if score > best_score:
                    best, best_score = col, score
        mapping[field] = best
        if best:
            used.add(best)
    return mapping


def safe_sheet_name(name: str) -> str:
    return re.sub(r"[\[\]\:\*\?\/\\]", "-", str(name))[:31] or "Sheet"


def to_excel_bytes(sheets: Dict[str, pd.DataFrame],
                   meta: Optional[Dict[str, str]] = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if meta:
            pd.DataFrame({"Field": list(meta), "Value": [str(v) for v in meta.values()]}
                         ).to_excel(writer, sheet_name="Cover", index=False)
        for name, frame in sheets.items():
            if frame is not None and not frame.empty:
                frame.to_excel(writer, sheet_name=safe_sheet_name(name), index=False)
    return buf.getvalue()


def to_csv_bytes(frame: pd.DataFrame) -> bytes:
    return frame.to_csv(index=False).encode("utf-8-sig")


# ==========================================================================
# FACILITY AUTO-DETECTION  (depot / refinery / import terminal / jetty / retail / transport)
# ==========================================================================
_SIG_WEIGHT = {"file": 4.0, "sheet": 3.0, "depot": 3.0, "header": 2.0,
               "product": 2.0, "reference": 1.0}
_SIG_LABEL = {"file": "file name", "sheet": "sheet name", "depot": "node name",
              "header": "column", "product": "product", "reference": "reference"}

FACILITY_SIGNALS: Dict[str, Dict[str, List[str]]] = {
    "REFINERY": {
        "name": ["refinery", "refining", "refined", "tor", "sentuo", "cdu", "fcc",
                 "reformer", "topping plant", "crude unit", "process unit"],
        "header": ["crude", "feedstock", "feed stock", "charge", "yield", "process loss",
                   "slate", "naphtha", "residue", "atmospheric", "refinery fuel",
                   "fuel and loss", "fuel & loss", "run"],
        "product": ["crude", "naphtha", "residue", "reformate", "slop", "vgo",
                    "atmospheric residue", "bitumen", "petcoke", "condensate", "rfg"],
    },
    "IMPORT": {
        "name": ["vessel", "cargo", "import", "imports", "bill of lading", "b/l",
                 "out-turn", "outturn", "out turn", "demurrage", "m/t", "mt", "m/v", "mv"],
        "header": ["vessel", "b/l", "bl quantity", "bill of lading", "out-turn", "outturn",
                   "out turn", "cargo", "demurrage", "voyage", "shore tank", "nor", "imo"],
        "product": [],
    },
    "MARINE": {
        "name": ["jetty", "berth", "marine", "bunker", "bunkering", "spm", "sbm",
                 "single point mooring", "offshore"],
        "header": ["jetty", "berth", "bunker", "loading arm", "line fill", "line-fill",
                   "vessel loaded", "draft", "mooring"],
        "product": ["bunker", "mgo", "ifo", "hfo", "marine gas oil"],
    },
    "RETAIL": {
        "name": ["retail", "outlet", "service station", "filling station", "forecourt",
                 "omc", "pump", "dealer", "station"],
        "header": ["pump", "nozzle", "outlet", "omc", "dealer", "forecourt", "meter reading",
                   "station", "shift", "litres sold", "cash sales"],
        "product": [],
    },
    "TRANSPORT": {
        "name": ["pipeline", "truck", "fleet", "haulage", "wapco", "brv", "bulk road",
                 "tanker truck", "road tanker", "transport", "logistics"],
        "header": ["truck", "waybill", "way bill", "trip", "driver", "registration",
                   "vehicle", "brv", "pipeline", "consignment", "transit", "haulage",
                   "destination"],
        "product": [],
    },
    "DEPOT": {
        "name": ["depot", "terminal", "bost", "tank farm", "storage", "installation",
                 "tema", "takoradi", "kumasi", "tamale", "buipe", "akosombo"],
        "header": ["dip", "ullage", "tank", "gantry", "ofd", "vcf", "density", "depot",
                   "tank no", "tank number", "opening dip", "closing dip"],
        "product": [],
    },
}


def _kw_hit(text: str, key: str) -> bool:
    k = key.strip().lower()
    if not k:
        return False
    pat = r"(?<![a-z0-9])" + re.escape(k) + (r"(?![a-z0-9])" if len(k) < 5 else "")
    return re.search(pat, text) is not None


def _score_facility(sources: Dict[str, List[str]]) -> Dict:
    """Score every facility type against the supplied text sources."""
    scores = {c: 0.0 for c in FACILITY_SIGNALS}
    evidence: Dict[str, List[str]] = {c: [] for c in FACILITY_SIGNALS}
    for stype, texts in sources.items():
        if stype not in _SIG_WEIGHT:
            continue
        kind = "header" if stype == "header" else ("product" if stype == "product" else "name")
        lowered = [str(t).lower() for t in texts if str(t).strip()]
        for code, sig in FACILITY_SIGNALS.items():
            hits: List[str] = []
            for key in sig.get(kind, []):
                if any(_kw_hit(t, key) for t in lowered):
                    hits.append(key)
            if hits:
                scores[code] += _SIG_WEIGHT[stype] * min(len(hits), 3)
                evidence[code] += [f"{_SIG_LABEL[stype]} “{h}”" for h in hits[:3]]
    top = max(scores.items(), key=lambda kv: kv[1])
    if top[1] <= 0:
        return {"code": "DEPOT", "confidence": 30, "scores": scores,
                "evidence": ["No facility keywords found — assumed Depot"]}
    total = sum(scores.values()) or 1.0
    share, strength = top[1] / total, 1 - math.exp(-top[1] / 6.0)
    conf = int(round(100 * (0.5 * share + 0.5 * strength)))
    return {"code": top[0], "confidence": max(35, min(conf, 97)),
            "scores": {k: round(v, 1) for k, v in scores.items()},
            "evidence": evidence[top[0]][:6]}


def detect_files(raw: pd.DataFrame) -> Dict[str, Dict]:
    """Per uploaded file: which kind of facility does it describe, and why."""
    out: Dict[str, Dict] = {}
    if raw is None or raw.empty or "__source_file" not in raw.columns:
        return out
    cols = [c for c in raw.columns if not str(c).startswith("__")]
    guess = detect_columns(cols)

    for fname, sub in raw.groupby("__source_file", sort=False):
        def uniq(col, n=150):
            if not col or col not in sub.columns:
                return []
            vals = sub[col].dropna().astype(str).str.strip()
            return list(dict.fromkeys(v for v in vals if v))[:n]

        sheets = (sub["__sheet"].dropna().astype(str).unique().tolist()
                  if "__sheet" in sub.columns else [])
        sources = {
            "file": [Path(str(fname)).stem],
            "sheet": sheets,
            "header": [c for c in cols if sub[c].notna().any()],
            "depot": uniq(guess.get("depot")),
            "product": uniq(guess.get("product")),
            "reference": uniq(guess.get("reference")),
        }
        out[str(fname)] = _score_facility(sources)
    return out


WEAK_DEPT_KEYS = {"bost", "tema", "takoradi", "kumasi", "tamale", "buipe"}


def infer_department_ex(name: str) -> Tuple[Optional[str], bool]:
    """(department, is_strong). Place names are only a weak hint for 'Depot'."""
    n = str(name).lower()
    for key, dept in DEPOT_KEYWORD_TO_DEPT.items():
        if key in n:
            return dept, key not in WEAK_DEPT_KEYS
    return None, False


def dataset_facility(fdf: pd.DataFrame, meta: Optional[Dict] = None,
                     override: Optional[str] = None) -> Dict:
    """Describe what kind of facility the (filtered) dataset represents."""
    meta = meta or {}
    thr = fdf.groupby("Department")["Throughput"].sum().clip(lower=0)
    if thr.sum() <= 0:
        thr = fdf["Department"].value_counts().astype(float)
    shares = (thr / thr.sum() * 100).sort_values(ascending=False)
    primary = str(shares.index[0])
    mixed = len(shares) > 1 and float(shares.iloc[1]) >= 20.0

    evidence: List[str] = []
    conf = int(round(float(shares.iloc[0])))
    try:
        det = json.loads(meta.get("facility_detection", "{}") or "{}")
    except Exception:
        det = {}
    files = set(fdf["Source File"].dropna().astype(str)) if "Source File" in fdf.columns else set()
    confs = []
    for f in files:
        d = det.get(f)
        if d and d.get("code") == primary:
            confs.append(int(d.get("confidence", 0)))
            evidence += d.get("evidence", [])
    if confs:
        conf = int(round(sum(confs) / len(confs)))
    if meta.get("department_source") == "column":
        evidence.insert(0, "Department column supplied in the file")
        conf = 99
    if mixed:
        conf = min(conf, 80)
    method = "auto"
    if override and override not in ("AUTO", "Auto-detect"):
        primary, conf, mixed, method = override, 100, False, "manual"
        evidence = ["Selected manually"]
    evidence = list(dict.fromkeys(evidence))[:6]
    m = dept_meta(primary)
    label, short = m["label"], m["short"]
    if mixed:
        tops = [dept_meta(k)["short"] for k in list(shares.index)[:3] if shares[k] >= 10]
        label, short = "Mixed facilities — " + " + ".join(tops), "Consolidated"
    return {"code": primary, "label": label, "short": short, "icon": m["icon"],
            "confidence": conf, "mixed": mixed, "method": method,
            "shares": {str(k): round(float(v), 1) for k, v in shares.items()},
            "evidence": evidence, "files": sorted(files)}


# ==========================================================================
# PERIOD ANALYSIS  (daily / weekly / monthly / yearly)
# ==========================================================================
PERIOD_KINDS = ["Daily", "Weekly", "Monthly", "Yearly"]
_PERIOD_FREQ = {"Daily": "D", "Weekly": "W", "Monthly": "M", "Yearly": "Y"}


def _to_period(dates: pd.Series, kind: str) -> pd.Series:
    freq = _PERIOD_FREQ[kind]
    try:
        return dates.dt.to_period(freq)
    except Exception:
        return dates.dt.to_period("A" if freq == "Y" else freq)


def _period_label(p: pd.Period, kind: str) -> str:
    s, e = p.start_time, p.end_time
    if kind == "Daily":
        return f"{s:%d %b %Y}"
    if kind == "Weekly":
        return f"{s:%d %b} - {e:%d %b %Y}"
    if kind == "Monthly":
        return f"{s:%b %Y}"
    return f"{s.year}"


def point_in_time(sub: pd.DataFrame) -> Tuple[float, float]:
    """Opening = first record, closing = last record per depot/product (stocks must not
    be added up across days — only flows are additive)."""
    op = cl = 0.0
    dated = sub["Date"].notna()
    if dated.any():
        d = sub[dated]
        key = ["Depot", "Product"]
        first = d.groupby(key, dropna=False)["Date"].transform("min")
        last = d.groupby(key, dropna=False)["Date"].transform("max")
        op += float(d.loc[d["Date"] == first, "Opening Stock"].sum())
        cl += float(d.loc[d["Date"] == last, "Actual Closing"].sum())
    if (~dated).any():
        u = sub[~dated]
        op += float(u["Opening Stock"].sum())
        cl += float(u["Actual Closing"].sum())
    return op, cl


def build_period_table(fdf: pd.DataFrame, kind: str) -> pd.DataFrame:
    d = fdf.dropna(subset=["Date"]).copy()
    cols = ["Period", "Records", "Opening", "Receipts", "Sales", "Closing",
            "Gain / (Loss)", "Throughput", "Var %", "Exceptions", "Compliance %"]
    if d.empty:
        return pd.DataFrame(columns=cols)
    d["_p"] = _to_period(d["Date"], kind)
    key = ["_p", "Depot", "Product"]
    first = d.groupby(key, dropna=False)["Date"].transform("min")
    last = d.groupby(key, dropna=False)["Date"].transform("max")
    d["_open"] = d["Opening Stock"].where(d["Date"] == first, 0.0)
    d["_close"] = d["Actual Closing"].where(d["Date"] == last, 0.0)
    d["_rec"] = d["Receipts"] + d["Transfers In"]
    d["_sal"] = d["Sales / Deliveries"] + d["Transfers Out"]
    d["_exc"] = d["Status"].isin(EXC_STATUSES)
    d["_ver"] = d["Status"] != "Unverified"
    d["_ok"] = d["Status"] == "Within Tolerance"
    g = d.groupby("_p").agg(
        Records=("_p", "size"), Opening=("_open", "sum"), Receipts=("_rec", "sum"),
        Sales=("_sal", "sum"), Closing=("_close", "sum"), GL=("Gain / (Loss)", "sum"),
        Throughput=("Throughput", "sum"), Exceptions=("_exc", "sum"),
        Ver=("_ver", "sum"), Ok=("_ok", "sum")).sort_index()
    out = pd.DataFrame({
        "Period": [_period_label(p, kind) for p in g.index],
        "Records": g["Records"].astype(int).values,
        "Opening": g["Opening"].values, "Receipts": g["Receipts"].values,
        "Sales": g["Sales"].values, "Closing": g["Closing"].values,
        "Gain / (Loss)": g["GL"].values, "Throughput": g["Throughput"].values,
        "Var %": np.where(g["Throughput"] > 0, g["GL"] / g["Throughput"] * 100, np.nan),
        "Exceptions": g["Exceptions"].astype(int).values,
        "Compliance %": np.where(g["Ver"] > 0, g["Ok"] / g["Ver"].replace(0, np.nan) * 100, np.nan),
    })
    return out


def period_tables(fdf: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    return {k: build_period_table(fdf, k) for k in PERIOD_KINDS}


def rows_in_period(fdf: pd.DataFrame, kind: str, label: str) -> pd.DataFrame:
    d = fdf.dropna(subset=["Date"])
    p = _to_period(d["Date"], kind)
    return d.loc[p.map(lambda x: _period_label(x, kind)) == label]


def period_display(pt: pd.DataFrame) -> pd.DataFrame:
    """Formatted copy for screen / PDF / Excel."""
    if pt.empty:
        return pt
    t = pt.copy()
    for c in ("Opening", "Receipts", "Sales", "Closing", "Gain / (Loss)", "Throughput"):
        t[c] = t[c].map(fmt)
    t["Var %"] = t["Var %"].map(fmt_pct)
    t["Compliance %"] = t["Compliance %"].map(fmt_comp)
    return t


# ==========================================================================
# ANALYSIS: KPIs + written findings per facility type
# ==========================================================================
def core_stats(sub: pd.DataFrame) -> Dict:
    opening, closing = point_in_time(sub)
    receipts = float(sub["Receipts"].sum() + sub["Transfers In"].sum())
    sales = float(sub["Sales / Deliveries"].sum() + sub["Transfers Out"].sum())
    gl = float(sub["Gain / (Loss)"].sum())
    thr = float(sub["Throughput"].sum())
    glc = sub["Gain / (Loss)"]
    return {
        "opening": opening, "closing": closing, "receipts": receipts, "sales": sales,
        "expected": opening + receipts - sales, "gl": gl, "thr": thr,
        "var_pct": gl / thr * 100 if thr > 0 else float("nan"),
        "losses": float(glc[glc < 0].sum()), "gains": float(glc[glc > 0].sum()),
        "n_exc": int(sub["Status"].isin(EXC_STATUSES).sum()),
        "n_unver": int((sub["Status"] == "Unverified").sum()), "n": len(sub),
        "compliance": compliance_pct(sub),
    }


def _span_days(sub: pd.DataFrame) -> int:
    d = sub["Date"].dropna()
    return int((d.max() - d.min()).days) + 1 if len(d) else 0


def _node_table(sub: pd.DataFrame) -> pd.DataFrame:
    g = sub.groupby("Depot", dropna=False).agg(
        GL=("Gain / (Loss)", "sum"), Thr=("Throughput", "sum"),
        Exc=("Status", lambda s: int(s.isin(EXC_STATUSES).sum())), N=("Status", "size"))
    g["Pct"] = np.where(g["Thr"] > 0, g["GL"] / g["Thr"] * 100, np.nan)
    return g


def facility_kpis(code: str, sub: pd.DataFrame, unit: str, s: Dict) -> List[Tuple[str, str]]:
    R, S, gl, closing = s["receipts"], s["sales"], s["gl"], s["closing"]
    days = _span_days(sub)
    nodes = int(sub["Depot"].nunique())
    ref_n = int(sub["Reference"].replace("", np.nan).dropna().nunique()) if "Reference" in sub else 0
    if code == "DEPOT":
        cover = closing / (S / days) if S > 0 and days and closing > 0 else float("nan")
        return [("Depot throughput", f"{fmt(s['thr'])} {unit}"),
                ("Gain / (Loss) % of throughput", fmt_pct(s["var_pct"])),
                ("Stock cover", "n/a" if np.isnan(cover) else f"{cover:,.1f} days"),
                ("Dip compliance", fmt_comp(100 * (1 - s["n_unver"] / s["n"]) if s["n"] else np.nan)),
                ("Depots reporting", f"{nodes}")]
    if code == "IMPORT":
        outturn = (R + gl) / R * 100 if R > 0 else float("nan")
        return [("Total discharged", f"{fmt(R)} {unit}"),
                ("Cargo / discharge records", f"{ref_n or s['n']:,}"),
                ("Out-turn vs receipts", "n/a" if np.isnan(outturn) else f"{outturn:,.3f}%"),
                ("Discharge variance", f"{fmt(gl)} {unit}"),
                ("Receiving terminals", f"{nodes}")]
    if code == "MARINE":
        return [("Jetty / bunkering volume", f"{fmt(S)} {unit}"),
                ("Receipts into shore", f"{fmt(R)} {unit}"),
                ("Line-fill variance", f"{fmt(gl)} {unit}"),
                ("Active jetties / nodes", f"{nodes}"),
                ("Variance % of throughput", fmt_pct(s["var_pct"]))]
    if code == "REFINERY":
        yld = S / R * 100 if R > 0 else float("nan")
        top = sub.groupby("Product")["Sales / Deliveries"].sum().sort_values(ascending=False)
        share = top.iloc[0] / top.sum() * 100 if len(top) and top.sum() > 0 else float("nan")
        return [("Feedstock / intake", f"{fmt(R)} {unit}"),
                ("Product output", f"{fmt(S)} {unit}"),
                ("Indicative yield", "n/a" if np.isnan(yld) else f"{yld:,.2f}%"),
                ("Process gain / (loss)", f"{fmt(gl)} {unit}"),
                ("Largest product in slate", "n/a" if np.isnan(share) else f"{top.index[0]} ({share:.1f}%)")]
    if code == "RETAIL":
        under = abs(s["losses"]) / R * 100 if R > 0 else float("nan")
        return [("Outlets reporting", f"{nodes}"), ("Outlet sales", f"{fmt(S)} {unit}"),
                ("Allocation vs lift", "n/a" if R <= 0 else f"{S / R * 100:,.1f}%"),
                ("Under-delivery %", "n/a" if np.isnan(under) else f"{under:,.3f}%"),
                ("Net gain / (loss)", f"{fmt(gl)} {unit}")]
    if code == "TRANSPORT":
        tl = abs(s["losses"]) / R * 100 if R > 0 else float("nan")
        return [("Volume hauled", f"{fmt(S)} {unit}"), ("Trips / records", f"{s['n']:,}"),
                ("Transit loss %", "n/a" if np.isnan(tl) else f"{tl:,.3f}%"),
                ("Net gain / (loss)", f"{fmt(gl)} {unit}"), ("Routes / nodes", f"{nodes}")]
    return [("Throughput", f"{fmt(s['thr'])} {unit}"), ("Net gain / (loss)", f"{fmt(gl)} {unit}"),
            ("Variance %", fmt_pct(s["var_pct"])), ("Records", f"{s['n']:,}"),
            ("Nodes", f"{nodes}")]


def _facility_lines(code: str, sub: pd.DataFrame, unit: str, s: Dict) -> List[str]:
    out: List[str] = []
    R, S = s["receipts"], s["sales"]
    nodes = _node_table(sub)
    if code == "DEPOT":
        days = _span_days(sub)
        if S > 0 and days:
            rows_ = []
            for dname, g in list(sub.groupby("Depot"))[:500]:
                _, cl = point_in_time(g)
                sl = float(g["Sales / Deliveries"].sum() + g["Transfers Out"].sum())
                if sl > 0 and cl > 0:
                    rows_.append((dname, cl / (sl / days)))
            low = [(n, c) for n, c in rows_ if c < 7]
            if low:
                n, c = min(low, key=lambda t: t[1])
                out.append(f"{len(low)} depot(s) hold less than 7 days of stock cover "
                           f"(lowest: {n} at {c:.1f} days).")
    elif code == "IMPORT" and R > 0:
        bad = int(((sub["Variance %"].abs() > sub["Tolerance %"])).sum())
        out.append(f"Overall out-turn is {(R + s['gl']) / R * 100:,.3f}% of quantity received; "
                   f"{bad} discharge record(s) are outside the {dept_meta('IMPORT')['tol']:.2f}% cargo tolerance.")
    elif code == "REFINERY" and R > 0:
        slate = (sub.groupby("Product")["Sales / Deliveries"].sum().sort_values(ascending=False))
        if slate.sum() > 0:
            top3 = ", ".join(f"{p} {v / slate.sum() * 100:.1f}%" for p, v in slate.head(3).items())
            out.append(f"Indicative yield is {S / R * 100:,.2f}% of intake (tank-movement basis). "
                       f"Product slate: {top3}.")
    elif code == "RETAIL":
        n_loss = int((nodes["GL"] < 0).sum())
        out.append(f"{n_loss} of {len(nodes)} outlet(s) show a net shortage; "
                   f"under-delivery totals {fmt(abs(s['losses']))} {unit}.")
    elif code == "TRANSPORT" and len(nodes.dropna(subset=['Pct'])):
        w = nodes["Pct"].idxmin()
        out.append(f"Highest transit loss is on {w} ({fmt_pct(nodes.loc[w, 'Pct'])}).")
    return out


def analyse_dataset(fdf: pd.DataFrame, fac: Dict, unit: str, issues: Optional[pd.DataFrame] = None) -> Dict:
    """KPIs, plain-language findings and a verdict for the dataset."""
    code = fac["code"]
    s = core_stats(fdf)
    m = dept_meta(code)
    mixed = bool(fac.get("mixed"))
    if mixed:        # throughput-weighted tolerance across the facility types present
        w = fdf.groupby("Department")["Throughput"].sum().clip(lower=0)
        tol = float(sum(dept_meta(k)["tol"] * v for k, v in w.items()) / w.sum()) if w.sum() > 0 \
            else m["tol"]
    else:
        tol = m["tol"]
    kpis = facility_kpis("MIXED" if mixed else code, fdf, unit, s)
    F: List[str] = []

    if fac.get("mixed"):
        mix = ", ".join(f"{dept_meta(k)['short']} {v:.0f}%" for k, v in list(fac["shares"].items())[:4])
        F.append(f"This file set covers several facility types ({mix} of throughput); "
                 f"figures below are consolidated; a breakdown by facility type follows.")

    vp = s["var_pct"]
    if np.isnan(vp):
        F.append("Closing stock / gain-loss could not be verified for these records, "
                 "so no balance conclusion can be drawn.")
    else:
        direction = "gain (overage)" if s["gl"] > 0 else "loss (shortage)"
        within = abs(vp) <= tol
        F.append(f"Net result is a {direction} of {fmt(abs(s['gl']))} {unit} "
                 f"({fmt_pct(vp)} of throughput), "
                 f"{'within' if within else 'outside'} the {tol:.2f}% tolerance "
                 f"{'(weighted across facility types)' if mixed else 'for ' + m['short']}.")
    if s["losses"] < 0 and s["gains"] > 0:
        F.append(f"Gross shortages total {fmt(abs(s['losses']))} {unit} against gross overages of "
                 f"{fmt(s['gains'])} {unit}; netting hides {fmt(min(abs(s['losses']), s['gains']))} {unit} "
                 f"of offsetting error.")

    nodes = _node_table(fdf)
    worst = nodes[nodes["GL"] < 0].sort_values("GL").head(3)
    if len(worst):
        F.append("Largest shortages: " + "; ".join(
            f"{n} ({fmt(r.GL)} {unit}, {fmt_pct(r.Pct)})" for n, r in worst.iterrows()) + ".")
    over = nodes[nodes["GL"] > 0].sort_values("GL", ascending=False).head(1)
    if len(over) and over["GL"].iloc[0] > 0.25 * abs(s["losses"] or 1):
        n, r = over.index[0], over.iloc[0]
        F.append(f"Largest overage: {n} (+{fmt(r.GL)} {unit}) — verify dips and meter readings.")

    if s["n_exc"]:
        top = nodes.sort_values("Exc", ascending=False).iloc[0]
        share = top["Exc"] / s["n_exc"] * 100
        F.append(f"{s['n_exc']:,} of {s['n']:,} records ({s['n_exc'] / s['n'] * 100:.1f}%) are exceptions"
                 + (f"; {share:.0f}% of them are at {nodes['Exc'].idxmax()}." if share >= 40 and len(nodes) > 1 else "."))
    prod = fdf.groupby("Product").agg(GL=("Gain / (Loss)", "sum"), Thr=("Throughput", "sum"))
    pw = prod[prod["GL"] < 0].sort_values("GL").head(1)
    if len(pw) and len(prod) > 1:
        F.append(f"Product with the largest net shortage: {pw.index[0]} "
                 f"({fmt(pw['GL'].iloc[0])} {unit}, "
                 f"{fmt_pct(pw['GL'].iloc[0] / pw['Thr'].iloc[0] * 100 if pw['Thr'].iloc[0] else np.nan)}).")

    mt = build_period_table(fdf, "Monthly")
    if len(mt) >= 3 and mt["Var %"].notna().sum() >= 3:
        last, prev = mt["Var %"].iloc[-1], mt["Var %"].iloc[:-1].mean()
        if pd.notna(last) and pd.notna(prev):
            trend = "improving" if abs(last) < abs(prev) else "deteriorating"
            F.append(f"Monthly trend is {trend}: {mt['Period'].iloc[-1]} at {fmt_pct(last)} "
                     f"versus an average of {fmt_pct(prev)} in the {len(mt) - 1} earlier month(s).")
    dt_ = build_period_table(fdf, "Daily")
    if len(dt_) >= 5:
        i = dt_["Gain / (Loss)"].abs().idxmax()
        F.append(f"Largest single-day variance: {dt_.loc[i, 'Period']} "
                 f"({fmt(dt_.loc[i, 'Gain / (Loss)'])} {unit}).")

    if mixed:
        for dcode in [d for d in DEPARTMENT_ORDER if fac["shares"].get(d, 0) >= 20]:
            sub = fdf[fdf["Department"] == dcode]
            ss = core_stats(sub)
            dm = dept_meta(dcode)
            vv = ss["var_pct"]
            F.append(f"{dm['short']}: net {'gain' if ss['gl'] > 0 else 'loss'} of {fmt(abs(ss['gl']))} {unit} "
                     f"({fmt_pct(vv)}), tolerance ±{dm['tol']:.2f}% — "
                     f"{'within' if pd.notna(vv) and abs(vv) <= dm['tol'] else 'OUTSIDE tolerance'}.")
            F += _facility_lines(dcode, sub, unit, ss)
    else:
        F += _facility_lines(code, fdf, unit, s)

    if s["n_unver"]:
        F.append(f"{s['n_unver']:,} record(s) ({s['n_unver'] / s['n'] * 100:.1f}%) have no closing stock "
                 f"or gain/loss and are excluded from variance totals.")
    if issues is not None and len(issues):
        serious = int(issues["Severity"].isin(["Critical", "High"]).sum())
        if serious:
            F.append(f"{serious} high-severity data-quality issue(s) were found — see the Data Quality sheet.")

    n_exc = s["n_exc"]
    verdict = ("PASS" if n_exc == 0 else "REVIEW REQUIRED" if n_exc < 10 else "ACTION REQUIRED")
    return {"stats": s, "kpis": kpis, "findings": F, "verdict": verdict, "tol": tol}


def facility_breakdown(fdf: pd.DataFrame, unit: str) -> List[List[str]]:
    """Header + rows: one line per facility type present in the data."""
    rows = [["Facility type", "Records", "Nodes", "Throughput", "Gain / (Loss)", "Var %",
             "Tolerance", "Status"]]
    for code in [d for d in DEPARTMENT_ORDER if d in set(fdf["Department"])]:
        sub = fdf[fdf["Department"] == code]
        s = core_stats(sub)
        tol = dept_meta(code)["tol"]
        vp = s["var_pct"]
        status = ("No data" if np.isnan(vp) else "Within" if abs(vp) <= tol
                  else "Loss breach" if vp < 0 else "Gain breach")
        rows.append([dept_meta(code)["label"], f"{s['n']:,}", f"{sub['Depot'].nunique()}",
                     fmt(s["thr"]), fmt(s["gl"]), fmt_pct(vp), f"±{tol:.2f}%", status])
    return rows


def methodology_text(code: str, unit: str) -> str:
    base = ("Method: opening and closing stock are point-in-time values (first and last record of "
            "each depot/product in the period); receipts and sales are summed. Gain/(Loss) is the sum "
            "of record-level variances (actual closing minus expected closing). Variance % = Gain/(Loss) "
            f"divided by throughput. All volumes are in {unit}. Records without a closing stock are "
            "'Unverified' and excluded from variance totals. ")
    extra = {
        "REFINERY": "Refinery yield is indicative: product issues divided by intake, from tank movements. ",
        "IMPORT": "Out-turn compares quantity received with the balance-implied quantity. ",
        "DEPOT": "Stock cover = closing stock / average daily deliveries over the period. ",
    }.get(code, "")
    return base + extra + "Tolerance follows the NPA demarcation for the detected facility type."


# ==========================================================================
# CHARTS FOR THE PDF  (matplotlib, 300 dpi — crisp when zoomed or printed)
# ==========================================================================
_CHART_DPI = 300


def _png(fig) -> Dict:
    from PIL import Image
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=_CHART_DPI, bbox_inches="tight", facecolor="white")
    data = buf.getvalue()
    w, h = Image.open(io.BytesIO(data)).size
    return {"png": data, "w": w, "h": h}


def _compact(x, _pos=None) -> str:
    a = abs(x)
    for div, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= div:
            v = x / div
            return (f"{v:.0f}" if abs(v - round(v)) < 0.05 else f"{v:.1f}") + suf
    return f"{x:.0f}"


def _style_axis(ax) -> None:
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#c9d3df")
    ax.tick_params(colors="#475467", labelsize=7)
    ax.grid(axis="y", color="#e7ecf2", linewidth=0.6)
    ax.set_axisbelow(True)


def chart_period(pt: pd.DataFrame, kind: str, unit: str, tol: float) -> Optional[Dict]:
    """Gain/(Loss) bars + variance-% line with the tolerance band."""
    if pt is None or len(pt) < 2:
        return None
    try:
        from matplotlib.figure import Figure
        from matplotlib.ticker import FuncFormatter
        fig = Figure(figsize=(7.4, 3.5))
        gs = fig.add_gridspec(2, 1, height_ratios=[3, 2], hspace=0.38)
        a1, a2 = fig.add_subplot(gs[0]), fig.add_subplot(gs[1])
        x = np.arange(len(pt))
        gl = pt["Gain / (Loss)"].fillna(0).values
        a1.bar(x, gl, color=np.where(gl < 0, "#b42318", "#157347"), width=0.72)
        a1.axhline(0, color="#667085", linewidth=0.7)
        a1.set_ylabel(f"Gain / (Loss) ({unit})", fontsize=7.5, color="#344054")
        a1.yaxis.set_major_formatter(FuncFormatter(_compact))
        a1.set_title(f"{kind} gain / (loss) and variance %", fontsize=9.5, loc="left",
                     color="#0a2945", fontweight="bold")
        vp = pt["Var %"].values
        a2.axhspan(-tol, tol, color="#d1fadf", alpha=0.7, label=f"±{tol:.2f}% tolerance")
        a2.plot(x, vp, color="#123b63", linewidth=1.4, marker="o", markersize=2.8)
        a2.axhline(0, color="#667085", linewidth=0.6)
        a2.set_ylabel("Variance %", fontsize=7.5, color="#344054")
        a2.legend(fontsize=6.5, frameon=False, loc="upper right")
        short = {"Daily": lambda s: s[:6], "Weekly": lambda s: s.split(" - ")[0],
                 "Monthly": lambda s: s, "Yearly": lambda s: s}[kind]
        labels = [short(s) for s in pt["Period"]]
        step = max(1, len(labels) // 14)
        for ax in (a1, a2):
            _style_axis(ax)
            ax.set_xticks(x[::step])
        a1.set_xticklabels([])
        a2.set_xticklabels(labels[::step], rotation=40, ha="right", fontsize=6.5)
        return _png(fig)
    except Exception:
        return None


def chart_nodes(sub: pd.DataFrame, unit: str) -> Optional[Dict]:
    """Gain/(Loss) by node and throughput by product."""
    try:
        from matplotlib.figure import Figure
        from matplotlib.ticker import FuncFormatter
        nodes = _node_table(sub).dropna(subset=["GL"])
        if nodes.empty:
            return None
        nodes = nodes.reindex(nodes["GL"].abs().sort_values(ascending=False).index).head(10).iloc[::-1]
        prod = sub.groupby("Product")["Throughput"].sum().sort_values(ascending=False).head(8).iloc[::-1]
        fig = Figure(figsize=(7.4, 3.3))
        a1, a2 = fig.add_subplot(1, 2, 1), fig.add_subplot(1, 2, 2)
        names = [str(n)[:22] for n in nodes.index]
        a1.barh(names, nodes["GL"].values,
                color=np.where(nodes["GL"].values < 0, "#b42318", "#157347"), height=0.7)
        a1.axvline(0, color="#667085", linewidth=0.7)
        a1.set_title(f"Gain / (loss) by node ({unit})", fontsize=8.5, loc="left",
                     color="#0a2945", fontweight="bold")
        a1.xaxis.set_major_formatter(FuncFormatter(_compact))
        a2.barh([str(p)[:20] for p in prod.index], prod.values, color="#123b63", height=0.7)
        a2.set_title(f"Throughput by product ({unit})", fontsize=8.5, loc="left",
                     color="#0a2945", fontweight="bold")
        a2.xaxis.set_major_formatter(FuncFormatter(_compact))
        from matplotlib.ticker import MaxNLocator
        for ax in (a1, a2):
            _style_axis(ax)
            ax.xaxis.set_major_locator(MaxNLocator(nbins=4))
            ax.grid(axis="x", color="#e7ecf2", linewidth=0.6)
            ax.grid(axis="y", visible=False)
            ax.tick_params(labelsize=6.5)
        fig.subplots_adjust(wspace=0.55)
        return _png(fig)
    except Exception:
        return None


# ==========================================================================
# PDF ENGINE (fpdf2) — embedded Unicode fonts, vector text, 300-dpi charts
# ==========================================================================
_EMOJI_RE = re.compile("[\U0001F000-\U0001FFFF\u2600-\u27BF\uFE0F\u200d]")


def _pdf_clean(txt) -> str:
    """Latin-1 fallback used only when the Unicode font cannot be loaded."""
    s = _EMOJI_RE.sub("", str(txt))
    for bad, good in (("\u2014", "-"), ("\u2013", "-"), ("\u2019", "'"), ("\u2018", "'"),
                      ("\u201c", '"'), ("\u201d", '"'), ("\u2026", "..."), ("\u00b1", "+/-"),
                      ("\u2022", "-"), ("\u2192", "->"), ("\u20b5", "GHS ")):
        s = s.replace(bad, good)
    return s.encode("latin-1", "replace").decode("latin-1")


def _pdf_fonts() -> Optional[Dict[str, str]]:
    cands = [Path(APP_DIR) / "assets" / "fonts"]
    try:
        import matplotlib
        cands.append(Path(matplotlib.get_data_path()) / "fonts" / "ttf")
    except Exception:
        pass
    names = {"": "DejaVuSans.ttf", "B": "DejaVuSans-Bold.ttf",
             "I": "DejaVuSans-Oblique.ttf", "BI": "DejaVuSans-BoldOblique.ttf"}
    for d in cands:
        paths = {k: d / v for k, v in names.items()}
        if all(p.exists() for p in paths.values()):
            return {k: str(p) for k, p in paths.items()}
    return None


@st.cache_data(show_spinner=False)
def _logo_png_for_pdf(path_str: str, mtime: float) -> Tuple[bytes, Tuple[int, int]]:
    """Logo → RGBA PNG, up-sampled with Lanczos so viewers never show blocky pixels."""
    from PIL import Image
    im = Image.open(path_str).convert("RGBA")
    w, h = im.size
    if max(w, h) < 1000:
        s = 1000 / max(w, h)
        im = im.resize((max(1, int(w * s)), max(1, int(h * s))), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    return buf.getvalue(), (w, h)


def get_pdf_logo() -> Optional[Dict]:
    """{'kind': 'png'|'svg', ...} for the official logo, or None."""
    try:
        ov = st.session_state.get("logo_url_override")
        if ov and ov.startswith("data:image/") and "svg" not in ov[:30]:
            from PIL import Image
            raw_bytes = base64.b64decode(ov.split(",", 1)[1])
            im = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
            w, h = im.size
            if max(w, h) < 1000:
                s = 1000 / max(w, h)
                im = im.resize((max(1, int(w * s)), max(1, int(h * s))), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return {"kind": "png", "data": buf.getvalue(), "orig": (w, h)}
        p = _find_logo_path()
        if p is None:
            return None
        if p.suffix.lower() == ".svg":
            return {"kind": "svg", "path": str(p)}
        data, orig = _logo_png_for_pdf(str(p), p.stat().st_mtime)
        return {"kind": "png", "data": data, "orig": orig}
    except Exception:
        return None


_STATUS_COLORS = {"loss": (180, 35, 24), "gain": (180, 105, 14), "within": (21, 115, 71),
                  "pass": (21, 115, 71), "unverified": (102, 112, 133)}


def _status_color(txt: str):
    t = str(txt).lower()
    for key in ("loss", "gain", "within", "unverified"):
        if t.startswith(key):
            return _STATUS_COLORS[key]
    return None


if _HAS_FPDF:
    from fpdf.fonts import FontFace

    class _NPAPdf(FPDF):
        def __init__(self, running_title: str, fonts: Optional[Dict[str, str]]):
            super().__init__(orientation="P", unit="mm", format="A4")
            self.running_title = running_title
            self.uni = fonts is not None
            self.F = "DejaVu" if self.uni else "Helvetica"
            if fonts:
                for style, path in fonts.items():
                    self.add_font("DejaVu", style, path)
            self.set_auto_page_break(auto=True, margin=17)
            self.set_margins(12, 12, 12)
            self.set_title(running_title)
            self.set_author("National Petroleum Authority")
            self.alias_nb_pages()

        def t(self, s) -> str:
            return _EMOJI_RE.sub("", str(s)) if self.uni else _pdf_clean(s)

        def header(self):
            if self.page_no() == 1:
                return
            self.set_fill_color(10, 41, 69)
            self.rect(0, 0, 210, 10, "F")
            self.set_fill_color(215, 167, 47)
            self.rect(0, 10, 210, 0.8, "F")
            self.set_xy(12, 2.6)
            self.set_font(self.F, "B", 8)
            self.set_text_color(255, 255, 255)
            self.cell(150, 5, self.t(f"NPA · {self.running_title}"))
            self.set_font(self.F, "", 8)
            self.cell(0, 5, f"Page {self.page_no()}", align="R")
            self.set_y(16)

        def footer(self):
            self.set_y(-12)
            self.set_draw_color(215, 167, 47)
            self.line(12, self.get_y(), 198, self.get_y())
            self.set_font(self.F, "", 7.5)
            self.set_text_color(102, 112, 133)
            self.cell(140, 6, self.t(f"Generated {dt.datetime.now():%Y-%m-%d %H:%M} · "
                                     f"NPA Enterprise Material Balance Portal"))
            self.cell(0, 6, f"Page {self.page_no()} of {{nb}}", align="R")

        # ---- building blocks -------------------------------------------------
        def need(self, h: float) -> None:
            if self.get_y() + h > self.h - 17:
                self.add_page()

        def h2(self, text: str, color=(10, 41, 69)) -> None:
            self.need(22)
            self.ln(3)
            self.set_font(self.F, "B", 11)
            self.set_text_color(*color)
            self.cell(0, 6, self.t(text.upper()), new_x="LMARGIN", new_y="NEXT")
            y = self.get_y()
            self.set_fill_color(215, 167, 47)
            self.rect(12, y, 18, 0.9, "F")
            self.set_y(y + 2.5)

        def para(self, text: str, size=9, color=(52, 64, 84), italic=False) -> None:
            self.set_font(self.F, "I" if italic else "", size)
            self.set_text_color(*color)
            self.multi_cell(0, 4.8, self.t(text), new_x="LMARGIN", new_y="NEXT")

        def bullets(self, items: List[str], size=9) -> None:
            mark = "•" if self.uni else "-"
            self.set_font(self.F, "", size)
            self.set_text_color(40, 50, 70)
            for it in items:
                self.need(10)
                x0 = self.l_margin
                self.set_x(x0)
                self.set_text_color(215, 167, 47)
                self.cell(5, 4.8, mark)
                self.set_text_color(40, 50, 70)
                self.multi_cell(0, 4.8, self.t(it), new_x="LMARGIN", new_y="NEXT")
                self.ln(0.8)

        def fit_text(self, text: str, w: float, style: str, size: float, minsize=6.0) -> float:
            s = size
            self.set_font(self.F, style, s)
            while s > minsize and self.get_string_width(self.t(text)) > w:
                s -= 0.5
                self.set_font(self.F, style, s)
            return s

        def cards(self, items: List[Tuple[str, str]], per_row=4) -> None:
            gap = 3.0
            w = (186 - gap * (per_row - 1)) / per_row
            h = 17.5
            for i in range(0, len(items), per_row):
                self.need(h + 4)
                y = self.get_y()
                for j, (label, value) in enumerate(items[i:i + per_row]):
                    x = 12 + j * (w + gap)
                    self.set_fill_color(244, 247, 251)
                    self.set_draw_color(221, 228, 236)
                    self.rect(x, y, w, h, "DF")
                    self.set_fill_color(215, 167, 47)
                    self.rect(x, y, 1.1, h, "F")
                    self.fit_text(label, w - 5, "", 7.2)
                    self.set_text_color(102, 112, 133)
                    self.set_xy(x + 3, y + 2.2)
                    self.cell(w - 4, 4, self.t(label))
                    self.fit_text(value, w - 5, "B", 11.5, 7.5)
                    self.set_text_color(10, 41, 69)
                    self.set_xy(x + 3, y + 8.6)
                    self.cell(w - 4, 6, self.t(value))
                self.set_y(y + h + 3)

        def banner(self, text: str, rgb) -> None:
            self.need(14)
            y = self.get_y()
            self.set_fill_color(*rgb)
            self.rect(12, y, 186, 9, "F")
            self.set_xy(15, y + 1.6)
            self.set_font(self.F, "B", 9.5)
            self.set_text_color(255, 255, 255)
            self.cell(180, 6, self.t(text))
            self.set_y(y + 12)

        def image_png(self, ch: Dict, caption: str = "") -> None:
            w = 186.0
            h = w * ch["h"] / ch["w"]
            self.need(h + 6)
            self.image(io.BytesIO(ch["png"]), x=12, y=self.get_y(), w=w, h=h)
            self.set_y(self.get_y() + h + 2)
            if caption:
                self.para(caption, 7.5, (102, 112, 133), italic=True)

        def grid(self, headers: List[str], rows: List[List[str]],
                 widths: Optional[List[float]] = None, size: float = 8.0) -> None:
            if not rows:
                return
            n = len(headers)
            widths = widths or [1.0] * n
            tot = float(sum(widths))
            widths = [w * 186.0 / tot for w in widths]
            num_re = re.compile(r"^[\-\+\(]?[\d,\.]+\)?%?(\s?[A-Za-z]{1,4})?$|^—$|^n/a$")
            aligns = []
            for ci in range(n):
                col = [str(r[ci]) for r in rows if ci < len(r)]
                aligns.append("R" if col and all(num_re.match(c.strip()) for c in col) else "L")
            self.set_font(self.F, "", size)
            self.set_text_color(40, 50, 70)
            self.set_fill_color(255, 255, 255)      # unfilled rows inherit this colour
            self.need(18)
            with self.table(
                    col_widths=tuple(widths), text_align=tuple(aligns),
                    headings_style=FontFace(emphasis="BOLD", color=(255, 255, 255),
                                            fill_color=(18, 59, 99)),
                    cell_fill_color=(244, 247, 251), cell_fill_mode="ROWS",
                    line_height=size * 0.46 + 1.1, borders_layout="HORIZONTAL_LINES",
                    padding=(1.2, 1.4), repeat_headings=1) as tb:
                hr = tb.row()
                for h in headers:
                    hr.cell(self.t(h))
                for r in rows:
                    row = tb.row()
                    for ci, c in enumerate(r):
                        col = _status_color(c) if str(headers[ci]).lower() == "status" else None
                        if col:
                            row.cell(self.t(c), style=FontFace(emphasis="BOLD", color=col))
                        else:
                            row.cell(self.t(c))
            self.ln(2)


def _draw_logo(pdf, logo: Optional[Dict], x: float, y: float, size: float) -> bool:
    pdf.set_fill_color(255, 255, 255)
    pdf.ellipse(x, y, size, size, "F")
    pad = size * 0.13
    box = size - 2 * pad
    try:
        if logo and logo["kind"] == "png":
            ow, oh = logo["orig"]
            k = min(box / ow, box / oh)
            w, h = ow * k, oh * k
            pdf.image(io.BytesIO(logo["data"]), x=x + (size - w) / 2, y=y + (size - h) / 2, w=w, h=h)
            return True
        if logo and logo["kind"] == "svg":
            pdf.image(logo["path"], x=x + pad, y=y + pad, w=box, h=box, keep_aspect_ratio=True)
            return True
    except Exception:
        pass
    pdf.set_text_color(18, 59, 99)
    pdf.set_font(pdf.F, "B", 11)
    pdf.set_xy(x, y + size / 2 - 2.5)
    pdf.cell(size, 5, "NPA", align="C")
    return False


def build_pdf_report(
    title: str,
    subtitle: str,
    meta_lines: List[Tuple[str, str]],
    kpi_rows: List[Tuple[str, str]],
    summary_table: List[List[str]],
    department_table: Optional[List[List[str]]] = None,
    exceptions: Optional[List[List[str]]] = None,
    note: str = "",
    *,
    facility: Optional[Dict] = None,
    facility_kpis_: Optional[List[Tuple[str, str]]] = None,
    findings: Optional[List[str]] = None,
    verdict: Optional[Tuple[str, Tuple[int, int, int]]] = None,
    period_sections: Optional[List[Dict]] = None,
    charts: Optional[List[Dict]] = None,
    extra_tables: Optional[List[Dict]] = None,
    methodology: str = "",
) -> bytes:
    """Professional PDF: embedded fonts, 300-dpi charts, wrapped tables, page numbers."""
    if not _HAS_FPDF:
        raise RuntimeError("fpdf2 not installed. Run: pip install fpdf2")

    pdf = _NPAPdf(title, _pdf_fonts())
    pdf.add_page()

    # ---- header band -------------------------------------------------------------
    pdf.set_fill_color(10, 41, 69)
    pdf.rect(0, 0, 210, 34, "F")
    pdf.set_fill_color(215, 167, 47)
    pdf.rect(0, 34, 210, 1.2, "F")
    logo = get_pdf_logo()
    _draw_logo(pdf, logo, 11, 5, 24)
    pdf.set_xy(40, 9)
    pdf.set_text_color(215, 167, 47)
    pdf.set_font(pdf.F, "B", 15)
    pdf.cell(0, 7, pdf.t("NATIONAL PETROLEUM AUTHORITY"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(40)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(pdf.F, "", 11)
    pdf.cell(0, 6, pdf.t(title), new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(40)
    pdf.set_font(pdf.F, "", 8)
    pdf.set_text_color(190, 205, 222)
    pdf.cell(0, 5, pdf.t(f"Generated {dt.datetime.now():%d %B %Y, %H:%M}"))

    pdf.set_y(42)
    pdf.set_text_color(10, 41, 69)
    pdf.set_font(pdf.F, "B", 16)
    pdf.cell(0, 8, pdf.t(subtitle), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font(pdf.F, "", 8.5)
    pdf.set_text_color(102, 112, 133)
    pdf.multi_cell(0, 4.6, pdf.t("   ·   ".join(f"{k}: {v}" for k, v in meta_lines)),
                   new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    # ---- auto-detected facility box ---------------------------------------------
    if facility:
        y = pdf.get_y()
        ev = "; ".join(facility.get("evidence", [])[:4]) or "—"
        mix = ""
        if facility.get("mixed"):
            mix = "Mixed data — " + ", ".join(f"{dept_meta(k)['short']} {v:.0f}%"
                                              for k, v in list(facility["shares"].items())[:4])
        box_h = 19 if mix else 15
        pdf.set_fill_color(234, 241, 250)
        pdf.set_draw_color(47, 111, 176)
        pdf.rect(12, y, 186, box_h, "DF")
        pdf.set_xy(15, y + 2)
        pdf.set_font(pdf.F, "B", 9)
        pdf.set_text_color(18, 59, 99)
        how = "Selected" if facility.get("method") == "manual" else "Auto-detected"
        pdf.cell(0, 5, pdf.t(f"{how} data source: {facility['label']}  ·  confidence {facility['confidence']}%"),
                 new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(15)
        pdf.set_font(pdf.F, "", 7.8)
        pdf.set_text_color(70, 85, 105)
        pdf.cell(180, 4.5, pdf.t(f"Evidence: {ev}")[:150], new_x="LMARGIN", new_y="NEXT")
        if mix:
            pdf.set_x(15)
            pdf.cell(180, 4.5, pdf.t(mix))
        pdf.set_y(y + box_h + 3)

    if verdict:
        pdf.banner(verdict[0], verdict[1])

    # ---- KPIs ------------------------------------------------------------------------
    if kpi_rows:
        pdf.h2("Key performance indicators")
        pdf.cards(kpi_rows[:8])
    if facility_kpis_:
        pdf.h2(f"{facility['short']} indicators" if facility else "Facility indicators")
        pdf.cards(facility_kpis_[:8], per_row=3 if len(facility_kpis_) <= 6 else 4)

    # ---- written analysis ---------------------------------------------------------------
    if findings:
        pdf.h2("Analysis and key findings")
        pdf.bullets(findings)

    if summary_table:
        pdf.h2("Balance summary")
        pdf.grid(["Item", "Volume"], summary_table, [3, 2], size=9)

    if department_table and len(department_table) > 1:
        pdf.h2("Departmental breakdown")
        pdf.grid(department_table[0], department_table[1:], [3.2, 1, 1, 1.8, 1.2, 1.4])

    for tb in extra_tables or []:
        pdf.h2(tb["title"])
        pdf.grid(tb["headers"], tb["rows"], tb.get("widths"), tb.get("size", 8.0))

    for ch in charts or []:
        if ch:
            pdf.need(60)
            pdf.image_png(ch)

    # ---- daily / weekly / monthly / yearly -------------------------------------------
    for sec in period_sections or []:
        pdf.add_page()
        pdf.h2(sec["title"])
        if sec.get("chart"):
            pdf.image_png(sec["chart"])
        pdf.grid(sec["headers"], sec["rows"], sec.get("widths"), sec.get("size", 7.2))
        if sec.get("note"):
            pdf.para(sec["note"], 7.5, (102, 112, 133), italic=True)

    if exceptions and len(exceptions) > 1:
        pdf.h2("Top exceptions", color=(180, 35, 24))
        pdf.grid(exceptions[0], exceptions[1:], [1.5, 1.6, 2.6, 1.4, 1.2, 1.7])

    if note:
        pdf.ln(2)
        pdf.para(note, 8.5, (102, 112, 133), italic=True)
    if methodology:
        pdf.h2("Methodology")
        pdf.para(methodology, 8, (102, 112, 133))

    return bytes(pdf.output())


PERIOD_PDF_HEADERS = ["Period", "Rows", "Opening", "Receipts", "Sales", "Closing",
                      "Gain / (Loss)", "Var %", "Exc.", "In tol."]
_PERIOD_LIMIT = {"Daily": 31, "Weekly": 26, "Monthly": 36, "Yearly": 15}


def pdf_extras(fdf: pd.DataFrame, meta: Dict, tol_pct: float, unit: str,
               override: str = "AUTO", issues: Optional[pd.DataFrame] = None) -> Dict:
    """Everything the PDF needs beyond the basic tables (detection, findings, charts)."""
    fac = dataset_facility(fdf, meta, override)
    an = analyse_dataset(fdf, fac, unit, issues)
    tol = an["tol"]
    sections = []
    for kind, pt in period_tables(fdf).items():
        if pt.empty:
            continue
        shown = pt.tail(_PERIOD_LIMIT[kind])
        disp = period_display(shown)
        rows = [[r["Period"], str(r["Records"]), r["Opening"], r["Receipts"], r["Sales"],
                 r["Closing"], r["Gain / (Loss)"], r["Var %"], str(r["Exceptions"]),
                 r["Compliance %"]] for _, r in disp.iterrows()]
        sections.append({
            "title": f"{kind} report ({unit})", "headers": PERIOD_PDF_HEADERS, "rows": rows,
            "widths": [2.2, 1.0, 1.9, 1.9, 1.9, 1.9, 2.1, 1.6, 0.9, 1.4],
            "chart": chart_period(shown, kind, unit, tol),
            "note": (f"Showing the latest {len(shown)} of {len(pt)} {kind.lower()} periods. "
                     f"The full series is in the Excel workbook." if len(pt) > len(shown) else "")})
    extra = []
    if fac.get("mixed"):
        bd = facility_breakdown(fdf, unit)
        extra.append({"title": "Breakdown by facility type", "headers": bd[0], "rows": bd[1:],
                      "widths": [3.2, 1, 0.9, 1.8, 1.9, 1.1, 1.2, 1.4], "size": 7.5})
    return {"facility": fac, "facility_kpis_": an["kpis"], "findings": an["findings"],
            "extra_tables": extra,
            "period_sections": sections, "charts": [chart_nodes(fdf, unit)],
            "methodology": methodology_text(fac["code"], unit), "_analysis": an}


@st.cache_data(show_spinner="Rendering high-resolution PDF…")
def cached_pdf_extras(fdf: pd.DataFrame, meta: Dict, tol_pct: float, unit: str,
                      override: str, issues: Optional[pd.DataFrame]) -> Dict:
    return pdf_extras(fdf, meta, tol_pct, unit, override, issues)


@st.cache_data(show_spinner=False)
def list_sheets(data: bytes, filename: str) -> List[str]:
    """Sheet names for Excel workbooks (empty list for CSV / text files)."""
    kind = _sniff_kind(data, filename)
    if kind not in ("xlsx", "xls"):
        return []
    try:
        engine = "openpyxl" if kind == "xlsx" else "xlrd"
        return pd.ExcelFile(io.BytesIO(data), engine=engine).sheet_names
    except Exception:
        return []


# ==========================================================================
# DATA INGESTION & CLEANING ENGINE
# --------------------------------------------------------------------------
#  Stage 1  read_upload()    structure : file sniffing, header detection, merged /
#                                        repeated headers, totals & footers
#  Stage 2  harmonize()      layout    : per-sheet column mapping so files with
#                                        different layouts can be combined
#  Stage 3  clean_dataset()  values    : text, numbers, dates, units, names,
#                                        duplicates, empty records
#  Stage 4  build_balance()  analysis  : material balance, tolerance, issues
# ==========================================================================
NULL_TOKENS = {
    "", "nan", "none", "null", "nil", "n/a", "na", "n.a", "n.a.", "-", "--", "---",
    "\u2014", "\u2013", "#n/a", "#ref!", "#div/0!", "#value!", "#name?", "#null!",
    "#num!", "tbc", "tbd", "?", "not available", "missing", "unknown",
}
NUM_FIELDS = ["opening_stock", "receipts", "transfers_in", "sales", "transfers_out",
              "closing_stock", "gains_losses", "throughput"]
FLOW_FIELDS = ["receipts", "transfers_in", "sales", "transfers_out"]
EXC_STATUSES = ("Gain (Over)", "Loss (Short)")      # statuses that count as exceptions

_ZERO_WIDTH = re.compile(r"[\u200b-\u200d\u2060\ufeff]")
_TOTAL_ROW = re.compile(
    r"^(grand\s+|sub[\s\-]?|overall\s+|net\s+)?totals?\s*:?$|^totals?\s+(for|of)\b.*|^sum\s*:?$",
    re.I)
_FOOTER_ROW = re.compile(
    r"^(prepared|checked|approved|verified|printed|signature|signed|source|note|notes|"
    r"remarks?|report generated|generated|page\s+\d+|\*)", re.I)
_SCI = re.compile(r"^[-+]?\d*\.?\d+[eE][-+]?\d+$")
_UNIT_WORDS = re.compile(
    r"(?i)(?<![A-Za-z])(litres?|liters?|ltrs?|ltr|lts|kl|m3|m\u00b3|cbm|cum|gal(?:lons?)?|"
    r"bbls?|barrels?|usd|ghs|ghc|cedis?|mt|tonnes?|tons?|l)(?![A-Za-z])")

_UNIT_ALIAS: Dict[str, str] = {}
for _key, _names in {
    "L": ["L", "LT", "LTR", "LTRS", "LTS", "LITRE", "LITRES", "LITER", "LITERS"],
    "KL": ["KL", "KLTR", "KILOLITRE", "KILOLITRES", "KILOLITER", "KILOLITERS"],
    "M3": ["M3", "CBM", "CUM", "CUBICMETRE", "CUBICMETRES", "CUBICMETER", "CUBICMETERS"],
    "GAL": ["GAL", "GALS", "GALLON", "GALLONS", "USGAL", "USG"],
    "BBL": ["BBL", "BBLS", "BARREL", "BARRELS"],
}.items():
    for _n in _names:
        _UNIT_ALIAS[_n] = _key
UNIT_CONVERSIONS.setdefault("KL", 1000.0)
UNIT_CONVERSIONS.setdefault("BBLS", 158.987)

_PRODUCT_SYNONYMS = {
    "PREMIUM":   ["premium", "pms", "petrol", "gasoline", "premium motor spirit",
                  "motor spirit", "super", "unleaded", "premium pms"],
    "GASOIL":    ["gasoil", "gas oil", "diesel", "ago", "automotive gas oil", "hsd",
                  "diesel oil"],
    "KEROSENE":  ["kerosene", "dpk", "household kerosene", "hhk", "paraffin"],
    "ATK":       ["atk", "jet a1", "jeta1", "jet fuel", "avtur",
                  "aviation turbine kerosene"],
    "LPG":       ["lpg", "lp gas", "liquefied petroleum gas", "cooking gas"],
    "RFO":       ["rfo", "hfo", "fuel oil", "residual fuel oil", "heavy fuel oil"],
    "MGO":       ["mgo", "marine gas oil"],
    "NAPHTHA":   ["naphtha"],
    "CRUDE OIL": ["crude", "crude oil"],
}
PRODUCT_ALIASES = {normalize(a): canon for canon, al in _PRODUCT_SYNONYMS.items()
                   for a in al + [canon]}

_DEPT_ALIASES: Dict[str, str] = {"ict": "IT", "depots": "DEPOT", "imports": "IMPORT",
                                 "refining": "REFINERY", "hse": "HSE"}
for _code, _m in DEPARTMENTS.items():
    for _a in (_code, _m["label"], _m["short"]):
        _DEPT_ALIASES[normalize(_a)] = _code


# ------------------------------------------------------------------ primitives
def _is_blank(x) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(x, str) and x.strip().lower() in NULL_TOKENS


def clean_text_value(x) -> str:
    """Unicode-normalise, strip invisible characters, collapse whitespace, null → ''."""
    if _is_blank(x):
        return ""
    t = unicodedata.normalize("NFKC", str(x))
    t = _ZERO_WIDTH.sub("", t)
    t = re.sub(r"\s+", " ", t).strip()
    if t.lower() in NULL_TOKENS:
        return ""
    if re.fullmatch(r"-?\d+\.0+", t):              # 101.0 → 101 (Excel float artefacts)
        t = t.split(".")[0]
    return t


def clean_text(s: pd.Series) -> pd.Series:
    return s.map(clean_text_value).astype(object)


_EURO_EVIDENCE = re.compile(r"(\d,\d{1,2}$)|(\d\.\d{3}(\.\d{3})*,\d+$)")
_US_EVIDENCE = re.compile(r"(\d\.\d{1,2}$)|(\d,\d{3}(,\d{3})*\.\d+$)")


def detect_euro(values) -> bool:
    """True when text numbers use '.' for thousands and ',' for decimals (1.234,50)."""
    eu = us = 0
    for v in values:
        if isinstance(v, str):
            s = v.strip()
            if _EURO_EVIDENCE.search(s):
                eu += 1
            elif _US_EVIDENCE.search(s):
                us += 1
    return eu > us


def parse_number(x, euro: bool = False) -> float:
    """Tolerant numeric parser.

    Handles 1,234.50 · 1.234,50 · 1 234 567 · (1,234) · 1,234- · 12 BBL · 'GHS 1,200' ·
    Excel errors (#N/A …) and placeholders ('-', 'nil', 'n/a') → NaN.
    Dates, IDs and free text are rejected (NaN) rather than mangled into numbers.
    """
    if x is None or isinstance(x, (bool, np.bool_)):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        v = float(x)
        return v if np.isfinite(v) else np.nan
    if isinstance(x, (dt.datetime, dt.date, pd.Timestamp)):
        return np.nan
    t = unicodedata.normalize("NFKC", str(x)).strip().replace("\u2212", "-")
    if t.lower() in NULL_TOKENS:
        return np.nan
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1].strip()
    t = _UNIT_WORDS.sub(" ", t)
    t = re.sub(r"[$\u20ac\u00a3\u00a2\u20b5]", "", t).strip()
    compact = t.replace(" ", "")
    if _SCI.fullmatch(compact):
        try:
            return float(compact)
        except ValueError:
            return np.nan
    t = re.sub(r"[^\d,\.\-\+\s']", "", t).strip()
    if not re.search(r"\d", t):
        return np.nan
    if t.endswith("-") and not t.startswith(("-", "+")):
        neg, t = (not neg), t[:-1].strip()
    sign = 1.0
    if t and t[0] in "+-":
        sign = -1.0 if t[0] == "-" else 1.0
        t = t[1:].strip()
    if "-" in t or "+" in t:                       # '2024-01-05', '12-34' → not a number
        return np.nan
    t = t.replace(" ", "").replace("'", "")
    if euro:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t and "." in t:
        dec = "," if t.rfind(",") > t.rfind(".") else "."
        t = t.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in t:
        parts = t.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:   # 12,5 → decimal comma
            t = parts[0] + "." + parts[1]
        else:                                        # 1,234 / 1,234,567 → thousands
            t = t.replace(",", "")
    elif t.count(".") > 1:                           # 1.234.567 → thousands
        t = t.replace(".", "")
    try:
        v = float(t) * sign * (-1.0 if neg else 1.0)
    except ValueError:
        return np.nan
    return v if np.isfinite(v) else np.nan


def parse_numeric_series(s: pd.Series, euro=False) -> pd.Series:
    """`euro` is a bool or a per-row boolean Series (decimal-comma files)."""
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        return pd.to_numeric(s, errors="coerce").astype(float).replace([np.inf, -np.inf], np.nan)
    flags = (euro.reindex(s.index).fillna(False).tolist() if isinstance(euro, pd.Series)
             else [bool(euro)] * len(s))
    cache: Dict = {}
    out = []
    for v, e in zip(s.tolist(), flags):
        k = (v if isinstance(v, str) else (type(v).__name__, str(v)), e)
        if k not in cache:
            cache[k] = parse_number(v, e)
        out.append(cache[k])
    return pd.Series(out, index=s.index, dtype=float)


_ISO_RE = re.compile(r"^\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}")
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def _parse_date_one(x, dayfirst: bool):
    if _is_blank(x):
        return pd.NaT
    ts = pd.NaT
    try:
        if isinstance(x, (pd.Timestamp, dt.datetime, dt.date)):
            ts = pd.Timestamp(x)
        elif isinstance(x, (int, float, np.integer, np.floating)) and not isinstance(x, bool):
            v = float(x)
            if 20000 <= v <= 80000:                          # Excel serial date
                ts = _EXCEL_EPOCH + pd.to_timedelta(v, unit="D")
            elif v == int(v) and 19000101 <= v <= 21001231:  # 20240105
                ts = pd.to_datetime(str(int(v)), format="%Y%m%d", errors="coerce")
        else:
            t = clean_text_value(x)
            if re.fullmatch(r"\d{8}", t):
                ts = pd.to_datetime(t, format="%Y%m%d", errors="coerce")
            elif re.fullmatch(r"\d{5}(\.\d+)?", t):
                ts = _EXCEL_EPOCH + pd.to_timedelta(float(t), unit="D")
            elif _ISO_RE.match(t):                           # ISO is never day-first
                ts = pd.to_datetime(t.replace("/", "-").replace(".", "-"),
                                    errors="coerce", dayfirst=False)
            else:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    ts = pd.to_datetime(t, errors="coerce", dayfirst=dayfirst)
    except Exception:
        return pd.NaT
    if ts is pd.NaT or pd.isna(ts):
        return pd.NaT
    if ts.tzinfo is not None:
        ts = ts.tz_localize(None)
    return ts.normalize()


def parse_dates(s: pd.Series, dayfirst: bool = True) -> pd.Series:
    cache: Dict = {}
    vals = []
    for v in s.tolist():
        k = (type(v).__name__, str(v))
        if k not in cache:
            cache[k] = _parse_date_one(v, dayfirst)
        vals.append(cache[k])
    return pd.to_datetime(pd.Series(vals, index=s.index), errors="coerce")


def canon_unit(x) -> str:
    t = clean_text_value(x).upper().translate(str.maketrans({"\u00b3": "3", "\u00b2": "2"}))
    t = re.sub(r"[\s\.\-_/'\u2019]", "", t)
    return _UNIT_ALIAS.get(t, "")


def unit_from_header(name: str) -> str:
    """'Receipts (BBL)' → 'BBL' ; 'Sales [m3]' → 'M3' ; '' when none."""
    for m in re.finditer(r"[\(\[\{]\s*([^\)\]\}]{1,18}?)\s*[\)\]\}]", str(name)):
        u = canon_unit(re.sub(r"(?i)'?000s?|thousand", "", m.group(1)))
        if u:
            return u
    m = re.search(r"(?i)(?<![A-Za-z])(bbls?|barrels?|gal(?:lons?)?|m3|m\u00b3|cbm|kl|"
                  r"litres?|liters?|ltrs?)(?![A-Za-z])", str(name))
    return canon_unit(m.group(1)) if m else ""


def scale_from_header(name: str) -> float:
    return 1000.0 if re.search(r"(?i)(?:'|\u2019|\(|\s)000s?\b|thousand", str(name)) else 1.0


def _log(log: List[dict], stage: str, action: str, count, detail: str = "",
         source: str = "All files") -> None:
    try:
        n = int(count)
    except Exception:
        n = 0
    if n:
        log.append({"Stage": stage, "Source": source, "Action": action,
                    "Affected": n, "Detail": detail})


# ---------------------------------------------------------------- Stage 1: read
def _sniff_kind(data: bytes, filename: str) -> str:
    if data[:2] == b"PK":
        return "xlsx"
    if data[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"
    head = data[:2048].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html", b"<table", b"<?xml")):
        return "html"
    return "text"


def _decode_text(data: bytes) -> str:
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return data.decode("utf-16", errors="replace")
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _text_to_grid(text: str) -> pd.DataFrame:
    sample = text[:20000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(io.StringIO(text), dialect))
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        return pd.DataFrame()
    rows = [r + [None] * (width - len(r)) for r in rows]
    return pd.DataFrame(rows, dtype=object)


def _clean_header(v, j: int) -> str:
    t = clean_text_value(v) if not isinstance(v, (int, float)) else str(v).strip()
    return t if t else f"Column {j + 1}"


def _make_unique(names: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out = []
    for n in names:
        if n in seen:
            seen[n] += 1
            out.append(f"{n} ({seen[n]})")
        else:
            seen[n] = 1
            out.append(n)
    return out


def _find_header_row(grid: pd.DataFrame, max_scan: int = 30) -> Tuple[int, int]:
    """Return (row_position, recognised_field_count) of the most header-like row."""
    best_i, best_score = 0, -1
    for i in range(min(len(grid), max_scan)):
        cells = [str(v).strip() for v in grid.iloc[i].tolist()
                 if isinstance(v, str) and v.strip()]
        if len(cells) < 2:
            continue
        score = sum(1 for v in detect_columns(cells).values() if v)
        if score > best_score:
            best_i, best_score = i, score
    if best_score >= 2:
        return best_i, best_score
    # fallback: first row (in the top 15) that is mostly text and about as wide as the data
    counts = grid.head(15).notna().sum(axis=1)
    widest = int(counts.max()) if len(counts) else 0
    for i in range(min(len(grid), 15)):
        row = grid.iloc[i]
        n_text = sum(isinstance(v, str) for v in row.tolist())
        if widest >= 2 and counts.iloc[i] >= 0.6 * widest and n_text >= 0.6 * counts.iloc[i]:
            return i, max(best_score, 0)
    return 0, max(best_score, 0)


def _grid_to_frame(grid: pd.DataFrame, header_row: int, filename: str, sheet: str,
                   drop_totals: bool, log: List[dict]) -> Tuple[Optional[pd.DataFrame], int]:
    grid = grid.reset_index(drop=True)
    grid.columns = range(grid.shape[1])
    grid = grid.replace(r"^\s*$", np.nan, regex=True)
    if grid.dropna(how="all").empty:
        return None, 0
    src = f"{filename} › {sheet}"
    if header_row < 0:
        hdr, score = _find_header_row(grid)
        if hdr > 0:
            _log(log, "Structure", "Skipped title / preamble rows above the header", hdr,
                 f"Header detected on row {hdr + 1}", src)
    else:
        hdr = header_row
        score = sum(1 for v in detect_columns(
            [str(v) for v in grid.iloc[hdr].tolist() if isinstance(v, str)]).values() if v) \
            if hdr < len(grid) else 0
    if hdr >= len(grid):
        _log(log, "Structure", "Header row is beyond the end of the sheet", 1, "Sheet ignored", src)
        return None, 0

    names = _make_unique([_clean_header(v, j) for j, v in enumerate(grid.iloc[hdr].tolist())])
    body = grid.iloc[hdr + 1:].copy()
    body.columns = names
    body["__row"] = body.index + 1                       # Excel row number (1-based)
    body = body.dropna(how="all", subset=names)
    keep = [c for c in names if body[c].notna().any()]
    if not keep or body.empty:
        return None, score
    dropped_cols = len(names) - len(keep)
    _log(log, "Structure", "Removed completely empty columns", dropped_cols, "", src)
    body = body[keep + ["__row"]]

    # repeated header rows (page breaks / pasted blocks)
    hdr_norm = {c: normalize(c) for c in keep}
    matches = pd.Series(0, index=body.index)
    for c in keep:
        if hdr_norm[c]:
            matches += (body[c].map(lambda v: normalize(v) if not _is_blank(v) else "")
                        == hdr_norm[c]).astype(int)
    nn = body[keep].notna().sum(axis=1)
    rep = (matches >= 2) & (matches >= 0.5 * nn)
    if rep.any():
        _log(log, "Structure", "Removed repeated header rows", rep.sum(), "", src)
        body = body.loc[~rep]

    if drop_totals and not body.empty:
        text_cols = keep[:3]
        is_total = pd.Series(False, index=body.index)
        for c in text_cols:
            is_total |= body[c].map(
                lambda v: isinstance(v, str) and bool(_TOTAL_ROW.match(v.strip()))).astype(bool)
        n_non = body[keep].notna().sum(axis=1)
        is_footer = pd.Series(False, index=body.index)
        for c in keep:
            is_footer |= (n_non == 1) & body[c].map(
                lambda v: isinstance(v, str) and bool(_FOOTER_ROW.match(v.strip()))).astype(bool)
        _log(log, "Structure", "Removed total / subtotal rows", is_total.sum(),
             "Prevents double-counting", src)
        _log(log, "Structure", "Removed footer / signature rows", (is_footer & ~is_total).sum(),
             "", src)
        body = body.loc[~(is_total | is_footer)]
    if body.empty:
        return None, score
    return body.reset_index(drop=True), score


@st.cache_data(show_spinner=False)
def read_upload(data: bytes, filename: str, sheets: Tuple[str, ...],
                header_row: int = -1, drop_totals: bool = True
                ) -> Tuple[pd.DataFrame, List[dict]]:
    """Read ANY tabular upload (xlsx/xlsm/xls/csv/tsv/txt/html-as-xls) into a tidy frame.

    header_row = -1 → auto-detect per sheet. Returns (frame, structural cleaning log).
    """
    log: List[dict] = []
    if not data:
        raise ValueError("The file is empty.")
    kind = _sniff_kind(data, filename)
    grids: Dict[str, pd.DataFrame] = {}

    if kind in ("xlsx", "xls"):
        engine = "openpyxl" if kind == "xlsx" else "xlrd"
        try:
            book = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None,
                                 dtype=object, engine=engine)
        except ImportError as exc:
            raise ValueError(f"Missing library to read this file type: {exc}")
        except Exception as exc:
            raise ValueError(
                f"Could not open the workbook ({type(exc).__name__}). It may be corrupt, "
                "password-protected, or an unsupported format such as .xlsb — please re-save "
                "it as .xlsx and try again.")
        grids = {str(k): v for k, v in book.items()}
    elif kind == "html":
        try:
            tables = pd.read_html(io.StringIO(_decode_text(data)), header=None)
        except Exception:
            raise ValueError("This file is HTML disguised as a spreadsheet and could not be "
                             "parsed. Open it in Excel and save it as .xlsx.")
        grids = {f"Table {i + 1}": t.astype(object) for i, t in enumerate(tables)}
    else:
        text = _decode_text(data)
        grid = _text_to_grid(text)
        if grid.empty:
            raise ValueError("No tabular data found in this file.")
        grids = {"(text)": grid}

    frames: List[Tuple[pd.DataFrame, int, str]] = []
    skipped_empty: List[str] = []
    for sheet_name, grid in grids.items():
        if kind in ("xlsx", "xls") and sheets and sheet_name not in sheets:
            continue
        frame, score = _grid_to_frame(grid, header_row, filename, sheet_name, drop_totals, log)
        if frame is None or frame.empty:
            skipped_empty.append(sheet_name)
            continue
        frames.append((frame, score, sheet_name))
    if skipped_empty:
        _log(log, "Structure", "Ignored empty sheets", len(skipped_empty),
             ", ".join(skipped_empty)[:200], filename)
    if not frames:
        raise ValueError("No usable data found (all sheets empty or deselected).")

    # drop sheets that look like notes / cover pages when real data sheets exist
    if header_row < 0 and any(sc >= 2 for _, sc, _ in frames):
        good = [(f, sc, n) for f, sc, n in frames if sc >= 2]
        dropped = [n for _, sc, n in frames if sc < 2]
        if dropped:
            _log(log, "Structure", "Ignored non-data sheets (no recognisable columns)",
                 len(dropped), ", ".join(dropped)[:200], filename)
        frames = good

    out = []
    for frame, _, sheet_name in frames:
        frame = frame.copy()
        frame["__source_file"] = filename
        frame["__sheet"] = sheet_name
        out.append(_normalize_string_dtypes(frame))
    return pd.concat(out, ignore_index=True, sort=False), log


# ------------------------------------------------------------- Stage 2: layout
def harmonize(raw: pd.DataFrame, mapping: Dict[str, Optional[str]],
              log: List[dict]) -> pd.DataFrame:
    """Build one canonical frame (a column per field) from possibly different sheet layouts.

    The user's column mapping is honoured wherever that column exists in a sheet; where it
    does not (another file used different headings) the sheet's own auto-detected column is
    used, so files with different layouts can be combined without losing data.
    """
    parts = []
    for (fname, sheet), sub in raw.groupby(["__source_file", "__sheet"], sort=False):
        gcols = [c for c in sub.columns
                 if not str(c).startswith("__") and sub[c].notna().any()]
        part = pd.DataFrame(index=sub.index)
        part["__source_file"] = fname
        part["__sheet"] = sheet
        part["__row"] = sub["__row"] if "__row" in sub.columns else np.arange(len(sub)) + 1
        used: set = set()
        for f in PRIORITY:
            want = mapping.get(f)
            chosen = None
            if want and want in gcols and want not in used:
                chosen = want
            elif want:                                   # mapped globally, absent here
                rest = [c for c in gcols if c not in used]
                chosen = detect_columns(rest).get(f) if rest else None
                if chosen:
                    _log(log, "Layout", f"Used '{chosen}' for {FIELD_LABELS[f]}", len(sub),
                         f"Mapped column '{want}' not present in this sheet",
                         f"{fname} › {sheet}")
            if chosen:
                used.add(chosen)
                part[f] = sub[chosen]
                if f in NUM_FIELDS:
                    part[f"__hu_{f}"] = unit_from_header(chosen)
                    part[f"__sc_{f}"] = scale_from_header(chosen)
            else:
                part[f] = np.nan
                if f in NUM_FIELDS:
                    part[f"__hu_{f}"] = ""
                    part[f"__sc_{f}"] = 1.0
        parts.append(part)
    canon = pd.concat(parts, sort=False)
    for f in NUM_FIELDS:
        canon[f"__hu_{f}"] = canon[f"__hu_{f}"].fillna("")
        canon[f"__sc_{f}"] = canon[f"__sc_{f}"].fillna(1.0)
    return canon


# ------------------------------------------------------------- Stage 3: values
def clean_dataset(raw: pd.DataFrame, mapping: Dict[str, Optional[str]],
                  opts: Dict) -> Tuple[pd.DataFrame, Dict[str, bool], List[dict], Dict]:
    """Return (clean canonical frame in LITRES, field-presence flags, log, stats)."""
    unit_override = opts.get("unit_override", "Auto")
    fill_merged = opts.get("fill_merged", True)
    drop_dupes = opts.get("drop_dupes", True)
    std_names = opts.get("std_names", True)
    dayfirst = opts.get("dayfirst", True)
    log: List[dict] = list(opts.get("structure_log", []))
    stats: Dict = {"rows_in": len(raw)}

    canon = harmonize(raw, mapping, log)
    present = {f: bool(canon[f].notna().any()) for f in PRIORITY}

    # ---- text fields ------------------------------------------------------
    for f in ("depot", "product", "department", "node_type", "reference", "units"):
        canon[f] = clean_text(canon[f])

    # ---- dates ------------------------------------------------------------
    raw_dates = canon["date"].copy()
    canon["date"] = parse_dates(canon["date"], dayfirst)
    bad = raw_dates.map(lambda v: not _is_blank(v)) & canon["date"].isna()
    _log(log, "Values", "Unreadable dates left blank", bad.sum(),
         "e.g. " + ", ".join(map(str, raw_dates[bad].head(3).tolist())))
    lo, hi = pd.Timestamp("2000-01-01"), pd.Timestamp(dt.date.today()) + pd.Timedelta(days=366)
    oob = canon["date"].notna() & ((canon["date"] < lo) | (canon["date"] > hi))
    _log(log, "Values", "Out-of-range dates (before 2000 / far future) left blank", oob.sum())
    canon.loc[oob, "date"] = pd.NaT

    # ---- fill blanks from merged / grouped cells ---------------------------------
    if fill_merged:
        grp = [canon["__source_file"], canon["__sheet"]]
        for f in ("depot", "product", "department", "node_type", "date"):
            if not present[f]:
                continue
            s = canon[f].replace("", np.nan) if f != "date" else canon[f]
            filled = s.groupby(grp, sort=False).ffill()
            n = int((s.isna() & filled.notna()).sum())
            if n:
                canon[f] = filled.fillna("") if f != "date" else filled
                _log(log, "Values", f"Filled blank {FIELD_LABELS[f]} from the cell above", n,
                     "Merged / grouped cells")

    # ---- numbers + units ----------------------------------------------------
    row_unit_raw = canon["units"]
    row_unit = row_unit_raw.map(canon_unit)
    unknown_unit = (row_unit_raw != "") & (row_unit == "")
    _log(log, "Values", "Unrecognised unit text — assumed litres", unknown_unit.sum(),
         "e.g. " + ", ".join(sorted(set(row_unit_raw[unknown_unit].head(5)))))
    default_unit = "L"
    euro_files = {}
    for fname, sub in canon.groupby("__source_file", sort=False):
        vals = [v for f in NUM_FIELDS if present[f] for v in sub[f].tolist()
                if isinstance(v, str)]
        euro_files[fname] = detect_euro(vals)
    euro_row = canon["__source_file"].map(euro_files).astype(bool)
    n_euro = [k for k, v in euro_files.items() if v]
    if n_euro:
        _log(log, "Values", "Detected decimal-comma number format (1.234,50)", len(n_euro),
             ", ".join(n_euro)[:200])
    conv_counter: Dict[str, int] = {}
    for f in NUM_FIELDS:
        if not present[f]:
            canon[f] = np.nan
            continue
        vals = parse_numeric_series(canon[f], euro_row)
        unparsable = canon[f].map(lambda v: not _is_blank(v)) & vals.isna()
        _log(log, "Values", f"Non-numeric entries in {FIELD_LABELS[f]} treated as blank",
             unparsable.sum(),
             "e.g. " + ", ".join(map(str, canon.loc[unparsable, f].head(3).tolist())))
        if unit_override != "Auto":
            eff = pd.Series(unit_override, index=canon.index)
        else:
            eff = row_unit.where(row_unit != "", canon[f"__hu_{f}"])
            eff = eff.where(eff != "", default_unit)
        factor = eff.map(lambda u: UNIT_CONVERSIONS.get(u, 1.0)).astype(float)
        scale = pd.to_numeric(canon[f"__sc_{f}"], errors="coerce").fillna(1.0)
        canon[f] = vals * factor * scale
        if f == "receipts" or (f not in ("receipts",) and not conv_counter):
            for u, n in eff[vals.notna()].value_counts().items():
                conv_counter[u] = conv_counter.get(u, 0) + int(n)
        if (scale != 1.0).any():
            _log(log, "Values", f"{FIELD_LABELS[f]} header says thousands — multiplied by 1,000",
                 int((scale != 1.0).sum()))
    converted = {u: n for u, n in conv_counter.items() if u != "L"}
    if converted:
        _log(log, "Values", "Rows converted to litres", sum(converted.values()),
             ", ".join(f"{u}: {n:,}" for u, n in converted.items()))
    stats["source_units"] = (", ".join(f"{u} ({n:,})" for u, n in
                                       sorted(conv_counter.items(), key=lambda kv: -kv[1]))
                             or "L")

    # ---- drop rows that carry no volume figures at all ----------------------
    num_present = [f for f in NUM_FIELDS if present[f]]
    if not num_present:
        raise ValueError("No volume columns were found (opening, receipts, sales, closing …). "
                         "Open the Column Mapping panel and map them, or check the header row.")
    if num_present:
        no_vol = canon[num_present].isna().all(axis=1)
        _log(log, "Rows", "Removed rows with no volume figures (blank lines, labels, notes)",
             no_vol.sum())
        canon = canon.loc[~no_vol]

    # ---- department / depot / product standardisation ---------------------------
    dept = canon["department"].map(lambda v: _DEPT_ALIASES.get(normalize(v), ""))
    canon["department"] = dept
    canon["depot"] = canon["depot"].replace("", "Unspecified")
    if std_names:
        before = canon["depot"].nunique()
        key = canon["depot"].map(normalize)
        best = (canon.assign(_k=key).groupby("_k")["depot"]
                .agg(lambda s: s.value_counts().index[0]))
        canon["depot"] = key.map(best)
        _log(log, "Values", "Merged depot spelling / case variants", before - canon["depot"].nunique(),
             "e.g. 'TEMA', 'Tema ', 'tema' → one depot")
        prod_raw = canon["product"].replace("", "UNSPECIFIED").str.upper()
        prod = prod_raw.map(lambda p: PRODUCT_ALIASES.get(normalize(p), p))
        _log(log, "Values", "Standardised product names", (prod != prod_raw).sum(),
             "e.g. PMS / Petrol → PREMIUM, AGO / Diesel → GASOIL")
        canon["product"] = prod
    else:
        canon["product"] = canon["product"].replace("", "UNSPECIFIED").str.upper()

    # ---- exact duplicate records ---------------------------------------------
    key_cols = ["date", "depot", "product", "reference"] + [f for f in NUM_FIELDS if present[f]]
    kf = canon[key_cols].copy()
    for f in key_cols:
        if f in NUM_FIELDS:
            kf[f] = kf[f].round(3)
    nonzero = canon[num_present].fillna(0).abs().sum(axis=1) > 0 if num_present else \
        pd.Series(False, index=canon.index)
    dup = kf.duplicated(keep="first") & nonzero
    stats["exact_duplicates"] = int(dup.sum())
    if drop_dupes and dup.any():
        _log(log, "Rows", "Removed exact duplicate records", dup.sum(),
             "Same date, depot, product, reference and every volume figure")
        canon = canon.loc[~dup]
    stats["rows_out"] = len(canon)
    return canon, present, log, stats


# ------------------------------------------------------------ Stage 4: analysis
@st.cache_data(show_spinner="Cleaning data and building the material balance…")
def build_balance(raw: pd.DataFrame, mapping: Dict[str, Optional[str]],
                  tol_pct: float, tol_abs: float, opts: Optional[Dict] = None
                  ) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, str], pd.DataFrame]:
    """Clean `raw`, then compute the material balance.

    Returns (balance, data-quality issues, meta, cleaning log). All volumes are in LITRES.
    """
    opts = dict(opts or {})
    meta: Dict[str, str] = {}
    canon, present, log, stats = clean_dataset(raw, mapping, opts)
    idx = canon.index
    out = pd.DataFrame(index=idx)

    out["Date"] = canon["date"]
    out["Depot"] = canon["depot"]
    out["Product"] = canon["product"]
    out["Reference"] = canon["reference"]

    # ---- facility auto-detection: explicit column > strong node-name keyword >
    #      file-level evidence (file/sheet name, headers, products) > Depot
    file_det = detect_files(raw)
    file_code = canon["__source_file"].astype(str).map(
        {f: d["code"] for f, d in file_det.items()}).fillna("DEPOT")
    name_map = {n: infer_department_ex(n) for n in out["Depot"].dropna().unique()}
    row_code = out["Depot"].map(lambda n: name_map.get(n, (None, False))[0])
    row_strong = out["Depot"].map(lambda n: name_map.get(n, (None, False))[1]).fillna(False).astype(bool)
    inferred = row_code.where(row_strong, file_code).fillna(file_code)
    out["Department"] = canon["department"].where(canon["department"] != "", inferred)
    n_inferred = int((canon["department"] == "").sum())
    meta["facility_detection"] = json.dumps(file_det)
    meta["department_source"] = ("column" if n_inferred == 0 else
                                 "mixed" if n_inferred < len(canon) else "inferred")

    nt = canon["node_type"].str.upper()
    default_nt = out["Department"].map(
        lambda d: {"DEPOT": "DEPOT", "IMPORT": "VESSEL", "MARINE": "JETTY",
                   "REFINERY": "REFINERY", "RETAIL": "RETAIL",
                   "TRANSPORT": "PIPELINE/FLEET"}.get(d, "NODE"))
    out["Node Type"] = nt.where(nt != "", default_nt)

    # all volumes are litres from here on
    meta["unit"] = "L"
    meta["source_units"] = stats["source_units"]
    meta["rows_in"] = f"{stats['rows_in']:,}"
    meta["rows_analysed"] = f"{stats['rows_out']:,}"

    def num(field: str) -> pd.Series:
        return canon[field].astype(float) if present[field] else pd.Series(np.nan, index=idx)

    opening_raw, closing_raw, gl_raw = num("opening_stock"), num("closing_stock"), num("gains_losses")
    through_given = num("throughput").fillna(0.0)
    receipts, t_in = num("receipts").fillna(0.0), num("transfers_in").fillna(0.0)
    sales, t_out = num("sales").fillna(0.0), num("transfers_out").fillna(0.0)
    opening = opening_raw.fillna(0.0)

    has_closing, has_gl = present["closing_stock"], present["gains_losses"]
    available = opening + receipts + t_in
    expected = available - sales - t_out

    if has_closing and has_gl:
        variance = (closing_raw - expected).where(closing_raw.notna(), gl_raw)
        closing = closing_raw.where(closing_raw.notna(), expected + gl_raw)
    elif has_closing:
        variance = closing_raw - expected                 # blank closing → NaN, never 0
        closing = closing_raw
    elif has_gl:
        variance = gl_raw
        closing = expected + gl_raw
    else:
        variance = pd.Series(np.nan, index=idx)
        closing = pd.Series(np.nan, index=idx)

    throughput = receipts + sales + t_in + t_out
    throughput = throughput.where(throughput != 0, through_given)
    throughput = throughput.where(throughput != 0, available)
    with np.errstate(divide="ignore", invalid="ignore"):
        var_pct = np.where(throughput > 0, variance.round(3) / throughput * 100.0, np.nan)

    out["Opening Stock"] = opening
    out["Receipts"] = receipts
    out["Transfers In"] = t_in
    out["Available"] = available
    out["Sales / Deliveries"] = sales
    out["Transfers Out"] = t_out
    out["Throughput"] = throughput
    out["Expected Closing"] = expected
    out["Actual Closing"] = closing
    out["Gain / (Loss)"] = variance.round(3)          # removes float noise from unit conversion
    out["Variance %"] = pd.Series(var_pct, index=idx).round(6)

    dept_tol = out["Department"].map(lambda d: DEPARTMENTS.get(d, {}).get("tol", tol_pct))
    dept_tol = pd.to_numeric(dept_tol, errors="coerce").fillna(tol_pct)
    out["Tolerance %"] = dept_tol

    verifiable = out["Gain / (Loss)"].notna()
    within = ((out["Variance %"].abs() <= dept_tol) | (out["Gain / (Loss)"].abs() <= tol_abs))
    out["Status"] = np.select(
        [~verifiable, within.fillna(False), out["Gain / (Loss)"] > 0],
        ["Unverified", "Within Tolerance", "Gain (Over)"], default="Loss (Short)")

    def _remark(status, gl, pct):
        if status == "Within Tolerance":
            return "Accepted — variance within tolerance"
        if status == "Unverified":
            return "Closing stock / gain-loss missing — balance cannot be verified"
        d = "Overage" if status.startswith("Gain") else "Shortage"
        pct_txt = f"{pct:,.2f}%" if pd.notna(pct) else "n/a"
        return f"{d} of {abs(gl):,.0f} L ({pct_txt}) — investigate"

    out["Remark"] = [_remark(s, g, p) for s, g, p in
                     zip(out["Status"], out["Gain / (Loss)"], out["Variance %"])]
    out["Source File"] = canon["__source_file"]
    out["Source Sheet"] = canon["__sheet"]
    out["Source Row"] = canon["__row"]

    # ------------------------------------------------------------------ issues
    issues: List[dict] = []

    def issue(sev: str, cat: str, title: str, detail: str) -> None:
        issues.append({"Severity": sev, "Category": cat, "Issue": title, "Detail": detail})

    if out.empty:
        issue("Critical", "No Data", "No usable rows after cleaning",
              "Check the column mapping and the header row.")
    if not has_closing and not has_gl:
        issue("Critical", "Balance Integrity", "No Closing Stock or Gain/Loss mapped",
              "Balance cannot be verified.")
    for f in REQUIRED_FIELDS:
        if not present[f]:
            issue("High", "Missing Column", f"{FIELD_LABELS[f]} not mapped",
                  "Rows default to zero for this field.")
    n_unver = int((~verifiable).sum())
    if n_unver and (has_closing or has_gl):
        issue("High", "Missing Closing", f"{n_unver:,} row(s) have no closing stock / gain-loss",
              "Excluded from variance totals instead of being counted as a 100% loss.")
    if present["opening_stock"]:
        n = int(opening_raw.isna().sum())
        if n:
            issue("Medium", "Missing Opening", f"{n:,} row(s) have a blank opening stock",
                  "Treated as zero — check the balance of these rows.")
        n = int((opening_raw < 0).sum())
        if n:
            issue("High", "Negative Stock", f"{n:,} row(s) with negative opening stock",
                  "Opening stock cannot be negative — check sign or posting.")
    if has_closing:
        n = int((closing_raw < 0).sum())
        if n:
            issue("High", "Negative Stock", f"{n:,} row(s) with negative closing stock",
                  "Closing stock cannot be negative — check sign or posting.")
    for f in FLOW_FIELDS:
        if present[f]:
            n = int((canon[f] < 0).sum())
            if n:
                issue("Medium", "Negative Movement",
                      f"{n:,} negative value(s) in {FIELD_LABELS[f]}",
                      "Kept as entered (possible reversals) — confirm they are intentional.")
    if has_closing and has_gl:
        calc = closing_raw - expected
        diff = ((calc - gl_raw).abs() > max(1.0, 0.0))
        n = int((diff & closing_raw.notna() & gl_raw.notna()).sum())
        if n:
            issue("Medium", "Inconsistent Figures",
                  f"{n:,} row(s) where reported gain/loss ≠ closing − expected",
                  "The computed value (closing − expected) is used.")
    dup = out.duplicated(subset=["Date", "Depot", "Product"], keep=False) & out["Date"].notna()
    if dup.any():
        issue("Medium", "Repeated Keys", f"{int(dup.sum()):,} rows share Date + Depot + Product",
              "Fine for multiple tanks; otherwise verify they are not double-posted.")
    if stats.get("exact_duplicates") and not opts.get("drop_dupes", True):
        issue("High", "Duplicate Records", f"{stats['exact_duplicates']:,} exact duplicate rows",
              "Duplicate removal is switched off — totals may be overstated.")
    if has_closing:
        big = max(tol_pct * 10, 5.0)
        n = int((out["Variance %"].abs() > big).sum())
        if n:
            issue("High", "Outliers", f"{n:,} row(s) with large variance (> {big:g}%)",
                  "Investigate for possible mis-postings or unit errors.")
    if n_inferred:
        issue("Info", "Demarcation",
              f"{out.loc[canon['department'] == '', 'Depot'].nunique()} depot(s) without a department",
              "Rows were auto-tagged using keyword inference.")
    if present["date"] and out["Date"].isna().any():
        issue("Medium", "Missing Dates", f"{int(out['Date'].isna().sum()):,} row(s) without a valid date",
              "Excluded from trend charts.")
    if (out["Product"] == "UNSPECIFIED").any():
        issue("Medium", "Missing Product", f"{int((out['Product'] == 'UNSPECIFIED').sum()):,} row(s) without a product",
              "Grouped under UNSPECIFIED.")
    # stock continuity: opening should equal the previous closing for the same depot/product
    if has_closing and present["opening_stock"] and out["Date"].notna().any():
        t = out.dropna(subset=["Date"]).sort_values("Date")
        prev = t.groupby(["Depot", "Product"])["Actual Closing"].shift(1)
        gap = (t["Opening Stock"] - prev).abs()
        n = int((gap > np.maximum(1.0, 0.0005 * t["Opening Stock"].abs())).sum())
        if n:
            issue("Info", "Continuity", f"{n:,} row(s) where opening ≠ previous closing",
                  "Possible missing days or unposted movements between consecutive records.")

    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["Severity", "Category", "Issue", "Detail"])
    if not issues_df.empty:
        issues_df["__o"] = issues_df["Severity"].map(SEVERITY_ORDER).fillna(9)
        issues_df = issues_df.sort_values("__o", kind="stable").drop(columns="__o").reset_index(drop=True)

    clean_log = pd.DataFrame(log, columns=["Stage", "Source", "Action", "Affected", "Detail"])
    return out, issues_df, meta, clean_log


# ==========================================================================
# Session bootstrap
# ==========================================================================
init_db()
for key, default in (("user", None), ("page", "Dashboard")):
    if key not in st.session_state:
        st.session_state[key] = default


# ==========================================================================
# AUTH
# ==========================================================================
def render_auth_screen() -> None:
    st.markdown(login_css(), unsafe_allow_html=True)
    left, right = st.columns([1.2, 1], gap="large")

    with left:
        hero = f"""
        <div class="hero">
          <div class="login-plate">
            {logo_html(64)}
            <div style="color:#fff">
              <div style="font-size:12.5px;letter-spacing:2px;font-weight:800">NATIONAL PETROLEUM AUTHORITY</div>
              <div style="font-size:11.5px;opacity:.65">Republic of Ghana</div>
            </div>
          </div>
          <div class="eyebrow"><i></i>Enterprise Platform &middot; Live</div>
          <h1 class="hero-title">Every litre,<br><span class="accent">accounted for.</span></h1>
          <p class="hero-sub">One secure platform to reconcile petroleum stock movements across every
             NPA department. Upload any Excel return &mdash; it is cleaned automatically and balanced in seconds.</p>
          <div class="stat-row">
            <div class="stat"><b>11</b><span>NPA departments</span></div>
            <div class="stat"><b>Any</b><span>Excel / CSV layout</span></div>
            <div class="stat"><b>Auto</b><span>Data cleaning</span></div>
          </div>
          <div class="preview">
            <div class="pv-head"><span>&#9981; Sample material balance &middot; PREMIUM</span><em>&#9679; Within tolerance</em></div>
            <div class="pv-eq">
              <div><small>Opening</small><b>1.00M</b></div>
              <div><small>+ Receipts</small><b>500K</b></div>
              <div><small>&minus; Sales</small><b>300K</b></div>
              <div><small>= Closing</small><b>1.198M</b></div>
            </div>
            <div class="pv-bar"><i></i></div>
            <div class="pv-foot"><span>Variance +0.13%</span><span>Tolerance &plusmn;0.35%</span></div>
          </div>
          <div class="hero-feat">
            <span>&#127981; Depots</span><span>&#128674; Imports</span><span>&#9875; Marine</span>
            <span>&#128293; Refinery</span><span>&#9981; Retail</span><span>&#128667; Transport</span>
            <span>&#129514; Quality</span><span>&#128176; Finance</span><span>&#128737;&#65039; Compliance</span>
            <span>&#128680; HSE</span><span>&#128187; ICT</span>
          </div>
          <div class="hero-trust">&#128274; Role-based access &middot; Encrypted credentials &middot; Full audit trail</div>
        </div>"""
        st.markdown("".join(line.strip() for line in hero.splitlines()), unsafe_allow_html=True)

    with right, st.container(key="login_card"):
        t_login, t_reg, t_help = st.tabs(["🔐  Sign In", "✨  Register", "ℹ️  Help"])

        with t_login:
            st.markdown('<div class="card-head"><div class="ico">&#128274;</div>'
                        '<h2>Welcome back</h2></div>'
                        '<p class="sub">Sign in with your email and staff ID to continue.</p>',
                        unsafe_allow_html=True)
            with st.form("login_form"):
                email = st.text_input("Email address", placeholder="you@npa.gov.gh")
                staff_id = st.text_input("Staff ID", placeholder="NPA-0001")
                password = st.text_input("Password", type="password", placeholder="••••••••")
                submit = st.form_submit_button("Sign In")
            if submit:
                if not (email and staff_id and password):
                    st.error("Please enter email, staff ID and password.")
                else:
                    rows = db_query(
                        "SELECT * FROM users WHERE lower(email)=lower(?) AND staff_id=?",
                        (email.strip(), staff_id.strip()))
                    if not rows:
                        verify_password(password, "00:00")      # equalise timing
                        st.error("Incorrect email, staff ID or password.")
                    else:
                        u = rows[0]
                        locked_until = u["locked_until"] if "locked_until" in u.keys() else None
                        if locked_until and locked_until > dt.datetime.now().isoformat(timespec="seconds"):
                            st.error(f"Too many failed attempts. Try again after "
                                     f"{locked_until[11:16]} (or ask a manager to reset).")
                        elif not verify_password(password, u["password_hash"]):
                            fails = int(u["failed_attempts"] or 0) + 1
                            if fails >= MAX_FAILED_LOGINS:
                                until = (dt.datetime.now() + dt.timedelta(minutes=LOCKOUT_MINUTES)
                                         ).isoformat(timespec="seconds")
                                db_execute("UPDATE users SET failed_attempts=0, locked_until=? WHERE id=?",
                                           (until, u["id"]))
                                log_action(u["email"], "login_locked", f"{fails} failed attempts")
                                st.error(f"Account locked for {LOCKOUT_MINUTES} minutes.")
                            else:
                                db_execute("UPDATE users SET failed_attempts=? WHERE id=?", (fails, u["id"]))
                                log_action(u["email"], "login_failed", f"attempt {fails}")
                                st.error("Incorrect email, staff ID or password.")
                        elif u["status"] == "pending":
                            st.warning("Account pending manager approval.")
                        elif u["status"] == "suspended":
                            st.error("Account suspended. Contact your manager.")
                        else:
                            db_execute("UPDATE users SET last_login=?, failed_attempts=0, locked_until=NULL WHERE id=?",
                                       (dt.datetime.now().isoformat(timespec="seconds"),
                                        u["id"]))
                            log_action(u["email"], "login", "successful sign-in")
                            st.session_state.user = dict(u)
                            st.rerun()

        with t_reg:
            st.markdown('<div class="card-head"><div class="ico">&#10024;</div>'
                        '<h2>Create your account</h2></div>'
                        '<p class="sub">Choose your role. Staff and company accounts '
                        'require manager approval.</p>', unsafe_allow_html=True)
            role = st.radio("I am registering as", ["staff", "company", "manager"],
                            format_func=lambda r: f'{ROLES[r]["icon"]} {ROLES[r]["label"]}',
                            horizontal=True, key="reg_role")
            companies = get_companies()
            company_options = {c["name"]: c["id"] for c in companies}

            with st.form("register_form"):
                full_name = st.text_input("Full name", placeholder="Ama Mensah")
                email = st.text_input("Email address", placeholder="you@company.com")
                staff_id = st.text_input("Staff ID", placeholder="NPA-1234")
                phone = st.text_input("Phone (optional)")
                password = st.text_input("Password", type="password")
                confirm = st.text_input("Confirm password", type="password")
                dept_choice = st.selectbox(
                    "Department",
                    [""] + DEPARTMENT_ORDER,
                    format_func=lambda d: "— none —" if d == ""
                    else f'{DEPARTMENTS[d]["icon"]} {DEPARTMENTS[d]["label"]}')
                invite = ""
                company_id: Optional[int] = None
                depot = ""
                new_company = ""
                company_code = ""
                if role == "manager":
                    invite = st.text_input("Manager invite code")
                elif role == "company":
                    new_company = st.text_input("Company name")
                    company_code = st.text_input("Company code (short)")
                else:
                    if company_options:
                        chosen = st.selectbox("Company", list(company_options.keys()))
                        company_id = company_options[chosen]
                    depot = st.text_input("Depot / Terminal")
                submit = st.form_submit_button("Create Account")

            if submit:
                errs = []
                if not full_name.strip():
                    errs.append("Full name is required.")
                if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()):
                    errs.append("Valid email address required.")
                if not staff_id.strip():
                    errs.append("Staff ID is required.")
                if len(password) < 8:
                    errs.append("Password must be at least 8 characters.")
                if password != confirm:
                    errs.append("Passwords do not match.")
                if role == "manager" and (not manager_invite_code() or not secrets.compare_digest(
                        invite.encode(), manager_invite_code().encode())):
                    errs.append("Invalid manager invite code.")
                if errs:
                    for e in errs:
                        st.error(e)
                else:
                    new_company_id = None
                    try:
                        if db_query("SELECT 1 FROM users WHERE lower(email)=lower(?)",
                                    (email.strip(),)):
                            raise ValueError("Email already registered.")
                        if db_query("SELECT 1 FROM users WHERE staff_id=?",
                                    (staff_id.strip(),)):
                            raise ValueError("Staff ID already registered.")
                        if role == "company":
                            if not new_company.strip():
                                raise ValueError("Company name required.")
                            if db_query("SELECT 1 FROM companies WHERE lower(name)=lower(?)",
                                        (new_company.strip(),)):
                                raise ValueError("A company with that name already exists.")
                            company_id = db_execute(
                                """INSERT INTO companies (name, code, contact_email, status)
                                   VALUES (?,?,?, 'pending')""",
                                (new_company.strip(), company_code.strip().upper() or None,
                                 email.strip()))
                            new_company_id = company_id
                            status = "pending"
                        elif role == "manager":
                            status = "active"
                        else:
                            status = "pending"
                        db_execute(
                            """INSERT INTO users (email, staff_id, password_hash, full_name,
                                                   role, company_id, depot, department,
                                                   phone, status)
                               VALUES (?,?,?,?,?,?,?,?,?,?)""",
                            (email.strip(), staff_id.strip(), hash_password(password),
                             full_name.strip(), role, company_id, depot.strip(),
                             dept_choice or None, phone.strip(), status))
                        log_action(email.strip(), "register", f"role={role} dept={dept_choice}")
                        st.success("Account created." if status == "active"
                                   else "Account created. Awaiting approval.")
                    except (sqlite3.IntegrityError, ValueError) as exc:
                        if new_company_id:                       # don't leave an orphan company
                            db_execute("DELETE FROM companies WHERE id=?", (new_company_id,))
                        st.error(f"Registration failed: {exc}" if isinstance(exc, sqlite3.IntegrityError)
                                 else str(exc))

        with t_help:
            st.markdown('<div class="card-head"><div class="ico">&#8505;&#65039;</div>'
                        '<h2>Getting started</h2></div>', unsafe_allow_html=True)
            st.markdown("""
**Staff** — Upload your depot's returns and view the balance.

**Company** — Manage your depots and staff, access consolidated reports.

**Manager** — Full oversight, approve registrations, audit activity.

**Departments covered**
Depot · Imports · Marine · Refinery · Retail · Transport · Quality · Finance · Compliance · HSE · ICT
""")
            if default_manager_unchanged() and _is_local_request():
                st.warning(f"First-run credentials (change after sign-in): "
                           f"`{DEFAULT_MANAGER_EMAIL}` · `NPA-0001` · `{DEFAULT_MANAGER_PASSWORD}`")

        st.markdown('<div class="card-foot">Need help? Contact your NPA system manager.</div>',
                    unsafe_allow_html=True)


# ==========================================================================
# MAIN APP
# ==========================================================================
def render_app(user: dict) -> None:
    st.markdown(base_css(), unsafe_allow_html=True)
    role = user["role"]
    role_meta = ROLES[role]
    user_dept = user.get("department") or ""

    st.markdown(f"""
    <div class="npa-header">
      {logo_html(58)}
      <div>
        <h1>Enterprise Material Balance Portal</h1>
        <small>{html.escape(role_meta['label'])} workspace ·
          {html.escape(user.get('company_name') or 'NPA Head Office')}
          {(' · ' + DEPARTMENTS.get(user_dept, {}).get('label','')) if user_dept else ''}</small>
      </div>
      <div class="who">
        <div><b>{html.escape(user['full_name'])}</b></div>
        <div>{html.escape(user['email'])} · {html.escape(user['staff_id'])}</div>
        <div style="margin-top:6px">{role_badge(role)}
          {('&nbsp;' + dept_badge(user_dept)) if user_dept else ''}</div>
      </div>
    </div>""", unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(f"""
        <div class="sb-brand">{logo_html(44)}
          <div><b>National Petroleum Authority</b><span>Material Balance Portal</span></div>
        </div>""", unsafe_allow_html=True)
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,{NAVY2},{NAVY});color:#fff;
                    padding:14px 16px;border-radius:10px;margin-bottom:14px">
          <div style="font-size:11px;opacity:.7;text-transform:uppercase;
                      letter-spacing:.6px">Signed in as</div>
          <div style="font-weight:800;font-size:14px;margin-top:3px">
            {html.escape(user['full_name'])}</div>
          <div style="font-size:11.5px;opacity:.85">{html.escape(user['email'])}</div>
          <div style="margin-top:8px">{role_badge(role)}</div>
        </div>""", unsafe_allow_html=True)

        if role == "manager":
            nav = ["Executive (NPA-wide)", "Departmental View",
                   "Dashboard", "Trends", "Material Balance", "Exceptions",
                   "Executive Report", "Demarcation Matrix",
                   "Submissions Log", "User Management", "Company Management",
                   "Audit Log", "Data Audit", "My Profile", "Help & Template"]
        elif role == "company":
            nav = ["Executive (NPA-wide)", "Departmental View",
                   "Dashboard", "Trends", "Material Balance", "Exceptions",
                   "Executive Report", "Demarcation Matrix",
                   "My Depots", "My Staff", "Submissions Log", "Data Audit",
                   "My Profile", "Help & Template"]
        else:
            nav = ["Departmental View", "Dashboard", "Trends", "Material Balance",
                   "Exceptions", "Executive Report",
                   "My Submissions", "Data Audit", "My Profile", "Help & Template"]

        page = st.radio("Navigation", nav, key=f"nav_{user['id']}",
                        label_visibility="collapsed")

        st.markdown("---")
        st.markdown("### 🎨 Branding")
        logo_file = st.file_uploader("NPA logo (PNG / SVG / WEBP / JPG)",
                                     type=list(LOGO_EXTS), key="logo_uploader",
                                     help="Transparent PNG or SVG works best.")
        if logo_file is not None:
            ext = logo_file.name.lower().rsplit(".", 1)[-1]
            st.session_state["logo_url_override"] = (
                f"data:{_MIME.get(ext, 'image/png')};base64,"
                + base64.b64encode(logo_file.getvalue()).decode())
            if role == "manager" and st.button("💾 Save as official logo", key="logo_save",
                                               width="stretch"):
                try:
                    save_logo_permanently(logo_file.getvalue(), logo_file.name)
                    st.session_state.pop("logo_url_override", None)
                    log_action(user["email"], "logo_updated", logo_file.name)
                    st.success("Official logo saved to ./assets.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        if st.session_state.get("logo_url_override") and st.button(
                "Reset logo", key="logo_reset", width="stretch"):
            st.session_state.pop("logo_url_override", None)
            st.rerun()
        if get_logo_url() is None:
            st.caption("No logo found — add `assets/npa_logo.png` or upload one above.")
        bg_file = st.file_uploader("Upload custom background",
                                   type=["png", "jpg", "jpeg", "webp"],
                                   key="bg_uploader",
                                   help="Overrides the default for this session only.")
        if bg_file is not None:
            data = base64.b64encode(bg_file.getvalue()).decode()
            ext = bg_file.name.lower().rsplit(".", 1)[-1]
            mime = {"png": "image/png", "jpg": "image/jpeg",
                    "jpeg": "image/jpeg", "webp": "image/webp"}.get(ext, "image/png")
            st.session_state["bg_url_override"] = f"data:{mime};base64,{data}"
        if st.session_state.get("bg_url_override"):
            if st.button("Reset background", key="bg_reset", width="stretch"):
                st.session_state.pop("bg_url_override", None)
                st.rerun()

        st.markdown("---")
        if st.button("🚪  Sign Out", width="stretch", key=f"so_{user['id']}"):
            log_action(user["email"], "logout")
            st.session_state.user = None
            st.rerun()

    raw = None
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📂 Data Source")
        uploaded = st.file_uploader("Upload material files (any layout)",
                                    type=["xlsx", "xlsm", "xls", "csv", "tsv", "txt"],
                                    accept_multiple_files=True,
                                    key=f"up_{user['id']}",
                                    help="Messy files are fine — titles, merged cells, totals, "
                                         "mixed units and text-formatted numbers are cleaned "
                                         "automatically. Report in the Data Audit page.")
        auto_header = st.checkbox("Auto-detect header row", value=True,
                                  key=f"ah_{user['id']}")
        header_row = -1
        if not auto_header:
            header_row = int(st.number_input("Header row (1 = first row)", 1, 200, 1, 1,
                                             key=f"hr_{user['id']}")) - 1
        with st.expander("🧹 Cleaning options", expanded=False):
            unit_override = st.selectbox(
                "Volume unit in the files", ["Auto", "L", "KL", "M3", "GAL", "BBL"],
                key=f"uo_{user['id']}",
                help="Auto uses a Units column or a unit in the header, e.g. 'Receipts (BBL)'. "
                     "Pick one to force it for every file.")
            dayfirst = st.checkbox("Day-first dates (31/01/2024)", True, key=f"df_{user['id']}")
            drop_totals = st.checkbox("Remove total / subtotal rows", True, key=f"dt_{user['id']}")
            drop_dupes = st.checkbox("Remove exact duplicate records", True, key=f"dd_{user['id']}")
            fill_merged = st.checkbox("Fill blank cells from merged / grouped cells", True,
                                      key=f"fm_{user['id']}")
            std_names = st.checkbox("Standardise depot & product names", True,
                                    key=f"sn_{user['id']}")
        sheet_sel: Tuple[str, ...] = ()
        if uploaded:
            seen: List[str] = []
            for f in uploaded:
                for s in list_sheets(f.getvalue(), f.name):
                    if s not in seen:
                        seen.append(s)
            if seen:
                sheet_sel = tuple(st.multiselect("Sheets", seen, default=seen,
                                                 key=f"sh_{user['id']}"))
        st.markdown("---")
        st.markdown("### ⚖️ Tolerance")
        tol_pct = st.slider("Default variance tolerance (%)", 0.0, 5.0, 0.50, 0.05,
                            key=f"tp_{user['id']}",
                            help="Overridden by per-department tolerances when known.")
        tol_abs = st.number_input("Absolute tolerance (L)", 0.0, value=0.0, step=100.0,
                                  key=f"ta_{user['id']}")

    if uploaded:
        frames, structure_log = [], []
        for f in uploaded:
            try:
                fr, flog = read_upload(f.getvalue(), f.name, tuple(sheet_sel),
                                       int(header_row), bool(drop_totals))
                frames.append(fr)
                structure_log.extend(flog)
            except Exception as exc:
                st.error(f"⚠️ {f.name}: {exc}")
        if frames:
            raw = pd.concat(frames, ignore_index=True, sort=False)

    NO_DATA_PAGES = {"User Management", "Company Management", "Audit Log", "My Staff",
                     "Submissions Log", "My Profile", "Help & Template"}
    no_data = raw is None or raw.empty
    if page in NO_DATA_PAGES or (page == "My Submissions" and no_data):
        render_no_data_page(page, role, user)      # admin / account pages need no upload
        render_footer()
        return

    if no_data:
        st.markdown('<div class="note">📂 Upload one or more depot material files '
                    'in the sidebar to begin — or start from the template.</div>',
                    unsafe_allow_html=True)
        c1, _ = st.columns([1, 2])
        with c1:
            template_download_button("empty_tpl")
        render_footer()
        return

    raw.columns = [str(c).strip() for c in raw.columns]
    data_cols = [c for c in raw.columns if not c.startswith("__")]
    detected = detect_columns(data_cols)
    available = ["— none —"] + data_cols

    with st.expander("🔧  Column Mapping", expanded=False):
        mapping: Dict[str, Optional[str]] = {}
        cols = st.columns(3)
        for i, f in enumerate(PRIORITY):
            default = detected.get(f)
            idx = available.index(default) if default in available else 0
            with cols[i % 3]:
                choice = st.selectbox(FIELD_LABELS[f], available, index=idx,
                                      key=f"map_{user['id']}_{f}")
                mapping[f] = None if choice == "— none —" else choice

    clean_opts = {"unit_override": unit_override, "dayfirst": bool(dayfirst),
                  "drop_dupes": bool(drop_dupes), "fill_merged": bool(fill_merged),
                  "std_names": bool(std_names), "structure_log": structure_log}
    try:
        balance, issues, meta, clean_log = build_balance(raw, mapping, tol_pct, tol_abs,
                                                         clean_opts)
    except Exception as exc:
        st.error("⚠️ The files could not be processed with the current column mapping "
                 f"({type(exc).__name__}: {exc}). Check the Column Mapping panel and the "
                 "header-row setting.")
        render_footer()
        return
    if balance.empty:
        st.error("No usable records were found after cleaning. Check the column mapping, "
                 "the header row and the selected sheets.")
        st.dataframe(clean_log, width="stretch", hide_index=True)
        render_footer()
        return
    _log_new_uploads(user, balance, uploaded)
    if user.get("depot") and role == "staff":
        only = balance[balance["Depot"].str.lower() == user["depot"].strip().lower()]
        if not only.empty:
            balance = only
    unit = meta.get("unit", "L")
    n_clean = int(clean_log["Affected"].sum()) if not clean_log.empty else 0
    st.caption(f"🧹 Auto-cleaned: {meta['rows_in']} rows in → {meta['rows_analysed']} analysed "
               f"· {n_clean:,} fixes applied · all volumes in litres "
               f"(source units: {meta['source_units']}) · details in **Data Audit**")

    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🎯 Filters")
        depts_present = [d for d in DEPARTMENT_ORDER
                         if d in balance["Department"].unique().tolist()]
        dept_sel = st.multiselect(
            "Departments", depts_present, default=depts_present,
            format_func=lambda d: f'{DEPARTMENTS[d]["icon"]} {DEPARTMENTS[d]["label"]}',
            key=f"fdp_{user['id']}")
        depots = sorted(balance["Depot"].dropna().unique().tolist())
        depot_sel = st.multiselect("Depots / Nodes", depots, default=depots,
                                   key=f"fd_{user['id']}")
        products = sorted(balance["Product"].dropna().unique().tolist())
        product_sel = st.multiselect("Products", products, default=products,
                                     key=f"fp_{user['id']}")
        status_sel = st.multiselect("Status",
                                    ["Within Tolerance", "Gain (Over)", "Loss (Short)", "Unverified"],
                                    default=["Within Tolerance", "Gain (Over)", "Loss (Short)", "Unverified"],
                                    key=f"fs_{user['id']}")

    mask = (balance["Department"].isin(dept_sel)
            & balance["Depot"].isin(depot_sel)
            & balance["Product"].isin(product_sel)
            & balance["Status"].isin(status_sel))
    fdf = balance.loc[mask].copy()
    if fdf.empty:
        st.warning("No rows match the current filters.")
        return

    ledger = fdf[["Date", "Department", "Depot", "Node Type", "Product", "Reference",
                  "Opening Stock", "Receipts", "Transfers In", "Available",
                  "Sales / Deliveries", "Transfers Out", "Expected Closing",
                  "Actual Closing", "Gain / (Loss)", "Variance %", "Tolerance %",
                  "Status", "Remark"]].sort_values(
        ["Date", "Department", "Depot", "Product"], na_position="last")

    if page == "Executive (NPA-wide)":
        render_executive_npa(fdf, tol_pct, unit, user)
    elif page == "Departmental View":
        render_departmental_view(fdf, tol_pct, unit, user)
    elif page == "Demarcation Matrix":
        render_demarcation_matrix(fdf, unit)
    elif page == "Dashboard":
        render_dashboard(fdf, tol_pct, unit)
    elif page == "Trends":
        render_trends(fdf, tol_pct, unit)
    elif page == "Material Balance":
        render_balance(fdf, ledger)
    elif page == "Exceptions":
        render_exceptions(fdf, issues)
    elif page == "Executive Report":
        render_report(fdf, issues, meta, tol_pct, tol_abs, ledger)
    elif page == "My Depots" and role == "company":
        render_my_depots(fdf)
    elif page == "My Submissions" and role == "staff":
        render_my_submissions(fdf, user)
    elif page == "Data Audit":
        render_data_audit(raw, mapping, issues, unit, clean_log, balance)
    else:
        render_dashboard(fdf, tol_pct, unit)
    render_footer()


# ==========================================================================
# EXECUTIVE NPA-WIDE
# ==========================================================================
def render_executive_npa(fdf: pd.DataFrame, tol_pct: float, unit: str, user: dict) -> None:
    st.markdown('<div class="section-title">Executive Command · NPA-Wide Consolidation</div>',
                unsafe_allow_html=True)
    st.caption(f"Consolidated across **{fdf['Department'].nunique()}** departments · "
               f"**{fdf['Depot'].nunique()}** nodes · **{len(fdf):,}** balance lines · "
               f"unit: **{unit}**")

    agg = fdf.groupby("Department", as_index=False).agg(**{
        "Records":     ("Department", "size"),
        "Nodes":       ("Depot", "nunique"),
        "Throughput":  ("Throughput", "sum"),
        "Receipts":    ("Receipts", "sum"),
        "Issues":      ("Sales / Deliveries", "sum"),
        "Closing":     ("Actual Closing", "sum"),
        "Variance":    ("Gain / (Loss)", "sum"),
    })
    agg["Var %"] = np.where(agg["Throughput"] > 0,
                            agg["Variance"] / agg["Throughput"] * 100, np.nan)
    agg["Tolerance"] = agg["Department"].map(
        lambda d: DEPARTMENTS.get(d, {}).get("tol", tol_pct))
    agg["Breach"] = agg.apply(
        lambda r: pd.notna(r["Var %"]) and abs(r["Var %"]) > r["Tolerance"], axis=1)
    agg = agg.sort_values("Variance", key=lambda s: s.abs(), ascending=False)

    total_tp  = fdf["Throughput"].sum()
    total_var = fdf["Gain / (Loss)"].sum()
    sys_pct   = total_var / total_tp * 100 if total_tp else np.nan
    breaches  = int(agg["Breach"].sum())
    losses    = fdf.loc[fdf["Gain / (Loss)"] < 0, "Gain / (Loss)"].sum()
    gains     = fdf.loc[fdf["Gain / (Loss)"] > 0, "Gain / (Loss)"].sum()
    compliance = compliance_pct(fdf)

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(kpi("National Throughput", fmt(total_tp),
                    f"{len(fdf):,} records", "gold"), unsafe_allow_html=True)
    k2.markdown(kpi("System Variance", fmt(total_var), fmt_pct(sys_pct),
                    "red" if pd.notna(sys_pct) and abs(sys_pct) > tol_pct else "green"),
                unsafe_allow_html=True)
    k3.markdown(kpi("Total Losses", fmt(losses), "shortages", "red"),
                unsafe_allow_html=True)
    k4.markdown(kpi("Total Gains", fmt(gains), "overages", "green"),
                unsafe_allow_html=True)
    k5.markdown(kpi("Departments", f"{len(agg)}",
                    f"{breaches} in breach", "amber" if breaches else "green"),
                unsafe_allow_html=True)
    k6.markdown(kpi("Compliance", f"{fmt_comp(compliance)}",
                    f"{int((fdf["Status"].isin(EXC_STATUSES)).sum())} exception(s)",
                    "green" if compliance >= 95 else ("amber" if compliance >= 80 else "red")),
                unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="section-title">Departmental League Table</div>',
                unsafe_allow_html=True)
    league_rows = []
    for _, r in agg.iterrows():
        m = dept_meta(r["Department"])
        status_txt = ("Loss Breach" if r["Variance"] < 0 else "Gain / Overage") \
            if r["Breach"] else "Within Tolerance"
        tone = "bad" if (r["Breach"] and r["Variance"] < 0) else \
               ("warn" if r["Breach"] else "ok")
        league_rows.append({
            "Department": f'{m["icon"]} {m["label"]}',
            "Records":    int(r["Records"]),
            "Nodes":      int(r["Nodes"]),
            "Throughput": fmt(r["Throughput"]),
            "Variance":   fmt(r["Variance"]),
            "Var %":      fmt_pct(r["Var %"]),
            "Tolerance":  f'±{r["Tolerance"]}%',
            "Status":     f'<span class="pill {tone}">{status_txt}</span>',
        })
    st.markdown(
        '<table style="width:100%;border-collapse:collapse;font-size:12.5px;background:#fff;'
        'border-radius:10px;overflow:hidden">'
        '<tr style="background:' + NAVY + ';color:#fff">'
        '<th style="text-align:left;padding:9px 12px">Department</th>'
        '<th style="text-align:right;padding:9px 12px">Records</th>'
        '<th style="text-align:right;padding:9px 12px">Nodes</th>'
        '<th style="text-align:right;padding:9px 12px">Throughput</th>'
        '<th style="text-align:right;padding:9px 12px">Variance</th>'
        '<th style="text-align:right;padding:9px 12px">Var %</th>'
        '<th style="text-align:right;padding:9px 12px">Tolerance</th>'
        '<th style="text-align:left;padding:9px 12px">Status</th></tr>'
        + "".join(
            f'<tr style="border-bottom:1px solid {LINE}">'
            f'<td style="padding:8px 12px;font-weight:600">{row["Department"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Records"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Nodes"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Throughput"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Variance"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Var %"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{row["Tolerance"]}</td>'
            f'<td style="padding:8px 12px">{row["Status"]}</td></tr>'
            for row in league_rows)
        + '</table>', unsafe_allow_html=True)

    st.markdown("")
    c1, c2 = st.columns([1.4, 1])
    with c1:
        fig = px.bar(agg.sort_values("Variance"),
                     x="Variance", y="Department", orientation="h",
                     color="Breach",
                     color_discrete_map={True: RED, False: BLUE, "True": RED, "False": BLUE},
                     template=PLOTLY_TEMPLATE,
                     title=f"Variance by Department ({unit})")
        fig.update_yaxes(tickmode="array",
                         tickvals=agg["Department"],
                         ticktext=[f'{DEPARTMENTS[d]["icon"]} {DEPARTMENTS[d]["short"]}'
                                   for d in agg["Department"]])
        fig.update_layout(height=420, showlegend=False,
                          margin=dict(l=0, r=10, t=50, b=0))
        st.plotly_chart(fig, width="stretch")

    with c2:
        fig = px.pie(agg, names="Department", values="Throughput", hole=0.55,
                     template=PLOTLY_TEMPLATE,
                     color="Department",
                     color_discrete_map={d: DEPARTMENTS[d]["color"]
                                         for d in agg["Department"]},
                     title="Throughput Share by Department")
        fig.update_traces(textposition="inside", textinfo="percent")
        fig.update_layout(height=420, showlegend=True,
                          margin=dict(l=0, r=0, t=50, b=0))
        st.plotly_chart(fig, width="stretch")

    st.markdown('<div class="section-title">Monthly Variance Trend by Department</div>',
                unsafe_allow_html=True)
    trend = monthly_dept_trend(fdf)
    if not trend.empty:
        fig = px.line(trend, x="Month", y="Var %", color="Department",
                      markers=True, template=PLOTLY_TEMPLATE,
                      color_discrete_map={d: DEPARTMENTS[d]["color"]
                                          for d in trend["Department"].unique()})
        fig.update_layout(height=380, margin=dict(l=0, r=10, t=20, b=0),
                          legend_title="")
        st.plotly_chart(fig, width="stretch")

    st.markdown('<div class="section-title">Product Performance Summary</div>',
                unsafe_allow_html=True)
    bp = fdf.groupby("Product", as_index=False).agg(**{
        "Throughput": ("Throughput", "sum"),
        "Variance":   ("Gain / (Loss)", "sum"),
        "Records":    ("Product", "size"),
    })
    bp["Var %"] = np.where(bp["Throughput"] > 0,
                           bp["Variance"] / bp["Throughput"] * 100, np.nan)
    st.dataframe(bp, width="stretch", hide_index=True)

    st.markdown(f"""
    <div class="note warn" style="margin-top:16px">
      <b>Executive note:</b> National system loss of <b>{fmt_pct(sys_pct)}</b>
      ({fmt(total_var)} {unit}) across {len(agg)} departments.
      <b>{breaches}</b> department(s) exceeded their demarcated tolerance.
      Losses of <b>{fmt(abs(losses))} {unit}</b> require departmental reconciliation
      before the next regulatory cycle.
    </div>""", unsafe_allow_html=True)

    # ---------- EXPORTS ----------
    st.markdown('<div class="section-title">Export Executive Pack</div>',
                unsafe_allow_html=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    sheets = {"Department League": agg}
    for dept, grp in fdf.groupby("Department"):
        sub = grp.groupby("Depot", as_index=False).agg(**{
            "Throughput": ("Throughput", "sum"),
            "Variance":   ("Gain / (Loss)", "sum"),
            "Records":    ("Depot", "size"),
        })
        sub["Var %"] = np.where(sub["Throughput"] > 0,
                                sub["Variance"] / sub["Throughput"] * 100, np.nan)
        sheets[DEPARTMENTS.get(dept, {}).get("short", dept)] = sub
    sheets["Exception Register"] = fdf[fdf["Status"].isin(EXC_STATUSES)][[
        "Date", "Department", "Depot", "Product", "Gain / (Loss)",
        "Variance %", "Tolerance %", "Status", "Remark"]]

    wb = to_excel_bytes(sheets, meta={
        "Report": "NPA Executive Material Balance Pack",
        "Generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Records": f"{len(fdf):,}", "Unit": unit,
        "Compliance %": f"{compliance:.2f}",
        "Departments": f"{len(agg)}",
        "Breaches": f"{breaches}"})

    # ---- PDF ----
    kpi_rows = [
        ("National Throughput", f"{fmt(total_tp)} {unit}"),
        ("System Variance",     f"{fmt(total_var)} {unit} ({fmt_pct(sys_pct)})"),
        ("Total Losses",        f"{fmt(losses)} {unit}"),
        ("Total Gains",         f"{fmt(gains)} {unit}"),
        ("Departments in scope", f"{len(agg)}"),
        ("Tolerance breaches",  f"{breaches}"),
        ("Compliance",          f"{fmt_comp(compliance)}"),
    ]
    summary_rows = [
        ["Opening + Receipts (Available)", fmt(fdf["Available"].sum())],
        ["Sales / Deliveries (Issues)",    fmt(fdf["Sales / Deliveries"].sum())],
        ["Expected Closing",               fmt(fdf["Expected Closing"].sum())],
        ["Actual Closing",                 fmt(fdf["Actual Closing"].sum())],
        ["Net Gain / (Loss)",              fmt(total_var)],
        ["Variance % of throughput",       fmt_pct(sys_pct)],
    ]
    dept_headers = ["Department", "Records", "Nodes", "Throughput", "Var %", "Status"]
    dept_rows = []
    for _, r in agg.iterrows():
        m = dept_meta(r["Department"])
        status_txt = ("Loss Breach" if r["Variance"] < 0 else "Gain/Overage") \
            if r["Breach"] else "Within"
        dept_rows.append([m["label"], str(int(r["Records"])), str(int(r["Nodes"])),
                          fmt(r["Throughput"]), fmt_pct(r["Var %"]), status_txt])
    exc_headers = ["Date", "Department", "Depot", "Product", "Var %", "Status"]
    exc_rows = []
    top_exc = fdf[fdf["Status"].isin(EXC_STATUSES)].head(15)
    for _, r in top_exc.iterrows():
        exc_rows.append([
            str(r["Date"])[:10] if pd.notna(r["Date"]) else "—",
            dept_meta(r["Department"])["short"],
            str(r["Depot"])[:22],
            str(r["Product"]),
            fmt_pct(r["Variance %"]),
            r["Status"],
        ])
    note_txt = (f"National system loss of {fmt_pct(sys_pct)} ({fmt(total_var)} {unit}) "
                f"across {len(agg)} departments. {breaches} department(s) exceeded "
                f"their demarcated tolerance.")

    try:
        _ex = cached_pdf_extras(fdf, {}, tol_pct, unit, "AUTO", None)
        _extras = {k: v for k, v in _ex.items() if k not in ("_analysis", "extra_tables")}
        pdf_bytes = build_pdf_report(
            title="Executive Material Balance Pack",
            subtitle="NPA-Wide Consolidation",
            meta_lines=[
                ("Generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
                ("Records", f"{len(fdf):,}"),
                ("Departments", f"{len(agg)}"),
                ("Nodes", f"{fdf['Depot'].nunique()}"),
                ("Unit", unit),
            ],
            kpi_rows=kpi_rows,
            summary_table=summary_rows,
            department_table=[dept_headers] + dept_rows,
            exceptions=[exc_headers] + exc_rows if exc_rows else None,
            note=note_txt, **_extras,
        )
        pdf_ok = True
        pdf_err = ""
    except Exception as e:
        pdf_bytes = b""
        pdf_ok = False
        pdf_err = str(e)

    d1, d2, d3 = st.columns(3)
    d1.download_button("⬇️  Download executive workbook",
                       wb, file_name=f"NPA_Executive_{stamp}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", type="primary", key="exec_wb")
    if pdf_ok:
        d2.download_button("📄  Download PDF report", pdf_bytes,
                           file_name=f"NPA_Executive_{stamp}.pdf",
                           mime="application/pdf",
                           width="stretch", key="exec_pdf")
    else:
        d2.button("📄  PDF unavailable", disabled=True,
                  width="stretch", key="exec_pdf_dis")
        if not _HAS_FPDF:
            d2.caption("Install `fpdf2`: `pip install fpdf2`")
        elif pdf_err:
            d2.caption(f"PDF error: {pdf_err[:120]}")
    d3.download_button("⬇️  Exception register (CSV)",
                       to_csv_bytes(sheets["Exception Register"]),
                       file_name=f"NPA_Exceptions_{stamp}.csv",
                       mime="text/csv", width="stretch", key="exec_exc")


def monthly_dept_trend(fdf: pd.DataFrame) -> pd.DataFrame:
    if fdf["Date"].isna().all():
        return pd.DataFrame()
    tmp = fdf.dropna(subset=["Date"]).copy()
    tmp["Month"] = tmp["Date"].dt.to_period("M").astype(str)
    grp = tmp.groupby(["Month", "Department"], as_index=False).agg(**{
        "Throughput": ("Throughput", "sum"),
        "Variance":   ("Gain / (Loss)", "sum"),
    })
    grp["Var %"] = np.where(grp["Throughput"] > 0,
                            grp["Variance"] / grp["Throughput"] * 100, np.nan)
    return grp


# ==========================================================================
# DEPARTMENTAL VIEW
# ==========================================================================
def render_departmental_view(fdf: pd.DataFrame, tol_pct: float, unit: str,
                             user: dict) -> None:
    st.markdown('<div class="section-title">Departmental Demarcation View</div>',
                unsafe_allow_html=True)
    st.caption("Each tab shows the KPI set, tolerance and scope of one NPA department. "
               "Material-balance departments consolidate their own nodes; cross-cutting "
               "departments consume the same data for their accountability lens.")

    present = [d for d in DEPARTMENT_ORDER if d in fdf["Department"].unique().tolist()]
    if not present:
        st.info("No department column detected — check mapping under Data Audit.")
        return

    tab_labels = [f'{DEPARTMENTS[d]["icon"]} {DEPARTMENTS[d]["short"]}' for d in present]
    tabs = st.tabs(tab_labels)

    for tab, dept in zip(tabs, present):
        with tab:
            sub = fdf[fdf["Department"] == dept].copy()
            if sub.empty:
                st.info(f"No records under {dept}.")
                continue
            m = dept_meta(dept)
            st.markdown(dept_banner(dept,
                f"· {len(sub):,} records · {sub['Depot'].nunique()} nodes · unit {unit}"),
                unsafe_allow_html=True)

            tp    = sub["Throughput"].sum()
            var   = sub["Gain / (Loss)"].sum()
            vp    = var / tp * 100 if tp else np.nan
            recv  = sub["Receipts"].sum() + sub["Transfers In"].sum()
            iss   = sub["Sales / Deliveries"].sum() + sub["Transfers Out"].sum()
            close = sub["Actual Closing"].sum()
            breach_count = int((sub["Status"].isin(EXC_STATUSES)).sum())
            breach = pd.notna(vp) and abs(vp) > m["tol"]

            if m["kind"] == "material":
                k1, k2, k3, k4 = st.columns(4)
                k1.markdown(kpi("Throughput", fmt(tp), f"{len(sub):,} records", ""),
                            unsafe_allow_html=True)
                k2.markdown(kpi("Receipts", fmt(recv), "into dept", "green"),
                            unsafe_allow_html=True)
                k3.markdown(kpi("Issues / Deliveries", fmt(iss), "out of dept"),
                            unsafe_allow_html=True)
                k4.markdown(kpi("Variance", fmt(var), fmt_pct(vp),
                                "red" if breach else "green"),
                            unsafe_allow_html=True)
                k5, k6, k7, k8 = st.columns(4)
                k5.markdown(kpi("Closing Stock", fmt(close), "actual physical", "gold"),
                            unsafe_allow_html=True)
                k6.markdown(kpi("Tolerance", f"±{m['tol']}%",
                                "breach" if breach else "within",
                                "red" if breach else "green"),
                            unsafe_allow_html=True)
                k7.markdown(kpi("Exceptions", f"{breach_count}",
                                f"of {len(sub):,}",
                                "red" if breach_count else "green"),
                            unsafe_allow_html=True)
                k8.markdown(kpi("Nodes", f"{sub['Depot'].nunique()}",
                                f"{sub['Product'].nunique()} products"),
                            unsafe_allow_html=True)
            else:
                k1, k2, k3, k4 = st.columns(4)
                k1.markdown(kpi("Records", f"{len(sub):,}", "in scope", "purple"),
                            unsafe_allow_html=True)
                k2.markdown(kpi("Nodes", f"{sub['Depot'].nunique()}", "touched"),
                            unsafe_allow_html=True)
                k3.markdown(kpi("Throughput", fmt(tp), "to review"),
                            unsafe_allow_html=True)
                k4.markdown(kpi("Exceptions", f"{breach_count}", "flagged",
                                "red" if breach_count else "green"),
                            unsafe_allow_html=True)

            st.markdown("")
            c1, c2 = st.columns([1.3, 1])
            with c1:
                by_node = sub.groupby("Depot", as_index=False).agg(**{
                    "Throughput": ("Throughput", "sum"),
                    "Variance":   ("Gain / (Loss)", "sum"),
                    "Records":    ("Depot", "size"),
                })
                by_node["Var %"] = np.where(by_node["Throughput"] > 0,
                                            by_node["Variance"] /
                                            by_node["Throughput"] * 100, np.nan)
                fig = px.bar(by_node.sort_values("Var %"),
                             x="Var %", y="Depot", orientation="h",
                             color="Var %", color_continuous_scale="RdYlGn_r",
                             template=PLOTLY_TEMPLATE,
                             title=f"{m['short']} — Variance % by Node")
                fig.update_layout(height=max(340, 40 * len(by_node)),
                                  showlegend=False,
                                  margin=dict(l=0, r=10, t=50, b=0))
                st.plotly_chart(fig, width="stretch")
            with c2:
                bp = sub.groupby("Product", as_index=False)["Sales / Deliveries"].sum()
                bp = bp[bp["Sales / Deliveries"] > 0]
                if not bp.empty:
                    fig = px.pie(bp, names="Product", values="Sales / Deliveries",
                                 hole=0.55, template=PLOTLY_TEMPLATE,
                                 color_discrete_sequence=PRODUCT_COLORS,
                                 title=f"{m['short']} — Product Mix")
                    fig.update_traces(textposition="inside", textinfo="percent+label")
                    fig.update_layout(height=max(340, 40 * len(by_node)),
                                      showlegend=False,
                                      margin=dict(l=0, r=0, t=50, b=0))
                    st.plotly_chart(fig, width="stretch")

            st.markdown(f"**{m['short']} — Node-Product Detail**")
            st.dataframe(
                sub[["Date", "Depot", "Node Type", "Product", "Opening Stock",
                     "Receipts", "Sales / Deliveries", "Actual Closing",
                     "Gain / (Loss)", "Variance %", "Tolerance %", "Status", "Remark"]]
                .sort_values("Variance %", key=lambda s: s.abs(), ascending=False),
                width="stretch", height=380)

            render_dept_extras(dept, sub, unit)

            st.markdown("")
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
            wb = to_excel_bytes({
                "Detail": sub,
                "By Node": by_node,
                "By Product": bp,
            }, meta={"Department": m["label"], "Tolerance %": f'{m["tol"]}',
                     "Generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     "Unit": unit})
            st.download_button(
                f"⬇️  Download {m['short']} workbook",
                wb, file_name=f"NPA_{dept}_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{dept}_{user['id']}")


def render_dept_extras(dept: str, sub: pd.DataFrame, unit: str) -> None:
    m = dept_meta(dept)
    if dept == "IMPORT":
        st.markdown(f"**{m['short']} — Cargo Out-turn & Demurrage**")
        st.caption("Vessel discharge reconciliation: BL quantity vs shore receipt.")
        cargo = sub.groupby("Depot", as_index=False).agg(**{
            "Receipts":  ("Receipts", "sum"),
            "Opening":   ("Opening Stock", "sum"),
        })
        cargo["Implied Discharge"] = cargo["Receipts"] + cargo["Opening"]
        st.dataframe(cargo, width="stretch", hide_index=True)
    elif dept == "RETAIL":
        st.markdown(f"**{m['short']} — Outlet Allocation vs Sales**")
        omc = sub.groupby("Depot", as_index=False).agg(**{
            "Allocated": ("Receipts", "sum"),
            "Sales":     ("Sales / Deliveries", "sum"),
            "Variance":  ("Gain / (Loss)", "sum"),
        })
        omc["Gap"] = omc["Allocated"] - omc["Sales"]
        omc["Gap %"] = np.where(omc["Allocated"] > 0,
                                omc["Gap"] / omc["Allocated"] * 100, np.nan)
        st.dataframe(omc, width="stretch", hide_index=True)
    elif dept == "TRANSPORT":
        st.markdown(f"**{m['short']} — Transit Loss Register**")
        st.caption("Volume lost between load and discharge points.")
        transit = sub[sub["Gain / (Loss)"] < 0][[
            "Date", "Depot", "Product", "Receipts", "Sales / Deliveries",
            "Gain / (Loss)", "Variance %"]]
        if transit.empty:
            st.success("No transit losses recorded.")
        else:
            st.dataframe(transit, width="stretch", hide_index=True)
    elif dept == "DEPOT":
        st.markdown(f"**{m['short']} — Dip Compliance & Stock Cover**")
        dip = sub.groupby("Depot", as_index=False).agg(**{
            "Records":   ("Depot", "size"),
            "Closing":   ("Actual Closing", "sum"),
            "Issues":    ("Sales / Deliveries", "sum"),
        })
        dip["Cover Days (approx)"] = np.where(
            dip["Issues"] > 0, dip["Closing"] / (dip["Issues"] / 30), np.nan)
        st.dataframe(dip, width="stretch", hide_index=True)
    elif dept == "REFINERY":
        st.markdown(f"**{m['short']} — Feedstock vs Product Slate**")
        slate = sub.groupby("Product", as_index=False).agg(**{
            "Output":   ("Sales / Deliveries", "sum"),
            "Variance": ("Gain / (Loss)", "sum"),
        })
        total_out = slate["Output"].sum() or 1
        slate["Yield %"] = slate["Output"] / total_out * 100
        st.dataframe(slate, width="stretch", hide_index=True)
    elif dept == "MARINE":
        st.markdown(f"**{m['short']} — Jetty Throughput by Product**")
        jet = sub.groupby("Product", as_index=False).agg(**{
            "Throughput": ("Throughput", "sum"),
            "Variance":   ("Gain / (Loss)", "sum"),
        })
        st.dataframe(jet, width="stretch", hide_index=True)
    elif dept == "QUALITY":
        st.markdown(f"**{m['short']} — Off-spec Watch List**")
        st.caption("Rows with variance outside tolerance may indicate density/"
                   "temperature correction errors — refer to lab.")
        watch = sub[sub["Status"].isin(EXC_STATUSES)][[
            "Date", "Depot", "Product", "Gain / (Loss)", "Variance %", "Remark"]]
        st.dataframe(watch, width="stretch", hide_index=True)
    elif dept == "FINANCE":
        st.markdown(f"**{m['short']} — Revenue at Risk by Node**")
        risk = sub.groupby("Depot", as_index=False)["Gain / (Loss)"].sum()
        risk["Loss Volume"] = risk["Gain / (Loss)"].apply(lambda v: min(v, 0))
        risk["Est. Levy @ 0.28/L"] = risk["Loss Volume"].abs() * 0.28
        st.dataframe(risk, width="stretch", hide_index=True)
    elif dept == "COMPLIANCE":
        st.markdown(f"**{m['short']} — Escalation Register**")
        esc = sub[sub["Status"] == "Loss (Short)"][[
            "Date", "Depot", "Product", "Gain / (Loss)", "Variance %",
            "Tolerance %", "Remark"]]
        st.dataframe(esc, width="stretch", hide_index=True)
    elif dept == "HSE":
        st.markdown(f"**{m['short']} — Product Loss to Environment (proxy)**")
        st.caption("Large negative variances may indicate spill events. Correlate "
                   "with incident reports.")
        hse = sub[sub["Variance %"] < -1.0][[
            "Date", "Depot", "Product", "Gain / (Loss)", "Variance %"]]
        st.dataframe(hse, width="stretch", hide_index=True)
    elif dept == "IT":
        st.markdown(f"**{m['short']} — Data Quality Flags**")
        flags = pd.DataFrame({
            "Check": ["Nodes reporting", "Rows with dates", "Rows missing product",
                      "Rows missing depot"],
            "Value": [sub["Depot"].nunique(),
                      int(sub["Date"].notna().sum()),
                      int((sub["Product"].isin(["", "UNSPECIFIED", "NAN"])).sum()),
                      int(sub["Depot"].isin(["", "Unspecified", "NAN"]).sum())],
        })
        st.dataframe(flags, width="stretch", hide_index=True)


# ==========================================================================
# DEMARCATION MATRIX
# ==========================================================================
def render_demarcation_matrix(fdf: pd.DataFrame, unit: str) -> None:
    st.markdown('<div class="section-title">NPA Departmental Demarcation Matrix</div>',
                unsafe_allow_html=True)
    st.caption("Formal ownership boundaries for material-balance nodes, KPI "
               "accountability and tolerance limits.")

    rows = []
    for code in DEPARTMENT_ORDER:
        m = dept_meta(code)
        sub = fdf[fdf["Department"] == code]
        tp = sub["Throughput"].sum() if not sub.empty else 0
        var = sub["Gain / (Loss)"].sum() if not sub.empty else 0
        vp = var / tp * 100 if tp else None
        breach = vp is not None and abs(vp) > m["tol"]
        rows.append({
            "Department": f'{m["icon"]} {m["label"]}',
            "Short":      m["short"],
            "Scope":      m["scope"],
            "Kind":       "Material balance" if m["kind"] == "material" else "Cross-cutting",
            "KPIs":       " · ".join(m["kpis"]),
            "Nodes":      sub["Depot"].nunique() if not sub.empty else 0,
            "Throughput": fmt(tp) if tp else "—",
            "Var %":      fmt_pct(vp) if vp is not None else "—",
            "Tolerance":  f"±{m['tol']}%",
            "Status":     ("Loss breach" if (breach and var < 0)
                           else "Gain breach" if breach
                           else "Within tolerance" if vp is not None
                           else "No data"),
        })

    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=520)

    st.markdown("""
    <div class="note">
      <b>Demarcation principle:</b> each material-balance node belongs to exactly one
      operations department. Cross-cutting departments (Quality, Finance, Compliance,
      HSE, ICT) consume the same balance data but are accountable for different KPI
      families. When a node's variance exceeds its owning department's tolerance, the
      exception is escalated to that department and mirrored in the Executive view.
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Department KPI Card Wall</div>',
                unsafe_allow_html=True)
    cols = st.columns(3)
    for i, code in enumerate(DEPARTMENT_ORDER):
        m = dept_meta(code)
        with cols[i % 3]:
            st.markdown(f"""
            <div class="kpi" style="border-left-color:{m['color']};margin-bottom:12px">
              <div class="l">{m['icon']} {html.escape(m['short'])}</div>
              <div class="v" style="font-size:14px">{html.escape(m['label'])}</div>
              <div class="d" style="margin-top:6px">{html.escape(m['scope'])}</div>
              <div class="d" style="margin-top:8px"><b>KPIs:</b> {html.escape(' · '.join(m['kpis']))}</div>
              <div class="d" style="margin-top:4px"><b>Tolerance:</b> ±{m['tol']}%</div>
            </div>""", unsafe_allow_html=True)


# ==========================================================================
# DASHBOARD / BALANCE / EXCEPTIONS / REPORT
# ==========================================================================
def render_dashboard(fdf, tol_pct, unit):
    opening = fdf["Opening Stock"].sum()
    receipts = fdf["Receipts"].sum() + fdf["Transfers In"].sum()
    sales = fdf["Sales / Deliveries"].sum() + fdf["Transfers Out"].sum()
    closing = fdf["Actual Closing"].sum()
    variance = fdf["Gain / (Loss)"].sum()
    throughput = fdf["Throughput"].sum()
    var_pct = variance / throughput * 100 if throughput else np.nan
    compliance = compliance_pct(fdf)
    n_exc = int((fdf["Status"].isin(EXC_STATUSES)).sum())

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(kpi("Opening Stock", fmt(opening), f"{len(fdf):,} records"), unsafe_allow_html=True)
    k2.markdown(kpi("Receipts", fmt(receipts), "incl. transfers in", "green"), unsafe_allow_html=True)
    k3.markdown(kpi("Sales / Deliveries", fmt(sales), "incl. transfers out"), unsafe_allow_html=True)
    k4.markdown(kpi("Closing Stock", fmt(closing), "actual physical", "gold"), unsafe_allow_html=True)
    tone = "green" if pd.notna(var_pct) and abs(var_pct) <= tol_pct else "red"
    k5.markdown(kpi("Net Gain / (Loss)", fmt(variance), fmt_pct(var_pct), tone), unsafe_allow_html=True)
    comp_tone = "green" if compliance >= 95 else ("amber" if compliance >= 80 else "red")
    k6.markdown(kpi("Compliance", f"{fmt_comp(compliance)}", f"{n_exc} exception(s)", comp_tone),
                unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="section-title">Depot & Product Analytics</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.15, 1])
    with c1:
        bd = fdf.groupby("Depot", as_index=False).agg(**{
            "Gain / (Loss)": ("Gain / (Loss)", "sum"),
            "Throughput": ("Throughput", "sum")})
        bd["Variance %"] = np.where(bd["Throughput"] > 0,
                                    bd["Gain / (Loss)"] / bd["Throughput"] * 100, np.nan)
        bd["Class"] = np.where(bd["Variance %"].abs() <= tol_pct, "Within Tolerance",
                               np.where(bd["Gain / (Loss)"] > 0, "Gain", "Loss"))
        fig = px.bar(bd.sort_values("Gain / (Loss)"), x="Gain / (Loss)", y="Depot",
                     orientation="h", color="Class",
                     color_discrete_map={"Within Tolerance": BLUE, "Gain": GREEN, "Loss": RED},
                     template=PLOTLY_TEMPLATE,
                     title=f"Net Gain / (Loss) by Depot ({unit})")
        fig.update_layout(height=380, legend_title="", margin=dict(l=0, r=10, t=50, b=0))
        st.plotly_chart(fig, width="stretch")
    with c2:
        bp = fdf.groupby("Product", as_index=False)["Sales / Deliveries"].sum()
        bp = bp[bp["Sales / Deliveries"] > 0]
        if not bp.empty:
            fig = px.pie(bp, names="Product", values="Sales / Deliveries", hole=0.55,
                         template=PLOTLY_TEMPLATE, color_discrete_sequence=PRODUCT_COLORS,
                         title="Sales Mix by Product")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(height=380, showlegend=False, margin=dict(l=0, r=0, t=50, b=0))
            st.plotly_chart(fig, width="stretch")

    st.markdown('<div class="section-title">Overall Balance Flow</div>', unsafe_allow_html=True)
    wf = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "total", "relative", "total"],
        x=["Opening", "Receipts", "Sales", "Expected Closing", "Gain / (Loss)", "Actual Closing"],
        y=[opening, receipts, -sales, 0, variance, 0],
        text=[fmt(opening), fmt(receipts), fmt(-sales), "", fmt(variance), fmt(closing)],
        textposition="outside", connector={"line": {"color": GREY}},
        increasing={"marker": {"color": GREEN}}, decreasing={"marker": {"color": RED}},
        totals={"marker": {"color": NAVY}}))
    wf.update_layout(template=PLOTLY_TEMPLATE, height=420, showlegend=False,
                     margin=dict(l=0, r=0, t=30, b=0), yaxis_title=f"Volume ({unit})")
    st.plotly_chart(wf, width="stretch")


def _balance_pivot(fdf: pd.DataFrame, index: str, columns: str, metric: str) -> pd.DataFrame:
    def _pv(col: str) -> pd.DataFrame:
        p = fdf.pivot_table(index=index, columns=columns, values=col, aggfunc="sum")
        if p.empty:
            return p
        p["TOTAL"] = p.sum(axis=1)
        p.loc["TOTAL"] = p.sum(axis=0)
        return p
    if metric == "Variance %":      # weighted: Σ variance ÷ Σ throughput (never sum percentages)
        g, t = _pv("Gain / (Loss)"), _pv("Throughput")
        return (g / t.where(t > 0)) * 100 if not g.empty else g
    return _pv(metric)


def render_balance(fdf, ledger):
    st.markdown('<div class="section-title">Department × Depot × Product Material Balance</div>',
                unsafe_allow_html=True)
    groupby = st.radio("Group by",
                       ["Department × Product", "Depot × Product", "Department × Depot"],
                       horizontal=True, label_visibility="collapsed")
    metric = st.radio("Metric",
                      ["Gain / (Loss)", "Variance %", "Actual Closing", "Receipts",
                       "Sales / Deliveries"],
                      horizontal=True, label_visibility="collapsed")

    if groupby == "Department × Product":
        index, columns = "Department", "Product"
    elif groupby == "Depot × Product":
        index, columns = "Depot", "Product"
    else:
        index, columns = "Department", "Depot"

    pivot = _balance_pivot(fdf, index, columns, metric)
    if not pivot.empty:
        fmt_str = "{:,.3f}" if metric == "Variance %" else "{:,.0f}"
        try:
            st.dataframe(pivot.style.format(fmt_str).background_gradient(
                cmap="RdYlGn_r" if metric == "Variance %" else "Blues", axis=None),
                width="stretch")
        except Exception:
            st.dataframe(pivot.style.format(fmt_str), width="stretch")
            heat = pivot.drop(index="TOTAL", errors="ignore").drop(columns="TOTAL", errors="ignore")
            if not heat.empty:
                fig = px.imshow(heat, text_auto=True, aspect="auto",
                                color_continuous_scale="RdYlGn_r" if metric == "Variance %" else "Blues",
                                template=PLOTLY_TEMPLATE, title=f"{metric} — {groupby}")
                fig.update_layout(height=max(340, 58 * len(heat)),
                                  margin=dict(l=0, r=0, t=50, b=0))
                st.plotly_chart(fig, width="stretch")

    st.markdown("---")
    st.markdown('<div class="section-title">Detailed Ledger</div>', unsafe_allow_html=True)
    q = st.text_input("Search", placeholder="Department, depot, product or reference…",
                      label_visibility="collapsed")
    view = ledger
    if q.strip():
        s = q.strip().lower()
        view = view[view["Depot"].str.lower().str.contains(s, na=False) |
                    view["Product"].str.lower().str.contains(s, na=False) |
                    view["Department"].str.lower().str.contains(s, na=False) |
                    view["Reference"].astype(str).str.lower().str.contains(s, na=False)]
    st.dataframe(view, width="stretch", height=520)


def render_exceptions(fdf, issues):
    exc = fdf[fdf["Status"].isin(EXC_STATUSES)].copy()
    exc = exc.sort_values("Variance %", key=lambda s: s.abs(), ascending=False)

    e1, e2, e3, e4 = st.columns(4)
    e1.markdown(kpi("Exceptions", f"{len(exc):,}", f"of {len(fdf):,} records",
                    "red" if len(exc) else "green"), unsafe_allow_html=True)
    e2.markdown(kpi("Total Gains",
                    fmt(fdf.loc[fdf["Gain / (Loss)"] > 0, "Gain / (Loss)"].sum()),
                    "overages", "green"), unsafe_allow_html=True)
    e3.markdown(kpi("Total Losses",
                    fmt(fdf.loc[fdf["Gain / (Loss)"] < 0, "Gain / (Loss)"].sum()),
                    "shortages", "red"), unsafe_allow_html=True)
    e4.markdown(kpi("Affected Departments",
                    f"{exc['Department'].nunique() if not exc.empty else 0}",
                    "across NPA"), unsafe_allow_html=True)

    st.markdown("")
    if exc.empty:
        st.markdown('<div class="note ok">✅ No rows breach the configured tolerance.</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="section-title">Exceptions by Department</div>',
                    unsafe_allow_html=True)
        by_dept = exc.groupby("Department", as_index=False).agg(**{
            "Exceptions": ("Department", "size"),
            "Net Variance (L)": ("Gain / (Loss)", "sum"),
        }).sort_values("Exceptions", ascending=False)
        st.dataframe(by_dept, width="stretch", hide_index=True)

        st.markdown('<div class="section-title">Exception Register</div>',
                    unsafe_allow_html=True)
        st.dataframe(exc[["Date", "Department", "Depot", "Product", "Reference",
                          "Opening Stock", "Receipts", "Sales / Deliveries",
                          "Expected Closing", "Actual Closing", "Gain / (Loss)",
                          "Variance %", "Tolerance %", "Status", "Remark"]],
                     width="stretch", height=440)

    st.markdown("---")
    st.markdown('<div class="section-title">Data Quality Observations</div>',
                unsafe_allow_html=True)
    if issues.empty:
        st.markdown('<div class="note ok">✅ No data quality issues detected.</div>',
                    unsafe_allow_html=True)
    else:
        sev = {"Critical": "bad", "High": "bad", "Medium": "warn", "Low": "info", "Info": "info"}
        for _, row in issues.iterrows():
            cls = sev.get(row["Severity"], "info")
            st.markdown(f'<div class="note {cls}"><span class="pill {cls}">'
                        f'{html.escape(row["Severity"])}</span> '
                        f'<b>{html.escape(row["Issue"])}</b><br>'
                        f'<span style="font-size:12.5px">{html.escape(row["Detail"])}</span>'
                        f'</div>', unsafe_allow_html=True)


def period_figure(pt: pd.DataFrame, kind: str, unit: str, tol: float):
    """Interactive gain/(loss) bars + variance-% line for one granularity."""
    from plotly.subplots import make_subplots
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.62, 0.38],
                        vertical_spacing=0.07)
    gl = pt["Gain / (Loss)"].fillna(0)
    fig.add_trace(go.Bar(x=pt["Period"], y=gl, name=f"Gain / (Loss) ({unit})",
                         marker_color=np.where(gl < 0, RED, GREEN)), row=1, col=1)
    fig.add_trace(go.Scatter(x=pt["Period"], y=pt["Var %"], mode="lines+markers",
                             name="Variance %", line=dict(color=NAVY, width=2)), row=2, col=1)
    fig.add_hrect(y0=-tol, y1=tol, fillcolor=GREEN, opacity=0.12, line_width=0, row=2, col=1)
    fig.update_layout(template=PLOTLY_TEMPLATE, height=420, showlegend=False,
                      margin=dict(l=10, r=10, t=30, b=10),
                      title=dict(text=f"{kind} gain / (loss) and variance %  (green band = ±{tol:.2f}% tolerance)",
                                 font=dict(size=13)))
    fig.update_yaxes(title_text=unit, row=1, col=1)
    fig.update_yaxes(title_text="Var %", row=2, col=1)
    return fig


_VERDICT_RGB = {"PASS": (21, 115, 71), "REVIEW REQUIRED": (180, 105, 14),
                "ACTION REQUIRED": (180, 35, 24)}


def render_report(fdf, issues, meta, tol_pct, tol_abs, ledger):
    st.markdown('<div class="section-title">Executive Report</div>', unsafe_allow_html=True)
    unit = meta.get("unit", "L")

    # ------------------------------------------------------------------ scope controls
    c1, c2, c3 = st.columns([1, 1.5, 1.5])
    scope_kind = c1.selectbox("Report period", ["Whole dataset"] + PERIOD_KINDS,
                              key="rep_scope_kind",
                              help="Report on everything, or on one specific day, week, month or year.")
    rdf, scope_label = fdf, "Whole dataset"
    if scope_kind != "Whole dataset":
        full_pt = build_period_table(fdf, scope_kind)
        if full_pt.empty:
            c2.warning("No dated records — showing the whole dataset.")
        else:
            labels = full_pt["Period"].tolist()[::-1]            # latest first
            pick = c2.selectbox(f"Select {scope_kind.lower()} period", labels,
                                key=f"rep_scope_{scope_kind}")
            rdf = rows_in_period(fdf, scope_kind, pick)
            scope_label = f"{scope_kind}: {pick}"
    else:
        c2.caption("Daily, weekly, monthly and yearly breakdowns are always included below.")
    codes = ["AUTO"] + [d for d in DEPARTMENT_ORDER if DEPARTMENTS[d]["kind"] == "material"]
    override = c3.selectbox(
        "Facility type", codes, key="rep_fac",
        format_func=lambda c: "Auto-detect from data" if c == "AUTO"
        else f'{DEPARTMENTS[c]["icon"]} {DEPARTMENTS[c]["label"]}',
        help="The system detects whether the data is a depot, refinery, import terminal, jetty, "
             "retail or transport file. Override only if the detection is wrong.")
    if rdf.empty:
        st.warning("No records in the selected period.")
        return
    ledger_r = ledger.loc[ledger.index.intersection(rdf.index)]

    fac = dataset_facility(rdf, meta, override)
    an = analyse_dataset(rdf, fac, unit, issues)
    s = an["stats"]
    pts = period_tables(rdf)
    ev = html.escape("; ".join(fac["evidence"][:4]) or "—")
    mix = ""
    if fac["mixed"]:
        mix = " · Mixed data: " + html.escape(", ".join(
            f"{dept_meta(k)['short']} {v:.0f}%" for k, v in list(fac["shares"].items())[:4]))
    how = "Selected" if fac["method"] == "manual" else "Auto-detected"
    st.markdown(f'<div class="note">🔎 <b>{how} data source:</b> {fac["icon"]} '
                f'<b>{html.escape(fac["label"])}</b> · confidence {fac["confidence"]}%{mix}'
                f'<br><span style="font-size:12px">Evidence: {ev}</span></div>',
                unsafe_allow_html=True)
    if fac["confidence"] < 50 and fac["method"] == "auto":
        st.caption("Low confidence — add a Department column or name the file/sheet "
                   "(e.g. “Tema Depot…”, “Refinery…”) for a more precise report, or pick the type above.")

    opening, receipts, sales, closing = s["opening"], s["receipts"], s["sales"], s["closing"]
    variance, throughput, var_pct = s["gl"], s["thr"], s["var_pct"]
    compliance, n_exc = s["compliance"], s["n_exc"]
    verdict = an["verdict"]
    status_color = {"PASS": GREEN, "REVIEW REQUIRED": AMBER, "ACTION REQUIRED": RED}[verdict]
    now = dt.datetime.now().strftime("%d %B %Y · %H:%M")
    gap = closing - s["expected"] - variance
    gap_row = (f"<tr><td>Continuity adjustment (opening ≠ previous closing)</td>"
               f"<td>{fmt(gap)}</td></tr>") if abs(gap) > max(1.0, 1e-4 * throughput) else ""

    st.markdown(f"""
    <div class="report-shell">
      <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px">
        <div><h2 style="margin:0;border:none;padding:0">NPA {html.escape(fac['short'])} Material Balance Report</h2>
             <div style="font-size:11.5px;color:{GREY};margin-top:6px">
               {html.escape(scope_label)} · Generated {now} ·
               {rdf['Department'].nunique()} facility type(s) · {rdf['Depot'].nunique()} nodes</div></div>
        <div style="text-align:right;font-size:12px;color:{GREY}">
          <div>Records: <b>{len(rdf):,}</b></div>
          <div>Tolerance: <b>±{an['tol']:.2f}%</b> or <b>±{fmt(tol_abs)} {unit}</b></div></div>
      </div>
      <div style="display:flex;gap:14px;margin:22px 0 8px 0;flex-wrap:wrap">
        <div style="flex:1;min-width:160px;background:{NAVY2};color:#fff;padding:16px 20px;border-radius:10px">
          <div style="font-size:11px;opacity:.75;text-transform:uppercase">Net Gain / (Loss)</div>
          <div style="font-size:22px;font-weight:800;margin-top:4px">{fmt(variance)} {unit}</div>
          <div style="font-size:12px;opacity:.85">{fmt_pct(var_pct)} of throughput</div></div>
        <div style="flex:1;min-width:160px;background:{status_color};color:#fff;padding:16px 20px;border-radius:10px">
          <div style="font-size:11px;opacity:.85;text-transform:uppercase">Compliance</div>
          <div style="font-size:22px;font-weight:800;margin-top:4px">{verdict}</div>
          <div style="font-size:12px;opacity:.9">{n_exc} exception(s) · {fmt_comp(compliance)} in tolerance</div></div>
        <div style="flex:1;min-width:160px;background:#f7f9fc;border:1px solid {LINE};padding:16px 20px;border-radius:10px">
          <div style="font-size:11px;color:{GREY};text-transform:uppercase">Closing Stock</div>
          <div style="font-size:22px;font-weight:800;margin-top:4px;color:{NAVY2}">{fmt(closing)} {unit}</div>
          <div style="font-size:12px;color:{GREY}">Opening {fmt(opening)} · Receipts {fmt(receipts)}</div></div>
      </div>
      <h3>Balance Overview</h3>
      <table>
        <tr><th>Item</th><th>Volume ({unit})</th></tr>
        <tr><td>Opening Stock (first record per depot / product)</td><td>{fmt(opening)}</td></tr>
        <tr><td>Total Receipts</td><td>{fmt(receipts)}</td></tr>
        <tr><td>Total Available</td><td>{fmt(opening + receipts)}</td></tr>
        <tr><td>Total Sales / Deliveries</td><td>{fmt(sales)}</td></tr>
        <tr><td>Expected Closing</td><td>{fmt(s['expected'])}</td></tr>
        <tr><td>Actual Closing (last record per depot / product)</td><td>{fmt(closing)}</td></tr>
        <tr><td><b>Gain / (Loss)</b></td><td><b>{fmt(variance)}</b></td></tr>
        {gap_row}
        <tr><td><b>Variance % of Throughput</b></td><td><b>{fmt_pct(var_pct)}</b></td></tr>
      </table>
    </div>""", unsafe_allow_html=True)

    # ------------------------------------------------------------------ facility KPIs + findings
    st.markdown(f'<div class="section-title" style="margin-top:18px">'
                f'{html.escape(fac["short"])} indicators</div>', unsafe_allow_html=True)
    kcols = st.columns(len(an["kpis"]))
    for col, (lab, val) in zip(kcols, an["kpis"]):
        col.markdown(kpi(lab, val), unsafe_allow_html=True)
    st.markdown('<div class="section-title" style="margin-top:18px">Analysis & key findings</div>',
                unsafe_allow_html=True)
    st.markdown("\n".join(f"- {f}" for f in an["findings"]))
    if fac["mixed"]:
        bd = facility_breakdown(rdf, unit)
        st.markdown("**Breakdown by facility type**")
        st.dataframe(pd.DataFrame(bd[1:], columns=bd[0]), width="stretch", hide_index=True)

    # ------------------------------------------------------------------ daily / weekly / monthly / yearly
    st.markdown('<div class="section-title" style="margin-top:18px">'
                'Daily · Weekly · Monthly · Yearly</div>', unsafe_allow_html=True)
    tabs = st.tabs(["📅 Daily", "🗓️ Weekly", "📆 Monthly", "🗂️ Yearly"])
    for tab, kind in zip(tabs, PERIOD_KINDS):
        with tab:
            pt = pts[kind]
            if pt.empty:
                st.info("No dated records — map a Transaction Date column to unlock period reports.")
                continue
            st.plotly_chart(period_figure(pt, kind, unit, an["tol"]), width="stretch",
                            key=f"rep_pf_{kind}")
            st.dataframe(period_display(pt).drop(columns=["Throughput"]),
                         width="stretch", hide_index=True,
                         height=min(420, 38 + 35 * len(pt)))
    undated = int(rdf["Date"].isna().sum())
    if undated:
        st.caption(f"{undated:,} record(s) without a valid date are excluded from the period tables.")

    # ------------------------------------------------------------------ downloads
    st.markdown("---")
    st.markdown('<div class="section-title">Download Reports</div>', unsafe_allow_html=True)

    dept_summary = rdf.groupby("Department", as_index=False).agg(**{
        "Records": ("Department", "size"), "Nodes": ("Depot", "nunique"),
        "Throughput": ("Throughput", "sum"), "Receipts": ("Receipts", "sum"),
        "Sales": ("Sales / Deliveries", "sum"), "Closing": ("Actual Closing", "sum"),
        "Variance": ("Gain / (Loss)", "sum")})
    dept_summary["Var %"] = np.where(dept_summary["Throughput"] > 0,
                                     dept_summary["Variance"] / dept_summary["Throughput"] * 100, np.nan)
    depot_summary = rdf.groupby(["Department", "Depot"], as_index=False).agg(**{
        "Opening Stock": ("Opening Stock", "sum"), "Receipts": ("Receipts", "sum"),
        "Sales / Deliveries": ("Sales / Deliveries", "sum"),
        "Expected Closing": ("Expected Closing", "sum"), "Actual Closing": ("Actual Closing", "sum"),
        "Gain / (Loss)": ("Gain / (Loss)", "sum"), "Throughput": ("Throughput", "sum")})
    depot_summary["Variance %"] = np.where(depot_summary["Throughput"] > 0,
                                           depot_summary["Gain / (Loss)"] / depot_summary["Throughput"] * 100, np.nan)
    product_summary = rdf.groupby("Product", as_index=False).agg(**{
        "Opening Stock": ("Opening Stock", "sum"), "Receipts": ("Receipts", "sum"),
        "Sales / Deliveries": ("Sales / Deliveries", "sum"),
        "Actual Closing": ("Actual Closing", "sum"), "Gain / (Loss)": ("Gain / (Loss)", "sum")})
    exception_report = rdf[rdf["Status"].isin(EXC_STATUSES)][[
        "Date", "Department", "Depot", "Product", "Reference", "Gain / (Loss)",
        "Variance %", "Tolerance %", "Status", "Remark"]]

    sheets = {"Findings": pd.DataFrame({"Finding": an["findings"]}),
              "Facility KPIs": pd.DataFrame(an["kpis"], columns=["Indicator", "Value"])}
    for kind in PERIOD_KINDS:
        if not pts[kind].empty:
            sheets[f"{kind} Report"] = pts[kind]
    sheets.update({"Department League": dept_summary, "Depot Summary": depot_summary,
                   "Product Summary": product_summary, "Exception Report": exception_report,
                   "Detailed Ledger": ledger_r, "Data Quality": issues})
    workbook = to_excel_bytes(
        sheets,
        meta={"Report": f"NPA {fac['short']} Material Balance",
              "Data source detected": f"{fac['label']} ({fac['confidence']}% confidence)",
              "Scope": scope_label, "Generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
              "Records": f"{len(rdf):,}", "Unit": unit,
              "Compliance %": fmt_comp(compliance), "Exceptions": f"{n_exc:,}",
              "Tolerance %": f"{an['tol']:.2f}"})

    kpi_rows = [("Net Gain / (Loss)", f"{fmt(variance)} {unit}"),
                ("Variance % of Throughput", fmt_pct(var_pct)),
                ("Compliance", fmt_comp(compliance)), ("Exceptions", f"{n_exc:,}"),
                ("Opening Stock", f"{fmt(opening)} {unit}"), ("Total Receipts", f"{fmt(receipts)} {unit}"),
                ("Total Sales / Deliveries", f"{fmt(sales)} {unit}"), ("Actual Closing", f"{fmt(closing)} {unit}")]
    summary_rows = [["Opening Stock (first record)", fmt(opening)], ["Total Receipts", fmt(receipts)],
                    ["Total Available", fmt(opening + receipts)], ["Total Sales / Deliveries", fmt(sales)],
                    ["Expected Closing", fmt(s["expected"])], ["Actual Closing (last record)", fmt(closing)],
                    ["Gain / (Loss)", fmt(variance)]]
    if gap_row:
        summary_rows.append(["Continuity adjustment", fmt(gap)])
    summary_rows.append(["Variance % of throughput", fmt_pct(var_pct)])
    dept_headers = ["Department", "Records", "Nodes", "Throughput", "Var %", "Status"]
    dept_rows = []
    for _, r in dept_summary.iterrows():
        m = dept_meta(r["Department"])
        st_txt = ("No data" if pd.isna(r["Var %"]) else "Within" if abs(r["Var %"]) <= m["tol"]
                  else "Loss breach" if r["Var %"] < 0 else "Gain breach")
        dept_rows.append([m["label"], f"{int(r['Records']):,}", str(int(r["Nodes"])),
                          fmt(r["Throughput"]), fmt_pct(r["Var %"]), st_txt])
    exc_rows = [[str(r["Date"])[:10] if pd.notna(r["Date"]) else "—", dept_meta(r["Department"])["short"],
                 str(r["Depot"])[:28], str(r["Product"]), fmt_pct(r["Variance %"]), r["Status"]]
                for _, r in exception_report.head(15).iterrows()]
    note_txt = (f"Compliance {fmt_comp(compliance)} — {n_exc} exception(s) recorded. "
                f"Net variance {fmt(variance)} {unit} ({fmt_pct(var_pct)}).")

    try:
        ex = cached_pdf_extras(rdf, meta, tol_pct, unit, override, issues)
        extras = {k: v for k, v in ex.items() if k != "_analysis"}
        pdf_bytes = build_pdf_report(
            title=f"{fac['short']} Material Balance Report",
            subtitle=fac["label"],
            meta_lines=[("Scope", scope_label), ("Records", f"{len(rdf):,}"),
                        ("Nodes", f"{rdf['Depot'].nunique()}"), ("Unit", unit)],
            kpi_rows=kpi_rows, summary_table=summary_rows,
            department_table=[dept_headers] + dept_rows if len(dept_rows) > 1 else None,
            exceptions=[["Date", "Department", "Depot", "Product", "Var %", "Status"]] + exc_rows
            if exc_rows else None,
            note=note_txt, verdict=(f"{verdict}  ·  {n_exc} exception(s)  ·  "
                                    f"{fmt_comp(compliance)} within tolerance", _VERDICT_RGB[verdict]),
            **extras)
        pdf_ok, pdf_err = True, ""
    except Exception as e:
        pdf_bytes, pdf_ok, pdf_err = b"", False, f"{type(e).__name__}: {e}"

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button("⬇️  Full Excel workbook", workbook, file_name=f"NPA_Material_Balance_{stamp}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", type="primary", key="rep_wb")
    if pdf_ok:
        d2.download_button("📄  PDF executive report", pdf_bytes, file_name=f"NPA_Report_{stamp}.pdf",
                           mime="application/pdf", width="stretch", key="rep_pdf")
    else:
        d2.button("📄  PDF unavailable", disabled=True, width="stretch", key="rep_pdf_dis")
        if not _HAS_FPDF:
            d2.caption("Install `fpdf2`: `pip install fpdf2`")
        elif pdf_err:
            d2.caption(f"PDF error: {pdf_err[:120]}")
    d3.download_button("⬇️  Exception report (CSV)", to_csv_bytes(exception_report),
                       file_name=f"NPA_Exceptions_{stamp}.csv", mime="text/csv",
                       width="stretch", key="rep_exc")
    d4.download_button("⬇️  Detailed ledger (CSV)", to_csv_bytes(ledger_r),
                       file_name=f"NPA_Ledger_{stamp}.csv", mime="text/csv",
                       width="stretch", key="rep_led")
    logo = get_pdf_logo()
    if logo and logo["kind"] == "png" and max(logo["orig"]) < 400:
        st.caption(f"⚠️ The logo is only {logo['orig'][0]}×{logo['orig'][1]} px, so the PDF header "
                   f"logo can look soft. Upload a PNG of 1000 px or more, or an SVG, for a razor-sharp header.")


# ==========================================================================
# USER / COMPANY / AUDIT / MY-* / DATA-AUDIT
# ==========================================================================
def render_user_management(user):
    st.markdown('<div class="section-title">User Management</div>', unsafe_allow_html=True)
    rows = get_users()
    df = pd.DataFrame([dict(r) for r in rows])
    if df.empty:
        st.info("No users yet.")
        return
    df["company_name"] = df["company_name"].fillna("NPA Head Office")

    # ---- Safe column guards (legacy DB without department / depot) --------
    if "department" not in df.columns:
        df["department"] = "—"
    else:
        df["department"] = df["department"].fillna("—")
    if "depot" not in df.columns:
        df["depot"] = ""
    else:
        df["depot"] = df["depot"].fillna("")
    # -----------------------------------------------------------------------

    counts = df["status"].value_counts()
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Total Users", f"{len(df):,}"), unsafe_allow_html=True)
    k2.markdown(kpi("Active", str(counts.get("active", 0)), "approved", "green"),
                unsafe_allow_html=True)
    k3.markdown(kpi("Pending", str(counts.get("pending", 0)), "awaiting", "amber"),
                unsafe_allow_html=True)
    k4.markdown(kpi("Suspended", str(counts.get("suspended", 0)), "blocked", "red"),
                unsafe_allow_html=True)
    st.markdown("")
    status_filter = st.multiselect("Filter by status", ["active", "pending", "suspended"],
                                   default=["active", "pending", "suspended"])
    view = df[df["status"].isin(status_filter)]

    show_cols = [c for c in ["id", "full_name", "email", "staff_id", "role",
                             "company_name", "depot", "department", "status",
                             "locked_until", "created_at", "last_login"] if c in df.columns]
    st.dataframe(view[show_cols], width="stretch", height=340)
    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Change status**")
        t = st.selectbox("User", view["email"].tolist(), key="um_t")
        s = st.selectbox("New status", ["active", "pending", "suspended"], key="um_s")
        if st.button("Apply", width="stretch", key="um_apply"):
            db_execute("UPDATE users SET status=? WHERE email=?", (s, t))
            log_action(user["email"], "user_status_change", f"{t} → {s}")
            st.success(f"{t} is now {s}.")
            st.rerun()
    with c2:
        st.markdown("**Change role / department**")
        t = st.selectbox("User", view["email"].tolist(), key="um_rt")
        r = st.selectbox("New role", list(ROLES.keys()), key="um_rr")
        d = st.selectbox("Department", [""] + DEPARTMENT_ORDER,
                         format_func=lambda x: "— none —" if x == ""
                         else f'{DEPARTMENTS[x]["icon"]} {DEPARTMENTS[x]["label"]}',
                         key="um_rd")
        if st.button("Update", width="stretch", key="um_r"):
            db_execute("UPDATE users SET role=?, department=? WHERE email=?",
                       (r, d or None, t))
            log_action(user["email"], "user_role_change",
                       f"{t} → role={r} dept={d}")
            st.success(f"{t} updated.")
            st.rerun()
    with c3:
        st.markdown("**Reset password**")
        t = st.selectbox("User", view["email"].tolist(), key="um_pt")
        pw = st.text_input("New password", type="password", key="um_pw")
        if st.button("Reset", width="stretch", key="um_p"):
            if len(pw) < 8:
                st.error("Password must be at least 8 characters.")
            else:
                db_execute("UPDATE users SET password_hash=? WHERE email=?",
                           (hash_password(pw), t))
                log_action(user["email"], "password_reset", t)
                st.success("Password updated.")
    _user_admin_tools(user, df)


def render_company_management(user):
    st.markdown('<div class="section-title">Company Management</div>', unsafe_allow_html=True)
    companies = get_companies()
    df = pd.DataFrame([dict(c) for c in companies])
    if not df.empty:
        st.dataframe(df, width="stretch", height=320)
    _company_admin_tools(user, df)
    st.markdown("---")
    with st.form("company_form"):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Company name")
            code = st.text_input("Company code")
            status = st.selectbox("Status", ["active", "pending", "suspended"])
        with c2:
            email = st.text_input("Contact email")
            phone = st.text_input("Contact phone")
            address = st.text_area("Address", height=90)
        submit = st.form_submit_button("Save Company")
    if submit:
        if not name.strip():
            st.error("Company name is required.")
        else:
            try:
                db_execute("""INSERT INTO companies
                              (name, code, contact_email, contact_phone, address, status)
                              VALUES (?,?,?,?,?,?)""",
                           (name.strip(), code.strip().upper(), email.strip(),
                            phone.strip(), address.strip(), status))
                log_action(user["email"], "company_created", name.strip())
                st.success(f"Company {name} added.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Company already exists.")


def render_audit_log():
    st.markdown('<div class="section-title">Audit Log</div>', unsafe_allow_html=True)
    rows = db_query("SELECT * FROM audit_log ORDER BY id DESC LIMIT 500")
    if not rows:
        st.info("No activity yet.")
        return
    df = pd.DataFrame([dict(r) for r in rows])
    st.dataframe(df, width="stretch", height=560)
    st.download_button("⬇️  Download audit log (CSV)", to_csv_bytes(df),
                       file_name=f"NPA_Audit_{dt.datetime.now():%Y%m%d}.csv",
                       mime="text/csv", width="stretch")


def render_my_depots(fdf):
    st.markdown('<div class="section-title">Depot Overview</div>', unsafe_allow_html=True)
    depots = fdf.groupby(["Department", "Depot"], as_index=False).agg(**{
        "Records": ("Depot", "size"),
        "Opening Stock": ("Opening Stock", "sum"),
        "Receipts": ("Receipts", "sum"),
        "Sales / Deliveries": ("Sales / Deliveries", "sum"),
        "Actual Closing": ("Actual Closing", "sum"),
        "Gain / (Loss)": ("Gain / (Loss)", "sum")})
    st.dataframe(depots, width="stretch")
    fig = px.bar(depots, x="Depot", y="Gain / (Loss)", color="Department",
                 template=PLOTLY_TEMPLATE, title="Variance by Depot",
                 color_discrete_map={d: DEPARTMENTS[d]["color"]
                                     for d in depots["Department"].unique()})
    fig.update_layout(height=380, margin=dict(l=0, r=0, t=50, b=0))
    st.plotly_chart(fig, width="stretch")


def render_my_staff(user):
    st.markdown('<div class="section-title">My Team</div>', unsafe_allow_html=True)
    rows = get_users(user.get("company_id"))
    df = pd.DataFrame([dict(r) for r in rows])
    if df.empty:
        st.info("No staff yet.")
        return

    # ---- Safe column guards ----------------------------------------------
    if "department" not in df.columns:
        df["department"] = "—"
    else:
        df["department"] = df["department"].fillna("—")
    if "depot" not in df.columns:
        df["depot"] = ""
    else:
        df["depot"] = df["depot"].fillna("")
    # -----------------------------------------------------------------------

    show_cols = [c for c in ["full_name", "email", "staff_id", "depot",
                             "department", "role", "status", "last_login"]
                 if c in df.columns]
    st.dataframe(df[show_cols], width="stretch", height=420)
    _company_staff_tools(user, df)


def render_my_submissions(fdf, user):
    st.markdown('<div class="section-title">My Submissions</div>', unsafe_allow_html=True)
    st.caption(f"Depot: **{user.get('depot') or 'Unassigned'}** · "
               f"Department: **{user.get('department') or '—'}**")
    render_submissions_log(user, embed=True)
    st.markdown("---")
    st.markdown('<div class="section-title">Current Upload</div>', unsafe_allow_html=True)
    if fdf is None:
        st.markdown('<div class="note">Upload a file in the sidebar to see its balance here.</div>',
                    unsafe_allow_html=True)
        template_download_button("mysub_tpl")
        return
    st.dataframe(fdf[["Date", "Department", "Depot", "Product", "Opening Stock",
                      "Receipts", "Sales / Deliveries", "Actual Closing",
                      "Gain / (Loss)", "Status"]],
                 width="stretch", height=480)


def render_data_audit(raw, mapping, issues, unit, clean_log=None, balance=None):
    st.markdown('<div class="section-title">Cleaning Report</div>', unsafe_allow_html=True)
    if clean_log is None or clean_log.empty:
        st.success("The files were already clean — no corrections were needed.")
    else:
        st.caption("Every correction applied before the material balance was calculated. "
                   "Your original files are never modified.")
        st.dataframe(clean_log, width="stretch", hide_index=True)
    if balance is not None and not balance.empty:
        st.download_button("⬇️  Download cleaned dataset (Excel)",
                           to_excel_bytes({"Cleaned Data": balance.drop(columns=["Remark"]),
                                           "Cleaning Log": clean_log if clean_log is not None else pd.DataFrame(),
                                           "Data Quality": issues}),
                           file_name=f"NPA_Cleaned_Data_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="clean_dl")
    st.markdown("---")
    st.markdown('<div class="section-title">Source Data Preview</div>', unsafe_allow_html=True)
    st.caption(f"{len(raw):,} raw rows · volumes analysed in **{unit}**")
    st.dataframe(raw.drop(columns=[c for c in ("__row",) if c in raw.columns]).head(500),
                 width="stretch", height=380)
    st.markdown("---")
    st.markdown('<div class="section-title">Column Mapping</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([{
        "Field": FIELD_LABELS[f],
        "Mapped column": mapping.get(f) or "— not mapped —",
        "Required": "yes" if f in REQUIRED_FIELDS or f in BALANCE_REQUIRES else "optional"
    } for f in PRIORITY]), width="stretch", hide_index=True)
    st.markdown("---")
    st.markdown('<div class="section-title">Data Quality Log</div>', unsafe_allow_html=True)
    if issues.empty:
        st.success("No issues detected.")
    else:
        st.dataframe(issues, width="stretch", hide_index=True)


# ==========================================================================
# v3.2 — NEW SECTIONS: footer, template, upload logging, trends, profile,
#        help, submissions log, admin tools
# ==========================================================================
def _s(v) -> str:
    """NaN/None-safe string for form defaults."""
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def render_footer() -> None:
    st.markdown(f"""
    <div class="app-footer">{logo_html(24, extra_style="box-shadow:none")}
      <span>National Petroleum Authority · Enterprise Material Balance Portal v3.2</span>
    </div>""", unsafe_allow_html=True)


def build_template_bytes() -> bytes:
    cols = ["Transaction Date", "Depot", "Department", "Node Type", "Product",
            "Opening Stock", "Receipts", "Transfers In", "Sales", "Transfers Out",
            "Closing Stock", "Units", "Reference"]
    today = dt.date.today()
    sample = pd.DataFrame([
        [today - dt.timedelta(days=1), "Tema Depot", "DEPOT", "DEPOT", "PREMIUM",
         1_000_000, 500_000, 0, 450_000, 0, 1_049_200, "L", "TKT-0001"],
        [today - dt.timedelta(days=1), "MT Example Vessel", "IMPORT", "VESSEL", "GASOIL",
         0, 2_000_000, 0, 1_996_000, 0, 3_500, "L", "BL-2026-014"],
    ], columns=cols)
    guide = pd.DataFrame([
        ["Transaction Date", "Recommended", "Any recognisable date (YYYY-MM-DD preferred). Needed for trends."],
        ["Depot", "Required", "Depot, terminal, vessel, jetty, outlet or node name."],
        ["Department", "Optional", "One of the codes on the Departments sheet. Inferred from the depot name if blank."],
        ["Node Type", "Optional", "DEPOT, VESSEL, JETTY, REFINERY, RETAIL, PIPELINE/FLEET …"],
        ["Product", "Required", "PREMIUM, DIESEL/GASOIL, LPG, KEROSENE, …"],
        ["Opening Stock", "Required", "Stock at start of the period."],
        ["Receipts", "Required", "Volume received (discharges, inflows)."],
        ["Transfers In / Out", "Optional", "Inter-depot movements."],
        ["Sales", "Required", "Deliveries / loadings / issues."],
        ["Closing Stock", "Required*", "Physical (dip) stock. *Or supply a Gain/Loss column instead."],
        ["Units", "Optional", "L, M3, GAL or BBL — converted to litres automatically."],
        ["Reference", "Optional", "Ticket, waybill or B/L number."],
    ], columns=["Column", "Requirement", "Notes"])
    depts = pd.DataFrame([{"Code": c, "Department": DEPARTMENTS[c]["label"],
                           "Tolerance %": DEPARTMENTS[c]["tol"],
                           "Scope": DEPARTMENTS[c]["scope"]} for c in DEPARTMENT_ORDER])
    return to_excel_bytes({"Template": sample, "Guide": guide, "Departments": depts})


def template_download_button(key: str, label: str = "⬇️  Download upload template") -> None:
    st.download_button(label, build_template_bytes(),
                       file_name="NPA_Material_Balance_Template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=key, width="stretch")


def _log_new_uploads(user: dict, balance: pd.DataFrame, uploaded) -> None:
    """Write one submissions row per newly-seen upload (once per session per file)."""
    seen = st.session_state.setdefault("logged_uploads", set())
    for f in uploaded or []:
        key = f"{user['id']}:{f.name}:{f.size}"
        if key in seen:
            continue
        sub = balance[balance["Source File"] == f.name]
        if sub.empty:
            continue
        record_submission(user, f.name, len(sub), sub["Throughput"].sum(),
                          sub["Gain / (Loss)"].sum(),
                          int(sub["Status"].isin(EXC_STATUSES).sum()))
        log_action(user["email"], "upload", f"{f.name} · {len(sub):,} rows")
        seen.add(key)


def render_no_data_page(page: str, role: str, user: dict) -> None:
    manager_only = {"User Management", "Company Management", "Audit Log"}
    if page in manager_only and role != "manager":
        st.error("You are not authorised to view this page.")
    elif page == "User Management":
        render_user_management(user)
    elif page == "Company Management":
        render_company_management(user)
    elif page == "Audit Log":
        render_audit_log()
    elif page == "My Staff" and role == "company":
        render_my_staff(user)
    elif page == "Submissions Log":
        render_submissions_log(user)
    elif page == "My Submissions":
        render_my_submissions(None, user)
    elif page == "My Profile":
        render_profile(user)
    elif page == "Help & Template":
        render_help()


# ---------------------------------------------------------------- Trends
def render_trends(fdf: pd.DataFrame, tol_pct: float, unit: str) -> None:
    from plotly.subplots import make_subplots

    st.markdown('<div class="section-title">Trends & Stock Cover</div>', unsafe_allow_html=True)
    dated = fdf.dropna(subset=["Date"]).copy()
    if dated.empty:
        st.markdown('<div class="note warn">No usable dates found. Map a <b>Transaction Date</b> '
                    'column under <i>Column Mapping</i> to unlock trend analysis.</div>',
                    unsafe_allow_html=True)
        return

    freq = st.radio("Granularity", ["Daily", "Weekly", "Monthly"], horizontal=True,
                    key="trend_freq", label_visibility="collapsed")
    per = {"Daily": "D", "Weekly": "W", "Monthly": "M"}[freq]
    dated["Period"] = dated["Date"].dt.to_period(per).dt.start_time

    ts = dated.groupby("Period", as_index=False).agg(**{
        "Throughput": ("Throughput", "sum"),
        "Variance": ("Gain / (Loss)", "sum"),
        "Exceptions": ("Status", lambda x: int(x.isin(EXC_STATUSES).sum())),
    })
    ts["Var %"] = np.where(ts["Throughput"] > 0, ts["Variance"] / ts["Throughput"] * 100, np.nan)
    ts["Rolling Var % (3)"] = ts["Var %"].rolling(3, min_periods=1).mean()

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Periods", f"{len(ts):,}",
                    f"{ts['Period'].min():%d %b %Y} → {ts['Period'].max():%d %b %Y}"),
                unsafe_allow_html=True)
    worst = ts.loc[ts["Variance"].idxmin()] if ts["Variance"].notna().any() else None
    k2.markdown(kpi("Worst period", fmt(worst["Variance"]) if worst is not None else "—",
                    f"{worst['Period']:%d %b %Y}" if worst is not None else "", "red"),
                unsafe_allow_html=True)
    k3.markdown(kpi("Avg variance %", fmt_pct(ts["Var %"].mean()), "per period",
                    "green" if pd.notna(ts["Var %"].mean()) and abs(ts["Var %"].mean()) <= tol_pct else "amber"),
                unsafe_allow_html=True)
    k4.markdown(kpi("Exceptions", f"{int(ts['Exceptions'].sum()):,}", "all periods",
                    "red" if ts["Exceptions"].sum() else "green"), unsafe_allow_html=True)

    st.markdown("")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=ts["Period"], y=ts["Throughput"], name=f"Throughput ({unit})",
                marker_color="rgba(47,111,176,.55)", secondary_y=False)
    fig.add_scatter(x=ts["Period"], y=ts["Var %"], name="Variance %", mode="lines+markers",
                    line=dict(color=RED, width=2), secondary_y=True)
    fig.add_scatter(x=ts["Period"], y=ts["Rolling Var % (3)"], name="3-period average",
                    mode="lines", line=dict(color=GOLD, width=3, dash="dot"), secondary_y=True)
    fig.add_hrect(y0=-tol_pct, y1=tol_pct, fillcolor=GREEN, opacity=.08, line_width=0,
                  secondary_y=True)
    fig.update_layout(template=PLOTLY_TEMPLATE, height=420, legend_title="",
                      title="Throughput vs variance % (green band = default tolerance)",
                      margin=dict(l=0, r=0, t=60, b=0))
    fig.update_yaxes(title_text=f"Throughput ({unit})", secondary_y=False)
    fig.update_yaxes(title_text="Variance %", secondary_y=True)
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        by_prod = dated.groupby(["Period", "Product"], as_index=False).agg(**{
            "T": ("Throughput", "sum"), "V": ("Gain / (Loss)", "sum")})
        by_prod["Var %"] = np.where(by_prod["T"] > 0, by_prod["V"] / by_prod["T"] * 100, np.nan)
        f2 = px.line(by_prod, x="Period", y="Var %", color="Product", markers=True,
                     template=PLOTLY_TEMPLATE, color_discrete_sequence=PRODUCT_COLORS,
                     title="Variance % by product")
        f2.update_layout(height=380, legend_title="", margin=dict(l=0, r=0, t=50, b=0))
        st.plotly_chart(f2, width="stretch")
    with c2:
        cum = ts.assign(Cumulative=ts["Variance"].cumsum())
        f3 = px.area(cum, x="Period", y="Cumulative", template=PLOTLY_TEMPLATE,
                     title=f"Cumulative gain / (loss) ({unit})")
        f3.update_traces(line_color=NAVY, fillcolor="rgba(18,59,99,.18)")
        f3.update_layout(height=380, margin=dict(l=0, r=0, t=50, b=0))
        st.plotly_chart(f3, width="stretch")

    st.markdown('<div class="section-title">Stock Cover by Product</div>', unsafe_allow_html=True)
    span_days = max(1, (dated["Date"].max() - dated["Date"].min()).days + 1)
    last_day = dated[dated["Date"] == dated["Date"].max()]
    cover = last_day.groupby("Product", as_index=False)["Actual Closing"].sum()
    daily_sales = (dated.groupby("Product")["Sales / Deliveries"].sum() / span_days)
    cover["Avg daily sales"] = cover["Product"].map(daily_sales)
    cover["Cover (days)"] = np.where(cover["Avg daily sales"] > 0,
                                     cover["Actual Closing"] / cover["Avg daily sales"], np.nan)
    if cover["Cover (days)"].notna().any():
        f4 = px.bar(cover.sort_values("Cover (days)"), x="Cover (days)", y="Product",
                    orientation="h", template=PLOTLY_TEMPLATE, color="Cover (days)",
                    color_continuous_scale="RdYlGn", title="Days of cover at latest reported date")
        f4.update_layout(height=max(300, 46 * len(cover)), coloraxis_showscale=False,
                         margin=dict(l=0, r=0, t=50, b=0))
        st.plotly_chart(f4, width="stretch")
    st.dataframe(cover, width="stretch", hide_index=True)

    st.markdown('<div class="section-title">Persistent Loss Hot-spots</div>', unsafe_allow_html=True)
    loss = fdf[fdf["Status"] == "Loss (Short)"]
    if loss.empty:
        st.markdown('<div class="note ok">✅ No loss events outside tolerance.</div>',
                    unsafe_allow_html=True)
    else:
        hot = loss.groupby(["Department", "Depot", "Product"], as_index=False).agg(**{
            "Loss events": ("Status", "size"), "Total loss": ("Gain / (Loss)", "sum"),
            "Worst %": ("Variance %", "min")}).sort_values(
            ["Loss events", "Total loss"], ascending=[False, True]).head(15)
        st.dataframe(hot, width="stretch", hide_index=True)


# ---------------------------------------------------------------- Profile
def render_profile(user: dict) -> None:
    st.markdown('<div class="section-title">My Profile</div>', unsafe_allow_html=True)
    rows = db_query("""SELECT u.*, c.name AS company_name FROM users u
                       LEFT JOIN companies c ON c.id=u.company_id WHERE u.id=?""", (user["id"],))
    if not rows:
        st.error("Account not found.")
        return
    me = dict(rows[0])
    d = dept_meta(me.get("department") or "")
    left, right = st.columns([1, 1.2], gap="large")
    with left:
        st.markdown(f"""
        <div class="report-shell">
          <div style="display:flex;align-items:center;gap:14px">{logo_html(60)}
            <div><div style="font-size:18px;font-weight:800;color:{NAVY2}">{html.escape(me['full_name'])}</div>
                 <div style="font-size:12.5px;color:{GREY}">{html.escape(me['email'])}</div></div></div>
          <div style="margin:14px 0 10px 0">{role_badge(me['role'])}
            {('&nbsp;' + dept_badge(me['department'])) if me.get('department') else ''}</div>
          <table>
            <tr><td>Staff ID</td><td>{html.escape(_s(me['staff_id']))}</td></tr>
            <tr><td>Company</td><td>{html.escape(_s(me.get('company_name')) or 'NPA Head Office')}</td></tr>
            <tr><td>Depot / Terminal</td><td>{html.escape(_s(me.get('depot')) or '—')}</td></tr>
            <tr><td>Department</td><td>{html.escape(d['label']) if me.get('department') else '—'}</td></tr>
            <tr><td>Status</td><td>{html.escape(_s(me['status']))}</td></tr>
            <tr><td>Member since</td><td>{html.escape(_s(me.get('created_at'))[:10])}</td></tr>
            <tr><td>Last sign-in</td><td>{html.escape(_s(me.get('last_login'))[:16] or '—')}</td></tr>
          </table>
        </div>""", unsafe_allow_html=True)
    with right:
        with st.form("profile_form"):
            st.markdown("**Contact details**")
            phone = st.text_input("Phone", value=_s(me.get("phone")))
            depot = st.text_input("Depot / Terminal", value=_s(me.get("depot")),
                                  disabled=(me["role"] == "manager"))
            save = st.form_submit_button("Save details")
        if save:
            db_execute("UPDATE users SET phone=?, depot=? WHERE id=?",
                       (phone.strip(), depot.strip(), me["id"]))
            log_action(me["email"], "profile_update", "contact details")
            fresh = db_query("SELECT * FROM users WHERE id=?", (me["id"],))
            st.session_state.user = dict(fresh[0])
            st.success("Details updated.")

        with st.form("password_form"):
            st.markdown("**Change password**")
            cur_pw = st.text_input("Current password", type="password")
            new_pw = st.text_input("New password (min 8 characters)", type="password")
            conf_pw = st.text_input("Confirm new password", type="password")
            change = st.form_submit_button("Update password")
        if change:
            if not verify_password(cur_pw, me["password_hash"]):
                st.error("Current password is incorrect.")
            elif len(new_pw) < 8:
                st.error("New password must be at least 8 characters.")
            elif new_pw != conf_pw:
                st.error("New passwords do not match.")
            elif new_pw == cur_pw:
                st.error("Choose a password different from the current one.")
            else:
                db_execute("UPDATE users SET password_hash=? WHERE id=?",
                           (hash_password(new_pw), me["id"]))
                log_action(me["email"], "password_change", "self-service")
                st.success("Password updated.")


# ---------------------------------------------------------------- Help
def render_help() -> None:
    st.markdown('<div class="section-title">Help & Upload Template</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.3, 1], gap="large")
    with c1:
        st.markdown("""
**How it works**

1. Download the template (right), fill in one row per depot × product × period.
2. Upload it from the sidebar — Excel (several sheets allowed) or CSV.
3. Check **Column Mapping** if a column was not recognised automatically.
4. Review **Dashboard**, **Trends**, **Exceptions** and the **Executive Report**.

**How the balance is computed**

`Expected closing = Opening + Receipts + Transfers in − Sales − Transfers out`  
`Gain / (Loss) = Actual closing − Expected closing`  
`Variance % = Gain / (Loss) ÷ Throughput × 100`

A row is *within tolerance* when |Variance %| is at or below its department tolerance
(or the absolute litre tolerance set in the sidebar).
""")
    with c2:
        st.markdown('<div class="note">📄 Includes a sample, a column guide and the '
                    'department code list.</div>', unsafe_allow_html=True)
        template_download_button("help_tpl")
        st.caption("Any layout is accepted. Units L, KL, M3, GAL, BBL are auto-converted to litres.")

    st.markdown('<div class="section-title">Departments & Tolerances</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([{
        "Code": c, "Department": f'{DEPARTMENTS[c]["icon"]} {DEPARTMENTS[c]["label"]}',
        "Tolerance": f'±{DEPARTMENTS[c]["tol"]}%',
        "Type": "Material balance" if DEPARTMENTS[c]["kind"] == "material" else "Cross-cutting",
        "Scope": DEPARTMENTS[c]["scope"]} for c in DEPARTMENT_ORDER]),
        width="stretch", hide_index=True)

    with st.expander("Frequently asked questions"):
        st.markdown("""
**My columns were not detected.** Use *Column Mapping* above the dashboard to choose them manually.

**Why do I see no trends?** Trends need a mapped date column.

**I forgot my password.** Ask a Manager to reset it from *User Management*. Accounts lock for
15 minutes after 5 failed sign-ins.

**Who can see my uploads?** Managers see every submission; companies see their own staff's; staff see their own.
""")


# ---------------------------------------------------------------- Submissions
def render_submissions_log(user: dict, embed: bool = False) -> None:
    role = user["role"]
    if role == "manager":
        rows = db_query("SELECT * FROM submissions ORDER BY id DESC LIMIT 1000")
        scope = "all users"
    elif role == "company":
        rows = db_query("SELECT * FROM submissions WHERE company_id=? ORDER BY id DESC LIMIT 1000",
                        (user.get("company_id"),))
        scope = "your company"
    else:
        rows = db_query("SELECT * FROM submissions WHERE lower(user_email)=lower(?) "
                        "ORDER BY id DESC LIMIT 1000", (user["email"],))
        scope = "you"
    title = "Upload History" if embed else "Submissions Log"
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if not rows:
        st.markdown(f'<div class="note">No uploads recorded yet for {scope}. '
                    'Files are logged automatically the first time you upload them.</div>',
                    unsafe_allow_html=True)
        return
    df = pd.DataFrame([dict(r) for r in rows]).rename(columns={
        "created_at": "Uploaded", "user_email": "User", "full_name": "Name",
        "filename": "File", "row_count": "Rows", "throughput": "Throughput",
        "variance": "Gain / (Loss)", "exceptions": "Exceptions",
        "depot": "Depot", "department": "Department"})
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Uploads", f"{len(df):,}", f"scope: {scope}"), unsafe_allow_html=True)
    k2.markdown(kpi("Rows submitted", fmt(df["Rows"].sum())), unsafe_allow_html=True)
    k3.markdown(kpi("Net variance", fmt(df["Gain / (Loss)"].sum()), "all uploads",
                    "red" if df["Gain / (Loss)"].sum() < 0 else "green"), unsafe_allow_html=True)
    k4.markdown(kpi("Last upload", _s(df["Uploaded"].iloc[0])[:16], _s(df["User"].iloc[0])),
                unsafe_allow_html=True)
    st.markdown("")
    cols = [c for c in ["Uploaded", "Name", "User", "Depot", "Department", "File", "Rows",
                        "Throughput", "Gain / (Loss)", "Exceptions"] if c in df.columns]
    st.dataframe(df[cols], width="stretch", hide_index=True, height=360)
    if not embed:
        st.download_button("⬇️  Download submissions (CSV)", to_csv_bytes(df[cols]),
                           file_name=f"NPA_Submissions_{dt.datetime.now():%Y%m%d}.csv",
                           mime="text/csv", key="subs_csv")


# ---------------------------------------------------------------- Admin tools
def _user_admin_tools(user: dict, df: pd.DataFrame) -> None:
    st.markdown("---")
    st.markdown('<div class="section-title">Approvals & Account Tools</div>', unsafe_allow_html=True)
    a1, a2 = st.columns(2, gap="large")
    pending = df[df["status"] == "pending"]
    with a1:
        st.markdown(f"**Pending approvals ({len(pending)})**")
        if pending.empty:
            st.caption("Nothing is waiting for approval.")
        else:
            st.dataframe(pending[["full_name", "email", "role", "company_name", "department"]],
                         hide_index=True, width="stretch",
                         height=min(260, 60 + 36 * len(pending)))
            if st.button(f"✅  Approve all {len(pending)} pending", key="um_aprv_all",
                         width="stretch", type="primary"):
                db_execute("UPDATE users SET status='active' WHERE status='pending'")
                db_execute("""UPDATE companies SET status='active' WHERE status='pending' AND id IN
                              (SELECT company_id FROM users WHERE role='company' AND status='active')""")
                log_action(user["email"], "bulk_approve", f"{len(pending)} users")
                st.rerun()
    with a2:
        st.markdown("**Unlock or delete an account**")
        t = st.selectbox("User", df["email"].tolist(), key="um_xt")
        b1, b2 = st.columns(2)
        if b1.button("🔓  Unlock", key="um_unlock", width="stretch"):
            db_execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE email=?", (t,))
            log_action(user["email"], "user_unlock", t)
            st.success(f"{t} unlocked.")
        confirm = st.checkbox("I understand deletion is permanent", key="um_xc")
        if b2.button("🗑️  Delete", key="um_del", width="stretch"):
            target = db_query("SELECT role FROM users WHERE email=?", (t,))
            managers = db_query("SELECT COUNT(*) AS n FROM users WHERE role='manager'")[0]["n"]
            if t.lower() == user["email"].lower():
                st.error("You cannot delete your own account.")
            elif not confirm:
                st.warning("Tick the confirmation box first.")
            elif target and target[0]["role"] == "manager" and managers <= 1:
                st.error("Cannot delete the last manager.")
            else:
                db_execute("DELETE FROM users WHERE email=?", (t,))
                log_action(user["email"], "user_deleted", t)
                st.success(f"{t} deleted.")
                st.rerun()


def _company_admin_tools(user: dict, df: pd.DataFrame) -> None:
    if df.empty:
        return
    st.markdown("---")
    st.markdown('<div class="section-title">Approve · Edit · Remove</div>', unsafe_allow_html=True)
    sel = st.selectbox("Company", df["name"].tolist(), key="cm_sel")
    row = df[df["name"] == sel].iloc[0]
    cid = int(row["id"])
    statuses = ["active", "pending", "suspended"]
    n_users = db_query("SELECT COUNT(*) AS n FROM users WHERE company_id=?", (cid,))[0]["n"]
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown(f"**Status** — {n_users} linked user(s)")
        new_status = st.selectbox("New status", statuses,
                                  index=statuses.index(row["status"]) if row["status"] in statuses else 0,
                                  key=f"cm_st_{cid}")
        cascade = st.checkbox("Also activate its pending users", value=True, key=f"cm_cas_{cid}")
        st.caption("Suspending a company also suspends its non-manager users.")
        if st.button("Apply status", key=f"cm_apply_{cid}", width="stretch"):
            db_execute("UPDATE companies SET status=? WHERE id=?", (new_status, cid))
            if new_status == "active" and cascade:
                db_execute("UPDATE users SET status='active' WHERE company_id=? AND status='pending'", (cid,))
            elif new_status == "suspended":
                db_execute("UPDATE users SET status='suspended' WHERE company_id=? AND role!='manager'", (cid,))
            log_action(user["email"], "company_status", f"{sel} → {new_status}")
            st.rerun()
        protected = sel == "NPA Head Office" or n_users > 0
        if st.button("🗑️  Delete company", key=f"cm_del_{cid}", width="stretch",
                     disabled=protected,
                     help="Only companies with no linked users can be deleted."):
            db_execute("DELETE FROM companies WHERE id=?", (cid,))
            log_action(user["email"], "company_deleted", sel)
            st.rerun()
    with c2:
        with st.form(f"cm_edit_{cid}"):
            st.markdown("**Edit details**")
            name = st.text_input("Name", value=_s(row["name"]))
            code = st.text_input("Code", value=_s(row.get("code")))
            email = st.text_input("Contact email", value=_s(row.get("contact_email")))
            phone = st.text_input("Contact phone", value=_s(row.get("contact_phone")))
            address = st.text_area("Address", value=_s(row.get("address")), height=70)
            save = st.form_submit_button("Save changes")
        if save:
            if not name.strip():
                st.error("Company name is required.")
            else:
                try:
                    db_execute("""UPDATE companies SET name=?, code=?, contact_email=?,
                                  contact_phone=?, address=? WHERE id=?""",
                               (name.strip(), code.strip().upper() or None, email.strip(),
                                phone.strip(), address.strip(), cid))
                    log_action(user["email"], "company_edited", name.strip())
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Another company already uses that name or code.")


def _company_staff_tools(user: dict, df: pd.DataFrame) -> None:
    """Companies may approve / suspend *staff* accounts of their own company."""
    staff = df[(df["role"] == "staff")]
    st.markdown("---")
    st.markdown('<div class="section-title">Approve / Suspend Staff</div>', unsafe_allow_html=True)
    if staff.empty:
        st.caption("No staff accounts registered under your company yet.")
        return
    c1, c2, c3 = st.columns([2, 1, 1])
    who = c1.selectbox("Staff member", staff["email"].tolist(), key="ms_who")
    new_status = c2.selectbox("Status", ["active", "pending", "suspended"], key="ms_st")
    c3.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if c3.button("Apply", key="ms_apply", width="stretch"):
        db_execute("UPDATE users SET status=? WHERE email=? AND company_id=? AND role='staff'",
                   (new_status, who, user.get("company_id")))
        log_action(user["email"], "staff_status_change", f"{who} → {new_status}")
        st.success(f"{who} is now {new_status}.")
        st.rerun()


# ==========================================================================
# Router
# ==========================================================================
if st.session_state.user is None:
    render_auth_screen()
else:
    u = dict(st.session_state.user)
    if u.get("company_id") and not u.get("company_name"):
        rows = db_query("SELECT name FROM companies WHERE id=?", (u["company_id"],))
        if rows:
            u["company_name"] = rows[0]["name"]
    render_app(u)