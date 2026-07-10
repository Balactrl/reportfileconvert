"""
Convert Sales Summary Report Excel files to Sales Transaction Summary text format.
"""

import os
import re
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
# Ensure the script's directory is on sys.path so local modules can be imported
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))

# Diagnostic info to help Streamlit deploy debugging
try:
    print("DEPLOY DEBUG: __file__:", Path(__file__).resolve())
    print("DEPLOY DEBUG: cwd:", os.getcwd())
    print("DEPLOY DEBUG: sys.path (head):", sys.path[:6])
    print("DEPLOY DEBUG: files in app dir:", [p.name for p in Path(__file__).resolve().parent.iterdir()])
except Exception:
    pass

# Robust import for ELTester: try normal import, then search parent folders and load by path
try:
    from ELTester import process_files, convert_to_excel
except Exception:
    import importlib.util

    def _load_module_from_search(name, filename):
        current = Path(__file__).resolve().parent
        for _ in range(5):
            candidate = current / filename
            if candidate.exists():
                spec = importlib.util.spec_from_file_location(name, str(candidate))
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                return mod
            current = current.parent
        return None

    _mod = _load_module_from_search("ELTester", "ELTester.py")
    if _mod:
        process_files = _mod.process_files
        convert_to_excel = _mod.convert_to_excel
    else:
        raise

PAGE_WIDTH = 180


def format_currency(value):
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return "0.00"


def _to_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.lower() in {"nan", "none"}:
        return 0
    try:
        return float(text)
    except ValueError:
        return 0


def _to_int(value):
    return int(round(_to_number(value)))


def _fix_line(line, width=PAGE_WIDTH):
    clean = line.rstrip("\n")
    if len(clean) < width:
        return clean + " " * (width - len(clean))
    return clean[:width]


def _normalize_col(name):
    return re.sub(r"[^A-Z0-9]", "", str(name).upper())


def _find_column(columns, aliases):
    normalized = {_normalize_col(col): col for col in columns}
    for alias in aliases:
        key = _normalize_col(alias)
        if key in normalized:
            return normalized[key]
    return None


def _parse_date_range(text):
    match = re.search(
        r"from\s+(\d{1,2}[-/]\w{3}[-/]\d{4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})\s+to\s+"
        r"(\d{1,2}[-/]\w{3}[-/]\d{4}|\d{4}[-/]\d{1,2}[-/]\d{1,2})",
        str(text),
        re.IGNORECASE,
    )
    if not match:
        return "", ""

    def convert_date(raw):
        raw = raw.strip()
        for fmt in ("%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return raw

    return convert_date(match.group(1)), convert_date(match.group(2))


def _clean_site_id(value):
    """Normalize site id from Excel numeric/text values like 10001.0 -> 10001."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    match = re.search(r"(\d{4,6})", text)
    if match:
        return match.group(1)
    return text


def _is_plausible_site_id(site_id, site_name=""):
    """Site IDs are usually 5 digits (e.g. 10001). Reject bare counts like 6419."""
    if not site_id or not re.fullmatch(r"\d{4,6}", site_id):
        return False
    if re.fullmatch(r"\d{5}", site_id):
        return True
    if site_name and re.search(r"[A-Za-z]{2,}", site_name):
        return True
    return False

# (app continues unchanged) -- kept full file content
