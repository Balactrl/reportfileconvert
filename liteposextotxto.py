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


def _parse_site_line(text):
    text = str(text).strip()
    if not text or text.lower() == "nan":
        return "", ""

    # SITEID: 10001  /  SITE ID - 10001
    match = re.search(r"site\s*id\s*[:\-]?\s*(\d{4,6})", text, re.IGNORECASE)
    if match:
        site_id = match.group(1)
        name_match = re.search(r"[-–—]\s*(.+)$", text)
        site_name = name_match.group(1).strip() if name_match else ""
        return site_id, site_name

    # 10001.0 - KONDAPURAM  /  10001-KONDAPURAM
    match = re.search(r"(\d{4,6})(?:\.0+)?\s*[-–—_]\s*(.+)", text)
    if match:
        site_id = match.group(1)
        site_name = match.group(2).strip()
        if _is_plausible_site_id(site_id, site_name):
            return site_id, site_name

    # 10001.0 KONDAPURAM
    match = re.search(r"^(\d{4,6})(?:\.0+)?\s+(.+)$", text)
    if match:
        site_id = match.group(1)
        site_name = match.group(2).strip()
        if _is_plausible_site_id(site_id, site_name):
            return site_id, site_name

    return "", ""


def _parse_site_from_row(row_values):
    """Parse site id and name from one Excel header row."""
    if not row_values:
        return "", ""

    cells = [str(v).strip() for v in row_values if str(v).strip().lower() != "nan"]
    if not cells:
        return "", ""

    # Site id in first cell, name in second cell (common Excel layout)
    site_id = _clean_site_id(cells[0])
    if re.fullmatch(r"\d{4,6}", site_id) and len(cells) >= 2:
        second = cells[1].strip().lstrip("-–—_ ").strip()
        if second and not re.fullmatch(r"\d{4,6}", _clean_site_id(second)):
            if _is_plausible_site_id(site_id, second):
                return site_id, second

    # Site id in first cell, dash in second, name in third: 10001 | - | KONDAPURAM
    if re.fullmatch(r"\d{4,6}", site_id) and len(cells) >= 3:
        third = cells[2].strip()
        if third and not re.fullmatch(r"\d{4,6}", _clean_site_id(third)):
            if _is_plausible_site_id(site_id, third):
                return site_id, third

    # Check each individual cell (merged layouts like "10001 - KONDAPURAM")
    for cell in cells:
        sid, sname = _parse_site_line(cell)
        if sid and _is_plausible_site_id(sid, sname):
            return sid, sname

    # Joined row text e.g. "10001 - KONDAPURAM"
    sid, sname = _parse_site_line(" ".join(cells))
    if sid and _is_plausible_site_id(sid, sname):
        return sid, sname

    return "", ""


def _is_site_header_row(row_values):
    joined = " ".join(str(v) for v in row_values).upper()
    skip_keywords = (
        "PHARMACIES",
        "SALES SUMMARY",
        "SALES TRANSACTION",
        "FROM ",
        "TO DATE",
        "GENERATED",
        "BILLTYPE",
        "USERNAME",
        "DATABASE",
        "TOTAL",
        "CARD",
        "CASH",
        "COD",
        "CREDIT",
        "GIFT",
        "UPI",
        "QR CODE",
    )
    if any(keyword in joined for keyword in skip_keywords):
        return False
    site_id, site_name = _parse_site_from_row(row_values)
    return bool(site_id) and _is_plausible_site_id(site_id, site_name)


def _extract_header_info(raw_df):
    company = "APOLLO PHARMACIES LIMITED"
    report_title = ""
    site_id = ""
    site_name = ""
    from_date = ""
    to_date = ""

    billtype_row = _find_table_start(raw_df)
    scan_rows = billtype_row if billtype_row is not None else min(8, len(raw_df))

    for idx in range(scan_rows):
        row_values = [v for v in raw_df.iloc[idx].tolist() if str(v).strip().lower() != "nan"]
        if not row_values:
            continue

        joined = " ".join(str(v).strip() for v in row_values)
        upper = joined.upper()

        if "PHARMACIES" in upper and "LIMITED" in upper:
            company = joined
        elif "SALES SUMMARY REPORT" in upper:
            report_title = joined
        elif "FROM" in upper and " TO " in upper:
            parsed_from, parsed_to = _parse_date_range(joined)
            if parsed_from and parsed_to:
                from_date, to_date = parsed_from, parsed_to
        elif _is_site_header_row(row_values):
            parsed_id, parsed_name = _parse_site_from_row(row_values)
            if parsed_id and _is_plausible_site_id(parsed_id, parsed_name):
                site_id, site_name = parsed_id, parsed_name

    # Second pass: only rows strictly between report title and billtype table
    if not site_id and billtype_row is not None:
        for idx in range(1, billtype_row):
            row_values = [v for v in raw_df.iloc[idx].tolist() if str(v).strip().lower() != "nan"]
            if not row_values:
                continue
            parsed_id, parsed_name = _parse_site_from_row(row_values)
            if parsed_id and _is_plausible_site_id(parsed_id, parsed_name):
                site_id, site_name = parsed_id, parsed_name
                break

    return company, report_title, site_id, site_name, from_date, to_date


def _find_table_start(raw_df):
    for idx in range(len(raw_df)):
        row = [_normalize_col(v) for v in raw_df.iloc[idx].tolist()]
        if "BILLTYPE" in row:
            return idx
    return None


def _find_total_row_index(raw_df, header_row):
    for idx in range(header_row + 1, len(raw_df)):
        first_cell = str(raw_df.iloc[idx, 0]).strip().upper()
        if first_cell in {"TOTAL", "TOTAL."}:
            return idx
    return None


def _row_non_empty_cells(row):
    cells = []
    for value in row:
        text = str(value).strip()
        if text and text.lower() != "nan":
            cells.append((text, value))
    return cells


def _format_value_display(value):
    num = _to_number(value)
    raw = str(value).strip().replace(",", "")
    if re.search(r"\.\d{3}\b", raw):
        return f"{num:,.3f}"
    return format_currency(num)


def _parse_footer_sections(raw_df, header_row):
    sections = {
        "sales": [],
        "healingcard": [],
        "oms": [],
        "ip": [],
        "total_cash_amount": None,
    }
    partner_start = None
    total_row = _find_total_row_index(raw_df, header_row)
    start_idx = (total_row + 1) if total_row is not None else header_row + 1

    current_section = None
    for idx in range(start_idx, len(raw_df)):
        row = raw_df.iloc[idx].tolist()
        cells = _row_non_empty_cells(row)
        if not cells:
            continue

        joined_upper = " ".join(text.upper() for text, _ in cells)

        if "PARTNER PROGRAM" in joined_upper:
            partner_start = idx
            break

        label = cells[0][0]
        label_upper = label.upper()
        value = cells[-1][1] if len(cells) > 1 else None

        if label_upper in {"SALES", "SALES :", "SALES :-", "SALES:-"}:
            current_section = "sales"
            continue
        if "HEALINGCARD" in label_upper:
            current_section = "healingcard"
            continue
        if "OMS COLLECTION" in label_upper:
            current_section = "oms"
            continue
        if label_upper.startswith("IP COLLECTION"):
            current_section = "ip"
            continue

        if "TOTAL CASH AMOUNT" in label_upper and value is not None:
            sections["total_cash_amount"] = value
            current_section = None
            continue

        if value is not None and current_section:
            sections[current_section].append((label, value))

    return sections, partner_start


def _parse_partner_program(raw_df, partner_start):
    if partner_start is None:
        return [], False

    header_idx = None
    for idx in range(partner_start, min(partner_start + 5, len(raw_df))):
        row_norm = [_normalize_col(v) for v in raw_df.iloc[idx].tolist()]
        joined = "".join(row_norm)
        if ("SRNO" in row_norm or "SNO" in joined) and "NAME" in row_norm:
            header_idx = idx
            break

    if header_idx is None:
        return [], False, None

    header_values = [str(v).strip() for v in raw_df.iloc[header_idx].tolist()]
    col_map = {
        "srno": _find_column(header_values, ["SRNO", "SR NO", "SNO", "SLNO"]),
        "name": _find_column(header_values, ["NAME"]),
        "noinv": _find_column(header_values, ["NOINV", "NO INV", "NOINVOICE"]),
        "amount": _find_column(header_values, ["AMOUNT", "AMT"]),
        "avg": _find_column(header_values, ["AVG", "AVERAGE"]),
    }

    has_avg = col_map["avg"] is not None
    name_idx = header_values.index(col_map["name"]) if col_map["name"] else 1
    noinv_idx = header_values.index(col_map["noinv"]) if col_map["noinv"] else 2
    amount_idx = header_values.index(col_map["amount"]) if col_map["amount"] else 3
    avg_idx = header_values.index(col_map["avg"]) if col_map["avg"] else None

    partners = []
    totals = None

    for idx in range(header_idx + 1, len(raw_df)):
        row = raw_df.iloc[idx].tolist()
        cells = _row_non_empty_cells(row)
        if not cells:
            continue

        first_text = cells[0][0].strip().upper()
        if first_text in {"TOTAL", "TOTAL AMOUNT", "TOTAL AMOUNT:"}:
            totals = {
                "noinv": row[noinv_idx] if noinv_idx < len(row) else 0,
                "amount": row[amount_idx] if amount_idx < len(row) else 0,
            }
            break

        name = str(row[name_idx]).strip() if name_idx < len(row) else ""
        if not name or name.lower() == "nan":
            continue

        partner = {
            "name": name,
            "noinv": _to_int(row[noinv_idx] if noinv_idx < len(row) else 0),
            "amount": _to_number(row[amount_idx] if amount_idx < len(row) else 0),
            "avg": _to_number(row[avg_idx] if avg_idx is not None and avg_idx < len(row) else 0),
        }
        partners.append(partner)

    return partners, has_avg, totals


def _read_excel_rows(excel_path):
    return pd.read_excel(excel_path, header=None, dtype=object)


def _read_all_sheets(excel_path):
    sheets = pd.read_excel(excel_path, header=None, dtype=object, sheet_name=None)
    if isinstance(sheets, pd.DataFrame):
        return {"Sheet1": sheets}
    return sheets


def _build_dataframe(raw_df):
    header_row = _find_table_start(raw_df)
    if header_row is None:
        raise ValueError("Could not find BILLTYPE header row in the Excel file.")

    header_values = [str(v).strip() if str(v).strip().lower() != "nan" else "" for v in raw_df.iloc[header_row].tolist()]
    data_df = raw_df.iloc[header_row + 1 :].copy()
    data_df.columns = header_values
    data_df = data_df.dropna(how="all")

    billtype_col = _find_column(data_df.columns, ["BILLTYPE", "BILL TYPE", "BILLTYPE."])
    if not billtype_col:
        raise ValueError("BILLTYPE column not found in Excel file.")

    col_map = {
        "S_NO": _find_column(data_df.columns, ["S_NO", "SNO", "S NO"]),
        "S_AMT": _find_column(data_df.columns, ["S_AMT", "SAMT", "S AMT"]),
        "S_DISC": _find_column(data_df.columns, ["S_DISC", "SDISC", "S DISC"]),
        "S_NET": _find_column(data_df.columns, ["S_NET", "SNET", "S NET"]),
        "R_NO": _find_column(data_df.columns, ["R_NO", "RNO", "R NO"]),
        "R_AMT": _find_column(data_df.columns, ["R_AMT", "RAMT", "R AMT"]),
        "R_DISC": _find_column(data_df.columns, ["R_DISC", "RDISC", "R DISC"]),
        "R_NET": _find_column(data_df.columns, ["R_NET", "RNET", "R NET"]),
    }

    rows = []
    for _, row in data_df.iterrows():
        billtype = str(row[billtype_col]).strip()
        if not billtype or billtype.lower() == "nan":
            continue

        rows.append(
            {
                "BILLTYPE": billtype.upper(),
                "S_NO": _to_int(row[col_map["S_NO"]] if col_map["S_NO"] else 0),
                "S_AMT": _to_number(row[col_map["S_AMT"]] if col_map["S_AMT"] else 0),
                "S_DISC": _to_number(row[col_map["S_DISC"]] if col_map["S_DISC"] else 0),
                "S_NET": _to_number(row[col_map["S_NET"]] if col_map["S_NET"] else 0),
                "R_NO": _to_int(row[col_map["R_NO"]] if col_map["R_NO"] else 0),
                "R_AMT": _to_number(row[col_map["R_AMT"]] if col_map["R_AMT"] else 0),
                "R_DISC": _to_number(row[col_map["R_DISC"]] if col_map["R_DISC"] else 0),
                "R_NET": _to_number(row[col_map["R_NET"]] if col_map["R_NET"] else 0),
            }
        )

        if billtype.upper() == "TOTAL":
            break

    if not rows:
        raise ValueError("No bill type rows found in Excel file.")

    return rows


def _section_value(section_items, label_keywords, default=0):
    for label, value in section_items:
        label_upper = label.upper()
        if any(keyword in label_upper for keyword in label_keywords):
            return value
    return default


def _format_summary_line(label, value, indent=7):
    spaces = " " * indent
    return f"{spaces}{label:<24}: {_format_value_display(value):>12}"


def _append_footer_sections(lines, sections, net_cash_sales):
    sales_items = sections["sales"]
    net_cash = _section_value(sales_items, ["NET CASH SALES"], net_cash_sales)
    paid_in = _section_value(sales_items, ["TOTAL PAID IN", "PAID IN"], 0)
    paid_out = _section_value(sales_items, ["TOTAL PAID OUT", "PAID OUT"], 0)
    total_sales = _section_value(sales_items, ["TOTAL SALES"], net_cash)

    lines.extend(
        [
            "\nSALES :-",
            "",
            _format_summary_line("Net Cash Sales", net_cash),
            _format_summary_line("Total Paid In", paid_in),
            _format_summary_line("Total Paid out", paid_out),
            _format_summary_line("Total Sales", total_sales),
            "",
            "HealingCard Collections:",
        ]
    )

    healing_defaults = [
        ("Cash Collections", 0),
        ("Credit Card Collections", 0),
        ("Total Collection", 0),
    ]
    healing_map = {label.upper(): value for label, value in sections["healingcard"]}
    for label, default in healing_defaults:
        value = healing_map.get(label.upper(), default)
        lines.append(_format_summary_line(label, value, indent=5))

    if sections["oms"]:
        lines.append("")
        lines.append("OMS Collections:")
        for label, value in sections["oms"]:
            lines.append(_format_summary_line(label, value, indent=5))

    if sections["ip"]:
        lines.append("")
        lines.append("IP Collection:")
        for label, value in sections["ip"]:
            lines.append(_format_summary_line(label, value, indent=5))

    total_cash = sections["total_cash_amount"]
    if total_cash is None:
        total_cash = total_sales
    lines.extend(["", f"Total Cash Amount            : {_format_value_display(total_cash)} ", "\n" + "-" * PAGE_WIDTH + "\n"])


def _append_partner_program(lines, partners, has_avg, totals):
    if not partners:
        return

    lines.append("\nPartner Program Summary  :\n")
    if has_avg:
        lines.append(
            " slno| Name                                     |     NoInv        |    Amount    |        Avg   |"
        )
    else:
        lines.append(" slno| Name                                     |     NoInv        |    Amount    |")
    lines.append("-" * PAGE_WIDTH)

    total_noinv = 0
    total_amount = 0.0
    for idx, partner in enumerate(partners, start=1):
        total_noinv += partner["noinv"]
        total_amount += partner["amount"]
        if has_avg:
            lines.append(
                f"{idx:6d} | {partner['name']:<38} | {partner['noinv']:16d} | "
                f"{format_currency(partner['amount']):>12} | {format_currency(partner['avg']):>12} |"
            )
        else:
            lines.append(
                f"{idx:6d} | {partner['name']:<38} | {partner['noinv']:16d} | "
                f"{format_currency(partner['amount']):>12} |"
            )

    lines.append("-" * (PAGE_WIDTH - 50))
    if totals:
        total_noinv = _to_int(totals.get("noinv", total_noinv))
        total_amount = _to_number(totals.get("amount", total_amount))

    lines.append(
        f"      TOTAL AMOUNT:                    {total_noinv:27d} | {format_currency(total_amount):>9} |"
    )
    lines.append("-" * (PAGE_WIDTH - 50))


def excel_to_text(excel_path, output_path=None):
    excel_path = Path(excel_path)
    all_sheets = _read_all_sheets(excel_path)
    first_sheet_name = next(iter(all_sheets))
    raw_df = all_sheets[first_sheet_name]

    company, _, site_id, site_name, from_date, to_date = _extract_header_info(raw_df)
    header_row = _find_table_start(raw_df)
    if header_row is None:
        raise ValueError("Could not find BILLTYPE header row in the Excel file.")
    rows = _build_dataframe(raw_df)

    footer_sections = {
        "sales": [],
        "healingcard": [],
        "oms": [],
        "ip": [],
        "total_cash_amount": None,
    }
    partner_start = None
    partner_df = raw_df

    for sheet_df in all_sheets.values():
        sheet_header = _find_table_start(sheet_df)
        start_row = sheet_header if sheet_header is not None else 0
        sections, partner_idx = _parse_footer_sections(sheet_df, start_row)

        for key in ("sales", "healingcard", "oms", "ip"):
            if sections[key]:
                footer_sections[key] = sections[key]
        if sections["total_cash_amount"] is not None:
            footer_sections["total_cash_amount"] = sections["total_cash_amount"]
        if partner_idx is not None:
            partner_start = partner_idx
            partner_df = sheet_df

    partners, has_avg, partner_totals = _parse_partner_program(partner_df, partner_start)

    now = datetime.now()
    lines = []
    lines.append(f"DATE: {now.strftime('%d/%m/%Y')}".rjust(PAGE_WIDTH))
    lines.append(f"TIME: {now.strftime('%I:%M %p')}".rjust(PAGE_WIDTH))
    lines.append("")
    lines.append(company.center(PAGE_WIDTH))
    if site_id and site_name:
        lines.append(f"{site_id} - {site_name}".center(PAGE_WIDTH))
    elif site_id:
        lines.append(site_id.center(PAGE_WIDTH))
    elif site_name:
        lines.append(site_name.center(PAGE_WIDTH))
    lines.append("")
    lines.append("Sales Transaction Summary Report".center(PAGE_WIDTH))
    if from_date and to_date:
        lines.append(f"From Date : {from_date}    To Date : {to_date}".center(PAGE_WIDTH))
    lines.append("-" * PAGE_WIDTH)

    lines.append(
        "|"
        + " SALES ".center(55)
        + "|"
        + " RETURNS ".center(55)
        + "|"
        + " NET ".center(55)
        + "|"
    )
    lines.append("-" * PAGE_WIDTH)
    lines.append(
        f"{'BILLTYPE':<17} |"
        f"{'NO':>8} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
        f"{'NO':>6} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
        f"{'NO':>6} |"
        f"{'AMT':>12} |"
        f"{'DISC':>12} |"
        f"{'NET':>12} |"
    )
    lines.append("-" * PAGE_WIDTH)

    detail_rows = [r for r in rows if r["BILLTYPE"] != "TOTAL"]
    total_row = next((r for r in rows if r["BILLTYPE"] == "TOTAL"), None)

    tot_sale_count = tot_sale_amt = tot_sale_disc = tot_sale_net = 0
    tot_ret_count = tot_ret_amt = tot_ret_disc = tot_ret_net = 0
    net_cash_sales = 0

    for row in detail_rows:
        net_no = row["S_NO"] + row["R_NO"]
        net_amt = row["S_AMT"] + row["R_AMT"]
        net_disc = row["S_DISC"] + row["R_DISC"]
        net_net = row["S_NET"] + row["R_NET"]

        lines.append(
            f"{row['BILLTYPE']:<17} |"
            f"{row['S_NO']:8d} |"
            f"{format_currency(row['S_AMT']):>12} |"
            f"{format_currency(row['S_DISC']):>12} |"
            f"{format_currency(row['S_NET']):>12} |"
            f"{row['R_NO']:6d} |"
            f"{format_currency(row['R_AMT']):>12} |"
            f"{format_currency(row['R_DISC']):>12} |"
            f"{format_currency(row['R_NET']):>12} |"
            f"{net_no:6d} |"
            f"{format_currency(net_amt):>12} |"
            f"{format_currency(net_disc):>12} |"
            f"{format_currency(net_net):>12} |"
        )

        tot_sale_count += row["S_NO"]
        tot_sale_amt += row["S_AMT"]
        tot_sale_disc += row["S_DISC"]
        tot_sale_net += row["S_NET"]
        tot_ret_count += row["R_NO"]
        tot_ret_amt += row["R_AMT"]
        tot_ret_disc += row["R_DISC"]
        tot_ret_net += row["R_NET"]

        if row["BILLTYPE"] == "CASH":
            net_cash_sales = row["S_NET"] + row["R_NET"]

    lines.append("-" * PAGE_WIDTH)

    if total_row:
        total_net_no = total_row["S_NO"] + total_row["R_NO"]
        total_net_amt = total_row["S_AMT"] + total_row["R_AMT"]
        total_net_disc = total_row["S_DISC"] + total_row["R_DISC"]
        total_net_net = total_row["S_NET"] + total_row["R_NET"]
        lines.append(
            f"{'TOTALAMOUNT   :':<17} |"
            f"{total_row['S_NO']:8d} |"
            f"{format_currency(total_row['S_AMT']):>12} |"
            f"{format_currency(total_row['S_DISC']):>12} |"
            f"{format_currency(total_row['S_NET']):>12} |"
            f"{total_row['R_NO']:6d} |"
            f"{format_currency(total_row['R_AMT']):>12} |"
            f"{format_currency(total_row['R_DISC']):>12} |"
            f"{format_currency(total_row['R_NET']):>12} |"
            f"{total_net_no:6d} |"
            f"{format_currency(total_net_amt):>12} |"
            f"{format_currency(total_net_disc):>12} |"
            f"{format_currency(total_net_net):>12} |"
        )
    else:
        total_net_no = tot_sale_count + tot_ret_count
        total_net_amt = tot_sale_amt + tot_ret_amt
        total_net_disc = tot_sale_disc + tot_ret_disc
        total_net_net = tot_sale_net + tot_ret_net
        lines.append(
            f"{'TOTALAMOUNT   :':<17} |"
            f"{tot_sale_count:8d} |"
            f"{format_currency(tot_sale_amt):>12} |"
            f"{format_currency(tot_sale_disc):>12} |"
            f"{format_currency(tot_sale_net):>12} |"
            f"{tot_ret_count:6d} |"
            f"{format_currency(tot_ret_amt):>12} |"
            f"{format_currency(tot_ret_disc):>12} |"
            f"{format_currency(tot_ret_net):>12} |"
            f"{total_net_no:6d} |"
            f"{format_currency(total_net_amt):>12} |"
            f"{format_currency(total_net_disc):>12} |"
            f"{format_currency(total_net_net):>12} |"
        )

    lines.append("-" * PAGE_WIDTH)
    _append_footer_sections(lines, footer_sections, net_cash_sales)
    _append_partner_program(lines, partners, has_avg, partner_totals)

    report_text = "\n".join(_fix_line(line) for line in lines)

    if output_path is None:
        output_path = excel_path.with_suffix(".txt")
    else:
        output_path = Path(output_path)

    output_path.write_text(report_text, encoding="utf-8")
    return output_path, report_text


def build_text_to_excel():
    st.subheader("Text to Excel Converter")
    st.write("Upload one or more plain text sales reports and generate a single Excel workbook.")

    if "text_upload_id" not in st.session_state:
        st.session_state["text_upload_id"] = 0

    if st.button("Clear", type="secondary"):
        st.session_state["text_upload_id"] += 1
        st.info("Upload cleared. You can re-upload the correct file.")

    uploader_key = f"uploaded_text_files_{st.session_state['text_upload_id']}"

    with st.form("text_to_excel_form"):
        uploaded_files = st.file_uploader(
            "Upload text files", type=["txt"], accept_multiple_files=True, key=uploader_key
        )
        st.caption("The generated file will be saved as Sales_Report.xlsx")

        submitted = st.form_submit_button("Convert to Excel")

    if submitted:
        if not uploaded_files:
            st.warning("Please upload one or more text files before converting.")
            return

        temp_folder = Path(tempfile.gettempdir()) / "sales_converter"
        temp_folder.mkdir(exist_ok=True, parents=True)
        output_name = "Sales_Report.xlsx"
        output_path_full = temp_folder / output_name

        with st.spinner("Extracting sales data and building Excel file..."):
            try:
                data_list = process_files(uploaded_files)
                if not data_list:
                    st.error("No valid sales data could be extracted from the uploaded files.")
                    return

                output_path, df = convert_to_excel(data_list, str(output_path_full))

                with open(output_path, "rb") as file_data:
                    excel_bytes = file_data.read()

                st.success("Excel file generated successfully.")
                st.download_button(
                    label="Download Excel File",
                    data=excel_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                with st.expander("Preview extracted data"):
                    st.dataframe(df)
            except Exception as exc:
                st.error(f"Conversion failed: {exc}")


def build_excel_to_txt():
    st.subheader("Excel to TXT Converter")
    st.write("Upload one or more Excel files to convert them into text report format. All outputs will be saved in a folder.")

    if "excel_upload_id" not in st.session_state:
        st.session_state["excel_upload_id"] = 0

    if st.button("Clear", type="secondary"):
        st.session_state["excel_upload_id"] += 1
        st.info("Upload cleared. You can re-upload the correct files.")

    uploader_key = f"uploaded_excel_files_{st.session_state['excel_upload_id']}"

    with st.form("excel_to_txt_form"):
        uploaded_excel_files = st.file_uploader(
            "Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True, key=uploader_key
        )
        st.caption("Generated TXT files will be available for download below.")

        submitted = st.form_submit_button("Convert Excel to TXT")

    if submitted:
        if not uploaded_excel_files:
            st.warning("Please upload one or more Excel files first.")
            return

        output_files = []

        with st.spinner(f"Converting {len(uploaded_excel_files)} file(s) to TXT..."):
            success_count = 0
            error_messages = []

            progress_bar = st.progress(0)
            for idx, uploaded_excel in enumerate(uploaded_excel_files):
                try:
                    base_name = Path(uploaded_excel.name).stem
                    output_name = f"{base_name}.txt"

                    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_excel.name).suffix) as tmp:
                        tmp.write(uploaded_excel.getvalue())
                        temp_excel_path = Path(tmp.name)

                    saved_path, report_text = excel_to_text(temp_excel_path, output_name)
                    output_files.append((output_name, report_text))
                    success_count += 1

                    try:
                        temp_excel_path.unlink(missing_ok=True)
                    except Exception:
                        pass

                except Exception as exc:
                    error_messages.append(f"{uploaded_excel.name}: {str(exc)}")

                progress_bar.progress((idx + 1) / len(uploaded_excel_files))

        st.success(f"Successfully converted {success_count} file(s)")

        if error_messages:
            with st.expander("View errors"):
                for error in error_messages:
                    st.error(error)

        with st.expander("View output files"):
            for output_name, report_text in output_files:
                st.write(f"✅ {output_name}")
                try:
                    file_bytes = report_text.encode("utf-8")
                    st.download_button(
                        label=f"Download {output_name}",
                        data=file_bytes,
                        file_name=output_name,
                        mime="text/plain",
                    )
                except Exception as exc:
                    st.error(f"Could not prepare download for {output_name}: {exc}")

        with st.expander("Preview first file"):
            if output_files:
                first_name, first_text = output_files[0]
                st.text_area("First TXT Preview", first_text, height=300)


def build_about():
    st.header("About")
    st.write(
        "This app combines two conversion tools:\n"
        "1. Text to Excel Converter\n"
        "2. Excel to TXT Converter\n"
        "\nUse the menu on the left to switch between tools."
    )


def main():
    st.set_page_config(page_title="Sales Converter", page_icon="🔄", layout="wide")

    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(180deg, #eef6fd 0%, #ffffff 100%);
        }
        div[data-testid="stSidebar"] {
            background-color: #f2f7ff;
        }
        .stButton>button {
            background-color: #0f4c81;
            color: white;
            border-radius: 10px;
            border: none;
            padding: 0.7rem 1rem;
        }
        .stButton>button:hover {
            background-color: #0b3a66;
        }
        .stFileUploader, .stTextInput>div>div {
            border-radius: 12px;
        }
        .css-1t8l2tu {
            background-color: #ffffff;
        }
        .css-1cpxqw2 {
            background-color: #f0f5fb;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.title("Sales Converter")
    st.sidebar.markdown(
        "Use the tool menu below to switch between converters and start conversion." 
    )
    option = st.sidebar.radio(
        "Choose a tool",
        ["Text to Excel Converter", "Excel to TXT Converter", "About"],
    )
    st.sidebar.divider()
    st.sidebar.info(
        "Upload files and click Convert. Output filenames are generated automatically."
    )

    st.markdown(
        "<div style='background: linear-gradient(90deg, #0f4c81, #2f6fb2); padding: 16px; border-radius: 16px; color: white;'>"
        "<h1 style='margin: 0; font-size: 2.3rem;'>Sales Report Converter</h1>"
        "<p style='margin: 4px 0 0; color: #d7e6ff;'>Simple text and Excel conversion tools in one app.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    if option == "Text to Excel Converter":
        build_text_to_excel()
    elif option == "Excel to TXT Converter":
        build_excel_to_txt()
    else:
        build_about()


if __name__ == "__main__":
    main()
