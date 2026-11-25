#!/usr/bin/env python3
"""SDS Biocompatibility Screening Tool — ACONIS | ISO 10993-1:2025"""

import streamlit as st
import anthropic
import fitz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.drawing.image import Image as XLImage
import pandas as pd
import json, io, datetime, re, os, traceback, base64

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="SDS Biocompatibility Screener", page_icon="🔬",
                   layout="wide", initial_sidebar_state="expanded")

# ─── File paths ───────────────────────────────────────────────────────────────
DIR         = os.path.dirname(__file__)
DB_FILE     = os.path.join(DIR, "databases.xlsx")
TMPL_FILE   = os.path.join(DIR, "template.xlsx")
PROMPT_FILE = os.path.join(DIR, "agent_prompt_generic.txt")

LOGO_B64 = "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"

# ─── Styles ───────────────────────────────────────────────────────────────────
DATA_FILL = PatternFill("solid", fgColor="C5E1FF")
WRAP_TOP  = Alignment(wrap_text=True, vertical="top")
DATE_FMT  = '[$-409]mmm dd\\, yyyy;@'

def seg(bold=False, italic=False, color="000000", size=10):
    return Font(name="Segoe UI Light", size=size, bold=bold, italic=italic, color=color)

def fr(formula, fill_hex=None, bold=False, italic=False, font_color="000000"):
    fill = PatternFill("solid", fgColor=fill_hex) if fill_hex else None
    font = Font(name="Segoe UI Light", size=10, bold=bold, italic=italic, color=font_color)
    return FormulaRule(formula=[formula], fill=fill, font=font)

def cidx(col): return column_index_from_string(col)

COLS = [
    ("A","sds_id"), ("B","product_name"), ("C","supplier"), ("D","sds_date"),
    ("E","composition"),
    ("F","cytotoxicity"), ("G","sensitisation"), ("H","skin_irritation"),
    ("I","eye_irritation"), ("J","acute_systemic_toxicity"),
    ("K","subacute_subchronic"), ("L","chronic_toxicity"),
    ("M","genotoxicity"), ("N","carcinogenicity"), ("O","reproductive_toxicity"),
    ("P","endocrine_disruption"), ("Q","bioaccumulation"),
    ("R","haemocompatibility"), ("S","pyrogenicity"), ("T","implantation"),
    ("U","immune_responses"), ("V","other_biological_effects"),
    ("W","cmr_regulatory_status"), ("X","azo_dyes"),
    ("Y","formaldehyde"), ("Z","heavy_metals"), ("AA","reach_svhc"),
    ("AC","alert_justification"),
]

ANALYSIS_FIELDS = [
    "cytotoxicity","sensitisation","skin_irritation","eye_irritation",
    "acute_systemic_toxicity","subacute_subchronic","chronic_toxicity",
    "genotoxicity","carcinogenicity","reproductive_toxicity",
    "endocrine_disruption","bioaccumulation","haemocompatibility",
    "pyrogenicity","implantation","immune_responses",
    "other_biological_effects","cmr_regulatory_status","azo_dyes",
    "formaldehyde","heavy_metals","reach_svhc",
]

ALERT_COLORS = {
    "CRITICAL":   ("#CC0000", "#FFFFFF", "🔴"),
    "MAJOR":      ("#E53935", "#FFFFFF", "🔴"),
    "MINOR":      ("#F57C00", "#FFFFFF", "🟠"),
    "NONE":       ("#2E7D32", "#FFFFFF", "🟢"),
    "NOT AN SDS": ("#808080", "#FFFFFF", "⬜"),
    "ERROR":      ("#888888", "#FFFFFF", "⚠️"),
}

DB_DISPLAY_NAMES = [
    ("CLP Annex VI ATP22",          "CLP Annex VI ATP22"),
    ("SVHC Candidates",             "SVHC Candidates (ECHA)"),
    ("REACH Annex XVII",            "REACH Annex XVII"),
    ("Aromatic Amines (Ent.43)",    "Aromatic Amines (Entry 43)"),
    ("Azo - Restricted Amines (JTF)", "Azo Restricted Amines (JTF)"),
    ("IARC",                        "IARC Monographs"),
    ("ED Assessment (ECHA)",        "ED Assessment (ECHA)"),
    ("TEDX",                        "TEDX"),
]

# ─── SDS format detection ─────────────────────────────────────────────────────
def detect_sds_format(text: str) -> str:
    """Detect SDS format — works for EN/FR/DE/ES/IT documents."""
    high_sections = 0
    for i in range(9, 17):
        patterns = [
            rf'section\s+{i}\b',       # EN
            rf'rubrique\s+{i}\b',      # FR
            rf'abschnitt\s+{i}\b',     # DE
            rf'secci[o\u00f3]n\s+{i}\b', # ES
            rf'sezione\s+{i}\b',       # IT
            rf'^\s*{i}[\s\.\:\-\)]', # Generic numbered
        ]
        if any(re.search(p, text, re.IGNORECASE | re.MULTILINE) for p in patterns):
            high_sections += 1
    return "16-section" if high_sections >= 4 else "8-section"

def normalize_db_version(raw: str) -> str:
    """Normalize database version/date to a clean display string."""
    import re as _re, datetime as _dt
    if not raw or str(raw).strip() in ("", "—", "None"):
        return "—"
    s = str(raw).strip()
    # Strip common verbose prefixes
    for prefix in [
        "Last table update:", "Last table update :",
        "Last updated", "Last update:", "Last update :",
        "Database contains",
    ]:
        if s.lower().startswith(prefix.lower()):
            s = s[len(prefix):].strip(" :,")
    # Take only first meaningful sentence
    s = _re.split(r'[.\n]', s)[0].strip()
    # Remove trailing time like "10:08 (CET)"
    s = _re.sub(r'\s+\d{1,2}:\d{2}.*$', '', s).strip()

    months = {"january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
              "july":7,"august":8,"september":9,"october":10,"november":11,"december":12}

    # YYYY-MM-DD
    m = _re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        try:
            d = _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return d.strftime("%B %d, %Y")
        except Exception:
            pass

    # DD Month YYYY  e.g. "10 December 2025"
    m = _re.match(r'(\d{1,2})\s+(\w+)\s+(\d{4})', s, _re.IGNORECASE)
    if m and m.group(2).lower() in months:
        try:
            d = _dt.date(int(m.group(3)), months[m.group(2).lower()], int(m.group(1)))
            return d.strftime("%B %d, %Y")
        except Exception:
            pass

    # Month DD, YYYY  e.g. "October 1, 2015"
    m = _re.match(r'(\w+)\s+(\d{1,2}),?\s+(\d{4})', s, _re.IGNORECASE)
    if m and m.group(1).lower() in months:
        try:
            d = _dt.date(int(m.group(3)), months[m.group(1).lower()], int(m.group(2)))
            return d.strftime("%B %d, %Y")
        except Exception:
            pass

    # Month YYYY  e.g. "September 2018"
    m = _re.match(r'(\w+)\s+(\d{4})$', s, _re.IGNORECASE)
    if m and m.group(1).lower() in months:
        return f"{m.group(1).capitalize()} {m.group(2)}"

    # Looks like a version code (e.g. "ATP 22") — no digits that look like a year
    if not _re.search(r'\b(19|20)\d{2}\b', s):
        return f"Version: {s}"

    return s


MSDS_8_MODIFIER = """

## IMPORTANT — OLD FORMAT MSDS DETECTED (pre-GHS / pre-2015)

This document appears to be an old-format MSDS (pre-GHS harmonisation, typically 8–16 sections
with non-standardised structure). Apply the following rules:

- Section mapping: hazard info may appear in "Section II/III", "Health Hazards", "Chemical Hazards"
  instead of current Sections 2 and 11.
- CLP H-phrases may be absent; old R-phrases may be used — map R-phrases to H-phrases where possible.
- Composition may be in "Section II" or "Ingredients" — extract all CAS numbers found.
- When a specific field CANNOT be found due to the old format, write EXACTLY:
  "Not available in MSDS format — [field name] section absent or non-standardised"
- Do NOT invent, extrapolate, or assume data not explicitly present in the document.
- Still output the complete JSON block with all required keys — use the above phrase for unavailable fields.

"""

# ─── Database helpers ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_db_dates(db_path: str) -> dict:
    """Read version/date from each database sheet (cell A4, fallback first 3 rows)."""
    if not os.path.exists(db_path):
        return {}
    dates = {}
    try:
        wb = load_workbook(db_path, read_only=True, data_only=True)
        for sheet_name, _ in DB_DISPLAY_NAMES:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            val = None
            try:
                val = ws["A4"].value
            except Exception:
                pass
            if not val or str(val).strip() in ("", "None"):
                for row in ws.iter_rows(max_row=3, values_only=True):
                    for cell in row:
                        if cell and str(cell).strip() not in ("", "None"):
                            val = str(cell).strip()
                            break
                    if val:
                        break
            dates[sheet_name] = str(val).strip() if val else "—"
        wb.close()
    except Exception:
        pass
    return dates


@st.cache_data(show_spinner=False)
def load_databases(db_bytes: bytes) -> dict:
    wb = load_workbook(io.BytesIO(db_bytes), read_only=True, data_only=True)
    dbs = {}

    def find_header_row(ws, keyword, max_rows=10):
        for i, row in enumerate(ws.iter_rows(max_row=max_rows, values_only=True), 1):
            if any(keyword.lower() in str(c).lower() for c in row if c):
                return i
        return 1

    def sheet_to_df(name, cas_hint="CAS"):
        if name not in wb.sheetnames:
            return pd.DataFrame()
        ws = wb[name]
        h = find_header_row(ws, cas_hint)
        rows = list(ws.iter_rows(min_row=h, values_only=True))
        if not rows:
            return pd.DataFrame()
        hdrs = [str(c).strip() if c else f"_c{i}" for i, c in enumerate(rows[0])]
        return pd.DataFrame(rows[1:], columns=hdrs)

    def register(key, sheet_name, cas_hint="CAS"):
        try:
            df = sheet_to_df(sheet_name, cas_hint)
            cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
            if cas_col is not None:
                df[cas_col] = df[cas_col].astype(str).str.strip()
                dbs[key]            = df
                dbs[f"{key}_cas"]   = cas_col
                dbs[f"{key}_sheet"] = sheet_name
        except Exception:
            pass

    register("clp",     "CLP Annex VI ATP22")
    register("svhc",    "SVHC Candidates")
    register("reach",   "REACH Annex XVII")
    register("amines",  "Aromatic Amines (Ent.43)")
    register("azo_jtf", "Azo - Restricted Amines (JTF)")
    register("iarc",    "IARC")
    register("ed",      "ED Assessment (ECHA)")
    register("tedx",    "TEDX")
    wb.close()
    return dbs


def lookup_all_cas(cas_list, dbs):
    if not cas_list:
        return "No CAS numbers extracted from the SDS text."
    today = datetime.date.today().strftime("%b %Y")
    lines = [f"PRE-COMPUTED DATABASE LOOKUP — {len(cas_list)} CAS extracted", "="*60]
    for cas in cas_list:
        hits = []
        if "clp" in dbs:
            cc = dbs["clp_cas"]
            m  = dbs["clp"][dbs["clp"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                codes = str(r.get("Hazard Statement Code(s)","")).replace("\n"," ")
                cat   = str(r.get("Hazard Class and Category Code(s)",""))
                name  = r.get("International Chemical Identification","")
                hits.append(f"  [CLP Annex VI ATP22] {name} | {cat} | H-codes: {codes}")
        if "svhc" in dbs:
            cc = dbs["svhc_cas"]
            m  = dbs["svhc"][dbs["svhc"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                reason = r.get("Reason for inclusion", r.get("reason",""))
                name   = r.get("Substance name", r.get("name",""))
                hits.append(f"  [SVHC Candidates] {name} — Reason: {reason}")
        if "iarc" in dbs:
            cc = dbs["iarc_cas"]
            m  = dbs["iarc"][dbs["iarc"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                grp   = r.get("Group", r.get("group",""))
                agent = r.get("Agent", r.get("agent",""))
                hits.append(f"  [IARC, {today}] Group {grp} — {agent}")
        if "reach" in dbs:
            cc = dbs["reach_cas"]
            m  = dbs["reach"][dbs["reach"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                entry = r.get("Entry", r.get("entry",""))
                restr = str(r.get("Restriction", r.get("restriction","")))[:120]
                hits.append(f"  [REACH Annex XVII] Entry {entry} — {restr}")
        if "amines" in dbs:
            cc = dbs["amines_cas"]
            m  = dbs["amines"][dbs["amines"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append("  [Aromatic Amines Ent.43] MATCH — restricted aromatic amine")
        if "azo_jtf" in dbs:
            cc = dbs["azo_jtf_cas"]
            m  = dbs["azo_jtf"][dbs["azo_jtf"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append("  [Azo - Restricted Amines JTF] MATCH — restricted azo amine")
        if "ed" in dbs:
            cc = dbs["ed_cas"]
            m  = dbs["ed"][dbs["ed"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                status = str(r.get("Status", r.get("status","")))
                hits.append(f"  [ED Assessment (ECHA), {today}] {status}")
        if "tedx" in dbs:
            cc = dbs["tedx_cas"]
            m  = dbs["tedx"][dbs["tedx"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append("  [TEDX] Listed as endocrine disruptor")
        if hits:
            lines.append(f"\nCAS {cas}:")
            lines.extend(hits)
        else:
            lines.append(f"\nCAS {cas}: Not found in CLP ATP22, SVHC, IARC, REACH Annex XVII, "
                         "Aromatic Amines, ED Assessment, TEDX")
    return "\n".join(lines)


# ─── Excel helpers ────────────────────────────────────────────────────────────
def parse_date(s):
    if not s or str(s).strip() in ("", "Not stated", "N/A", "not stated", "—"):
        return str(s) if s else "Not stated"
    if isinstance(s, (datetime.datetime, datetime.date)):
        return datetime.datetime(s.year, s.month, s.day)
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y",
                "%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y"):
        try:
            return datetime.datetime.strptime(str(s).strip(), fmt)
        except Exception:
            pass
    return str(s)


def compute_alert_level(data):
    all_text = " ".join(str(data.get(f,"")) for f in ANALYSIS_FIELDS)
    if "🔴" in all_text: return "CRITICAL"
    if "🟠" in all_text: return "MAJOR"
    if "🟡" in all_text: return "MINOR"
    return "NONE"


def copy_row_format(ws, source_row, target_row):
    """Repeat the first data-row formatting on subsequent table rows."""
    from copy import copy
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def enforce_report_formatting(wb):
    """Apply Segoe UI Light and top-left wrapping to all report data tables."""
    from copy import copy
    from openpyxl.styles import Alignment
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None or cell.has_style:
                    f = copy(cell.font)
                    f.name = "Segoe UI Light"
                    cell.font = f
        if ws.title == "Project":
            ranges = ((17, 24, 1, 2),)
        elif ws.title == "SDS Dashboard":
            ranges = ((4, max(4, ws.max_row), 1, 10),)
        elif ws.title == "SDS Analysis":
            ranges = ((5, max(5, ws.max_row), 1, 34),)
        else:
            ranges = ()
        for r1, r2, c1, c2 in ranges:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = ws.cell(r, c)
                    a = copy(cell.alignment)
                    a.horizontal = "left"
                    a.vertical = "top"
                    a.wrap_text = True
                    cell.alignment = a


def add_report_logos(wb):
    """Embed the ACONIS logo as a real Excel image on each report sheet."""
    logo_bytes = base64.b64decode(LOGO_B64)
    for sheet_name in ("Project", "SDS Dashboard", "SDS Analysis"):
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else None
        if ws is None:
            continue
        ws["A1"].value = None
        ws._images = []
        img = XLImage(io.BytesIO(logo_bytes))
        img.width = 125
        img.height = 38
        img.anchor = "A1"
        ws.add_image(img)


def fit_row_height(ws, row, first_col, last_col, min_height=22, max_height=420):
    """Estimate a readable row height from wrapped cell contents."""
    lines = 1
    for col in range(first_col, last_col + 1):
        value = ws.cell(row, col).value
        if value in (None, ""):
            continue
        width = ws.column_dimensions[ws.cell(row, col).column_letter].width or 10
        chars_per_line = max(12, int(width * 1.05))
        text = str(value)
        for paragraph in text.split("\n"):
            lines = max(lines, max(1, len(paragraph) // chars_per_line + 1))
    ws.row_dimensions[row].height = min(max_height, max(min_height, lines * 14.5))


def apply_table_font_and_alignment(wb):
    """Set table font size 10, titles size 11, top-left wrapped data."""
    from copy import copy
    thin = Side(style="thin", color="8FAADC")
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None or cell.has_style:
                    f = copy(cell.font)
                    f.name = "Segoe UI Light"
                    if ws.title in ("SDS Dashboard", "SDS Analysis") and cell.row >= (4 if ws.title == "SDS Dashboard" else 4):
                        f.sz = 10
                    cell.font = f
        if ws.title == "Project":
            for r in range(17, 25):
                for c in range(1, 3):
                    cell = ws.cell(r, c)
                    a = copy(cell.alignment); a.horizontal = "left"; a.vertical = "top"; a.wrap_text = True
                    cell.alignment = a
                    cell.border = table_border
                    f = copy(cell.font); f.name = "Segoe UI Light"; f.sz = 10; cell.font = f
                fit_row_height(ws, r, 1, 2)
        elif ws.title == "SDS Dashboard":
            for r in range(4, ws.max_row + 1):
                for c in range(1, 11):
                    cell = ws.cell(r, c)
                    a = copy(cell.alignment); a.horizontal = "left"; a.vertical = "top"; a.wrap_text = True
                    cell.alignment = a
                    cell.border = table_border
                    f = copy(cell.font); f.name = "Segoe UI Light"; f.sz = 10; cell.font = f
                fit_row_height(ws, r, 1, 10)
        elif ws.title == "SDS Analysis":
            for r in range(5, ws.max_row + 1):
                for c in range(1, 35):
                    cell = ws.cell(r, c)
                    a = copy(cell.alignment); a.horizontal = "left"; a.vertical = "top"; a.wrap_text = True
                    cell.alignment = a
                    cell.border = table_border
                    f = copy(cell.font); f.name = "Segoe UI Light"; f.sz = 10; cell.font = f
                fit_row_height(ws, r, 1, 34)
        if ws.title in ("Project", "SDS Dashboard", "SDS Analysis"):
            for coord in ("B1",):
                f = copy(ws[coord].font); f.name = "Segoe UI Light"; f.sz = 11; ws[coord].font = f
        if ws.title == "SDS Analysis":
            # Restore the readable layout for the long SDS analysis fields.
            # The template defaults many of these columns to width 13.
            for col_idx in range(5, 28):  # E:AA
                ws.column_dimensions[get_column_letter(col_idx)].width = 36

            # AB contains CRITICAL / MAJOR / MINOR / NONE.
            ws.column_dimensions["AB"].width = 18
            ws.column_dimensions["AF"].width = max(ws.column_dimensions["AF"].width or 0, 15)
            ws.column_dimensions["AG"].width = max(ws.column_dimensions["AG"].width or 0, 24)
            ws.column_dimensions["AH"].width = max(ws.column_dimensions["AH"].width or 0, 14)
        elif ws.title == "SDS Dashboard":
            ws.column_dimensions["E"].width = max(ws.column_dimensions["E"].width or 0, 15)
            ws.column_dimensions["H"].width = max(ws.column_dimensions["H"].width or 0, 15)
            ws.column_dimensions["I"].width = max(ws.column_dimensions["I"].width or 0, 24)
            ws.column_dimensions["J"].width = max(ws.column_dimensions["J"].width or 0, 14)


def apply_cf(ws):
    if hasattr(ws.conditional_formatting, "_cf_rules"):
        ws.conditional_formatting._cf_rules.clear()
    ws.conditional_formatting.add("AB5:AB9999", fr('$AB5="CRITICAL"',   "FF0000", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AB5:AB9999", fr('$AB5="MAJOR"',      "E53935", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AB5:AB9999", fr('$AB5="MINOR"',      "F57C00", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AB5:AB9999", fr('$AB5="NONE"',       "2E7D32", False, False, "FFFFFF"))
    ws.conditional_formatting.add("AB5:AB9999", fr('$AB5="NOT AN SDS"', "808080", False, False, "FFFFFF"))
    ws.conditional_formatting.add("A5:AH9999",  fr('$AB5="CRITICAL"', "FFE6E6"))
    ws.conditional_formatting.add("A5:AH9999",  fr('$AB5="MAJOR"',    "FFE5E5"))
    ws.conditional_formatting.add("A5:AH9999",  fr('$AB5="MINOR"',    "FFF3E0"))
    ws.conditional_formatting.add("AG5:AG9999", fr('$AG5="AI - SDS Screening agent only"', "FFF2CC", False, True))
    ws.conditional_formatting.add("AG5:AG9999", fr('AND($AG5<>"",$AG5<>"AI - SDS Screening agent only")', "E2EFDA"))
    ws.conditional_formatting.add("AD5:AD9999", fr('$AD5<>""', "E2EFDA", True))
    ws.conditional_formatting.add("AH5:AH9999", fr('$AH5="Draft"',      "BDD7EE"))
    ws.conditional_formatting.add("AH5:AH9999", fr('$AH5="Validated"',  "00B050", True, False, "FFFFFF"))
    ws.conditional_formatting.add("AH5:AH9999", fr('$AH5="Superseded"', "D9D9D9", False, False, "808080"))
    ws.conditional_formatting.add("AH5:AH9999", fr('$AH5="Archived"',   "595959", False, False, "FFFFFF"))
    ws.conditional_formatting.add("D5:D9999", fr('AND(D5<>"",ISNUMBER(D5),D5<TODAY()-1095)', "FFE0E0", False, False, "CC0000"))
    ws.conditional_formatting.add("E5:AB9999", FormulaRule(
        formula=['NOT(ISERROR(SEARCH("TO VERIFY",E5)))'],
        fill=PatternFill("solid", fgColor="E2CCFF"),
        font=Font(name="Segoe UI Light", size=10),
    ))


def write_row(wb, data, row):
    ws = wb["SDS Analysis"]
    if row > 5:
        copy_row_format(ws, 5, row)
    for col_letter, key in COLS:
        val = data.get(key,"")
        if key == "sds_date":
            val = parse_date(str(val)) if val else "Not stated"
        c = ws.cell(row=row, column=cidx(col_letter), value=val)
        # Keep the user's template font, fill, borders and alignment.
        # Only enforce the Excel date format for the SDS date column.
        if key == "sds_date":
            c.number_format = DATE_FMT

    ab = ws.cell(row=row, column=cidx("AB"))
    # Write the calculated value directly, rather than relying on Excel
    # recalculation or on the user's Excel language settings.
    alert_level = compute_alert_level(data)
    ab.value = alert_level
    alert_style = {
        "CRITICAL": ("FF0000", "FFFFFF"),
        "MAJOR": ("E53935", "FFFFFF"),
        "MINOR": ("F57C00", "FFFFFF"),
        "NONE": ("2E7D32", "FFFFFF"),
        "NOT AN SDS": ("808080", "FFFFFF"),
    }.get(alert_level)
    if alert_style:
        fill_hex, font_hex = alert_style
        ab.fill = PatternFill("solid", fgColor=fill_hex)
        ab.font = Font(name="Segoe UI Light", size=10, bold=True, color=font_hex)
        ab.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    ws.cell(row=row, column=cidx("AD"), value="")
    ws.cell(row=row, column=cidx("AE"), value="")

    today = datetime.date.today()
    af = ws.cell(row=row, column=cidx("AF"),
                 value=datetime.datetime(today.year, today.month, today.day))
    af.number_format = DATE_FMT

    ws.cell(row=row, column=cidx("AG"), value="AI - SDS Screening agent only")
    ws.cell(row=row, column=cidx("AH"), value="Draft")


def extend_dashboard(wb, analysis_row, dashboard_row):
    if "SDS Dashboard" not in wb.sheetnames:
        return
    ws = wb["SDS Dashboard"]
    if dashboard_row > 4:
        copy_row_format(ws, 4, dashboard_row)
    formulas = [
        ("A", f"=ROW()-3"),
        ("B", f"='SDS Analysis'!$A{analysis_row}"),
        ("C", f"='SDS Analysis'!$B{analysis_row}"),
        ("D", f"='SDS Analysis'!$C{analysis_row}"),
        ("E", f"='SDS Analysis'!$D{analysis_row}"),
        ("F", f"=IF('SDS Analysis'!$AE{analysis_row}<>\"\",'SDS Analysis'!$AE{analysis_row},'SDS Analysis'!$AB{analysis_row})"),
        ("G", f"='SDS Analysis'!$AC{analysis_row}&IF('SDS Analysis'!$AD{analysis_row}<>\"\",CHAR(10)&CHAR(10)&\"Toxicologist comment: \"&'SDS Analysis'!$AD{analysis_row},\"\")"),
        ("H", f"='SDS Analysis'!$AF{analysis_row}"),
        ("I", f"='SDS Analysis'!$AG{analysis_row}"),
        ("J", f"='SDS Analysis'!$AH{analysis_row}"),
    ]
    for col_l, val in formulas:
        c = ws.cell(row=dashboard_row, column=cidx(col_l), value=val)
        # Keep the user's Dashboard formatting and table style.
        if col_l in ("E", "H"):
            c.number_format = DATE_FMT


def update_project_tab(wb, db_path):
    """Preserve the user's Project layout and update only the database list."""
    if "Project" not in wb.sheetnames:
        return
    ws = wb["Project"]
    db_dates = get_db_dates(db_path)
    loaded_sheet_names = set(db_dates.keys())

    # The user's template has the database table header on row 16 and data rows 17:24.
    for row in range(17, 25):
        ws.cell(row=row, column=1).value = None
        ws.cell(row=row, column=2).value = None

    output_row = 17
    for sheet_name, display_name in DB_DISPLAY_NAMES:
        if sheet_name not in loaded_sheet_names:
            continue
        ws.cell(row=output_row, column=1).value = display_name
        ws.cell(row=output_row, column=2).value = normalize_db_version(db_dates.get(sheet_name, "—"))
        output_row += 1
        if output_row > 24:
            break


# ─── Prompt builder ───────────────────────────────────────────────────────────
def build_system_prompt(base_prompt, config, sds_format="16-section"):
    azo_instr = (
        "Col X (Azo dyes / Aromatic amines): assess only from evidence in the SDS, "
        "the declared composition and the Aromatic Amines regulatory database. "
        "Do not infer relevance from device category, textile status, dyeing-agent "
        "or fiber-related criteria. If no relevant substance is identified, write "
        "exactly: `N/A - no relevant azo dye or aromatic amine identified in the "
        "SDS or database checks [Agent assessment]`."
    )

    prompt = base_prompt

    prompt = prompt.replace(
        "<<CLIENT_NAME>>",
        config.get("client_name", "CLIENT")
    )
    prompt = prompt.replace(
        "<<DEVICE_DESCRIPTION>>",
        config.get("device_description", "medical device")
    )
    prompt = prompt.replace(
        "<<CONTACT_SCENARIO>>",
        "General screening - contact scenario not specified"
    )
    prompt = prompt.replace(
        "<<AZO_DYES_INSTRUCTION>>",
        azo_instr
    )

    if sds_format == "8-section":
        prompt += MSDS_8_MODIFIER

    json_schema = """\n\n---\n\n## PYTHON APP — MANDATORY JSON OUTPUT\n
Return ONLY one valid JSON object, with no screening report, no Step 0A, no database freshness commentary, no Markdown fences, and no text before or after the JSON.
Use EXACTLY these keys. Values must include all markers (🔴🟠🟡), source tags, newlines as \\n.\n
```json
{
  "sds_id": "", "product_name": "", "supplier": "", "sds_date": "", "composition": "",
  "cytotoxicity": "", "sensitisation": "", "skin_irritation": "", "eye_irritation": "",
  "acute_systemic_toxicity": "", "subacute_subchronic": "", "chronic_toxicity": "",
  "genotoxicity": "", "carcinogenicity": "", "reproductive_toxicity": "",
  "endocrine_disruption": "", "bioaccumulation": "", "haemocompatibility": "",
  "pyrogenicity": "", "implantation": "", "immune_responses": "",
  "other_biological_effects": "", "cmr_regulatory_status": "", "azo_dyes": "",
  "formaldehyde": "", "heavy_metals": "", "reach_svhc": "", "alert_justification": ""
}
```
"""
    return prompt + json_schema


def extract_json_response(raw: str) -> dict:
    """Extract the first valid JSON object from Claude's text response."""
    if not raw:
        return {}
    candidates = []
    fenced = re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", raw, re.DOTALL | re.IGNORECASE)
    candidates.extend(fenced)
    candidates.append(raw.strip())
    # Also try every plausible object beginning in the response.
    for start, char in enumerate(raw):
        if char == "{":
            candidates.append(raw[start:])
    for candidate in candidates:
        try:
            value = json.loads(candidate)
            if isinstance(value, dict):
                return value
        except Exception:
            continue
    return {}


# ─── Email ─────────────────────────────────────────────────────────────────────

# ── Notion registration helpers ────────────────────────────────────────────────
def _notion_register(first_name: str, last_name: str, company: str, email: str):
    """Create one Notion row for this trial and return its page ID."""
    import requests as _req
    from datetime import datetime, timezone
    token = st.secrets.get("NOTION_TOKEN", "")
    db_id = st.secrets.get("NOTION_DB_ID", "")
    if not token or not db_id:
        return True, None
    admin = st.secrets.get("ADMIN_EMAIL", "elodie.saudrais@aconis.fr")
    role  = "Admin" if email.strip().lower() == admin.strip().lower() else "Beta tester"
    payload = {
        "parent": {"database_id": db_id},
        "properties": {
            "Name":          {"title":     [{"text": {"content": first_name + " " + last_name}}]},
            "First Name":    {"rich_text": [{"text": {"content": first_name}}]},
            "Last Name":     {"rich_text": [{"text": {"content": last_name}}]},
            "Company":       {"rich_text": [{"text": {"content": company}}]},
            "Email":         {"email": email},
            "Registered at": {"date": {"start": datetime.now(timezone.utc).isoformat()}},
            "Role":          {"select": {"name": role}},
        },
    }
    try:
        r = _req.post(
            "https://api.notion.com/v1/pages",
            headers={"Authorization": "Bearer " + token,
                     "Content-Type": "application/json",
                     "Notion-Version": "2022-06-28"},
            json=payload, timeout=10,
        )
        if r.status_code in (200, 201):
            return True, r.json().get("id")
        return False, None
    except Exception:
        return True, None


def _notion_submit_feedback(email: str, company: str, rating: int, comments: str):
    """Update the registration row for this trial with its feedback."""
    import requests as _req
    from datetime import datetime, timezone

    token = st.secrets.get("NOTION_TOKEN", "")
    db_id = st.secrets.get("NOTION_FEEDBACK_DB_ID", "") or st.secrets.get("NOTION_DB_ID", "")
    if not token or not db_id:
        return False, "Feedback is not configured yet. Please ask the administrator to set NOTION_TOKEN and NOTION_DB_ID."

    headers = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }
    try:
        schema_response = _req.get(
            f"https://api.notion.com/v1/databases/{db_id}",
            headers=headers, timeout=10,
        )
        if schema_response.status_code != 200:
            return False, "The SDS Screener Notion database could not be accessed. Check NOTION_DB_ID and Notion permissions."
        properties = schema_response.json().get("properties", {})

        required = ["Name", "Email", "Rating", "Submitted at"]
        missing = [name for name in required if name not in properties]
        if missing:
            return False, "Missing Notion feedback properties: " + ", ".join(missing)

        def rich_text(value):
            return [{"type": "text", "text": {"content": str(value)[:2000]}}]

        def property_value(name, value):
            prop_type = properties[name].get("type")
            if prop_type == "title":
                return {"title": rich_text(value)}
            if prop_type == "rich_text":
                return {"rich_text": rich_text(value)}
            if prop_type == "email":
                return {"email": value or None}
            if prop_type == "number":
                return {"number": int(value) if value not in (None, "") else None}
            if prop_type == "select":
                return {"select": {"name": str(value)}}
            if prop_type == "date":
                return {"date": {"start": value}}
            return None

        # The registration page ID is stored in session state, so feedback is
        # attached to the exact trial that produced it. For an older session,
        # recover the most recently created row for the same email.
        page_id = st.session_state.get("notion_page_id", "")
        if not page_id:
            query = {
                "filter": {"property": "Email", "email": {"equals": email}},
                "sorts": [{"timestamp": "created_time", "direction": "descending"}],
                "page_size": 1,
            }
            q = _req.post(
                f"https://api.notion.com/v1/databases/{db_id}/query",
                headers=headers, json=query, timeout=10,
            )
            if q.status_code == 200:
                rows = q.json().get("results", [])
                if rows:
                    page_id = rows[0].get("id", "")

        now = datetime.now(timezone.utc).isoformat()
        values = {
            "Feedback": "Beta tester feedback",
            "Rating": rating,
            "Comments": comments.strip(),
            "Submitted at": now,
            "App version": "v2.0",
        }
        payload_properties = {}
        for name, value in values.items():
            if name in properties:
                encoded = property_value(name, value)
                if encoded is not None:
                    payload_properties[name] = encoded

        if page_id:
            response = _req.patch(
                f"https://api.notion.com/v1/pages/{page_id}",
                headers=headers,
                json={"properties": payload_properties},
                timeout=10,
            )
            if response.status_code in (200, 201):
                return True, "Thank you for your feedback."
            return False, "Feedback could not be saved in the trial row in Notion. Please try again later."

        # Fallback only if the registration row cannot be found. This still
        # creates one complete trial row, rather than a feedback-only row.
        fallback_values = {
            "Name": f"SDS Screener trial - {company or email} - {now[:10]}",
            "Email": email,
            "Company": company,
            "Role": "Beta tester",
            **values,
        }
        create_properties = {}
        for name, value in fallback_values.items():
            if name in properties:
                encoded = property_value(name, value)
                if encoded is not None:
                    create_properties[name] = encoded
        response = _req.post(
            "https://api.notion.com/v1/pages",
            headers=headers,
            json={"parent": {"database_id": db_id}, "properties": create_properties},
            timeout=10,
        )
        if response.status_code in (200, 201):
            return True, "Thank you for your feedback."
        return False, "Feedback could not be saved in Notion. Please try again later."
    except Exception:
        return False, "Feedback could not be saved in Notion. The analysis and Excel download remain available."

def _show_registration_gate():
    st.markdown(
        "<h1 style='color:#007AFF;font-weight:700;margin:0;padding-top:4px;'>"
        "Biocompatibility SDS Screener</h1>"
        "<p style='color:#007AFF;font-size:0.82em;margin:2px 0 0 0;'>"
        "ACONIS — Automated SDS analysis for ISO 10993-1:2025 / MDR 2017/745 compliance</p>",
        unsafe_allow_html=True,
    )
    st.divider()
    st.markdown(
        "<div style='background:#F0F6FF;border-left:4px solid #007AFF;border-radius:8px;"
        "padding:22px 26px;margin:10px 0 28px 0;max-width:600px;'>"
        "<p style='margin:0 0 8px 0;font-size:1.05em;color:#007AFF;font-weight:600;'>"
        "🧪 Beta access</p>"
        "<p style='margin:0;color:#333;line-height:1.6;'>"
        "Please fill in your details to access the application. "
        "Your information will be stored securely and used only to manage beta access."
        "</p></div>",
        unsafe_allow_html=True,
    )
    with st.form("registration_form"):
        c1, c2 = st.columns(2)
        with c1:
            first_name = st.text_input("First name *", placeholder="Marie")
        with c2:
            last_name  = st.text_input("Last name *", placeholder="Dupont")
        company  = st.text_input("Company *", placeholder="Acme Medical")
        email    = st.text_input("Professional email *", placeholder="marie.dupont@acme.com")
        submitted = st.form_submit_button("Request access →", type="primary", use_container_width=True)
        if submitted:
            if not all([first_name.strip(), last_name.strip(), company.strip(), email.strip()]):
                st.error("Please fill in all fields.")
            elif "@" not in email or "." not in email.split("@")[-1]:
                st.error("Please enter a valid email address.")
            else:
                with st.spinner("Registering…"):
                    registration_ok, notion_page_id = _notion_register(
                        first_name.strip(), last_name.strip(),
                        company.strip(), email.strip()
                    )
                st.session_state["notion_page_id"] = notion_page_id or ""
                st.session_state["registered"]     = True
                st.session_state["reg_first_name"] = first_name.strip()
                st.session_state["reg_last_name"]  = last_name.strip()
                st.session_state["reg_company"]    = company.strip()
                st.session_state["reg_email"]      = email.strip()
                st.rerun()


# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI Light', 'Helvetica Neue', sans-serif; }
[data-testid="stSidebar"] { background-color: #F0F6FF; }
[data-testid="stSidebar"] img {
    mix-blend-mode: normal !important;
    display: block !important;
}
.main .block-container { padding-top: 0.6rem !important; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color: #007AFF; }
h1 { color: #007AFF !important; font-weight: 300; }
h2, h3 { color: #007AFF; font-weight: 300; }
.stButton > button[kind="primary"],
.stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button,
button[data-testid="baseButton-primary"] {
    background-color: #007AFF !important;
    color: #FFFFFF !important;
    border: 1px solid #007AFF !important;
    font-family: 'Segoe UI Light', sans-serif;
    letter-spacing: 0.03em;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button:hover,
[data-testid="stFormSubmitButton"] > button:hover,
button[data-testid="baseButton-primary"]:hover {
    background-color: #005FCC !important;
    border-color: #005FCC !important;
    color: #FFFFFF !important;
}
.footer-aconis {
    margin-top: 40px; padding: 18px 0 8px 0;
    border-top: 1px solid #DCE8FF; color: #888;
    font-size: 0.78em; text-align: center; line-height: 1.7;
}
</style>
""", unsafe_allow_html=True)

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_bytes = base64.b64decode(LOGO_B64)
    st.image(logo_bytes, width=200)
    st.markdown("---")
    st.caption("v2.0 — Claude Sonnet 5.5 — ISO 10993-1:2025")
    st.markdown("---")
    st.markdown(
        '<a href="mailto:elodie.saudrais@aconis.fr?subject=SDS%20Screener%20-%20Contact" '
        'style="display:block;text-align:center;background:#007AFF;color:#fff;'
        'padding:8px 0;border-radius:6px;text-decoration:none;font-size:0.9em;">'
        '✉️ Contacter ACONIS</a>',
        unsafe_allow_html=True,
    )
    st.markdown("---")
    st.markdown(
        "<p style=\'font-size:0.85em;color:#007AFF;font-weight:600;margin:0 0 6px 0;\'>"
        "About this application</p>"
        "<p style=\'font-size:0.80em;color:#444;line-height:1.65;margin:0 0 12px 0;\'>"
        "Supports the biocompatibility assessment of medical devices "
        "(ISO&nbsp;10993-1:2025 / MDR&nbsp;2017/745). "
        "Automatically <b>extracts and reports biocompatibility data</b> across "
        "17&nbsp;endpoints and <b>cross-references each compound against "
        "8 regulatory databases</b>.</p>"
        "<p style=\'font-size:0.85em;color:#007AFF;font-weight:600;margin:0 0 6px 0;\'>"
        "How it works</p>"
        "<div style=\'font-size:0.80em;color:#444;line-height:1.85;\'>"
        "📄 SDS PDF upload<br>"
        "↓ 🔢 Text &amp; CAS extraction<br>"
        "↓ 🔍 8 regulatory databases<br>"
        "↓ 🤖 AI-assisted synthesis<br>"
        "↓ 📊 Excel report download"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    db_dates_sb = get_db_dates(DB_FILE) if os.path.exists(DB_FILE) else {}
    db_rows_sb  = "".join(
        "<tr style=\'border-bottom:1px solid #DCE8FF;\'>"
        "<td style=\'padding:5px 8px;color:#222;font-size:0.78em;font-weight:500;\'>"
        + disp +
        "</td><td style=\'padding:5px 8px;color:#555;font-size:0.75em;\'>"
        + normalize_db_version(db_dates_sb.get(sheet, "—")) +
        "</td></tr>"
        for sheet, disp in DB_DISPLAY_NAMES
    )
    st.markdown(
        "<p style=\'font-size:0.85em;color:#007AFF;font-weight:600;margin:0 0 4px 0;\'>"
        "Regulatory databases</p>"
        "<p style=\'font-size:0.72em;color:#666;line-height:1.45;margin:0 0 8px 0;\'>"
        "Database extraction date: <b>4 May 2026</b><br>"
        "Some source databases have older update dates; their contents were checked in May 2026."
        "</p>"
        "<table style=\'border-collapse:collapse;width:100%;border:1px solid #DCE8FF;\'>"
        "<tr style=\'background:#007AFF;color:#fff;\'>"
        "<th style=\'padding:5px 8px;font-size:0.78em;font-weight:500;\'>Database</th>"
        "<th style=\'padding:5px 8px;font-size:0.78em;font-weight:500;\'>Version</th>"
        "</tr>"
        + db_rows_sb +
        "</table>",
        unsafe_allow_html=True,
    )

api_key     = st.secrets.get("ANTHROPIC_API_KEY","")
ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL","elodie.saudrais@aconis.fr")
BETA_LIMIT  = 5

# ─── Registration gate ─────────────────────────────────────────────────────────
if not st.session_state.get("registered", False):
    _show_registration_gate()
    st.stop()

# ─── Main area ────────────────────────────────────────────────────────────────
st.markdown(
    "<h1 style='color:#007AFF;font-weight:700;margin:0 0 4px 0;'>"
    "Biocompatibility SDS Screener</h1>"
    "<p style='color:#007AFF;font-size:0.82em;margin:2px 0 0 0;font-weight:400;'>"
    "ACONIS — Automated SDS analysis for ISO 10993-1:2025 / MDR 2017/745 compliance</p>",
    unsafe_allow_html=True,
)

# ── 1. Presentation ───────────────────────────────────────────────────────────
db_dates = get_db_dates(DB_FILE) if os.path.exists(DB_FILE) else {}

# ── Client info from registration ─────────────────────────────────────────────
client_name  = st.session_state.get("reg_company", "")
client_email = st.session_state.get("reg_email", "")
config = {"client_name": client_name, "client_email": client_email}

# Keep the last completed report available across Streamlit reruns, including
# download and feedback submissions.
if "last_results" not in st.session_state:
    st.session_state["last_results"] = []
if "last_excel_bytes" not in st.session_state:
    st.session_state["last_excel_bytes"] = None
if "last_excel_name" not in st.session_state:
    st.session_state["last_excel_name"] = None

def _section_title(text):
    st.markdown(
        f"<h2 style='color:#007AFF;font-weight:700;margin:24px 0 4px 0;'>{text}</h2>",
        unsafe_allow_html=True,
    )

def _build_results_table(results):
    thead = (
        '<table style="width:100%;border-collapse:collapse;'
        'border:1px solid #DCE8FF;font-family:Segoe UI Light,Arial,sans-serif;">'
        '<thead><tr style="background:#007AFF;color:#fff;text-align:left;">'
        '<th style="padding:10px 8px;width:36px;">#</th>'
        '<th style="padding:10px 8px;width:130px;">SDS ID</th>'
        '<th style="padding:10px 8px;width:170px;">Product</th>'
        '<th style="padding:10px 8px;width:130px;">Supplier</th>'
        '<th style="padding:10px 8px;width:95px;">SDS Date</th>'
        '<th style="padding:10px 8px;width:105px;">Alert Level</th>'
        '<th style="padding:10px 8px;">Alert Justification</th>'
        '</tr></thead><tbody>'
    )
    rows = []
    for i, r in enumerate(results):
        lvl = r["alert_level"]
        bg, fg, emoji = ALERT_COLORS.get(lvl, ("#888", "#fff", "⚠️"))
        badge = (
            '<span style="background:' + bg + ';color:' + fg + ';'
            'padding:3px 10px;border-radius:4px;font-size:0.82em;'
            'font-weight:600;white-space:nowrap;">' + emoji + ' ' + lvl + '</span>'
        )
        justif_html = r.get("alert_justification", "").replace("\n", "<br>")
        row_bg = (
            "#FFF8F8" if lvl == "CRITICAL" else
            "#FFE5E5" if lvl == "MAJOR" else
            "#FFF3E0" if lvl == "MINOR" else
            "#F5FFF8" if lvl == "NONE" else "#F5F5F5"
        )
        rows.append(
            '<tr style="background:' + row_bg + ';vertical-align:top;border-bottom:1px solid #EEF4FF;">'
            '<td style="padding:10px 8px;text-align:center;color:#666;font-size:0.85em;">' + str(i+1) + '</td>'
            '<td style="padding:10px 8px;font-size:0.82em;color:#444;word-break:break-all;">' + str(r.get("sds_id", "") or "—") + '</td>'
            '<td style="padding:10px 8px;font-weight:600;color:#222;">' + str(r.get("product", "") or "—") + '</td>'
            '<td style="padding:10px 8px;color:#444;">' + str(r.get("supplier", "") or "—") + '</td>'
            '<td style="padding:10px 8px;color:#444;white-space:nowrap;">' + str(r.get("sds_date", "") or "—") + '</td>'
            '<td style="padding:10px 8px;text-align:center;">' + badge + '</td>'
            '<td style="padding:10px 8px;font-size:0.85em;color:#333;line-height:1.6;">' + justif_html + '</td>'
            '</tr>'
        )
    return thead + "".join(rows) + "</tbody></table>"

# ── 1. SDS Upload ─────────────────────────────────────────────────────────────
st.markdown(
    "<h2 style=\'color:#007AFF;font-weight:700;margin:24px 0 4px 0;\'>1 — SDS Upload</h2>",
    unsafe_allow_html=True,
)
st.caption("Max 5 files per session (beta)")
pdf_uploads = st.file_uploader(
    "Upload one or more SDS files (PDF)",
    type=["pdf"], accept_multiple_files=True)

st.divider()

# ── 4. Analysis ───────────────────────────────────────────────────────────────
st.markdown(
    "<h2 style=\'color:#007AFF;font-weight:700;margin:24px 0 4px 0;\'>2 — Analysis</h2>",
    unsafe_allow_html=True,
)
st.caption("Each SDS is processed in ~1–2 min. The application extracts the chemical composition and collects data across 17 biocompatibility endpoints, cross-references each compound against 8 databases, and generates an AI-assisted risk assessment.")
can_run = bool(api_key and pdf_uploads)
run = st.button("🚀 Run Analysis", disabled=not can_run, type="primary", use_container_width=True)
if not can_run and pdf_uploads is not None:
    missing = [x for x, ok in [("Anthropic API key", api_key), ("SDS PDF(s)", pdf_uploads)] if not ok]
    if missing:
        st.info(f"Missing: {', '.join(missing)}")

# ── 5. Processing ─────────────────────────────────────────────────────────────
results = []
excel_bytes = None
fname = None
if run:
    # ── Beta limit check ─────────────────────────────────────────────────────
    is_admin     = client_email.strip().lower() == ADMIN_EMAIL.strip().lower()
    max_files    = None if is_admin else BETA_LIMIT
    if not is_admin and len(pdf_uploads) > BETA_LIMIT:
        st.warning(
            f"⚠️ Beta version: analysis limited to **{BETA_LIMIT} SDS per session**. "
            f"{len(pdf_uploads) - BETA_LIMIT} additional file(s) ignored. "
            
            
        )
        pdf_uploads = pdf_uploads[:BETA_LIMIT]

    missing_files = [(p,n) for p,n in [
        (PROMPT_FILE,"agent_prompt_generic.txt"),
        (DB_FILE,    "databases.xlsx"),
        (TMPL_FILE,  "template.xlsx"),
    ] if not os.path.exists(p)]
    if missing_files:
        for _,n in missing_files:
            st.error(f"❌ Fichier manquant : **{n}**")
        st.stop()

    with open(PROMPT_FILE, "r", encoding="utf-8") as f:
        base_prompt = f.read()

    with st.spinner("Loading databases…"):
        with open(DB_FILE, "rb") as f:
            db_bytes = f.read()
        dbs = load_databases(db_bytes)

    with open(TMPL_FILE, "rb") as f:
        tmpl_bytes = f.read()

    wb = load_workbook(io.BytesIO(tmpl_bytes))
    if "SDS Analysis" in wb.sheetnames:
        apply_cf(wb["SDS Analysis"])

    client_api = anthropic.Anthropic(api_key=api_key)

    for idx, pdf_file in enumerate(pdf_uploads):
        st.subheader(f"📄 {pdf_file.name}  ({idx+1}/{len(pdf_uploads)})")
        prog = st.progress(0)
        msg  = st.empty()

        try:
            msg.info("📖 Extracting PDF text…")
            prog.progress(10)
            pdf_bytes = pdf_file.read()
            doc   = fitz.open(stream=pdf_bytes, filetype="pdf")
            pages = [f"--- PAGE {i+1} ---\n{p.get_text()}" for i,p in enumerate(doc)]
            doc.close()
            pdf_text = "\n".join(pages)

            if len(pdf_text.strip()) < 200:
                st.warning("⚠️ Very short text — PDF may be image-based (scanned).")

            sds_format = detect_sds_format(pdf_text)
            fmt_label  = "16-section GHS" if sds_format == "16-section" else "Old MSDS format (8 sections)"

            msg.info("🔢 Extracting CAS numbers…")
            prog.progress(20)
            cas_list = sorted(set(re.findall(r'\b\d{1,7}-\d{2}-\d\b', pdf_text)))
            with st.expander("Technical details"):
                st.caption(f"📋 Format detected: {fmt_label}")
                st.caption(f"{len(cas_list)} CAS found: "
                           f"{', '.join(cas_list[:12])}{'…' if len(cas_list)>12 else ''}") 

            msg.info("📊 Cross-referencing databases…")
            prog.progress(35)
            lookup_report = lookup_all_cas(cas_list, dbs)

            msg.info("🤖 AI analysis in progress (~1–2 min)…")
            prog.progress(50)

            system_prompt = build_system_prompt(base_prompt, config, sds_format)

            config_block = (
                f"DEVICE CONFIGURATION FOR THIS ANALYSIS:\n"
                f"- Client: {config['client_name']}\n"
                f"- SDS Format: {fmt_label}\n"
                f"- Contact Scenario: General screening\n"
            )
            db_metadata = "\n".join(
                f"- {disp}: source update shown as {normalize_db_version(db_dates.get(sheet, '—'))}"
                for sheet, disp in DB_DISPLAY_NAMES
            )
            user_msg = (
                f"{config_block}\n"
                "=== DATABASE EXTRACTION METADATA ===\n"
                "The regulatory database files supplied to this application were extracted "
                "and verified on 4 May 2026. Some source update dates may be older; do not "
                "treat an older date as a missing database.\n"
                f"{db_metadata}\n\n"
                f"=== SDS TEXT ===\n{pdf_text[:60000]}\n\n"
                f"=== PRE-COMPUTED DATABASE LOOKUP RESULTS ===\n{lookup_report}\n\n"
                "Use the database results above as authoritative sources. "
                "Return ONLY the required JSON object, with no Step 0A, no report, "
                "no Markdown, and no commentary before or after it."
            )

            response = client_api.messages.create(
                model="claude-sonnet-5-5",
                max_tokens=16000,
                system=system_prompt,
                messages=[{"role":"user","content":user_msg}],
            )
            # Sonnet 5.5 may return a ThinkingBlock before the final TextBlock.
            # Read only text blocks instead of assuming content[0] has a .text attribute.
            raw = "\n".join(
                block.text
                for block in response.content
                if getattr(block, "type", None) == "text"
                and getattr(block, "text", None)
            ).strip()
            if not raw:
                raise RuntimeError("Anthropic response contained no text block.")
            prog.progress(80)

            msg.info("📋 Parsing JSON response…")
            data = extract_json_response(raw)

            # Sonnet 5.5 can occasionally spend the first response on a diagnostic
            # explanation instead of returning the application JSON. Retry once with
            # an explicit JSON-only instruction before marking the SDS as failed.
            if not data:
                msg.info("🔁 Retrying structured JSON response…")
                retry_msg = (
                    user_msg
                    + "\n\nURGENT RETRY: The previous response did not contain the required JSON. "
                    "Return only one valid JSON object using exactly the required keys. "
                    "Do not output Step 0A, a database freshness report, Markdown, or commentary."
                )
                retry_response = client_api.messages.create(
                    model="claude-sonnet-5-5",
                    max_tokens=16000,
                    system=system_prompt,
                    messages=[{"role": "user", "content": retry_msg}],
                )
                raw = "\n".join(
                    block.text
                    for block in retry_response.content
                    if getattr(block, "type", None) == "text"
                    and getattr(block, "text", None)
                ).strip()
                data = extract_json_response(raw)

            if not data:
                error_reason = "JSON non trouvé dans la réponse de l'IA · JSON not found in AI response"
                st.warning(f"⚠️ {error_reason}")
                with st.expander("Technical details: raw AI response", expanded=False):
                    st.text_area("Raw AI response", raw, height=300)
                results.append({
                    "sds_id": "—", "product": pdf_file.name, "supplier": "—",
                    "sds_date": "—", "alert_level": "ERROR",
                    "alert_justification": f"⚠️ Traitement échoué : {error_reason}",
                    "tox_comment": "", "status": "Error", "format": fmt_label,
                })
                prog.progress(100); msg.empty()
                continue

            ws_sds   = wb["SDS Analysis"]
            next_row = 5
            while ws_sds.cell(row=next_row, column=1).value is not None:
                next_row += 1
            dashboard_row = next_row - 1

            msg.info("📝 Writing to Excel…")
            prog.progress(90)
            write_row(wb, data, next_row)
            extend_dashboard(wb, next_row, dashboard_row)

            alert_level = compute_alert_level(data)
            results.append({
                "sds_id":             data.get("sds_id",""),
                "product":            data.get("product_name", pdf_file.name),
                "supplier":           data.get("supplier",""),
                "sds_date":           data.get("sds_date",""),
                "alert_level":        alert_level,
                "alert_justification": data.get("alert_justification",""),
                "tox_comment":        "",
                "status":             "Draft",
                "format":             fmt_label,
            })

            prog.progress(100); msg.empty()
            st.success(f"✅ **{data.get('product_name', pdf_file.name)}** — row {next_row} written.")


        except anthropic.AuthenticationError:
            st.error("❌ Invalid API key."); break
        except anthropic.RateLimitError:
            st.error("❌ Rate limit reached — wait 30 s."); break
        except Exception as exc:
            error_reason = str(exc)
            st.error(f"❌ Error: {exc}")
            results.append({
                "sds_id":"—","product":pdf_file.name,"supplier":"—","sds_date":"—",
                "alert_level":"ERROR",
                "alert_justification":f"⚠️ Technical error: {error_reason[:200]}",
                "tox_comment":"","status":"Error","format":"—",
            })
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    # ── 6. Results summary ────────────────────────────────────────────────────
    if results:
        update_project_tab(wb, DB_FILE)
        enforce_report_formatting(wb)
        apply_table_font_and_alignment(wb)
        add_report_logos(wb)
        if "Project" in wb.sheetnames:
            wb.active = wb.sheetnames.index("Project")
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.calculation.calcMode = "auto"
        except Exception:
            pass
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        excel_bytes = buf.getvalue()
        fname = (f"SDS_Analysis_{config.get('client_name','').replace(' ','_')}"
                 f"_{datetime.date.today().strftime('%Y%m%d')}.xlsx")
        st.session_state["last_results"] = results
        st.session_state["last_excel_bytes"] = excel_bytes
        st.session_state["last_excel_name"] = fname

        st.divider()
        _section_title(f"3 — Preview — {len(results)} SDS analysed")
        st.caption(
            "The summary table is displayed below. Download the complete Excel report "
            "for the full analysis and regulatory database screening results."
        )

        st.markdown(_build_results_table(results), unsafe_allow_html=True)
        for r in results:
            if r["alert_level"] == "ERROR":
                st.warning("⚠️ **" + r["product"] + "** — " + r["alert_justification"])

        st.download_button(
            label=f"📥 Download {fname}",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            on_click="ignore",
        )

        st.divider()
        _section_title("4 — Beta tester feedback")
        st.write("How would you rate your experience with the SDS Screener?")
        st.caption("Please do not include confidential client information or confidential SDS content in your comments.")
        with st.form("feedback_form", clear_on_submit=False):
            rating = st.selectbox(
                "Rating *",
                options=[1, 2, 3, 4, 5],
                index=None,
                placeholder="Select a rating",
                format_func=lambda value: "★" * value + "☆" * (5 - value),
            )
            comments = st.text_area("Comments or suggestions", max_chars=2000)
            feedback_submitted = st.form_submit_button("Submit feedback", type="primary")
            if feedback_submitted:
                if rating is None:
                    st.error("Please select a rating before submitting your feedback.")
                else:
                    ok, message = _notion_submit_feedback(
                        client_email, client_name, int(rating), comments
                    )
                    if ok:
                        st.success("Thank you for your feedback.")
                    else:
                        st.warning(message)

# Re-render the completed report after download or feedback form reruns.
if not run and st.session_state.get("last_results") and st.session_state.get("last_excel_bytes"):
    results = st.session_state["last_results"]
    excel_bytes = st.session_state["last_excel_bytes"]
    fname = st.session_state["last_excel_name"]
    st.divider()
    _section_title(f"3 — Preview — {len(results)} SDS analysed")
    st.caption(
        "The summary table is displayed below. Download the complete Excel report "
        "for the full analysis and regulatory database screening results."
    )
    st.markdown(_build_results_table(results), unsafe_allow_html=True)
    for r in results:
        if r["alert_level"] == "ERROR":
            st.warning("⚠️ **" + r["product"] + "** — " + r["alert_justification"])
    st.download_button(
        label=f"📥 Download {fname}",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        on_click="ignore",
        key="cached_download",
    )
    _section_title("4 — Beta tester feedback")
    st.write("How would you rate your experience with the SDS Screener?")
    st.caption("Please do not include confidential client information or confidential SDS content in your comments.")
    with st.form("feedback_form", clear_on_submit=False):
        rating = st.selectbox(
            "Rating *", [1, 2, 3, 4, 5], index=None,
            placeholder="Select a rating",
            format_func=lambda value: "★" * value + "☆" * (5 - value),
        )
        comments = st.text_area("Comments or suggestions", max_chars=2000)
        feedback_submitted = st.form_submit_button("Submit feedback", type="primary")
        if feedback_submitted:
            if rating is None:
                st.error("Please select a rating before submitting your feedback.")
            else:
                ok, message = _notion_submit_feedback(
                    client_email, client_name, int(rating), comments
                )
                if ok:
                    st.success("Thank you for your feedback.")
                else:
                    st.warning(message)

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-aconis">
    Aconis SAS — 6, rue Bouchard 69510 Messimy, France<br>
    RCS Lyon n° 944 426 113 — SIRET 944 426 113 00014<br>
    <a href="mailto:elodie.saudrais@aconis.fr" style="color:#007AFF;">elodie.saudrais@aconis.fr</a>
    &nbsp;·&nbsp; +33 6 11 38 53 65
</div>
""", unsafe_allow_html=True)