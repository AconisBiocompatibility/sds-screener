#!/usr/bin/env python3
"""SDS Biocompatibility Screening Tool — ACONIS | Claude AI | ISO 10993-1:2025"""

import streamlit as st
import anthropic
import fitz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import column_index_from_string
from openpyxl.formatting.rule import FormulaRule
import pandas as pd
import json, io, datetime, re, os, traceback

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="SDS Biocompatibility Screener", page_icon="🔬",
                   layout="wide", initial_sidebar_state="expanded")

# ─── File paths ───────────────────────────────────────────────────────────────
DIR         = os.path.dirname(__file__)
DB_FILE     = os.path.join(DIR, "databases.xlsx")
TMPL_FILE   = os.path.join(DIR, "template.xlsx")
PROMPT_FILE = os.path.join(DIR, "agent_prompt_generic.txt")

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

# ─── Column map — ACONIS structure (A–AH, 34 cols) ───────────────────────────
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
    # AB = ALERT LEVEL formula (written separately)
    ("AC","alert_justification"),
    # AD = Toxicologist Comment — always empty
    # AE = Corrected Alert Level — blank for reviewer
    # AF = Analysis Date — auto
    # AG = Analyst — auto
    # AH = Row Status — always Draft
]

# ─── ISO 10993-1:2025 options ─────────────────────────────────────────────────
CONTACT_NATURE_OPTIONS = [
    "Medical devices in contact with intact skin",
    "Medical devices in contact with intact mucosal membrane",
    "Medical devices in contact with either breached or compromised surfaces "
    "(skin or mucosal membranes) or internal tissues other than circulating blood",
    "Medical devices in contact with circulating blood",
]
CONTACT_DURATION_OPTIONS = [
    "A – limited (≤24 h)",
    "B – prolonged (>24 h to 30 d)",
    "C – long-term (>30 d)",
]

# ─── Database loading ─────────────────────────────────────────────────────────
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

    register("clp",      "CLP Annex VI ATP22")
    register("svhc",     "SVHC Candidates")
    register("reach",    "REACH Annex XVII")
    register("amines",   "Aromatic Amines (Ent.43)")
    register("azo_jtf",  "Azo - Restricted Amines (JTF)")
    register("iarc",     "IARC")
    register("ed",       "ED Assessment (ECHA)")
    register("tedx",     "TEDX")

    wb.close()
    return dbs


def lookup_all_cas(cas_list: list, dbs: dict) -> str:
    if not cas_list:
        return "No CAS numbers extracted from the SDS text."
    today = datetime.date.today().strftime("%b %Y")
    lines = [
        f"PRE-COMPUTED DATABASE LOOKUP — {len(cas_list)} CAS extracted",
        "=" * 60,
    ]
    for cas in cas_list:
        hits = []
        # CLP
        if "clp" in dbs:
            cc = dbs["clp_cas"]
            m  = dbs["clp"][dbs["clp"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                codes = str(r.get("Hazard Statement Code(s)", "")).replace("\n", " ")
                cat   = str(r.get("Hazard Class and Category Code(s)", ""))
                name  = r.get("International Chemical Identification", "")
                hits.append(f"  [CLP Annex VI ATP22] {name} | {cat} | H-codes: {codes}")
        # SVHC
        if "svhc" in dbs:
            cc = dbs["svhc_cas"]
            m  = dbs["svhc"][dbs["svhc"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                reason = r.get("Reason for inclusion", r.get("reason", ""))
                name   = r.get("Substance name", r.get("name", ""))
                hits.append(f"  [SVHC Candidates] {name} — Reason: {reason}")
        # IARC
        if "iarc" in dbs:
            cc = dbs["iarc_cas"]
            m  = dbs["iarc"][dbs["iarc"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                grp   = r.get("Group", r.get("group", ""))
                agent = r.get("Agent", r.get("agent", ""))
                hits.append(f"  [IARC, {today}] Group {grp} — {agent}")
        # REACH Annex XVII
        if "reach" in dbs:
            cc = dbs["reach_cas"]
            m  = dbs["reach"][dbs["reach"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                entry = r.get("Entry", r.get("entry", ""))
                restr = str(r.get("Restriction", r.get("restriction", "")))[:120]
                hits.append(f"  [REACH Annex XVII] Entry {entry} — {restr}")
        # Aromatic Amines
        if "amines" in dbs:
            cc = dbs["amines_cas"]
            m  = dbs["amines"][dbs["amines"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append(f"  [Aromatic Amines Ent.43] MATCH — restricted aromatic amine")
        # Azo JTF
        if "azo_jtf" in dbs:
            cc = dbs["azo_jtf_cas"]
            m  = dbs["azo_jtf"][dbs["azo_jtf"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append(f"  [Azo - Restricted Amines JTF] MATCH — restricted azo amine")
        # ED Assessment
        if "ed" in dbs:
            cc = dbs["ed_cas"]
            m  = dbs["ed"][dbs["ed"][cc].str.contains(cas, regex=False, na=False)]
            for _, r in m.iterrows():
                status = str(r.get("Status", r.get("status", "")))
                hits.append(f"  [ED Assessment (ECHA), {today}] {status}")
        # TEDX
        if "tedx" in dbs:
            cc = dbs["tedx_cas"]
            m  = dbs["tedx"][dbs["tedx"][cc].str.contains(cas, regex=False, na=False)]
            if not m.empty:
                hits.append(f"  [TEDX] Listed as endocrine disruptor")

        if hits:
            lines.append(f"\nCAS {cas}:")
            lines.extend(hits)
        else:
            lines.append(
                f"\nCAS {cas}: Not found in CLP ATP22, SVHC, IARC, "
                "REACH Annex XVII, Aromatic Amines, ED Assessment, TEDX"
            )
    return "\n".join(lines)


# ─── Excel helpers ────────────────────────────────────────────────────────────
def parse_date(s: str) -> str:
    if not s or str(s).strip() in ("", "Not stated", "N/A", "not stated"):
        return str(s) if s else "Not stated"
    for fmt in ("%d/%m/%Y","%m/%d/%Y","%Y-%m-%d","%d-%m-%Y",
                "%B %d, %Y","%b %d, %Y","%d %B %Y","%d %b %Y"):
        try:
            return datetime.datetime.strptime(str(s).strip(), fmt).strftime("%b %d, %Y")
        except Exception:
            pass
    return str(s)


def apply_cf(ws):
    """Apply all conditional formatting (ACONIS structure)."""
    # AB: Alert Level cell colors (1-5)
    ws.conditional_formatting.add("AB3:AB9999", fr('$AB3="CRITICAL"',   "FF0000", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AB3:AB9999", fr('$AB3="MAJOR"',      "FF8C00", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AB3:AB9999", fr('$AB3="MINOR"',      "FFC000", True,  False, "000000"))
    ws.conditional_formatting.add("AB3:AB9999", fr('$AB3="NONE"',       "00B050", False, False, "FFFFFF"))
    ws.conditional_formatting.add("AB3:AB9999", fr('$AB3="NOT AN SDS"', "808080", False, False, "FFFFFF"))
    # Full row background (6-8) — lock column, not row
    ws.conditional_formatting.add("A3:AH9999", fr('$AB3="CRITICAL"', "FFE6E6"))
    ws.conditional_formatting.add("A3:AH9999", fr('$AB3="MAJOR"',    "FFF2E6"))
    ws.conditional_formatting.add("A3:AH9999", fr('$AB3="MINOR"',    "FFFDE6"))
    # AG: Analyst indicator (9-10)
    ws.conditional_formatting.add("AG3:AG9999",
        fr('$AG3="AI - SDS Screening agent only"', "FFF2CC", False, True))
    ws.conditional_formatting.add("AG3:AG9999",
        fr('AND($AG3<>"",$AG3<>"AI - SDS Screening agent only")', "E2EFDA"))
    # AD: Toxicologist comment (11)
    ws.conditional_formatting.add("AD3:AD9999", fr('$AD3<>""', "E2EFDA", True))
    # AH: Row Status (12-15)
    ws.conditional_formatting.add("AH3:AH9999", fr('$AH3="Draft"',      "BDD7EE"))
    ws.conditional_formatting.add("AH3:AH9999", fr('$AH3="Validated"',  "00B050", True,  False, "FFFFFF"))
    ws.conditional_formatting.add("AH3:AH9999", fr('$AH3="Superseded"', "D9D9D9", False, False, "808080"))
    ws.conditional_formatting.add("AH3:AH9999", fr('$AH3="Archived"',   "595959", False, False, "FFFFFF"))
    # D: SDS date > 3 years
    ws.conditional_formatting.add("D3:D9999",
        fr('AND(D3<>"",ISNUMBER(D3),D3<TODAY()-1095)', "FFE0E0", False, False, "CC0000"))
    # "TO VERIFY" cells — lavender
    ws.conditional_formatting.add("E3:AC9999", FormulaRule(
        formula=['NOT(ISERROR(SEARCH("TO VERIFY",E3)))'],
        fill=PatternFill("solid", fgColor="E2CCFF"),
        font=Font(name="Segoe UI Light", size=10),
    ))


def write_row(wb, data: dict, row: int):
    ws = wb["SDS Analysis"]
    for col_letter, key in COLS:
        val = data.get(key, "")
        if key == "sds_date":
            val = parse_date(str(val)) if val else "Not stated"
        c = ws.cell(row=row, column=cidx(col_letter), value=val)
        c.fill      = DATA_FILL
        c.alignment = WRAP_TOP
        c.font      = seg()

    # AB: ALERT LEVEL formula scanning F:AA
    ab = ws.cell(row=row, column=cidx("AB"))
    ab.value = (
        f'=IF(COUNTIF(F{row}:AA{row},"*🔴*")>0,"CRITICAL",'
        f'IF(COUNTIF(F{row}:AA{row},"*🟠*")>0,"MAJOR",'
        f'IF(COUNTIF(F{row}:AA{row},"*🟡*")>0,"MINOR","NONE")))'
    )
    ab.fill      = DATA_FILL
    ab.alignment = Alignment(horizontal="center", vertical="center")
    ab.font      = seg(bold=True)

    # AD: Toxicologist Comment — always empty
    ws.cell(row=row, column=cidx("AD"), value="").fill = DATA_FILL
    # AE: Corrected Alert Level — blank for reviewer
    ws.cell(row=row, column=cidx("AE"), value="").fill = DATA_FILL

    # AF: Analysis Date
    today = datetime.date.today()
    af = ws.cell(row=row, column=cidx("AF"),
                 value=datetime.datetime(today.year, today.month, today.day))
    af.fill          = DATA_FILL
    af.alignment     = WRAP_TOP
    af.font          = seg()
    af.number_format = DATE_FMT

    # AG: Analyst
    ag = ws.cell(row=row, column=cidx("AG"), value="AI - SDS Screening agent only")
    ag.fill      = DATA_FILL
    ag.alignment = WRAP_TOP
    ag.font      = seg(italic=True)

    # AH: Row Status
    ah = ws.cell(row=row, column=cidx("AH"), value="Draft")
    ah.fill      = DATA_FILL
    ah.alignment = WRAP_TOP
    ah.font      = seg()


def extend_dashboard(wb, n: int):
    if "SDS Dashboard" not in wb.sheetnames:
        return
    ws = wb["SDS Dashboard"]
    if ws.cell(row=n, column=1).value is not None:
        return
    formulas = [
        ("A", f"=ROW()-2"),
        ("B", f"='SDS Analysis'!$A{n}"),
        ("C", f"='SDS Analysis'!$B{n}"),
        ("D", f"='SDS Analysis'!$C{n}"),
        ("E", f"='SDS Analysis'!$D{n}"),
        ("F", f"='SDS Analysis'!$AE{n}"),
        ("G", f"='SDS Analysis'!$AC{n}&IF('SDS Analysis'!$AD{n}<>\"\",CHAR(10)&'SDS Analysis'!$AD{n},\"\")"),
        ("H", f"='SDS Analysis'!$AF{n}"),
        ("I", f"='SDS Analysis'!$AG{n}"),
        ("J", f"='SDS Analysis'!$AH{n}"),
    ]
    for col_l, val in formulas:
        c = ws.cell(row=n, column=cidx(col_l), value=val)
        c.fill      = DATA_FILL
        c.alignment = WRAP_TOP
        c.font      = seg()
        if col_l in ("E", "H"):
            c.number_format = DATE_FMT


def write_project_tab(wb, config: dict, results: list):
    if "Project" not in wb.sheetnames:
        wb.create_sheet("Project", 0)
    ws = wb["Project"]
    ws.delete_rows(1, ws.max_row + 1)

    blue  = PatternFill("solid", fgColor="007AFF")
    light = PatternFill("solid", fgColor="EAF4FF")

    def row_w(r, a, b="", a_fill=None, a_font=None, b_font=None):
        ca = ws.cell(row=r, column=1, value=a)
        cb = ws.cell(row=r, column=2, value=b)
        ca.fill = a_fill or PatternFill()
        cb.fill = PatternFill()
        ca.font = a_font or seg()
        cb.font = b_font or seg()
        ca.alignment = Alignment(vertical="center", wrap_text=True)
        cb.alignment = Alignment(vertical="center", wrap_text=True)

    def section(r, label):
        row_w(r, label, a_fill=blue,
              a_font=seg(bold=True, color="FFFFFF", size=11))

    row_w(1, "ACONIS BIOCOMPATIBILITY SCREENING — SESSION RECORD",
          a_fill=blue, a_font=seg(bold=True, color="FFFFFF", size=13))
    row_w(2, "")
    row_w(3, "Generated:",    datetime.datetime.now().strftime("%b %d, %Y — %H:%M"))
    row_w(4, "Tool Version:", "v2.0 — Claude Sonnet 4.5 | ISO 10993-1:2025 / MDR 2017/745")
    row_w(5, "")
    section(6,  "DEVICE CONFIGURATION")
    row_w(7,  "Client Name:",        config.get("client_name", ""))
    row_w(8,  "Device Description:", config.get("device_description", ""))
    row_w(9,  "Device Type:",        config.get("device_type", ""))
    row_w(10, "")
    section(11, "CONTACT CLASSIFICATION — ISO 10993-1:2025")
    row_w(12, "Nature of Contact:", config.get("contact_nature", ""))
    row_w(13, "Duration Category:", config.get("contact_duration", ""))
    row_w(14, "")
    section(15, "ANALYSIS SUMMARY")
    row_w(16, "Files Analyzed:", str(len(results)))
    row_w(17, "")

    hdrs = ["#", "SDS ID", "Product Name", "Supplier", "Analyst"]
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(row=18, column=i, value=h)
        c.fill      = blue
        c.font      = seg(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")

    for i, r in enumerate(results, 1):
        row_data = [i, r.get("sds_id",""), r.get("product",""),
                    r.get("supplier",""), "AI - SDS Screening agent only"]
        for j, val in enumerate(row_data, 1):
            c = ws.cell(row=18+i, column=j, value=val)
            c.fill      = light
            c.font      = seg()
            c.alignment = Alignment(vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 25
    ws.column_dimensions["E"].width = 30


# ─── Prompt builder ───────────────────────────────────────────────────────────
def build_system_prompt(base_prompt: str, config: dict) -> str:
    is_textile = config.get("device_type", "Non-textile") == "Textile"
    azo_instr = (
        "Col X (Azo dyes / Aromatic amines): run normal analysis — "
        "check Aromatic Amines (Ent.43) and Azo - Restricted Amines (JTF) lookup results above."
        if is_textile else
        "Col X (Azo dyes / Aromatic amines): write exactly "
        "`N/A — not applicable: non-textile device [Agent assessment]`"
    )
    contact_scenario = (
        f"{config.get('contact_nature','')} — {config.get('contact_duration','')}"
    )

    prompt = base_prompt
    prompt = prompt.replace("<<CLIENT_NAME>>",          config.get("client_name", "CLIENT"))
    prompt = prompt.replace("<<DEVICE_DESCRIPTION>>",   config.get("device_description", "medical device"))
    prompt = prompt.replace("<<CONTACT_SCENARIO>>",     contact_scenario)
    prompt = prompt.replace("<<AZO_DYES_INSTRUCTION>>", azo_instr)

    json_schema = """\n\n---\n\n## PYTHON APP — MANDATORY JSON OUTPUT\n
After your screening report, output a JSON block in ```json ... ```.
Use EXACTLY these keys. Values must include all markers (🔴🟠🟡), source tags, newlines as \\n.\n
```json
{
  "sds_id": "",
  "product_name": "",
  "supplier": "",
  "sds_date": "",
  "composition": "",
  "cytotoxicity": "",
  "sensitisation": "",
  "skin_irritation": "",
  "eye_irritation": "",
  "acute_systemic_toxicity": "",
  "subacute_subchronic": "",
  "chronic_toxicity": "",
  "genotoxicity": "",
  "carcinogenicity": "",
  "reproductive_toxicity": "",
  "endocrine_disruption": "",
  "bioaccumulation": "",
  "haemocompatibility": "",
  "pyrogenicity": "",
  "implantation": "",
  "immune_responses": "",
  "other_biological_effects": "",
  "cmr_regulatory_status": "",
  "azo_dyes": "",
  "formaldehyde": "",
  "heavy_metals": "",
  "reach_svhc": "",
  "alert_justification": ""
}
```
"""
    return prompt + json_schema


# ─── UI ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { min-width: 350px; }
h1 { color: #007AFF; }
</style>
""", unsafe_allow_html=True)

st.title("🔬 SDS Biocompatibility Screener")
st.caption("ACONIS — ISO 10993-1:2025 / MDR 2017/745 — Claude AI")

with st.sidebar:
    st.header("⚙️ Configuration")
    # Lire la clé depuis les secrets Streamlit (invisible pour l'utilisateur)
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    st.divider()

    st.subheader("🏢 Client")
    client_name        = st.text_input("Client Name", placeholder="e.g. Acme Medical")
    device_description = st.text_input("Device Description",
                                       placeholder="e.g. medical compression hosiery")
    device_type        = st.radio("Device Type", ["Textile", "Non-textile"], horizontal=True)
    st.divider()

    st.subheader("📋 ISO 10993-1:2025 Classification")
    contact_nature   = st.selectbox("Nature of Contact", CONTACT_NATURE_OPTIONS)
    contact_duration = st.radio("Contact Duration", CONTACT_DURATION_OPTIONS)
    st.divider()
    st.caption("v2.0 — Claude Sonnet 4.5 — ISO 10993-1:2025")

    config = {
        "client_name":        client_name,
        "device_description": device_description,
        "device_type":        device_type,
        "contact_nature":     contact_nature,
        "contact_duration":   contact_duration,
    }

left, right = st.columns([3, 1])
with left:
    st.subheader("📄 SDS Files (PDF)")
    pdf_uploads = st.file_uploader(
        "Upload one or more SDS files", type=["pdf"], accept_multiple_files=True)
with right:
    st.subheader("Quick Guide")
    st.markdown("""
**1.** Configure client & device  
**2.** Upload SDS PDF(s)  
**3.** Click **Analyze**  
**4.** Download Excel  

⏱ ~1 min per SDS  
🔒 No data stored
""")

can_run = bool(api_key and pdf_uploads)
run = st.button("🚀 Analyze", disabled=not can_run, type="primary", use_container_width=True)

if not can_run:
    missing = [x for x, ok in [("Anthropic API key", api_key), ("SDS PDF(s)", pdf_uploads)] if not ok]
    if missing:
        st.info(f"Missing: {', '.join(missing)}")

# ─── Processing ───────────────────────────────────────────────────────────────
if run:
    st.divider()

    # Validate required files
    missing_files = [(p, n) for p, n in [
        (PROMPT_FILE, "agent_prompt_generic.txt"),
        (DB_FILE,     "databases.xlsx"),
        (TMPL_FILE,   "template.xlsx"),
    ] if not os.path.exists(p)]
    if missing_files:
        for _, n in missing_files:
            st.error(f"❌ Missing file: **{n}** — upload it to GitHub in the same folder as app.py")
        st.stop()

    with open(PROMPT_FILE, "r", encoding="utf-8") as f:
        base_prompt = f.read()
    system_prompt = build_system_prompt(base_prompt, config)

    with st.spinner("Loading databases..."):
        with open(DB_FILE, "rb") as f:
            db_bytes = f.read()
        dbs = load_databases(db_bytes)
        loaded = [dbs[k].replace("_sheet","") for k in dbs if k.endswith("_sheet")]
    st.success(f"✅ {len(loaded)} databases loaded")

    with open(TMPL_FILE, "rb") as f:
        tmpl_bytes = f.read()
    wb = load_workbook(io.BytesIO(tmpl_bytes))

    if "SDS Analysis" in wb.sheetnames:
        apply_cf(wb["SDS Analysis"])

    client  = anthropic.Anthropic(api_key=api_key)
    results = []

    for idx, pdf_file in enumerate(pdf_uploads):
        st.subheader(f"📄 {pdf_file.name}  ({idx+1}/{len(pdf_uploads)})")
        prog = st.progress(0)
        msg  = st.empty()

        try:
            msg.info("📖 Extracting PDF text…")
            prog.progress(10)
            pdf_bytes = pdf_file.read()
            doc   = fitz.open(stream=pdf_bytes, filetype="pdf")
            pages = [f"--- PAGE {i+1} ---\n{p.get_text()}" for i, p in enumerate(doc)]
            doc.close()
            pdf_text = "\n".join(pages)

            if len(pdf_text.strip()) < 200:
                st.warning("⚠️ Very short text — PDF may be image-based (scanned).")

            msg.info("🔢 Extracting CAS numbers…")
            prog.progress(20)
            cas_list = sorted(set(re.findall(r'\b\d{1,7}-\d{2}-\d\b', pdf_text)))
            st.caption(f"{len(cas_list)} CAS found: "
                       f"{', '.join(cas_list[:12])}{'…' if len(cas_list)>12 else ''}")

            msg.info("📊 Database lookups…")
            prog.progress(35)
            lookup_report = lookup_all_cas(cas_list, dbs)

            msg.info("🤖 AI analysis in progress (~1–2 min)…")
            prog.progress(50)

            # Device config header for user message
            config_block = (
                f"DEVICE CONFIGURATION FOR THIS ANALYSIS:\n"
                f"- Client: {config['client_name']}\n"
                f"- Device: {config['device_description']}\n"
                f"- Device Type: {config['device_type']}\n"
                f"- Contact Nature (ISO 10993-1:2025): {config['contact_nature']}\n"
                f"- Contact Duration: {config['contact_duration']}\n"
            )
            if config["device_type"] == "Non-textile":
                config_block += (
                    "- Col X INSTRUCTION: write exactly "
                    "`N/A — not applicable: non-textile device [Agent assessment]`\n"
                )

            user_msg = (
                f"{config_block}\n"
                f"=== SDS TEXT ===\n{pdf_text[:60000]}\n\n"
                f"=== PRE-COMPUTED DATABASE LOOKUP RESULTS ===\n{lookup_report}\n\n"
                "Use the database results above as authoritative sources for "
                "CLP, SVHC, IARC, REACH, Aromatic Amines, ED Assessment lookups. "
                "Complete the full analysis then output the JSON block."
            )

            response = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=8000,
                system=system_prompt,
                messages=[{"role": "user", "content": user_msg}],
            )
            raw = response.content[0].text
            prog.progress(80)

            msg.info("📋 Parsing JSON…")
            m = re.search(r'```json\s*(.*?)\s*```', raw, re.DOTALL)
            data = {}
            if m:
                try:
                    data = json.loads(m.group(1))
                except Exception:
                    pass

            if not data:
                st.warning("⚠️ No JSON found in response — displaying raw report.")
                st.text_area("Claude response", raw, height=400)
                prog.progress(100); msg.empty()
                continue

            # Find next empty row in SDS Analysis
            ws_sds  = wb["SDS Analysis"]
            next_row = 3
            while ws_sds.cell(row=next_row, column=1).value is not None:
                next_row += 1

            msg.info("📝 Writing to Excel…")
            prog.progress(90)
            write_row(wb, data, next_row)
            extend_dashboard(wb, next_row)

            results.append({
                "sds_id":   data.get("sds_id", ""),
                "product":  data.get("product_name", pdf_file.name),
                "supplier": data.get("supplier", ""),
                "row":      next_row,
            })

            prog.progress(100); msg.empty()
            st.success(f"✅ **{data.get('product_name', pdf_file.name)}** — row {next_row} written.")

            with st.expander("📊 Full screening report"):
                st.markdown(raw[:6000] + ("…" if len(raw) > 6000 else ""))

        except anthropic.AuthenticationError:
            st.error("❌ Invalid API key — check your key."); break
        except anthropic.RateLimitError:
            st.error("❌ Rate limit reached — wait 30 s and retry."); break
        except Exception as exc:
            st.error(f"❌ Error: {exc}")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    if results:
        write_project_tab(wb, config, results)
        st.divider()
        st.subheader(f"🎉 {len(results)} SDS analyzed")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        fname = (
            f"SDS_Analysis_{config.get('client_name','').replace(' ','_')}"
            f"_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        )
        st.download_button(
            label=f"📥 Download {fname}",
            data=buf.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.table([{
            "#": i+1,
            "Product": r["product"],
            "Supplier": r["supplier"],
            "Excel Row": r["row"],
        } for i, r in enumerate(results)])
