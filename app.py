#!/usr/bin/env python3
"""
SDS Biocompatibility Screening Tool
SIGVARIS GROUP — Powered by Claude AI
MDR 2017/745 / ISO 10993-1:2025
"""

import streamlit as st
import anthropic
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import column_index_from_string
import pandas as pd
import json
import io
import datetime
import re
import os
import traceback

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SDS Biocompatibility Screener",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f4f7fb; }
.block-title { font-size: 1.6rem; font-weight: 700; color: #007AFF; margin-bottom: 0.2rem; }
.block-sub  { font-size: 0.95rem; color: #555; margin-bottom: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ── Excel constants ───────────────────────────────────────────────────────────
DATA_FILL = PatternFill("solid", fgColor="C5E1FF")

# Ordered column map: (Excel letter, JSON key)
# AC is the ALERT LEVEL formula — written separately
# AF is always empty (toxicologist) — written separately
COLS = [
    ("A",  "sds_id"),
    ("B",  "product_name"),
    ("C",  "supplier"),
    ("D",  "product_type"),
    ("E",  "sds_date"),
    ("F",  "composition"),
    ("G",  "cytotoxicity"),
    ("H",  "sensitisation"),
    ("I",  "skin_irritation"),
    ("J",  "eye_irritation"),
    ("K",  "acute_systemic_toxicity"),
    ("L",  "subacute_subchronic"),
    ("M",  "chronic_toxicity"),
    ("N",  "genotoxicity"),
    ("O",  "carcinogenicity"),
    ("P",  "reproductive_toxicity"),
    ("Q",  "endocrine_disruption"),
    ("R",  "bioaccumulation"),
    ("S",  "haemocompatibility"),
    ("T",  "pyrogenicity"),
    ("U",  "implantation"),
    ("V",  "immune_responses"),
    ("W",  "other_biological_effects"),
    ("X",  "cmr_regulatory_status"),
    ("Y",  "azo_dyes"),
    ("Z",  "formaldehyde"),
    ("AA", "heavy_metals"),
    ("AB", "reach_svhc"),
    # AC = formula
    ("AD", "alert_justification"),
    ("AE", "residual_on_fiber"),
    # AF = empty (toxicologist)
    # AG = corrected alert level — left blank for human reviewer
]

# ── Helper: load all lookup databases ────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_databases(excel_bytes: bytes) -> dict:
    """Load lookup databases from the Excel workbook."""
    wb = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    dbs = {}

    def rows_to_df(sheet_name: str, header_row: int) -> pd.DataFrame:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < header_row:
            return pd.DataFrame()
        hdrs = [str(h).strip() if h else f"_col{i}" for i, h in enumerate(rows[header_row - 1])]
        return pd.DataFrame(rows[header_row:], columns=hdrs)

    # SVHC Candidates — header row 6
    try:
        df = rows_to_df("SVHC Candidates", 6)
        df["CAS No."] = df["CAS No."].astype(str).str.strip()
        dbs["svhc"] = df
    except Exception:
        pass

    # CLP Annex VI ATP22 — header row 6
    try:
        df = rows_to_df("CLP Annex VI ATP22", 6)
        df["CAS No"] = df["CAS No"].astype(str).str.strip()
        dbs["clp"] = df
    except Exception:
        pass

    # IARC — header row 6
    try:
        df = rows_to_df("IARC", 6)
        df["CAS No."] = df["CAS No."].astype(str).str.strip()
        dbs["iarc"] = df
    except Exception:
        pass

    # Aromatic Amines (Ent.43) — find header dynamically
    try:
        ws = wb["Aromatic Amines (Ent.43)"]
        rows = list(ws.iter_rows(values_only=True))
        for i, row in enumerate(rows):
            if row[0] and "name" in str(row[0]).lower():
                hdrs = [str(h).strip() if h else f"_col{j}" for j, h in enumerate(row)]
                dbs["aromatic_amines"] = pd.DataFrame(rows[i + 1:], columns=hdrs)
                break
    except Exception:
        pass

    # Agent rules (specific) — residual lookup table (rows ~64–120)
    try:
        ws = wb["Agent rules (specific)"]
        rows = list(ws.iter_rows(values_only=True))
        residual = {}
        for row in rows[62:]:  # starts around row 63 in the sheet
            if row[0] and row[3] is not None:
                key = str(row[0]).strip().lstrip("@ ").strip().lower()
                val = row[3]
                if val == 1:
                    residual[key] = "Yes - Sigvaris confirmed"
                elif val == 0:
                    residual[key] = "No - Sigvaris confirmed"
        dbs["residual"] = residual
    except Exception:
        pass

    wb.close()
    return dbs


def lookup_residual(product_name: str, residual_db: dict) -> str:
    key = product_name.strip().lower()
    return residual_db.get(key, "To provide - Sigvaris classification required")


def lookup_cas_in_databases(cas_list: list, dbs: dict) -> str:
    """Search CAS numbers in databases and return a compact text report."""
    if not cas_list:
        return "No CAS numbers found in SDS text."

    lines = [f"PRE-COMPUTED DATABASE LOOKUP — {len(cas_list)} CAS number(s) extracted from SDS:", "=" * 60]

    for cas in cas_list:
        found = False
        entry = [f"\nCAS {cas}:"]

        if "svhc" in dbs:
            hits = dbs["svhc"][dbs["svhc"]["CAS No."].str.contains(cas, regex=False, na=False)]
            for _, r in hits.iterrows():
                found = True
                entry.append(f"  [SVHC] {r.get('Substance name', '')} — Reason: {r.get('Reason for inclusion', '')}")

        if "clp" in dbs:
            hits = dbs["clp"][dbs["clp"]["CAS No"].str.contains(cas, regex=False, na=False)]
            for _, r in hits.iterrows():
                found = True
                h_codes = str(r.get("Hazard Statement Code(s)", "")).replace("\n", " ").strip()
                hazard  = str(r.get("Hazard Class and Category Code(s)", "")).replace("\n", " / ").strip()
                entry.append(f"  [CLP ATP22] {r.get('International Chemical Identification', '')} | {hazard} | {h_codes}")

        if "iarc" in dbs:
            hits = dbs["iarc"][dbs["iarc"]["CAS No."].str.contains(cas, regex=False, na=False)]
            for _, r in hits.iterrows():
                found = True
                entry.append(f"  [IARC] Group {r.get('Group', '')} — {r.get('Agent', '')}")

        if not found:
            entry.append("  Not found in SVHC, CLP ATP22, or IARC")

        lines.extend(entry)

    return "\n".join(lines)


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages of a PDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        pages.append(f"--- PAGE {i + 1} ---\n{page.get_text()}")
    doc.close()
    return "\n".join(pages)


def extract_cas_numbers(text: str) -> list:
    """Return unique CAS numbers found in text."""
    return sorted(set(re.findall(r'\b\d{1,7}-\d{2}-\d\b', text)))


def parse_json_block(text: str) -> dict:
    """Extract the JSON block from Claude's response."""
    m = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    # Fallback: try to find raw JSON object
    m = re.search(r'\{[^{}]{50,}\}', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return {}


def write_row_to_excel(wb, data: dict) -> int:
    """Append a new analysis row to SDS Analysis tab. Returns the row number written."""
    ws = wb["SDS Analysis"]

    # Find next empty row (data starts at row 3)
    row = 3
    while ws.cell(row=row, column=1).value is not None:
        row += 1

    def cidx(col_letter: str) -> int:
        return column_index_from_string(col_letter)

    # Write data columns
    for col_letter, json_key in COLS:
        value = data.get(json_key, "")
        cell = ws.cell(row=row, column=cidx(col_letter), value=value)
        cell.fill = DATA_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Col AC — ALERT LEVEL formula (auto-calculated)
    ac = ws.cell(row=row, column=cidx("AC"))
    ac.value = (
        f'=IF(COUNTIF(G{row}:AB{row},"*🔴*")>0,"CRITICAL",'
        f'IF(COUNTIF(G{row}:AB{row},"*🟠*")>0,"MAJOR",'
        f'IF(COUNTIF(G{row}:AB{row},"*🟡*")>0,"MINOR","NONE")))'
    )
    ac.fill = DATA_FILL
    ac.alignment = Alignment(horizontal="center", vertical="center")

    # Col AF — always empty (reserved for toxicologist)
    af = ws.cell(row=row, column=cidx("AF"), value="/")
    af.fill = DATA_FILL
    af.alignment = Alignment(vertical="top")

    # Col AG — blank (corrected alert level, filled by toxicologist)
    ws.cell(row=row, column=cidx("AG"), value="").fill = DATA_FILL

    # Col AH — Analysis Date
    ah = ws.cell(row=row, column=cidx("AH"), value=datetime.date.today().strftime("%d/%m/%Y"))
    ah.fill = DATA_FILL
    ah.alignment = Alignment(vertical="top")

    # Col AI — Analyst
    ai = ws.cell(row=row, column=cidx("AI"), value="AI - SDS Screening agent only")
    ai.fill = DATA_FILL
    ai.alignment = Alignment(vertical="top")

    return row


def build_system_prompt(base_prompt: str, client_config: dict) -> str:
    """Replace client placeholders and append JSON output instruction."""
    prompt = base_prompt
    prompt = prompt.replace("<<CLIENT_NAME>>",             client_config.get("client_name", "CLIENT"))
    prompt = prompt.replace("<<DEVICE_DESCRIPTION>>",      client_config.get("device_description", "medical device"))
    prompt = prompt.replace("<<CONTACT_SCENARIO>>",        client_config.get("contact_scenario", "External skin contact — long-term use"))
    prompt = prompt.replace("<<PRODUCT_TYPE_VOCABULARY>>", client_config.get("product_type_vocabulary", ""))
    return prompt + """

---

## PYTHON APP — MANDATORY JSON OUTPUT

After your screening report, output a JSON block wrapped in ```json ... ```.
Use EXACTLY these keys. Values must match what you would write in the Excel cells
(with markers 🔴🟠🟡, newlines as \\n, source tags, etc.).

```json
{
  "sds_id": "",
  "product_name": "",
  "supplier": "",
  "product_type": "",
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
  "alert_justification": "",
  "residual_on_fiber": ""
}
```
"""


# ── MAIN APP ──────────────────────────────────────────────────────────────────

st.markdown('<p class="block-title">🔬 SDS Biocompatibility Screener</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="block-sub">SIGVARIS GROUP — Analyse automatique des FDS '
    '(MDR 2017/745 / ISO 10993-1:2025)</p>',
    unsafe_allow_html=True,
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")

    api_key = st.text_input(
        "Clé API Anthropic",
        type="password",
        placeholder="sk-ant-...",
        help="Depuis console.anthropic.com > API Keys",
    )

    st.divider()
    st.header("📁 Fichier Excel")
    excel_upload = st.file_uploader(
        "SIGVARIS FDS.xlsx (avec bases de données)",
        type=["xlsx"],
        help="Votre fichier Excel contenant les onglets SVHC, CLP, IARC, etc.",
    )
    if excel_upload:
        st.success(f"✅ {excel_upload.name}")

    st.divider()
    st.divider()
    st.header("🏢 Configuration client")

    client_name = st.text_input(
        "Nom du client",
        value="SIGVARIS GROUP",
        help="Remplace <<CLIENT_NAME>> dans le prompt",
    )
    device_description = st.text_input(
        "Description du dispositif",
        value="medical compression hosiery",
        help="Remplace <<DEVICE_DESCRIPTION>> dans le prompt",
    )
    contact_scenario = st.text_input(
        "Scénario de contact",
        value="External skin contact with damaged/compromised skin — long-term chronic use",
        help="Remplace <<CONTACT_SCENARIO>> dans le prompt",
    )

    DEFAULT_PRODUCT_TYPES = """| Value | Definition |
|---|---|
| `Fibre - Elastic` | Elastane / Spandex / Lycra fibers |
| `Fibre - Cellulosic` | Cotton / Linen / Viscose / Bamboo viscose / Rayon |
| `Fibre - Non-cellulosic synthetic` | Polyamide / Polyester / PBT / PTFE / Polypropylene / Aramid / Silk |
| `Fibre - Silicone` | Silicone bands / grip bands / border elements |
| `Dyeing agent` | Colorant designed to fix to fiber |
| `Auxiliary dyeing agent` | Process chemical used in dye bath alongside colorant |
| `Processing aid` | Lubricant / oil / finish applied outside dye bath |
| `To verify` | Cannot be determined from SDS data alone |"""

    product_type_vocabulary = st.text_area(
        "Vocabulaire Product Type (Markdown)",
        value=DEFAULT_PRODUCT_TYPES,
        height=180,
        help="Remplace <<PRODUCT_TYPE_VOCABULARY>> dans le prompt. Adaptez selon le client.",
    )

    client_config = {
        "client_name":            client_name,
        "device_description":     device_description,
        "contact_scenario":       contact_scenario,
        "product_type_vocabulary": product_type_vocabulary,
    }

    st.divider()
    st.caption("v1.0 — Claude Sonnet 4.6 — ISO 10993-1:2025")

# ── Main columns ─────────────────────────────────────────────────────────────
left, right = st.columns([3, 1])

with left:
    st.subheader("📄 Déposer les FDS à analyser")
    pdf_uploads = st.file_uploader(
        "Fichiers PDF (une ou plusieurs FDS)",
        type=["pdf"],
        accept_multiple_files=True,
    )

with right:
    st.subheader("Guide rapide")
    st.markdown("""
**1.** Clé API dans la barre latérale  
**2.** Chargez votre Excel SIGVARIS  
**3.** Déposez vos PDF  
**4.** Cliquez **Analyser**  
**5.** Téléchargez l'Excel mis à jour

⏱ ~1 min par FDS  
🔒 Données non stockées
""")

# ── Run button ────────────────────────────────────────────────────────────────
can_run = bool(api_key and excel_upload and pdf_uploads)
run = st.button("🚀 Analyser", disabled=not can_run, type="primary", use_container_width=True)

if not can_run and not run:
    missing = []
    if not api_key:     missing.append("clé API Anthropic")
    if not excel_upload: missing.append("fichier Excel")
    if not pdf_uploads:  missing.append("FDS PDF")
    if missing:
        st.info(f"Manquant(s) : {', '.join(missing)}")

# ── Processing ────────────────────────────────────────────────────────────────
if run:
    st.divider()

    # Load agent prompt
    prompt_file = os.path.join(os.path.dirname(__file__), "agent_prompt_generic.txt")
    if not os.path.exists(prompt_file):
        st.error(
            "❌ Fichier **agent_prompt_generic.txt** introuvable.\n\n"
            "Placez-le dans le même dossier que app.py."
        )
        st.stop()

    with open(prompt_file, "r", encoding="utf-8") as f:
        base_prompt = f.read()
    system_prompt = build_system_prompt(base_prompt, client_config)

    # Load Excel
    excel_bytes = excel_upload.read()

    with st.spinner("Chargement des bases de données..."):
        try:
            dbs = load_databases(excel_bytes)
            db_names = [k for k in dbs if k != "residual"]
            st.success(f"✅ Bases chargées : {', '.join(db_names)} + table résidualité")
        except Exception as e:
            st.error(f"Erreur chargement Excel : {e}")
            st.stop()

    wb = load_workbook(io.BytesIO(excel_bytes))
    client = anthropic.Anthropic(api_key=api_key)
    results = []

    for idx, pdf_file in enumerate(pdf_uploads):
        st.subheader(f"📄 {pdf_file.name}  ({idx + 1} / {len(pdf_uploads)})")
        prog = st.progress(0)
        msg  = st.empty()

        try:
            # 1. Extract text
            msg.info("📖 Extraction du texte PDF...")
            prog.progress(10)
            pdf_bytes_data = pdf_file.read()
            pdf_text = extract_pdf_text(pdf_bytes_data)

            if len(pdf_text.strip()) < 200:
                st.warning("⚠️ Peu de texte extrait — le PDF est peut-être scanné en image.")

            # 2. Extract CAS numbers
            msg.info("🔢 Extraction des numéros CAS...")
            prog.progress(20)
            cas_list = extract_cas_numbers(pdf_text)
            st.caption(f"{len(cas_list)} CAS trouvé(s) : {', '.join(cas_list[:15])}{'…' if len(cas_list) > 15 else ''}")

            # 3. Database lookups
            msg.info("📊 Recherche dans les bases de données...")
            prog.progress(35)
            lookup_report = lookup_cas_in_databases(cas_list, dbs)

            # 4. Call Claude
            msg.info("🤖 Analyse IA en cours... (1-2 min)")
            prog.progress(50)

            user_msg = f"""Please analyze the Safety Data Sheet below.

=== SDS TEXT ===
{pdf_text[:60000]}

=== PRE-COMPUTED DATABASE RESULTS ===
{lookup_report}

Note: SVHC, CLP ATP22, and IARC lookups above are pre-computed from your Excel databases.
For REACH Annex XVII, Aromatic Amines Ent.43, ED Assessment (ECHA), TEDX, and Oeko-Tex,
apply your embedded knowledge based on the CAS numbers listed.

Please complete the full analysis and output the JSON block at the end.
"""

            response = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=8000,
                system=system_prompt,
                messages=[{"role": "user", "content": user_msg}],
            )
            response_text = response.content[0].text
            prog.progress(80)

            # 5. Parse JSON
            msg.info("📋 Structuration des données...")
            data = parse_json_block(response_text)

            if not data:
                st.warning("⚠️ Impossible d'extraire le JSON. Rapport brut ci-dessous :")
                st.text_area("Réponse Claude", response_text, height=400)
                prog.progress(100)
                msg.empty()
                continue

            # 6. Residual lookup override if product name known
            product_name = data.get("product_name", "")
            if product_name and "residual" in dbs:
                residual_val = lookup_residual(product_name, dbs["residual"])
                if residual_val != "To provide - Sigvaris classification required":
                    data["residual_on_fiber"] = residual_val

            # 7. Write to Excel
            msg.info("📝 Écriture dans l'Excel...")
            prog.progress(90)
            row_num = write_row_to_excel(wb, data)

            results.append({"file": pdf_file.name, "product": product_name, "row": row_num})

            prog.progress(100)
            msg.empty()
            st.success(
                f"✅ **{product_name or pdf_file.name}** — "
                f"Analyse terminée, ligne **{row_num}** ajoutée."
            )

            with st.expander("📊 Rapport de screening complet"):
                st.markdown(response_text[:6000] + ("…" if len(response_text) > 6000 else ""))

        except anthropic.AuthenticationError:
            st.error("❌ Clé API invalide. Vérifiez votre clé Anthropic.")
            break
        except anthropic.RateLimitError:
            st.error("❌ Limite de requêtes atteinte. Attendez 30 s puis relancez.")
            break
        except Exception as exc:
            st.error(f"❌ Erreur inattendue : {exc}")
            with st.expander("Détail technique"):
                st.code(traceback.format_exc())

    # ── Download ──────────────────────────────────────────────────────────────
    if results:
        st.divider()
        st.subheader(f"🎉 {len(results)} FDS analysée(s) avec succès")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        filename = f"SIGVARIS_FDS_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label=f"📥 Télécharger {filename}",
            data=buf.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.table([{"FDS": r["file"], "Produit": r["product"], "Ligne Excel": r["row"]} for r in results])
