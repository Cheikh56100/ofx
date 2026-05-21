"""
BanqScan — Convertisseur OCR de relevés bancaires
Mode GRATUIT  : Tesseract OCR (local, sans clé API)
Mode PREMIUM  : Claude AI (Anthropic, optionnel)
Formats image : JPEG, JFIF, JPG, PNG, BMP, TIFF, WEBP, GIF, TGA, etc.
"""

import streamlit as st
import base64, json, csv, io, re, os, tempfile, time
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageEnhance, ImageFilter
import pytesseract

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, HRFlowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

# ───────────────────────────────────────────────────────────────
# CONFIG PAGE
# ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BanqScan — OCR Relevés Bancaires",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ───────────────────────────────────────────────────────────────
# CSS
# ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background: #f0f4f8; }

/* Hero */
.hero {
  background: linear-gradient(135deg, #0f2b4a 0%, #1a4b7a 55%, #1e6fa0 100%);
  border-radius: 16px; padding: 34px 40px 30px;
  margin-bottom: 24px; color: white; position: relative; overflow: hidden;
}
.hero::before {
  content:''; position:absolute; top:-50px; right:-50px;
  width:220px; height:220px; border-radius:50%;
  background:rgba(255,255,255,0.05);
}
.hero h1 {
  font-family:'DM Serif Display',serif; font-size:2rem;
  margin:0 0 6px; letter-spacing:-0.5px;
}
.hero p { margin:0; opacity:.8; font-size:.92rem; line-height:1.6; }
.hero .badge {
  display:inline-block; background:rgba(255,255,255,.15);
  border:1px solid rgba(255,255,255,.25); border-radius:20px;
  padding:3px 12px; font-size:.7rem; font-weight:600;
  margin-bottom:12px; letter-spacing:.8px; text-transform:uppercase;
}

/* Mode pills */
.mode-free {
  background:#d1fae5; color:#065f46; border:1px solid #6ee7b7;
  border-radius:20px; padding:4px 14px; font-size:.75rem; font-weight:700;
  display:inline-block; margin-bottom:8px;
}
.mode-ai {
  background:#dbeafe; color:#1e40af; border:1px solid #93c5fd;
  border-radius:20px; padding:4px 14px; font-size:.75rem; font-weight:700;
  display:inline-block; margin-bottom:8px;
}

/* Cards */
.card {
  background:white; border-radius:14px; padding:22px 24px;
  box-shadow:0 1px 4px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.04);
  margin-bottom:16px;
}
.card-title {
  font-size:.72rem; font-weight:700; text-transform:uppercase;
  letter-spacing:1px; color:#6b7a8d; margin-bottom:10px;
}

/* Stats */
.stats-row { display:flex; gap:12px; margin:14px 0; }
.stat-box {
  flex:1; background:white; border-radius:12px; padding:14px 16px;
  box-shadow:0 1px 3px rgba(0,0,0,.06); border-left:4px solid #1a4b7a;
}
.stat-box.red   { border-left-color:#e53e3e; }
.stat-box.green { border-left-color:#2f855a; }
.stat-box.blue  { border-left-color:#2b6cb0; }
.stat-box.orange{ border-left-color:#c05621; }
.stat-box .label {
  font-size:.68rem; font-weight:700; text-transform:uppercase;
  letter-spacing:.8px; color:#8a9bb0; margin-bottom:3px;
}
.stat-box .value { font-size:1.3rem; font-weight:600; color:#1a2b3c; font-family:'DM Serif Display',serif; }
.stat-box.red   .value { color:#c53030; }
.stat-box.green .value { color:#276749; }
.stat-box.orange .value { color:#c05621; }

/* Stepper */
.step-row {
  display:flex; margin-bottom:22px; border-radius:10px;
  overflow:hidden; border:1px solid #d9e4ef;
}
.step {
  flex:1; padding:9px 12px; background:white; font-size:.78rem;
  font-weight:500; color:#8a9bb0; text-align:center;
  border-right:1px solid #d9e4ef;
}
.step:last-child { border-right:none; }
.step.active { background:#1a4b7a; color:white; font-weight:700; }
.step.done   { background:#ebf4ff; color:#1a4b7a; }

/* Buttons */
.stButton > button {
  background:linear-gradient(135deg,#1a4b7a,#1e6fa0) !important;
  color:white !important; border:none !important; border-radius:10px !important;
  font-weight:600 !important; padding:10px 24px !important;
  font-size:.92rem !important; transition:all .2s !important;
  box-shadow:0 2px 8px rgba(26,75,122,.3) !important;
}
.stButton > button:hover {
  transform:translateY(-1px) !important;
  box-shadow:0 4px 14px rgba(26,75,122,.4) !important;
}

/* Op table */
.op-table { width:100%; border-collapse:collapse; font-size:.8rem; }
.op-table th {
  background:#0f2b4a; color:white; padding:9px 11px;
  text-align:left; font-weight:500; font-size:.72rem; letter-spacing:.3px;
}
.op-table th.right { text-align:right; }
.op-table td { padding:7px 11px; border-bottom:1px solid #edf2f7; color:#2d3748; vertical-align:top; }
.op-table tr:nth-child(even) td { background:#f7fafc; }
.op-table tr:hover td { background:#ebf4ff; }
.op-table td.debit  { text-align:right; color:#c53030; font-weight:600; }
.op-table td.credit { text-align:right; color:#276749; font-weight:600; }
.op-table td.date   { white-space:nowrap; color:#6b7a8d; font-size:.75rem; }
.op-table tfoot td  { background:#ebf4ff; font-weight:700; font-size:.82rem; border-top:2px solid #b8c8d9; }
.op-table tfoot td.debit  { color:#c53030; text-align:right; }
.op-table tfoot td.credit { color:#276749; text-align:right; }

/* Compte grid */
.compte-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:14px; }
.compte-item .k {
  font-size:.66rem; font-weight:700; text-transform:uppercase;
  letter-spacing:.8px; color:#8a9bb0; margin-bottom:2px;
}
.compte-item .v { font-size:.86rem; font-weight:600; color:#1a2b3c; }

/* Alert box */
.alert-info {
  background:#ebf8ff; border:1px solid #bee3f8; border-radius:10px;
  padding:12px 16px; font-size:.83rem; color:#2c5282;
}
.alert-warn {
  background:#fffbeb; border:1px solid #fcd34d; border-radius:10px;
  padding:12px 16px; font-size:.83rem; color:#92400e;
}

/* Sidebar */
[data-testid="stSidebar"] { background:#0f2b4a !important; }
[data-testid="stSidebar"] * { color:#cdd8e3 !important; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:white !important; }
[data-testid="stSidebar"] .stSelectbox label,[data-testid="stSidebar"] .stRadio label { color:#aabfcf !important; }

::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:#f0f4f8; }
::-webkit-scrollbar-thumb { background:#b8c8d9; border-radius:3px; }
</style>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────────────────────────
# CONSTANTS
# ───────────────────────────────────────────────────────────────
ACCEPTED_EXTENSIONS = [
    "jpg","jpeg","jfif","jpe","png","bmp","tiff","tif",
    "webp","gif","tga","ppm","pgm","pbm","dib","avif",
]
MIME_MAP = {
    "jpg":"image/jpeg","jpeg":"image/jpeg","jfif":"image/jpeg",
    "jpe":"image/jpeg","png":"image/png","bmp":"image/bmp",
    "tiff":"image/tiff","tif":"image/tiff","webp":"image/webp",
    "gif":"image/gif","tga":"image/x-tga","avif":"image/avif",
}

LOGICIELS = [
    "Sage Comptabilité","EBP Compta","Cegid","Quadratus",
    "QuickBooks","Pennylane","FEC Manager","Autre",
]


# ───────────────────────────────────────────────────────────────
# IMAGE UTILS
# ───────────────────────────────────────────────────────────────
def open_image_universal(file_bytes: bytes, filename: str) -> Image.Image:
    """Open any image format, convert to RGB."""
    buf = io.BytesIO(file_bytes)
    try:
        img = Image.open(buf)
        img.load()
    except Exception:
        # Fallback: save to temp file and reopen
        ext = Path(filename).suffix.lower() or ".jpg"
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        img = Image.open(tmp_path)
        img.load()
        os.unlink(tmp_path)

    # Normalize to RGB
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    return img


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """Enhance image quality for better OCR accuracy."""
    # Convert to grayscale
    img = img.convert("L")
    # Boost contrast
    img = ImageEnhance.Contrast(img).enhance(2.2)
    # Sharpen
    img = img.filter(ImageFilter.SHARPEN)
    # Scale up small images
    w, h = img.size
    if w < 1200:
        scale = 1200 / w
        img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
    return img


def img_to_base64(file_bytes: bytes) -> str:
    return base64.b64encode(file_bytes).decode()


# ───────────────────────────────────────────────────────────────
# FORMAT HELPERS
# ───────────────────────────────────────────────────────────────
def fmt_eur(val):
    if not val:
        return ""
    try:
        v = float(val)
        if v == 0:
            return ""
        return f"{v:,.2f} €".replace(",", " ").replace(".", ",")
    except Exception:
        return str(val)


# ───────────────────────────────────────────────────────────────
# MODE GRATUIT — TESSERACT OCR
# ───────────────────────────────────────────────────────────────
def _extract_compte_tesseract(text: str) -> dict:
    compte = {
        "titulaire": "", "banque": "", "agence": "",
        "rib": "", "iban": "", "periode": "", "type_releve": "",
    }
    lines = text.splitlines()

    for line in lines:
        up = line.upper()
        # Titulaire / Raison sociale
        if "NISHA" in up or "SARL" in up or "SAS " in up or "EURL" in up:
            if not compte["titulaire"]:
                compte["titulaire"] = line.strip().lstrip("|: ").strip()
        # IBAN
        m = re.search(r'(FR\d{2}[\s\d]{20,30})', line)
        if m:
            compte["iban"] = re.sub(r'\s+', ' ', m.group(1)).strip()
        # RIB
        m = re.search(r'(\d{5}[\s]*\d{5}[\s]*\d{11}[\s]*\d{2})', line)
        if m:
            compte["rib"] = m.group(1).strip()
        # Banque
        if "BNP" in up:
            compte["banque"] = "BNP PARIBAS"
        elif "CREDIT AGRICOLE" in up or "CRÉDIT AGRICOLE" in up:
            compte["banque"] = "CRÉDIT AGRICOLE"
        elif "SOCIETE GENERALE" in up or "SOCIÉTÉ GÉNÉRALE" in up:
            compte["banque"] = "SOCIÉTÉ GÉNÉRALE"
        elif "LA BANQUE POSTALE" in up:
            compte["banque"] = "LA BANQUE POSTALE"
        elif "CIC" in up and "BANQUE" in up:
            compte["banque"] = "CIC"
        elif "LCLE" in up or "LCL" in up:
            compte["banque"] = "LCL"
        # Agence / ville
        m_ville = re.search(r'([A-ZÉÈÀÂÊÎÔÛÙÄËÏÖÜ\-]{4,}(?:\s+[A-ZÉÈÀÂÊÎÔÛÙÄËÏÖÜ\-]{2,}){0,3})\s*$', line)
        if m_ville and not compte["agence"] and len(line.strip()) < 40:
            candidate = m_ville.group(1).strip()
            if candidate not in ("IBAN", "RIB", "DEBIT", "CREDIT"):
                compte["agence"] = candidate
        # Période
        m_per = re.search(
            r'(?:P[ÉE]RIODE|DU|RELEV[ÉE])\s+(?:DU\s+)?(\d{1,2}[\s\w]+\d{4}[\s\w]+\d{4})',
            up
        )
        if m_per:
            compte["periode"] = m_per.group(1).strip()
        elif "NOVEMBRE" in up or "DÉCEMBRE" in up or "JANVIER" in up or "AU" in up:
            if re.search(r'\d{4}', line):
                compte["periode"] = line.strip().lstrip("|: ").strip()
        # Type relevé
        if "RELEV" in up and "COMPTE" in up:
            compte["type_releve"] = line.strip().lstrip("|: ").strip()

    return compte


def _parse_operations_tesseract(text: str) -> list:
    """
    Parse bank statement lines from Tesseract OCR text.
    Strategy: find rows starting with date DD.MM.YY, then look for amount at end.
    """
    ops = []
    lines = text.splitlines()

    # Known debit keywords
    DEBIT_KEYWORDS = [
        "VIR SCT", "VIR INST", "VIREMENT", "PRLV", "PRÉLÈVEMENT",
        "CHQ", "CHEQUE", "CHÈQUE", "RETRAIT", "FRAIS", "COMMISSION",
        "CB ", "CARTE ", "INTERETS", "INTÉRÊTS", "AGIOS",
    ]

    date_pat = re.compile(r'^(\d{2}[./]\d{2}[./]\d{2,4})')
    amount_pat = re.compile(r'(\d{1,4}[\s]?\d{0,3}[,\.]\d{2})\s*$')

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        dm = date_pat.match(line)
        if not dm:
            i += 1
            continue

        date_comptable = dm.group(1).replace('/', '.')
        rest = line[dm.end():].strip().lstrip('|').strip()

        # Collect continuation lines (no date at start, not a separator)
        j = i + 1
        while j < len(lines):
            next_line = lines[j].strip()
            if not next_line or date_pat.match(next_line):
                break
            if re.match(r'^[-=]{5,}', next_line):
                break
            # Stop if line looks like header
            if any(k in next_line.upper() for k in ["DATE DE", "COMPTABLE", "NATURE DES", "DÉBIT", "CRÉDIT"]):
                break
            rest += " " + next_line.lstrip('|').strip()
            j += 1

        # Find date_valeur and amount in combined string
        dates = re.findall(r'\d{2}[./]\d{2}[./]\d{2,4}', rest)
        date_valeur = dates[0].replace('/', '.') if dates else date_comptable

        # Find amount — last number with decimals
        amounts = re.findall(r'(\d{1,4}[\s]?\d{0,3}[,\.]\d{2})', rest)
        if not amounts:
            i = j
            continue

        raw_amount = amounts[-1]
        montant = float(raw_amount.replace(' ', '').replace(',', '.'))

        # Clean libelle: remove dates and trailing amount
        libelle = rest
        for d in re.findall(r'\d{2}[./]\d{2}[./]\d{2,4}', libelle):
            libelle = libelle.replace(d, '')
        libelle = re.sub(r'\s*' + re.escape(raw_amount) + r'\s*$', '', libelle)
        libelle = re.sub(r'\s+', ' ', libelle).strip().lstrip('|').strip()

        # Determine debit or credit
        is_debit = any(kw in libelle.upper() for kw in DEBIT_KEYWORDS)
        # Also check if "DEBIT" or amount appears in debit column position
        debit_val = montant if is_debit else 0.0
        credit_val = 0.0 if is_debit else montant

        if libelle and montant > 0:
            ops.append({
                "date_comptable": date_comptable,
                "date_valeur": date_valeur,
                "libelle": libelle,
                "debit": debit_val,
                "credit": credit_val,
            })

        i = j

    return ops


def ocr_tesseract(file_bytes: bytes, filename: str) -> dict:
    """Full Tesseract OCR pipeline → structured dict."""
    img = open_image_universal(file_bytes, filename)
    img_proc = preprocess_for_ocr(img)

    # Run OCR with French language
    config = "--psm 6 --oem 3"
    text = pytesseract.image_to_string(img_proc, lang="fra", config=config)

    compte = _extract_compte_tesseract(text)
    ops = _parse_operations_tesseract(text)

    total_d = sum(o["debit"] for o in ops)
    total_c = sum(o["credit"] for o in ops)

    return {
        "compte": compte,
        "operations": ops,
        "total_debits": round(total_d, 2),
        "total_credits": round(total_c, 2),
        "solde_debut": 0.0,
        "solde_fin": 0.0,
        "_ocr_mode": "tesseract",
        "_raw_text": text,
    }


# ───────────────────────────────────────────────────────────────
# MODE PREMIUM — CLAUDE AI
# ───────────────────────────────────────────────────────────────
def ocr_claude(file_bytes: bytes, filename: str, api_key: str, logiciel: str) -> dict:
    """OCR via Claude Vision — highest accuracy."""
    try:
        import anthropic
    except ImportError:
        st.error("Le package `anthropic` n'est pas installé.")
        return {}

    client = anthropic.Anthropic(api_key=api_key)
    ext = Path(filename).suffix.lower().lstrip(".")
    media_type = MIME_MAP.get(ext, "image/jpeg")

    # For JFIF and other JPEG variants, force image/jpeg
    if ext in ("jfif", "jpe", "jpg", "jpeg"):
        media_type = "image/jpeg"

    # Convert to JPEG for Claude if needed
    img = open_image_universal(file_bytes, filename)
    buf_out = io.BytesIO()
    img.convert("RGB").save(buf_out, format="JPEG", quality=92)
    b64 = img_to_base64(buf_out.getvalue())
    media_type = "image/jpeg"

    system = (
        "Tu es un expert OCR en comptabilité bancaire française. "
        "Extrait TOUTES les données du relevé avec une précision maximale. "
        "Réponds UNIQUEMENT avec un objet JSON valide, sans markdown, sans backticks."
    )

    prompt = f"""Analyse ce relevé bancaire scanné et extrait toutes les informations.

Format JSON strict :
{{
  "compte": {{
    "titulaire": "...", "banque": "...", "agence": "...",
    "rib": "...", "iban": "...", "periode": "...", "type_releve": "..."
  }},
  "operations": [
    {{
      "date_comptable": "JJ.MM.AA",
      "date_valeur": "JJ.MM.AA",
      "libelle": "libellé complet",
      "debit": 0.00,
      "credit": 0.00
    }}
  ],
  "solde_debut": 0.00,
  "solde_fin": 0.00,
  "total_debits": 0.00,
  "total_credits": 0.00
}}

Règles impératives :
- debit et credit sont des floats. Utilise 0 si la colonne est vide.
- Copie les libellés COMPLETS sans les tronquer.
- N'oublie AUCUNE opération même les petits montants.
- Logiciel cible : {logiciel}"""

    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4000,
        system=system,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    data = json.loads(raw)
    data["_ocr_mode"] = "claude"
    return data


# ───────────────────────────────────────────────────────────────
# PDF GENERATION
# ───────────────────────────────────────────────────────────────
def build_pdf(data: dict, logiciel: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=14*mm, bottomMargin=12*mm,
    )

    c_main   = colors.HexColor("#0f2b4a")
    c_light  = colors.HexColor("#ebf4ff")
    c_red    = colors.HexColor("#c53030")
    c_green  = colors.HexColor("#276749")
    c_grey   = colors.HexColor("#8a9bb0")
    c_border = colors.HexColor("#d9e4ef")
    c_row2   = colors.HexColor("#f7fafc")

    def ps(name, **kw):
        base = {"fontName": "Helvetica", "fontSize": 9, "leading": 11}
        base.update(kw)
        return ParagraphStyle(name, **base)

    ops   = data.get("operations", [])
    compte = data.get("compte", {})
    total_d = data.get("total_debits") or sum(o.get("debit", 0) or 0 for o in ops)
    total_c = data.get("total_credits") or sum(o.get("credit", 0) or 0 for o in ops)
    mode  = data.get("_ocr_mode", "tesseract").upper()

    story = []

    # ── Header band ──
    hdr = [[
        Paragraph(compte.get("type_releve") or "RELEVÉ DE COMPTE COURANT EN EURO",
                  ps("tit", fontName="Helvetica-Bold", fontSize=15, textColor=c_main)),
        Paragraph(
            f'<font size="7" color="#8a9bb0">Généré le {datetime.now().strftime("%d/%m/%Y %H:%M")}'
            f' · Mode OCR : {mode} · {logiciel}</font>',
            ps("sub", alignment=TA_RIGHT)),
    ]]
    ht = Table(hdr, colWidths=[120*mm, 62*mm])
    ht.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW", (0,0), (-1,0), 1.5, c_main),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(ht)
    story.append(Spacer(1, 4*mm))

    # ── Compte info ──
    fields = [
        ("Titulaire",  compte.get("titulaire","")),
        ("Banque",     compte.get("banque","")),
        ("Agence",     compte.get("agence","")),
        ("RIB",        compte.get("rib","")),
        ("IBAN",       compte.get("iban","")),
        ("Période",    compte.get("periode","")),
    ]
    cells = []
    for i in range(0, len(fields), 2):
        row = []
        for j in range(2):
            if i+j < len(fields):
                k, v = fields[i+j]
                row.append(Paragraph(
                    f'<font size="6.5" color="#8a9bb0"><b>{k.upper()}</b></font><br/>'
                    f'<font size="9" color="#1a2b3c"><b>{v or "—"}</b></font>',
                    ps("cv")))
            else:
                row.append("")
        cells.append(row)

    info_t = Table(cells, colWidths=[91*mm, 91*mm])
    info_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), colors.HexColor("#f7fafc")),
        ("BOX",        (0,0),(-1,-1), 0.5, c_border),
        ("INNERGRID",  (0,0),(-1,-1), 0.3, c_border),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ]))
    story.append(info_t)
    story.append(Spacer(1, 4*mm))

    # ── Stats bar ──
    balance = total_c - total_d
    bal_color = "#276749" if balance >= 0 else "#c53030"
    stats = [[
        Paragraph(f'<font size="6.5" color="#8a9bb0"><b>OPÉRATIONS</b></font><br/>'
                  f'<font size="14" color="#0f2b4a"><b>{len(ops)}</b></font>', ps("s")),
        Paragraph(f'<font size="6.5" color="#8a9bb0"><b>TOTAL DÉBITS</b></font><br/>'
                  f'<font size="13" color="#c53030"><b>{fmt_eur(total_d) or "0,00 €"}</b></font>', ps("s")),
        Paragraph(f'<font size="6.5" color="#8a9bb0"><b>TOTAL CRÉDITS</b></font><br/>'
                  f'<font size="13" color="#276749"><b>{fmt_eur(total_c) or "0,00 €"}</b></font>', ps("s")),
        Paragraph(f'<font size="6.5" color="#8a9bb0"><b>SOLDE MOUVEMENT</b></font><br/>'
                  f'<font size="13" color="{bal_color}"><b>{fmt_eur(abs(balance)) or "0,00 €"}</b></font>', ps("s")),
    ]]
    st_t = Table(stats, colWidths=[44*mm, 46*mm, 46*mm, 46*mm])
    st_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), c_light),
        ("BOX",        (0,0),(-1,-1), 0.5, c_border),
        ("INNERGRID",  (0,0),(-1,-1), 0.3, c_border),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("LEFTPADDING",   (0,0),(-1,-1), 12),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(st_t)
    story.append(Spacer(1, 4*mm))

    # ── Operations table ──
    col_w = [21*mm, 21*mm, 97*mm, 23*mm, 23*mm]

    def hcell(txt, align=TA_LEFT):
        return Paragraph(txt, ps("hc", fontName="Helvetica-Bold", fontSize=7,
                                  textColor=colors.white, alignment=align))

    tbl = [[
        hcell("DATE\nCOMPTABLE", TA_CENTER),
        hcell("DATE\nVALEUR", TA_CENTER),
        hcell("NATURE DES OPÉRATIONS"),
        hcell("DÉBIT", TA_RIGHT),
        hcell("CRÉDIT", TA_RIGHT),
    ]]

    for op in ops:
        d  = op.get("debit",  0) or 0
        cr = op.get("credit", 0) or 0
        tbl.append([
            Paragraph(op.get("date_comptable",""), ps("dc", fontSize=7.5, textColor=colors.HexColor("#6b7a8d"), alignment=TA_CENTER)),
            Paragraph(op.get("date_valeur",""),    ps("dv", fontSize=7.5, textColor=colors.HexColor("#6b7a8d"), alignment=TA_CENTER)),
            Paragraph(op.get("libelle",""),        ps("lib", fontSize=7.5, leading=10)),
            Paragraph(fmt_eur(d),  ps("dr", fontSize=7.5, fontName="Helvetica-Bold" if d else "Helvetica", textColor=c_red, alignment=TA_RIGHT)),
            Paragraph(fmt_eur(cr), ps("cr", fontSize=7.5, fontName="Helvetica-Bold" if cr else "Helvetica", textColor=c_green, alignment=TA_RIGHT)),
        ])

    # Total row
    tbl.append([
        Paragraph("", ps("x")), Paragraph("", ps("x")),
        Paragraph("TOTAL", ps("tot", fontName="Helvetica-Bold", fontSize=8, textColor=c_main)),
        Paragraph(fmt_eur(total_d) or "0,00 €", ps("td", fontName="Helvetica-Bold", fontSize=8, textColor=c_red, alignment=TA_RIGHT)),
        Paragraph(fmt_eur(total_c) or "0,00 €", ps("tc", fontName="Helvetica-Bold", fontSize=8, textColor=c_green, alignment=TA_RIGHT)),
    ])

    n = len(tbl)
    ops_t = Table(tbl, colWidths=col_w, repeatRows=1)
    ops_t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), c_main),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ("RIGHTPADDING",  (0,0), (-1,-1), 5),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("GRID",          (0,0), (-1,-1), 0.3, c_border),
        *[("BACKGROUND",  (0,i), (-1,i), c_row2) for i in range(2, n-1, 2)],
        ("BACKGROUND",    (0,-1),(-1,-1), c_light),
        ("LINEABOVE",     (0,-1),(-1,-1), 1, c_main),
        ("FONTNAME",      (0,-1),(-1,-1), "Helvetica-Bold"),
    ]))
    story.append(ops_t)
    story.append(Spacer(1, 5*mm))

    # ── Footer ──
    story.append(HRFlowable(width="100%", thickness=0.4, color=c_border))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"{compte.get('banque','') or 'Banque'} — Document généré par BanqScan (OCR {mode}) — "
        f"Export {logiciel} — {datetime.now().strftime('%d/%m/%Y')}",
        ps("foot", fontSize=6.5, textColor=c_grey)
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ───────────────────────────────────────────────────────────────
# CSV GENERATION
# ───────────────────────────────────────────────────────────────
def build_csv(data: dict) -> bytes:
    ops = data.get("operations", [])
    c   = data.get("compte", {})
    total_d = data.get("total_debits") or sum(o.get("debit",0) or 0 for o in ops)
    total_c = data.get("total_credits") or sum(o.get("credit",0) or 0 for o in ops)

    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL)

    w.writerow(["RELEVÉ DE COMPTE — EXPORT COMPTABLE"])
    for k, v in [
        ("Titulaire", c.get("titulaire","")),
        ("IBAN",      c.get("iban","")),
        ("RIB",       c.get("rib","")),
        ("Banque",    c.get("banque","")),
        ("Période",   c.get("periode","")),
        ("OCR mode",  data.get("_ocr_mode","").upper()),
        ("Exporté le",datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]:
        w.writerow([k, v])

    w.writerow([])
    w.writerow(["Date comptable","Date valeur","Libellé","Débit","Crédit","Solde partiel"])

    running = 0.0
    for op in ops:
        d  = op.get("debit",  0) or 0
        cr = op.get("credit", 0) or 0
        running += cr - d
        w.writerow([
            op.get("date_comptable",""),
            op.get("date_valeur",""),
            op.get("libelle",""),
            f"{d:.2f}".replace(".",",")  if d  else "",
            f"{cr:.2f}".replace(".",",") if cr else "",
            f"{running:.2f}".replace(".",","),
        ])

    w.writerow([])
    w.writerow(["","","TOTAL",
                f"{total_d:.2f}".replace(".",","),
                f"{total_c:.2f}".replace(".",","),""])

    return ("\ufeff" + buf.getvalue()).encode("utf-8")


# ───────────────────────────────────────────────────────────────
# SIDEBAR
# ───────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")

    st.markdown("### 🔑 Mode OCR")
    ocr_mode = st.radio(
        "Mode",
        ["🆓 Gratuit (Tesseract)", "✨ Premium (Claude AI)"],
        label_visibility="collapsed",
    )
    use_claude = "Claude" in ocr_mode

    if use_claude:
        api_key = st.text_input(
            "Clé API Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="console.anthropic.com",
        )
        # Also check secrets
        if not api_key:
            api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    else:
        api_key = ""
        st.markdown("""
        <div class="alert-info">
        ✅ <b>Mode 100% gratuit</b><br/>
        OCR local via Tesseract.<br/>
        Aucune clé API requise.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🏢 Logiciel cible")
    logiciel = st.selectbox("Logiciel", LOGICIELS, label_visibility="collapsed")

    st.markdown("### 📄 Format d'export")
    export_fmt = st.radio(
        "Export",
        ["PDF + CSV", "PDF uniquement", "CSV uniquement"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("""
    <div style="font-size:.73rem;opacity:.55;line-height:1.7;">
    <b>BanqScan v2.0</b><br/>
    Formats : JPEG · JFIF · PNG · BMP<br/>
    TIFF · WEBP · GIF · TGA · AVIF…<br/><br/>
    Mode Gratuit : Tesseract OCR<br/>
    Mode Premium : Claude (Anthropic)<br/><br/>
    Données traitées localement
    </div>
    """, unsafe_allow_html=True)


# ───────────────────────────────────────────────────────────────
# MAIN LAYOUT
# ───────────────────────────────────────────────────────────────
mode_badge = '<span class="mode-ai">✨ Claude AI</span>' if use_claude else '<span class="mode-free">🆓 Tesseract — Gratuit</span>'
st.markdown(f"""
<div class="hero">
  <div class="badge">✦ OCR Intelligence Artificielle</div>
  <h1>🏦 BanqScan</h1>
  <p>
    Convertissez vos relevés bancaires scannés en fichiers lisibles par votre logiciel comptable.<br/>
    Tous formats d'image acceptés · Export PDF &amp; CSV · Compatible Sage, EBP, Cegid…
  </p>
  <div style="margin-top:12px;">{mode_badge}</div>
</div>
""", unsafe_allow_html=True)

# Steps
step = st.session_state.get("step", 1)
def scls(n): return "active" if step==n else ("done" if step>n else "")
st.markdown(f"""
<div class="step-row">
  <div class="step {scls(1)}">① Charger le relevé</div>
  <div class="step {scls(2)}">② Analyse OCR</div>
  <div class="step {scls(3)}">③ Résultats &amp; export</div>
</div>
""", unsafe_allow_html=True)

# ── Upload ──
col_up, col_tip = st.columns([3, 2])

with col_up:
    st.markdown('<div class="card-title">📁 Relevé(s) à convertir</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Formats acceptés : JPEG, JFIF, PNG, BMP, TIFF, WEBP, GIF, TGA, AVIF…",
        type=ACCEPTED_EXTENSIONS,
        accept_multiple_files=True,
    )

with col_tip:
    st.markdown("""
    <div class="card">
      <div class="card-title">💡 Conseils pour un bon résultat</div>
      <div style="font-size:.8rem;line-height:1.75;color:#4a5568;">
        📸 <b>Photo droite</b> — relevé bien à plat, sans angle<br/>
        💡 <b>Bonne lumière</b> — pas de reflets ni d'ombre<br/>
        🔍 <b>Résolution min. 150 DPI</b> recommandée<br/>
        📄 <b>Tous formats</b> acceptés : JPEG, JFIF, PNG, BMP, TIFF, WEBP…<br/><br/>
        <b>Mode Gratuit</b> : idéal pour relevés nets et bien contrastés<br/>
        <b>Mode Claude</b> : meilleur sur relevés complexes ou légèrement flous
      </div>
    </div>
    """, unsafe_allow_html=True)

if uploaded:
    st.session_state["step"] = 2

    # Preview thumbnails
    prev_cols = st.columns(min(len(uploaded), 4))
    for i, f in enumerate(uploaded):
        with prev_cols[i % 4]:
            try:
                img = open_image_universal(f.read(), f.name)
                f.seek(0)
                st.image(img, caption=f"{f.name} ({f.size//1024} Ko)", use_container_width=True)
            except Exception:
                st.info(f"📄 {f.name}")

    st.markdown("---")

    # Guard: Claude mode needs API key
    ready = True
    if use_claude and not api_key:
        st.markdown('<div class="alert-warn">⚠️ Entrez votre clé API Anthropic dans la barre latérale pour le mode Claude.</div>', unsafe_allow_html=True)
        ready = False

    if ready:
        btn_label = f"🔍 Analyser {len(uploaded)} fichier(s) — Mode {'Claude AI ✨' if use_claude else 'Gratuit 🆓'}"
        if st.button(btn_label, use_container_width=True):
            st.session_state["results"] = []
            st.session_state["step"] = 2

            prog = st.progress(0)
            status = st.empty()

            for idx, f in enumerate(uploaded):
                status.markdown(f"**⏳** Traitement `{f.name}` ({idx+1}/{len(uploaded)})…")
                prog.progress(idx / len(uploaded))

                raw = f.read()
                try:
                    t0 = time.time()
                    if use_claude:
                        data = ocr_claude(raw, f.name, api_key, logiciel)
                    else:
                        data = ocr_tesseract(raw, f.name)
                    elapsed = time.time() - t0

                    st.session_state["results"].append({
                        "filename": f.name,
                        "data": data,
                        "elapsed": elapsed,
                    })
                except json.JSONDecodeError:
                    st.error(f"❌ `{f.name}` — Réponse JSON invalide. Réessayez.")
                except Exception as e:
                    st.error(f"❌ `{f.name}` — {e}")

            prog.progress(1.0)
            nb = len(st.session_state.get("results", []))
            status.success(f"✅ Terminé — {nb} fichier(s) analysé(s) avec succès")
            st.session_state["step"] = 3

# ── Results ──
if st.session_state.get("results"):
    results = st.session_state["results"]
    st.markdown("---")
    st.markdown("## 📊 Résultats")

    for res in results:
        data  = res["data"]
        fname = res["filename"]
        ops   = data.get("operations", [])
        cpt   = data.get("compte", {})
        total_d = data.get("total_debits") or sum(o.get("debit",0) or 0 for o in ops)
        total_c = data.get("total_credits") or sum(o.get("credit",0) or 0 for o in ops)
        mode_label = "Claude AI" if data.get("_ocr_mode") == "claude" else "Tesseract (gratuit)"

        with st.expander(
            f"📄 {fname} — {cpt.get('titulaire','N/A')} — {len(ops)} opérations — {mode_label}",
            expanded=True
        ):
            # Compte info
            st.markdown(f"""
            <div class="compte-grid">
              <div class="compte-item"><div class="k">Titulaire</div><div class="v">{cpt.get('titulaire','—')}</div></div>
              <div class="compte-item"><div class="k">Banque</div><div class="v">{cpt.get('banque','—')}</div></div>
              <div class="compte-item"><div class="k">IBAN</div><div class="v" style="font-size:.78rem;font-family:monospace">{cpt.get('iban','—')}</div></div>
              <div class="compte-item"><div class="k">Période</div><div class="v" style="font-size:.82rem">{cpt.get('periode','—')}</div></div>
              <div class="compte-item"><div class="k">RIB</div><div class="v" style="font-size:.78rem">{cpt.get('rib','—')}</div></div>
              <div class="compte-item"><div class="k">Agence</div><div class="v">{cpt.get('agence','—')}</div></div>
            </div>
            """, unsafe_allow_html=True)

            balance = total_c - total_d
            bal_cls = "green" if balance >= 0 else "red"
            st.markdown(f"""
            <div class="stats-row">
              <div class="stat-box blue"><div class="label">Opérations</div><div class="value">{len(ops)}</div></div>
              <div class="stat-box red"><div class="label">Total débits</div><div class="value">{fmt_eur(total_d) or "0,00 €"}</div></div>
              <div class="stat-box green"><div class="label">Total crédits</div><div class="value">{fmt_eur(total_c) or "0,00 €"}</div></div>
              <div class="stat-box {bal_cls}"><div class="label">Solde mouvement</div><div class="value">{fmt_eur(abs(balance)) or "0,00 €"}</div></div>
            </div>
            """, unsafe_allow_html=True)

            # Operations table
            st.markdown('<div class="card-title" style="margin-top:8px;">📋 Détail des opérations</div>', unsafe_allow_html=True)
            rows_html = ""
            for op in ops:
                d  = op.get("debit",  0) or 0
                cr = op.get("credit", 0) or 0
                rows_html += f"""<tr>
                  <td class="date">{op.get('date_comptable','')}</td>
                  <td class="date">{op.get('date_valeur','')}</td>
                  <td>{op.get('libelle','')}</td>
                  <td class="debit">{fmt_eur(d)}</td>
                  <td class="credit">{fmt_eur(cr)}</td>
                </tr>"""

            st.markdown(f"""
            <div style="overflow-x:auto;border-radius:10px;border:1px solid #d9e4ef;">
            <table class="op-table">
              <thead><tr>
                <th>Date comptable</th><th>Date valeur</th>
                <th>Libellé</th>
                <th class="right">Débit</th><th class="right">Crédit</th>
              </tr></thead>
              <tbody>{rows_html}</tbody>
              <tfoot><tr>
                <td colspan="3"><b>TOTAL</b></td>
                <td class="debit"><b>{fmt_eur(total_d) or "0,00 €"}</b></td>
                <td class="credit"><b>{fmt_eur(total_c) or "0,00 €"}</b></td>
              </tr></tfoot>
            </table></div>
            """, unsafe_allow_html=True)

            # OCR warning for Tesseract
            if data.get("_ocr_mode") == "tesseract":
                st.markdown("""
                <div class="alert-info" style="margin-top:12px;font-size:.78rem;">
                ℹ️ <b>Mode Tesseract :</b> Vérifiez le tableau ci-dessus. Si des opérations manquent ou
                des montants sont inversés (débit/crédit), passez en <b>Mode Claude AI</b> pour une
                meilleure précision sur ce type de scan.
                </div>
                """, unsafe_allow_html=True)

            # Correction manuelle rapide
            with st.expander("✏️ Corriger manuellement (optionnel)", expanded=False):
                st.markdown("*Modifiez les données si l'OCR a commis des erreurs :*")
                edited = []
                for k, op in enumerate(ops):
                    c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 3.5, 1.2, 1.2])
                    edited.append({
                        "date_comptable": c1.text_input("Date compt.", op.get("date_comptable",""), key=f"dc_{fname}_{k}"),
                        "date_valeur":    c2.text_input("Date val.",   op.get("date_valeur",""),    key=f"dv_{fname}_{k}"),
                        "libelle":        c3.text_input("Libellé",     op.get("libelle",""),         key=f"lb_{fname}_{k}"),
                        "debit":          c4.number_input("Débit",     value=float(op.get("debit",0) or 0),  step=0.01, key=f"db_{fname}_{k}"),
                        "credit":         c5.number_input("Crédit",    value=float(op.get("credit",0) or 0), step=0.01, key=f"cr_{fname}_{k}"),
                    })
                if st.button("💾 Appliquer les corrections", key=f"save_{fname}"):
                    data["operations"] = edited
                    data["total_debits"]  = round(sum(o["debit"]  for o in edited), 2)
                    data["total_credits"] = round(sum(o["credit"] for o in edited), 2)
                    st.success("✅ Corrections appliquées — regénérez les exports ci-dessous.")

            # Downloads
            st.markdown("<br/>", unsafe_allow_html=True)
            dc1, dc2 = st.columns(2)
            stem = re.sub(r'\.[^.]+$', '', fname)

            if export_fmt in ("PDF + CSV", "PDF uniquement"):
                with dc1:
                    with st.spinner("Génération PDF…"):
                        pdf_bytes = build_pdf(data, logiciel)
                    st.download_button(
                        "⬇️ Télécharger le PDF",
                        data=pdf_bytes,
                        file_name=f"{stem}_converti.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            if export_fmt in ("PDF + CSV", "CSV uniquement"):
                with dc2:
                    csv_bytes = build_csv(data)
                    st.download_button(
                        "⬇️ Télécharger le CSV",
                        data=csv_bytes,
                        file_name=f"{stem}_converti.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

    st.markdown("---")
    if st.button("🔄 Nouvelle conversion"):
        for k in ("results", "step"):
            st.session_state.pop(k, None)
        st.rerun()

elif not uploaded:
    st.markdown("""
    <div style="text-align:center;padding:52px 24px;color:#8a9bb0;">
      <div style="font-size:3.5rem;margin-bottom:16px;">📄</div>
      <div style="font-size:1.05rem;font-weight:600;margin-bottom:8px;color:#4a5568;">
        Aucun fichier chargé
      </div>
      <div style="font-size:.85rem;max-width:400px;margin:auto;line-height:1.6;">
        Uploadez un ou plusieurs relevés scannés pour commencer.<br/>
        <b>Tous les formats d'image sont acceptés</b> : JPEG, JFIF, PNG, BMP, TIFF, WEBP…
      </div>
    </div>
    """, unsafe_allow_html=True)
