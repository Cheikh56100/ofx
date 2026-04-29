# -*- coding: utf-8 -*-
"""
OFX Bridge — Interface Streamlit
Convertisseur de relevés bancaires PDF vers OFX
Supporte : Qonto, LCL, CA, CE, BP, CIC, CM (Crédit Mutuel), CGD, LBP, SG, BNP, myPOS, Shine,
           CBAO, Ecobank, BCI, Coris, UBA, Orabank, BOA, ATB, BSIC, BIS, BNDE
           + Parseur universel automatique pour toute autre banque
"""

import io
import re
import hashlib
import logging
import tempfile
import os
from datetime import datetime
from pathlib import Path

import streamlit as st

# ── Logging minimal (Streamlit gère son propre stdout) ───────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(funcName)s — %(message)s")
logger = logging.getLogger("ofxbridge")

# ── Imports avec gestion d'erreur propre ─────────────────────────────────────
try:
    import pdfplumber
    _PDFPLUMBER_OK = True
except ImportError:
    _PDFPLUMBER_OK = False

_OCR_AVAILABLE = False
try:
    import pytesseract
    from pdf2image import convert_from_path
    _OCR_AVAILABLE = True
except ImportError:
    pass

try:
    from pydantic import BaseModel, field_validator
    _PYDANTIC_OK = True
except ImportError:
    _PYDANTIC_OK = False

# ── OCR via PyMuPDF (fitz) — fallback sans Tesseract ─────────────────────────
_FITZ_OK = False
try:
    import fitz  # PyMuPDF
    _FITZ_OK = True
except ImportError:
    pass

import base64
import json
import urllib.request

# ════════════════════════════════════════════════════════════════════════════
# MODÈLE PYDANTIC
# ════════════════════════════════════════════════════════════════════════════
if _PYDANTIC_OK:
    class Transaction(BaseModel):
        date:   str
        type:   str
        amount: float
        name:   str
        memo:   str = ""
        fitid:  str

        @field_validator('date')
        @classmethod
        def date_must_be_8_digits(cls, v):
            if not re.match(r'^\d{8}$', v):
                raise ValueError(f"Date OFX invalide: '{v}'")
            return v

        @field_validator('type')
        @classmethod
        def type_must_be_valid(cls, v):
            if v not in ('CREDIT', 'DEBIT'):
                raise ValueError(f"Type invalide: '{v}'")
            return v

        @field_validator('amount')
        @classmethod
        def amount_must_be_nonzero(cls, v):
            if v == 0.0:
                raise ValueError("Montant nul détecté")
            return v


# ════════════════════════════════════════════════════════════════════════════
# UTILITAIRES COMMUNS (identiques à la version Tkinter)
# ════════════════════════════════════════════════════════════════════════════

def extract_words_by_page(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_words(keep_blank_chars=False))
    return pages

def extract_text_by_page(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages

def parse_amount(s):
    s = re.sub(r'\s+', ' ', str(s))  # normalise \n \t \r en espace
    s = s.replace('\xa0','').replace(' ','').replace('*','').strip()
    if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', s):
        return float(s.replace('.','').replace(',','.'))
    if re.match(r'^\d+,\d{2}$', s):
        return float(s.replace(',','.'))
    if re.match(r'^\d+\.\d{2}$', s):
        return float(s)
    cleaned = re.sub(r'[^\d,.]', '', s)
    cleaned = cleaned.replace(',', '.')
    try:
        return float(cleaned)
    except ValueError:
        return None

def group_words_by_row(words, tol=3.0):
    if not words:
        return []
    rows, cur, top = [], [words[0]], words[0]['top']
    for w in words[1:]:
        if abs(w['top'] - top) <= tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda x: x['x0']))
            cur, top = [w], w['top']
    if cur:
        rows.append(sorted(cur, key=lambda x: x['x0']))
    return sorted(rows, key=lambda r: r[0]['top'])

def clean_label(s):
    return re.sub(r'\s+', ' ', s).strip()

def _is_technical_label(label):
    if not label:
        return True
    if re.match(r'^\d{6}\s+CB\*+\d+\s+\w+\s*$', label):
        return True
    if not re.search(r'[A-Za-zÀ-ÿ]{3,}', label):
        return True
    return False

def _is_human_readable(label):
    if not label:
        return False
    if re.search(r'[A-Z0-9]{15,}', label):
        return False
    if re.match(r'^[\d\s\-\/.,]+$', label):
        return False
    readable_words = [w for w in label.split() if re.search(r'[A-Za-zÀ-ÿ]{2,}', w)
                      and not re.match(r'^\d', w)]
    return len(readable_words) >= 2

def smart_label(main_label, memo_lines):
    label = clean_label(main_label)
    memos = [clean_label(m) for m in memo_lines if clean_label(m)]
    if _is_technical_label(label) and memos:
        for candidate in memos:
            if _is_human_readable(candidate):
                remaining = ' | '.join(m for m in memos if m != candidate and m)
                return candidate, (label + (' | ' + remaining if remaining else ''))
        return label, ' | '.join(memos)
    return label, ' | '.join(memos)

def make_fitid(date, label, amount):
    return hashlib.md5(f"{date}{label}{amount:.2f}".encode()).hexdigest()

def date_jjmm_to_ofx(jjmm, year):
    p = jjmm.replace('.', '/').split('/')
    if len(p) == 2:
        return f"{year}{p[1].zfill(2)}{p[0].zfill(2)}"
    return f"{year}0101"

def date_full_to_ofx(date_str):
    date_str = date_str.replace('.', '/')
    p = date_str.split('/')
    if len(p) == 3:
        return f"{p[2]}{p[1].zfill(2)}{p[0].zfill(2)}"
    return datetime.now().strftime('%Y%m%d')

def extract_iban(text):
    """
    Extrait le premier IBAN valide trouvé dans le texte.
    Stratégies (ordre de priorité) :
      1. Mot-clé IBAN + lecture ligne par ligne (évite la fusion avec le BIC)
      2. IBAN nu — format groupé par 4 sans mot-clé
      3. Fallback UEMOA/BCEAO sans mot-clé
    Tronque à la longueur IBAN réglementaire du pays (ex: FR=27, SN=28…).
    """
    # Longueurs IBAN officielles par code pays
    _IBAN_LEN = {
        'FR':27,'BE':16,'DE':22,'ES':24,'IT':27,'NL':18,'PT':25,'GB':22,
        'CH':21,'IE':22,'LU':20,'AT':20,'DK':18,'FI':18,'NO':15,'SE':24,
        'SN':28,'CI':28,'BJ':28,'TG':28,'ML':28,'BF':28,'NE':28,
        'GW':25,'GN':26,'MR':27,'CM':27,'MA':28,'TN':24,'DZ':24,
    }

    def _clean_and_truncate(raw):
        """Nettoie et tronque à la bonne longueur selon le code pays."""
        raw = re.sub(r'\s+', '', raw).upper()
        raw = re.sub(r'[^A-Z0-9]', '', raw)
        if len(raw) < 14 or not re.match(r'^[A-Z]{2}\d{2}', raw):
            return ''
        max_len = _IBAN_LEN.get(raw[:2], 34)
        return raw[:max_len]

    # Normaliser les espaces insécables
    text_clean = text.replace('\xa0', ' ').replace('\u202f', ' ')

    # ── 1. Avec mot-clé : traiter ligne par ligne pour éviter la fusion avec BIC ─
    # Le BIC est souvent sur la ligne suivante : "IBAN: FR76 ...\nBIC: QNTO..."
    for line in text_clean.split('\n'):
        line = line.strip()
        m = re.search(
            r'(?:IBAN|I\.?B\.?A\.?N\.?)\s*[:\-]?\s*'
            r'([A-Z]{2}[\s]?\d{2}[\s\dA-Z]{10,38})',
            line, re.IGNORECASE
        )
        if m:
            result = _clean_and_truncate(m.group(1))
            if result:
                return result

    # ── 2. IBAN nu sans mot-clé — format groupé par 4 (ex: FR76 3000 4028…) ───
    # On cherche ligne par ligne pour ne pas déborder sur la ligne suivante
    for line in text_clean.split('\n'):
        m2 = re.search(
            r'\b([A-Z]{2}\d{2}(?:\s?[A-Z0-9]{4}){3,7}(?:\s?[A-Z0-9]{1,4})?)\b',
            line.upper()
        )
        if m2:
            result = _clean_and_truncate(m2.group(1))
            if result:
                return result

    # ── 3. IBAN UEMOA/BCEAO avec mot-clé (peut contenir lettres dans le BBAN) ──
    # Ex : "IBAN : SN08 SN035 01010 0022015022 03" ou compact
    uemoa_cc = r'(?:SN|CI|BJ|TG|ML|BF|NE|GW|GN|MR|CM|MA|TN|DZ)'
    for line in text_clean.split('\n'):
        m3a = re.search(
            r'(?:IBAN|I\.?B\.?A\.?N\.?)\s*[:\-]?\s*'
            r'(' + uemoa_cc + r'[\s\dA-Z]{14,36})',
            line, re.IGNORECASE
        )
        if m3a:
            result = _clean_and_truncate(m3a.group(1))
            if result:
                return result

    # ── 4. Fallback UEMOA/BCEAO sans mot-clé (scan global) ───────────────────
    m4 = re.search(
        r'\b(' + uemoa_cc + r'\d{2}[A-Z0-9\s]{15,35})\b',
        text_clean.upper()
    )
    if m4:
        result = _clean_and_truncate(m4.group(1))
        if result:
            return result

    # ── 5. Numéro de compte RIB brut BCEAO (si aucun IBAN trouvé) ────────────
    # Format BCEAO numérique pur : ex "SN011 01005 005000458982 90"
    # Certains PDF africains écrivent le RIB sans mentionner "IBAN"
    m5 = re.search(
        r'\b([A-Z]{2}\d{3,5})\s+(\d{4,6})\s+(\d{8,14})\s+(\d{2})\b',
        text_clean.upper()
    )
    if m5:
        raw = ''.join(m5.groups())
        if len(raw) >= 14 and re.match(r'^[A-Z]{2}', raw):
            result = _clean_and_truncate(raw)
            if result:
                return result

    return ''

# Codes pays UEMOA → (longueur IBAN, longueur compte)
_UEMOA_IBAN = {
    # Format BCEAO : CC KK BBB AAAAA NNNNNNNNNNNN CC
    # CC=pays(2) KK=clé(2) BBB=banque(3-5) AAAAA=agence(4-5) N=compte(11-12) CC=clé RIB(2)
    'SN': 28,  # Sénégal
    'CI': 28,  # Côte d'Ivoire
    'BJ': 28,  # Bénin
    'TG': 28,  # Togo
    'ML': 28,  # Mali
    'BF': 28,  # Burkina Faso
    'NE': 28,  # Niger
    'GW': 25,  # Guinée-Bissau
    'GN': 26,  # Guinée
    'MR': 27,  # Mauritanie
    'CM': 27,  # Cameroun
    'MA': 28,  # Maroc
    'TN': 24,  # Tunisie
    'DZ': 24,  # Algérie
}

def iban_to_rib(iban, info=None):
    """
    Décompose un IBAN en (banque, agence, compte) pour l'OFX.
    Priorité : champs _rib_* extraits directement du PDF (via _afr_header).
    Sinon : France (FR27), UEMOA/BCEAO (SN28…), fallback numérique.
    """
    # ── Priorité : RIB directement extrait du PDF ────────────────────────────
    if info and info.get('_rib_bank') and info.get('_rib_account'):
        return (info['_rib_bank'],
                info.get('_rib_agency', '00000'),
                info['_rib_account'])

    c = iban.replace(' ', '').upper()
    c = re.sub(r'[^A-Z0-9]', '', c)   # purge tout caractère invalide

    # ── France ──────────────────────────────────────────────────────────────
    if c.startswith('FR') and len(c) == 27:
        r = c[4:]
        return r[0:5], r[5:10], r[10:21]

    # ── UEMOA/BCEAO ─────────────────────────────────────────────────────────
    country = c[:2]
    if country in _UEMOA_IBAN and len(c) >= 20:
        bban = c[4:]   # tout après CC+KK
        # Format BCEAO : BB(5 alphan.) + Agence(5 num.) + Compte(11-12 num.) + [Cle(2)]
        # Certains IBAN BSIC font 26 chars sans cle RIB -> ne pas tronquer avec [-2]
        code_banque = bban[0:5]
        agence      = bban[5:10]
        expected_len = _UEMOA_IBAN.get(country, 28)
        iban_has_key = len(c) >= expected_len
        compte_raw = bban[10:-2] if iban_has_key else bban[10:]
        compte     = re.sub(r'[A-Z]', '', compte_raw)
        return code_banque, agence, compte

    # ── Fallback numérique ───────────────────────────────────────────────────
    # Si on a un code pays UEMOA mais que l'IBAN était trop court pour le bloc
    # ci-dessus, on retente avec le BBAN brut (lettres conservées pour code banque)
    if country in _UEMOA_IBAN and len(c) >= 12:
        bban = c[4:]
        code_banque = bban[0:5]
        agence      = re.sub(r'[^0-9]', '', bban[5:10]) or '00000'
        expected_len = _UEMOA_IBAN.get(country, 28)
        iban_has_key = len(c) >= expected_len
        compte_raw  = bban[10:-2] if iban_has_key else bban[10:]
        compte      = re.sub(r'[A-Z]', '', compte_raw)
        if code_banque:
            return code_banque, agence, compte
    digits = re.sub(r'[^0-9]', '', c)
    if len(digits) >= 15:
        return digits[0:5], digits[5:10], digits[10:22]
    if len(digits) >= 5:
        return digits[0:5], '00000', digits[5:]
    return '00000', '00000', c[:20] if c else '00000'

def _year_from_text(text):
    m = re.search(r'\b(20\d{2})\b', text)
    return int(m.group(1)) if m else datetime.now().year

def _parse_col_amount(words):
    if not words:
        return None
    full = ' '.join(w['text'] for w in words).replace('\xa0', ' ').strip()
    if full in ('.', ',', ''):
        return None
    m = re.search(r'(\d{1,3}(?:[.\s]\d{3})+,\d{2})', full)
    if m:
        val = parse_amount(m.group(1).replace(' ', '.'))
        if val is not None and val > 0:
            return val
    m2 = re.search(r'(\d+,\d{2})', full)
    if m2:
        val = parse_amount(m2.group(1))
        if val is not None and val > 0:
            return val
    return None

def _parse_signed_amount(words):
    if not words:
        return None
    full = ' '.join(w['text'] for w in words).replace('\xa0', ' ').strip()
    m = re.search(r'([+\-])\s*([\d\s]+[,.][\d]{2})', full)
    if m:
        sign = 1.0 if m.group(1) == '+' else -1.0
        val = parse_amount(m.group(2))
        if val is not None:
            return sign * val
    m2 = re.search(r'([\d\s]+[,.][\d]{2})', full)
    if m2:
        val = parse_amount(m2.group(1))
        if val is not None:
            return val
    return None

def _make_txn(date_ofx, amount, label, memo=''):
    txn_dict = {
        'date':   date_ofx,
        'type':   'CREDIT' if amount >= 0 else 'DEBIT',
        'amount': amount,
        'name':   clean_label(label)[:64],
        'memo':   clean_label(memo)[:128],
        'fitid':  make_fitid(date_ofx, label, amount)
    }
    if _PYDANTIC_OK:
        try:
            Transaction(**txn_dict)
        except Exception as exc:
            logger.warning("Transaction ignorée [%s | %s | %.2f] : %s",
                           date_ofx, label[:40], amount, exc)
            return None
    return txn_dict

def _pdf_has_text(pages_text, min_chars=300):
    total = sum(len(p.strip()) for p in pages_text)
    return total >= min_chars

def _ocr_pdf(pdf_path):
    if not _OCR_AVAILABLE:
        raise RuntimeError(
            "Ce PDF semble scanné (aucun texte extractible). "
            "Les outils OCR (pytesseract, pdf2image, Tesseract) ne sont pas installés sur ce serveur."
        )
    images = convert_from_path(pdf_path, dpi=300)
    return [pytesseract.image_to_string(img, lang='fra+eng') for img in images]


# ════════════════════════════════════════════════════════════════════════════
# DÉTECTION DE LA BANQUE
# ════════════════════════════════════════════════════════════════════════════

def detect_bank(pages_text):
    text = pages_text[0][:3000].upper()
    if 'QONTO' in text or 'QNTOFRP' in text:
        return 'QONTO'
    if 'CREDIT LYONNAIS' in text or ('LCL' in text and 'RELEVE DE COMPTE COURANT' in text):
        return 'LCL'
    text_nospace = text.replace(' ', '')
    if ('SOCIETE GENERALE' in text or 'SOCIÉTÉ GÉNÉRALE' in text
            or 'SOCIETEGENERALE' in text_nospace) and (
            'SENEGAL' in text or 'SÉNÉGAL' in text or 'COTE D' in text
            or "CÔTE D'" in text or 'CAMEROUN' in text or 'DAKAR' in text
            or 'ABIDJAN' in text or 'DOUALA' in text or 'LOME' in text
            or 'BAMAKO' in text):
        return 'SG_AFRIQUE'
    if ('SOCIETE GENERALE' in text or 'SOCIÉTÉ GÉNÉRALE' in text
            or '552 120 222' in text or 'SOCIETEGENERALE' in text_nospace
            or 'SG.FR' in text or 'PROFESSIONNELS.SG.FR' in text):
        return 'SG'
    if 'CREDIT AGRICOLE' in text or 'AGRIFRPP' in text:
        return 'CA'
    if 'CAIXA GERAL' in text or 'CGDIFRPP' in text or 'CGD' in text[:500]:
        return 'CGD'
    if "CAISSE D'EPARGNE" in text or "CAISSE D.EPARGNE" in text or 'CEPAFRPP' in text:
        return 'CE'
    if 'BANQUE POPULAIRE' in text or 'CCBPFRPP' in text:
        return 'BP'
    if 'BANQUE POSTALE' in text or 'PSSTFRPP' in text or 'LABANQUEPOSTALE' in text:
        return 'LBP'
    if 'CREDIT INDUSTRIEL' in text or 'CMCIFRPP' in text or ('CIC' in text and 'RELEVE' in text):
        return 'CIC'
    if ('CREDIT MUTUEL' in text or 'CRÉDIT MUTUEL' in text
            or 'CMCIFR2A' in text or 'CREDITMUTUEL' in text_nospace
            or 'CAISSE DE CREDIT MUTUEL' in text):
        return 'CM'
    if ('BNP PARIBAS' in text or 'BNPAFRPP' in text or 'BNP' in text[:500]
            or 'BANQUE NATIONALE DE PARIS' in text):
        return 'BNP'
    if 'MYPOS' in text or 'MYPOS LTD' in text or 'MY POS' in text:
        return 'MYPOS'
    if ('SNNNFR22XXX' in text or 'SHINE.FR' in text or 'SHINE SAS' in text
            or ('SHINE' in text and ('RELEVE' in text or 'SNNN' in text or '1741' in text))):
        return 'SHINE'
    if 'NSIA BANQUE' in text or ('NSIA' in text and ('RELEVE DE COMPTE' in text or 'SOLDE DEBUT' in text)):
        return 'NSIA'
    # Détection NSIA par structure du relevé (le logo NSIA est souvent une image)
    if ('SOLDE DEBUT' in text or 'SOLDE DÉBUT' in text) and 'MOUV' in text and (
            'MOUV. DÉBIT' in text or 'MOUV. DEBIT' in text or
            'MOV. DEBIT' in text or 'NOMBRE DEBIT' in text or 'NOMBRE CRÉDIT' in text):
        return 'NSIA'
    if 'CBAO' in text or 'COMPAGNIE BANCAIRE DE L' in text:
        return 'CBAO'
    if 'ECOBANK' in text or 'ECOBANK SENEGAL' in text:
        return 'ECOBANK'
    if 'BANQUE POUR LE COMMERCE' in text and 'INDUSTRIE' in text:
        return 'BCI'
    if 'CORIS BANK' in text or 'CORISBANK' in text_nospace:
        return 'CORIS'
    if 'UNITED BANK FOR AFRICA' in text or 'UNAFSNDA' in text or ('UBA' in text[:400] and 'BANK' in text):
        return 'UBA'
    if 'ORABANK' in text:
        return 'ORABANK'
    if 'BANK OF AFRICA' in text:
        return 'BOA'
    if 'ARAB TUNISIAN BANK' in text:
        return 'ATB'
    if ('BSIC' in text or 'BANQUE SAHELO' in text or 'SN08SN111' in text_nospace):
        return 'BSIC'
    if ('BANQUE ISLAMIQUE DU SENEGAL' in text or 'ISLAMIQUE' in text and 'SENEGAL' in text):
        return 'BIS'
    if 'BNDE' in text or 'BANQUE NATIONALE POUR LE DEVELOPPEMENT' in text:
        return 'BNDE'
    # Détection BNDE par code banque IBAN (SN08SN169...) ou structure colonnes
    if ('SN08SN169' in text_nospace or 'SN169' in text_nospace) and 'EXTRAIT DE COMPTE' in text:
        return 'BNDE'
    if ('DÉBIT (XOF)' in text or 'DEBIT (XOF)' in text) and ('CRÉDIT (XOF)' in text or 'CREDIT (XOF)' in text) and 'EXTRAIT DE COMPTE' in text:
        return 'BNDE'
    return 'UNIVERSAL'


# ════════════════════════════════════════════════════════════════════════════
# PARSEURS BANCAIRES (identiques à la version Tkinter — non modifiés)
# ════════════════════════════════════════════════════════════════════════════

def parse_qonto(pages_words, pages_text):
    info = _extract_qonto_header(pages_text[0])
    year = int(info['period_start'].split('/')[2]) if info.get('period_start') else _year_from_text(pages_text[0])
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _qonto_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 130 <= w['x0'] < 410).strip()
            amount = _qonto_amount(row)
            memo = ''
            j = i + 1
            while j < len(rows) and not _qonto_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 130 <= w['x0'] < 410).strip()
                na = _qonto_amount(rows[j])
                if na is not None and amount is None:
                    amount = na; memo = nl; j += 1; break
                elif na is None and nl:
                    memo = nl; j += 1; break
                else:
                    break
            i = j
            if amount is None or not label or label in ('Transactions', 'Date de valeur'):
                continue
            memo_clean = memo if memo.strip() not in ('', '-', '+') else ''
            name, memo_out = smart_label(label, [memo_clean] if memo_clean else [])
            txns.append(_make_txn(date_jjmm_to_ofx(date_str, year), amount, name, memo_out))
    return info, [t for t in txns if t is not None]

def _qonto_date(row):
    for w in row:
        if w['x0'] < 120 and re.match(r'^\d{2}/\d{2}$', w['text']):
            return w['text']
    return ''

def _qonto_amount(row):
    aw = [w for w in row if w['x0'] >= 400]
    if not aw: return None
    full = ' '.join(w['text'] for w in aw).replace('EUR','').replace('\xa0',' ').strip()
    m = re.search(r'([+\-])\s*([\d\s]+[.,]\d{2})', full)
    if m:
        sign = 1.0 if m.group(1)=='+' else -1.0
        try: return sign * float(m.group(2).replace(' ','').replace(',','.'))
        except: pass
    m2 = re.search(r'([\d\s]+[.,]\d{2})', full)
    if m2:
        sign = 1.0
        for w in aw:
            if w['text'] in ('+','-'): sign = 1.0 if w['text']=='+' else -1.0; break
            sm = re.match(r'^([+\-])([\d,.]+)$', w['text'])
            if sm: sign = 1.0 if sm.group(1)=='+' else -1.0; break
        try: return sign * float(m2.group(1).replace(' ','').replace(',','.'))
        except: pass
    return None

def _extract_qonto_header(text):
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    m = re.search(r'Du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
    if m: info['period_start'], info['period_end'] = m.group(1), m.group(2)
    bals = re.findall(r'Solde au \d{2}/\d{2}\s*[+\-]\s*([\d]+\.[\d]{2})\s*EUR', text)
    if len(bals) >= 1: info['balance_open']  = float(bals[0])
    if len(bals) >= 2: info['balance_close'] = float(bals[-1])
    return info

def parse_lcl(pages_words, pages_text):
    info = _extract_lcl_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []

    # Labels à ignorer absolument (lignes de solde, totaux, en-têtes)
    SKIP_LABELS = {
        'DEBIT', 'CREDIT', 'VALEUR', 'DATE', 'LIBELLE', 'ANCIEN SOLDE',
        'SOLDE EN EUROS', 'TOTAUX', 'NOUVEAU SOLDE', 'SOLDE INITIAL',
        'SOLDE FINAL', 'TOTAL', 'REPORT', 'A REPORTER',
    }
    # Préfixes de labels de solde ou total (comparaison startswith)
    SKIP_PREFIXES = ('SOLDE', 'TOTAUX', 'TOTAL', 'ANCIEN', 'NOUVEAU')

    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _lcl_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 360).strip()

            # Ignorer les lignes de solde/totaux même si elles ont une date
            label_up = label.upper().strip()
            if (label_up in SKIP_LABELS
                    or any(label_up.startswith(p) for p in SKIP_PREFIXES)
                    or not label_up):
                i += 1; continue

            debit_words = [w for w in row if 360 <= w['x0'] < 490
                           and not re.match(r'^\d{2}\.\d{2}(\.\d{2,4})?$', w['text'])]
            debit_amt  = _parse_col_amount(debit_words)
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 490])

            memo = ''
            j = i + 1
            while j < len(rows) and not _lcl_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 70 <= w['x0'] < 360).strip()
                nl_up = nl.upper().strip()
                if nl and nl_up not in SKIP_LABELS and not any(nl_up.startswith(p) for p in SKIP_PREFIXES):
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j

            date_ofx = date_jjmm_to_ofx(date_str, year)
            name, memo_out = smart_label(label, [memo] if memo else [])
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo_out))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo_out))
    return info, [t for t in txns if t is not None]

def _lcl_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}\.\d{2}$', w['text']):
            return w['text']
    return ''

def _extract_lcl_header(pages_text):
    text = ' '.join(pages_text)
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)

    # Période : "du 03.10.2025 au 31.10.2025"
    m = re.search(r'du\s+(\d{2}\.\d{2}\.\d{4})\s+au\s+(\d{2}\.\d{2}\.\d{4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1).replace('.','/')
        info['period_end']   = m.group(2).replace('.','/')

    # Solde d'ouverture : "ANCIEN SOLDE  40 978,70" ou "ANCIEN SOLDE 40978,70"
    m_open = re.search(r'ANCIEN SOLDE\s+([\d\s]+[,\.]\d{2})', text)
    if m_open:
        v = parse_amount(m_open.group(1).replace(' ', ''))
        if v: info['balance_open'] = v

    # Solde de clôture : "SOLDE EN EUROS  46 862,54" (dernière ligne du relevé)
    for m_close in re.finditer(r'SOLDE EN EUROS\s+([\d\s]+[,\.]\d{2})', text):
        v = parse_amount(m_close.group(1).replace(' ', ''))
        if v: info['balance_close'] = v

    return info

def _ca_parse_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max and re.match(r'^\d', w['text'])]
    if not col: return None
    last = col[-1]['text']
    if not re.match(r'^\d+,\d{2}$', last): return None
    if len(col) == 1: return parse_amount(last)
    prefix_tokens = [w['text'] for w in col[:-1]]
    if all(re.match(r'^\d+$', p) for p in prefix_tokens):
        try:
            return float(''.join(prefix_tokens) + last.replace(',', '.'))
        except ValueError:
            pass
    return None

def parse_ca(pages_words, pages_text):
    info = _extract_ca_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'Débit','Crédit','Date','Libellé','Total des opérations','Nouveau solde',
            'opé.','valeur','Libellé des opérations','Ancien solde débiteur','Nouveau solde débiteur'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _ca_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 420).strip()
            debit_amt  = _ca_parse_zone(row, 415, 490)
            credit_amt = _ca_parse_zone(row, 490, 560)
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _ca_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 70 <= w['x0'] < 420).strip()
                if nl and nl not in SKIP and len(nl) > 1:
                    memo_parts.append(nl)
                j += 1
            i = j
            if not label or label in SKIP: continue
            date_ofx = date_jjmm_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _ca_date(row):
    for w in row:
        if w['x0'] < 50 and re.match(r'^\d{2}\.\d{2}$', w['text']):
            return w['text']
    return ''

def _extract_ca_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    mois_map = {'janvier':'01','février':'02','mars':'03','avril':'04','mai':'05','juin':'06',
                'juillet':'07','août':'08','septembre':'09','octobre':'10','novembre':'11','décembre':'12'}
    m = re.search(r'Date d.arrêté\s*:\s*(\d+)\s+(\w+)\s+(\d{4})', text)
    if m:
        mn = mois_map.get(m.group(2).lower(), '01')
        info['period_end']   = f"{m.group(1).zfill(2)}/{mn}/{m.group(3)}"
        info['period_start'] = f"01/{mn}/{m.group(3)}"
    return info

def parse_ce(pages_words, pages_text):
    info = _extract_ce_header(pages_text)
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _ce_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 155 <= w['x0'] < 500).strip()
            amount = _parse_signed_amount([w for w in row if w['x0'] >= 500])
            memo = ''
            j = i + 1
            while j < len(rows) and not _ce_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 155 <= w['x0'] < 500).strip()
                if nl and len(nl) > 2:
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j
            if not label or amount is None: continue
            skip_kw = {'DATE','VALEUR','MONTANT','OPERATIONS','SOLDE','TOTAL','DETAIL'}
            if any(s in label.upper() for s in skip_kw): continue
            date_ofx = date_full_to_ofx(date_str)
            name, memo_out = smart_label(label, [memo] if memo else [])
            txns.append(_make_txn(date_ofx, amount, name, memo_out))
    return info, [t for t in txns if t is not None]

def _ce_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
            return w['text']
    return ''

def _extract_ce_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    return info

def parse_bp(pages_words, pages_text):
    info = _extract_bp_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP_KW = {'DATE','LIBELLE','REFERENCE','COMPTA','VALEUR','MONTANT','SOLDE','TOTAL','DETAIL','OPERATION'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        skip_from = None
        for idx, row in enumerate(rows):
            row_text = ' '.join(w['text'] for w in row).upper()
            if 'DETAIL DE VOS MOUVEMENTS SEPA' in row_text or 'DETAIL DE VOS PRELEVEMENTS SEPA' in row_text:
                skip_from = idx; break
        i = 0
        while i < len(rows):
            if skip_from is not None and i >= skip_from: break
            row = rows[i]
            date_str = _bp_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 80 <= w['x0'] < 355).strip()
            amount = _bp_amount([w for w in row if w['x0'] >= 490])
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _bp_date(rows[j]):
                if skip_from is not None and j >= skip_from: break
                nl = ' '.join(w['text'] for w in rows[j] if 80 <= w['x0'] < 355).strip()
                if nl and len(nl) > 2 and not re.match(r'^[\d\s.,€%=\-EUR]+$', nl):
                    memo_parts.append(nl)
                j += 1
            i = j
            if not label or amount is None: continue
            if any(s in label.upper() for s in SKIP_KW): continue
            date_ofx = date_jjmm_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            txns.append(_make_txn(date_ofx, amount, name, memo))
    return info, [t for t in txns if t is not None]

def _bp_date(row):
    for w in row:
        if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}$', w['text']):
            return w['text']
    return ''

def _bp_amount(words):
    if not words: return None
    full = ' '.join(w['text'] for w in words).replace('€','').replace('\xa0',' ').strip()
    m = re.search(r'-\s*([\d\s]+[,.][\d]{2})', full)
    if m:
        try: return -abs(float(m.group(1).replace(' ','').replace(',','.')))
        except: pass
    m2 = re.search(r'\+\s*([\d\s]+[,.][\d]{2})', full)
    if m2:
        try: return abs(float(m2.group(1).replace(' ','').replace(',','.')))
        except: pass
    m3 = re.search(r'([\d\s]+[,.][\d]{2})', full)
    if m3:
        try:
            val = float(m3.group(1).replace(' ','').replace(',','.'))
            return val if val > 0 else None
        except: pass
    return None

def _extract_bp_header(pages_text):
    text = pages_text[0] if pages_text else ''
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    return info

def parse_cic(pages_words, pages_text):
    info = _extract_cic_header(pages_text)
    txns = []
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _cic_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 140 <= w['x0'] < 430).strip()
            debit_amt  = _parse_col_amount([w for w in row if 420 <= w['x0'] < 500])
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 500])
            memo = ''
            j = i + 1
            while j < len(rows) and not _cic_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 140 <= w['x0'] < 430).strip()
                if nl and len(nl) > 2 and not re.match(r'^[\d.,]+$', nl):
                    memo = (memo + ' ' + nl).strip()
                j += 1
            i = j
            if not label: continue
            skip_kw = {'DATE','DÉBIT','CRÉDIT','EUROS','SOLDE CREDITEUR','CREDIT INDUSTRIEL','TOTAL DES MOUVEMENTS'}
            if any(s in label.upper() for s in skip_kw): continue
            date_ofx = date_full_to_ofx(date_str)
            name, memo_out = smart_label(label, [memo] if memo else [])
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo_out))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo_out))
    return info, [t for t in txns if t is not None]

def _cic_date(row):
    for w in row:
        if w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
            return w['text']
    return ''

def _extract_cic_header(pages_text):
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    for pt in reversed(pages_text):
        iban = extract_iban(pt)
        if iban:
            info['iban'] = iban; break
    return info

def parse_cgd(pages_words, pages_text):
    info = _extract_cgd_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'A REPORTER','REPORT','TOTAL','NOUVEAU','ANCIEN','SARL','CPT ORD'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (len(row) >= 2
                    and re.match(r'^\d{2}$', row[0]['text']) and row[0]['x0'] < 50
                    and re.match(r'^\d{2}$', row[1]['text']) and row[1]['x0'] < 55):
                i += 1; continue
            dd, mm = row[0]['text'], row[1]['text']
            label = ' '.join(w['text'] for w in row if 70 <= w['x0'] < 310).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _cgd_amount_in_zone(row, 395, 500)
            credit_amt = _cgd_amount_in_zone(row, 500, 570)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if len(r2) >= 2 and re.match(r'^\d{2}$', r2[0]['text']) and r2[0]['x0'] < 50: break
                nl = ' '.join(w['text'] for w in r2 if 70 <= w['x0'] < 310).strip()
                if nl: memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = f"{year}{mm.zfill(2)}{dd.zfill(2)}"
            name, memo = smart_label(label, memo_parts)
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _cgd_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max and re.match(r'^\d', w['text'])]
    if not col: return None
    return parse_amount(col[-1]['text'])

def _extract_cgd_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    return info

def parse_lbp(pages_words, pages_text):
    info = _extract_lbp_header(pages_text)
    year = _year_from_text(pages_text[0])
    txns = []
    SKIP = {'TOTAL DES','NOUVEAU SOLDE','ANCIEN SOLDE','VOS OPERATIONS','DATE OPERATION','SITUATION DU','PAGE'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (row[0]['x0'] < 60 and re.match(r'^\d{2}/\d{2}$', row[0]['text'])):
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 85 <= w['x0'] < 430).strip()
            label = re.sub(r'\(cid:\d+\)', '', label).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _lbp_amount_in_zone(row, 430, 500)
            credit_amt = _lbp_amount_in_zone(row, 500, 560)
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2[0]['x0'] < 60 and re.match(r'^\d{2}/\d{2}$', r2[0]['text']): break
                j += 1
            i = j
            date_ofx = f"{year}{row[0]['text'][3:5]}{row[0]['text'][:2]}"
            name, memo = smart_label(label, [])
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _lbp_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max and re.match(r'^\d', w['text'])]
    if not col: return None
    last = col[-1]['text']
    if not re.match(r'^\d+,\d{2}$', last): return None
    if len(col) == 1: return parse_amount(last)
    prefix_tokens = [w['text'] for w in col[:-1]]
    if all(re.match(r'^\d+$', p) for p in prefix_tokens):
        try: return float(''.join(prefix_tokens) + last.replace(',', '.'))
        except: pass
    return parse_amount(last)

def _extract_lbp_header(pages_text):
    text = re.sub(r'\(cid:\d+\)', ' ', ' '.join(pages_text[:2]))
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    return info

def parse_sg(pages_words, pages_text):
    info = _extract_sg_header(pages_text)
    txns = []
    SKIP = {'TOTAUX DES','NOUVEAU SOLDE','SOLDE PRECEDENT','PROGRAMME DE','RAPPEL DES','MONTANT CUMULE'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            if not (row[0]['x0'] < 45 and re.match(r'^\d{2}/\d{2}/\d{4}$', row[0]['text'])):
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 120 <= w['x0'] < 430).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _sg_amount_in_zone(row, 430, 510)
            credit_amt = _sg_amount_in_zone(row, 510, 570)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2[0]['x0'] < 45 and re.match(r'^\d{2}/\d{2}/\d{4}$', r2[0]['text']): break
                nl = ' '.join(w['text'] for w in r2 if 120 <= w['x0'] < 430).strip()
                if nl and not any(s in nl.upper() for s in ('TOTAUX','NOUVEAU','PROGRAMME','RAPPEL')):
                    memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = date_full_to_ofx(row[0]['text'])
            name, memo = smart_label(label, memo_parts)
            if debit_amt:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _uba_join_amount(words):
    """
    Reconstitue un montant XOF/EUR fragmente en plusieurs tokens pdfplumber.
    Gere les montants avec decimales : ['41', '295,00'] → 41295.0
    ET les montants entiers XOF     : ['15', '400']     → 15400.0
    ET les grands entiers           : ['2', '316', '032,00'] → 2316032.0
    Strategie :
      1. Cherche d'abord un pattern avec decimales (ex: '15 400,00')
      2. Sinon, si tous les tokens sont numeriques purs, les concatene
         en verifiant que c'est coherent (pas un solde, pas une date)
    """
    if not words:
        return None
    full = ' '.join(w['text'] for w in words).replace('\xa0', ' ').strip()
    if not full or full in ('.', ','):
        return None

    # --- Cas 0 : format BIS iMAL (virgule séparateur milliers, pas de décimales) ---
    # Ex: "2,362,500"  "660,000"  "1,350,412"
    m0 = re.match(r'^(\d{1,3}(?:,\d{3})+)$', full.replace(' ', ''))
    if m0:
        try:
            val = float(m0.group(1).replace(',', ''))
            if val > 0:
                return val
        except ValueError:
            pass

    # --- Cas 1 : montant avec decimales (EUR ou XOF) ---
    # Ex: "15 400,00"  "2 316 032,00"  "41,95"
    m = re.search(r'(\d[\d\s]*[\.,]\d{2})\b', full)
    if m:
        raw = m.group(1).strip()
        normalized = re.sub(r'\s+', '', raw).replace(',', '.')
        try:
            val = float(normalized)
            if val > 0:
                return val
        except ValueError:
            pass

    # --- Cas 2 : montant entier XOF (pas de decimales) ---
    # Tous les tokens doivent etre purement numeriques
    # Ex: ['15', '400']  ['369', '630']  ['5', '850']
    texts = [w['text'] for w in words]
    if all(re.match(r'^\d+$', t) for t in texts) and len(texts) >= 2:
        # Verifier que la concatenation donne un nombre raisonnable
        # Le 2e token et suivants doivent avoir exactement 3 chiffres
        # (separateur milliers) pour eviter de concatener une date valeur
        valid = True
        for t in texts[1:]:
            if len(t) != 3:
                valid = False
                break
        if valid:
            try:
                val = float(''.join(texts))
                if val > 0:
                    return val
            except ValueError:
                pass
    elif len(texts) == 1 and re.match(r'^\d+$', texts[0]):
        # Montant entier sur un seul token
        try:
            val = float(texts[0])
            if val > 0:
                return val
        except ValueError:
            pass

    return None

def _sg_amount_in_zone(row, x_min, x_max):
    col = [w for w in row if x_min <= w['x0'] < x_max]
    if not col: return None
    # Enlever le '*' (opérations exonérées de commission) avant de parser
    col_clean = [w for w in col if w['text'].strip() not in ('*', '')]
    if not col_clean: return None

    # Reconstituer la chaîne complète et tenter parse_amount directement
    # (gère correctement "1.082,92", "29.117,17", etc.)
    full = ' '.join(w['text'] for w in col_clean).replace('*', '').replace('\xa0', ' ').strip()
    # Format français groupé par 3 avec point comme séparateur de milliers : 1.082,92
    m_fr = re.search(r'(\d{1,3}(?:\.\d{3})+,\d{2})', full)
    if m_fr:
        v = parse_amount(m_fr.group(1))
        if v is not None and v > 0:
            return v
    # Format simple : 100,75  ou  312,48
    m_simple = re.search(r'(\d+,\d{2})', full)
    if m_simple:
        v = parse_amount(m_simple.group(1))
        if v is not None and v > 0:
            return v
    # Fallback : montants fragmentés en plusieurs tokens (ex: "5 850" → 5850)
    v = _uba_join_amount(col_clean)
    if v is not None:
        return v
    return None

def _extract_sg_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}

    # ── IBAN (présent sur certains relevés SG) ────────────────────────────────
    info['iban'] = extract_iban(text)

    # ── RIB SG : "n° 30003 03320 00020641644 69" ─────────────────────────────
    # Format : n° BBBBB GGGGG CCCCCCCCCCC KK  (Banque 5 + Guichet 5 + Compte 11 + Clé 2)
    # pdfplumber peut lire avec espaces OU en un seul token compact "n°30003033200002064164469"
    if not info['iban']:
        # Cas 1 : avec espaces
        rib_m = re.search(
            r'n[°o]\s*(\d{5})\s+(\d{5})\s+(\d{10,12})\s+(\d{2})\b', text)
        # Cas 2 : compact sans espaces (23 chiffres : 5+5+11+2)
        if not rib_m:
            rib_m = re.search(r'n[°o]\s*(\d{5})(\d{5})(\d{11})(\d{2})\b', text)
        if rib_m:
            banque  = rib_m.group(1)
            guichet = rib_m.group(2)
            compte  = rib_m.group(3)
            cle     = rib_m.group(4)
            info['_rib_bank']    = banque
            info['_rib_agency']  = guichet
            info['_rib_account'] = compte
            info['_rib_key']     = cle
            info['iban'] = f"{banque} {guichet} {compte} {cle}"

    # ── Période : "du 01/03/2026 au 31/03/2026" ──────────────────────────────
    # Tolérance aux espaces multiples et aux retours à la ligne (layout multi-col SG)
    m = re.search(
        r'du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})',
        text, re.IGNORECASE
    )
    if m:
        info['period_start'] = m.group(1)
        info['period_end']   = m.group(2)

    # ── Solde de clôture : "NOUVEAU SOLDE AU 31/03/2026 + 85.536,72" ─────────
    m_bal = re.search(
        r'NOUVEAU SOLDE\s+AU\s+\d{2}/\d{2}/\d{4}\s+[+\-]?\s*([\d\s]+[,\.]\d{2})',
        text, re.IGNORECASE
    )
    if m_bal:
        v = parse_amount(m_bal.group(1).replace(' ', '').replace('\xa0', ''))
        if v:
            info['balance_close'] = v

    return info

def parse_bnp(pages_words, pages_text):
    info = _extract_bnp_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []
    SKIP = {'DATE','LIBELLE','VALEUR','DEBIT','CREDIT','EUROS','SOLDE','TOTAL','OPERATIONS',
            'ANCIEN SOLDE','NOUVEAU SOLDE','VIREMENT RECU','RELEVE DE COMPTE'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _bnp_date(row)
            if not date_str:
                i += 1; continue
            label = ' '.join(w['text'] for w in row if 85 <= w['x0'] < 430).strip()
            if not label or any(s in label.upper() for s in SKIP):
                i += 1; continue
            debit_amt  = _parse_col_amount([w for w in row if 480 <= w['x0'] < 560])
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 560])
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _bnp_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 85 <= w['x0'] < 430).strip()
                if nl and len(nl) > 2:
                    memo_parts.append(nl)
                j += 1
            i = j
            date_ofx = _bnp_date_to_ofx(date_str, year)
            name, memo = smart_label(label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _bnp_date(row):
    for w in row:
        if w['x0'] < 80:
            if re.match(r'^\d{2}/\d{2}/\d{2}$', w['text']): return w['text']
            if re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']): return w['text']
    return ''

def _bnp_date_to_ofx(date_str, year_hint):
    parts = date_str.split('/')
    if len(parts) == 3:
        dd, mm, yy = parts[0].zfill(2), parts[1].zfill(2), parts[2]
        if len(yy) == 2:
            full_year = (2000 + int(yy)) if int(yy) <= 30 else (1900 + int(yy))
        else:
            full_year = int(yy)
        return f"{full_year}{mm}{dd}"
    return str(year_hint) + '0101'

def _extract_bnp_header(pages_text):
    text = ' '.join(pages_text[:2])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)
    return info

def parse_mypos(pages_words, pages_text):
    info = _extract_mypos_header(pages_text)
    txns = []
    full_text = '\n'.join(pages_text)
    lines = [l.strip() for l in full_text.splitlines()]
    txn_re = re.compile(
        r'^(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}\s+'
        r'(System Fee|myPOS Payment|Glass Payment|Outgoing bank transfer|POS Payment|Mobile)\s*'
        r'.*?1\.0000\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*$'
    )
    for idx, line in enumerate(lines):
        m = txn_re.match(line)
        if not m: continue
        date_raw = m.group(1)
        txn_type = m.group(2).strip()
        try:
            debit_val  = float(m.group(3).replace(',', ''))
            credit_val = float(m.group(4).replace(',', ''))
        except ValueError:
            continue
        date_ofx = f"{date_raw[6:10]}{date_raw[3:5]}{date_raw[0:2]}"
        description = ''
        for back in (1, 2):
            if idx >= back:
                prev = lines[idx - back].strip()
                if prev and not re.match(r'^\d{2}\.\d{2}\.\d{4}', prev):
                    description = prev; break
        if txn_type == 'System Fee':
            name, memo = 'myPOS Fee', description
        else:
            name, memo = description or txn_type, description
        amount = -debit_val if debit_val > 0 else (credit_val if credit_val > 0 else None)
        if amount is None: continue
        txns.append(_make_txn(date_ofx, amount, name[:64], memo[:128]))
    return info, [t for t in txns if t is not None]

def _extract_mypos_header(pages_text):
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    text = pages_text[0] if pages_text else ''
    m = re.search(r'IBAN\s*:?\s*(IE\d{2}[A-Z0-9]+)', text)
    if m: info['iban'] = m.group(1).replace(' ','')
    m2 = re.search(r'Monthly statement\s*[-–]\s*(\d{2})\.(\d{4})', text, re.IGNORECASE)
    if m2:
        import calendar
        month, year = m2.group(1), m2.group(2)
        last_day = calendar.monthrange(int(year), int(month))[1]
        info['period_start'] = f"01/{month}/{year}"
        info['period_end']   = f"{last_day:02d}/{month}/{year}"
    return info

def parse_shine(pages_words, pages_text):
    info = _extract_shine_header(pages_text)
    txns = []
    SKIP = {'DATE','TYPE','OPÉRATION','OPERATION','DÉBIT','DEBIT','CRÉDIT','CREDIT',
            '(EURO)','SOLDE','TOTAL','NOUVEAU','COMMISSIONS','MOUVEMENTS','PAGE','LES','RELEVÉ'}
    for pw in pages_words:
        rows = group_words_by_row(pw)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _shine_date(row)
            if not date_str:
                i += 1; continue
            txn_type = ' '.join(w['text'] for w in row if 95 <= w['x0'] < 160).strip()
            label    = ' '.join(w['text'] for w in row if 160 <= w['x0'] < 453).strip()
            debit_amt  = _parse_col_amount([w for w in row if 453 <= w['x0'] < 513])
            credit_amt = _parse_col_amount([w for w in row if w['x0'] >= 513])
            memo_parts = []
            j = i + 1
            while j < len(rows) and not _shine_date(rows[j]):
                nl = ' '.join(w['text'] for w in rows[j] if 95 <= w['x0'] < 453).strip()
                if nl and len(nl) > 2:
                    memo_parts.append(nl)
                j += 1
            i = j
            full_label = (txn_type + ' ' + label).strip() if txn_type else label
            if any(s in full_label.upper() for s in SKIP) or len(full_label) < 2: continue
            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(full_label, memo_parts)
            if debit_amt is not None:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt is not None:
                txns.append(_make_txn(date_ofx,  credit_amt, name, memo))
    return info, [t for t in txns if t is not None]

def _shine_date(row):
    for w in row:
        if w['x0'] < 60 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
            return w['text']
    return ''

def _extract_shine_header(pages_text):
    text = ' '.join(pages_text[:3])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    info['iban'] = extract_iban(text)

    # Période : "De 01/01/2026 à 31/01/2026"
    m = re.search(r'De\s+(\d{2}/\d{2}/\d{4})\s+[àa]\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
    if m:
        info['period_start'] = m.group(1)
        info['period_end']   = m.group(2)

    # Solde d'ouverture : "Solde au JJ/MM/AAAA X,XX €"
    m_open = re.search(r'Solde\s+au\s+\d{2}/\d{2}/\d{4}\s+([\d\s]+[,\.]\d{2})', text, re.IGNORECASE)
    if m_open:
        v = parse_amount(m_open.group(1))
        if v is not None: info['balance_open'] = v

    # Nouveau solde (clôture)
    m_close = re.search(r'Nouveau\s+solde\s+([\d\s]+[,\.]\d{2})', text, re.IGNORECASE)
    if m_close:
        v = parse_amount(m_close.group(1))
        if v is not None: info['balance_close'] = v
    elif m_open:
        # Si pas de "Nouveau solde", solde de clôture = solde d'ouverture (mois vide)
        info['balance_close'] = info['balance_open']

    # Détecter explicitement un mois sans mouvement pour produire un OFX valide
    if re.search(r'Total\s+des\s+mouvements\s+0[,\.]00\s+0[,\.]00', text, re.IGNORECASE):
        info['_empty_period'] = True

    return info


# ════════════════════════════════════════════════════════════════════════════
# PARSEUR UNIVERSEL
# ════════════════════════════════════════════════════════════════════════════

_COL_SYNONYMS = {
    'date':   ['date','date opé','date opé.','date opération','date val','date valeur','date comptable','valeur','jour','date op'],
    'label':  ['libellé','libelle','opération','operation','description','motif','désignation','nature','détail','detail','mouvement','intitulé','label','wording','particulars','narration'],
    'debit':  ['débit','debit','débit (euro)','debit (euro)','sorties','sortie','retrait','retraits','paiements','débit fcfa','débit xof','withdrawals','withdrawal','payments','dr','déb','deb'],
    'credit': ['crédit','credit','crédit (euro)','credit (euro)','entrées','entrée','versement','versements','encaissements','crédit fcfa','crédit xof','deposits','deposit','receipts','cr','cré','cred'],
    'amount': ['montant','amount','somme','mouvement','débit/crédit','debit/credit','montant net','net'],
    'balance':['solde','balance','solde après','running balance'],
}

_DATE_PATTERNS = [
    (r'^(\d{2})[/\-\.](\d{2})[/\-\.](\d{4})$', 'dmy4'),
    (r'^(\d{4})[/\-\.](\d{2})[/\-\.](\d{2})$', 'ymd4'),
    (r'^(\d{2})[/\-\.](\d{2})[/\-\.](\d{2})$', 'dmy2'),
    (r'^(\d{2})[/\-\.](\d{2})$',                'dm'),
    (r'^(\d{8})$',                               'Ymd8'),
]

def _match_col(cell_text, col_type):
    if not cell_text: return False
    t = str(cell_text).strip().lower()
    for syn in _COL_SYNONYMS[col_type]:
        if t == syn or t.startswith(syn + ' ') or t.startswith(syn + '('): return True
    return False

def _detect_header_row(table):
    for row_idx, row in enumerate(table[:20]):
        col_map = {}
        for col_idx, cell in enumerate(row):
            if not cell: continue
            for ctype in ('date','label','debit','credit','amount','balance'):
                if ctype not in col_map and _match_col(str(cell), ctype):
                    col_map[ctype] = col_idx
        has_date  = 'date' in col_map
        has_money = ('debit' in col_map and 'credit' in col_map) or 'amount' in col_map
        if has_date and has_money:
            return row_idx, col_map
    return None, {}

def _parse_date_universal(raw, year_hint=None):
    if not raw: return None
    raw = str(raw).strip()
    raw = re.sub(r'^[A-Za-zÀ-ÿ]+\.?\s*', '', raw).strip()
    raw = raw.split('\n')[0].strip()
    for pattern, fmt in _DATE_PATTERNS:
        m = re.match(pattern, raw)
        if not m: continue
        if fmt == 'dmy4': return f"{m.group(3)}{m.group(2).zfill(2)}{m.group(1).zfill(2)}"
        elif fmt == 'ymd4': return f"{m.group(1)}{m.group(2).zfill(2)}{m.group(3).zfill(2)}"
        elif fmt == 'dmy2':
            yy = int(m.group(3))
            return f"{2000+yy if yy<=30 else 1900+yy}{m.group(2).zfill(2)}{m.group(1).zfill(2)}"
        elif fmt == 'dm':
            yr = str(year_hint) if year_hint else str(datetime.now().year)
            return f"{yr}{m.group(2).zfill(2)}{m.group(1).zfill(2)}"
        elif fmt == 'Ymd8':
            s = m.group(0); return f"{s[:4]}{s[4:6]}{s[6:8]}"
    return None

def _parse_amount_cell(cell_text):
    if not cell_text: return None
    s = str(cell_text).strip().replace('\xa0',' ').replace('\u202f',' ').replace('\n',' ').strip()
    s = re.sub(r'[€$£FCFAXOF]','',s,flags=re.IGNORECASE).strip().replace('*','').strip()
    if not s or s in ('.', ',', '-', '—', '–', ''): return None
    negative = False
    if s.startswith('(') and s.endswith(')'): s = s[1:-1].strip(); negative = True
    if s.startswith('-'): negative = True; s = s[1:].strip()
    elif s.startswith('+'): s = s[1:].strip()
    s_nospace = s.replace(' ','')
    m = re.match(r'^(\d{1,3}(?:[.,]\d{3})+)[,.](\d{2})$', s_nospace)
    if m:
        integer_part = re.sub(r'[,.]','',m.group(1))
        val = float(f"{integer_part}.{m.group(2)}")
        return -val if negative else val
    m2 = re.match(r'^(\d+)[,.](\d{1,2})$', s_nospace)
    if m2:
        val = float(f"{m2.group(1)}.{m2.group(2)}")
        return -val if negative else val
    m3 = re.match(r'^\d[\d\s]*\d$|^\d$', s)
    if m3:
        val = float(s.replace(' ',''))
        return -val if negative else val
    return None

def _extract_universal_header(pages_text):
    text = ' '.join(pages_text[:3])
    info = {'iban':'','period_start':'','period_end':'','balance_open':0.0,'balance_close':0.0}
    # Utilise extract_iban() centralisé pour cohérence avec tous les autres parsers
    info['iban'] = extract_iban(text)
    m1 = re.search(
        r'(?:du|from|de|period[e]?\s*:?)\s*(\d{1,2}[/\-.]\d{2}[/\-.]\d{2,4})'
        r'\s*(?:au|to|[àa]|\-)\s*(\d{1,2}[/\-.]\d{2}[/\-.]\d{2,4})',
        text, re.IGNORECASE)
    if m1:
        info['period_start'] = m1.group(1).replace('-','/').replace('.','/') 
        info['period_end']   = m1.group(2).replace('-','/').replace('.','/')
    return info

def _universal_parse_path(pdf_path, pages_text):
    info = _extract_universal_header(pages_text)
    year_hint = _year_from_text(' '.join(pages_text[:2]))
    txns = []
    SKIP_LABELS = {'TOTAL','TOTAUX','SOLDE','SOUS-TOTAL','REPORT','A REPORTER',
                   'NOUVEAU SOLDE','ANCIEN SOLDE','SOLDE INITIAL','SOLDE FINAL'}
    TABLE_SETTINGS_LIST = [
        {"vertical_strategy":"text","horizontal_strategy":"text","snap_tolerance":4,"join_tolerance":4},
        {"vertical_strategy":"lines","horizontal_strategy":"lines","snap_tolerance":3},
        {"vertical_strategy":"lines","horizontal_strategy":"text","snap_tolerance":4},
    ]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = None
            for settings in TABLE_SETTINGS_LIST:
                t = page.extract_table(settings)
                if t and len(t) >= 3:
                    table = t; break
            if not table: continue
            table_clean = [[str(c).replace('\n',' ').strip() if c else '' for c in row] for row in table]
            header_idx, col_map = _detect_header_row(table_clean)
            if header_idx is None: continue
            for row in table_clean[header_idx + 1:]:
                if not any(row): continue
                date_col = col_map.get('date')
                if date_col is None or date_col >= len(row): continue
                date_ofx = _parse_date_universal(row[date_col], year_hint)
                if not date_ofx: continue
                label_col = col_map.get('label')
                label = row[label_col].strip() if (label_col is not None and label_col < len(row)) else row[date_col]
                label_up = label.upper().strip()
                if not label or len(label) < 2: continue
                if any(skip in label_up for skip in SKIP_LABELS): continue
                if re.match(r'^[\d\s.,\-]+$', label): continue
                amount = None
                if 'debit' in col_map and 'credit' in col_map:
                    d_col, c_col = col_map['debit'], col_map['credit']
                    dv = _parse_amount_cell(row[d_col] if d_col < len(row) else '')
                    cv = _parse_amount_cell(row[c_col] if c_col < len(row) else '')
                    if dv and dv > 0: amount = -dv
                    elif cv and cv > 0: amount = cv
                elif 'amount' in col_map:
                    a_col = col_map['amount']
                    amount = _parse_amount_cell(row[a_col] if a_col < len(row) else '')
                if amount is None or amount == 0.0: continue
                name, memo = smart_label(label, [])
                txn = _make_txn(date_ofx, amount, name, memo)
                if txn: txns.append(txn)

    # ── Fallback texte ligne par ligne (banques non reconnues ou sans table) ──
    # Analyse chaque ligne à la recherche du pattern : DATE ... MONTANT
    # Fonctionne pour tout relevé avec des montants en fin de ligne.
    if not txns:
        full = '\n'.join(pages_text)
        full = full.replace('\xa0', ' ').replace('\u202f', ' ')
        SKIP_TEXT = ('SOLDE', 'TOTAL', 'TOTAUX', 'REPORT', 'DATE', 'VALEUR',
                     'LIBELLÉ', 'LIBELLE', 'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT',
                     'EUROS', 'MONTANT', 'PAGE', 'SUITE', 'VERSO', 'REF :',
                     'IBAN', 'BIC', 'AGENCE', 'COMPTE', 'TITULAIRE', 'ADRESSE',)
        # Regex générique : DD/MM/YYYY (optionnellement suivi d'une 2e date) + libellé + montant(s)
        DATE_RE = re.compile(r'^(\d{2}[/\-.]\d{2}[/\-.]\d{2,4})')
        AMT_RE  = re.compile(r'([\d]{1,3}(?:[.\s]\d{3})*[,]\d{2}|[\d]+[,]\d{2}|[\d]+[.]\d{2})')

        prev_solde = None
        lines = full.splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            m_date = DATE_RE.match(line)
            if not m_date:
                continue
            date_raw = m_date.group(1).replace('.', '/').replace('-', '/')
            date_ofx = _parse_date_universal(date_raw, year_hint)
            if not date_ofx:
                continue

            line_up = line.upper()
            if any(kw in line_up for kw in SKIP_TEXT):
                continue

            # Trouver tous les montants dans la ligne
            amounts_found = AMT_RE.findall(line)
            if not amounts_found:
                continue
            amounts_vals = [v for v in (parse_amount(a) for a in amounts_found) if v and v > 0.5]
            if not amounts_vals:
                continue

            # Extraire le libellé : retirer date(s) en début + montants en fin
            label_part = line
            # Supprimer jusqu'à 2 dates en début de ligne
            label_part = DATE_RE.sub('', label_part, count=1).strip()
            label_part = DATE_RE.sub('', label_part, count=1).strip()
            # Supprimer les montants numériques de fin
            label_part = re.sub(r'[\d\s.,]+$', '', label_part).strip()
            # Supprimer les caractères parasites restants
            label_part = re.sub(r'^[^\w]+', '', label_part).strip()

            if not label_part or len(label_part) < 3:
                continue
            if not re.search(r'[A-Za-zÀ-ÿ]{2,}', label_part):
                continue
            label_up2 = label_part.upper()
            if any(kw in label_up2 for kw in SKIP_TEXT):
                continue
            if re.match(r'^[\d\s/\-.,]+$', label_part):
                continue

            # Déterminer sens (débit/crédit) par variation de solde ou mots-clés
            is_credit = None
            if len(amounts_vals) >= 2:
                # Dernier montant = solde courant, avant-dernier = opération
                solde_courant = amounts_vals[-1]
                montant_op    = amounts_vals[-2]
                if prev_solde is not None:
                    diff = solde_courant - prev_solde
                    if diff > 0.5:
                        is_credit = True
                    elif diff < -0.5:
                        is_credit = False
                prev_solde = solde_courant
                amt = montant_op
            else:
                amt = amounts_vals[0]

            # Fallback mots-clés si sens non déterminé par solde
            if is_credit is None:
                CREDIT_KW = ('VIR ', 'VIREMENT', 'REGLEMENT', 'REMBOURSEMENT',
                             'VERSEMENT', 'AVOIR', 'RETOUR', 'REMISE', 'CREDIT',
                             'EDENRED', 'DELIVEROO', 'PLUXEE', 'BIMPLI', 'UBER',
                             'QUATRA', 'SCI ', 'RECETTE')
                DEBIT_KW  = ('PRLV', 'PRELEVEMENT', 'PAIEMENT CB', 'PAIEMENT PSC',
                             'PREL ', 'FACT ', 'COTISATION', 'ABONNEMENT',
                             'COMMISSION', 'FRAIS', 'AGIOS', 'RETRAIT',
                             'CHEQUE', 'LOYER', 'EDF', 'ORANGE', 'DGFIP',
                             'GENERALI', 'MAXANCE', 'SURAVENIR')
                if any(k in label_up2 for k in CREDIT_KW):
                    is_credit = True
                elif any(k in label_up2 for k in DEBIT_KW):
                    is_credit = False
                else:
                    is_credit = False  # défaut conservateur

            signed = amt if is_credit else -amt
            name, memo = smart_label(label_part, [])
            txn = _make_txn(date_ofx, signed, name, memo)
            if txn:
                txns.append(txn)

    return info, [t for t in txns if t is not None]

# ════════════════════════════════════════════════════════════════════════════
# CRÉDIT MUTUEL
# Format : Date | Date valeur | Opération | Débit EUROS | Crédit EUROS
# Relevé « RELEVE ET INFORMATIONS BANCAIRES » — Eurocompte Pro / Compte courant
# Particularités :
#   • La date opération et la date valeur sont toutes deux au format DD/MM/YYYY
#   • Le libellé principal est sur la ligne de date ; les lignes suivantes
#     (sans date) sont des lignes de continuation (mémo : référence, ICS, RUM…)
#   • Deux colonnes montant distinctes (Débit / Crédit) en fin de ligne
# Positions mesurées sur relevés CM réels :
#   Date op    : x0 < 70   (DD/MM/YYYY)
#   Date val   : x0 ≈ 70–130 (DD/MM/YYYY)
#   Libellé    : x0 ≈ 130–430
#   Débit      : x0 ≈ 430–530
#   Crédit     : x0 ≈ 530+
# ════════════════════════════════════════════════════════════════════════════

def _extract_cm_header(pages_text):
    text = ' '.join(pages_text[:3])
    info = {'iban': '', 'period_start': '', 'period_end': '',
            'balance_open': 0.0, 'balance_close': 0.0}

    # IBAN
    info['iban'] = extract_iban(text)

    # Période — "31 octobre 2025" → on déduit fin de mois, ou cherche dates explicites
    m_per = re.search(
        r'(?:du|Du)\s+(\d{2}/\d{2}/\d{4})\s+(?:au|Au)\s+(\d{2}/\d{2}/\d{4})',
        text, re.IGNORECASE
    )
    if m_per:
        info['period_start'] = m_per.group(1)
        info['period_end']   = m_per.group(2)
    else:
        # Chercher la date du relevé (ex: "31 octobre 2025")
        MOIS = {'janvier':'01','février':'02','fevrier':'02','mars':'03','avril':'04',
                'mai':'05','juin':'06','juillet':'07','août':'08','aout':'08',
                'septembre':'09','octobre':'10','novembre':'11','décembre':'12','decembre':'12'}
        m_date = re.search(
            r'(\d{1,2})\s+(' + '|'.join(MOIS.keys()) + r')\s+(20\d{2})',
            text, re.IGNORECASE
        )
        if m_date:
            day = m_date.group(1).zfill(2)
            month = MOIS.get(m_date.group(2).lower(), '01')
            year  = m_date.group(3)
            info['period_end'] = f"{day}/{month}/{year}"
            info['period_start'] = f"01/{month}/{year}"

    # Soldes — "SOLDE CREDITEUR AU 30/09/2025  4.286,81"
    for pat in [
        r'SOLDE\s+(?:CREDITEUR|DEBITEUR)\s+AU\s+\d{2}/\d{2}/\d{4}\s+([\d\s.,]+)',
        r'SOLDE\s+(?:INITIAL|D\'OUVERTURE)\s+([\d\s.,]+)',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            v = parse_amount(m.group(1).strip().split()[0])
            if v: info['balance_open'] = v; break

    # Solde final
    m_close = re.search(
        r'SOLDE\s+CREDITEUR\s+AU\s+\d{2}/\d{2}/\d{4}\s+([\d\s.,]+)',
        text, re.IGNORECASE
    )
    if m_close:
        vals = re.findall(r'[\d]+[.,][\d]{2}', text)
        # Prendre la dernière occurrence de "SOLDE CREDITEUR"
        all_solde = re.findall(
            r'SOLDE\s+CREDITEUR\s+AU\s+\d{2}/\d{2}/\d{4}\s+([\d.,\s]+)',
            text, re.IGNORECASE
        )
        if all_solde:
            v = parse_amount(all_solde[-1].strip().split()[0])
            if v: info['balance_close'] = v

    # "Total des mouvements  12.525,90  14.417,56" + "SOLDE CREDITEUR AU 31/10/2025  6.178,47"
    m_final = re.search(
        r'R[eé]f\s*:\s*\d+\s+SOLDE\s+CREDITEUR\s+AU\s+\d{2}/\d{2}/\d{4}\s+([\d.,\s]+)',
        text, re.IGNORECASE
    )
    if m_final:
        v = parse_amount(m_final.group(1).strip().split()[0])
        if v: info['balance_close'] = v

    return info


def parse_cm(pages_words, pages_text):
    """Parseur Crédit Mutuel — format RELEVE ET INFORMATIONS BANCAIRES.

    Stratégie :
    1. Parsing mot-par-mot (pdfplumber words) avec distinction colonne débit/crédit
       par mots-clés d'opération (plus fiable que la position x qui varie).
    2. Fallback : parsing texte brut ligne par ligne.
    """
    info = _extract_cm_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []

    SKIP_UP = {
        'SOLDE', 'TOTAL', 'TOTAUX', 'DATE', 'VALEUR', 'OPERATION', 'OPÉRATION',
        'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT', 'EUROS', 'LIBELLÉ', 'LIBELLE',
        'REPORT', 'A REPORTER', 'SUITE', 'VERSO', 'PAGE', 'RELEVE',
        'RELEVÉ', 'INFORMATIONS', 'BANCAIRES',
    }
    # Seuil de séparation débit/crédit mesuré sur relevés CM réels :
    # Débit  : x0 ≈ 428–499  (colonne gauche)
    # Crédit : x0 ≈ 500–535  (colonne droite)
    CM_CREDIT_X_MIN = 500  # si x0 du montant >= cette valeur → crédit

    # Mots-clés de secours quand la position est ambiguë ou le montant absent
    CREDIT_KW = (
        'REGLEMENT AFFILIES', 'STICHTING CUSTODIAN', 'EDENRED FRANCE',
        'DELIVEROO FRANCE', 'PLUXEE FRANCE', 'QUATRA FRANCE',
    )
    DEBIT_KW = (
        'PRLV SEPA', 'PRLV ', 'PRELEVEMENT', 'PAIEMENT CB', 'PAIEMENT PSC',
        'PREL EURO', 'FACT SGT', 'LOYER LOCAL', 'VIR SEPA LOYER',
    )

    def _cm_is_skip(label):
        up = label.upper().strip()
        if up in SKIP_UP: return True
        for kw in ('SOLDE ', 'TOTAL ', 'SUITE AU', '<<SUITE', 'SOUS RESERVE',
                   'DONT TVA', 'INFORMATION SUR'):
            if up.startswith(kw): return True
        return False

    def _cm_date(row):
        for w in row:
            if w['x0'] < 95 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']):
                return w['text']
        return ''

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=3.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_str = _cm_date(row)
            if not date_str:
                i += 1; continue

            # Libellé : mots entre fin des dates et début colonnes montants
            # Dates : x0 < 145, montants : x0 >= 415
            label_words = [w for w in row
                           if 145 <= w['x0'] < 415
                           and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            label = ' '.join(w['text'] for w in label_words).strip()

            if _cm_is_skip(label):
                i += 1; continue

            # Montant : tous les tokens à x0 >= 415
            amount_words = [w for w in row if w['x0'] >= 415]
            raw_amount_text = ' '.join(w['text'] for w in amount_words).strip()
            amt = parse_amount(raw_amount_text) if raw_amount_text else None
            # Position du premier token montant pour déduire la colonne
            amount_x0 = amount_words[0]['x0'] if amount_words else None

            # Lignes de continuation (mémo) : pas de date op
            memo_parts = []
            j = i + 1
            while j < len(rows):
                nr = rows[j]
                if _cm_date(nr):
                    break
                cont_words = [w for w in nr
                              if 145 <= w['x0'] < 415
                              and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                cont = ' '.join(w['text'] for w in cont_words).strip()
                if cont and not _cm_is_skip(cont):
                    memo_parts.append(cont)
                # Si montant absent de la ligne principale, chercher ici
                if amt is None:
                    alt_amount_words = [w for w in nr if w['x0'] >= 415]
                    alt_raw = ' '.join(w['text'] for w in alt_amount_words).strip()
                    if alt_raw:
                        amt = parse_amount(alt_raw)
                        if alt_amount_words:
                            amount_x0 = alt_amount_words[0]['x0']
                j += 1
            i = j

            if not label or not re.search(r'[A-Za-zÀ-ÿ]{2,}', label):
                continue
            if amt is None or amt == 0:
                continue

            date_ofx = date_full_to_ofx(date_str)
            if not re.match(r'^\d{8}$', date_ofx):
                continue

            # Déterminer débit / crédit — CRITÈRE PRINCIPAL : position x du montant
            # Crédit si x0 >= CM_CREDIT_X_MIN (colonne droite), sinon débit
            label_up = label.upper()
            if amount_x0 is not None:
                is_credit = (amount_x0 >= CM_CREDIT_X_MIN)
            else:
                # Fallback mots-clés si pas de position
                if any(k in label_up for k in CREDIT_KW):
                    is_credit = True
                elif any(k in label_up for k in DEBIT_KW):
                    is_credit = False
                else:
                    is_credit = False

            signed = amt if is_credit else -amt
            name, memo = smart_label(label, memo_parts)
            txns.append(_make_txn(date_ofx, signed, name, memo))

    # Fallback texte brut si pdfplumber n'a pas donné de words utilisables
    if not txns:
        full = '\n'.join(pages_text)
        full = full.replace('\xa0', ' ').replace('\u202f', ' ')
        SKIP_TEXT = ('SOLDE', 'TOTAL', 'TOTAUX', 'RELEVE', 'RELEVÉ', 'DATE',
                     'VALEUR', 'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT', 'EUROS',
                     'SUITE', 'VERSO', 'PAGE', 'INFORMATIONS', 'BANCAIRES', 'REF :',
                     'DONT TVA', 'SOUS RESERVE', 'INFORMATION SUR')
        CREDIT_KW2 = ('VIR ', 'REGLEMENT', 'EDENRED', 'DELIVEROO', 'PLUXEE',
                      'BIMPLI', 'UBER', 'M&M', 'QUATRA', 'SCI ')
        DEBIT_KW2  = ('PRLV', 'PAIEMENT', 'PREL ', 'FACT SGT', 'LOYER')
        for line in full.splitlines():
            line = line.strip()
            m = re.match(r'^(\d{2}/\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}\s+(.+?)\s+([\d.]+[,]\d{2})\s*$', line)
            if not m:
                m = re.match(r'^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d.]+[,]\d{2})\s*$', line)
            if not m: continue
            date_str, label, amt_str = m.group(1), m.group(2).strip(), m.group(3)
            label_up = label.upper()
            if any(kw in label_up for kw in SKIP_TEXT): continue
            if not re.search(r'[A-Za-zÀ-ÿ]{2,}', label): continue
            amt = parse_amount(amt_str)
            if not amt: continue
            date_ofx = date_full_to_ofx(date_str)
            if not re.match(r'^\d{8}$', date_ofx): continue
            is_credit = (any(k in label_up for k in CREDIT_KW2) and
                         not any(k in label_up for k in DEBIT_KW2))
            signed = amt if is_credit else -amt
            txns.append(_make_txn(date_ofx, signed, label))

    return info, [t for t in txns if t is not None]


# ════════════════════════════════════════════════════════════════════════════
# PARSEURS AFRICAINS DÉDIÉS
# ════════════════════════════════════════════════════════════════════════════

def _afr_header(pages_text):
    """Header commun pour les banques africaines."""
    text = ' '.join(pages_text[:3])
    info = {'iban': '', 'period_start': '', 'period_end': '',
            'balance_open': 0.0, 'balance_close': 0.0}

    # ── 1. RIB explicitement labellisé : Code Banque / Agence / Compte / Clé ──
    # Ex tableau : "Code Banque  Agence   Compte         Clé RIB"
    #              "SN213        01001    02341624101    33"
    rib_table = re.search(
        r'(?:Code\s*Banque|Banque)\s*[:\|]?\s*([A-Z]{0,2}\d{3,5})'
        r'.{0,40?}(?:Agence|Guichet)\s*[:\|]?\s*(\d{4,6})'
        r'.{0,40?}(?:N[°o]?\s*(?:de\s*)?[Cc]ompte|Compte)\s*[:\|]?\s*(\d{8,14})'
        r'.{0,40?}(?:Cl[eé]\s*(?:RIB)?|RIB)\s*[:\|]?\s*(\d{2})\b',
        text, re.IGNORECASE | re.DOTALL
    )
    if rib_table:
        info['_rib_bank']    = rib_table.group(1)
        info['_rib_agency']  = rib_table.group(2)
        info['_rib_account'] = rib_table.group(3)
        info['_rib_key']     = rib_table.group(4)

    # ── 2. RIB compact sur une ligne : "SN213 01001 02341624101 33" ───────────
    _SEP = r'[\s\t\-]+'
    if not info.get('_rib_bank'):
        rib_match = re.search(
            r'\b([A-Z]{0,2}\d{3,5})' + _SEP +
            r'(\d{4,6})'             + _SEP +
            r'(\d{8,14})'            + _SEP +
            r'(\d{2})\b',
            text
        )
        if rib_match:
            info['_rib_bank']    = rib_match.group(1)
            info['_rib_agency']  = rib_match.group(2)
            info['_rib_account'] = rib_match.group(3)
            info['_rib_key']     = rib_match.group(4)

    # ── 2b. RIB format tiret BSIC/ECOBANK : "01001-00100029193-76" ─────────────
    if not info.get('_rib_bank'):
        rib_tiret = re.search(
            r'(?:Num[\xe9e]ro\s+de\s+compte|N[o\xb0]\s*compte)\s*[:\-]?\s*'
            r'(\d{5})-(\d{8,14})-(\d{2})',
            text, re.IGNORECASE
        )
        if rib_tiret:
            info['_rib_agency']  = rib_tiret.group(1)
            info['_rib_account'] = rib_tiret.group(2)
            info['_rib_key']     = rib_tiret.group(3)

    # ── 3. Extraction IBAN centralisée ───────────────────────────────────────
    if not info['iban']:
        info['iban'] = extract_iban(text)

    # ── 4. Si IBAN trouvé mais pas de code banque, dériver depuis l'IBAN ──────
    # Ne pas ecraser _rib_account si deja extrait par regex tiret (plus fiable)
    if info['iban'] and not info.get('_rib_bank'):
        try:
            b, ag, ac = iban_to_rib(info['iban'])
            if b and b != '00000':
                info['_rib_bank']   = b
                info['_rib_agency'] = info.get('_rib_agency') or ag
                # Ne pas ecraser le compte si deja capture par la regex tiret
                if not info.get('_rib_account'):
                    info['_rib_account'] = ac
                info['_rib_key']    = info.get('_rib_key') or ''
        except Exception:
            pass

    # ── 5. Numéro de compte brut (dernier recours) ───────────────────────────
    if not info['iban']:
        m4 = re.search(
            r'(?:N[°o°]\.?\s*(?:de\s*)?compte|Compte|COMPTE)\s*[:\-]?\s*'
            r'([\d]{5,20}(?:[\s\-]?\d{1,6})*)',
            text, re.IGNORECASE
        )
        if m4:
            info['iban'] = re.sub(r'[\s\-]', '', m4.group(1))

    # ── 6. Période ───────────────────────────────────────────────────────────
    m5 = re.search(r'(?:du|[Pp]ériode du?|[Pp]our la p[ée]riode du?|[Dd]u)\s+'
                   r'(\d{1,2}[/\-\.]\d{2}[/\-\.]\d{2,4})'
                   r'\s*(?:au|[àa]|[Aa]u)\s*'
                   r'(\d{1,2}[/\-\.]\d{2}[/\-\.]\d{2,4})',
                   text, re.IGNORECASE)
    if m5:
        info['period_start'] = m5.group(1).replace('-', '/').replace('.', '/')
        info['period_end']   = m5.group(2).replace('-', '/').replace('.', '/')
    return info


# ── BSIC : format « Date Op | Date Val | Libellé | Débit | Crédit | Solde »
# Positions mesurées sur PDF réel :
#   Date opération : x0 ≈ 32    (dd/mm/yyyy)
#   Date valeur    : x0 ≈ 80    (dd/mm/yyyy) — à exclure du libellé
#   Libellé        : x0 ≈ 128-320
#   Débit          : x0 ≈ 355-410
#   Crédit         : x0 ≈ 440-500
#   Solde          : x0 ≈ 510+  (ignoré)
def parse_bsic(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []
    SKIP = {'TOTAL','SOLDE','A REPORTER','REPORT','DATE','VALEUR','LIBELLÉ','LIBELLE',
            'DÉBIT','DEBIT','CRÉDIT','CREDIT','EXTRAIT','PÉRIODE','CODE','NOM',
            'PAGE SUR'}  # "PAGE" et "SUR" uniquement en mots entiers via regex ci-dessous

    def _is_skip_label(lbl):
        """Retourne True si le libellé est une ligne d'en-tête/pied à ignorer."""
        lbl_up = lbl.upper().strip()
        # Mots-clés exacts (libellé = entièrement ce mot)
        if lbl_up in SKIP:
            return True
        # Mots-clés en début de libellé (solde, total…)
        for kw in ('SOLDE', 'TOTAL', 'A REPORTER', 'REPORT', 'DÉBIT', 'DEBIT',
                   'CRÉDIT', 'CREDIT', 'DATE', 'VALEUR', 'LIBEL', 'EXTRAIT'):
            if lbl_up.startswith(kw):
                return True
        # "Page N sur N" — mot isolé SUR/PAGE uniquement s'il est le seul mot significatif
        if re.match(r'^PAGE\s+\d+\s+SUR\s+\d+$', lbl_up):
            return True
        return False

    # ── Extraction solde final depuis le texte ────────────────────────────────
    # "Solde (XOF) au 31/01/2024 : 8 728 070"
    full_text = ' '.join(pages_text)
    m_bal = re.search(
        r'Solde\s+\([A-Z]+\)\s+au\s+\d{2}/\d{2}/\d{4}\s*:\s*([\d\s]+)',
        full_text, re.IGNORECASE
    )
    if m_bal:
        raw_bal = re.sub(r'\s+', '', m_bal.group(1))
        try:
            info['balance_close'] = float(raw_bal)
        except ValueError:
            pass

    # ── Extraction RIB/IBAN BSIC ─────────────────────────────────────────────
    # Format numéro de compte : "01001-00100029193-76 XOF"
    #   → Guichet=01001, Compte=00100029193, Clé=76
    # IBAN collé possible : "Code Iban : SN08SN11101001000100029193"

    m_compte = re.search(
        r'Num[eé]ro\s+de\s+compte\s*:\s*(\d{5})-(\d{8,14})-(\d{2})',
        full_text, re.IGNORECASE
    )
    if m_compte:
        info['_rib_agency']  = m_compte.group(1)   # ex: 01001
        info['_rib_account'] = m_compte.group(2)   # ex: 00100029193
        info['_rib_key']     = m_compte.group(3)   # ex: 76
        if not info.get('_rib_bank'):
            iban_raw = re.sub(r'\s+', '', info.get('iban', '')).upper()
            if re.match(r'^[A-Z]{2}\d{2}', iban_raw) and len(iban_raw) >= 9:
                info['_rib_bank'] = iban_raw[4:9]
            else:
                m_ib = re.search(r'Code\s+[Ii]ban\s*:\s*([A-Z]{2}\d{2}[A-Z0-9]{3,5})',
                                  full_text, re.IGNORECASE)
                if m_ib:
                    info['_rib_bank'] = m_ib.group(1)[4:9]

    # Valider/compléter l'IBAN depuis "Code Iban : SN08SN111..."
    if not info.get('iban') or not re.match(r'^[A-Z]{2}\d{2}', re.sub(r'\s+','',info.get('iban','')).upper()):
        m_ci = re.search(r'Code\s+[Ii]ban\s*:\s*([A-Z]{2}\d{2}[A-Z0-9\s]{14,32})',
                          full_text, re.IGNORECASE)
        if m_ci:
            raw = re.sub(r'\s+', '', m_ci.group(1)).upper()
            raw = re.sub(r'[^A-Z0-9]', '', raw)
            if len(raw) >= 15:
                info['iban'] = raw[:28]

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=2.0)  # réduit pour BSIC (lignes serrées)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date opération : dd/mm/yyyy à x0 ≈ 32 (tolérance 70 pts)
            date_w = [w for w in row if w['x0'] < 70
                      and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_w:
                i += 1; continue

            date_str = date_w[0]['text']
            date_ofx = date_full_to_ofx(date_str)

            # Libellé : x0 entre 120 et 335 sur la ligne courante
            label_words = [w for w in row if 120 <= w['x0'] < 335]
            label = ' '.join(w['text'] for w in label_words).strip()

            # Si le libellé est vide, chercher dans la ou les lignes précédentes
            # (cas BSIC : "Virmt fav. ALIOS FINANCE" sur la ligne juste avant la date)
            if not label and i > 0:
                # Chercher jusqu'à 3 lignes en arrière (sans date op)
                for k in range(i - 1, max(i - 4, -1), -1):
                    prev_row = rows[k]
                    prev_date = [w for w in prev_row if w['x0'] < 70
                                 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                    if prev_date:
                        break  # une autre date op → stop
                    prev_label_words = [w for w in prev_row if 120 <= w['x0'] < 335]
                    prev_label = ' '.join(w['text'] for w in prev_label_words).strip()
                    if prev_label and not _is_skip_label(prev_label):
                        label = prev_label
                        break

            # Récupérer les lignes de continuation qui suivent (sans date op)
            j = i + 1
            memo_parts = []
            while j < len(rows):
                next_row = rows[j]
                next_date = [w for w in next_row if w['x0'] < 70
                             and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                if next_date:
                    break
                cont_words = [w for w in next_row if 120 <= w['x0'] < 335]
                cont = ' '.join(w['text'] for w in cont_words).strip()
                if cont and not _is_skip_label(cont):
                    memo_parts.append(cont)
                j += 1
            i = j

            if not label or _is_skip_label(label):
                continue
            if re.match(r'^[\d\s/\-]+$', label):
                continue

            # Débit (x0 ≈ 355-415) / Crédit (x0 ≈ 440-510)
            debit_words  = [w for w in row if 340 <= w['x0'] < 430]
            credit_words = [w for w in row if 430 <= w['x0'] < 515]

            debit_amt  = _uba_join_amount(debit_words)
            credit_amt = _uba_join_amount(credit_words)

            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    # ── Fallback texte brut (PDF scanné via OCR ou pages_words vides) ──────────
    # Quand pages_words est vide (scan OCR), on parse le texte ligne par ligne.
    # Format BSIC OCR attendu par ligne :
    #   "01/10/2025 30/09/2025 RET ESP CP 2149951 MESSIN DRA   80 000   9 262 621"
    #   "07/10/2025 07/10/2025 V/V FAC FV2025 09 000260        867 300   948 931"
    # → date_op  date_val  libellé  [débit ou crédit]  solde
    if not txns:
        full_text_ocr = '\n'.join(pages_text)
        # Pré-traitement : normaliser espaces insécables
        full_text_ocr = full_text_ocr.replace('\xa0', ' ').replace('\u202f', ' ')

        # Extraire toutes les lignes de transaction + la suivante (solde précédent
        # pour déduire le sens débit/crédit)
        lines = full_text_ocr.split('\n')
        prev_solde = None

        for idx, line in enumerate(lines):
            line = line.strip()
            # Doit commencer par date_op dd/mm/yyyy
            m_date = re.match(r'^(\d{2}/\d{2}/\d{4})', line)
            if not m_date:
                continue
            date_str = m_date.group(1)
            date_ofx = date_full_to_ofx(date_str)

            line_up = line.upper()
            if any(kw in line_up for kw in ('SOLDE', 'TOTAL', 'LIBELLÉ', 'LIBELLE',
                                             'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT',
                                             'EXTRAIT', 'PAGE', 'PÉRIODE', 'CODE CLIENT')):
                continue

            # Extraire tous les montants de la ligne (≥ 100 pour ignorer numéros)
            # On cherche des nombres avec éventuels espaces comme séparateurs de milliers
            raw_amounts = re.findall(
                r'\b(\d{1,3}(?:\s\d{3})*(?:,\d{2})?)\b', line)
            amounts_vals = []
            for a in raw_amounts:
                v = parse_amount(a.replace(' ', '.'))
                if v is not None and v >= 100:
                    amounts_vals.append(v)

            if not amounts_vals:
                continue

            # Libellé : retirer date_op + date_val, puis enlever les montants de fin
            label_part = re.sub(r'^\d{2}/\d{2}/\d{4}\s*', '', line)    # date op
            label_part = re.sub(r'^\d{2}/\d{2}/\d{4}\s*', '', label_part)  # date val
            # Retirer les montants numériques de fin de chaîne
            label_part = re.sub(r'[\d\s,]+$', '', label_part).strip()
            label_part = clean_label(label_part)

            if not label_part or len(label_part) < 3 or _is_skip_label(label_part):
                continue

            # ── Déduire débit/crédit ─────────────────────────────────────────
            # Méthode 1 : variation de solde
            # Format : libellé  [montant_op]  solde_courant
            # Si on a ≥ 2 montants : le dernier est le solde, l'avant-dernier est l'opération
            # Si solde_courant < solde_précédent → débit, sinon crédit
            is_credit = None
            if len(amounts_vals) >= 2:
                solde_courant = amounts_vals[-1]
                montant_op    = amounts_vals[-2]
                if prev_solde is not None:
                    diff = solde_courant - prev_solde
                    # Tolérance de 10 XOF pour arrondi OCR
                    if diff > 10:
                        is_credit = True
                    elif diff < -10:
                        is_credit = False
                prev_solde = solde_courant
                amt = montant_op
            else:
                amt = amounts_vals[0]

            # Méthode 2 : mots-clés du libellé (si méthode 1 insuffisante)
            if is_credit is None:
                credit_kw = ('V/V', 'FAC FV', 'VIREMENT ENTRANT', 'REMISE',
                             'VERSEMENT', 'AVOIR', 'RETOUR', 'REMBOURSEMENT',
                             'VIREMENT REÇU', 'CREDIT')
                debit_kw  = ('VIREMENT W', 'FACTURE ARC', 'CHQ', 'PREST',
                             'RET ESP', 'FRAIS', 'RETRAIT', 'COMMISSION',
                             'COTISATION', 'ABONNEMENT', 'TENUE DE COMPTE',
                             'SOUSCRIPTION', 'SMS')
                lbl_up2 = label_part.upper()
                if any(kw in lbl_up2 for kw in credit_kw):
                    is_credit = True
                elif any(kw in lbl_up2 for kw in debit_kw):
                    is_credit = False
                else:
                    is_credit = False  # défaut : débit

            signed_amt = amt if is_credit else -amt
            txns.append(_make_txn(date_ofx, signed_amt, label_part))

    # Fallback universel si toujours rien trouvé
    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ── BIS (Banque Islamique du Sénégal) : format iMAL*CSM
# Positions mesurées sur PDF réel :
#   Date+Valeur : x0 ≈ 19-103 (format "dd/mm/yyyydd/mm/yyyy" concaténé ou simple)
#   No Trs      : x0 ≈ 111-152
#   Description : x0 ≈ 157-400
#   Mnt Crédit  : x0 ≈ 320-415  ('0' = colonne vide)
#   Mnt Débit   : x0 ≈ 415-500  ('0' = colonne vide)
#   Solde       : x0 ≈ 509+
def parse_bis(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    txns = []
    SKIP = {'TOTAL','SOLDE','DATE','VALEUR','NO TRS','DESCRIPTION','MNT','DÉBIT',
            'RÉSUMÉ','BANQUE','ISLAMIQUE','ECOLE','CIF','NO. CPTE','DISPONIBLE',
            'Solde','Bénéficiaire','Page','de','**'}

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=5.0)
        for row in rows:
            # Date : token en x0 < 110 commençant par dd/mm/yyyy
            # Peut être concaténé "02/05/202402/05/2024" (date trs + date valeur)
            date_w = None
            for w in row:
                if w['x0'] < 110:
                    m = re.match(r'^(\d{2}/\d{2}/\d{4})', w['text'])
                    if m:
                        date_w = w; break
            if not date_w:
                continue
            date_str = date_w['text'][:10]  # prendre dd/mm/yyyy

            # Ignorer les lignes de pied de page (contenant heure HH:MM:SS)
            if any(re.match(r'^\d{2}:\d{2}:\d{2}$', w['text']) for w in row):
                continue

            # Description : x0 ~157-400
            desc_words = [w for w in row if 150 <= w['x0'] < 400]
            label = ' '.join(w['text'] for w in desc_words).strip()
            if not label:
                continue
            if re.match(r'^[0-9#,\s]+$', label):
                continue
            if any(s in label for s in SKIP):
                continue

            # Crédit (Mnt Cr) : x0 ≈ 320-415  |  Débit (Mnt débit) : x0 ≈ 415-500
            credit_words = [w for w in row if 320 <= w['x0'] < 415]
            debit_words  = [w for w in row if 415 <= w['x0'] < 500]

            credit_raw = ' '.join(w['text'] for w in credit_words).strip()
            debit_raw  = ' '.join(w['text'] for w in debit_words).strip()

            credit_amt = _uba_join_amount(credit_words) if credit_raw and credit_raw != '0' else None
            debit_amt  = _uba_join_amount(debit_words)  if debit_raw  and debit_raw  != '0' else None

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, [])
            if credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
            elif debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ── BNDE : format tableau Date | Libellé | Valeur | Débit | Crédit | Solde
# PDF natif pdfplumber → parsing mot-par-mot d'abord, texte brut en fallback
def parse_bnde(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []

    SKIP = {'TOTAL','SOLDE','DATE','LIBELLÉ','LIBELLE','VALEUR','DÉBIT','DEBIT',
            'CRÉDIT','CREDIT','A REPORTER','SOLDE À REPORTER','TITULAIRE','RELEVE',
            'VEUILLEZ','PAGE','BNDE','AGENCE','COMPTE','DEVISE','DOMICILIATION',
            'SIÈGE','SOCIAL','RCCM','NINEA'}
    CREDIT_KW = {'VERSEMENT','VIREMENT','RECU','REMISE','CREDIT','DÉBLOCAGE',
                 'DEBLOCAGE','ANNUL CHQ','RECOUVREMENT','SWIFT','PAR :',
                 'CNCA THIES','REMISE DE CHEQUES','TIRE :'}
    DEBIT_KW  = {'RETRAIT','CHEQ','CHQ','AGIOS','FRAIS','COMMISSION',
                 'ABONNEMENT','ROUTAGE','FAV :','BENEF :','VIREMENT AUTRE BANQUE',
                 '000090','000110','000104','000103','000101','000106',
                 '000115','000111','000112','000114','000113','000123',
                 '000131','000130','000132','000124','000128','000905',
                 '000906'}

    # ── Essai 1 : parsing pdfplumber words ───────────────────────────────────
    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date : DD/MM/YYYY en x0 < 100
            date_words = [w for w in row if w['x0'] < 100
                          and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_words:
                i += 1; continue
            date_str = date_words[0]['text']

            label_words = [w for w in row if 70 <= w['x0'] < 360
                           and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()

            # Si libellé vide/numérique/trop court, chercher dans lignes adjacentes sans date
            if not label or len(label) < 3 or re.match(r'^[\d\s/\-.,]+$', label):
                for k in list(range(i - 1, max(i - 4, -1), -1)) + list(range(i + 1, min(i + 4, len(rows)))):
                    r2 = rows[k]
                    has_date2 = any(w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                                    for w in r2)
                    if has_date2:
                        continue
                    adj_words = [w for w in r2 if 70 <= w['x0'] < 360
                                 and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                    adj = ' '.join(w['text'] for w in adj_words).strip()
                    if adj and len(adj) >= 3 and re.search(r'[A-Za-zÀ-ÿ]{2,}', adj):
                        label = adj
                        break
                label_up = label.upper()

            if not label or any(s == label_up for s in SKIP):
                i += 1; continue
            if re.match(r'^[\d\s/\-.,]+$', label):
                i += 1; continue

            # Montants : débit (x0 342-421) et crédit (x0 421-503)
            # Positions mesurées sur PDF BNDE réel (SUP DECO THIES) :
            #   Débit  header x0=342, Crédit header x0=421, Solde header x0=503
            debit_words  = [w for w in row if 335 <= w['x0'] < 420]
            credit_words = [w for w in row if 420 <= w['x0'] < 503]
            debit_amt  = _uba_join_amount(debit_words)
            credit_amt = _uba_join_amount(credit_words)

            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and any(w['x0'] < 100 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                              for w in r2):
                    break
                nl = ' '.join(w['text'] for w in r2 if 70 <= w['x0'] < 360).strip()
                if nl and not any(s in nl.upper() for s in SKIP) and len(nl) > 2:
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
            else:
                # Fallback sémantique si colonnes non distinctes
                all_right = [w for w in row if w['x0'] >= 360]
                amt = _uba_join_amount(all_right)
                if amt and amt > 100:
                    is_credit = any(k in label_up for k in CREDIT_KW)
                    is_debit  = any(k in label_up for k in DEBIT_KW)
                    if is_credit and not is_debit:
                        txns.append(_make_txn(date_ofx, amt, name, memo))
                    elif is_debit and not is_credit:
                        txns.append(_make_txn(date_ofx, -amt, name, memo))

    # ── Essai 2 : parseur universel table ─────────────────────────────────────
    if not txns and _pdf_path and Path(_pdf_path).exists():
        result_info, txns_u = _universal_parse_path(_pdf_path, pages_text)
        if txns_u:
            result_info.update({k: v for k, v in info.items() if v})
            return result_info, txns_u

    # ── Essai 3 : fallback texte brut DD/MM/YYYY ──────────────────────────────
    if not txns:
        full_text = '\n'.join(pages_text)
        date_re = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+')
        for line in full_text.splitlines():
            line = line.strip()
            if not line: continue
            m = date_re.match(line)
            if not m: continue
            date_str = m.group(1)
            rest = line[m.end():]
            # Sauter la date valeur si présente
            rest = re.sub(r'^\d{2}/\d{2}/\d{4}\s*', '', rest).strip()
            rest_up = rest.upper()
            if any(s in rest_up for s in SKIP): continue

            amounts_raw = re.findall(r'\d[\d\s]*[,.]\d{2}|\d+', rest)
            amounts_parsed = [v for a in amounts_raw
                              for v in [parse_amount(a.strip().replace(' ', '').replace(',', '.'))]
                              if v and v >= 100]
            if not amounts_parsed: continue

            label_m = re.match(r'^([A-Za-zÀ-ÿ\s\-\'\./,:&°0-9]+?)(?=\s{2,}\d|\s+\d{3,})', rest)
            label = label_m.group(1).strip() if label_m else rest[:60].strip()
            if not label or len(label) < 3: continue
            if any(s in label.upper() for s in SKIP): continue
            if not re.search(r'[A-Za-zÀ-ÿ]{2,}', label): continue

            is_credit = any(k in label.upper() for k in CREDIT_KW)
            is_debit  = any(k in label.upper() for k in DEBIT_KW)
            amt = amounts_parsed[0]
            date_ofx = date_full_to_ofx(date_str)
            name, memo_out = smart_label(label, [])
            if is_credit and not is_debit:
                txns.append(_make_txn(date_ofx, amt, name, memo_out))
            elif is_debit and not is_credit:
                txns.append(_make_txn(date_ofx, -amt, name, memo_out))

    # ── Extraction solde depuis le texte (BNDE) ───────────────────────────────
    full_text_bnde = ' '.join(pages_text)
    m_close = re.search(
        r'Solde\s+\([A-Z]+\)\s+au\s+\d{2}/\d{2}/\d{4}\s*:\s*([\d\s]+)',
        full_text_bnde, re.IGNORECASE)
    if m_close:
        v = parse_amount(re.sub(r'\s+', '', m_close.group(1).strip()))
        if v and v > 0:
            info['balance_close'] = v
    m_open = re.search(
        r'Solde\s+initial\s+\([A-Z]+\)\s*:\s*([\d\s]+)',
        full_text_bnde, re.IGNORECASE)
    if m_open:
        v = parse_amount(re.sub(r'\s+', '', m_open.group(1).strip()))
        if v and v > 0:
            info['balance_open'] = v

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ── UBA : format Extrait de compte
# Structure réelle : le libellé est sur les lignes PRÉCÉDANT la ligne de date.
# Ligne de date: date (x0<90), date valeur (x0≈269), débit (x0≈330-430), crédit (x0≈430-520), solde (x0≈520+)
def parse_uba(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    txns = []
    CREDIT_KW = {'MAINLEVEE','BLOCAGE','RECU','SWIFT','DECAISSEMENT','VERSEMENT','REMISE','CREDIT'}
    DEBIT_KW  = {'VISA','FRAIS','COMMISSION','CHEQUE','TOB','ECHEANCE','PRET',
                 'FACTURATION','REDEVANCE','RETRAIT','AGIOS'}
    HEADER_SKIP = {'SOLDE','DÉBIT','DEBIT','CRÉDIT','CREDIT','VALEUR','DATE',
                   'OPÉRATION','OPERATION','INSTR','AGENCE','COMPTE','PÉRIODE',
                   'POUR','RELEVÉ','EXTRAIT','TITULAIRE','RIB','BIC'}

    def _group_num_tokens(words):
        blocks, cur, prev_x1 = [], [], None
        for w in words:
            is_num = bool(re.match(r'^[\d\s,\.]+$', w['text']) and re.search(r'\d', w['text']))
            if not is_num:
                if cur: blocks.append(cur); cur = []
                prev_x1 = None; continue
            if prev_x1 is not None and (w['x0'] - prev_x1) > 35:
                if cur: blocks.append(cur)
                cur = [w]
            else:
                cur.append(w)
            prev_x1 = w.get('x1', w['x0'] + 20)
        if cur: blocks.append(cur)
        return blocks

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=5.0)

        # Identifier les indices des lignes de date (dd/mm/yyyy à x0 < 90)
        date_row_indices = []
        for idx, row in enumerate(rows):
            dw = [w for w in row if w['x0'] < 90
                  and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if dw:
                date_row_indices.append(idx)

        for i, date_idx in enumerate(date_row_indices):
            row = rows[date_idx]
            date_str = [w for w in row if w['x0'] < 90
                        and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])][0]['text']

            # Libellé = lignes entre date row précédente et celle-ci
            prev_date_idx = date_row_indices[i - 1] if i > 0 else -1
            label_parts = []
            for j in range(prev_date_idx + 1, date_idx):
                r = rows[j]
                lw = [w for w in r if 115 <= w['x0'] < 410]
                text = ' '.join(w['text'] for w in lw).strip()
                if not text or re.match(r'^[\d/,\.\s]+$', text):
                    continue
                text_up = text.upper()
                if any(s in text_up for s in HEADER_SKIP):
                    continue
                if re.match(r'^[au\s\d/]+$', text, re.IGNORECASE):
                    continue
                label_parts.append(text)
            label = ' '.join(label_parts).strip()

            if not label:
                row_label = ' '.join(w['text'] for w in row if 80 <= w['x0'] < 260).strip()
                label = row_label

            if not label:
                continue
            label_up = label.upper()
            has_useful = bool(re.search(r'[A-Za-zÀ-ÿ]{3,}', label))
            if not has_useful:
                continue

            # Montants sur la ligne de date (à droite de x0≥250)
            right_words = sorted([w for w in row if w['x0'] >= 250], key=lambda w: w['x0'])
            blocks = _group_num_tokens(right_words)
            numeric_blocks = []
            for b in blocks:
                joined = ''.join(w['text'] for w in b)
                if re.match(r'^\d{2}/\d{2}/\d{4}$', joined):
                    continue
                v = _uba_join_amount(b)
                if v is not None:
                    numeric_blocks.append((v, b[0]['x0']))

            if not numeric_blocks:
                continue

            # Débit x0≈330-430, Crédit x0≈430-520, Solde x0≈520+
            debit_cands  = [(v, x) for v, x in numeric_blocks if 330 <= x < 430]
            credit_cands = [(v, x) for v, x in numeric_blocks if 430 <= x < 520]

            debit_amt  = debit_cands[0][0]  if debit_cands  else None
            credit_amt = credit_cands[0][0] if credit_cands else None

            if debit_amt is None and credit_amt is None and numeric_blocks:
                val, x0 = numeric_blocks[0]
                is_credit = any(k in label_up for k in CREDIT_KW)
                is_debit  = any(k in label_up for k in DEBIT_KW)
                if is_credit and not is_debit:
                    credit_amt = val
                elif is_debit and not is_credit:
                    debit_amt = val

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, [])
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ── SG Afrique (SGBS Sénégal) : Date | Libellé | Débit | Crédit
# Particularités : montants XOF entiers sans décimales ('15 400', '369 630'),
# fragmentés en plusieurs tokens par pdfplumber.
def parse_sg_afrique(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    txns = []
    SKIP = {'TOTAUX DES','NOUVEAU SOLDE','SOLDE PRECEDENT','PROGRAMME DE',
            'RAPPEL DES','MONTANT CUMULE','SOLDE AU','DATE D','DATE DE','LIBELLÉ',
            'TOTAL DES','DÉBIT','CRÉDIT','SOLDE','TOTAL','LIBELLE'}

    # Mots-clés sémantiques pour déduire le sens quand la position seule est ambiguë
    CREDIT_KW = {'VIREMENT','VERSEMENT','REMISE','CREDIT','RECU','SWIFT',
                 'DECAISSEMENT','MAINLEVEE','DEBLOCAGE','TRF-RECU','TRF RECU',
                 'REM.CHQ','REMISE CHQ','TRF-REÇU','VIRT RECU','CAPI SENEGAL',
                 'SOCOCIM','SENEGALAISE','ASS MEDIA','EDWIGE','DIOP EDWIGE',
                 'C LINES INTER', 'S.A.S. C', 'VERSEMENT DIOP', 'VERSEMENT EDWIGE',
                 'TIRE :', 'BENEF :', 'SORT CHEQUES', 'CNCA'}
    DEBIT_KW  = {'REDEVANCE','CHEQUE','COMMISSION','FRAIS','ECHEANCE','PRET',
                 'FACTURATION','RETRAIT','AGIOS','PRELEVEMENT','PRLV',
                 'CHQ COMP','RETRAIT ESPECES','IMPAYE','IBE','ABONNEMENT',
                 'FRAIS TELECOMP','FRAIS VIRTREG','VIREMENT REG PRO ENERGY',
                 'REG PRO ENERGY','LUFTHANSA'}

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date : dd/mm/yyyy en x0 < 75
            if not (row and row[0]['x0'] < 75
                    and re.match(r'^\d{2}/\d{2}/\d{4}$', row[0]['text'])):
                i += 1; continue

            label = ' '.join(w['text'] for w in row if 115 <= w['x0'] < 370).strip()
            label_up = label.upper()
            if not label or any(s in label_up for s in SKIP):
                i += 1; continue

            # --- Détection robuste des montants ---
            # SG Sénégal : Date | Date_valeur (79–114) | Libellé (115–370)
            # Débit  ≈ x0 370–420, Crédit ≈ x0 445–490, Solde ≈ x0 520+
            # Zones mesurées sur PDF réel SGBS Sénégal
            right_words = sorted([w for w in row if w['x0'] >= 360],
                                 key=lambda w: w['x0'])

            # Regrouper en blocs contigus (gap > 30px = nouveau bloc)
            def _make_blocks(words, gap=30):
                blocks, cur, prev_x1 = [], [], None
                for w in words:
                    is_num = bool(re.match(r'^[\d,\.]+$', w['text'])
                                  and re.search(r'\d', w['text']))
                    if not is_num:
                        if cur: blocks.append(cur); cur = []
                        prev_x1 = None; continue
                    if prev_x1 is not None and (w['x0'] - prev_x1) > gap:
                        if cur: blocks.append(cur)
                        cur = [w]
                    else:
                        cur.append(w)
                    prev_x1 = w.get('x1', w['x0'] + len(w['text']) * 7)
                if cur: blocks.append(cur)
                return blocks

            blocks = _make_blocks(right_words)

            # Résoudre chaque bloc en valeur numérique
            resolved = []  # liste de (valeur, x0_du_bloc)
            for b in blocks:
                v = _uba_join_amount(b)
                if v is not None:
                    resolved.append((v, b[0]['x0']))

            if not resolved:
                i += 1; continue

            debit_amt = credit_amt = None
            nb = len(resolved)

            # SG Sénégal (SGBS) — positions mesurées sur PDF réel SGBS Jan-2026 (TRC) :
            #   Débit  : x0 ≈ 355–435   (colonne "Débit" à x0=358)
            #   Crédit : x0 ≈ 435–510   (colonne "Crédit" à x0=437)
            #   Solde  : x0 ≥ 510       (colonne "Solde" à x0=507)
            DEBIT_X_MAX  = 435   # colonne débit : [360, 435)
            CREDIT_X_MIN = 435   # colonne crédit : [435, 510)
            SOLDE_X_MIN  = 505   # colonne solde : ≥ 505 (ignorée)

            is_credit_kw = any(k in label_up for k in CREDIT_KW)
            is_debit_kw  = any(k in label_up for k in DEBIT_KW)

            # Filtrer le solde (bloc le plus à droite si x0 ≥ SOLDE_X_MIN)
            # mais seulement quand on a plus de 2 blocs
            candidates = resolved if nb <= 2 else [r for r in resolved if r[1] < SOLDE_X_MIN]
            if not candidates:
                candidates = resolved[:-1] if nb >= 2 else resolved

            if len(candidates) == 1:
                val, x0 = candidates[0]
                if x0 >= CREDIT_X_MIN:
                    credit_amt = val
                elif x0 < DEBIT_X_MAX:
                    # Position clairement dans la colonne débit
                    debit_amt = val
                else:
                    # Zone ambiguë : fallback sémantique
                    if is_credit_kw and not is_debit_kw:
                        credit_amt = val
                    else:
                        debit_amt = val

            elif len(candidates) == 2:
                val_l, x0_l = min(candidates, key=lambda x: x[1])
                val_r, x0_r = max(candidates, key=lambda x: x[1])
                if x0_l < DEBIT_X_MAX and x0_r >= CREDIT_X_MIN:
                    # Colonnes bien séparées : débit gauche, crédit droite
                    debit_amt  = val_l
                    credit_amt = val_r
                elif x0_l >= CREDIT_X_MIN:
                    # Les deux en zone crédit (montant fragmenté)
                    credit_amt = val_l  # le plus à gauche est le 1er token
                else:
                    # Les deux en zone débit
                    debit_amt = val_l

            else:
                # ≥ 3 candidats après filtrage solde : prendre les deux plus distincts
                x_min_c = min(candidates, key=lambda x: x[1])
                x_max_c = max(candidates, key=lambda x: x[1])
                if x_min_c[1] < DEBIT_X_MAX and x_max_c[1] >= CREDIT_X_MIN:
                    debit_amt  = x_min_c[0]
                    credit_amt = x_max_c[0]
                elif x_max_c[1] >= CREDIT_X_MIN:
                    credit_amt = x_max_c[0]
                else:
                    if is_credit_kw and not is_debit_kw:
                        credit_amt = x_max_c[0]
                    else:
                        debit_amt = x_min_c[0]

            # Mémo lignes suivantes
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and r2[0]['x0'] < 75 and re.match(r'^\d{2}/\d{2}/\d{4}$', r2[0]['text']):
                    break
                nl = ' '.join(w['text'] for w in r2 if 115 <= w['x0'] < 370).strip()
                if nl and not any(s in nl.upper() for s in SKIP):
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(row[0]['text'])
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    # ── Extraction du solde de clôture ─────────────────────────────────────
    full_text_sg = ' '.join(pages_text)
    for pat in [
        r'Nouveau\s+solde\s+en\s+FRANC[^0-9]+([\d][\d\s]{5,})',
        r'Nouveau\s+solde\b[^0-9]+([\d][\d\s]{5,})',
        r'NOUVEAU\s+SOLDE\b[^0-9]+([\d][\d\s]{5,})',
        r'Solde\s+au\s+\d{2}/\d{2}/\d{4}\s+([\d][\d\s]{5,})',
    ]:
        matches = list(re.finditer(pat, full_text_sg, re.IGNORECASE))
        if matches:
            # Prendre le dernier match (solde final en bas du relevé)
            raw = matches[-1].group(1).strip().split()
            candidate = ''.join(raw[:3])
            v = parse_amount(candidate)
            if v and v > 0:
                info['balance_close'] = v
                break

    return info, [t for t in txns if t is not None]
def _make_african_parser(bank_name):
    def _parser(pages_words, pages_text, _pdf_path=''):
        if _pdf_path and Path(_pdf_path).exists():
            return _universal_parse_path(_pdf_path, pages_text)
        return _afr_header(pages_text), []
    return _parser

parse_bci       = _make_african_parser('BCI')
parse_atb       = _make_african_parser('ATB')
parse_universal = _make_african_parser('Universal')


# ════════════════════════════════════════════════════════════════════════════
# CBAO — Compagnie Bancaire de l'Afrique Occidentale
# Format : Date | Valeur | Libellé | Débit (XOF) | Crédit (XOF) | Solde (XOF)
# En-tête   : "EXTRAIT DE COMPTE" + "Nom du client", "Numéro de compte"
# Positions mesurées sur relevé CBAO CB MOTORS (mars 2026) :
#   Date       : x0 < 80   (DD/MM/YYYY)
#   Valeur     : x0 ≈ 80–150 (DD/MM/YYYY) — ignoré
#   Libellé    : x0 ≈ 150–370
#   Débit      : x0 ≈ 370–460
#   Crédit     : x0 ≈ 460–550
#   Solde      : x0 ≥ 550  (ignoré)
# ════════════════════════════════════════════════════════════════════════════
def parse_cbao(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []

    SKIP = {'TOTAL', 'SOLDE', 'DATE', 'VALEUR', 'LIBELLÉ', 'LIBELLE', 'DÉBIT', 'DEBIT',
            'CRÉDIT', 'CREDIT', 'SOLDE (XOF)', 'DÉBIT (XOF)', 'CRÉDIT (XOF)',
            'EXTRAIT', 'PÉRIODE', 'CODE', 'NOM', 'PAGE', 'NUMÉRO', 'NUMERO'}
    SKIP_START = ('SOLDE', 'TOTAL', 'FRAIS/', 'FRA/COM')

    # Mots-clés sémantiques CBAO
    CREDIT_KW = {'VIREMENT RECU', 'VIRMT ORD', 'CREDIT', 'VERSEMENT', 'REMISE',
                 'RTRACTAFRIC', 'VIREMENT RTRACTAFRIC'}
    DEBIT_KW  = {'FRAIS', 'COMMISSION', 'SAISIE TRF', 'RETRAIT', 'CHEQUE',
                 'FRA/COM', 'VIREMENT W'}

    # Lecture du solde final
    full_text = ' '.join(pages_text)
    m_bal = re.search(r'Solde\s+\([A-Z]+\)\s+au\s+[\d/]+\s*:\s*([\d\s]+)', full_text, re.IGNORECASE)
    if not m_bal:
        m_bal = re.search(r'Solde\s+initial\s+\([A-Z]+\)\s*:\s*([\d\s]+)', full_text, re.IGNORECASE)
    if m_bal:
        v = parse_amount(m_bal.group(1).strip().replace(' ', ''))
        if v:
            info['balance_close'] = v

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date opération : DD/MM/YYYY en x0 < 80
            date_words = [w for w in row if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_words:
                i += 1; continue
            date_str = date_words[0]['text']

            # Libellé : x0 ≈ 150–410 (élargi)
            label_words = [w for w in row if 150 <= w['x0'] < 410]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()
            if not label or any(s in label_up for s in SKIP):
                i += 1; continue
            if any(label_up.startswith(s) for s in SKIP_START):
                i += 1; continue

            # Montants — chercher de manière adaptive dans toute la zone droite
            # CBAO CB MOTORS: les colonnes varient, on regroupe par blocs
            right_words = sorted([w for w in row if w['x0'] >= 410 and re.search(r'\d', w['text'])],
                                 key=lambda w: w['x0'])

            # Regrouper en blocs contigus (gap > 25px)
            blocs, cur, prev_x1 = [], [], None
            for w in right_words:
                is_num = bool(re.match(r'^[\d\s,\.]+$', w['text']) and re.search(r'\d', w['text']))
                if not is_num:
                    if cur: blocs.append(cur); cur = []
                    prev_x1 = None; continue
                if prev_x1 is not None and (w['x0'] - prev_x1) > 25:
                    if cur: blocs.append(cur)
                    cur = [w]
                else:
                    cur.append(w)
                prev_x1 = w.get('x1', w['x0'] + len(w['text']) * 7)
            if cur: blocs.append(cur)

            resolved = []
            for b in blocs:
                v = _uba_join_amount(b)
                if v is not None:
                    resolved.append((v, b[0]['x0']))

            debit_amt = credit_amt = None
            if len(resolved) == 0:
                i += 1; continue
            elif len(resolved) == 1:
                val, x0 = resolved[0]
                # Déduire via sémantique
                is_credit = any(k in label_up for k in CREDIT_KW)
                is_debit  = any(k in label_up for k in DEBIT_KW)
                if is_credit and not is_debit:
                    credit_amt = val
                elif is_debit and not is_credit:
                    debit_amt = val
                else:
                    # Position : x0 > 520 = crédit probable, sinon débit
                    if x0 > 520:
                        credit_amt = val
                    else:
                        debit_amt = val
            elif len(resolved) >= 2:
                # CBAO : colonnes Débit (≈410–490) | Crédit (≈490–570) | Solde (≥570)
                # Un seul montant est renseigné par ligne (l'autre est vide).
                # Stratégie : trier par x0, identifier la zone de chaque bloc.
                CBAO_DEBIT_MAX  = 510   # x0 < 510 → débit
                CBAO_CREDIT_MIN = 490   # x0 ≥ 490 → crédit possible
                CBAO_SOLDE_MIN  = 560   # x0 ≥ 560 → solde (ignorer)

                # Filtrer le solde (blocs très à droite)
                non_solde = [(v, x) for v, x in resolved if x < CBAO_SOLDE_MIN]
                if not non_solde:
                    # Tout est dans la zone solde — prendre le premier par position
                    # et déduire via sémantique
                    val, x0 = resolved[0]
                    is_credit = any(k in label_up for k in CREDIT_KW)
                    if is_credit:
                        credit_amt = val
                    else:
                        debit_amt = val
                elif len(non_solde) == 1:
                    val, x0 = non_solde[0]
                    if x0 < CBAO_DEBIT_MAX:
                        debit_amt = val
                    else:
                        credit_amt = val
                else:
                    # 2 montants hors solde → débit (gauche) et crédit (droite)
                    c_left  = min(non_solde, key=lambda x: x[1])
                    c_right = max(non_solde, key=lambda x: x[1])
                    debit_amt  = c_left[0]
                    credit_amt = c_right[0]

            # Mémo lignes suivantes
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and any(w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text']) for w in r2):
                    break
                nl = ' '.join(w['text'] for w in r2 if 150 <= w['x0'] < 410).strip()
                if nl and not any(s in nl.upper() for s in SKIP):
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ════════════════════════════════════════════════════════════════════════════
# BOA — Bank of Africa Sénégal
# Format : Date op | Description | Référence | Date valeur | Débit | Crédit | Solde courant
# En-tête  : "BANK OF AFRICA - SENEGAL"
# Positions mesurées sur relevé BOA VILLA YEMAYA (janvier 2026) :
#   Date op     : x0 < 60   (DD/MM/YY ou DD/MM/YYYY)
#   Description : x0 ≈ 60–370
#   Référence   : x0 ≈ 370–440 (ignoré)
#   Date valeur : x0 ≈ 440–490 (ignoré)
#   Débit       : x0 ≈ 490–575  (montants débits, peuvent être négatifs)
#   Crédit      : x0 ≈ 575–660
#   Solde       : x0 ≥ 660 (ignoré)
# IMPORTANT : La colonne "Solde courant" est à droite et NE DOIT PAS être utilisée
#             comme montant de transaction. On filtre strictement par position.
# ════════════════════════════════════════════════════════════════════════════
def parse_boa(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    txns = []

    SKIP = {'TOTAL', 'SOLDE', 'DATE', 'VALEUR', 'DESCRIPTION', 'RÉFÉRENCE', 'REFERENCE',
            'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT', 'SOLDE COURANT', 'BANK OF AFRICA',
            'BMCE GROUP', 'BMCE'}
    SKIP_START = ('SOLDE', 'TOTAL')

    # Mots-clés sémantiques pour déduire le sens (Crédit = entrée d'argent)
    CREDIT_KW = {'VERSEMENT', 'VIR RECU', 'VIREMENT RECU', 'VIR.RECU', 'CREDIT',
                 'REMISE', 'SWIFT', 'DEPOT', 'RECOUVREMENT'}
    DEBIT_KW  = {'RETRAIT', 'ACHAT', 'CHEQUE', 'FRAIS', 'TAXE', 'COMMISSION',
                 'DROIT TIMBRE', 'PRELEVEMENT', 'PRELEV', 'VIR BOAWEB',
                 'ABONNEMENT', 'PRIME ASSURANCE'}

    # Solde de clôture depuis le texte
    full_text = ' '.join(pages_text)
    m_close = re.search(r'Solde\s+de\s+cl[ôo]ture\s*[:\-]?\s*([\d\s,\.]+)\s*XOF', full_text, re.IGNORECASE)
    if m_close:
        v = parse_amount(m_close.group(1).replace(' ', ''))
        if v:
            info['balance_close'] = v

    def _boa_parse_col(words):
        """Parse un montant BOA dans une colonne précise.
        Retourne (valeur_absolue, est_négatif) ou (None, False)."""
        if not words:
            return None, False
        # Reconstituer le texte de la colonne
        full = ' '.join(w['text'] for w in sorted(words, key=lambda w: w['x0']))
        full = full.replace('\xa0', ' ').strip()
        neg = full.startswith('-')
        full_clean = full.lstrip('-').strip()
        # Essayer de parser en XOF (entier ou décimal, espaces comme séparateurs de milliers)
        # Formats : "200 000,00" ou "200000,00" ou "200 000" ou "200000"
        m = re.search(r'([\d][\d\s]*[\d](?:,\d{2})?)', full_clean)
        if m:
            v = parse_amount(m.group(1).replace(' ', ''))
            if v is not None and v > 0:
                return v, neg
        return None, False

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date op : x0 < 60, format DD/MM/YY ou DD/MM/YYYY
            date_words = [w for w in row if w['x0'] < 60
                          and re.match(r'^\d{2}/\d{2}/\d{2,4}$', w['text'])]
            if not date_words:
                i += 1; continue
            raw_date = date_words[0]['text']
            # Normaliser DD/MM/YY → DD/MM/YYYY
            parts = raw_date.split('/')
            if len(parts) == 3 and len(parts[2]) == 2:
                raw_date = f"{parts[0]}/{parts[1]}/20{parts[2]}"

            # Description : x0 ≈ 60–370
            label_words = [w for w in row if 60 <= w['x0'] < 370]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()
            if not label or len(label) < 3:
                i += 1; continue
            if any(label_up.startswith(s) for s in SKIP_START):
                i += 1; continue
            if any(s in label_up for s in SKIP) and len(label) < 20:
                i += 1; continue

            # ── Montants BOA ──────────────────────────────────────────────────
            # Le PDF BOA a : Débit | Crédit | Solde courant
            # On identifie les colonnes numériques à droite du libellé.
            # Stratégie : regrouper tous les mots numériques ≥ x0 355 en blocs,
            # puis affecter selon la position :
            #   Débit  : x0 355–450  (col "Débit" à x0=380)
            #   Crédit : x0 450–515  (col "Crédit" à x0=453)
            #   Solde  : x0 515+     (ignorer)
            right_words = sorted(
                [w for w in row if w['x0'] >= 355
                 and re.search(r'[\d,\.]', w['text'])
                 and not re.match(r'^\d{2}/\d{2}/\d{2,4}$', w['text'])],
                key=lambda w: w['x0']
            )

            # Regrouper en blocs (gap > 20px = nouveau bloc)
            blocs, cur, prev_x1 = [], [], None
            for w in right_words:
                if prev_x1 is not None and (w['x0'] - prev_x1) > 20:
                    if cur: blocs.append(cur)
                    cur = [w]
                else:
                    cur.append(w)
                prev_x1 = w.get('x1', w['x0'] + max(len(w['text']) * 6, 20))
            if cur: blocs.append(cur)

            # Résoudre chaque bloc en valeur + position x0
            resolved = []  # [(valeur, x0_debut, is_neg)]
            for b in blocs:
                x0_b = b[0]['x0']
                val, neg = _boa_parse_col(b)
                if val is not None:
                    resolved.append((val, x0_b, neg))

            # Mémo lignes suivantes (description multi-ligne)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and any(w['x0'] < 60 and re.match(r'^\d{2}/\d{2}/\d{2,4}$', w['text'])
                               for w in r2):
                    break
                nl = ' '.join(w['text'] for w in r2 if 60 <= w['x0'] < 370).strip()
                if nl and not any(s == nl.upper() for s in SKIP):
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(raw_date)
            name, memo = smart_label(label, memo_parts)
            label_up = label.upper()

            if not resolved:
                continue

            # ── Affecter débit / crédit selon position ────────────────────────
            # Le relevé BOA a toujours : [Débit | Crédit | Solde]
            # Un seul montant non-nul par ligne (l'autre colonne est vide).
            # Avec 1 bloc : position détermine tout (< 575 → débit, ≥ 575 → crédit)
            # Avec 2 blocs : 1er = débit/crédit, 2ème = solde (ignorer)
            # Avec 3 blocs : 1er = débit, 2ème = crédit, 3ème = solde (ignorer)
            # Seuil débit/crédit mesuré sur PDF : x0 ≈ 490–575 = débit, 575–660 = crédit
            DEBIT_X_MIN  = 355
            DEBIT_X_MAX  = 442   # en-tête Débit x0=380, tokens débit finissent avant 442
            CREDIT_X_MIN = 442   # en-tête Crédit x0=453, tokens crédit démarrent à ~445
            CREDIT_X_MAX = 515

            debit_amt = credit_amt = None

            if len(resolved) == 1:
                val, x0, neg = resolved[0]
                if DEBIT_X_MIN <= x0 < DEBIT_X_MAX:
                    # Dans la colonne débit : montant négatif = débit (sortie d'argent)
                    # Le signe '-' dans le PDF BOA indique simplement le sens
                    debit_amt = val
                elif CREDIT_X_MIN <= x0 < CREDIT_X_MAX:
                    credit_amt = val
                else:
                    # Fallback sémantique
                    is_credit = any(k in label_up for k in CREDIT_KW)
                    is_debit  = any(k in label_up for k in DEBIT_KW)
                    if is_credit and not is_debit:
                        credit_amt = val
                    else:
                        debit_amt = val

            elif len(resolved) == 2:
                # Prob: [Débit_ou_Crédit, Solde] — prendre seulement le premier
                val, x0, neg = resolved[0]
                if DEBIT_X_MIN <= x0 < DEBIT_X_MAX:
                    debit_amt = val
                elif x0 >= CREDIT_X_MIN and x0 < CREDIT_X_MAX:
                    credit_amt = val
                else:
                    # Fallback sémantique
                    is_credit = any(k in label_up for k in CREDIT_KW)
                    if is_credit:
                        credit_amt = val
                    else:
                        debit_amt = val

            elif len(resolved) >= 3:
                # [Débit, Crédit, Solde] — prendre les deux premiers
                # Trier par x0 pour être sûr
                sorted_r = sorted(resolved, key=lambda x: x[1])
                val0, x0_0, neg0 = sorted_r[0]
                val1, x0_1, neg1 = sorted_r[1]
                if x0_0 < DEBIT_X_MAX:
                    debit_amt = val0
                if x0_1 >= CREDIT_X_MIN and x0_1 < CREDIT_X_MAX:
                    credit_amt = val1

            # Émettre la transaction
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ════════════════════════════════════════════════════════════════════════════
# ORABANK Sénégal
# Format : Date | Libellé opération | Valeur | Débit | Crédit | Solde
# En-tête : "Orabank", "EXTRAIT DE COMPTE"
# Positions mesurées sur relevés Orabank AHC (déc 2025) et SUNBEAM (jan 2026) :
#   Date       : x0 < 75   (DD/MM/YYYY)
#   Libellé    : x0 ≈ 75–370
#   Valeur     : x0 ≈ 370–430 (date valeur — ignoré)
#   Débit      : x0 ≈ 430–510
#   Crédit     : x0 ≈ 510–590
#   Solde      : x0 ≥ 590  (ignoré)
# ════════════════════════════════════════════════════════════════════════════
def parse_orabank(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []

    SKIP = {'TOTAL', 'SOLDE', 'DATE', 'VALEUR', 'LIBELLÉ', 'LIBELLE', 'DÉBIT', 'DEBIT',
            'CRÉDIT', 'CREDIT', 'EXTRAIT', 'ORABANK', 'PAGE', 'RELEVÉ', 'RELEVE',
            'TITULAIRE', 'CODE', 'IBAN', 'BIC', 'AGENCE', 'DESTINATAIRE', 'COMPTE',
            'TOTAL GÉNÉRAL', 'TOTAL GENERAL'}
    SKIP_START = ('SOLDE', 'TOTAL', 'ANCIEN SOLDE', 'NOUVEAU SOLDE')

    CREDIT_KW = {'VIREMENT', 'TRF RECU', 'TRF-RECU', 'TRF RECU AUTO', 'SORT CHEQUES',
                 'VERSEMENT', 'REMISE', 'CREDIT', 'SWIFT', 'RECOUVREMENT',
                 'W FREIGHT', 'W FREIGHT PA', 'AU426', 'VIREMENT INTER AGENCE',
                 'VIREMENT W', 'CHQ RECU COMPENSE', 'REMISE CHQ INTERNE',
                 'REMISE CHQ'}
    DEBIT_KW  = {'LUFTHANSA', 'RETRAIT', 'FRAIS', 'COMMISSION', 'AGIOS',
                 'ABONNEMENT', 'PACK PRO', 'CHEQUE', 'RETRAIT ESPECES',
                 'RETRAIT DAB', 'RETRAIT GAB', 'CHQ N.', 'RET. GAB',
                 'PAIEMENT VISA', 'PAIEMENT HORS', 'VIRMT FAV.'}

    full_text = ' '.join(pages_text)
    # "Solde (XOF) au 31/01/2026 : 5 272 384"  ou  "Solde au 31/01/2026 : ..."
    for bal_pat in [
        r'Solde\s+\([A-Z]+\)\s+au\s+[\d/]+\s*:\s*([\d\s]+)',
        r'Solde\s+au\s*[:\-]?\s*[\d/]+\s+([\d\s]+)',
    ]:
        m_bal = re.search(bal_pat, full_text, re.IGNORECASE)
        if m_bal:
            v = parse_amount(re.sub(r'\s+', '', m_bal.group(1).strip()))
            if v and v > 0:
                info['balance_close'] = v
                break

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date : x0 < 80, DD/MM/YYYY
            date_words = [w for w in row if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_words:
                i += 1; continue
            date_str = date_words[0]['text']

            # Libellé : x0 ≈ 75–400 (élargi pour capturer les deux formats)
            label_words = [w for w in row if 75 <= w['x0'] < 400
                           and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()

            # Si libellé vide ou trop court sur la ligne de date,
            # chercher dans les lignes précédentes sans date (pattern Orabank SUNBEAM/BNDE)
            if not label or len(label) < 3:
                for k in range(i - 1, max(i - 6, -1), -1):
                    r2 = rows[k]
                    has_date2 = any(w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                                    for w in r2)
                    if has_date2:
                        break
                    adj_words = [w for w in r2 if 75 <= w['x0'] < 400
                                 and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                    adj = ' '.join(w['text'] for w in adj_words).strip()
                    if adj and len(adj) >= 3 and re.search(r'[A-Za-zÀ-ÿ]{2,}', adj):
                        label = adj
                        break
                label_up = label.upper()

            if not label or len(label) < 3:
                i += 1; continue
            if any(label_up.startswith(s) for s in SKIP_START):
                i += 1; continue
            if any(s == label_up for s in SKIP):
                i += 1; continue

            # ── Montants : stratégie adaptive ────────────────────────────────
            # Positions mesurées sur PDF Orabank réel (SUNBEAM Jan-2026) :
            #   Débit  header x0=342, tokens à x0 ≈ 330–420
            #   Crédit header x0=421, tokens à x0 ≈ 420–500
            #   Solde  header x0=503, tokens à x0 ≥ 500 (ignorer)
            DEBIT_X_MAX  = 420
            CREDIT_X_MIN = 420
            CREDIT_X_MAX = 500
            SOLDE_X_MIN  = 500

            right_words = sorted(
                [w for w in row if w['x0'] >= 330
                 and re.search(r'\d', w['text'])
                 and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])],
                key=lambda w: w['x0']
            )

            # Regrouper en blocs contigus (gap > 20px)
            blocs, cur, prev_x1 = [], [], None
            for w in right_words:
                if prev_x1 is not None and (w['x0'] - prev_x1) > 20:
                    if cur: blocs.append(cur)
                    cur = [w]
                else:
                    cur.append(w)
                prev_x1 = w.get('x1', w['x0'] + max(len(w['text']) * 6, 15))
            if cur: blocs.append(cur)

            resolved = []
            for b in blocs:
                v = _uba_join_amount(b)
                if v is not None:
                    resolved.append((v, b[0]['x0']))

            debit_amt = credit_amt = None
            non_solde = [(v, x) for v, x in resolved if x < SOLDE_X_MIN]

            if len(non_solde) == 1:
                val, x0 = non_solde[0]
                if x0 < DEBIT_X_MAX:
                    debit_amt = val
                elif CREDIT_X_MIN <= x0 < CREDIT_X_MAX:
                    credit_amt = val
                else:
                    # Ambiguïté positionnelle → sémantique
                    is_cr = any(k in label_up for k in CREDIT_KW)
                    is_db = any(k in label_up for k in DEBIT_KW)
                    if is_cr and not is_db:
                        credit_amt = val
                    elif is_db and not is_cr:
                        debit_amt = val
                    else:
                        # Défaut : débit si dans la moitié gauche, crédit sinon
                        if x0 < (DEBIT_X_MAX + CREDIT_X_MAX) / 2:
                            debit_amt = val
                        else:
                            credit_amt = val
            elif len(non_solde) >= 2:
                c_left  = min(non_solde, key=lambda x: x[1])
                c_right = max(non_solde, key=lambda x: x[1])
                if c_left[1] < DEBIT_X_MAX and c_right[1] >= CREDIT_X_MIN:
                    debit_amt  = c_left[0]
                    credit_amt = c_right[0]
                elif c_right[1] >= CREDIT_X_MIN:
                    credit_amt = c_right[0]
                elif c_left[1] < DEBIT_X_MAX:
                    debit_amt = c_left[0]
                else:
                    is_cr = any(k in label_up for k in CREDIT_KW)
                    if is_cr:
                        credit_amt = c_right[0]
                    else:
                        debit_amt = c_left[0]
            elif resolved:
                val, x0 = resolved[0]
                is_cr = any(k in label_up for k in CREDIT_KW)
                is_db = any(k in label_up for k in DEBIT_KW)
                if is_cr and not is_db:
                    credit_amt = val
                elif is_db and not is_cr:
                    debit_amt = val

            # Mémo lignes suivantes
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and any(w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                               for w in r2):
                    break
                nl = ' '.join(w['text'] for w in r2 if 75 <= w['x0'] < 400).strip()
                if nl and not any(s in nl.upper() for s in SKIP):
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ════════════════════════════════════════════════════════════════════════════
# NSIA Banque (Sénégal)
# Format : Date Transact | Détails Transaction | Cheque N° | Agence |
#          Date Valeur | Mouv. Débit | Mouv. Crédit | Solde
# En-tête : "NSIA BANQUE", "RELEVE DE COMPTE"
# Positions mesurées sur relevés NSIA IROKO BEACH et AAG SENEGAL (mars 2026) :
#   Date Transact : x0 < 80   (DD/MM/YYYY)
#   Détails       : x0 ≈ 80–380
#   Cheque N°     : x0 ≈ 200–280 (optionnel)
#   Agence        : x0 ≈ 280–320 (ignoré)
#   Date Valeur   : x0 ≈ 320–390 (ignoré)
#   Mouv. Débit   : x0 ≈ 390–475
#   Mouv. Crédit  : x0 ≈ 475–560
#   Solde         : x0 ≥ 560  (ignoré)
# ════════════════════════════════════════════════════════════════════════════
def parse_nsia(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    year = _year_from_text(' '.join(pages_text[:2]))
    txns = []

    SKIP = {'TOTAL', 'SOLDE', 'DATE', 'VALEUR', 'DÉTAILS', 'DETAILS', 'DÉBIT', 'DEBIT',
            'CRÉDIT', 'CREDIT', 'CHEQUE', 'AGENCE', 'TRANSACTION', 'MOUV',
            'RELEVE', 'RELEVÉ', 'NSIA', 'PAGE', 'TITULAIRE', 'IBAN', 'BIC',
            'COMPTE', 'NUMÉRO', 'NUMERO', 'DEVISE', 'TOTAL DEBIT', 'TOTAL CREDIT'}
    SKIP_START = ('SOLDE', 'TOTAL DEBIT', 'TOTAL CREDIT', 'CHERS', 'VEUILLEZ',
                  'SEN/BANM', 'PAGE ', 'OPID:')

    # Mots-clés sémantiques NSIA
    CREDIT_KW = {'VIREMENT PERMANENT', 'VIREMENT', 'REMISE', 'CREDIT', 'VERSEMENT',
                 'SWIFT', 'RECOUVREMENT', 'SV CLEARING', '-PURCHASE'}
    DEBIT_KW  = {'PAIEMENT DES CHQ', 'RETRAIT PAR CHQ', 'RETRAIT', 'FRAIS',
                 'COMMISSION MANUELLE', 'TAF SUR', 'TAX SUR', 'FRAIS DE TENUE',
                 'FRAIS DE GESTION', 'MAINTENANCE', 'PRELEV', 'PRELEVEMENT'}

    # Solde final depuis le texte
    full_text = ' '.join(pages_text)
    # Format NSIA : "SOLDE  1 136 108" en fin de relevé, ou "Solde Début Période  8 066 238"
    for pat in [
        r'SOLDE\s+(\d[\d\s]{3,})\s*$',
        r'SOLDE\s*\n\s*(\d[\d\s]{3,})',
        r'Solde\s+D[ée]but\s+P[ée]riode\s+([\d\s]+)',
        r'Solde\s+Final\s*[:\-]?\s*([\d\s]+)',
    ]:
        m = re.search(pat, full_text, re.IGNORECASE | re.MULTILINE)
        if m:
            v = parse_amount(m.group(1).strip().replace(' ', ''))
            if v:
                info['balance_close'] = v
                break

    # Récupérer aussi le solde depuis la dernière ligne "SOLDE XXXX" du tableau
    solde_lines = re.findall(r'SOLDE\s+([\d][\d\s]{2,})', full_text, re.IGNORECASE)
    if solde_lines:
        last_v = parse_amount(solde_lines[-1].strip().replace(' ', ''))
        if last_v and last_v > 0:
            info['balance_close'] = last_v

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            # Date Transact : x0 < 80, DD/MM/YYYY
            date_words = [w for w in row if w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_words:
                i += 1; continue
            date_str = date_words[0]['text']

            # Détails Transaction : x0 ≈ 80–380 (élargi pour capturer descriptions longues)
            # Le libellé NSIA est souvent sur PLUSIEURS lignes précédant la ligne de date.
            # La date est sur une ligne qui contient aussi le n° agence et les montants.
            label_words = [w for w in row if 80 <= w['x0'] < 380
                           and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                           and not re.match(r'^\d{2,3}$', w['text'])]  # exclure n° agence (2-3 chiffres)
            label = ' '.join(w['text'] for w in label_words).strip()

            # Si libellé vide ou trop court sur la ligne de date, chercher dans les lignes précédentes
            if (not label or len(label) < 3) and i > 0:
                # Chercher jusqu'à 5 lignes en arrière
                parts = []
                for k in range(i - 1, max(i - 6, -1), -1):
                    prev_row = rows[k]
                    prev_date = [w for w in prev_row if w['x0'] < 80
                                 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                    if prev_date:
                        break  # autre transaction → stop
                    prev_label_words = [w for w in prev_row if 80 <= w['x0'] < 380
                                        and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                                        and not re.match(r'^\d{2,3}$', w['text'])]
                    pl = ' '.join(w['text'] for w in prev_label_words).strip()
                    if pl and not any(s in pl.upper() for s in SKIP):
                        parts.insert(0, pl)
                if parts:
                    label = ' '.join(parts)

            label_up = label.upper()
            if not label or len(label) < 3:
                i += 1; continue
            if any(label_up.startswith(s) for s in SKIP_START):
                i += 1; continue
            if label_up in SKIP:
                i += 1; continue

            # Mouv. Débit : x0 ≈ 380–450
            # Mouv. Crédit : x0 ≈ 450–525
            # Solde : x0 ≥ 525 (ignorer absolument)
            # Stratégie : regrouper tous les mots numériques en blocs, puis
            # affecter selon la position x0 de chaque bloc.
            # IMPORTANT : ne pas filtrer les tokens à 3 chiffres comme '000' car
            # ils font partie des montants (ex: '750' + '000' = 750 000)
            right_words = sorted(
                [w for w in row if w['x0'] >= 380
                 and re.search(r'^\d+$', w['text'])  # uniquement les tokens purement numériques
                 and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])],
                key=lambda w: w['x0']
            )
            # Exclure le numéro d'agence (3-4 chiffres, x0 < 340) uniquement dans la zone libellé
            # Les agences NSIA ont un numéro court (ex: 311, 301, 307) à x0 ≈ 285-300
            # Ils sont déjà exclus car on commence à x0 >= 380

            # Regrouper en blocs contigus (gap > 18px = nouveau bloc)
            nsia_blocs, cur_b, prev_x1_b = [], [], None
            for w in right_words:
                if prev_x1_b is not None and (w['x0'] - prev_x1_b) > 18:
                    if cur_b: nsia_blocs.append(cur_b)
                    cur_b = [w]
                else:
                    cur_b.append(w)
                prev_x1_b = w.get('x1', w['x0'] + max(len(w['text']) * 6, 15))
            if cur_b: nsia_blocs.append(cur_b)

            # Résoudre chaque bloc en (valeur, x0_debut)
            nsia_resolved = []
            for b in nsia_blocs:
                x0_b = b[0]['x0']
                v = _uba_join_amount(b)
                if v is not None and v > 0:
                    nsia_resolved.append((v, x0_b))

            NSIA_DEBIT_MIN   = 380
            NSIA_DEBIT_MAX   = 450
            NSIA_CREDIT_MIN  = 450
            NSIA_CREDIT_MAX  = 525
            NSIA_SOLDE_MIN   = 525   # ignorer tout ce qui est ≥ 525

            debit_amt  = None
            credit_amt = None

            # Stratégie NSIA : séparer par position absolue (pas par gap)
            # car les tokens crédit et solde peuvent être très proches
            debit_words_nsia  = [w for w in right_words if NSIA_DEBIT_MIN <= w['x0'] < NSIA_DEBIT_MAX]
            credit_words_nsia = [w for w in right_words if NSIA_CREDIT_MIN <= w['x0'] < NSIA_CREDIT_MAX]

            debit_amt  = _uba_join_amount(debit_words_nsia)  if debit_words_nsia  else None
            credit_amt = _uba_join_amount(credit_words_nsia) if credit_words_nsia else None

            # Filtrer les montants nuls
            if debit_amt is not None and debit_amt <= 0:
                debit_amt = None
            if credit_amt is not None and credit_amt <= 0:
                credit_amt = None

            # Reconstruire nsia_resolved pour le fallback
            nsia_resolved_all = []
            for b_words, x0_b in [(debit_words_nsia, NSIA_DEBIT_MIN),
                                   (credit_words_nsia, NSIA_CREDIT_MIN)]:
                if b_words:
                    v = _uba_join_amount(b_words)
                    if v and v > 0:
                        nsia_resolved_all.append((v, b_words[0]['x0']))
            nsia_resolved = nsia_resolved_all
            non_solde = nsia_resolved_all

            # Mémo lignes suivantes (continuation de la description)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if r2 and any(w['x0'] < 80 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                               for w in r2):
                    break
                nl_words = [w for w in r2 if 80 <= w['x0'] < 380
                            and not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])
                            and not re.match(r'^\d{3,4}$', w['text'])]
                nl = ' '.join(w['text'] for w in nl_words).strip()
                if nl and not any(s in nl.upper() for s in SKIP) and len(nl) > 2:
                    memo_parts.append(nl)
                j += 1
            i = j

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, memo_parts)

            # ── Règle SV CLEARING NSIA ────────────────────────────────────────
            # Le terminal TPE NSIA génère TOUJOURS 2 ou 3 lignes par transaction :
            #   1. :PURCHASE  OPID=xxxx → frais réseau (comm ou TVA) → IGNORER TOUJOURS
            #   2. :PURCHASE  OPID=xxxx → TVA frais réseau          → IGNORER TOUJOURS
            #   3. :-PURCHASE OPID=xxxx → montant principal          → GARDER (crédit)
            # RÈGLE ABSOLUE : toute ligne SV CLEARING avec ":PURCHASE" mais SANS ":-PURCHASE"
            # est un frais réseau à ignorer, quel que soit le montant.
            # Seules les lignes ":-PURCHASE" (avec tiret) sont des crédits à conserver.
            is_tpe_frais = False
            if 'SV CLEARING' in label_up or 'CLEARING' in label_up:
                is_minus_purchase = bool(re.search(r':-\s*PURCHASE', label_up))
                if not is_minus_purchase:
                    # Toute ligne :PURCHASE (sans tiret) du SV CLEARING = frais réseau → IGNORER
                    is_tpe_frais = True

            if is_tpe_frais:
                pass  # Ignorer les lignes de frais/commissions TPE réseau
            elif debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))
            else:
                # Fallback sémantique : si aucun bloc dans les zones exactes,
                # prendre le premier bloc non-solde et utiliser les mots-clés
                if non_solde:
                    amt = non_solde[0][0]
                    is_credit = any(k in label_up for k in CREDIT_KW)
                    is_debit  = any(k in label_up for k in DEBIT_KW)
                    # SV CLEARING :-PURCHASE = entrée (crédit) — garder toujours
                    if re.search(r':-\s*PURCHASE', label_up):
                        txns.append(_make_txn(date_ofx, amt, name, memo))
                    elif ('SV CLEARING' in label_up or 'CLEARING' in label_up):
                        pass  # :PURCHASE sans tiret = frais réseau → ignorer
                    elif is_credit and not is_debit:
                        txns.append(_make_txn(date_ofx, amt, name, memo))
                    elif is_debit and not is_credit:
                        txns.append(_make_txn(date_ofx, -amt, name, memo))
                else:
                    # Dernier recours : blocs résolus complets (incluant solde)
                    if nsia_resolved:
                        amt = nsia_resolved[0][0]
                        is_credit = any(k in label_up for k in CREDIT_KW)
                        is_debit  = any(k in label_up for k in DEBIT_KW)
                        if re.search(r':-\s*PURCHASE', label_up):
                            txns.append(_make_txn(date_ofx, amt, name, memo))
                        elif ('SV CLEARING' in label_up or 'CLEARING' in label_up):
                            pass  # :PURCHASE sans tiret = frais réseau → ignorer
                        elif is_credit and not is_debit:
                            txns.append(_make_txn(date_ofx, amt, name, memo))
                        elif is_debit and not is_credit:
                            txns.append(_make_txn(date_ofx, -amt, name, memo))

    # ── Fallback texte brut pour les opérations SV CLEARING (format IROKO) ──
    # Structure : ligne N = libellé (ex: "VIREMENT. SV CLEARING :-PURCHASE")
    #             ligne N+1 = "DD/MM/YYYY OPID:... TERM:... REF AGENCE DD/MM/YYYY MONTANT SOLDE"
    # Ce fallback ne s'exécute QUE si le parseur word-by-row n'a capturé aucune transaction
    # valide, pour éviter les doublons sur les relevés comme AAG SENEGAL.
    txns_real = [t for t in txns if t]
    if not txns_real:
        full_text_nsia = '\n'.join(pages_text)
        lines_nsia = full_text_nsia.splitlines()

        for idx, line in enumerate(lines_nsia):
            line = line.strip()
            if idx == 0:
                continue

            # Ligne de transaction : commence par DD/MM/YYYY
            m_date = re.match(r'^(\d{2}/\d{2}/\d{4})\s', line)
            if not m_date:
                continue

            date_str = m_date.group(1)

            # Chercher le libellé dans les lignes précédentes (jusqu'à 5 en arrière)
            label_raw = ''
            for k in range(idx - 1, max(idx - 6, -1), -1):
                prev = lines_nsia[k].strip()
                if not prev:
                    continue
                # Stop si c'est une ligne de date (autre transaction)
                if re.match(r'^\d{2}/\d{2}/\d{4}\s', prev):
                    break
                # Garder si contient des lettres (libellé)
                if re.search(r'[A-Za-zÀ-ÿ]{3,}', prev):
                    label_raw = prev
                    break

            if not label_raw:
                continue
            label_up_raw = label_raw.upper()
            if any(s in label_up_raw for s in SKIP_START):
                continue
            if not any(k in label_up_raw for k in list(CREDIT_KW) + list(DEBIT_KW)
                       + [':-PURCHASE', ':PURCHASE', 'CLEARING']):
                continue

            # Extraire la date valeur et les tokens numériques après
            rest = line[m_date.end():]
            # Retirer OPID:... TERM:... si présents
            rest = re.sub(r'OPID:\s*\S+\s*', '', rest)
            rest = re.sub(r'TERM:\s*\S+\s*', '', rest)
            # Retirer la référence (alphanumérique long)
            rest = re.sub(r'\b[A-Z0-9]{8,}\b\s*', '', rest)
            # Retirer le numéro d'agence (3-4 chiffres seuls)
            rest = re.sub(r'\b\d{3,4}\b\s+', '', rest, count=1)
            # Retirer la date valeur DD/MM/YYYY
            rest = re.sub(r'\d{2}/\d{2}/\d{4}\s*', '', rest)

            # Extraire tous les tokens numériques restants
            num_tokens = re.findall(r'\d+', rest)
            if not num_tokens:
                continue

            # Reconstituer montant et solde :
            # En NSIA XOF, les montants sont typiquement :
            #   - montant TPE : 5000–500000 (1-3 tokens)
            #   - solde : 100000–50000000 (3-9 tokens)
            # Stratégie : le solde occupe les N derniers tokens (typiquement 3 tokens)
            # et le montant occupe les M premiers tokens.
            # On teste toutes les coupures et on prend celle qui donne
            # un montant raisonnable ET un solde raisonnable.
            def _smart_recompose(tokens):
                n = len(tokens)
                for cut in range(1, n):
                    m_str = ''.join(tokens[:cut])
                    s_str = ''.join(tokens[cut:])
                    try:
                        mv = float(m_str)
                        sv = float(s_str)
                        # Critères : montant 100–5000000, solde 100000–100000000
                        if 100 <= mv <= 5_000_000 and 100_000 <= sv <= 100_000_000:
                            return mv, sv
                    except Exception:
                        pass
                # Fallback : prendre le premier token comme montant
                if n >= 2:
                    try:
                        mv = float(''.join(tokens[:min(3, n-1)]))
                        if 100 <= mv <= 5_000_000:
                            return mv, None
                    except Exception:
                        pass
                return None, None

            amount_val, _ = _smart_recompose(num_tokens)
            if not amount_val or amount_val <= 0:
                continue

            # Déterminer débit/crédit
            # :-PURCHASE (avec tiret) = crédit TPE principal → garder
            # :PURCHASE  (sans tiret) = frais réseau TPE     → IGNORER toujours
            _has_minus_purchase = bool(re.search(r':-\s*PURCHASE', label_up_raw))
            _is_sv_clearing = ('SV CLEARING' in label_up_raw or 'CLEARING' in label_up_raw)
            _has_plain_purchase = (':PURCHASE' in label_up_raw and not _has_minus_purchase)
            if _is_sv_clearing and _has_plain_purchase:
                continue  # frais réseau TPE → ignorer
            is_credit_raw = (_has_minus_purchase or
                             any(k in label_up_raw for k in CREDIT_KW))
            is_debit_raw  = any(k in label_up_raw for k in DEBIT_KW) if not is_credit_raw else False

            date_ofx = date_full_to_ofx(date_str)
            name_raw, memo_raw = smart_label(label_raw, [])

            # Éviter les doublons
            sign = amount_val if is_credit_raw else -amount_val
            fitid_raw = make_fitid(date_ofx, label_raw, sign)
            if any(t and t.get('fitid') == fitid_raw for t in txns):
                continue

            if is_credit_raw and not is_debit_raw:
                txns.append(_make_txn(date_ofx, amount_val, name_raw, memo_raw))
            elif is_debit_raw and not is_credit_raw:
                txns.append(_make_txn(date_ofx, -amount_val, name_raw, memo_raw))

    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]


# ── CORIS BANK : format Date | Libellé | Valeur | Débit | Crédit | Solde
# Montants XOF entiers (sans décimales), dates DD/MM/YYYY.
# Positions mesurées sur relevé Coris Bank Sénégal réel :
#   Date      : x0 < 75
#   Libellé   : x0 ≈ 75–310
#   Valeur    : x0 ≈ 310–390  (date valeur — on l'ignore)
#   Débit     : x0 ≈ 390–470
#   Crédit    : x0 ≈ 470–555
#   Solde     : x0 ≥ 555
def parse_coris(pages_words, pages_text, _pdf_path=''):
    info = _afr_header(pages_text)
    txns = []

    SKIP = {'SOLDE', 'TOTAL', 'TOTAUX', 'DATE', 'LIBELLÉ', 'LIBELLE', 'VALEUR',
            'DÉBIT', 'DEBIT', 'CRÉDIT', 'CREDIT', 'REPORT', 'A REPORTER',
            'NOMBRE', 'MOUVEMENTS', 'ANCIEN', 'NOUVEAU', 'EXTRAIT', 'COMPTE',
            'COMPTES COURANTS', 'RELEVE', 'RELEVÉ', 'PRÉCÉDENT', 'PRECEDENT'}

    CREDIT_KW = {'VERSEMENT', 'VERS.', 'VRT RECU', 'VIREMENT RECU', 'CREDIT',
                 'REMISE', 'SWIFT', 'DÉBLOCAGE', 'DEBLOCAGE', 'RECOUVREMENT',
                 'ANNULATION', 'RETROCESSION'}
    DEBIT_KW  = {'RET ', 'RETRAIT', 'CHEQUE', 'CHEQ', 'FRAIS', 'COMMISSION',
                 'VRT EMIS', 'VIREMENT EMIS', 'PRELEVEMENT', 'AGIOS',
                 'ABONNEMENT', 'IMPAYE', 'PACKAGE', 'CAUTION'}

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]

            # Date : DD/MM/YYYY en x0 < 75
            date_w = [w for w in row
                      if w['x0'] < 75 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
            if not date_w:
                i += 1; continue

            date_str = date_w[0]['text']

            # Libellé : x0 entre 75 et 310
            label_words = [w for w in row if 75 <= w['x0'] < 310]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()

            if not label or any(s in label_up for s in SKIP):
                i += 1; continue

            # Collecter les lignes de suite (mémo, libellé multi-lignes)
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                # Stopper si nouvelle ligne de date
                if r2 and r2[0]['x0'] < 75 and re.match(r'^\d{2}/\d{2}/\d{4}$', r2[0]['text']):
                    break
                nl = ' '.join(w['text'] for w in r2 if 75 <= w['x0'] < 310).strip()
                if nl and not any(s in nl.upper() for s in SKIP):
                    memo_parts.append(nl)
                j += 1
            i = j

            # --- Montants ---
            # Positions mesurées sur PDF réel Coris Bank Sénégal :
            #   Débit  : x0 ≈ 363–420  (ex: '2','100','000' → 2 100 000)
            #   Crédit : x0 ≈ 445–495  (ex: '12','940','000' → 12 940 000)
            #   Solde  : x0 ≈ 520–575  (ex: '27','010','116' → ignoré)
            debit_words  = [w for w in row if 355 <= w['x0'] < 430]
            credit_words = [w for w in row if 440 <= w['x0'] < 510]

            debit_amt  = _uba_join_amount(debit_words)
            credit_amt = _uba_join_amount(credit_words)

            # Fallback sémantique si les colonnes se chevauchent ou sont absentes
            if debit_amt is None and credit_amt is None:
                # Tous les mots numériques à droite
                right_words = sorted([w for w in row if w['x0'] >= 380],
                                     key=lambda w: w['x0'])
                # Exclure les dates (valeur)
                right_words = [w for w in right_words
                               if not re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])]
                if right_words:
                    # Regrouper en blocs (gap > 25px)
                    blocs, cur, prev_x1 = [], [], None
                    for w in right_words:
                        is_num = bool(re.match(r'^[\d\s,\.]+$', w['text'])
                                      and re.search(r'\d', w['text']))
                        if not is_num:
                            if cur: blocs.append(cur); cur = []
                            prev_x1 = None; continue
                        if prev_x1 is not None and (w['x0'] - prev_x1) > 25:
                            if cur: blocs.append(cur)
                            cur = [w]
                        else:
                            cur.append(w)
                        prev_x1 = w.get('x1', w['x0'] + len(w['text']) * 7)
                    if cur: blocs.append(cur)

                    resolved = [(v, b[0]['x0']) for b in blocs
                                for v in [_uba_join_amount(b)] if v is not None]

                    if resolved:
                        is_credit = any(k in label_up for k in CREDIT_KW)
                        is_debit  = any(k in label_up for k in DEBIT_KW)

                        if len(resolved) == 1:
                            val, x0 = resolved[0]
                            if is_credit and not is_debit:
                                credit_amt = val
                            elif is_debit and not is_credit:
                                debit_amt = val
                            # sinon ambigu sans position fiable
                        elif len(resolved) >= 2:
                            # Exclure le solde (dernier bloc, le plus à droite)
                            candidates = resolved[:-1]
                            if len(candidates) == 1:
                                val, x0 = candidates[0]
                                if x0 >= 440:
                                    credit_amt = val
                                else:
                                    debit_amt = val
                            else:
                                c_left  = min(candidates, key=lambda x: x[1])
                                c_right = max(candidates, key=lambda x: x[1])
                                if is_credit and not is_debit:
                                    credit_amt = c_right[0]
                                elif is_debit and not is_credit:
                                    debit_amt = c_left[0]
                                else:
                                    if c_right[1] >= 440:
                                        credit_amt = c_right[0]
                                    else:
                                        debit_amt = c_left[0]

            date_ofx = date_full_to_ofx(date_str)
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    # Fallback universel si rien trouvé
    if not txns and _pdf_path and Path(_pdf_path).exists():
        return _universal_parse_path(_pdf_path, pages_text)
    return info, [t for t in txns if t is not None]

def parse_ecobank(pages_words, pages_text, _pdf_path=''):
    """
    Parser dédié Ecobank Sénégal.
    Supporte deux formats :
      - Format anglais (Mega Max) : Date "30-May-2025", "Account Number", colonnes en pts
      - Format français            : Date DD/MM/YY, "Numéro de compte"
    Montants XOF entiers fragmentés. Pas d'IBAN : numéro de compte brut.
    """
    info = _afr_header(pages_text)
    full_text = ' '.join(pages_text)

    # ── Numéro de compte : supporte "Account Number" (anglais) et "Numéro de compte" ──
    if not info.get('iban') and not info.get('_rib_account'):
        m_cpte = re.search(
            r'(?:Account\s+Number|Num[eé]ro\s+de\s+compte)\s*[:\-]?\s*([\d]{6,20})',
            full_text, re.IGNORECASE
        )
        if m_cpte:
            num = m_cpte.group(1).strip()
            info['iban']         = num
            info['_rib_bank']    = '00000'
            info['_rib_agency']  = '00000'
            info['_rib_account'] = num
            info['_rib_key']     = ''

    # ── Période : formats français "Du 01/01/2025 Au 31/01/2025"
    #             et anglais "Statement From Date 01-05-2025 Statement To Date 31-05-2025" ──
    if not info.get('period_start'):
        # Format français
        m_per = re.search(
            r'(?:Du|du)\s+(\d{2}/\d{2}/\d{4})\s+(?:Au|au)\s+(\d{2}/\d{2}/\d{4})',
            full_text
        )
        if m_per:
            info['period_start'] = m_per.group(1)
            info['period_end']   = m_per.group(2)
        else:
            # Format anglais Ecobank : "01-05-2025"
            m_per2 = re.search(
                r'Statement\s+From\s+Date\s+(\d{2}-\d{2}-\d{4}).*?Statement\s+To\s+Date\s+(\d{2}-\d{2}-\d{4})',
                full_text, re.IGNORECASE
            )
            if m_per2:
                def _eco_reformat(s):
                    p = s.split('-')
                    return f"{p[0]}/{p[1]}/{p[2]}" if len(p)==3 else s
                info['period_start'] = _eco_reformat(m_per2.group(1))
                info['period_end']   = _eco_reformat(m_per2.group(2))

    # ── Solde de clôture : "Closing Balance XOF6,095,325.00" ou "Solde de clôture …" ──
    m_bal_en = re.search(r'Closing\s+Balance\s+(?:XOF)?([\d,\.]+)', full_text, re.IGNORECASE)
    if m_bal_en:
        raw = m_bal_en.group(1).replace(',', '')
        try: info['balance_close'] = float(raw)
        except ValueError: pass
    if not info.get('balance_close'):
        m_bal = re.search(r'Solde\s+de\s+cl[oô]ture\s+([\d\s]+)', full_text, re.IGNORECASE)
        if m_bal:
            raw = re.sub(r'\s+', '', m_bal.group(1))
            try: info['balance_close'] = float(raw)
            except ValueError: pass

    # ── Solde d'ouverture : "Opening Balance XOF8,383,905.00" ──
    m_open_en = re.search(r'Opening\s+Balance\s+(?:XOF)?([\d,\.]+)', full_text, re.IGNORECASE)
    if m_open_en:
        raw = m_open_en.group(1).replace(',', '')
        try: info['balance_open'] = float(raw)
        except ValueError: pass

    txns = []
    year = _year_from_text(full_text)

    # Mois anglais → numéro
    _MONTH_EN = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
                 'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}

    SKIP_UP = {'TOTAL','SOLDE','DATE','DATEVAL','TRANSACTION',
               'DÉBIT','DEBIT','CRÉDIT','CREDIT','PERIODE','PÉRIODE',
               'PAYMENTS','DEPOSITS','BALANCE','OPENING','CLOSING',
               'ACCOUNT','STATEMENT','DESCRIPTION','VALUE','REFERENCE'}

    def _eco_date(row):
        """Détecte une date en colonne gauche.
        Supporte DD/MM/YY, DD/MM/YYYY (format français)
        et DD-Mon-YYYY comme '30-May-2025' (format anglais Ecobank)."""
        for w in row:
            if w['x0'] < 90:
                # Format français
                if re.match(r'^\d{2}/\d{2}/(\d{2}|\d{4})$', w['text']):
                    return ('fr', w['text'])
                # Format anglais : "30-May-2025"
                m = re.match(r'^(\d{2})-([A-Za-z]{3})-(\d{4})$', w['text'])
                if m:
                    return ('en', w['text'])
        return None

    def _eco_date_ofx(date_info):
        mode, ds = date_info
        if mode == 'fr':
            p = ds.split('/')
            if len(p) == 3:
                dd, mm, yy = p
                yr = (2000+int(yy)) if len(yy)==2 and int(yy)<=30 else (1900+int(yy)) if len(yy)==2 else int(yy)
                return f"{yr}{mm.zfill(2)}{dd.zfill(2)}"
        elif mode == 'en':
            p = ds.split('-')
            if len(p) == 3:
                dd, mon, yyyy = p
                mm = _MONTH_EN.get(mon[:3].capitalize(), '01')
                return f"{yyyy}{mm}{dd.zfill(2)}"
        return str(year)+'0101'

    for pw in pages_words:
        rows = group_words_by_row(pw, tol=4.0)
        i = 0
        while i < len(rows):
            row = rows[i]
            date_info = _eco_date(row)
            if not date_info:
                i += 1; continue
            label_words = [w for w in row if 90 <= w['x0'] < 350]
            label = ' '.join(w['text'] for w in label_words).strip()
            label_up = label.upper()
            if not label or any(s in label_up for s in SKIP_UP) or re.match(r'^[\d\s/\-]+$', label):
                i += 1; continue
            memo_parts = []
            j = i + 1
            while j < len(rows):
                r2 = rows[j]
                if _eco_date(r2): break
                cont = ' '.join(w['text'] for w in r2 if 90 <= w['x0'] < 350).strip()
                if cont and not any(s in cont.upper() for s in SKIP_UP):
                    memo_parts.append(cont)
                j += 1
            i = j

            # ── Colonnes montants ────────────────────────────────────────────
            # Format français :  Débit x0≈390–465 / Crédit x0≈465–535
            # Format anglais  :  Payments x0≈420–510 / Deposits x0≈510–590
            # On essaie les deux fenêtres et on garde la plus large
            if date_info[0] == 'en':
                debit_words  = [w for w in row if 420 <= w['x0'] < 510]
                credit_words = [w for w in row if 510 <= w['x0'] < 595]
            else:
                debit_words  = [w for w in row if 390 <= w['x0'] < 465]
                credit_words = [w for w in row if 465 <= w['x0'] < 535]

            # Montants en format anglais : "XOF6,500.00" → parse_amount gère déjà
            def _eco_parse_en(words):
                full = ' '.join(w['text'] for w in words).strip()
                # Retirer le préfixe devise "XOF"
                full = re.sub(r'^XOF', '', full, flags=re.IGNORECASE).strip()
                # Format "6,500.00" → remplacer virgule-milliers, point-décimal
                full = full.replace(',', '')
                try:
                    v = float(full)
                    return v if v > 0 else None
                except ValueError:
                    return None

            if date_info[0] == 'en':
                debit_amt  = _eco_parse_en(debit_words)
                credit_amt = _eco_parse_en(credit_words)
            else:
                debit_amt  = _uba_join_amount(debit_words)
                credit_amt = _uba_join_amount(credit_words)
                # Montant négatif dans colonne débit = remboursement (crédit)
                if debit_amt is None:
                    raw_d = ' '.join(w['text'] for w in debit_words)
                    if '- ' in raw_d or raw_d.strip().startswith('-'):
                        nums = re.sub(r'[^\d]', '', raw_d)
                        try: credit_amt = float(nums); debit_amt = None
                        except: pass

            date_ofx = _eco_date_ofx(date_info)
            name, memo = smart_label(label, memo_parts)
            if debit_amt and debit_amt > 0:
                txns.append(_make_txn(date_ofx, -debit_amt, name, memo))
            elif credit_amt and credit_amt > 0:
                txns.append(_make_txn(date_ofx, credit_amt, name, memo))

    # Fallback universel si rien extrait
    if not txns and _pdf_path and Path(_pdf_path).exists():
        _, txns2 = _universal_parse_path(_pdf_path, pages_text)
        if txns2:
            return info, txns2

    return info, [t for t in txns if t is not None]


# ════════════════════════════════════════════════════════════════════════════
# DEVISE & LABELS
# ════════════════════════════════════════════════════════════════════════════

BANK_CURRENCY = {
    'QONTO':'EUR','LCL':'EUR','CA':'EUR','CE':'EUR','BP':'EUR','CIC':'EUR',
    'CM':'EUR','CGD':'EUR','LBP':'EUR','SG':'EUR','BNP':'EUR','MYPOS':'EUR','SHINE':'EUR',
    'CBAO':'XOF','ECOBANK':'XOF','BCI':'XOF','CORIS':'XOF','UBA':'XOF',
    'ORABANK':'XOF','BOA':'XOF','ATB':'TND','SG_AFRIQUE':'XOF','BSIC':'XOF',
    'BIS':'XOF','BNDE':'XOF','UNIVERSAL':'XOF','NSIA':'XOF',
}

BANK_LABELS = {
    'QONTO':'Qonto','LCL':'LCL (Crédit Lyonnais)','CA':'Crédit Agricole',
    'CE':"Caisse d'Épargne",'BP':'Banque Populaire','CIC':'CIC',
    'CM':'Crédit Mutuel',
    'CGD':'Caixa Geral de Depositos','LBP':'La Banque Postale',
    'SG':'Société Générale','BNP':'BNP Paribas','MYPOS':'myPOS',
    'SHINE':'Shine (néo-banque pro)','CBAO':'CBAO (Sénégal)',
    'ECOBANK':'Ecobank','BCI':'BCI','CORIS':'Coris Bank','UBA':'UBA',
    'ORABANK':'Orabank','BOA':'Bank of Africa','ATB':'Arab Tunisian Bank',
    'SG_AFRIQUE':'Société Générale Afrique','BSIC':'BSIC (Sénégal)',
    'BIS':'Banque Islamique du Sénégal','BNDE':'BNDE','UNIVERSAL':'Format universel',
    'NSIA':'NSIA Banque',
}

AFRICAN_BANKS = {'CBAO','ECOBANK','BCI','CORIS','UBA','ORABANK','BOA','ATB',
                 'SG_AFRIQUE','UNIVERSAL','BSIC','BIS','BNDE','NSIA'}

PARSERS = {
    'QONTO':parse_qonto,'LCL':parse_lcl,'CA':parse_ca,'CE':parse_ce,
    'BP':parse_bp,'CIC':parse_cic,'CM':parse_cm,'CGD':parse_cgd,'LBP':parse_lbp,
    'SG':parse_sg,'BNP':parse_bnp,'MYPOS':parse_mypos,'SHINE':parse_shine,
    'CBAO':parse_cbao,'ECOBANK':parse_ecobank,'BCI':parse_bci,'CORIS':parse_coris,
    'UBA':parse_uba,'ORABANK':parse_orabank,'BOA':parse_boa,'ATB':parse_atb,
    'SG_AFRIQUE':parse_sg_afrique,'UNIVERSAL':parse_universal,
    'BSIC':parse_bsic,'BIS':parse_bis,'BNDE':parse_bnde,'NSIA':parse_nsia,
}


# ════════════════════════════════════════════════════════════════════════════
# GÉNÉRATION OFX
# ════════════════════════════════════════════════════════════════════════════

def period_to_ofx(date_str):
    try:
        p = date_str.split('/')
        return f"{p[2]}{p[1].zfill(2)}{p[0].zfill(2)}"
    except:
        return datetime.now().strftime('%Y%m%d')

def _clamp_balance_for_ofx(bal: float) -> tuple:
    """
    Quadra/Cegid limite le champ BALAMT à 11 caractères total
    (signe éventuel + chiffres + point + 2 décimales).
    - Avec décimales : max 9 999 999.99  (format "9999999.99"  = 10 chars)
    - Sans décimales : max 999 999 999   (format "999999999"   = 9 chars)
    En pratique Quadra bloque à partir de 10 chiffres entiers (ex: 1 000 000 000).
    On tronque à 9 999 999.99 pour laisser une marge.
    Retourne (valeur_clampée, True si tronquée).
    """
    QUADRA_MAX = 9_999_999_999.99   # ~10 milliards — limite sûre Quadra
    if abs(bal) > QUADRA_MAX:
        sign = -1 if bal < 0 else 1
        return sign * QUADRA_MAX, True
    return bal, False


def _format_balamt(bal: float, currency: str) -> str:
    """
    Formate le solde pour le champ BALAMT OFX.
    - XOF / devises sans centimes : entier sans décimales
    - EUR et autres : 2 décimales
    Quadra accepte mieux les entiers pour les devises africaines.
    """
    if currency in ('XOF', 'XAF', 'GNF', 'MGA'):
        return f"{int(round(bal))}"
    return f"{bal:.2f}"


def generate_ofx(info, txns, target='quadra', currency='EUR'):
    iban_full = info.get('iban', '') or ''
    bid, brid, aid = iban_to_rib(iban_full, info=info)

    # ── ACCTID pour Quadra / Money ────────────────────────────────────────────
    # Priorite :
    #   1. RIB extrait directement du PDF :
    #      - Banques africaines (IBAN >= 26 chars ou IBAN alphanum) : ACCTID =
    #        numéro de compte seul (ce que Money utilise pour identifier le client)
    #      - Banques FR (IBAN=27 chars, tout numérique) : ACCTID = Banque+Guichet+Compte+Cle
    #   2. IBAN court (<= 22 chars) -> utiliser l'IBAN compact directement.
    #   3. IBAN long (> 22 chars, UEMOA SN=28...) sans _rib_account -> BBAN compte
    #   4. Fallback fragment iban_to_rib().
    iban_is_real  = bool(re.match(r'^[A-Z]{2}\d{2}', iban_full.replace(' ','')))
    iban_compact  = re.sub(r'\s+', '', iban_full)
    # Detecter si c'est une banque africaine : IBAN > 25 chars OU BBAN contient des lettres
    bban_part = iban_compact[4:] if len(iban_compact) > 4 else ''
    is_african_iban = len(iban_compact) > 25 or bool(re.search(r'[A-Z]', bban_part))

    if info.get('_rib_account'):
        bank = info.get('_rib_bank', '') or ''
        if is_african_iban or not iban_is_real:
            # Banque africaine OU num compte brut (pas d'IBAN) :
            # Quadra reconnait le compte via le numero seul
            acctid = info['_rib_account']
        elif bank and bank != '00000':
            # Banque francaise avec RIB complet : Banque+Guichet+Compte+Cle
            acctid = (bank +
                      info.get('_rib_agency','') +
                      info.get('_rib_account','') +
                      info.get('_rib_key',''))
        else:
            acctid = info['_rib_account']
    elif iban_is_real and len(iban_compact) <= 22:
        # IBAN court (FR, BE, NL...) -> IBAN compact
        acctid = iban_compact
    elif iban_is_real and len(iban_compact) > 22:
        # IBAN long (UEMOA SN=28...) -> extraire le numero de compte du BBAN
        # Format BCEAO : CC(2)+KK(2)+Banque(5)+Agence(5)+Compte(11-12)+Cle(2)
        bban = iban_compact[4:]
        compte_bceao = bban[10:-2] if len(bban) >= 13 else bban[10:]
        acctid = compte_bceao if compte_bceao else iban_compact[:22]
    else:
        # Fallback fragment RIB
        acctid = aid

    # ── ACCTID : limiter a 22 caracteres (limite OFX SGML) ───────────────────
    acctid = acctid[:22] if acctid else '0000000000'

    # ── BANKID / BRANCHID : purger les caractères non-alphanumériques ─────────
    bid  = re.sub(r'[^A-Z0-9]', '', bid.upper())[:11]  if bid  else '00000'
    brid = re.sub(r'[^A-Z0-9]', '', brid.upper())[:11] if brid else '00000'

    ds  = period_to_ofx(info.get('period_start',''))
    de  = period_to_ofx(info.get('period_end',''))
    dn  = datetime.now().strftime('%Y%m%d%H')

    # ── Solde : limité à 13 chiffres pour compatibilité Quadra/Cegid ──────────
    bal_raw = info.get('balance_close', 0.0)
    bal, _bal_truncated = _clamp_balance_for_ofx(bal_raw)

    memo_carries_label = target in ('myunisoft','sage','ebp')
    lines = [
        'OFXHEADER:100','DATA:OFXSGML','VERSION:102','SECURITY:NONE',
        'ENCODING:USASCII','CHARSET:1252','COMPRESSION:NONE',
        'OLDFILEUID:NONE','NEWFILEUID:NONE',
        '<OFX>','<SIGNONMSGSRSV1>','<SONRS>','<STATUS>',
        '<CODE>0','<SEVERITY>INFO','</STATUS>',
        f'<DTSERVER>{dn}','<LANGUAGE>FRA',
        '</SONRS>','</SIGNONMSGSRSV1>',
        '<BANKMSGSRSV1>','<STMTTRNRS>','<TRNUID>00',
        '<STATUS>','<CODE>0','<SEVERITY>INFO','</STATUS>',
        '<STMTRS>',f'<CURDEF>{currency}','<BANKACCTFROM>',
        f'<BANKID>{bid}',f'<BRANCHID>{brid}',
        f'<ACCTID>{acctid}','<ACCTTYPE>CHECKING','</BANKACCTFROM>',
        '<BANKTRANLIST>',f'<DTSTART>{ds}',f'<DTEND>{de}',
    ]
    for t in txns:
        name = t['name']
        memo = t.get('memo', '') or ''
        if memo_carries_label:
            name_tag = name
            memo_tag = (name + ' | ' + memo) if memo else name
        else:
            name_tag = name
            memo_tag = memo
        lines += [
            '<STMTTRN>',
            f"<TRNTYPE>{t['type']}",
            f"<DTPOSTED>{t['date']}",
            f"<TRNAMT>{_format_balamt(t['amount'], currency)}",
            f"<FITID>{t['fitid']}",
            '<NAME>' + name_tag,
            '<MEMO>' + memo_tag,
            '</STMTTRN>',
        ]
    bal_fmt = _format_balamt(bal, currency)
    # TRNAMT : aussi entier pour devises sans centimes
    lines += [
        '</BANKTRANLIST>',
        f'<LEDGERBAL>',f'<BALAMT>{bal_fmt}',f'<DTASOF>{dn}','</LEDGERBAL>',
        f'<AVAILBAL>',f'<BALAMT>{bal_fmt}',f'<DTASOF>{dn}','</AVAILBAL>',
        '</STMTRS>','</STMTTRNRS>','</BANKMSGSRSV1>','</OFX>',
    ]
    return '\n'.join(lines) + '\n'


# ════════════════════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE DE CONVERSION (avec cache Streamlit)
# ════════════════════════════════════════════════════════════════════════════

def _pdf_to_images_base64(pdf_path, dpi=200):
    """Convertit chaque page du PDF en image base64 via PyMuPDF."""
    if not _FITZ_OK:
        return []
    images = []
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        images.append(base64.b64encode(img_bytes).decode())
    doc.close()
    return images


def _ocr_via_claude(pdf_path):
    """
    OCR de secours : envoie les pages du PDF (images) à l'API Anthropic
    et retourne une liste de textes (une entrée par page).
    Nécessite PyMuPDF (fitz). Utilise l'endpoint /v1/messages sans clé API
    (le proxy Streamlit Cloud l'injecte automatiquement via ANTHROPIC_API_KEY).
    """
    images_b64 = _pdf_to_images_base64(pdf_path)
    if not images_b64:
        return []

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        logger.warning("ANTHROPIC_API_KEY absent — OCR via Claude impossible")
        return []

    pages_text = []
    for img_b64 in images_b64:
        payload = {
            "model": "claude-sonnet-4-6",
            "max_tokens": 4096,
            "messages": [{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": (
                            "Tu es un expert en extraction de données de relevés bancaires. "
                            "Analyse cette image de relevé bancaire et extrais TOUTES les transactions.\n\n"
                            "INSTRUCTIONS STRICTES :\n"
                            "1. Transcris chaque ligne de transaction sur une seule ligne\n"
                            "2. Format obligatoire par ligne : DATE | LIBELLÉ | DÉBIT | CRÉDIT | SOLDE\n"
                            "   - DATE : format DD/MM/YYYY ou DD/MM/YY\n"
                            "   - Utilise le séparateur ' | ' entre chaque colonne\n"
                            "   - Si une colonne est vide, écris 'N/A'\n"
                            "   - Les montants : chiffres uniquement avec virgule décimale (ex: 1234,56)\n"
                            "3. Avant les transactions, extrais sur des lignes séparées :\n"
                            "   IBAN: [valeur]\n"
                            "   PERIODE: [DD/MM/YYYY] au [DD/MM/YYYY]\n"
                            "   SOLDE_OUVERTURE: [montant]\n"
                            "   SOLDE_CLOTURE: [montant]\n"
                            "4. Ensuite, commence les transactions avec la ligne '=== TRANSACTIONS ==='\n"
                            "5. Ignore les lignes de totaux, sous-totaux et en-têtes de colonnes\n"
                            "6. Réponds UNIQUEMENT avec ces données, sans commentaire."
                        )
                    }
                ]
            }]
        }
        try:
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                "https://api.anthropic.com/v1/messages",
                data=data,
                headers={
                    "Content-Type": "application/json",
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                },
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                result = json.loads(resp.read().decode())
                page_text = result["content"][0]["text"]
                pages_text.append(page_text)
        except Exception as exc:
            logger.warning("OCR Claude page échouée : %s", exc)
            pages_text.append("")
    return pages_text


def _parse_structured_ocr_text(pages_text):
    """
    Parse le texte structuré retourné par l'OCR Claude Vision (format pipe-delimited).
    Retourne (info_dict, txns_list) ou (None, []) si format non reconnu.
    """
    full = '\n'.join(pages_text)
    info = {'iban': '', 'period_start': '', 'period_end': '',
            'balance_open': 0.0, 'balance_close': 0.0}

    m_iban = re.search(r'^IBAN\s*:\s*(.+)$', full, re.MULTILINE | re.IGNORECASE)
    if m_iban:
        info['iban'] = re.sub(r'\s+', '', m_iban.group(1)).upper()

    m_per = re.search(r'^PERIODE\s*:\s*(\d{2}/\d{2}/\d{4})\s*au\s*(\d{2}/\d{2}/\d{4})',
                      full, re.MULTILINE | re.IGNORECASE)
    if m_per:
        info['period_start'] = m_per.group(1)
        info['period_end']   = m_per.group(2)

    m_so = re.search(r'^SOLDE_OUVERTURE\s*:\s*([\d\s,\.]+)', full, re.MULTILINE | re.IGNORECASE)
    if m_so:
        v = parse_amount(m_so.group(1).strip())
        if v: info['balance_open'] = v

    m_sc = re.search(r'^SOLDE_CLOTURE\s*:\s*([\d\s,\.]+)', full, re.MULTILINE | re.IGNORECASE)
    if m_sc:
        v = parse_amount(m_sc.group(1).strip())
        if v: info['balance_close'] = v

    txns = []
    in_section = False
    year = _year_from_text(full)
    SKIP_KW = {'TOTAL', 'SOLDE', 'SOUS-TOTAL', 'REPORT', 'DATE', 'DÉBIT',
               'DEBIT', 'CRÉDIT', 'CREDIT', 'LIBELLÉ', 'LIBELLE', 'MONTANT'}

    for line in full.split('\n'):
        line = line.strip()
        if '=== TRANSACTIONS ===' in line:
            in_section = True
            continue
        if not in_section:
            continue
        if not line or line.startswith('==='):
            continue

        parts = [p.strip() for p in line.split('|')]
        if len(parts) >= 3:
            date_raw = parts[0]
            label    = parts[1] if len(parts) > 1 else ''
            debit_s  = parts[2] if len(parts) > 2 else 'N/A'
            credit_s = parts[3] if len(parts) > 3 else 'N/A'

            if not label or any(kw in label.upper() for kw in SKIP_KW):
                continue

            date_ofx = date_full_to_ofx(date_raw.replace('.', '/').replace('-', '/'))
            if not re.match(r'^\d{8}$', date_ofx):
                m_dm = re.match(r'^(\d{1,2})/(\d{1,2})$', date_raw)
                if m_dm:
                    date_ofx = f"{year}{m_dm.group(2).zfill(2)}{m_dm.group(1).zfill(2)}"
                else:
                    continue

            debit_v  = None if debit_s  in ('N/A', '', '-', '—') else parse_amount(debit_s)
            credit_v = None if credit_s in ('N/A', '', '-', '—') else parse_amount(credit_s)

            name, memo = smart_label(label, [])
            if debit_v and debit_v > 0:
                txn = _make_txn(date_ofx, -debit_v, name, memo)
                if txn: txns.append(txn)
            elif credit_v and credit_v > 0:
                txn = _make_txn(date_ofx, credit_v, name, memo)
                if txn: txns.append(txn)

    if txns:
        return info, txns
    return None, []


@st.cache_data(show_spinner=False)
def process_pdf(file_bytes: bytes, filename: str, force_ocr: bool = False):
    """
    Traite un PDF (bytes) et retourne (bank, info, txns, error).
    Mis en cache par Streamlit — évite de ré-analyser si le fichier n'a pas changé.
    Si force_ocr=True, ignore la détection automatique et active l'OCR directement.
    """
    if not _PDFPLUMBER_OK:
        return None, {}, [], "pdfplumber n'est pas installé. Ajoutez-le à requirements.txt."

    # Écrire dans un fichier temporaire (pdfplumber a besoin d'un chemin)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        pages_words = extract_words_by_page(tmp_path)
        pages_text  = extract_text_by_page(tmp_path)
        _ocr_used = None

        if force_ocr or not _pdf_has_text(pages_text):
            # Essai 1 : Tesseract (si disponible localement)
            if _OCR_AVAILABLE:
                pages_text  = _ocr_pdf(tmp_path)
                pages_words = []
                _ocr_used = 'tesseract'
            # Essai 2 : OCR via API Claude Vision (PyMuPDF requis)
            elif _FITZ_OK and os.environ.get("ANTHROPIC_API_KEY"):
                logger.info("PDF scanné détecté — OCR via Claude Vision")
                pages_text  = _ocr_via_claude(tmp_path)
                pages_words = []
                _ocr_used = 'claude_vision'
                if not _pdf_has_text(pages_text, min_chars=20):
                    return None, {}, [], (
                        "OCR via Claude Vision n'a pas pu extraire de texte. "
                        "Vérifiez la qualité du scan."
                    )
                # Essayer le parseur structuré en premier (format pipe du nouveau prompt)
                info_struct, txns_struct = _parse_structured_ocr_text(pages_text)
                if txns_struct:
                    bank_struct = detect_bank(pages_text)
                    info_struct['_ocr_mode'] = 'claude_vision_structured'
                    return bank_struct, info_struct, txns_struct, None
            else:
                return None, {}, [], (
                    "⚠️ PDF scanné (image uniquement) — aucun texte extractible. "
                    "Solutions : (1) installez Tesseract + pytesseract + pdf2image, "
                    "ou (2) installez PyMuPDF (pip install pymupdf) et définissez "
                    "ANTHROPIC_API_KEY pour activer l'OCR via Claude Vision."
                )

        bank = detect_bank(pages_text)

        if bank in AFRICAN_BANKS:
            info, txns = PARSERS[bank](pages_words, pages_text, _pdf_path=tmp_path)
        else:
            info, txns = PARSERS[bank](pages_words, pages_text)

        # ── Fallback OCR automatique si la lecture standard retourne 0 transaction ──
        # Ne pas déclencher l'OCR si c'est un relevé explicitement vide (aucun mouvement)
        if not txns and not force_ocr and not _ocr_used and not info.get('_empty_period'):
            logger.info("Lecture standard : 0 transaction — tentative OCR automatique")
            if _OCR_AVAILABLE:
                pages_text_ocr  = _ocr_pdf(tmp_path)
                pages_words_ocr = []
                bank_ocr = detect_bank(pages_text_ocr)
                if bank_ocr in PARSERS:
                    if bank_ocr in AFRICAN_BANKS:
                        info2, txns2 = PARSERS[bank_ocr](pages_words_ocr, pages_text_ocr, _pdf_path=tmp_path)
                    else:
                        info2, txns2 = PARSERS[bank_ocr](pages_words_ocr, pages_text_ocr)
                    if txns2:
                        info, txns, bank, _ocr_used = info2, txns2, bank_ocr, 'tesseract_auto'
            elif _FITZ_OK and os.environ.get('ANTHROPIC_API_KEY'):
                pages_text_ocr = _ocr_via_claude(tmp_path)
                bank_ocr = detect_bank(pages_text_ocr)
                # Essai 1 : parser structuré (format pipe issu du nouveau prompt)
                info_struct, txns_struct = _parse_structured_ocr_text(pages_text_ocr)
                if txns_struct:
                    info2, txns2 = info_struct, txns_struct
                    bank_ocr = bank_ocr  # conserver la détection
                elif bank_ocr in PARSERS:
                    if bank_ocr in AFRICAN_BANKS:
                        info2, txns2 = PARSERS[bank_ocr]([], pages_text_ocr, _pdf_path=tmp_path)
                    else:
                        info2, txns2 = PARSERS[bank_ocr]([], pages_text_ocr)
                else:
                    info2, txns2 = {}, []
                if txns2:
                    info, txns, bank, _ocr_used = info2, txns2, bank_ocr, 'claude_vision_auto'

        # ── Stratégie supplémentaire : parseur universel si toujours 0 txn ──────
        if not txns and not _ocr_used and not info.get('_empty_period'):
            logger.info("Tous les parseurs ont échoué — tentative parseur universel direct")
            try:
                info_u, txns_u = _universal_parse_path(tmp_path, pages_text)
                if txns_u:
                    info, txns, bank = info_u, txns_u, bank
                    info['_fallback_universal'] = True
            except Exception as ex:
                logger.warning("Parseur universel direct échoué : %s", ex)

        if _ocr_used:
            info['_ocr_mode'] = _ocr_used
        return bank, info, txns, None

    except Exception as e:
        logger.error("Erreur traitement %s : %s", filename, e, exc_info=True)
        return None, {}, [], str(e)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


# ════════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ════════════════════════════════════════════════════════════════════════════

def fmt_amount(amount: float, currency: str) -> str:
    if currency == 'EUR':
        return f"{abs(amount):,.2f} €".replace(",", "\u202f")
    else:
        return f"{abs(amount):,.0f} {currency}".replace(",", "\u202f")


def main():
    st.set_page_config(
        page_title="OFX Bridge — PDF vers OFX",
        page_icon="💱",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── CSS Modern Premium ────────────────────────────────────────────────────
    st.markdown("""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

      /* ── Reset & base ── */
      html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
      .stApp { background: linear-gradient(135deg, #f0f4ff 0%, #f5f3ff 50%, #faf5ff 100%) !important; }
      .block-container { padding-top: 1.5rem !important; padding-bottom: 3rem !important; max-width: 1120px !important; }

      /* ── Textes globaux ── */
      h1, h2, h3, h4 { color: #1e1b4b !important; font-family: 'Inter', sans-serif !important; }
      p, label, .stMarkdown p { color: #4a5568 !important; }

      /* ── Sidebar ── */
      [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e1b4b 0%, #312e81 60%, #4c1d95 100%) !important;
        border-right: none !important;
      }
      [data-testid="stSidebar"] * { color: #e8e8ff !important; }
      [data-testid="stSidebar"] h2,
      [data-testid="stSidebar"] h3,
      [data-testid="stSidebar"] strong { color: #ffffff !important; }
      [data-testid="stSidebar"] .stSelectbox label { color: #c4b5fd !important; font-size: 0.82rem !important; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; }
      [data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div {
        background: rgba(255,255,255,0.1) !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        border-radius: 8px !important;
        color: #fff !important;
      }
      [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12) !important; }

      /* ── Metric cards ── */
      [data-testid="metric-container"] {
        background: #ffffff !important;
        border: 1px solid #e0e7ff !important;
        border-radius: 16px !important;
        padding: 18px 20px !important;
        box-shadow: 0 4px 16px rgba(99,102,241,0.08) !important;
        transition: transform 0.15s, box-shadow 0.15s;
      }
      [data-testid="metric-container"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(99,102,241,0.14) !important;
      }
      [data-testid="stMetricValue"] { color: #1e1b4b !important; font-size: 1.4rem !important; font-weight: 800 !important; }
      [data-testid="stMetricLabel"] { color: #6366f1 !important; font-size: 0.78rem !important; font-weight: 700 !important; text-transform: uppercase; letter-spacing: 0.06em; }

      /* ── Download button ── */
      .stDownloadButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #7c3aed 100%) !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        padding: 0.7rem 2rem !important;
        letter-spacing: 0.01em !important;
        box-shadow: 0 4px 16px rgba(99,102,241,0.35) !important;
        transition: all 0.2s ease !important;
      }
      .stDownloadButton > button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
      }
      .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #4f46e5 0%, #6d28d9 100%) !important;
        box-shadow: 0 8px 24px rgba(99,102,241,0.45) !important;
        transform: translateY(-2px) !important;
      }

      /* ── Primary buttons ── */
      .stButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #7c3aed 100%) !important;
        color: #fff !important; border: none !important;
        border-radius: 12px !important; font-weight: 700 !important;
        padding: 0.6rem 1.6rem !important;
        box-shadow: 0 4px 14px rgba(99,102,241,0.3) !important;
        transition: all 0.2s !important;
      }
      .stButton > button:hover {
        background: linear-gradient(135deg, #4f46e5 0%, #6d28d9 100%) !important;
        box-shadow: 0 8px 20px rgba(99,102,241,0.4) !important;
        transform: translateY(-2px) !important;
      }

      /* ── File uploader ── */
      [data-testid="stFileUploader"] {
        background: rgba(255,255,255,0.8) !important;
        border: 2px dashed #a5b4fc !important;
        border-radius: 20px !important;
        padding: 1.5rem !important;
        box-shadow: 0 4px 20px rgba(99,102,241,0.08) !important;
        backdrop-filter: blur(10px);
        transition: border-color 0.2s, box-shadow 0.2s;
      }
      [data-testid="stFileUploader"]:hover {
        border-color: #6366f1 !important;
        box-shadow: 0 8px 32px rgba(99,102,241,0.15) !important;
      }
      [data-testid="stFileUploader"] label { color: #1e1b4b !important; font-weight: 700 !important; }
      [data-testid="stFileDropzone"] { background: #f5f3ff !important; }

      /* ── Dataframe / data editor ── */
      [data-testid="stDataFrame"],
      [data-testid="stDataEditor"] {
        border: 1px solid #e0e7ff !important;
        border-radius: 14px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 14px rgba(99,102,241,0.07) !important;
      }

      /* ── Alert boxes ── */
      [data-testid="stAlert"] { border-radius: 14px !important; }

      /* ── Separators ── */
      hr { border-color: #e0e7ff !important; margin: 1.8rem 0 !important; }

      /* ── Custom components ── */
      .ofx-hero {
        background: linear-gradient(135deg, #1e1b4b 0%, #4338ca 45%, #7c3aed 80%, #9333ea 100%);
        border-radius: 24px;
        padding: 2.6rem 3rem;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 20px 60px rgba(99,102,241,0.3);
      }
      .ofx-hero::before {
        content: '';
        position: absolute; top: -60px; right: -60px;
        width: 280px; height: 280px;
        background: rgba(255,255,255,0.06);
        border-radius: 50%;
      }
      .ofx-hero::after {
        content: '';
        position: absolute; bottom: -80px; left: 60px;
        width: 180px; height: 180px;
        background: rgba(255,255,255,0.04);
        border-radius: 50%;
      }
      .ofx-hero h1 {
        color: #ffffff !important;
        font-size: 2.1rem !important;
        font-weight: 800 !important;
        margin: 0 0 0.5rem 0 !important;
        letter-spacing: -0.03em;
      }
      .ofx-hero p {
        color: rgba(255,255,255,0.8) !important;
        font-size: 1.05rem !important;
        margin: 0 !important;
      }
      .ofx-hero .badges {
        margin-top: 1.4rem;
        display: flex; gap: 10px; flex-wrap: wrap;
      }
      .ofx-hero .badge {
        background: rgba(255,255,255,0.14);
        border: 1px solid rgba(255,255,255,0.22);
        color: #fff !important;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 600;
        backdrop-filter: blur(4px);
      }

      .step-card {
        background: rgba(255,255,255,0.85);
        border: 1px solid #e0e7ff;
        border-radius: 16px;
        padding: 1.3rem 1.5rem;
        display: flex; align-items: flex-start; gap: 14px;
        box-shadow: 0 4px 16px rgba(99,102,241,0.07);
        backdrop-filter: blur(8px);
        transition: transform 0.15s, box-shadow 0.15s;
      }
      .step-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 30px rgba(99,102,241,0.14);
      }
      .step-num {
        background: linear-gradient(135deg, #6366f1, #7c3aed);
        color: #fff; font-weight: 800; font-size: 1rem;
        width: 34px; height: 34px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0;
        box-shadow: 0 4px 10px rgba(99,102,241,0.35);
      }
      .step-text strong { color: #1e1b4b !important; font-size: 0.95rem; }
      .step-text small { color: #6366f1 !important; font-size: 0.81rem; }

      .bank-badge {
        background: linear-gradient(135deg, #ede9fe, #ddd6fe);
        color: #5b21b6 !important;
        border: 1px solid #c4b5fd;
        padding: 5px 14px; border-radius: 20px;
        font-size: 0.84rem; font-weight: 700;
        display: inline-block;
        letter-spacing: 0.01em;
      }

      .result-card {
        background: rgba(255,255,255,0.92);
        border: 1px solid #e0e7ff;
        border-radius: 20px;
        padding: 2rem 2.2rem;
        margin-bottom: 1.8rem;
        box-shadow: 0 8px 32px rgba(99,102,241,0.09);
        backdrop-filter: blur(12px);
      }
      .result-card .file-title {
        font-size: 1rem; font-weight: 700; color: #1e1b4b !important;
        display: flex; align-items: center; gap: 8px; margin-bottom: 1rem;
      }

      .info-row {
        background: #f5f3ff;
        border: 1px solid #e0e7ff;
        border-radius: 12px;
        padding: 10px 16px;
        margin-bottom: 1rem;
        display: flex; gap: 2rem; flex-wrap: wrap;
      }
      .info-item { font-size: 0.85rem; color: #4a5568 !important; }
      .info-item strong { color: #1e1b4b !important; }

      .security-box {
        background: linear-gradient(135deg, #ede9fe, #f0fdf4);
        border: 1px solid #c4b5fd;
        border-radius: 14px;
        padding: 1rem 1.2rem;
        margin-top: 0.5rem;
      }
      .security-box p { color: #4c1d95 !important; font-size: 0.88rem !important; margin: 0 !important; }

      .sidebar-label {
        color: rgba(196,181,253,0.9) !important;
        font-size: 0.72rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.09em !important;
        text-transform: uppercase !important;
        margin-bottom: 0.3rem !important;
      }

      .footer-bar {
        background: rgba(255,255,255,0.85);
        border: 1px solid #e0e7ff;
        border-radius: 14px;
        padding: 1rem 1.5rem;
        display: flex; justify-content: space-between; align-items: center;
        flex-wrap: wrap; gap: 0.5rem;
        backdrop-filter: blur(8px);
      }
      .footer-bar span { color: #7c8db0 !important; font-size: 0.82rem !important; }
      .footer-bar .highlight { color: #6366f1 !important; font-weight: 700 !important; }

      .empty-state {
        background: rgba(255,255,255,0.7);
        border: 2px dashed #c4b5fd;
        border-radius: 24px;
        padding: 3.5rem 2rem;
        text-align: center;
        backdrop-filter: blur(8px);
      }
      .empty-state .icon { font-size: 3.5rem; margin-bottom: 1rem; }
      .empty-state h3 { color: #1e1b4b !important; font-size: 1.3rem !important; margin-bottom: 0.5rem !important; }
      .empty-state p { color: #7c8db0 !important; font-size: 0.92rem !important; }

      /* ── Data editor ── */
      [data-testid="stDataEditor"] [data-testid="glideDataEditor"] {
        border-radius: 14px !important;
      }
      .edit-notice {
        background: #fffbeb; border: 1px solid #fcd34d;
        border-radius: 12px; padding: 10px 16px;
        font-size: 0.85rem; color: #92400e;
        margin: 0.5rem 0;
      }

      /* Spinner override */
      .stSpinner > div { border-top-color: #6366f1 !important; }

      /* Section headers */
      .section-header {
        font-size: 1rem; font-weight: 700; color: #1e1b4b !important;
        margin: 0 0 0.8rem 0; padding-bottom: 0.5rem;
        border-bottom: 2px solid #e0e7ff;
        display: flex; align-items: center; gap: 8px;
      }

      /* ── Confidence badge ── */
      .conf-badge {
        display: inline-flex; align-items: center; gap: 6px;
        padding: 3px 12px; border-radius: 20px;
        font-size: 0.78rem; font-weight: 700;
        letter-spacing: 0.02em;
      }
      .conf-high   { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; }
      .conf-medium { background:#fef3c7; color:#92400e; border:1px solid #fcd34d; }
      .conf-low    { background:#fee2e2; color:#991b1b; border:1px solid #fca5a5; }

      /* ── OCR badge ── */
      .ocr-badge {
        display: inline-flex; align-items: center; gap: 5px;
        background: #ede9fe; color: #5b21b6;
        border: 1px solid #c4b5fd;
        padding: 3px 10px; border-radius: 20px;
        font-size: 0.76rem; font-weight: 700;
        margin-left: 8px;
      }

      /* ── Flux chart bar ── */
      .flux-bar-wrap { margin: 1rem 0 0.8rem; }
      .flux-bar-label { font-size: 0.78rem; font-weight: 700; color: #64748b !important;
                        text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 4px; }
      .flux-bar-track {
        height: 12px; border-radius: 6px;
        background: #e2e8f0; overflow: hidden; position: relative;
      }
      .flux-bar-fill-debit  { height: 100%; border-radius: 6px;
                               background: linear-gradient(90deg,#f87171,#ef4444); }
      .flux-bar-fill-credit { height: 100%; border-radius: 6px;
                               background: linear-gradient(90deg,#34d399,#10b981); }
    </style>
    """, unsafe_allow_html=True)

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="padding: 0.5rem 0 1rem 0;">
          <div style="font-size:1.6rem; font-weight:800; color:#fff; letter-spacing:-0.02em;">💱 OFX Bridge</div>
          <div style="font-size:0.78rem; color:rgba(255,255,255,0.5); margin-top:3px; font-weight:500;">v2.4 — Convertisseur PDF → OFX</div>
          <div style="margin-top:8px; display:flex; gap:6px; flex-wrap:wrap;">
            <span style="background:rgba(99,102,241,0.3);color:#c7d2fe;padding:2px 8px;border-radius:12px;font-size:0.7rem;font-weight:700;border:1px solid rgba(99,102,241,0.4)">✦ IA Vision</span>
            <span style="background:rgba(16,185,129,0.2);color:#6ee7b7;padding:2px 8px;border-radius:12px;font-size:0.7rem;font-weight:700;border:1px solid rgba(16,185,129,0.3)">✓ 24 banques</span>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        st.markdown('<div class="sidebar-label">⚙️ Logiciel comptable cible</div>', unsafe_allow_html=True)
        target = st.selectbox(
            "Logiciel cible",
            options=["Quadra / Cegid", "MyUnisoft", "Sage", "EBP"],
            index=0,
            label_visibility="collapsed",
            help="Affecte la disposition NAME/MEMO dans le fichier OFX."
        )
        target_map = {"Quadra / Cegid": "quadra", "MyUnisoft": "myunisoft",
                      "Sage": "sage", "EBP": "ebp"}
        target_code = target_map[target]

        st.divider()

        st.markdown('<div class="sidebar-label">🔍 Options OCR</div>', unsafe_allow_html=True)
        force_ocr = st.checkbox(
            "Forcer l'OCR (relevé scanné)",
            value=False,
            help=(
                "Cochez cette option si votre relevé est un PDF scanné (image) "
                "et que la détection automatique ne trouve aucune transaction. "
                "Nécessite Tesseract installé sur le serveur."
            )
        )
        if force_ocr:
            st.markdown(
                "<div style='font-size:0.76rem; color:#fcd34d; margin-top:-0.3rem; margin-bottom:0.4rem;'>"
                "⚡ OCR activé — le traitement sera plus lent.</div>",
                unsafe_allow_html=True
            )

        st.divider()

        st.markdown("""
        <div class="security-box" style="background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);">
          <div style="font-size:0.95rem; font-weight:700; color:#fff; margin-bottom:5px;">🔒 100 % local & sécurisé</div>
          <div style="font-size:0.82rem; color:rgba(255,255,255,0.7); line-height:1.5;">
            Vos relevés sont traités <strong style="color:#7dd3fc">directement sur ce serveur</strong>.
            Aucun fichier n'est envoyé vers le cloud. Aucune IA distante. Aucune inscription requise.
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        st.markdown("""
        <div style="font-size:0.78rem; color:rgba(255,255,255,0.45); line-height:1.7;">
          <div style="font-weight:700; color:rgba(255,255,255,0.6); margin-bottom:5px; font-size:0.72rem; text-transform:uppercase; letter-spacing:0.08em;">Détection automatique</div>
          Qonto · LCL · Crédit Agricole · Caisse d'Épargne · Banque Populaire · CIC · La Banque Postale · Société Générale · BNP Paribas · CGD · myPOS · Shine · CBAO · Ecobank · BCI · Coris · UBA · Orabank · BOA · ATB · BSIC · BIS · BNDE · + Format universel
        </div>
        """, unsafe_allow_html=True)

    # ── HERO HEADER ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="ofx-hero">
      <h1>💱 OFX Bridge <span style="font-size:1rem;font-weight:500;opacity:0.7;letter-spacing:0;margin-left:8px">PDF → OFX</span></h1>
      <p>Importez vos relevés bancaires directement dans Quadra, MyUnisoft, Sage ou EBP — en quelques secondes.</p>
      <div class="badges">
        <span class="badge">⚡ Détection automatique de la banque</span>
        <span class="badge">🔒 100 % local &amp; sécurisé</span>
        <span class="badge">📂 Multi-fichiers</span>
        <span class="badge">🤖 OCR IA pour scans</span>
        <span class="badge">✏️ Édition interactive</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Vérification dépendances
    if not _PDFPLUMBER_OK:
        st.error("❌ **pdfplumber n'est pas installé.** Ajoutez `pdfplumber` à votre `requirements.txt` et relancez l'application.")
        st.stop()

    # ── Comment ça marche ─────────────────────────────────────────────────────
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        st.markdown("""<div class="step-card">
          <div class="step-num">1</div>
          <div class="step-text"><strong>Déposez votre PDF</strong><br><small>Relevé natif ou scanné — détection automatique de la banque</small></div>
        </div>""", unsafe_allow_html=True)
    with col_s2:
        st.markdown("""<div class="step-card">
          <div class="step-num">2</div>
          <div class="step-text"><strong>Vérifiez &amp; corrigez</strong><br><small>Tableau éditable : date, libellé, montant, sens de l'opération</small></div>
        </div>""", unsafe_allow_html=True)
    with col_s3:
        st.markdown("""<div class="step-card">
          <div class="step-num">3</div>
          <div class="step-text"><strong>Téléchargez l'OFX</strong><br><small>Prêt pour Quadra · MyUnisoft · Sage · EBP</small></div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📂 Sélection des fichiers PDF</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Glissez-déposez un ou plusieurs relevés bancaires au format PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="PDF natif (texte extractible) recommandé. Les PDF scannés nécessitent Tesseract OCR."
    )

    if not uploaded_files:
        st.markdown("""
        <div class="empty-state">
          <div class="icon">📄</div>
          <h3>Aucun fichier sélectionné</h3>
          <p>Glissez-déposez vos relevés PDF ci-dessus pour démarrer la conversion.<br>
          La banque est détectée automatiquement — aucune configuration requise.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Transactions", "—")
        with c2: st.metric("Total Débits", "—")
        with c3: st.metric("Total Crédits", "—")
        st.stop()

    # ── Traitement de chaque fichier ──────────────────────────────────────────
    import pandas as pd

    for uploaded_file in uploaded_files:
        st.markdown(f"<div style='height:0.5rem'></div>", unsafe_allow_html=True)

        file_bytes = uploaded_file.read()

        with st.spinner(f"Analyse de « {uploaded_file.name} »…"):
            bank, info, txns, error = process_pdf(file_bytes, uploaded_file.name, force_ocr=force_ocr)

        # ── Carte de résultat ──────────────────────────────────────────────────
        st.markdown('<div class="result-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="file-title">📄 {uploaded_file.name}</div>', unsafe_allow_html=True)

        if error:
            st.error(f"❌ **Erreur :** {error}")
            st.markdown('</div>', unsafe_allow_html=True)
            continue

        if not txns:
            # ── Relevé vide légitimement (aucun mouvement sur la période) ──────
            if info.get('_empty_period') or (info.get('iban') and info.get('period_start')):
                # On a un IBAN et une période : on peut générer un OFX vide valide
                st.info(
                    f"ℹ️ **Relevé sans mouvement** — Ce relevé couvre la période "
                    f"{info.get('period_start','')} → {info.get('period_end','')} "
                    f"et ne contient aucune transaction. Un OFX vide (solde uniquement) "
                    f"sera généré — cela est normal et importable dans votre logiciel comptable."
                )
                txns = []  # OFX vide mais valide
            # Distinguer PDF scanné sans OCR vs. relevé non reconnu
            elif not _OCR_AVAILABLE:
                st.warning(
                    "⚠️ **PDF scanné détecté** — Ce relevé est une image (scan). "
                    "L'OCR (Tesseract) n'est pas installé sur ce serveur. "
                    "Pour traiter ce fichier, installez `pytesseract`, `pdf2image` et `tesseract-ocr` "
                    "puis redémarrez l'application."
                )
                st.markdown('</div>', unsafe_allow_html=True)
                continue
            else:
                st.warning("⚠️ Aucune transaction détectée. Vérifiez qu'il s'agit bien d'un relevé bancaire.")
                st.markdown('</div>', unsafe_allow_html=True)
                continue

        # Badge banque + infos
        bank_label = BANK_LABELS.get(bank, bank)
        currency   = BANK_CURRENCY.get(bank, 'EUR')
        file_key   = uploaded_file.name.replace('.', '_').replace(' ', '_')

        period_str = ""
        if info.get('period_start') and info.get('period_end'):
            period_str = f"{info['period_start']} → {info['period_end']}"
        iban_detected = info.get('iban', '') or ''

        # Indicateur de confiance basé sur les données extraites
        has_iban    = bool(iban_detected)
        has_period  = bool(info.get('period_start'))
        has_balance = bool(info.get('balance_close'))
        is_fallback = bool(info.get('_fallback_universal'))
        conf_score  = sum([has_iban, has_period, has_balance, not is_fallback, len(txns) > 5])
        if conf_score >= 4:
            conf_html = '<span class="conf-badge conf-high">✓ Extraction haute confiance</span>'
        elif conf_score >= 2:
            conf_html = '<span class="conf-badge conf-medium">⚠ Confiance moyenne — vérifiez</span>'
        else:
            conf_html = '<span class="conf-badge conf-low">⚡ Confiance faible — corrections conseillées</span>'

        ocr_mode = info.get('_ocr_mode', '')
        ocr_badge = ""
        if 'claude_vision' in ocr_mode:
            ocr_badge = '<span class="ocr-badge">🤖 OCR Claude Vision</span>'
        elif 'tesseract' in ocr_mode:
            ocr_badge = '<span class="ocr-badge" style="background:#dcfce7;color:#166534;border-color:#86efac">🔍 OCR Tesseract</span>'

        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:1rem; flex-wrap:wrap;">
          <span class="bank-badge">🏦 {bank_label}</span>
          {conf_html}
          {ocr_badge}
          <span style="font-size:0.83rem; color:#64748b">
            {'📅 ' + period_str if period_str else ''}
          </span>
        </div>
        """, unsafe_allow_html=True)

        # ── Champ IBAN/RIB éditable ────────────────────────────────────────────
        iban_key = f"iban_{file_key}"
        if iban_key not in st.session_state:
            st.session_state[iban_key] = iban_detected

        col_iban1, col_iban2 = st.columns([3, 1])
        with col_iban1:
            iban_label = "🔑 RIB (banque guichet compte clé)" if info.get('_rib_bank') else "🔑 IBAN"
            iban_help  = ("RIB extrait du relevé — format : BBBBB GGGGG CCCCCCCCCCC KK\n"
                          "Modifiez si la détection est incorrecte."
                          if info.get('_rib_bank') else
                          "IBAN extrait du relevé. Modifiez si la détection est incorrecte.")
            iban_value = st.text_input(
                iban_label,
                value=st.session_state[iban_key],
                key=f"iban_input_{file_key}",
                help=iban_help,
                placeholder="Ex: 30003 03320 00020641644 69 (RIB) ou FR76 ... (IBAN)",
            )
        with col_iban2:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            if st.button("↺ Réinitialiser", key=f"iban_reset_{file_key}",
                         help="Revenir à la valeur détectée automatiquement"):
                st.session_state[iban_key] = iban_detected
                st.rerun()

        # Mettre à jour info avec l'IBAN/RIB saisi manuellement
        if iban_value != iban_detected:
            iban_clean = iban_value.strip()
            # RIB français : 5+5+11+2 (tout numérique)
            rib_fr  = re.match(r'^(\d{5})\s+(\d{5})\s+(\d{10,12})\s+(\d{2})$', iban_clean)
            # RIB africain : code banque alphanumérique + agence + compte + clé
            # Ex: "SN213 01001 02341624101 33"
            rib_afr = re.match(r'^([A-Z]{0,2}\d{3,5})\s+(\d{4,6})\s+(\d{8,14})\s+(\d{2})$', iban_clean)
            if rib_fr:
                info['_rib_bank']    = rib_fr.group(1)
                info['_rib_agency']  = rib_fr.group(2)
                info['_rib_account'] = rib_fr.group(3)
                info['_rib_key']     = rib_fr.group(4)
                info['iban']         = iban_clean
            elif rib_afr:
                info['_rib_bank']    = rib_afr.group(1)
                info['_rib_agency']  = rib_afr.group(2)
                info['_rib_account'] = rib_afr.group(3)
                info['_rib_key']     = rib_afr.group(4)
                info['iban']         = iban_clean
            else:
                info['iban'] = iban_clean
                # Nettoyer les _rib_* pour forcer iban_to_rib à re-dériver depuis l'IBAN
                info.pop('_rib_bank', None)
                info.pop('_rib_agency', None)
                info.pop('_rib_account', None)
                info.pop('_rib_key', None)
        elif iban_detected:
            info['iban'] = iban_detected

        iban_display = info.get('iban', '') or '—'
        # ── Diagnostic IBAN/RIB ───────────────────────────────────────────────
        rib_bank    = info.get('_rib_bank', '')
        rib_agency  = info.get('_rib_agency', '')
        rib_account = info.get('_rib_account', '')
        rib_key     = info.get('_rib_key', '')

        if rib_bank and rib_account:
            # Afficher décomposition RIB
            st.markdown(
                f"<div style='font-size:0.78rem; color:#7489b0; margin-top:-0.4rem; margin-bottom:0.4rem;'>"
                f"🔑 RIB détecté&nbsp;: "
                f"<code style='color:#1a56db'>Banque&nbsp;<b>{rib_bank}</b></code> "
                f"<code style='color:#1a56db'>Guichet&nbsp;<b>{rib_agency}</b></code> "
                f"<code style='color:#1a56db'>Compte&nbsp;<b>{rib_account}</b></code> "
                f"<code style='color:#1a56db'>Clé&nbsp;<b>{rib_key}</b></code>"
                f"</div>",
                unsafe_allow_html=True
            )
        elif iban_display != '—':
            st.markdown(
                f"<div style='font-size:0.78rem; color:#7489b0; margin-top:-0.4rem; margin-bottom:0.4rem;'>"
                f"🔑 IBAN utilisé dans l'OFX&nbsp;: <code style='color:#1a56db'>{iban_display}</code></div>",
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                "<div style='font-size:0.78rem; color:#e53e3e; margin-top:-0.4rem; margin-bottom:0.4rem;'>"
                "⚠️ <b>Aucun IBAN/RIB détecté</b> — saisissez-le manuellement ci-dessus pour un OFX valide.</div>",
                unsafe_allow_html=True
            )

        # ── Avertissement solde trop grand (Quadra) ───────────────────────────
        bal_raw_check = info.get('balance_close', 0.0)
        _, bal_will_truncate = _clamp_balance_for_ofx(bal_raw_check)
        if bal_will_truncate:
            st.warning(
                f"⚠️ **Solde trop grand pour Quadra/Cegid** — La valeur `{bal_raw_check:,.2f}` "
                f"dépasse la limite du champ BALAMT (~10 milliards). "
                f"La valeur sera plafonnée à `9 999 999 999.99` dans l'OFX. "
                f"Les **transactions individuelles** ne sont pas affectées — l'import reste fonctionnel."
            )

        # ── Métriques ──────────────────────────────────────────────────────────
        total_debit  = sum(abs(t['amount']) for t in txns if t['type'] == 'DEBIT')
        total_credit = sum(t['amount']      for t in txns if t['type'] == 'CREDIT')
        balance      = total_credit - total_debit

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Transactions", f"{len(txns)}")
        with col2:
            st.metric("Total Débits", fmt_amount(total_debit, currency))
        with col3:
            st.metric("Total Crédits", fmt_amount(total_credit, currency))
        with col4:
            delta_label = "positif" if balance >= 0 else "négatif"
            st.metric("Solde net", fmt_amount(abs(balance), currency),
                      delta=f"{'+'  if balance >= 0 else '-'}{fmt_amount(abs(balance), currency)}",
                      delta_color="normal" if balance >= 0 else "inverse")

        # ── Barre de flux visuelle ─────────────────────────────────────────────
        flux_total = total_debit + total_credit
        if flux_total > 0:
            pct_debit  = int(100 * total_debit  / flux_total)
            pct_credit = int(100 * total_credit / flux_total)
            st.markdown(f"""
            <div class="flux-bar-wrap">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                <span class="flux-bar-label">🔴 Débits {pct_debit}%</span>
                <span style="font-size:0.72rem;color:#94a3b8;font-weight:600">Flux des opérations</span>
                <span class="flux-bar-label" style="text-align:right">Crédits {pct_credit}% 🟢</span>
              </div>
              <div style="height:14px;border-radius:7px;background:#e2e8f0;overflow:hidden;display:flex;">
                <div style="width:{pct_debit}%;background:linear-gradient(90deg,#f87171,#ef4444);transition:width 0.4s;"></div>
                <div style="width:{pct_credit}%;background:linear-gradient(90deg,#34d399,#10b981);transition:width 0.4s;"></div>
              </div>
            </div>""", unsafe_allow_html=True)

        # ── Tableau ÉDITABLE ───────────────────────────────────────────────────
        st.markdown(f"""
        <div class="section-header" style="margin-top:1.2rem">
          ✏️ Vérification &amp; correction des transactions
          <span style="font-size:0.78rem; font-weight:500; color:#7489b0; margin-left:10px;">
            Cliquez sur une cellule pour modifier • Supprimer une ligne avec la case à gauche
          </span>
        </div>""", unsafe_allow_html=True)

        # Construction du DataFrame éditable (montants numériques)
        edit_rows = []
        for t in txns:
            d = t['date']
            date_fmt = f"{d[6:8]}/{d[4:6]}/{d[0:4]}"
            is_debit = t['type'] == 'DEBIT'
            edit_rows.append({
                "📅 Date":    date_fmt,
                "Type":       "Débit" if is_debit else "Crédit",
                "Libellé":    t['name'],
                "Mémo":       t.get('memo', '') or '',
                "Montant":    round(abs(t['amount']), 2),
            })

        df_edit = pd.DataFrame(edit_rows)

        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            height=min(480, 60 + 38 * len(df_edit)),
            hide_index=False,
            num_rows="dynamic",
            key=f"editor_{file_key}",
            column_config={
                "📅 Date":  st.column_config.TextColumn(
                    "📅 Date", width=105,
                    help="Format JJ/MM/AAAA — modifiable directement",
                ),
                "Type": st.column_config.SelectboxColumn(
                    "Type", width=100,
                    options=["Débit", "Crédit"],
                    required=True,
                    help="Inverser Débit ↔ Crédit si mal détecté",
                ),
                "Libellé": st.column_config.TextColumn(
                    "Libellé", width=260,
                    help="Nom de la transaction exporté dans NAME",
                ),
                "Mémo": st.column_config.TextColumn(
                    "Mémo", width=200,
                    help="Champ MEMO dans l'OFX",
                ),
                "Montant": st.column_config.NumberColumn(
                    f"Montant ({currency})", width=130,
                    min_value=0.0, step=0.01, format="%.2f",
                    help="Valeur absolue — le signe est géré par le Type",
                ),
            },
        )

        # Indicateur de modifications
        orig_hash = str(df_edit.values.tolist())
        edit_hash = str(edited_df.values.tolist())
        has_changes = (orig_hash != edit_hash)
        nb_edited   = len(edited_df)

        if has_changes:
            st.markdown("""
            <div style="background:#fffbeb; border:1px solid #fcd34d; border-radius:10px;
                        padding:10px 16px; margin:0.5rem 0; font-size:0.85rem; color:#92400e;">
              ✏️ <strong>Modifications détectées</strong> — L'OFX sera généré depuis le tableau corrigé ci-dessus.
            </div>""", unsafe_allow_html=True)

        # ── Reconstruction des transactions depuis le tableau édité ────────────
        txns_final = []
        errors_edit = []
        for idx, row in edited_df.iterrows():
            try:
                raw_date  = str(row.get("📅 Date", "")).strip()
                txn_type  = str(row.get("Type", "Débit")).strip()
                label     = str(row.get("Libellé", "")).strip()
                memo      = str(row.get("Mémo", "")).strip()
                montant   = float(row.get("Montant", 0) or 0)

                if not raw_date or not label or montant == 0:
                    continue

                # Conversion date JJ/MM/AAAA → AAAAMMJJ
                date_ofx = date_full_to_ofx(raw_date)
                if not re.match(r'^\d{8}$', date_ofx):
                    errors_edit.append(f"Ligne {idx+1} : date invalide « {raw_date} »")
                    continue

                amount = -montant if txn_type == "Débit" else montant
                txn = _make_txn(date_ofx, amount, label[:64], memo[:128])
                if txn:
                    txns_final.append(txn)
            except Exception as ex:
                errors_edit.append(f"Ligne {idx+1} : {ex}")

        if errors_edit:
            for err in errors_edit:
                st.warning(f"⚠️ {err}")

        # Métriques recalculées sur les données éditées
        if has_changes and txns_final:
            total_debit_e  = sum(abs(t['amount']) for t in txns_final if t['type'] == 'DEBIT')
            total_credit_e = sum(t['amount']      for t in txns_final if t['type'] == 'CREDIT')
            balance_e      = total_credit_e - total_debit_e
            col1e, col2e, col3e, col4e = st.columns(4)
            with col1e: st.metric("Transactions (éditées)", f"{len(txns_final)}")
            with col2e: st.metric("Total Débits",  fmt_amount(total_debit_e,  currency))
            with col3e: st.metric("Total Crédits", fmt_amount(total_credit_e, currency))
            with col4e: st.metric("Solde net", fmt_amount(abs(balance_e), currency),
                                  delta=f"{'+'if balance_e>=0 else '-'}{fmt_amount(abs(balance_e),currency)}",
                                  delta_color="normal" if balance_e >= 0 else "inverse")

        # OFX généré depuis les données (potentiellement éditées)
        txns_for_ofx = txns_final if txns_final else txns

        # ── Téléchargement OFX ─────────────────────────────────────────────────
        st.markdown('<div class="section-header" style="margin-top:1.2rem">⬇️ Télécharger le fichier OFX</div>', unsafe_allow_html=True)

        ofx_content = generate_ofx(info, txns_for_ofx, target=target_code, currency=currency)
        ofx_bytes   = ofx_content.encode('latin-1', errors='replace')
        ofx_name    = Path(uploaded_file.name).stem + ".ofx"

        col_dl1, col_dl2 = st.columns([2, 3])
        with col_dl1:
            st.download_button(
                label=f"⬇️  Télécharger {ofx_name}",
                data=ofx_bytes,
                file_name=ofx_name,
                mime="application/x-ofx",
                key=f"dl_{file_key}",
                use_container_width=True,
            )
        with col_dl2:
            st.markdown(f"""
            <div style="padding:0.6rem 0; font-size:0.83rem; color:#7489b0; line-height:1.7;">
              <b style="color:#0f2d6b">Format :</b> OFX Standard ({target})
              &nbsp;&nbsp;·&nbsp;&nbsp;
              <b style="color:#0f2d6b">Devise :</b> {currency}
              &nbsp;&nbsp;·&nbsp;&nbsp;
              <b style="color:#0f2d6b">{len(txns_for_ofx)} transactions</b>
              {'&nbsp;&nbsp;·&nbsp;&nbsp;<span style="color:#d97706;font-weight:600">✏️ Données corrigées</span>' if has_changes else ''}
              &nbsp;&nbsp;·&nbsp;&nbsp;
              Encodage Latin-1 (Cegid / MyUnisoft)
            </div>
            """, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)  # .result-card

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div class="footer-bar">
      <span>💱 <span class="highlight">OFX Bridge</span> v2.4</span>
      <span>🔒 Traitement 100 % local — Aucune donnée envoyée vers un serveur externe</span>
      <span>✉️ Compatible Quadra · MyUnisoft · Sage · EBP</span>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()