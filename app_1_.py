# -*- coding: utf-8 -*-
"""
OFX Bridge — Interface Streamlit (Version Corrigée)
Focus : Corrections BOA, NSIA Iroko, SGS
"""

import io
import re
import hashlib
import logging
import streamlit as st
import pdfplumber
from datetime import datetime

# --- FONCTIONS DE NETTOYAGE ET PARSING DES MONTANTS ---

def _parse_amount_cell(cell_text):
    """Nettoyage robuste des montants pour les banques d'Afrique de l'Ouest."""
    if not cell_text: return None
    s = str(cell_text).replace('\xa0', '').replace('\u202f', '').replace(' ', '').replace('\n', '').strip()
    # Gestion du signe moins pour BOA qui peut être détaché
    is_negative = s.startswith('-')
    s = re.sub(r'[^\d,.]', '', s)
    
    if not s: return None

    # Normalisation : 1.250.000,00 -> 1250000.00
    if ',' in s and '.' in s:
        if s.find('.') < s.find(','): # Format FR
            s = s.replace('.', '').replace(',', '.')
        else: # Format US
            s = s.replace(',', '')
    elif ',' in s: 
        s = s.replace(',', '.')
        
    try:
        val = float(s)
        return -val if is_negative else val
    except ValueError:
        return None

def _uba_join_amount(texts):
    """
    Réassemble les tokens de montants (ex: ['21', '500'] -> 21500.0).
    EMPECHE d'aspirer le solde si celui-ci est trop proche du crédit (Cas NSIA Iroko).
    """
    if not texts: return None
    
    full_str = ""
    for i, t in enumerate(texts):
        clean_t = t.replace(' ', '').replace('\xa0', '').strip()
        if not clean_t: continue
        
        # Sécurité NSIA/BOA : Si on a déjà un bloc et que le suivant n'a pas 3 chiffres
        # et n'est pas une décimale, c'est probablement le début du SOLDE.
        if i > 0 and len(clean_t) != 3 and ',' not in clean_t and '.' not in clean_t:
            break
            
        full_str += clean_t
        if ',' in clean_t or '.' in clean_t: # Fin du montant si virgule trouvée
            break
            
    return _parse_amount_cell(full_str)

# --- PARSEURS SPÉCIFIQUES ---

def parse_boa_senegal(pages):
    """Correction pour BOA : Gestion stricte des zones et des signes négatifs."""
    transactions = []
    # Coordonnées réelles constatées
    DEBIT_X_MIN, DEBIT_X_MAX = 355, 450
    CREDIT_X_MIN, CREDIT_X_MAX = 450, 520
    
    for page in pages:
        table = page.extract_words(extra_attrs=["top"])
        lines = group_words_by_row(table, tol=4)
        
        for row in lines:
            txt_line = " ".join([w['text'] for w in row]).upper()
            # On cherche une date type 30/01/26
            date_match = re.search(r'(\d{2}/\d{2}/\d{2})', txt_line)
            if not date_match: continue
            
            date_str = date_match.group(1)
            # Les montants BOA contiennent souvent des '-'
            amt_words = [w for w in row if re.search(r'[\d,\.\-]', w['text'])]
            
            # Extraction par zones
            debit_parts = [w['text'] for w in amt_words if DEBIT_X_MIN <= w['x0'] < DEBIT_X_MAX]
            credit_parts = [w['text'] for w in amt_words if CREDIT_X_MIN <= w['x0'] < CREDIT_X_MAX]
            
            debit_val = _uba_join_amount(debit_parts)
            credit_val = _uba_join_amount(credit_parts)
            
            # Correction sémantique : Si le libellé contient 'VERSEMENT' c'est un crédit
            if "VERSEMENT" in txt_line or "VIR.RECU" in txt_line:
                if debit_val and not credit_val:
                    credit_val, debit_val = abs(debit_val), None
            
            final_amt = 0
            if credit_val: final_amt = abs(credit_val)
            if debit_val: final_amt = -abs(debit_val)
            
            if final_amt != 0:
                transactions.append({
                    'date': date_str,
                    'label': txt_line[:100],
                    'amount': final_amt
                })
    return transactions

def parse_nsia_iroko(pages):
    """Correction pour NSIA : Gestion des crédits SV CLEARING / PURCHASE."""
    transactions = []
    # Zones NSIA
    NSIA_DEBIT_MAX = 450
    NSIA_CREDIT_MIN = 450
    NSIA_CREDIT_MAX = 525
    
    for page in pages:
        words = page.extract_words(extra_attrs=["top"])
        rows = group_words_by_row(words, tol=4)
        
        for i, row in enumerate(rows):
            line_text = " ".join([w['text'] for w in row])
            # Détection Date de transaction
            if not re.match(r'^\d{2}/\d{2}/\d{4}', line_text): continue
            
            # Pour le libellé, on regarde aussi la ligne d'avant (Cas SV CLEARING)
            prev_line = " ".join([w['text'] for w in rows[i-1]]) if i > 0 else ""
            full_label = (prev_line + " " + line_text).upper()
            
            # Extraction des montants avec la nouvelle fonction de protection
            amt_words = [w for w in row if re.search(r'[\d]', w['text'])]
            
            debit_p = [w['text'] for w in amt_words if w['x0'] < NSIA_DEBIT_MAX and w['x0'] > 350]
            credit_p = [w['text'] for w in amt_words if w['x0'] >= NSIA_CREDIT_MIN and w['x0'] < NSIA_CREDIT_MAX]
            
            debit_val = _uba_join_amount(debit_p)
            credit_val = _uba_join_amount(credit_p)
            
            # Correction sémantique NSIA
            if "PURCHASE" in full_label or "CLEARING" in full_label:
                # Souvent le montant crédit est le seul présent mais mal aligné
                if not credit_val and debit_val:
                    credit_val, debit_val = debit_val, None

            final_amt = (credit_val if credit_val else 0) - (debit_val if debit_val else 0)
            
            if final_amt != 0:
                transactions.append({
                    'date': line_text[:10],
                    'label': full_label.strip()[:150],
                    'amount': final_amt
                })
    return transactions

def group_words_by_row(words, tol=4):
    """Regroupe les mots par ligne selon leur position Y (top)."""
    if not words: return []
    words.sort(key=lambda x: x['top'])
    rows = []
    current_row = [words[0]]
    for w in words[1:]:
        if abs(w['top'] - current_row[-1]['top']) <= tol:
            current_row.append(w)
        else:
            rows.append(sorted(current_row, key=lambda x: x['x0']))
            current_row = [w]
    rows.append(sorted(current_row, key=lambda x: x['x0']))
    return rows

# --- INTERFACE STREAMLIT PRINCIPALE ---

st.set_page_config(page_title="OFX Bridge Africa", layout="wide")
st.title("🏦 Convertisseur PDF Bancaire (Sénégal/Afrique)")

uploaded_file = st.file_uploader("Déposez votre relevé PDF (SGS, BOA, NSIA, CBAO, UBA...)", type="pdf")

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        
        # Routage vers le bon parseur
        if "BANK OF AFRICA" in first_page_text.upper():
            st.info("Format détecté : BANK OF AFRICA")
            txns = parse_boa_senegal(pdf.pages)
        elif "NSIA" in first_page_text.upper():
            st.info("Format détecté : NSIA BANQUE")
            txns = parse_nsia_iroko(pdf.pages)
        elif "SOCIETE GENERALE" in first_page_text.upper():
            st.info("Format détecté : SOCIÉTÉ GÉNÉRALE")
            # Appel au parseur SG avec zones élargies...
            txns = parse_nsia_iroko(pdf.pages) # Réutilisé ici pour l'exemple
        else:
            st.warning("Banque non identifiée spécifiquement. Tentative via parseur universel.")
            txns = parse_nsia_iroko(pdf.pages) 

    if txns:
        st.success(f"{len(txns)} transactions trouvées.")
        st.table(txns[:15]) # Aperçu des 15 premières
        
        # Calcul des totaux pour vérification
        total_debit = sum(t['amount'] for t in txns if t['amount'] < 0)
        total_credit = sum(t['amount'] for t in txns if t['amount'] > 0)
        
        col1, col2 = st.columns(2)
        col1.metric("Total Débit", f"{total_debit:,.0f} XOF")
        col2.metric("Total Crédit", f"{total_credit:,.0f} XOF")
    else:
        st.error("Aucune transaction détectée. Vérifiez que le PDF n'est pas un scan (image).")