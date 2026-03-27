# -*- coding: utf-8 -*-
# app.py — Duty online version (Streamlit)
# Basé sur la logique de Duty.py, mais avec upload web au lieu de D:\IMA

from __future__ import annotations

import io
import re
from decimal import Decimal, InvalidOperation
from collections import defaultdict

import fitz  # pymupdf
import pandas as pd
import streamlit as st

# =========================
# CONFIG / REGEX
# =========================

# nombres acceptés : 419.83 / 419,83 / 16 245,10 / 0 / 0,00
NUM_TOKEN = re.compile(r"^\d{1,3}(?:\s?\d{3})*(?:[.,]\d+)?$")

# Détection flexible conteneur / référence
# Exemples:
# - SK013
# - OOCU9729247
# - AB12345678
PREFIX_RE = re.compile(r"^(SK\d{3}|[A-Z]{4}\d{7}|[A-Z0-9]{8,15})", re.IGNORECASE)


# =========================
# OUTILS EXTRACTION
# =========================

def to_decimal(s: str) -> Decimal | None:
    s = s.replace(" ", "").replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def has_currency_near(words, idx: int, radius: int = 3) -> bool:
    """Vrai si 'EUR'/'USD' est très proche sur la même ligne."""
    line = words[idx][6]
    lo = max(0, idx - radius)
    hi = min(len(words), idx + radius + 1)
    for j in range(lo, hi):
        if j == idx:
            continue
        if words[j][6] == line and str(words[j][4]).lower() in {"eur", "usd"}:
            return True
    return False


def in_rect(w, rect) -> bool:
    rx0, ry0, rx1, ry1 = rect
    return (w[0] >= rx0 and w[2] <= rx1 and w[1] >= ry0 and w[3] <= ry1)


def find_amount_right_below(words, anchor_bbox) -> Decimal | None:
    ax0, ay0, ax1, ay1 = anchor_bbox

    # Zone à droite et zone en dessous du texte "Paiement comptant"
    right_rect = (ax1, ay0 - 5, ax1 + 200, ay1 + 60)
    below_rect = (ax0 - 80, ay1, ax1 + 80, ay1 + 140)

    candidates = []

    for k, w in enumerate(words):
        token = str(w[4]).strip()

        if not NUM_TOKEN.match(token):
            continue

        if in_rect(w, right_rect) or in_rect(w, below_rect):
            # si une devise est proche, on garde surtout la zone à droite
            if has_currency_near(words, k) and not in_rect(w, right_rect):
                continue

            dy = max(0, min(abs(w[1] - ay1), abs(w[3] - ay1)))
            val = to_decimal(token)
            if val is not None:
                candidates.append((dy, w[0], val))

    if not candidates:
        return None

    candidates.sort()
    return candidates[0][2]


def find_paiement_comptant_amount(page) -> Decimal | None:
    words = page.get_text("words")
    if not words:
        return None

    words.sort(key=lambda w: (round(w[1], 1), w[0]))
    lower = [(w[0], w[1], w[2], w[3], str(w[4]).lower(), w[5], w[6], w[7]) for w in words]

    for i in range(len(lower) - 1):
        x0, y0, x1, y1, t0, b0, l0, w0 = lower[i]
        x0n, y0n, x1n, y1n, t1, b1, l1, w1 = lower[i + 1]

        if t0 == "paiement" and (t1.startswith("comptant") or t1 == "comptant:"):
            ax0 = min(x0, x0n)
            ay0 = min(y0, y0n)
            ax1 = max(x1, x1n)
            ay1 = max(y1, y1n)

            amt = find_amount_right_below(lower, (ax0, ay0, ax1, ay1))
            if amt is not None:
                return amt

    return None


def detect_prefix(filename: str) -> str | None:
    stem = filename.rsplit(".", 1)[0].upper()
    m = PREFIX_RE.match(stem)
    if not m:
        return None
    return m.group(1)


def process_uploaded_pdf(uploaded_file) -> dict:
    """
    Retourne un dict détaillé par fichier:
    - filename
    - prefix
    - amount
    - status
    - error
    """
    result = {
        "filename": uploaded_file.name,
        "prefix": None,
        "amount": Decimal("0"),
        "status": "OK",
        "error": "",
    }

    prefix = detect_prefix(uploaded_file.name)
    if not prefix:
        result["status"] = "IGNORED"
        result["error"] = "Aucun préfixe détecté"
        return result

    result["prefix"] = prefix

    try:
        file_bytes = uploaded_file.read()

        amount = None
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            for page in doc:
                amount = find_paiement_comptant_amount(page)
                if amount is not None:
                    break

        if amount is None:
            amount = Decimal("0")

        result["amount"] = amount
        return result

    except Exception as e:
        result["status"] = "ERROR"
        result["error"] = str(e)
        result["amount"] = Decimal("0")
        return result


def build_summary(details: list[dict]) -> pd.DataFrame:
    totals = defaultdict(lambda: Decimal("0"))
    counts = defaultdict(int)

    for row in details:
        if row["prefix"] is None:
            continue
        counts[row["prefix"]] += 1
        totals[row["prefix"]] += row["amount"]

    summary_rows = []
    for prefix in sorted(totals.keys()):
        summary_rows.append({
            "Préfixe": prefix,
            "Total EUR": float(totals[prefix]),
            "Nombre de fichiers": counts[prefix],
        })

    return pd.DataFrame(summary_rows)


def details_to_dataframe(details: list[dict]) -> pd.DataFrame:
    rows = []
    for row in details:
        rows.append({
            "Fichier": row["filename"],
            "Préfixe": row["prefix"] or "",
            "Montant EUR": float(row["amount"]),
            "Statut": row["status"],
            "Erreur": row["error"],
        })
    return pd.DataFrame(rows)


def make_excel_file(summary_df: pd.DataFrame, details_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        details_df.to_excel(writer, index=False, sheet_name="Details")
    output.seek(0)
    return output.getvalue()


def format_amount(x: float) -> str:
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


# =========================
# UI STREAMLIT
# =========================

st.set_page_config(
    page_title="Duty PDF Analyzer",
    layout="wide"
)

st.title("Duty PDF Analyzer")
st.caption("Analyse des PDF et récapitulatif par conteneur / référence")

with st.expander("Comment ça marche", expanded=False):
    st.write(
        """
- Upload un ou plusieurs PDF
- L'application cherche le montant près de **Paiement comptant**
- Elle détecte le préfixe depuis le nom du fichier
- Elle génère :
  - un détail par fichier
  - un récap par préfixe
  - un export Excel / CSV
        """
    )

uploaded_files = st.file_uploader(
    "Glisse tes PDF ici",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} fichier(s) prêt(s) à analyser.")

    if st.button("Analyser les PDF", type="primary"):
        details = []

        progress = st.progress(0)
        status_box = st.empty()

        total_files = len(uploaded_files)

        for i, up in enumerate(uploaded_files, start=1):
            status_box.write(f"Lecture : {up.name}")
            details.append(process_uploaded_pdf(up))
            progress.progress(i / total_files)

        summary_df = build_summary(details)
        details_df = details_to_dataframe(details)

        st.success("Analyse terminée.")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("PDF traités", len(details))
        with col2:
            ok_count = int((details_df["Statut"] == "OK").sum()) if not details_df.empty else 0
            st.metric("PDF OK", ok_count)
        with col3:
            total_eur = summary_df["Total EUR"].sum() if not summary_df.empty else 0
            st.metric("Total EUR", format_amount(total_eur))

        st.subheader("Récap par préfixe")
        if summary_df.empty:
            st.warning("Aucun préfixe valide trouvé.")
        else:
            st.dataframe(summary_df, use_container_width=True)

        st.subheader("Détail par fichier")
        st.dataframe(details_df, use_container_width=True)

        # CSV summary
        summary_csv = summary_df.to_csv(index=False).encode("utf-8-sig")
        details_csv = details_df.to_csv(index=False).encode("utf-8-sig")
        excel_file = make_excel_file(summary_df, details_df)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                "Télécharger Summary CSV",
                data=summary_csv,
                file_name="duty_summary.csv",
                mime="text/csv"
            )
        with c2:
            st.download_button(
                "Télécharger Details CSV",
                data=details_csv,
                file_name="duty_details.csv",
                mime="text/csv"
            )
        with c3:
            st.download_button(
                "Télécharger Excel",
                data=excel_file,
                file_name="duty_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("Ajoute au moins un PDF pour commencer.")
