"""
app.py - Interface utilisateur Streamlit.

Ce module gère UNIQUEMENT la couche de présentation :
- Upload et déduplication des fichiers sources.
- Mapping interactif des colonnes.
- Déclenchement du pipeline (délégué à utils.py).
- Affichage des résultats et téléchargements.

Toute logique de calcul, nettoyage ou export est dans utils.py.
Toute configuration est dans config.py.
"""

from __future__ import annotations

import logging
from datetime import datetime

import pandas as pd
import streamlit as st

from config import SETTINGS
from utils import (
    add_month_column,
    aggregate_by_month,
    aggregate_by_salesperson,
    build_excel_bytes,
    build_pdf_bytes,
    clean_data,
    format_excel_bytes,
    hash_bytes,
    read_uploaded_excels,
    reconcile_data,
)

logging.getLogger().setLevel(logging.ERROR)

# ---------------------------------------------------------------------------
# PAGE
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Reporting (Excel)", layout="centered")
st.title("📊 Reporting - Excel-only")
st.write("Upload des .xlsx → mapping colonnes → génération → téléchargement Excel + PDF + log.")


# ---------------------------------------------------------------------------
# FONCTIONS HELPERS (UI uniquement - logique propre à l'interface)
# ---------------------------------------------------------------------------

def _normalize_and_merge_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise les noms de colonnes et fusionne les doublons pour le mapping UI.

    Cette étape est pré-mapping (avant clean_data()) : elle sert uniquement
    à alimenter les selectbox avec des noms de colonnes propres.
    clean_data() refait cette normalisation en interne sur les données réelles.

    Args:
        df: DataFrame brut avec colonnes potentiellement "sales".

    Returns:
        DataFrame avec noms de colonnes normalisés et doublons fusionnés.
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()

    if df.columns.duplicated().any():
        for col_name in df.columns[df.columns.duplicated()].unique():
            sub_df = df.loc[:, df.columns == col_name]
            merged_value = sub_df.bfill(axis=1).iloc[:, 0]
            df = df.drop(columns=col_name)
            df[col_name] = merged_value
    return df


def _default_selectbox_index(
    cols: list[str],
    preferred: list[str],
    fallback: int = 0,
) -> int:
    """Retourne l'index de la première colonne préférée trouvée dans la liste.

    Permet une pré-sélection intelligente des selectbox lors du mapping :
    si la colonne "date" existe déjà, elle est sélectionnée par défaut.

    Args:
        cols: Liste des colonnes disponibles.
        preferred: Noms de colonnes à rechercher en priorité.
        fallback: Index par défaut si aucune colonne préférée n'est trouvée.

    Returns:
        Index de la colonne à pré-sélectionner dans le selectbox.
    """
    for p in preferred:
        if p in cols:
            return cols.index(p)
    return fallback


def _format_currency_fr(x: float) -> str:
    """Formate un float en montant français lisible : 1234567.89 → '1 234 567,89 €'.

    Args:
        x: Montant numérique.

    Returns:
        Chaîne formatée avec séparateur de milliers (espace) et décimale (virgule).
    """
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{s} €"


def _build_log_text(lines: list[str]) -> str:
    """Assemble une liste de lignes en texte de log téléchargeable.

    Args:
        lines: Lignes du log.

    Returns:
        Texte complet terminé par un saut de ligne.
    """
    return "\n".join(lines) + "\n"


def _compute_quality_metrics(df_mapped: pd.DataFrame) -> dict[str, int]:
    """Calcule les métriques qualité AVANT nettoyage final (pour affichage + log).

    Ces métriques sont calculées sur df_mapped (post-renommage, pré-clean_data)
    pour permettre à l'utilisateur de comprendre ce que le pipeline va corriger.

    Args:
        df_mapped: DataFrame après renommage des colonnes par l'utilisateur,
            avant appel à clean_data().

    Returns:
        Dictionnaire avec les compteurs : dates_invalides, montants_vides_source,
        montants_invalides_hors_vides.
    """
    metrics: dict[str, int] = {
        "dates_invalides": 0,
        "montants_vides_source": 0,
        "montants_invalides_hors_vides": 0,
    }

    if "date" in df_mapped.columns:
        parsed_dates = pd.to_datetime(df_mapped["date"], dayfirst=True, errors="coerce")
        metrics["dates_invalides"] = int(parsed_dates.isna().sum())

    if "montant" in df_mapped.columns:
        raw_amounts = df_mapped["montant"]
        amounts_str = raw_amounts.astype(str).str.strip()
        empty_mask = raw_amounts.isna() | (amounts_str == "")
        metrics["montants_vides_source"] = int(empty_mask.sum())

        num = pd.to_numeric(
            amounts_str.str.replace(",", ".", regex=False)
            .str.replace("\u00a0", "", regex=False),
            errors="coerce",
        )
        metrics["montants_invalides_hors_vides"] = int((num.isna() & ~empty_mask).sum())

    return metrics


def _friendly_file_error(err: str) -> str:
    """Traduit un message d'erreur technique en libellé compréhensible PME.

    Args:
        err: Message d'erreur technique levé lors de la lecture du fichier.

    Returns:
        Message en français adapté à un utilisateur non technique.
    """
    e = (err or "").lower()
    if "excel file format cannot be determined" in e:
        return "Format Excel non reconnu (fichier cassé, mauvais format ou renommé en .xlsx)."
    if "password" in e or "encrypted" in e:
        return "Fichier protégé par mot de passe (impossible à lire automatiquement)."
    if "permission" in e or "access is denied" in e:
        return "Accès refusé (droits insuffisants ou fichier verrouillé)."
    if "zipfile" in e:
        return "Fichier Excel corrompu (structure ZIP interne illisible)."
    if "openpyxl" in e:
        return "Erreur de lecture openpyxl (fichier potentiellement incompatible)."
    return "Le fichier n'est pas un .xlsx valide (souvent : renommé ou corrompu)."


# ---------------------------------------------------------------------------
# UPLOAD & DÉDUPLICATION
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Fichiers Excel à consolider (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("Ajoute 1+ fichiers Excel pour commencer.")
    st.stop()

# Déduplication par hash SHA-256 du contenu.
# Deux fichiers au même contenu sont des doublons (même si noms différents).
unique_files = []
seen_hashes: dict[str, str] = {}
duplicates: list[str] = []

for f in uploaded:
    file_hash = hash_bytes(f.getvalue())
    if file_hash in seen_hashes:
        duplicates.append(f.name)
    else:
        seen_hashes[file_hash] = f.name
        unique_files.append(f)

if duplicates:
    st.warning(f"Doublons ignorés (même contenu) : {', '.join(duplicates)}")

# ---------------------------------------------------------------------------
# LECTURE DES FICHIERS
# ---------------------------------------------------------------------------

df_raw, failed_files = read_uploaded_excels(unique_files)

if failed_files:
    st.warning("⚠️ Certains fichiers ont été ignorés car illisibles :")
    for name, err in failed_files:
        st.write(f"- **{name}** - {_friendly_file_error(err)}")
    with st.expander("Détails techniques", expanded=False):
        for name, err in failed_files:
            st.write(f"- {name} - {err}")

valid_file_count = len(unique_files) - len(failed_files)
source_row_count = len(df_raw)
st.success(
    f"{valid_file_count} fichier(s) OK lus. "
    f"Total lignes (avant nettoyage) : {source_row_count}"
)

# ---------------------------------------------------------------------------
# MAPPING DES COLONNES
# ---------------------------------------------------------------------------

df_base = _normalize_and_merge_columns(df_raw)
cols = list(df_base.columns)

st.subheader("🧩 Mapping des colonnes")

with st.form("mapping_form"):
    date_col = st.selectbox(
        "Colonne DATE", options=cols, index=_default_selectbox_index(cols, ["date"])
    )
    amount_col = st.selectbox(
        "Colonne MONTANT", options=cols, index=_default_selectbox_index(cols, ["montant"])
    )
    salesperson_col = st.selectbox(
        "Colonne COMMERCIAL (optionnel)",
        options=["(aucune)"] + cols,
        index=(["(aucune)"] + cols).index("commercial") if "commercial" in cols else 0,
    )
    submitted = st.form_submit_button("⚙️ Générer le reporting")

# Logique de persistance de l'affichage :
#   - submitted=True  → on génère + on stocke dans session_state
#   - submitted=False + "generated" en session_state → résultats déjà là → on continue
#   - submitted=False + pas de "generated" → premier chargement → st.stop()
if not submitted:
    if "generated" not in st.session_state:
        st.stop()

if submitted and date_col == amount_col:
    st.error("Tu as sélectionné la même colonne pour DATE et MONTANT.")
    st.stop()

# ---------------------------------------------------------------------------
# GÉNÉRATION DU REPORTING
# ---------------------------------------------------------------------------

if submitted:
    rename_map = {date_col: "date", amount_col: "montant"}
    if salesperson_col != "(aucune)":
        rename_map[salesperson_col] = "commercial"

    df_mapped = df_base.rename(columns=rename_map)

    missing_cols = [c for c in ["date", "montant"] if c not in df_mapped.columns]
    if missing_cols:
        st.error(f"Colonnes manquantes après mapping : {missing_cols}.")
        st.stop()

    quality = _compute_quality_metrics(df_mapped)

    # --- Construction du log d'exécution ---
    now = datetime.now()
    log_lines: list[str] = []

    log_lines += [
        "=== LOG D'EXECUTION (STREAMLIT) ===",
        f"Application : {SETTINGS.app_name} v{SETTINGS.app_version}",
        f"Horodatage  : {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Fichiers uploadés (uniques) : {len(unique_files)}",
        f"Fichiers OK lus : {valid_file_count}",
        "",
        "--- Empreintes SHA-256 (audit d'intégrité des sources) ---",
    ]
    for f in unique_files:
        h = hash_bytes(f.getvalue())
        log_lines.append(f"  {f.name} | {SETTINGS.hash_algorithm.upper()} : {h}")

    log_lines += [
        "",
        "=== FICHIERS EN ERREUR ===",
        *([f"  - {n} | {e}" for n, e in failed_files] if failed_files else ["  Aucun"]),
        "",
        f"Lignes avant nettoyage : {source_row_count}",
        "",
        "=== MAPPING ===",
        f"  DATE       : {date_col} → date",
        f"  MONTANT    : {amount_col} → montant",
        f"  COMMERCIAL : {salesperson_col} → "
        f"{'commercial' if salesperson_col != '(aucune)' else '(aucune)'}",
        "",
        "=== QUALITE DES DONNEES (avant nettoyage) ===",
        f"  Dates invalides                : {quality['dates_invalides']}",
        f"  Montants vides (source)        : {quality['montants_vides_source']}",
        f"  Montants invalides (hors vides): {quality['montants_invalides_hors_vides']}",
    ]

    try:
        df = clean_data(df_mapped)
        df = add_month_column(df)

        # --- CHECKPOINT D'INTÉGRITÉ DES DONNEES ---
        # df_mapped = source normalisée pré-nettoyage
        # df        = données après clean_data() + add_month_column()
        reconciliation_report = reconcile_data(df_mapped, df)

        report_months = aggregate_by_month(df)
        report_salespeople = aggregate_by_salesperson(df)

        processed_row_count = len(df)
        period_min = df["mois"].min()
        period_max = df["mois"].max()
        total_amount = float(df["montant"].sum())
        salesperson_count = (
            int(df["commercial"].nunique()) if "commercial" in df.columns else 0
        )
        tag = (
            f"{period_min}"
            if period_min == period_max
            else f"{period_min}_to_{period_max}"
        )

        excel_bytes = build_excel_bytes(df, report_months, report_salespeople)
        excel_bytes = format_excel_bytes(excel_bytes)
        pdf_bytes = build_pdf_bytes(report_months, report_salespeople)

        log_lines += [
            "",
            "=== RÉSULTATS ===",
            f"  Lignes après nettoyage : {processed_row_count}",
            f"  Lignes supprimées      : {source_row_count - processed_row_count}",
            "  Règle : lignes sans date valide supprimées ; "
            "montants vides → NaN, exclus du CA",
            f"  Période détectée       : {period_min} → {period_max}",
            f"  Chiffre d'affaires     : {_format_currency_fr(total_amount)}",
        ]
        if salesperson_col != "(aucune)":
            log_lines.append(f"  Commerciaux distincts  : {salesperson_count}")

        log_lines += [
            "",
            "=== RÉCONCILIATION DES DONNEES (Checkpoint d'intégrité) ===",
            f"  {reconciliation_report.message}",
            f"  Somme source attendue  : "
            f"{_format_currency_fr(reconciliation_report.expected_source_sum)}",
            f"  Somme traitée          : "
            f"{_format_currency_fr(reconciliation_report.processed_sum)}",
            f"  Écart absolu           : {reconciliation_report.absolute_gap:.4f} €",
            f"  Dates invalides        : {reconciliation_report.invalid_date_count}",
            f"  Montants invalides     : {reconciliation_report.invalid_amount_count}",
        ]

        log_bytes = _build_log_text(log_lines).encode("utf-8")

        # --- Persistance dans session_state ---
        st.session_state.update({
            "generated": True,
            "status": "ok",
            "error_message": None,
            "tag": tag,
            "df": df,
            "report_months": report_months,
            "report_salespeople": report_salespeople,
            "reconciliation_report": reconciliation_report,
            "summary": {
                "valid_file_count": valid_file_count,
                "source_row_count": source_row_count,
                "processed_row_count": processed_row_count,
                "dropped_row_count": source_row_count - processed_row_count,
                "period_min": period_min,
                "period_max": period_max,
                "total_amount": total_amount,
                "salesperson_count": salesperson_count,
                "has_salesperson": (salesperson_col != "(aucune)"),
            },
            "quality": quality,
            "excel_bytes": excel_bytes,
            "pdf_bytes": pdf_bytes,
            "log_bytes": log_bytes,
        })

    except Exception as e:
        log_lines += ["", "=== ERREUR ===", str(e)]
        st.session_state.update({
            "generated": True,
            "status": "error",
            "error_message": str(e),
            "tag": now.strftime("%Y-%m-%d_%H-%M-%S"),
            "log_bytes": _build_log_text(log_lines).encode("utf-8"),
        })
        st.error(f"Erreur : {e}")
        st.stop()


# ---------------------------------------------------------------------------
# AFFICHAGE PERSISTANT DES RÉSULTATS
# ---------------------------------------------------------------------------

if "generated" not in st.session_state:
    st.stop()

tag = st.session_state.tag

if st.session_state.get("status") == "error":
    st.subheader("❌ Génération impossible")
    st.error(st.session_state.get("error_message", "Erreur inconnue."))
    st.download_button(
        "⬇️ Télécharger le log d'exécution",
        data=st.session_state.log_bytes,
        file_name=f"log_{tag}.txt",
        mime="text/plain",
    )
    st.stop()

s = st.session_state.summary
q = st.session_state.quality
recon = st.session_state.reconciliation_report

# --- Résumé ---
st.subheader("✅ Résumé")
st.write(f"- Fichiers OK lus : **{s['valid_file_count']}**")
st.write(f"- Lignes avant nettoyage : **{s['source_row_count']}**")
st.write(f"- Lignes après nettoyage : **{s['processed_row_count']}**")
st.write(f"- Lignes supprimées : **{s['dropped_row_count']}**")
st.write(
    "- Règle : **lignes sans date valide supprimées** ; "
    "montants vides → NaN, **exclus du CA**."
)
st.write(f"- Période détectée : **{s['period_min']} → {s['period_max']}**")
st.write(f"- Chiffre d'affaires total : **{_format_currency_fr(s['total_amount'])}**")
if s["has_salesperson"]:
    st.write(f"- Nombre de commerciaux : **{s['salesperson_count']}**")

# --- Qualité des données ---
st.subheader("🧪 Qualité des données")
st.write(f"- Dates invalides (parse impossible) : **{q['dates_invalides']}**")
st.write(f"- Montants vides (source) : **{q['montants_vides_source']}**")
st.write(f"- Montants invalides (hors vides) : **{q['montants_invalides_hors_vides']}**")

# --- Réconciliation des données ---
st.subheader("🔐 Réconciliation des données")
if recon.integrity_ok:
    st.success(recon.message)
else:
    st.error(recon.message)
    st.warning(
        "⚠️ Un écart anormal a été détecté. "
        "Télécharge le log pour investiguer avant de diffuser ce reporting."
    )

with st.expander("Détails du checkpoint d'intégrité", expanded=False):
    st.write(
        f"- Somme source attendue (lignes à date valide) : "
        f"**{_format_currency_fr(recon.expected_source_sum)}**"
    )
    st.write(
        f"- Somme traitée (après nettoyage) : "
        f"**{_format_currency_fr(recon.processed_sum)}**"
    )
    st.write(
        f"- Écart absolu : **{recon.absolute_gap:.4f} €** "
        f"({recon.gap_pct * 100:.4f} %)"
    )
    st.write(
        f"- Tolérance appliquée : "
        f"**{SETTINGS.reconciliation_tolerance_pct * 100:.2f} %**"
    )
    st.write(f"- Lignes supprimées (dates invalides) : **{recon.invalid_date_count}**")
    st.write(f"- Montants non convertibles (→ NaN) : **{recon.invalid_amount_count}**")

# --- Aperçu et tableaux ---
st.subheader("Aperçu (après nettoyage)")
st.dataframe(st.session_state.df.head(30), use_container_width=True)

st.subheader("Total par mois")
st.dataframe(st.session_state.report_months, use_container_width=True)

if s["has_salesperson"]:
    st.subheader("Total par commercial")
    st.dataframe(st.session_state.report_salespeople, use_container_width=True)

# --- Téléchargements ---
st.download_button(
    "⬇️ Télécharger l'Excel multi-feuilles (formaté)",
    data=st.session_state.excel_bytes,
    file_name=f"reporting_{tag}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.download_button(
    "⬇️ Télécharger le PDF",
    data=st.session_state.pdf_bytes,
    file_name=f"rapport_{tag}.pdf",
    mime="application/pdf",
)
st.download_button(
    "⬇️ Télécharger le log d'exécution",
    data=st.session_state.log_bytes,
    file_name=f"log_{tag}.txt",
    mime="text/plain",
)