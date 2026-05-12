"""
app.py - Interface utilisateur Streamlit.

Ce module gère UNIQUEMENT la couche de présentation :
- Upload et déduplication des fichiers sources (.xlsx et .csv).
  → Déduplication niveau bytes (SHA-256) : même contenu binaire.
  → Déduplication niveau métier : même données, formats différents (.xlsx/.csv).
- Mapping interactif des colonnes.
- Déclenchement du pipeline (délégué à utils.py).
- Affichage des résultats et téléchargements.

Toute logique de calcul, nettoyage ou export est dans utils.py.
Toute configuration est dans config.py.
"""

from __future__ import annotations

import hashlib
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
    read_one_uploaded_file,
    reconcile_data,
)

logging.getLogger().setLevel(logging.ERROR)

# ---------------------------------------------------------------------------
# PAGE
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Reporting (Excel/CSV)", layout="centered")
st.title("📊 Reporting - Excel & CSV")
st.write(
    "Upload des .xlsx ou .csv → mapping colonnes → génération → "
    "téléchargement Excel + PDF + log."
)


# ---------------------------------------------------------------------------
# FONCTIONS HELPERS (UI uniquement)
# ---------------------------------------------------------------------------

def _normalize_and_merge_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise les noms de colonnes et fusionne les doublons pour le mapping UI."""
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
    """Retourne l'index de la première colonne préférée trouvée dans la liste."""
    for p in preferred:
        if p in cols:
            return cols.index(p)
    return fallback


def _format_currency_fr(x: float) -> str:
    """Formate un float en montant français lisible : 1234567.89 → '1 234 567,89 €'."""
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{s} €"


def _build_log_text(lines: list[str]) -> str:
    """Assemble une liste de lignes en texte de log téléchargeable."""
    return "\n".join(lines) + "\n"


def _compute_quality_metrics(df_mapped: pd.DataFrame) -> dict[str, int]:
    """Calcule les métriques qualité AVANT nettoyage final (pour affichage + log)."""
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
    """Traduit un message d'erreur technique en libellé compréhensible PME."""
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
    if "encodages testés" in e or "unicodedecode" in e:
        return (
            "Impossible de détecter l'encodage du CSV. "
            "Essayez de l'ouvrir dans Excel et de le réenregistrer en 'CSV UTF-8'."
        )
    if "non supporté" in e or "non supportée" in e:
        return "Format non supporté. Seuls les fichiers .xlsx et .csv sont acceptés."
    return "Le fichier n'est pas lisible (format non reconnu, corrompu, ou vide)."


def _normalize_col_for_fingerprint(series: pd.Series) -> pd.Series:
    """Normalise une colonne en une représentation canonique inter-formats.

    C'est le cœur de la détection inter-formats. Le même fichier lu depuis
    .xlsx et depuis .csv produit des types pandas différents pour les mêmes
    données :
    - Nombres : Excel → float  (100.0)  /  CSV → string ("100" ou "100,0")
    - Dates   : Excel → datetime64      /  CSV → string ("01/01/2025")
    - Texte   : identique dans les deux cas

    L'approche est donc de détecter le type sémantique de chaque colonne
    (numérique → date → texte) et d'appliquer une normalisation cohérente.

    Args:
        series: Colonne brute d'un DataFrame avant clean_data().

    Returns:
        Series de strings canoniques comparables quel que soit le format source.
    """
    # Représentation string de base : strip pour tous les types
    s = series.astype(str).str.strip()

    # --- Essai numérique ---
    # On prétraite comme clean_data() le ferait : virgule → point, espaces milliers.
    # Ratio > 0.5 : la majorité des valeurs non-vides sont numériques → colonne numérique.
    numeric_attempt = pd.to_numeric(
        s.str.replace(",", ".", regex=False)
         .str.replace("\u00a0", "", regex=False)
         .str.replace(" ", "", regex=False),
        errors="coerce",
    )
    non_empty = s.replace({"nan": "", "none": "", "": pd.NA}).notna().sum()
    numeric_ratio = numeric_attempt.notna().sum() / max(non_empty, 1)
    if numeric_ratio >= 0.5:
        # Arrondi à 4 décimales pour absorber les micro-différences de précision
        # flottante IEEE 754 entre les deux lectures (ex: 100.5000000001 vs 100.5).
        return numeric_attempt.round(4).astype(str).str.replace("nan", "", regex=False)

    # --- Essai date ---
    # dayfirst=True : cohérent avec clean_data() et les exports FR (JJ/MM/AAAA).
    # Résultat normalisé en ISO 8601 (YYYY-MM-DD) : même valeur quel que soit
    # le format source ("01/01/2025" et "2025-01-01 00:00:00" → "2025-01-01").
    date_attempt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    date_ratio = date_attempt.notna().sum() / max(non_empty, 1)
    if date_ratio >= 0.5:
        return date_attempt.dt.strftime("%Y-%m-%d").fillna("")

    # --- Fallback texte ---
    return s.str.lower()


def _compute_df_fingerprint(df: pd.DataFrame) -> str:
    """Calcule une empreinte normalisée du contenu d'un DataFrame.

    Permet de détecter les doublons inter-formats : un fichier .xlsx et son
    export .csv ont des bytes totalement différents (SHA-256 bytes ≠), mais
    leurs données normalisées par ``_normalize_col_for_fingerprint`` sont
    identiques → même empreinte métier.

    **Algorithme :**
    1. Normalisation des noms de colonnes (strip + lower + tri alphabétique).
    2. Normalisation sémantique de chaque colonne via ``_normalize_col_for_fingerprint``
       (numérique → date → texte) : c'est ce qui résout le problème
       "100.0" (Excel) ≠ "100" (CSV) de la version précédente.
    3. Hash SHA-256 sur forme + colonnes + échantillon (25+25 lignes).

    Args:
        df: DataFrame brut tel que retourné par la lecture (avant clean_data).

    Returns:
        Empreinte hexadécimale SHA-256 du contenu normalisé (64 caractères).
    """
    df_norm = df.copy()
    # Normalisation des colonnes : strip + lower + tri alphabétique
    df_norm.columns = df_norm.columns.astype(str).str.strip().str.lower()
    df_norm = df_norm.reindex(sorted(df_norm.columns), axis=1)
    # Normalisation sémantique colonne par colonne
    df_norm = df_norm.apply(_normalize_col_for_fingerprint)
    df_norm = df_norm.fillna("")
    # Échantillonnage pour la performance sur les gros fichiers
    if len(df_norm) > 50:
        sample = pd.concat([df_norm.head(25), df_norm.tail(25)])
    else:
        sample = df_norm
    fingerprint_str = (
        f"shape={df_norm.shape}"
        f"|cols={list(df_norm.columns)}"
        f"|data={sample.to_csv(index=False)}"
    )
    return hashlib.sha256(fingerprint_str.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Fichiers sources à consolider (.xlsx ou .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
)

# Quand l'uploader est vide (fichiers supprimés par l'utilisateur), on efface
# tous les résultats de la session pour qu'ils ne réapparaissent pas au prochain
# upload sans nouvelle génération.
if not uploaded:
    for key in list(st.session_state.keys()):
        st.session_state.pop(key)
    st.info("Ajoute 1+ fichiers Excel (.xlsx) ou CSV (.csv) pour commencer.")
    st.stop()


# ---------------------------------------------------------------------------
# DÉDUPLICATION NIVEAU BYTES (même contenu binaire, tous formats confondus)
# ---------------------------------------------------------------------------

unique_files = []
seen_hashes: dict[str, str] = {}   # bytes_hash → nom du fichier conservé
duplicates_bytes: list[str] = []

for f in uploaded:
    file_hash = hash_bytes(f.getvalue())
    if file_hash in seen_hashes:
        duplicates_bytes.append(f.name)
    else:
        seen_hashes[file_hash] = f.name
        unique_files.append(f)

if duplicates_bytes:
    st.warning(
        "Doublons ignorés (contenu identique octet pour octet) : "
        f"{', '.join(duplicates_bytes)}"
    )

# Empreinte de l'ensemble de fichiers actuellement uploadé.
# Sert à invalider les résultats si les fichiers changent entre deux générations.
current_upload_sig: frozenset[str] = frozenset(seen_hashes.keys())


# ---------------------------------------------------------------------------
# LECTURE INDIVIDUELLE + DÉDUPLICATION NIVEAU CONTENU (inter-formats)
#
# On lit chaque fichier séparément pour calculer une empreinte *métier*
# par fichier. Cela permet de détecter qu'un .xlsx et un .csv contiennent
# les mêmes données (même fichier converti d'un format à l'autre).
# ---------------------------------------------------------------------------

per_file_dfs: list[tuple[str, pd.DataFrame]] = []
failed_files: list[tuple[str, str]] = []
seen_content_hashes: dict[str, str] = {}           # content_hash → nom conservé
duplicates_content: list[tuple[str, str]] = []     # (nom_doublon, nom_original)

for f in unique_files:
    filename = f.name
    try:
        df_file = read_one_uploaded_file(f)
        content_hash = _compute_df_fingerprint(df_file)

        if content_hash in seen_content_hashes:
            duplicates_content.append((filename, seen_content_hashes[content_hash]))
        else:
            seen_content_hashes[content_hash] = filename
            per_file_dfs.append((filename, df_file))

    except Exception as exc:
        failed_files.append((filename, str(exc)))

# Affichage des doublons inter-formats
if duplicates_content:
    for dup_name, orig_name in duplicates_content:
        st.warning(
            f"⚠️ Doublon inter-formats ignoré : **{dup_name}** contient les mêmes "
            f"données que **{orig_name}** (probablement le même fichier converti)."
        )

# Affichage des fichiers illisibles
if failed_files:
    st.warning("⚠️ Certains fichiers ont été ignorés car illisibles :")
    for name, err in failed_files:
        st.write(f"- **{name}** — {_friendly_file_error(err)}")
    with st.expander("Détails techniques", expanded=False):
        for name, err in failed_files:
            st.write(f"- {name} — {err}")

if not per_file_dfs:
    st.error("Aucun fichier lisible. Vérifie les formats et les messages ci-dessus.")
    st.stop()

# Consolidation en un seul DataFrame brut
df_raw = pd.concat([df for _, df in per_file_dfs], ignore_index=True)
valid_file_count = len(per_file_dfs)
source_row_count = len(df_raw)

xlsx_count = sum(1 for name, _ in per_file_dfs if name.lower().endswith(".xlsx"))
csv_count  = sum(1 for name, _ in per_file_dfs if name.lower().endswith(".csv"))
format_parts = (["Excel"] if xlsx_count else []) + (["CSV"] if csv_count else [])

st.success(
    f"{valid_file_count} fichier(s) OK ({' + '.join(format_parts) or '?'}). "
    f"Total lignes avant nettoyage : {source_row_count}"
)


# ---------------------------------------------------------------------------
# MAPPING DES COLONNES
# ---------------------------------------------------------------------------

df_base = _normalize_and_merge_columns(df_raw)
cols = list(df_base.columns)

st.subheader("🧩 Mapping des colonnes")

with st.form("mapping_form"):
    date_col = st.selectbox(
        "Colonne DATE", options=cols,
        index=_default_selectbox_index(cols, ["date"]),
    )
    amount_col = st.selectbox(
        "Colonne MONTANT", options=cols,
        index=_default_selectbox_index(cols, ["montant"]),
    )
    salesperson_col = st.selectbox(
        "Colonne COMMERCIAL (optionnel)",
        options=["(aucune)"] + cols,
        index=(["(aucune)"] + cols).index("commercial") if "commercial" in cols else 0,
    )
    submitted = st.form_submit_button("⚙️ Générer le reporting")

if submitted and date_col == amount_col:
    st.error("Tu as sélectionné la même colonne pour DATE et MONTANT.")
    st.stop()


# ---------------------------------------------------------------------------
# INVALIDATION DES RÉSULTATS SI LES FICHIERS ONT CHANGÉ
#
# Scénario visé :
#   1. Génération avec fichiers A → résultats stockés en session_state
#   2. Upload de fichiers B sans reclicker Générer
#   → on bloque l'affichage des résultats de A et on invite à regénérer
#
# Mécanisme : à chaque génération, on stocke l'empreinte de l'ensemble de
# fichiers utilisé (generation_upload_sig). Si elle diffère de l'empreinte
# courante (current_upload_sig), les résultats affichés ne correspondraient
# plus aux fichiers présents → on n'affiche rien.
# ---------------------------------------------------------------------------

files_changed = (
    "generation_upload_sig" in st.session_state
    and st.session_state["generation_upload_sig"] != current_upload_sig
)

if not submitted and files_changed:
    st.info(
        "📂 Les fichiers ont changé depuis la dernière génération. "
        "Appuie sur **⚙️ Générer le reporting** pour traiter les nouveaux fichiers."
    )
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
    now = datetime.now()
    log_lines: list[str] = []

    log_lines += [
        "=== LOG D'EXECUTION (STREAMLIT) ===",
        f"Application : {SETTINGS.app_name} v{SETTINGS.app_version}",
        f"Horodatage  : {now.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Fichiers uploadés (uniques) : {len(unique_files)} "
        f"({xlsx_count} Excel, {csv_count} CSV)",
        f"Fichiers conservés après déduplication contenu : {valid_file_count}",
        "",
        "--- Empreintes SHA-256 (audit d'intégrité des sources) ---",
    ]
    for name, _ in per_file_dfs:
        file_hash = next(h for h, n in seen_hashes.items() if n == name)
        fmt = "CSV" if name.lower().endswith(".csv") else "XLSX"
        log_lines.append(
            f"  {name} [{fmt}] | {SETTINGS.hash_algorithm.upper()} : {file_hash}"
        )
    if duplicates_content:
        log_lines += ["", "--- Doublons inter-formats ignorés ---"]
        for dup, orig in duplicates_content:
            log_lines.append(f"  {dup} → même contenu que {orig}")

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
        reconciliation_report = reconcile_data(df_mapped, df)
        report_months      = aggregate_by_month(df)
        report_salespeople = aggregate_by_salesperson(df)

        processed_row_count = len(df)
        period_min   = df["mois"].min()
        period_max   = df["mois"].max()
        total_amount = float(df["montant"].sum())
        salesperson_count = (
            int(df["commercial"].nunique()) if "commercial" in df.columns else 0
        )
        tag = (
            f"{period_min}" if period_min == period_max
            else f"{period_min}_to_{period_max}"
        )

        excel_bytes = build_excel_bytes(df, report_months, report_salespeople)
        excel_bytes = format_excel_bytes(excel_bytes)
        pdf_bytes   = build_pdf_bytes(report_months, report_salespeople)

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

        st.session_state.update({
            "generated": True,
            "status": "ok",
            "error_message": None,
            # Empreinte des fichiers utilisés pour CETTE génération.
            # Comparée à current_upload_sig lors des rechargements suivants.
            "generation_upload_sig": current_upload_sig,
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
                "xlsx_count": xlsx_count,
                "csv_count": csv_count,
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
            "generation_upload_sig": current_upload_sig,
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

s     = st.session_state.summary
q     = st.session_state.quality
recon = st.session_state.reconciliation_report

# --- Résumé ---
st.subheader("✅ Résumé")
st.write(
    f"- Fichiers OK traités : **{s['valid_file_count']}** "
    f"({s['xlsx_count']} Excel, {s['csv_count']} CSV)"
)
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

# --- Réconciliation ---
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