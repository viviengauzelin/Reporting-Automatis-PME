import logging
from pathlib import Path
from datetime import datetime

import pandas as pd


# -----------------------------
# LOGGING (Batch — exécution planifiée)
# -----------------------------
def setup_logging(base_dir="output"):
    Path(base_dir).mkdir(exist_ok=True)
    today = datetime.today().strftime("%Y-%m-%d")
    log_path = Path(base_dir) / f"log_{today}.txt"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )
    return log_path


# -----------------------------
# UTILS
# -----------------------------
def check_columns(df: pd.DataFrame, required: list[str]):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes: {missing}")


# -----------------------------
# CHARGEMENT
# -----------------------------
def charger_fichiers_robuste(folder_path: str) -> pd.DataFrame:
    """
    Batch: lit tous les .xlsx du dossier data/.
    Continue même si un fichier est cassé (illisible/corrompu/protégé).
    """
    import hashlib

    def file_sha256(path: Path) -> str:
        with open(path, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()

    files = list(Path(folder_path).glob("*.xlsx"))
    logging.info(f"{len(files)} fichier(s) détecté(s) dans {folder_path}")

    dfs = []
    ko = 0

    for file in files:
        # Parfois Excel laisse des fichiers temporaires ~$ventes_...xlsx et pd.read_excel peut planter
        if file.name.startswith("~$"):
            continue
        try:
            h = file_sha256(file)
            logging.info(f"Fichier: {file.name} | SHA256: {h}")

            df = pd.read_excel(file)
            dfs.append(df)
            logging.info(f"OK lecture: {file.name} ({len(df)} lignes)")
        except Exception as e:
            ko += 1
            logging.error(f"KO lecture: {file.name} | {e}")

    if not dfs:
        raise ValueError("Aucun fichier lisible. Vérifie le dossier data/")

    if ko:
        logging.warning(f"Fichiers en erreur: {ko}")

    return pd.concat(dfs, ignore_index=True)



def lire_excels_upload(uploaded_files):
    """
    Streamlit: lit des fichiers uploadés (.xlsx) et concatène.
    Continue même si un fichier est cassé (illisible/corrompu/protégé).
    Retourne (df_concat, fichiers_ko).
    """
    dfs = []
    fichiers_ko = []

    for f in uploaded_files:
        try:
            df = pd.read_excel(f, engine="openpyxl")  # engine explicite
            dfs.append(df)
        except Exception as e:
            name = getattr(f, "name", "fichier_sans_nom.xlsx")
            fichiers_ko.append((name, str(e)))

    if not dfs:
        raise ValueError("Aucun fichier Excel lisible parmi ceux uploadés.")

    return pd.concat(dfs, ignore_index=True), fichiers_ko



# -----------------------------
# NETTOYAGE
# -----------------------------
def nettoyer(df: pd.DataFrame) -> pd.DataFrame:
    # Normaliser les noms de colonnes (gestion espaces / casse)
    df.columns = df.columns.str.strip().str.lower()

    # Fusionner les colonnes  quées au lieu d’en supprimer une (ex: "montant" + "montant")
    if df.columns.duplicated().any():
        dup_names = df.columns[df.columns.duplicated()].unique().tolist()
        logging.warning(
            f"Colonnes dupliquées après normalisation: {dup_names} (fusion des valeurs)"
        )

        for name in dup_names:
            cols = df.loc[:, df.columns == name]  # toutes les colonnes portant ce nom
            merged = cols.bfill(axis=1).iloc[:, 0]  # première valeur non vide ligne par ligne
            df = df.drop(columns=name)
            df[name] = merged

    df = df.dropna(how="all")

    check_columns(df, ["date", "montant"])

    # --- Montant : garder une copie brute pour debug ---
    df["montant_raw"] = df["montant"]

    # --- Nettoyage texte ---
    s = df["montant"].astype(str).str.strip()
    s = s.str.replace(",", ".", regex=False)
    s = s.str.replace("\u00a0", "", regex=False)  # espace insécable parfois

    # --- Conversion numérique ---
    df["montant"] = pd.to_numeric(s, errors="coerce")

    # --- Stats / logs ---
    raw = df["montant_raw"]

    # Compter les montants vides dans la source
    nb_vides = int(raw.isna().sum() + (raw.astype(str).str.strip() == "").sum())
    logging.info(f"Montants vides (source): {nb_vides}")

    # Montants manquants après conversion + exemples (non vides)
    nb_nan = int(df["montant"].isna().sum())
    if nb_nan:
        logging.warning(f"montants manquants (vides) -> NaN: {nb_nan}")

        mask_examples = df["montant"].isna() & raw.notna() & (raw.astype(str).str.strip() != "")
        exemples = raw.loc[mask_examples].head(10).tolist()
        logging.warning(f"Exemples montants manquants (non vides): {exemples}")

    # Retirer la colonne debug dans l'Excel final
    df = df.drop(columns=["montant_raw"])

    # --- Dates ---
    df["date"] = pd.to_datetime(df["date"], dayfirst=True, errors="coerce")
    nb_date_nat = int(df["date"].isna().sum())
    if nb_date_nat:
        logging.warning(f"dates invalides -> NaT: {nb_date_nat}")

    before = len(df)
    df = df.dropna(subset=["date"])
    logging.info(f"Lignes supprimées sans date: {before - len(df)}")

    # --- Normalisation texte ---
    if "commercial" in df.columns:
        df["commercial"] = df["commercial"].astype(str).str.strip().str.title()

    return df


def ajouter_mois(df: pd.DataFrame) -> pd.DataFrame:
    df["mois"] = df["date"].dt.to_period("M").astype(str)
    return df


# -----------------------------
# REPORTINGS
# -----------------------------
def total_par_mois(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("mois")["montant"].sum().reset_index()


def total_par_commercial(df: pd.DataFrame) -> pd.DataFrame:
    if "commercial" not in df.columns:
        return pd.DataFrame(columns=["commercial", "montant"])
    return df.groupby("commercial")["montant"].sum().reset_index()


# -----------------------------
# EXPORT EXCEL (Batch)
# -----------------------------
def exporter_excel(
    df: pd.DataFrame,
    report_mois: pd.DataFrame,
    report_commercial: pd.DataFrame,
    filepath: Path,
) -> Path:
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)

    try:
        with pd.ExcelWriter(filepath) as writer:
            df.to_excel(writer, sheet_name="Donnees_brutes", index=False)

            # Renommage uniquement pour l'export (on garde "montant" en interne)
            report_mois_export = report_mois.copy()
            report_commercial_export = report_commercial.copy()

            if "montant" in report_mois_export.columns:
                report_mois_export = report_mois_export.rename(columns={"montant": "montant en euros"})

            if "montant" in report_commercial_export.columns:
                report_commercial_export = report_commercial_export.rename(columns={"montant": "montant en euros"})

            report_mois_export.to_excel(writer, sheet_name="Total_par_mois", index=False)
            report_commercial_export.to_excel(writer, sheet_name="Total_par_commercial", index=False)

        logging.info(f"Excel exporté: {filepath}")
        return filepath
    except Exception as e:
        logging.error(f"Erreur export Excel: {e}")
        raise


# -----------------------------
# FORMAT EXCEL (openpyxl) — Batch
# -----------------------------
def formater_excel(filepath: Path) -> Path:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, numbers

    wb = load_workbook(filepath)

    for ws in wb.worksheets:
        # Header en gras
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Ligne 1 figée
        ws.freeze_panes = "A2"

        # Ajustement largeur colonnes
        for col_cells in ws.columns:
            col_letter = col_cells[0].column_letter
            header = str(col_cells[0].value) if col_cells[0].value else ""

            max_len = len(header)
            for cell in col_cells[1:]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            width = max_len + 4
            width = max(width, 12)

            if "montant" in header.lower():
                width = max(width, 18)

            width = min(width, 40)
            ws.column_dimensions[col_letter].width = width

        # Format dates / monnaie
        headers = [cell.value for cell in ws[1]]
        header_to_index = {h: i + 1 for i, h in enumerate(headers) if h}

        if "date" in header_to_index:
            c = header_to_index["date"]
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = "DD/MM/YYYY"

        for money_col in ("montant", "montant en euros", "total"):
            if money_col in header_to_index:
                c = header_to_index[money_col]
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, c).number_format = numbers.FORMAT_CURRENCY_EUR_SIMPLE

    wb.save(filepath)
    logging.info(f"Excel formaté: {filepath}")
    return filepath


# -----------------------------
# PDF (Batch) + PDF bytes (Streamlit)
# -----------------------------
def exporter_pdf(report_mois: pd.DataFrame, report_commercial: pd.DataFrame, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)

    # Période réelle couverte par les données (basée sur le reporting par mois)
    min_mois = report_mois["mois"].min()
    max_mois = report_mois["mois"].max()

    if min_mois == max_mois:
        filename = f"rapport_{min_mois}.pdf"
    else:
        filename = f"rapport_{min_mois}_to_{max_mois}.pdf"

    pdf_path = out_dir / filename
    _build_pdf_to_path(report_mois, report_commercial, pdf_path)

    logging.info(f"PDF généré: {pdf_path}")
    return pdf_path


def build_pdf_bytes(report_mois: pd.DataFrame, report_commercial: pd.DataFrame) -> bytes:
    import io

    buffer = io.BytesIO()
    _build_pdf_to_buffer(report_mois, report_commercial, buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _build_pdf_to_path(report_mois: pd.DataFrame, report_commercial: pd.DataFrame, pdf_path: Path):
    from reportlab.platypus import SimpleDocTemplate
    from reportlab.lib.pagesizes import A4

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4)
    elements = _pdf_elements(report_mois, report_commercial)
    doc.build(elements)


def _build_pdf_to_buffer(report_mois: pd.DataFrame, report_commercial: pd.DataFrame, buffer):
    from reportlab.platypus import SimpleDocTemplate
    from reportlab.lib.pagesizes import A4

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = _pdf_elements(report_mois, report_commercial)
    doc.build(elements)


def _pdf_elements(report_mois: pd.DataFrame, report_commercial: pd.DataFrame):
    from reportlab.platypus import Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors

    styles = getSampleStyleSheet()
    elements = []

    # Déterminer l'année à partir des données
    min_mois = report_mois["mois"].min()
    year = min_mois.split("-")[0]

    elements.append(Paragraph(f"Rapport annuel — {year}", styles["Heading1"]))
    elements.append(Paragraph(f"Généré le {datetime.today().strftime('%Y-%m-%d')}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Copies avant arrondis (évite les effets de bord)
    report_mois = report_mois.copy() if report_mois is not None else report_mois
    report_commercial = report_commercial.copy() if report_commercial is not None else report_commercial

    # Arrondir montants + renommer la colonne pour l'affichage
    for d in (report_mois, report_commercial):
        if d is not None and "montant" in d.columns:
            d["montant"] = pd.to_numeric(d["montant"], errors="coerce").round(2)
            d.rename(columns={"montant": "montant en euros"}, inplace=True)

    def add_table(df, title):
        elements.append(Paragraph(title, styles["Heading2"]))
        elements.append(Spacer(1, 6))

        if df is None or df.empty:
            elements.append(Paragraph("Aucune donnée.", styles["Normal"]))
            elements.append(Spacer(1, 12))
            return

        data = [list(df.columns)] + df.values.tolist()
        table = Table(data, hAlign="LEFT")
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(table)
        elements.append(Spacer(1, 12))

    add_table(report_mois, "Total par mois")
    add_table(report_commercial, "Total par commercial")

    return elements


# -----------------------------
# STREAMLIT: Excel bytes (multi-feuilles)
# -----------------------------
def build_excel_bytes(df: pd.DataFrame, report_mois: pd.DataFrame, report_commercial: pd.DataFrame) -> bytes:
    import io

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Donnees_brutes", index=False)

        # Renommage uniquement pour l'export (download Streamlit)
        report_mois_export = report_mois.copy()
        report_commercial_export = report_commercial.copy()

        if "montant" in report_mois_export.columns:
            report_mois_export = report_mois_export.rename(columns={"montant": "montant en euros"})

        if "montant" in report_commercial_export.columns:
            report_commercial_export = report_commercial_export.rename(columns={"montant": "montant en euros"})

        report_mois_export.to_excel(writer, sheet_name="Total_par_mois", index=False)
        report_commercial_export.to_excel(writer, sheet_name="Total_par_commercial", index=False)

    buffer.seek(0)
    return buffer.getvalue()
