"""
main.py - Point d'entrée pour l'exécution en mode Batch (automatisé).

Usage direct :
    python main.py

Usage planifié (Windows - Planificateur de tâches) :
    Lancer RUN_BATCH.bat selon la fréquence souhaitée.

Ce script est intentionnellement minimal : il orchestre les appels à utils.py
sans contenir de logique métier. Toute la logique est dans utils.py.
Les chemins et constantes proviennent de config.SETTINGS (surchargeable via .env).

Formats sources supportés : .xlsx et .csv (détectés automatiquement).
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import utils
from config import SETTINGS


def main() -> None:
    """Exécute le pipeline complet de reporting en mode Batch.

    Pipeline :
        1. Initialisation du logging (fichier horodaté dans output/).
        2. Chargement robuste des fichiers sources (.xlsx et .csv depuis data/).
        3. Nettoyage et enrichissement (colonne mois).
        4. Checkpoint de réconciliation des données (audit).
        5. Calcul des reportings (par mois, par commercial).
        6. Export Excel multi-feuilles + formatage openpyxl.
        7. Export PDF.

    En cas d'alerte d'intégrité (écart des données anormal), le script
    loggue une erreur critique et quitte avec le code 2 pour permettre
    une détection par un orchestrateur ou le Planificateur de tâches.

    Returns:
        None. Effets de bord : écriture de fichiers dans output/<annee>/.

    Raises:
        SystemExit(0): Si aucun fichier n'est trouvé (pas une erreur).
        SystemExit(2): Si le checkpoint de réconciliation échoue.
    """
    # 1. Logging
    log_path = utils.setup_logging(base_dir=str(SETTINGS.output_dir))
    logging.info(f"=== DÉMARRAGE - {SETTINGS.app_name} v{SETTINGS.app_version} ===")
    logging.info(f"Log : {log_path}")
    logging.info(f"Répertoire source : {SETTINGS.data_dir}")
    logging.info(f"Répertoire sortie : {SETTINGS.output_dir}")
    logging.info(
        f"Formats sources acceptés : {sorted(SETTINGS.csv_source_extensions)}"
    )

    # 2. Chargement robuste (.xlsx + .csv)
    # ValueError si aucun fichier lisible → sortie propre (code 0, pas une erreur fatale).
    try:
        df_source = utils.load_source_files(str(SETTINGS.data_dir))
    except ValueError as e:
        logging.warning(f"{e} → Fin sans traitement.")
        logging.info("=== FIN (AUCUN FICHIER) ===")
        sys.exit(0)

    source_row_count = len(df_source)
    logging.info(f"Lignes chargées (avant nettoyage) : {source_row_count}")

    # 3. Nettoyage + enrichissement
    # clean_data() est format-agnostique : fonctionne identiquement qu'il
    # provienne de fichiers .xlsx, .csv, ou d'un mix des deux.
    df = utils.clean_data(df_source)
    df = utils.add_month_column(df)

    # 4. Checkpoint de réconciliation des données
    recon_report = utils.reconcile_data(df_source, df)

    if not recon_report.integrity_ok:
        logging.critical(
            f"ALERTE INTÉGRITÉ - Écart de {recon_report.absolute_gap:.4f} € "
            f"({recon_report.gap_pct * 100:.2f} %) détecté. "
            "Le reporting NE DOIT PAS être diffusé avant investigation. "
            f"Somme attendue : {recon_report.expected_source_sum:.2f} € | "
            f"Somme traitée : {recon_report.processed_sum:.2f} €"
        )
        logging.info("=== FIN (ÉCHEC INTÉGRITÉ) ===")
        sys.exit(2)

    # 5. Reportings
    report_months = utils.aggregate_by_month(df)
    report_salespeople = utils.aggregate_by_salesperson(df)

    total_amount = float(df["montant"].sum())
    logging.info(f"Chiffre d'affaires total : {total_amount:,.2f} €")
    logging.info(f"Période : {df['mois'].min()} → {df['mois'].max()}")

    # 6. Export Excel
    min_month = df["mois"].min()
    max_month = df["mois"].max()
    year = str(min_month).split("-")[0]
    excel_name = (
        f"reporting_{min_month}.xlsx"
        if min_month == max_month
        else f"reporting_{min_month}_to_{max_month}.xlsx"
    )

    out_dir = SETTINGS.output_dir / year
    out_dir.mkdir(parents=True, exist_ok=True)

    excel_path = out_dir / excel_name
    utils.export_excel(df, report_months, report_salespeople, excel_path)
    utils.format_excel_file(excel_path)

    # 7. Export PDF
    utils.export_pdf(report_months, report_salespeople, out_dir)

    logging.info("=== FIN OK ===")


if __name__ == "__main__":
    main()