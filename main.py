from pathlib import Path
import logging
import sys

import utils


def main():
    # 1) logging
    log_path = utils.setup_logging(base_dir="output")
    logging.info("=== DÉMARRAGE BATCH REPORTING ===")
    logging.info(f"Log: {log_path}")

    # 2) chemins
    data_dir = Path("data")

    # 3) chargement robuste (gérer data vide / tout KO)
    try:
        df = utils.charger_fichiers_robuste(str(data_dir))
    except ValueError as e:
        logging.warning(f"{e} -> Fin sans traitement.")
        logging.info("=== FIN (AUCUN FICHIER) ===")
        sys.exit(0)  # pas d'échec: juste rien à traiter

    # 4) nettoyage + enrichissement
    df = utils.nettoyer(df)
    df = utils.ajouter_mois(df)

    # 5) reportings
    report_mois = utils.total_par_mois(df)
    report_commercial = utils.total_par_commercial(df)

    # 6) période analysée (basée sur les données)
    min_mois = df["mois"].min()
    max_mois = df["mois"].max()
    year = str(min_mois).split("-")[0]

    # 7) dossier de sortie cohérent avec l'année des données
    out_base = Path("output") / year
    out_base.mkdir(parents=True, exist_ok=True)

    # 8) export Excel + formatage (nom basé sur période)
    if min_mois == max_mois:
        excel_name = f"reporting_{min_mois}.xlsx"
    else:
        excel_name = f"reporting_{min_mois}_to_{max_mois}.xlsx"

    excel_path = out_base / excel_name
    utils.exporter_excel(df, report_mois, report_commercial, excel_path)
    utils.formater_excel(excel_path)

    # 9) export PDF
    utils.exporter_pdf(report_mois, report_commercial, out_base)

    logging.info("=== FIN OK ===")


if __name__ == "__main__":
    main()
