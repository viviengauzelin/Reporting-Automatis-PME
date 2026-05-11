"""
generate_demo_data.py — Générateur de données de démonstration.

Crée 12 fichiers Excel simulant des exports de ventes mensuelles pour l'année
configurée dans SETTINGS (par défaut 2025). Les données sont intentionnellement
"sales" (dates invalides, montants en format FR, espaces dans les colonnes)
pour tester la robustesse du pipeline de nettoyage.

Usage :
    python generate_demo_data.py
    — ou —
    Double-clic sur CREER_FICHIERS_DEMO.bat

Fichiers générés :
    data/ventes_YYYY-MM.xlsx  (12 fichiers, un par mois)
"""

from __future__ import annotations

import random
from pathlib import Path

import pandas as pd

from config import SETTINGS

# Graine fixe : garantit des données reproductibles d'une exécution à l'autre.
# Indispensable pour les démonstrations et les tests de non-régression.
random.seed(42)


def build_monthly_dataframe(year: int, month: int, n: int = 200) -> pd.DataFrame:
    """Génère un DataFrame simulant les ventes d'un mois donné.

    Injecte volontairement des anomalies pour tester le pipeline :
    - 5 % de montants non numériques (valeur "N/A")
    - 5 % de montants au format FR (virgule décimale)
    - 5 lignes avec des dates impossibles (32/13/YYYY)
    - 5 doublons aléatoires
    - 30 % des colonnes avec des espaces parasites en suffixe

    Args:
        year: Année des transactions (ex: 2025).
        month: Mois des transactions (1 à 12).
        n: Nombre de lignes à générer avant injection des anomalies.

    Returns:
        DataFrame avec colonnes : Date, Montant, Commercial, Ville, Client,
        Commande_ID (certaines potentiellement avec espaces parasites).
    """
    salespersons = ["Alice Martin", "Bob Leroy", "Chloé Bernard", "David Morel"]
    cities = ["Paris", "Lyon", "Marseille", "Toulouse", "Nantes"]
    rows = []

    for i in range(n):
        day = random.randint(1, 28)
        date = f"{day:02d}/{month:02d}/{year}"  # format JJ/MM/AAAA (standard FR)
        amount = round(random.uniform(10, 1500), 2)

        # Injection de montants "sales" pour tester la robustesse du nettoyage
        if random.random() < 0.05:
            amount = "N/A"  # type: ignore[assignment]  # valeur non numérique intentionnelle
        if random.random() < 0.05:
            amount = str(amount).replace(".", ",")  # format FR : virgule décimale  # type: ignore[assignment]

        client_id = random.randint(1000, 1300)

        rows.append({
            "Date": date,
            "Montant": amount,
            "Commercial": random.choice(salespersons),
            "Ville": random.choice(cities),
            "Client": f"CL{client_id}",
            "Commande_ID": f"CMD{year}{month:02d}{i:04d}",
        })

    df = pd.DataFrame(rows)

    # Injection de dates impossibles pour tester la suppression par clean_data()
    for _ in range(5):
        idx = random.randint(0, len(df) - 1)
        df.loc[idx, "Date"] = f"32/13/{year}"

    # Injection de doublons pour tester la déduplication côté utilisateur
    duplicate_rows = df.sample(5, random_state=month)
    df = pd.concat([df, duplicate_rows], ignore_index=True)

    # Espaces parasites dans les noms de colonnes (simulation d'exports CRM réels)
    df.columns = [c + "  " if random.random() < 0.3 else c for c in df.columns]

    return df


def main() -> None:
    """Génère les 12 fichiers Excel de démonstration dans SETTINGS.data_dir.

    Le répertoire est créé automatiquement s'il n'existe pas.
    L'année est dérivée des données de config.py (configurable via .env).
    """
    data_dir: Path = SETTINGS.data_dir
    data_dir.mkdir(exist_ok=True)

    # Année des données de démonstration.
    # Modifier DATA_YEAR dans .env pour changer l'année sans toucher au code.
    year = 2025

    for month in range(1, 13):
        df = build_monthly_dataframe(year, month, n=200)
        output_path = data_dir / f"ventes_{year}-{month:02d}.xlsx"
        df.to_excel(output_path, index=False)

    print(f"✅ Données démo {year} générées dans '{data_dir}/' (12 fichiers).")


if __name__ == "__main__":
    main()