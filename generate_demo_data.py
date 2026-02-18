from pathlib import Path
import random
import pandas as pd

random.seed(42)

def make_month_df(year, month, n=200):
    commerciaux = ["Alice Martin", "Bob Leroy", "Chloé Bernard", "David Morel"]
    villes = ["Paris", "Lyon", "Marseille", "Toulouse", "Nantes"]
    rows = []

    for i in range(n):
        day = random.randint(1, 28)
        date = f"{day:02d}/{month:02d}/{year}"  # format FR
        montant = round(random.uniform(10, 1500), 2)

        # injecter des montants "sales"
        if random.random() < 0.05:
            montant = "N/A"
        if random.random() < 0.05:
            montant = str(montant).replace(".", ",")  # virgule

        client_id = random.randint(1000, 1300)

        rows.append({
            "Date": date,
            "Montant": montant,
            "Commercial": random.choice(commerciaux),
            "Ville": random.choice(villes),
            "Client": f"CL{client_id}",
            "Commande_ID": f"CMD{year}{month:02d}{i:04d}"
        })

    df = pd.DataFrame(rows)

    # injecter quelques dates cassées (cohérentes avec l'année choisie)
    for _ in range(5):
        idx = random.randint(0, len(df)-1)
        df.loc[idx, "Date"] = f"32/13/{year}"

    # injecter doublons
    dup = df.sample(5, random_state=month)
    df = pd.concat([df, dup], ignore_index=True)

    # colonnes avec espaces (piège classique)
    df.columns = [c + "  " if random.random() < 0.3 else c for c in df.columns]

    return df


def main():
    data_dir = Path("data")
    data_dir.mkdir(exist_ok=True)

    year = 2025  # ← ICI on met 2025

    for month in range(1, 13):
        df = make_month_df(year, month, n=200)
        path = data_dir / f"ventes_{year}-{month:02d}.xlsx"
        df.to_excel(path, index=False)

    print("✅ Données démo 2025 générées dans data/ (12 fichiers).")


if __name__ == "__main__":
    main()
