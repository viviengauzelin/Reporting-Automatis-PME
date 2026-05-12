"""
config.py - Source de vérité unique du projet Reporting demo.

Ce module est le seul endroit où modifier :
- Les chemins de fichiers et constantes d'application (SETTINGS).
- Le schéma métier des colonnes attendues (DATA_DICTIONARY).

Philosophie : « convention over configuration ».
Les valeurs par défaut fonctionnent partout.
Les surcharges (modifications des valeurs par défaut) passent par variables d'environnement ou fichier `.env`.

Usage :
    from config import SETTINGS, DATA_DICTIONARY, REQUIRED_COLUMNS
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

# ---------------------------------------------------------------------------
# Support .env (optionnel - pip install python-dotenv)
# Si python-dotenv n'est pas installé, on continue sans lever d'erreur.
# ---------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv non installé → valeurs par défaut utilisées


# ---------------------------------------------------------------------------
# SETTINGS - Paramètres de l'application
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class AppSettings:
    """Paramètres globaux de l'application.

    ``frozen=True`` garantit l'immuabilité après instanciation :
    le code ne peut pas modifier accidentellement un chemin en cours d'exécution.
    C'est essentiel pour la traçabilité en contexte d'audit.

    Attributes:
        data_dir: Répertoire des fichiers sources Excel/CSV (mode Batch).
        output_dir: Répertoire des fichiers générés (Excel, PDF, logs).
        app_name: Nom de l'application (affiché dans les rapports).
        app_version: Version sémantique (tracée dans les logs).
        log_level: Niveau de log Python (INFO, DEBUG, WARNING…).
        excel_max_col_width: Largeur maximale d'une colonne Excel (en caractères).
        excel_min_col_width: Largeur minimale d'une colonne Excel.
        excel_money_min_width: Largeur minimale pour les colonnes monétaires.
        pdf_pagesize: Format de page PDF (A4 ou LETTER).
        hash_algorithm: Algorithme de hachage pour l'empreinte des fichiers sources.
        reconciliation_tolerance_pct: Tolérance des données pour la réconciliation.
            Exemple : 0.001 = 0.1 % d'écart autorisé (arrondis flottants IEEE 754).
            Au-delà → alerte d'intégrité levée.
        csv_encodings_fallback: Séquence d'encodages testés lors de la lecture d'un CSV.
            Justification : les exports Windows (Excel, CRM, ERP) produisent souvent
            du latin-1 ou cp1252, tandis que les exports modernes utilisent utf-8
            avec ou sans BOM (utf-8-sig). L'ordre reflète la fréquence rencontrée
            dans les PME françaises.
        csv_source_extensions: Extensions de fichiers sources acceptées (mode Batch).
            Permet d'étendre facilement le support à d'autres formats sans
            modifier la logique de chargement.
    """

    # --- Chemins (surchargeables via .env) ---
    data_dir: Path = field(
        default_factory=lambda: Path(os.getenv("DATA_DIR", "data"))
    )
    output_dir: Path = field(
        default_factory=lambda: Path(os.getenv("OUTPUT_DIR", "output"))
    )

    # --- Application ---
    app_name: str = "Reporting démo"
    app_version: str = "2.1.0"
    log_level: str = field(default_factory=lambda: os.getenv("LOG_LEVEL", "INFO"))

    # --- Mise en forme Excel ---
    excel_max_col_width: int = 40
    excel_min_col_width: int = 12
    excel_money_min_width: int = 18  # colonnes monétaires plus larges par défaut

    # --- PDF ---
    pdf_pagesize: str = "A4"

    # --- Audit ---
    hash_algorithm: str = "sha256"

    # --- Réconciliation des données ---
    # Justification métier : 0.1 % couvre les écarts d'arrondi IEEE 754 inhérents
    # aux flottants Python (ex: 0.1 + 0.2 ≠ 0.3 exactement).
    # Toute valeur supérieure indique une anomalie réelle (ligne perdue, valeur modifiée).
    reconciliation_tolerance_pct: float = field(
        default_factory=lambda: float(os.getenv("RECON_TOLERANCE_PCT", "0.001"))
    )

    # --- Lecture CSV ---
    # Ordre des encodages testés : du plus courant en contexte français au plus rare.
    # utf-8-sig : utf-8 avec BOM (exports Excel "Enregistrer sous CSV UTF-8")
    # utf-8     : exports modernes (Python, API, Linux)
    # latin-1   : anciens exports ERP/CRM Windows (ISO-8859-1)
    # cp1252    : Windows Western European (surensemble de latin-1, caractères €, etc.)
    csv_encodings_fallback: tuple[str, ...] = (
        "utf-8-sig", "utf-8", "latin-1", "cp1252"
    )

    # Extensions acceptées en mode Batch (glob patterns sans le point).
    # L'ordre détermine la priorité d'affichage dans les logs, pas le traitement.
    csv_source_extensions: tuple[str, ...] = ("xlsx", "csv")


# Singleton global - importé par utils.py et app.py
SETTINGS: AppSettings = AppSettings()


# ---------------------------------------------------------------------------
# DATA DICTIONARY - Schéma métier des colonnes
# ---------------------------------------------------------------------------

# Type alias pour les dtypes pandas acceptés après nettoyage.
DtypeLiteral = Literal["datetime64[ns]", "float64", "str", "int64", "object"]


@dataclass(frozen=True)
class ColSpec:
    """Spécification complète d'une colonne du schéma de données.

    Attributes:
        description: Description métier (pour la documentation et les messages d'erreur).
        dtype: Type pandas attendu après nettoyage.
        is_required: Si ``True``, l'absence de cette colonne bloque le traitement.
        aliases: Noms alternatifs reconnus dans les fichiers sources.
    """

    description: str
    dtype: DtypeLiteral
    is_required: bool
    aliases: tuple[str, ...] = field(default_factory=tuple)


# Schéma canonique - clé = nom de colonne interne normalisé (après strip + lower).
# Les clés sont des noms de colonnes métier français, conformes aux fichiers sources Excel/CSV.
DATA_DICTIONARY: dict[str, ColSpec] = {
    "date": ColSpec(
        description=(
            "Date de la transaction commerciale. "
            "Formats acceptés dans les sources : JJ/MM/AAAA, YYYY-MM-DD, etc. "
            "Les lignes sans date valide sont supprimées (règle d'audit : "
            "une transaction non datée ne peut pas être comptabilisée)."
        ),
        dtype="datetime64[ns]",
        is_required=True,
        aliases=("date_vente", "date_commande", "date_facture", "transaction_date"),
    ),
    "montant": ColSpec(
        description=(
            "Montant financier de la transaction en euros (HT ou TTC selon source). "
            "Accepte les formats FR (virgule décimale, espace milliers) et EN (point décimal). "
            "Les montants non convertibles deviennent NaN et sont exclus du CA "
            "(tracés dans le log d'exécution pour audit)."
        ),
        dtype="float64",
        is_required=True,
        aliases=("montant_ht", "montant_ttc", "ca", "chiffre_affaires", "amount", "total"),
    ),
    "commercial": ColSpec(
        description=(
            "Nom du commercial responsable de la vente. "
            "Normalisé en Title Case (ex: 'alice martin' → 'Alice Martin'). "
            "Optionnel : si absent, les rapports par commercial ne sont pas générés."
        ),
        dtype="str",
        is_required=False,
        aliases=("vendeur", "representant", "account_manager", "salesperson"),
    ),
    "ville": ColSpec(
        description="Ville du client. Conservé pour analyses géographiques futures.",
        dtype="str",
        is_required=False,
        aliases=("city", "localite", "agence"),
    ),
    "client": ColSpec(
        description="Identifiant ou nom du client. Clé de jointure potentielle avec un CRM.",
        dtype="str",
        is_required=False,
        aliases=("client_id", "customer", "compte", "customer_id"),
    ),
    "commande_id": ColSpec(
        description=(
            "Identifiant unique de la commande/facture. "
            "Permet la déduplication en cas de double export."
        ),
        dtype="str",
        is_required=False,
        aliases=("order_id", "ref_commande", "num_commande", "invoice_id"),
    ),
}

# Colonnes obligatoires - dérivées dynamiquement du dictionnaire.
# Avantage : ajouter une colonne required dans DATA_DICTIONARY la propage automatiquement
# dans toute la chaîne de validation sans toucher à d'autres fichiers.
REQUIRED_COLUMNS: list[str] = [
    name for name, spec in DATA_DICTIONARY.items() if spec.is_required
]