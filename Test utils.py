"""
test_utils.py - Suite de tests unitaires pour utils.py.

Philosophie de test :
- Isolation totale : aucun fichier réel sur le disque, aucune dépendance à
  Streamlit. Toutes les données sont générées par des fixtures pytest en mémoire.
- Couverture des cas nominaux ET des cas limites : un test qui ne tente
  jamais de casser le code ne prouve rien.
- Tests d'intégrité des données : la fonction reconcile_data est
  testée pour prouver qu'elle détecte bien toute anomalie silencieuse.

Lancer les tests :
    pytest test_utils.py -v
    pytest test_utils.py -v --tb=short   # traces courtes
    pytest test_utils.py --cov=utils     # couverture (pip install pytest-cov)
"""

from __future__ import annotations

import io
from typing import Any

import pandas as pd
import pytest

from utils import (
    ReconciliationReport,
    add_month_column,
    aggregate_by_month,
    aggregate_by_salesperson,
    check_required_columns,
    clean_data,
    read_uploaded_excels,
    reconcile_data,
    validate_schema,
)
from config import DATA_DICTIONARY, REQUIRED_COLUMNS


# ===========================================================================
# FIXTURES - Données simulées (aucun fichier sur le disque)
# ===========================================================================

@pytest.fixture
def df_clean() -> pd.DataFrame:
    """DataFrame propre : dates valides, montants numériques, commercial présent.

    5 lignes, somme des montants = 1 500.0 €.
    Sert de base pour la majorité des tests nominaux.
    """
    return pd.DataFrame({
        "date": ["01/01/2025", "15/02/2025", "10/03/2025", "20/04/2025", "05/05/2025"],
        "montant": [100.0, 200.0, 300.0, 400.0, 500.0],
        "commercial": ["Alice Martin", "Bob Leroy", "Alice Martin", "Charlie Dupont", "Bob Leroy"],
    })


@pytest.fixture
def df_with_invalid_dates() -> pd.DataFrame:
    """DataFrame avec 2 dates invalides sur 5 lignes.

    Lignes valides = 3 (index 0, 2, 4), somme montants valides = 100 + 300 + 500 = 900.
    Les lignes 1 (32/13/2025) et 3 (nope) doivent être supprimées par clean_data().
    """
    return pd.DataFrame({
        "date": ["01/01/2025", "32/13/2025", "10/03/2025", "nope", "05/05/2025"],
        "montant": [100.0, 200.0, 300.0, 400.0, 500.0],
        "commercial": ["Alice", "Bob", "Alice", "Charlie", "Bob"],
    })


@pytest.fixture
def df_text_amounts() -> pd.DataFrame:
    """DataFrame avec montants non numériques (N/A, texte) et un montant virgule FR.

    Ligne 0 : "N/A"      → NaN (conservé, exclu du CA)
    Ligne 1 : "abc"      → NaN (conservé, exclu du CA)
    Ligne 2 : "300,50"   → 300.50 (format FR, doit être converti correctement)
    """
    return pd.DataFrame({
        "date": ["01/01/2025", "15/02/2025", "10/03/2025"],
        "montant": ["N/A", "abc", "300,50"],
        "commercial": ["Alice", "Bob", "Alice"],
    })


@pytest.fixture
def df_fr_formatted_amounts() -> pd.DataFrame:
    """DataFrame avec montants au format français : espaces milliers + virgule décimale."""
    return pd.DataFrame({
        "date": ["01/01/2025", "15/02/2025"],
        "montant": ["1 234,56", "999,99"],
        "commercial": ["Alice", "Bob"],
    })


@pytest.fixture
def df_spaces_in_columns() -> pd.DataFrame:
    """DataFrame avec espaces parasites dans les noms de colonnes (cas export CRM)."""
    return pd.DataFrame({
        "  Date  ": ["01/01/2025", "15/02/2025"],
        "Montant ": [100.0, 200.0],
        " Commercial": ["Alice", "Bob"],
    })


@pytest.fixture
def df_duplicate_columns() -> pd.DataFrame:
    """DataFrame avec colonnes dupliquées après normalisation.

    Simule un concat de deux exports dont l'un a des espaces dans les noms.
    La fusion bfill doit conserver la première valeur non-nulle.
    """
    df1 = pd.DataFrame({"date": ["01/01/2025"], "montant": [100.0]})
    df2 = pd.DataFrame({"date": ["15/02/2025"], "montant": [200.0]})
    return pd.concat([df1, df2], ignore_index=True)


@pytest.fixture
def df_empty() -> pd.DataFrame:
    """DataFrame vide avec les colonnes requises (aucune ligne)."""
    return pd.DataFrame(columns=["date", "montant", "commercial"])


@pytest.fixture
def df_missing_columns() -> pd.DataFrame:
    """DataFrame sans aucune des colonnes requises."""
    return pd.DataFrame({
        "produit": ["Widget A", "Gadget B"],
        "quantite": [10, 20],
        "prix_unitaire": [9.99, 24.99],
    })


@pytest.fixture
def df_cleaned(df_clean: pd.DataFrame) -> pd.DataFrame:
    """df_clean après passage dans clean_data() - base pour les tests d'agrégation."""
    return clean_data(df_clean)


@pytest.fixture
def df_cleaned_with_month(df_cleaned: pd.DataFrame) -> pd.DataFrame:
    """df_cleaned après add_month_column() - base pour les tests de reporting."""
    return add_month_column(df_cleaned)


def _make_mock_upload(df: pd.DataFrame, name: str = "test.xlsx") -> Any:
    """Crée un faux fichier uploadé Streamlit à partir d'un DataFrame.

    Écrit le DataFrame en mémoire sous forme Excel (.xlsx) et retourne
    un objet BytesIO enrichi d'un attribut ``name`` pour simuler
    ``streamlit.UploadedFile``.

    Args:
        df: DataFrame à sérialiser.
        name: Nom du fichier simulé.

    Returns:
        Objet compatible avec ``read_uploaded_excels``.
    """
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    buf.name = name  # type: ignore[attr-defined]
    return buf


# ===========================================================================
# TESTS - Validation de schéma
# ===========================================================================

class TestValidateSchema:
    """Tests de validate_schema() et check_required_columns()."""

    def test_schema_ok_when_all_required_columns_present(
        self, df_cleaned: pd.DataFrame
    ) -> None:
        """Toutes les colonnes requises présentes → liste vide retournée."""
        missing = validate_schema(df_cleaned)
        assert missing == [], f"Colonnes inattendues manquantes : {missing}"

    def test_schema_reports_missing_required_columns(
        self, df_missing_columns: pd.DataFrame
    ) -> None:
        """Toutes les colonnes requises absentes → liste non vide."""
        missing = validate_schema(df_missing_columns)
        for col in REQUIRED_COLUMNS:
            assert col in missing, f"La colonne requise '{col}' devrait être signalée manquante"

    def test_schema_uses_data_dictionary_as_reference(
        self, df_missing_columns: pd.DataFrame
    ) -> None:
        """validate_schema utilise bien DATA_DICTIONARY comme référence."""
        custom_dict = {
            "col_fictive": type(DATA_DICTIONARY["date"])(
                description="Test", dtype="str", is_required=True, aliases=()
            )
        }
        missing = validate_schema(df_missing_columns, dictionary=custom_dict)
        assert "col_fictive" in missing

    def test_check_required_columns_legacy_list_mode(
        self, df_clean: pd.DataFrame
    ) -> None:
        """check_required_columns avec liste explicite (mode legacy) ne lève pas d'erreur."""
        df_normalized = df_clean.copy()
        df_normalized.columns = df_normalized.columns.str.strip().str.lower()
        check_required_columns(df_normalized, required=["date", "montant"])

    def test_check_required_columns_raises_value_error(
        self, df_missing_columns: pd.DataFrame
    ) -> None:
        """check_required_columns lève ValueError si colonnes requises manquantes."""
        with pytest.raises(ValueError, match="manquantes"):
            check_required_columns(df_missing_columns)


# ===========================================================================
# TESTS - Nettoyage
# ===========================================================================

class TestCleanData:
    """Tests de clean_data()."""

    def test_nominal_returns_dataframe(self, df_clean: pd.DataFrame) -> None:
        """Cas nominal : clean_data() retourne un DataFrame non vide."""
        result = clean_data(df_clean)
        assert isinstance(result, pd.DataFrame)
        assert len(result) == 5

    def test_normalizes_column_names(self, df_spaces_in_columns: pd.DataFrame) -> None:
        """Les espaces et majuscules dans les noms de colonnes sont normalisés."""
        result = clean_data(df_spaces_in_columns)
        assert "date" in result.columns
        assert "montant" in result.columns
        assert "commercial" in result.columns
        for col in result.columns:
            assert col == col.strip(), f"Colonne '{col}' contient des espaces"

    def test_date_column_is_datetime_type(self, df_clean: pd.DataFrame) -> None:
        """La colonne date est bien convertie en datetime64."""
        result = clean_data(df_clean)
        assert pd.api.types.is_datetime64_any_dtype(result["date"])

    def test_amount_column_is_float_type(self, df_clean: pd.DataFrame) -> None:
        """La colonne montant est bien convertie en float64."""
        result = clean_data(df_clean)
        assert pd.api.types.is_float_dtype(result["montant"])

    def test_drops_rows_with_invalid_dates(
        self, df_with_invalid_dates: pd.DataFrame
    ) -> None:
        """Les lignes avec dates invalides sont supprimées."""
        result = clean_data(df_with_invalid_dates)
        # 2 dates invalides sur 5 → 3 lignes restantes
        assert len(result) == 3
        assert result["date"].isna().sum() == 0

    def test_fr_decimal_comma_converted_to_float(
        self, df_text_amounts: pd.DataFrame
    ) -> None:
        """Le montant '300,50' (format FR) est correctement converti en 300.5."""
        result = clean_data(df_text_amounts)
        amount_300 = result.loc[result["montant"].notna(), "montant"].values
        assert 300.5 in amount_300, f"300,50 FR devrait être converti. Valeurs : {amount_300}"

    def test_fr_thousand_spaces_converted(
        self, df_fr_formatted_amounts: pd.DataFrame
    ) -> None:
        """Les montants '1 234,56' et '999,99' (format FR) sont correctement convertis."""
        result = clean_data(df_fr_formatted_amounts)
        assert result["montant"].notna().all()
        assert pytest.approx(result["montant"].iloc[0], abs=0.01) == 1234.56
        assert pytest.approx(result["montant"].iloc[1], abs=0.01) == 999.99

    def test_invalid_amounts_become_nan(self, df_text_amounts: pd.DataFrame) -> None:
        """'N/A' et 'abc' deviennent NaN (ligne conservée, montant exclu du CA)."""
        result = clean_data(df_text_amounts)
        nan_count = result["montant"].isna().sum()
        assert nan_count == 2, f"2 montants invalides attendus, {nan_count} trouvés"

    def test_salesperson_normalized_to_title_case(self, df_clean: pd.DataFrame) -> None:
        """Le champ commercial est normalisé en Title Case."""
        df = df_clean.copy()
        df["commercial"] = [
            "alice martin", "BOB LEROY", "  charlie dupont  ", "alice", "bob"
        ]
        result = clean_data(df)
        assert result["commercial"].iloc[0] == "Alice Martin"
        assert result["commercial"].iloc[1] == "Bob Leroy"
        assert result["commercial"].iloc[2] == "Charlie Dupont"

    def test_fully_empty_rows_are_dropped(self, df_clean: pd.DataFrame) -> None:
        """Les lignes entièrement NaN sont supprimées avant traitement."""
        empty_row = pd.DataFrame([[None, None, None]], columns=df_clean.columns)
        df_with_empty_row = pd.concat([df_clean, empty_row], ignore_index=True)
        result = clean_data(df_with_empty_row)
        assert len(result) == 5

    def test_does_not_mutate_input_dataframe(self, df_clean: pd.DataFrame) -> None:
        """clean_data() ne modifie pas le DataFrame d'entrée (copie interne)."""
        columns_before = list(df_clean.columns)
        clean_data(df_clean)
        assert list(df_clean.columns) == columns_before

    def test_missing_required_columns_raises_value_error(
        self, df_missing_columns: pd.DataFrame
    ) -> None:
        """DataFrame sans colonnes requises → ValueError explicite."""
        with pytest.raises(ValueError, match="manquantes"):
            clean_data(df_missing_columns)

    def test_empty_dataframe_returns_empty_result(self, df_empty: pd.DataFrame) -> None:
        """DataFrame vide (colonnes OK, 0 lignes) → retourne un DataFrame vide."""
        result = clean_data(df_empty)
        assert isinstance(result, pd.DataFrame)
        assert len(result) == 0

    def test_all_invalid_dates_returns_empty_result(self) -> None:
        """Si toutes les dates sont invalides → DataFrame vide (0 lignes)."""
        df = pd.DataFrame({
            "date": ["32/13/2025", "nope", ""],
            "montant": [100.0, 200.0, 300.0],
        })
        result = clean_data(df)
        assert len(result) == 0


# ===========================================================================
# TESTS - Enrichissement
# ===========================================================================

class TestAddMonthColumn:
    """Tests de add_month_column()."""

    def test_month_column_is_added(self, df_cleaned: pd.DataFrame) -> None:
        """La colonne 'mois' est bien créée."""
        result = add_month_column(df_cleaned)
        assert "mois" in result.columns

    def test_month_format_is_yyyy_mm(self, df_cleaned: pd.DataFrame) -> None:
        """Le format de 'mois' est YYYY-MM (tri lexicographique = tri chronologique)."""
        result = add_month_column(df_cleaned)
        for value in result["mois"]:
            assert len(value) == 7, f"Format inattendu : '{value}'"
            assert value[4] == "-", f"Format inattendu : '{value}'"

    def test_month_matches_each_date(self, df_cleaned: pd.DataFrame) -> None:
        """Chaque ligne a le bon mois extrait de sa date."""
        result = add_month_column(df_cleaned)
        for _, row in result.iterrows():
            expected_month = row["date"].strftime("%Y-%m")
            assert row["mois"] == expected_month

    def test_does_not_mutate_input_dataframe(self, df_cleaned: pd.DataFrame) -> None:
        """add_month_column() ne modifie pas le DataFrame d'entrée."""
        columns_before = set(df_cleaned.columns)
        add_month_column(df_cleaned)
        assert set(df_cleaned.columns) == columns_before


# ===========================================================================
# TESTS - Agrégations
# ===========================================================================

class TestAggregateByMonth:
    """Tests de aggregate_by_month()."""

    def test_returns_one_row_per_month(
        self, df_cleaned_with_month: pd.DataFrame
    ) -> None:
        """Un mois distinct → une ligne dans le rapport."""
        result = aggregate_by_month(df_cleaned_with_month)
        distinct_month_count = df_cleaned_with_month["mois"].nunique()
        assert len(result) == distinct_month_count

    def test_sum_is_arithmetically_correct(self) -> None:
        """La somme par mois est arithmétiquement exacte."""
        df = pd.DataFrame({
            "date": pd.to_datetime(["2025-01-01", "2025-01-15", "2025-02-01"]),
            "montant": [100.0, 200.0, 300.0],
            "mois": ["2025-01", "2025-01", "2025-02"],
        })
        result = aggregate_by_month(df)
        jan = result.loc[result["mois"] == "2025-01", "montant"].values[0]
        feb = result.loc[result["mois"] == "2025-02", "montant"].values[0]
        assert pytest.approx(jan, abs=0.01) == 300.0
        assert pytest.approx(feb, abs=0.01) == 300.0

    def test_nan_amounts_excluded_from_sum(self) -> None:
        """Les montants NaN sont exclus du CA (comportement pandas sum par défaut)."""
        df = pd.DataFrame({
            "date": pd.to_datetime(["2025-01-01", "2025-01-15"]),
            "montant": [100.0, float("nan")],
            "mois": ["2025-01", "2025-01"],
        })
        result = aggregate_by_month(df)
        assert pytest.approx(result["montant"].iloc[0], abs=0.01) == 100.0

    def test_sorted_chronologically(self) -> None:
        """Les mois sont triés chronologiquement (YYYY-MM lexicographique)."""
        df = pd.DataFrame({
            "date": pd.to_datetime(["2025-03-01", "2025-01-01", "2025-02-01"]),
            "montant": [300.0, 100.0, 200.0],
            "mois": ["2025-03", "2025-01", "2025-02"],
        })
        result = aggregate_by_month(df)
        assert list(result["mois"]) == ["2025-01", "2025-02", "2025-03"]


class TestAggregateBySalesperson:
    """Tests de aggregate_by_salesperson()."""

    def test_returns_one_row_per_salesperson(
        self, df_cleaned_with_month: pd.DataFrame
    ) -> None:
        """Un commercial distinct → une ligne dans le rapport."""
        result = aggregate_by_salesperson(df_cleaned_with_month)
        salesperson_count = df_cleaned_with_month["commercial"].nunique()
        assert len(result) == salesperson_count

    def test_sum_matches_total(self, df_cleaned_with_month: pd.DataFrame) -> None:
        """La somme globale par commercial = somme globale du DataFrame."""
        result = aggregate_by_salesperson(df_cleaned_with_month)
        assert pytest.approx(result["montant"].sum(), abs=0.01) == (
            df_cleaned_with_month["montant"].sum()
        )

    def test_without_salesperson_column_returns_empty(
        self, df_cleaned: pd.DataFrame
    ) -> None:
        """Sans colonne 'commercial', retourne un DataFrame vide avec les bonnes colonnes."""
        df_without_salesperson = df_cleaned.drop(columns=["commercial"], errors="ignore")
        result = aggregate_by_salesperson(df_without_salesperson)
        assert isinstance(result, pd.DataFrame)
        assert len(result) == 0
        assert "commercial" in result.columns
        assert "montant" in result.columns

    def test_sorted_by_amount_descending(self) -> None:
        """Le rapport est trié par montant décroissant (top performers en premier)."""
        df = pd.DataFrame({
            "date": pd.to_datetime(["2025-01-01"] * 3),
            "montant": [500.0, 100.0, 300.0],
            "mois": ["2025-01"] * 3,
            "commercial": ["Alice", "Bob", "Charlie"],
        })
        result = aggregate_by_salesperson(df)
        amounts = list(result["montant"])
        assert amounts == sorted(amounts, reverse=True)


# ===========================================================================
# TESTS - Réconciliation financière (Checkpoint d'intégrité)
# ===========================================================================

class TestReconcileData:
    """Tests du checkpoint d'intégrité financière.

    Ces tests constituent la preuve formelle que reconcile_data()
    détecte correctement les anomalies et valide les pipelines propres.
    """

    def test_integrity_ok_on_clean_pipeline(self, df_clean: pd.DataFrame) -> None:
        """Pipeline propre (aucune perte anormale) → integrity_ok = True."""
        df_processed = add_month_column(clean_data(df_clean))
        report = reconcile_data(df_clean, df_processed)

        assert report.integrity_ok, (
            f"L'intégrité devrait être OK sur un pipeline propre. "
            f"Message : {report.message}"
        )
        assert report.absolute_gap < 0.01

    def test_returns_reconciliation_report_instance(
        self, df_clean: pd.DataFrame
    ) -> None:
        """reconcile_data() retourne bien un ReconciliationReport."""
        df_processed = add_month_column(clean_data(df_clean))
        report = reconcile_data(df_clean, df_processed)
        assert isinstance(report, ReconciliationReport)

    def test_row_counts_are_correct(self, df_clean: pd.DataFrame) -> None:
        """source_row_count et processed_row_count sont correctement renseignés."""
        df_processed = add_month_column(clean_data(df_clean))
        report = reconcile_data(df_clean, df_processed)

        assert report.source_row_count == 5
        assert report.processed_row_count == 5
        assert report.dropped_row_count == 0

    def test_invalid_dates_counted_correctly(
        self, df_with_invalid_dates: pd.DataFrame
    ) -> None:
        """Les dates invalides sont correctement comptées dans le rapport."""
        df_processed = add_month_column(clean_data(df_with_invalid_dates))
        report = reconcile_data(df_with_invalid_dates, df_processed)

        assert report.invalid_date_count == 2
        assert report.dropped_row_count == 2

    def test_expected_sum_excludes_invalid_date_rows(
        self, df_with_invalid_dates: pd.DataFrame
    ) -> None:
        """La somme source attendue ne compte que les lignes à date valide."""
        df_processed = add_month_column(clean_data(df_with_invalid_dates))
        report = reconcile_data(df_with_invalid_dates, df_processed)

        # Lignes valides : index 0 (100), 2 (300), 4 (500) = 900
        assert pytest.approx(report.expected_source_sum, abs=0.01) == 900.0
        assert pytest.approx(report.processed_sum, abs=0.01) == 900.0
        assert report.integrity_ok

    # --- TESTS D'INTÉGRITÉ NÉGATIFS (détection d'anomalie) ---

    def test_detects_silently_modified_amount(self, df_clean: pd.DataFrame) -> None:
        """🚨 CRITIQUE : détecte une modification silencieuse d'un montant.

        Simule la corruption la plus insidieuse : une valeur modifiée sans suppression
        de ligne. Ce type d'anomalie passe souvent inaperçu sans réconciliation.
        Écart = 100 € sur 1500 € = 6.7 % >> tolérance de 0.1 %.
        """
        df_processed = add_month_column(clean_data(df_clean.copy()))
        # Corruption silencieuse : 100 € → 0 € (perte de 100 € non expliquée)
        df_corrupted = df_processed.copy()
        df_corrupted.loc[0, "montant"] = 0.0

        report = reconcile_data(df_clean, df_corrupted)

        assert not report.integrity_ok, "La réconciliation DOIT détecter cet écart"
        assert pytest.approx(report.absolute_gap, abs=0.01) == 100.0

    def test_detects_unjustified_row_deletion(self, df_clean: pd.DataFrame) -> None:
        """🚨 CRITIQUE : détecte une ligne supprimée de façon non justifiée.

        La date est valide → la suppression ne peut pas être expliquée par le
        filtre des dates invalides. La réconciliation doit flaguer l'anomalie.
        Écart = 500 € sur 1500 € = 33.3 % >> tolérance.
        """
        df_processed = add_month_column(clean_data(df_clean.copy()))
        # Suppression de la dernière ligne (montant 500 €) sans raison valable
        df_incomplete = df_processed.iloc[:-1].copy()

        report = reconcile_data(df_clean, df_incomplete)

        assert not report.integrity_ok, "La réconciliation DOIT détecter la ligne manquante"
        assert report.absolute_gap > 0

    def test_alert_message_contains_alert_marker(self, df_clean: pd.DataFrame) -> None:
        """Le message d'alerte contient '🚨' ou 'ALERTE' pour visibilité dans les logs."""
        df_processed = add_month_column(clean_data(df_clean.copy()))
        df_corrupted = df_processed.copy()
        df_corrupted.loc[0, "montant"] = 0.0

        report = reconcile_data(df_clean, df_corrupted)
        assert "🚨" in report.message or "ALERTE" in report.message

    def test_ok_message_contains_ok_marker(self, df_clean: pd.DataFrame) -> None:
        """Le message OK contient '✅' ou 'OK' pour visibilité dans les logs."""
        df_processed = add_month_column(clean_data(df_clean))
        report = reconcile_data(df_clean, df_processed)
        assert "✅" in report.message or "OK" in report.message

    def test_float_rounding_within_tolerance(self) -> None:
        """Les micro-écarts d'arrondi IEEE 754 ne déclenchent pas d'alerte.

        Ex: sum([0.1, 0.2, 0.3, ...]) peut différer légèrement selon l'ordre.
        La tolérance de 0.1 % absorbe ces imperfections inévitables.
        """
        df_source = pd.DataFrame({
            "date": ["01/01/2025"] * 100,
            "montant": [0.1] * 100,  # 100 × 0.1 = 10.0 (en théorie)
        })
        df_processed = add_month_column(clean_data(df_source.copy()))

        report = reconcile_data(df_source, df_processed)
        assert report.integrity_ok

    def test_empty_source_dataframe_no_exception(self) -> None:
        """Un DataFrame source vide ne doit pas lever d'exception."""
        df_source = pd.DataFrame(columns=["date", "montant"])
        df_processed = pd.DataFrame(columns=["date", "montant", "mois"])

        report = reconcile_data(df_source, df_processed)
        assert report.integrity_ok  # 0 == 0 → pas d'écart
        assert report.source_row_count == 0

    def test_gap_percentage_calculated_correctly(self, df_clean: pd.DataFrame) -> None:
        """L'écart en pourcentage est arithmétiquement correct."""
        df_processed = add_month_column(clean_data(df_clean.copy()))
        df_corrupted = df_processed.copy()
        df_corrupted.loc[0, "montant"] = 0.0  # perte de 100 €

        report = reconcile_data(df_clean, df_corrupted)

        # somme_source_attendue = 1500, ecart = 100, pct = 100/1500 ≈ 0.0667
        assert pytest.approx(report.gap_pct, abs=0.001) == 100.0 / 1500.0


# ===========================================================================
# TESTS - Chargement (read_uploaded_excels)
# ===========================================================================

class TestReadUploadedExcels:
    """Tests de read_uploaded_excels() avec des fichiers simulés en mémoire."""

    def test_reads_single_valid_file(self, df_clean: pd.DataFrame) -> None:
        """Un fichier Excel valide est lu sans erreur."""
        mock_file = _make_mock_upload(df_clean, "test.xlsx")
        result, failed = read_uploaded_excels([mock_file])

        assert isinstance(result, pd.DataFrame)
        assert len(result) == 5
        assert failed == []

    def test_concatenates_multiple_files(self, df_clean: pd.DataFrame) -> None:
        """Deux fichiers Excel valides sont concaténés correctement."""
        f1 = _make_mock_upload(df_clean, "f1.xlsx")
        f2 = _make_mock_upload(df_clean, "f2.xlsx")
        result, failed = read_uploaded_excels([f1, f2])

        assert len(result) == 10  # 5 + 5
        assert failed == []

    def test_corrupted_file_added_to_failed_files(
        self, df_clean: pd.DataFrame
    ) -> None:
        """Un fichier corrompu est ajouté à failed_files sans crasher."""
        valid_file = _make_mock_upload(df_clean, "ok.xlsx")

        # Simule un fichier corrompu (contenu non-Excel)
        corrupted = io.BytesIO(b"ceci n'est pas un Excel")
        corrupted.name = "corrompu.xlsx"  # type: ignore[attr-defined]

        result, failed = read_uploaded_excels([valid_file, corrupted])

        assert len(result) == 5  # seul le fichier OK est lu
        assert len(failed) == 1
        assert failed[0][0] == "corrompu.xlsx"

    def test_all_files_corrupted_raises_value_error(self) -> None:
        """Si tous les fichiers sont illisibles → ValueError."""
        corrupted = io.BytesIO(b"pas un excel")
        corrupted.name = "invalide.xlsx"  # type: ignore[attr-defined]

        with pytest.raises(ValueError, match="lisible"):
            read_uploaded_excels([corrupted])