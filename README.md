📊 Reporting Automatisé PME
Outil automatisé de consolidation et reporting Excel pour PME
🎯 Objectif

Cette solution permet de :

Consolider automatiquement plusieurs fichiers Excel (.xlsx)

Nettoyer et normaliser les données

Générer un reporting mensuel consolidé

Produire un rapport PDF synthétique

Assurer une traçabilité complète (journal + empreinte des fichiers)

🔹 MODE 1 — Traitement Automatique (Batch)

Idéal pour un usage simple et rapide.

📂 Déposer les fichiers

Déposez vos exports Excel (.xlsx) dans un dossier à créer :

data/

▶️ Lancer le traitement

Ouvrir un terminal dans le dossier du projet puis exécuter :

py main.py

📁 Résultats générés

Les fichiers sont créés dans :

output/<ANNEE>/


Vous y trouverez :

reporting_YYYY-MM.xlsx → Excel consolidé multi-feuilles

rapport_YYYY-MM.pdf → Rapport synthétique PDF

log_YYYY-MM-DD.txt → Journal détaillé d’exécution

🔹 MODE 2 — Interface Graphique (Streamlit)

Permet :

Import direct des fichiers

Mapping des colonnes (date, montant, commercial)

Visualisation des données

Téléchargement immédiat des exports

▶️ Lancer l’interface

Dans le dossier du projet :

streamlit run app.py


Un navigateur s’ouvre automatiquement.

📤 Étapes

Importer les fichiers Excel

Sélectionner les colonnes nécessaires

Générer le reporting

Télécharger :

Excel consolidé

Rapport PDF

Log d’exécution


🧪 Mode Démonstration (données test)

Pour tester la solution sans utiliser vos données :

Lancer :

py generate_demo_data.py


Cela génère automatiquement plusieurs fichiers Excel de démonstration dans un dossier :

data/


Vous pouvez ensuite :

Lancer le traitement batch (py main.py)

Ou tester l’interface (streamlit run app.py)


🔎 Contrôle Qualité & Traçabilité

Chaque exécution inclut :

Liste des fichiers utilisés

Empreinte SHA256 de chaque fichier

Nombre de lignes traitées

Nombre de lignes supprimées

Période analysée

Journal d’exécution détaillé

Permet :

✔ Audit
✔ Vérification interne
✔ Résolution rapide en cas d’anomalie

⚠️ En cas de problème

Transmettre le fichier :

output/log_YYYY-MM-DD.txt


(ou le log téléchargeable via l’interface)

🛠️ Installation (si nécessaire)

Créer un environnement virtuel :

python -m venv venv
venv\Scripts\activate


Installer les dépendances :

pip install -r requirements.txt

🔒 Sécurité

Traitement local uniquement

Aucun accès réseau requis

Aucune exécution de code externe

Fichiers sources non modifiés