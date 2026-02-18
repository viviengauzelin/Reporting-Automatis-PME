# 📊 Reporting Automatisé PME

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Interface-Streamlit-red)
![Statut](https://img.shields.io/badge/Statut-Demo%20Professionnelle-success)

Solution d’automatisation de consolidation et de reporting Excel destinée aux PME.

Objectif : transformer des exports Excel bruts en reporting exploitable, propre et traçable, en quelques secondes.

---

# 🎯 Problématique PME

De nombreuses PME :

- Consolident manuellement plusieurs exports Excel
- Refont les mêmes manipulations chaque mois
- Perdent du temps sur le nettoyage des données
- Manquent de traçabilité en cas d’erreur
- N’ont pas d’outil simple pour produire un reporting clair

Cette solution automatise l’ensemble du processus.

---

# ✅ Fonctionnalités

✔ Consolidation automatique de multiples fichiers Excel  
✔ Nettoyage et normalisation des données  
✔ Détection d’erreurs (dates invalides, montants incorrects)  
✔ Reporting mensuel  
✔ Reporting par commercial  
✔ Export Excel multi-feuilles  
✔ Génération PDF  
✔ Log d’exécution détaillé (audit & traçabilité)  
✔ Empreinte SHA256 des fichiers source  

---

# 🚀 Modes de fonctionnement

## 1️⃣ Mode Batch (automatisation locale)

Lecture automatique des fichiers déposés dans (dossier à créer) :

data/


Génération des résultats dans :

output/<ANNEE>/


Fichiers produits :

- reporting_YYYY-MM_to_YYYY-MM.xlsx
- rapport_YYYY-MM_to_YYYY-MM.pdf
- log_YYYY-MM-DD.txt

### Lancer :

```bash
python main.py
Idéal pour :

Exécution planifiée

Traitement mensuel

Intégration dans un flux interne

2️⃣ Interface Web (Streamlit)
Interface utilisateur interactive :

Upload des fichiers Excel

Mapping des colonnes

Contrôle qualité en temps réel

Génération instantanée

Téléchargement Excel / PDF / Log

Lancer :
streamlit run app.py
Idéal pour :

Utilisateur non technique

Traitement ponctuel

Analyse exploratoire


---

🖥 Aperçu de l’interface

1️⃣ Upload des fichiers Excel

![Upload](assets/streamlit_automatisation_demo_1.png)


Interface permettant l’import de plusieurs fichiers `.xlsx` simultanément, avec détection automatique des doublons.


2️⃣ Mapping des colonnes

<img src="assets/streamlit_automatisation_demo_2.png" width="900">

Sélection guidée des colonnes nécessaires (Date, Montant, Commercial) avec validation des incohérences.


3️⃣ Résumé & Reporting

<img src="assets/streamlit_automatisation_demo_3.png" width="900">

Affichage des indicateurs clés


4️⃣ Téléchargement des résultats

<img src="assets/streamlit_automatisation_demo_4.png" width="900">

Export immédiat :

- Excel multi-feuilles formaté  
- Rapport PDF  
- Log d’exécution complet (audit & traçabilité)  


🧪 Données de démonstration
Pour tester le projet :

python generate_demo_data.py
Cela crée automatiquement plusieurs fichiers Excel simulés dans :

data/
🏗 Architecture du projet
project/
│
├── app.py                  # Interface Streamlit
├── main.py                 # Mode batch
├── utils.py                # Fonctions métier (lecture, nettoyage, reporting)
├── generate_demo_data.py   # Génération de données de démo
├── requirements.txt
├── README.md
│
├── data/                   # Fichiers source (non versionnés)
├── output/                 # Résultats générés (non versionnés)
└── venv/                   # Environnement virtuel (non versionné)
⚙ Installation
1️⃣ Créer un environnement virtuel
python -m venv venv
Activation (Windows) :

venv\Scripts\activate
Si PowerShell bloque :

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
2️⃣ Installer les dépendances
pip install -r requirements.txt
🧾 Traçabilité & Audit
Chaque exécution enregistre :

Horodatage précis

Liste des fichiers traités

Hash SHA256 des fichiers

Statistiques de qualité des données

Nombre de lignes supprimées

Résumé financier

Objectif : pouvoir justifier un résultat à tout moment.

🔐 Sécurité & Bonnes pratiques
Aucun code client n’est exécuté

Validation des types et conversions sécurisées

Gestion robuste des erreurs

Données non versionnées

Logs exploitables en cas de contrôle

💼 Cas d’usage
Consolidation mensuelle des ventes

Reporting commercial multi-fichiers

Préparation reporting expert-comptable

Vérification cohérence exports CRM

Analyse interne direction

📈 Valeur ajoutée
Gain estimé :

1 à 3 heures économisées par mois

Réduction du risque d’erreur humaine

Meilleure traçabilité

Standardisation du reporting

🧠 Technologies
Python 3.10+

Pandas

OpenPyXL

ReportLab

Streamlit

Git

👨‍💻 Auteur
Vivien Gauzelin
Ingénieur – Automatisation & Reporting PME

Projet démonstration dans le cadre d’une activité freelance spécialisée en automatisation de processus et reporting.