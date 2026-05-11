# 📊 Reporting Automatisé PME

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Interface-Streamlit-red)
![Pytest](https://img.shields.io/badge/Tests-Pytest-green)
![Statut](https://img.shields.io/badge/Statut-Demo%20Professionnelle-success)

Outil automatisé de consolidation et reporting Excel pour PME.

Objectif : transformer des exports Excel bruts en reporting exploitable, propre et traçable, en quelques secondes.

---

# 🎯 Problématique PME

De nombreuses PME :

- Consolident manuellement plusieurs exports Excel
- Refont les mêmes manipulations chaque mois
- Perdent du temps sur le nettoyage des données
- Manquent de traçabilité en cas d'erreur
- N'ont pas d'outil simple pour produire un reporting clair

Cette solution automatise l'ensemble du processus.

---

# ✅ Fonctionnalités

✔ Consolidation automatique de multiples fichiers Excel  
✔ Nettoyage et normalisation des données  
✔ Détection d'erreurs (dates invalides, montants incorrects)  
✔ **Réconciliation des données** : checkpoint d'intégrité prouvant mathématiquement qu'aucune valeur n'a été perdue  
✔ Reporting mensuel  
✔ Reporting par commercial  
✔ Export Excel multi-feuilles  
✔ Génération PDF  
✔ Log d'exécution détaillé (audit & traçabilité)  
✔ Empreinte SHA-256 des fichiers source  
✔ Suite de tests unitaires (47 tests, couverture Pytest)  
✔ Configuration centralisée via `.env`  

⚠️ Remarque : Cette démo est conçue pour fonctionner avec des fichiers Excel respectant le format généré par le script de démonstration. Les fichiers doivent avoir les colonnes attendues (date, montant, commercial, etc.) et un format compatible `.xlsx`.

---

# 📦 Prérequis d'installation

- Ordinateur Windows (code adaptable à MacOS et Linux avec quelques changements)
- Python 3.10+ installé
- Option "Add Python to PATH" cochée
- Droits d'installation initiaux (ou intervention IT)
- Connexion Internet uniquement lors de la première installation

---

# 🧪 Données de démonstration

Pour tester le projet :

```bash
python generate_demo_data.py
```

Cela crée automatiquement 12 fichiers Excel simulés (un par mois) dans :

```
data/
```

---

# 🚀 Modes de fonctionnement

## 1️⃣ Mode Batch (automatisation locale)

Lecture automatique des fichiers déposés dans :

```
data/
```

Génération des résultats dans :

```
output/<ANNEE>/
```

Fichiers produits :

- `reporting_YYYY-MM_to_YYYY-MM.xlsx`
- `reporting_YYYY-MM_to_YYYY-MM.pdf`
- `log_YYYY-MM-DD.txt`

Le mode Batch intègre un **checkpoint d'intégrité des données** : si un écart anormal est détecté entre la somme source et la somme traitée, le script se termine avec le code de sortie `2` (détectable par un planificateur ou un système de supervision).

### Lancer :

```bash
python main.py
```

Idéal pour :

- Exécution planifiée via le Planificateur de tâches Windows
- Traitement mensuel automatisé
- Intégration dans un flux interne

---

## 2️⃣ Interface Web (Streamlit)

Interface utilisateur interactive :

- Upload des fichiers Excel
- Mapping des colonnes
- Contrôle qualité en temps réel
- Réconciliation des données avec verdict d'intégrité
- Génération instantanée
- Téléchargement Excel / PDF / Log

Lancer :

```bash
streamlit run app.py
```

Idéal pour :

- Utilisateur non technique
- Traitement ponctuel
- Analyse exploratoire

---

# 🖥 Aperçu de l'interface

## 1️⃣ Upload des fichiers Excel

<img src="assets/streamlit_automatisation_demo_1.png" width="900">

Interface permettant l'import de plusieurs fichiers `.xlsx` simultanément, avec détection automatique des doublons.

---

## 2️⃣ Mapping des colonnes

<img src="assets/streamlit_automatisation_demo_2.png" width="900">

Sélection guidée des colonnes nécessaires (Date, Montant, Commercial) avec validation des incohérences.

---

## 3️⃣ Résumé & Reporting

<img src="assets/streamlit_automatisation_demo_3.png" width="900">
---
<img src="assets/streamlit_automatisation_demo_4.png" width="900">

Affichage des indicateurs clés :

- Nombre de fichiers traités
- Lignes avant/après nettoyage
- Qualité des données
- **Verdict de réconciliation des données** (✅ OK ou 🚨 Alerte)
- Chiffre d'affaires total
- Reporting mensuel et par commercial

---

## 4️⃣ Téléchargement des résultats

<img src="assets/streamlit_automatisation_demo_5.png" width="900">

Export immédiat :

- Excel multi-feuilles formaté
- Rapport PDF
- Log d'exécution complet (audit & traçabilité)

---

# 🏗 Architecture du projet

```
project/
│
├── app.py                  # Interface Streamlit (présentation)
├── main.py                 # Point d'entrée mode Batch
├── utils.py                # Moteur : chargement, nettoyage, exports, audit
├── config.py               # Configuration centralisée (SETTINGS, DATA_DICTIONARY)
├── generate_demo_data.py   # Générateur de données de test
├── test_utils.py           # Suite de tests unitaires (pytest, 47 tests)
│
├── requirements.txt        # Dépendances Python
├── .env.template           # Modèle de configuration (à copier en .env)
├── README.md
│
├── INSTALLER.bat           # Installation en un clic (Windows)
├── RUN_STREAMLIT.bat       # Lancement interface Web
├── RUN_BATCH.bat           # Lancement mode automatisé
├── CREER_FICHIERS_DEMO.bat # Génération des données de test
│
├── data/                   # Fichiers source (non versionnés)
├── output/                 # Résultats générés (non versionnés)
└── venv/                   # Environnement virtuel (non versionné)
```

---

# ⚙ Installation

## Installation manuelle

**1️⃣ Créer un environnement virtuel**

```bash
python -m venv venv
```

Activation (Windows) :

```bash
venv\Scripts\activate
```

Si PowerShell bloque :

```bash
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**2️⃣ Installer les dépendances**

```bash
pip install -r requirements.txt
```

**3️⃣ Configurer l'environnement (optionnel)**

```bash
cp .env.template .env
# Éditer .env selon ton environnement (chemins, niveau de log, etc.)
```

---

## ⚙️ Installation & Lancement simplifiés (.bat)

Pour une utilisation simple côté PME, l'outil peut être installé et lancé sans terminal.

### 🔹 1. Installation (une seule fois)

Double-cliquer sur `INSTALLER.bat`.

Ce script :
- Crée un environnement virtuel (venv)
- Installe automatiquement les dépendances
- Prépare l'environnement d'exécution

---

### 🔹 2. Création des fichiers Excel de démo

Double-cliquer sur `CREER_FICHIERS_DEMO.bat`.

Ce script :
- Crée le dossier `data/` avec 12 fichiers Excel de démonstration

---

### 🔹 3a. Lancement – Mode Interface (recommandé)

Double-cliquer sur `RUN_STREAMLIT.bat`.

Cela :
- Lance l'application Streamlit
- Ouvre automatiquement le navigateur en local (localhost)

---

### 🔹 3b. Lancement – Mode Batch (automatique)

Double-cliquer sur `RUN_BATCH.bat`.

Le script :
- Traite automatiquement tous les fichiers `.xlsx` du dossier `data/`
- Génère les reportings dans `output/`
- Affiche une alerte si le checkpoint d'intégrité des données détecte un écart

⚠️ En environnement planifié (Planificateur de tâches Windows), `RUN_BATCH.bat` peut être exécuté automatiquement. Le code de sortie `2` signale une alerte d'intégrité détectable par le planificateur.

---

# 🧪 Lancer les tests

```bash
pytest test_utils.py -v
```

Avec rapport de couverture :

```bash
pytest test_utils.py --cov=utils --cov-report=term-missing
```

La suite couvre 47 cas : nominal, colonnes manquantes, données vides, formats invalides, et détection d'anomalie dans les valeurs des données par la réconciliation.

---

# 🧾 Traçabilité & Audit

Chaque exécution enregistre :

- Horodatage précis
- Liste des fichiers traités
- Empreinte SHA-256 des fichiers source
- Statistiques de qualité des données
- Nombre de lignes supprimées et raison
- **Résultat de la réconciliation des données** (somme source vs somme traitée, écart en € et %)
- Résumé du chiffre d'affaires

Objectif : pouvoir justifier un résultat à tout moment, y compris face à un expert-comptable ou un auditeur.

---

# 🔐 Sécurité & Bonnes pratiques

- Aucun code client n'est exécuté
- Validation des types et conversions sécurisées
- Gestion robuste des erreurs (fichiers corrompus, protégés, mal formés)
- Données et fichiers `.env` non versionnés (`.gitignore`)
- Logs exploitables en cas de contrôle

---

# 💼 Cas d'usage

- Consolidation mensuelle des ventes
- Reporting commercial multi-fichiers
- Préparation reporting expert-comptable
- Contrôle et fiabilisation des exports CRM
- Analyse interne direction

---

# 📈 Valeur ajoutée

Gain estimé :

- 1 à 3 heures économisées par mois
- Réduction du risque d'erreur humaine
- Meilleure traçabilité et auditabilité
- Standardisation du reporting

---

# 🧠 Technologies

- Python 3.10+
- Pandas
- OpenPyXL
- ReportLab
- Streamlit
- python-dotenv
- Pytest
- Git

---

# 👨‍💻 Auteur

Vivien Gauzelin  
Data Analyst | Automatisation & fiabilisation de données  

Projet démonstration dans le cadre d'une activité freelance spécialisée en automatisation de processus et reporting.