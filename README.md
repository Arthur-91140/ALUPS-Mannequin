# ALUPS — Suivi des Mannequins

Application web de traçabilité des mannequins de formation (nettoyage, changement des poumons, réparation).

## Fonctionnalités

- **Formulaire public** (mobile) : saisie des interventions avec signature tactile
- **Liens directs** par mannequin via paramètres d'URL (pour QR Codes)
- **Panel admin** (desktop) : gestion des mannequins, historique, export Excel

## Installation

```bash
pip install -r requirements.txt
python3 app.py
```

## Déploiement (production)

```bash
gunicorn -w 2 -b 0.0.0.0:5000 app:app
```

## Stack

- Python / Flask
- SQLite
- Bootstrap 5
- signature_pad.js
- openpyxl (export Excel)
