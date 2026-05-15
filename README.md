# Reporting SOMATRIN

> **Sujet de stage** : Conception et développement d'une application web de reporting multi-services connectée à l'ERP Odoo via API XML-RPC

**Dépôt GitHub :** [https://github.com/fadoulimane1-debug/Somatrin-dashboard](https://github.com/fadoulimane1-debug/Somatrin-dashboard)

Application web de reporting interne développée pour **SOMATRIN — Exploitation de Carrières**, permettant la visualisation et l'analyse des données métier extraites depuis l'ERP Odoo, sans passer par son interface native.

> **Accès au dépôt :** projet hébergé en dépôt **privé** sur GitHub. L'encadrant doit être ajouté en collaborateur (*Settings → Collaborators*) pour consulter le code.

---

## Contexte

SOMATRIN est une entreprise spécialisée dans l'exploitation de carrières au Maroc. Face à la multiplicité des services opérationnels (transport, maintenance, production, achats, finance, QHSE, etc.) et à la dispersion des données dans l'ERP Odoo, la direction a exprimé le besoin d'une plateforme centralisée de reporting, accessible, lisible et adaptée aux besoins métier de chaque service.

Ce projet répond à ce besoin en développant une application web **Django** connectée à Odoo via son API **XML-RPC** standard, sans nécessiter de développement côté Odoo.

---

## Objectifs

- Centraliser les données de tous les services dans une interface unique
- Offrir des tableaux de bord avec KPI, filtres dynamiques et exports (Excel, PDF, CSV)
- Détecter automatiquement les anomalies (ex. : écarts de consommation gasoil)
- Fournir une architecture extensible pour intégrer de nouveaux modules facilement
- Permettre une consultation sans accès direct à l'interface Odoo

---

## Outils et technologies

### Backend

| Outil | Version | Rôle |
|---|---|---|
| **Python** | 3.12+ | Langage principal |
| **Django** | 4.2 | Framework web — routes, vues, templates, authentification |
| **xmlrpc.client** | stdlib | Connexion et interrogation Odoo via XML-RPC |
| **SQLite** | — | Base locale Django (sessions, utilisateurs) |
| **openpyxl / reportlab / weasyprint** | — | Exports Excel et PDF |

### Frontend

| Outil | Version | Rôle |
|---|---|---|
| **HTML5 / CSS3** | — | Structure et styles par module |
| **Bootstrap** | 5.3 | Mise en page responsive, composants UI |
| **Bootstrap Icons** | 1.11 | Icônes |
| **JavaScript** | ES6 | Interactions dynamiques, graphiques |

### ERP & données

| Outil | Version | Rôle |
|---|---|---|
| **Odoo** | 19 | ERP source des données métier (migration depuis v16) |
| **API XML-RPC** | — | Échange Django ↔ Odoo |

### Environnement

| Outil | Rôle |
|---|---|
| **Visual Studio Code / Cursor** | Édition du code |
| **Git / GitHub** | Versionnement et hébergement |
| **PowerShell** | Terminal (Windows) |
| **pip** | Dépendances Python |

---

## Modules de l'application

Légende : ✅ opérationnel · 🔧 en cours d'enrichissement · 📋 planifié

### ✅ Module 1 — Gasoil

- Entrées, sorties, bilan
- Filtres dynamiques, KPI, détection d'anomalies
- Exports CSV / PDF

### ✅ Module 2 — Transport & logistique

- Bons de transport, gasoil, coûts par nature
- Facturation client, rentabilité

### ✅ Module 3 — Production

- Tableau de bord, gasoil, pointages foration
- Machines (heures / tonnages), coûts par nature, ratios, IPC
- Facturation ventes, rentabilité, rapports, sites
- Exports dédiés

### ✅ Module 4 — Parc & maintenance

- Demandes, équipements, disponibilité
- Ordres de travail, interventions, fournisseurs, coûts
- Exports PDF / Excel / CSV

### 🔧 Module 5 — Achats & approvisionnement

- Vue d'ensemble, demandes d'achat, demandes de prix
- Bons de commande, suivi livraisons, fournisseurs
- Synthèse PDF

### ✅ Module 6 — Comptabilité

- Factures fournisseurs / client, décaissements, encaissements
- Trésorerie, analyse projet, TVA

### ✅ Module 7 — Finance

- Tableau de bord, factures clients / fournisseurs
- Avoirs, paiements, rapports, configuration
- Détail facture, exports Excel / PDF

### ✅ Module 8 — QHSE

- Tableau de bord, incidents & accidents, plan d'actions
- Achats QHSE, consommations, produits HSE
- Indicateurs, audits qualité, entrées / sorties / bilan

### 📋 Modules planifiés

- **Ressources humaines** — effectifs, absences, pointage, formations
- **Système d'information** — projets SI, parc informatique, tickets

### Assistants & API

- Chatbot métier (`/chatbot/`, `/api/soma-ai/chat/`)
- Endpoints KPI Production et QHSE

---

## Architecture du projet

```
Somatrin-dashboard/
│
├── somatrin/                      # Configuration Django
│   ├── urls.py                    # Routes racine (auth, reporting)
│   ├── wsgi.py
│   └── settings_local.py          # Fichier LOCAL uniquement (non versionné)
│
├── core/                          # Auth, rôles, context processors
│
├── reporting/                     # Application principale
│   ├── views.py                   # Vues gasoil, transport, achats, finance…
│   ├── urls.py                    # Routage principal
│   ├── production_urls.py         # Sous-module Production
│   ├── production_views.py
│   ├── qhse_urls.py               # Sous-module QHSE
│   ├── qhse_views.py
│   ├── models.py
│   ├── migrations/
│   ├── services/                  # Couche métier Odoo
│   │   ├── odoo_service.py
│   │   ├── finance_service.py
│   │   ├── parc_service.py
│   │   ├── production_service.py
│   │   ├── qhse_service.py
│   │   └── soma_ai_v2.py
│   └── utils/export_utils.py
│
├── templates/                     # Gabarits HTML par module
│   ├── base.html
│   ├── accueil.html
│   ├── gasoil/ · transport/ · production/
│   ├── parc/ · achats/ · comptabilite/
│   └── finance/ · qhse/
│
├── static/
│   ├── css/ · js/
│   └── images/logo_somatrin.png
│
├── manage.py
└── requirements.txt
```

**Principe :** les données métier résident dans **Odoo** ; Django est une couche de **présentation et reporting**.

---

## Connexion Odoo via XML-RPC

```python
import xmlrpc.client

common = xmlrpc.client.ServerProxy('http://odoo-url/xmlrpc/2/common')
uid = common.authenticate(db, user, password, {})

models = xmlrpc.client.ServerProxy('http://odoo-url/xmlrpc/2/object')
data = models.execute_kw(db, uid, password, 'stock.move', 'search_read', [domain], {})
```

Connexion centralisée dans `reporting/services/odoo_service.py` (JSON-RPC) et `reporting/views.py` (XML-RPC).

---

## Migration Odoo 16 → 19

L'entreprise a migré l'ERP vers **Odoo 19**. L'application SOMATRIN reste compatible via les API standard (`xmlrpc/2` et `jsonrpc`) ; après migration, vérifier :

- Les **identifiants** et l'**URL** dans `somatrin/settings_local.py` (nouvelle instance Odoo 19)
- Les **modèles métier** toujours présents (`stock.move`, `quality.check`, champs `x_*` custom)
- Les **modules Odoo** installés (Qualité, Maintenance, Comptabilité, Achats, etc.)
- Les écrans qui renvoient une erreur Odoo (logs Django / message dans l'interface)

En cas de changement de nom de modèle ou de champ entre v16 et v19, adapter les services dans `reporting/services/`.

---

## Installation

### Prérequis

- Python 3.12+
- Instance **Odoo 19** accessible
- Git

### Étapes

```bash
git clone https://github.com/fadoulimane1-debug/Somatrin-dashboard.git
cd Somatrin-dashboard

python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
```

Créer `somatrin/settings_local.py` (non fourni sur GitHub) :

```python
ODOO_URL = 'http://127.0.0.1:8001'
ODOO_DB = 'somatrin'
ODOO_USER = 'admin'
ODOO_PASS = 'votre_mot_de_passe'
```

```bash
python manage.py migrate --settings=somatrin.settings_local
python manage.py runserver 8091 --settings=somatrin.settings_local
```

**Accès :** [http://127.0.0.1:8091](http://127.0.0.1:8091)

---

## Fichiers exclus du dépôt

| Fichier / dossier | Raison |
|---|---|
| `somatrin/settings_local.py` | Identifiants Odoo |
| `db.sqlite3` | Base locale |
| `__pycache__/`, `*.pyc` | Cache Python |
| `memory/` | Notes locales |

---

## Auteur

Projet réalisé dans le cadre d'un **stage de fin d'études** — **SOMATRIN**, Exploitation de Carrières, Maroc.
