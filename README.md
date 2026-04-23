# Reporting SOMATRIN

> **Sujet de stage** : Conception et Développement d'une Application Web de Reporting Multi-Services connectée à l'ERP Odoo via API XML-RPC

Application web de reporting interne développée pour **SOMATRIN — Exploitation de Carrières**, permettant la visualisation et l'analyse des données métier extraites en temps réel depuis l'ERP Odoo, sans passer par son interface native.

---

## Contexte

SOMATRIN est une entreprise spécialisée dans l'exploitation de carrières au Maroc. Face à la multiplicité des services opérationnels (transport, maintenance, production, RH, achats, etc.) et à la dispersion des données dans l'ERP Odoo, la direction a exprimé le besoin d'une plateforme centralisée de reporting, accessible, lisible et adaptée aux besoins métier de chaque service.

Ce projet répond à ce besoin en développant une application web Django connectée à Odoo via son API XML-RPC standard, sans nécessiter de développement côté Odoo.

---

## Objectifs

- Centraliser les données de tous les services dans une interface unique
- Offrir des tableaux de bord avec KPI, filtres dynamiques et exports
- Détecter automatiquement les anomalies (ex. : écarts de consommation gasoil)
- Fournir une architecture extensible pour intégrer de nouveaux modules facilement
- Permettre une consultation sans accès direct à Odoo

---

## Outils et technologies utilisés

### Backend
| Outil | Version | Rôle |
|---|---|---|
| **Python** | 3.14 | Langage principal du projet |
| **Django** | 4.2 | Framework web — gestion des routes, vues, templates |
| **xmlrpc.client** | stdlib | Connexion et interrogation de l'ERP Odoo via API XML-RPC |
| **SQLite** | — | Base de données locale (sessions, utilisateurs Django) |

### Frontend
| Outil | Version | Rôle |
|---|---|---|
| **HTML5 / CSS3** | — | Structure et style des pages |
| **Bootstrap** | 5.3 | Framework CSS — mise en page responsive, composants UI |
| **Bootstrap Icons** | 1.11 | Icônes vectorielles |
| **JavaScript** | ES6 | Interactions dynamiques côté client |

### ERP & Données
| Outil | Version | Rôle |
|---|---|---|
| **Odoo** | 16 | ERP source de toutes les données métier |
| **API XML-RPC** | — | Protocole d'échange entre Django et Odoo |

### Environnement de développement
| Outil | Rôle |
|---|---|
| **Visual Studio Code** | Éditeur de code principal |
| **Git / GitHub** | Versionnement et hébergement du code source |
| **PowerShell** | Terminal de commandes (Windows) |
| **pip** | Gestionnaire de paquets Python |

---

## Modules de l'application

### ✅ Module 1 — Gasoil (terminé)

- Entrées
- Sorties
- Bilan

---

### ✅ Module 2 — Transport & Logistique (terminé)

- Bons transport
- Gasoil
- Coûts par nature
- Facturation client
- Rentabilité

---

### 🔧 Module 3 — Production (en cours)

- Gasoil
- Coûts par nature
- Facturation ventes
- Rentabilité
- Sites

---

### 🔧 Module 4 — Parc & Maintenance

- Ordres de travail
- Disponibilité équipements
- Historique

---

### 🔧 Module 5 — Achats & Approvisionnement

- Bons de commande
- Fournisseurs
- Analyse des dépenses

---

### 🔧 Module 6 — Comptabilité

- Indicateurs financiers
- Analytique
- Coûts de revient

---

### 🔧 Module 7 — QHSE

- Indicateurs HSE
- Gestion alertes
- Suivi qualité

---

### 🔧 Module 8 — Ressources Humaines

- Effectifs
- Absences
- Pointage
- Formations

---

### 🔧 Module 9 — Système d'Information

- Projets SI
- Parc informatique
- Tickets incidents

---

## Architecture

```
somatrin_project/
│
├── somatrin/                  # Configuration Django
│   ├── settings_local.py      # Paramètres locaux (Odoo, BDD, etc.)
│   └── urls.py                # Routes principales
│
├── reporting/                 # App principale
│   ├── views.py               # Logique métier + appels Odoo XML-RPC
│   └── urls.py                # Routes du module reporting
│
├── templates/                 # Templates HTML
│   ├── base.html              # Layout global (navbar, breadcrumb)
│   ├── accueil.html           # Page d'accueil avec KPI et modules
│   └── gasoil/
│       └── sorties.html       # Sorties gasoil avec filtres
│
├── static/
│   └── images/
│       └── logo_somatrin.png
│
├── manage.py
└── requirements.txt
```

---

## Connexion Odoo via XML-RPC

L'application interroge Odoo via son API standard (`xmlrpc/2`). Aucun développement côté Odoo n'est requis — seule une connexion réseau et un compte utilisateur suffisent.

```python
import xmlrpc.client

common = xmlrpc.client.ServerProxy('http://odoo-url/xmlrpc/2/common')
uid = common.authenticate(db, user, password, {})

models = xmlrpc.client.ServerProxy('http://odoo-url/xmlrpc/2/object')
data = models.execute_kw(db, uid, password, 'stock.move', 'search_read', [domain], {})
```

---

## Installation

```bash
# 1. Cloner le projet
git clone https://github.com/ton-user/somatrin-reporting.git
cd somatrin-reporting

# 2. Installer les dépendances
pip install -r requirements.txt

# 3. Configurer la connexion Odoo dans somatrin/settings_local.py
ODOO_URL  = 'http://127.0.0.1:8001'
ODOO_DB   = 'somatrin'
ODOO_USER = 'admin'
ODOO_PASS = 'admin'

# 4. Initialiser et lancer
python manage.py migrate --settings=somatrin.settings_local
python manage.py runserver 8091 --settings=somatrin.settings_local
```

Accès : [http://127.0.0.1:8091](http://127.0.0.1:8091)

---

## Auteur

Projet réalisé dans le cadre d'un stage de fin d'études — SOMATRIN, Exploitation de Carrières, Maroc.
