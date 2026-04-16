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

### ✅ Module 1 — Gasoil (pilote, en production)

Suivi de la consommation de carburant multi-sites.

- Bons de sortie gasoil avec filtres (date, site, chauffeur, ouvrage, catégorie)
- KPI : total bons, litres consommés, anomalies détectées, consommation moyenne
- Détection automatique des écarts entre compteur initial et compteur actuel
- Tableau avec mise en évidence des lignes anormales

---

### 🔧 Module 2 — Transport & Flotte

Suivi des bons de transport et de l'activité des véhicules.

- Liste des bons de transport par date, chauffeur, véhicule, destination
- Tonnages transportés par site et par période
- Suivi des chauffeurs : nombre de rotations, tonnage cumulé
- KPI : total rotations, tonnage total, véhicules actifs

---

### 🔧 Module 3 — Maintenance / GMAO

Suivi des interventions de maintenance sur les équipements.

- Liste des ordres de maintenance (préventive / corrective) par équipement
- Suivi des interventions : durée, technicien, statut, coût estimé
- Taux de disponibilité des équipements
- KPI : interventions du mois, équipements en panne, délai moyen de résolution

---

### 🔧 Module 4 — Achats

Suivi des commandes fournisseurs et des délais de livraison.

- Liste des bons de commande par fournisseur, date, statut
- Analyse des délais : commande → livraison réelle vs prévue
- Répartition des achats par catégorie de produit
- KPI : commandes en cours, montant total achats, fournisseurs actifs, retards

---

### 🔧 Module 5 — Comptabilité

Suivi des dépenses et des clôtures mensuelles.

- Dépenses par compte, par service, par mois
- Suivi des factures fournisseurs : émises, validées, en attente
- Évolution mensuelle des charges
- KPI : total dépenses du mois, factures en attente, écart budget/réel

---

### 🔧 Module 6 — Production / MRP

Suivi des quantités produites par site et par période.

- Production journalière et mensuelle par carrière
- Comparatif objectif vs réalisé
- Répartition par type de produit (granulats, concassés, etc.)
- KPI : tonnage produit, taux de réalisation, sites actifs

---

### 🔧 Module 7 — Qualité / QHSE

Suivi des incidents, non-conformités et contrôles qualité.

- Enregistrement et suivi des incidents par site, type, gravité
- Contrôles qualité : résultats des analyses sur les produits
- Tableau de bord sécurité : fréquence des incidents, jours sans accident
- KPI : incidents du mois, taux de conformité, actions correctives en cours

---

### 🔧 Module 8 — RH

Suivi des effectifs, présences et éléments de paie.

- Effectifs par site et par département
- Suivi des présences et absences par période
- Éléments variables de paie (heures supplémentaires, primes)
- KPI : effectif total, taux d'absentéisme, heures supplémentaires du mois

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
python manage.py runserver 8090 --settings=somatrin.settings_local
```

Accès : [http://127.0.0.1:8090](http://127.0.0.1:8090)

---

## Auteur

Projet réalisé dans le cadre d'un stage de fin d'études — SOMATRIN, Exploitation de Carrières, Maroc.
