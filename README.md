# Reporting SOMATRIN — Guide d'installation complet

## Structure du projet

```
D:\projets\Somatrin\
│
├── somatrin/
│   ├── __init__.py
│   ├── urls.py
│   └── settings_local.py
│
├── reporting/
│   ├── __init__.py
│   ├── apps.py
│   ├── urls.py
│   └── views.py
│
├── templates/
│   ├── base.html
│   ├── accueil.html
│   └── gasoil/
│       └── sorties.html
│
├── static/
│   └── images/
│       └── logo_somatrin.png   ← mettre le vrai logo ici
│
├── requirements.txt
└── manage.py
```

---

## Étapes d'installation

### 1. Installer les dépendances
```bash
pip install -r requirements.txt
```

### 2. Copier les fichiers
Copie tous les fichiers fournis dans les bons dossiers selon la structure ci-dessus.

### 3. Mettre le logo
Place le fichier `Logo_SOMATRIN_RVB.png` dans :
```
static/images/logo_somatrin.png
```

### 4. Configurer Odoo dans settings_local.py
```python
ODOO_URL  = 'http://127.0.0.1:8001'   # URL de ton Odoo
ODOO_DB   = 'somatrin'                 # Nom de ta base
ODOO_USER = 'admin'                    # Ton utilisateur
ODOO_PASS = 'admin'                    # Ton mot de passe
```

### 5. Adapter les noms de champs Odoo
Les champs personnalisés dans views.py (x_chauffeur, x_affectation, etc.)
doivent correspondre aux vrais noms dans ton Odoo.

Pour vérifier les vrais noms, lance dans un shell Python :
```python
import xmlrpc.client
common = xmlrpc.client.ServerProxy('http://127.0.0.1:8001/xmlrpc/2/common')
uid = common.authenticate('somatrin', 'admin', 'admin', {})
models = xmlrpc.client.ServerProxy('http://127.0.0.1:8001/xmlrpc/2/object')
fields = models.execute_kw('somatrin', uid, 'admin', 'stock.move', 'fields_get', [], {'attributes': ['string', 'type']})
for k, v in fields.items():
    if k.startswith('x_'):
        print(k, '->', v['string'])
```

### 6. Créer les migrations et lancer
```bash
python manage.py migrate --settings=somatrin.settings_local
python manage.py runserver --settings=somatrin.settings_local
```

### 7. Accéder à l'application
Ouvre : http://127.0.0.1:8000

---

## Pages disponibles

| URL | Description |
|-----|-------------|
| `/` | Page d'accueil |
| `/gasoil/sorties/` | Liste des sorties gasoil avec filtres |

## Filtres disponibles sur /gasoil/sorties/

| Paramètre GET | Description |
|---------------|-------------|
| `date_debut` | Date de début (YYYY-MM-DD) |
| `date_fin` | Date de fin (YYYY-MM-DD) |
| `site` | Filtrer par site (LHOUJ, LHMEK...) |
| `categorie` | Filtrer par UDM (H...) |
| `chauffeur` | Recherche par nom chauffeur |
| `ouvrage` | Recherche par affectation/ouvrage |
| `anomalie` | OK / Anomalie / vide = tous |
