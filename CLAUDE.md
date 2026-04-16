# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SOMATRIN is a Django-based fuel (gasoil) reporting application that fetches data from an Odoo ERP backend via XML-RPC and displays it with filters and KPIs.

## Commands

Always use `--settings=somatrin.settings_local` with Django management commands:

```bash
# Install dependencies
pip install -r requirements.txt

# Run development server
python manage.py runserver --settings=somatrin.settings_local

# Apply migrations
python manage.py migrate --settings=somatrin.settings_local

# Create superuser
python manage.py createsuperuser --settings=somatrin.settings_local
```

URLs:
- Home: `http://127.0.0.1:8000/`
- Fuel report: `http://127.0.0.1:8000/gasoil/sorties/`
- Admin: `http://127.0.0.1:8000/admin/`

## Architecture

### Data Flow

All business data lives in Odoo (not the local SQLite DB). Django is a thin presentation layer:

1. `somatrin/urls.py` → routes to `reporting/urls.py`
2. `reporting/views.py` → `get_odoo_connection()` authenticates via XML-RPC
3. Views build an Odoo search domain from GET filters, call `stock.move.search_read()`
4. Results are post-processed (anomaly filtering, KPI aggregation) and passed to templates

### Odoo Integration

Configured in `somatrin/settings_local.py`:
```python
ODOO_URL  = 'http://127.0.0.1:8001'
ODOO_DB   = 'somatrin'
ODOO_USER = 'admin'
ODOO_PASS = 'admin'
```

The app queries `stock.move` records with custom fields prefixed `x_` (e.g., `x_chauffeur`, `x_affectation`, `x_cpt_initial`, `x_anomalie`).

### Apps

- **`reporting/`** — Main app: views, URL patterns, fuel reporting logic
- **`core/`** — Secondary app: currently a placeholder for auth/shared functionality
- **`somatrin/`** — Django project config: `settings_local.py`, root `urls.py`

### Templates

Located in `templates/`. The main template is `gasoil/sorties.html` (filter form + KPI cards + data table). A `base.html` is referenced but must exist for template inheritance.

### Settings Module

The project uses `somatrin.settings_local` (not `somatrin.settings`). Always specify this explicitly.
