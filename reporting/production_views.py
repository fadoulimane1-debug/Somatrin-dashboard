"""
Vues Production — module autonome SOMATRIN.
Vues utilisant ProductionService pour dashboard/ratios/ipc/rapports.
Vues complexes (Odoo live) déléguées à reporting.views.
"""
from __future__ import annotations

import json
import logging
from datetime import datetime

from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_GET, require_http_methods

from reporting.utils.export_utils import (
    export_to_excel,
    export_to_csv,
    export_to_pdf,
    export_pointages_operations_pdf,
    pointages_rows_to_export_dicts,
    POINTAGES_EXPORT_COLUMNS,
)
from reporting.views import pointages_operations_request_context

from reporting.services.production_service import ProductionService

logger = logging.getLogger(__name__)

POINTAGES_EXPORT_MAX_ROWS = 8000


def _pointages_export_subtitle_lines(ctx):
    """Résumé des filtres pour PDF / métadonnées export."""
    lines = []
    d0, d1 = ctx.get('date_debut') or '', ctx.get('date_fin') or ''
    if d0 or d1:
        lines.append(f"Période : du {d0 or '—'} au {d1 or '—'}")
    if ctx.get('site'):
        lines.append(f"Site : {ctx['site']}")
    if ctx.get('chauffeur'):
        lines.append(f"Chauffeur / équipe : {ctx['chauffeur']}")
    if ctx.get('societe'):
        lines.append(f"Société : {ctx['societe']}")
    if ctx.get('ouvrage'):
        ov = ctx['ouvrage']
        lines.append(f"Ouvrage : {ov[:120]}{'…' if len(ov) > 120 else ''}")
    if ctx.get('engin'):
        lines.append(f"Engin : {ctx['engin']}")
    if ctx.get('operation_q'):
        lines.append(f"N° opération / bon : {ctx['operation_q']}")
    ff = (ctx.get('foration_filtre') or '').lower()
    if ff == 'oui':
        lines.append('Foration : oui')
    elif ff == 'non':
        lines.append('Foration : non')
    an = (ctx.get('anomalie') or '').lower()
    if an == 'ok':
        lines.append('Anomalie : OK uniquement')
    elif an == 'anomalie':
        lines.append('Anomalie : anomalies uniquement')
    return lines


def _pointages_export_source_note(ctx, *, used_fallback_rows: bool):
    parts = []
    if ctx.get('demo_notice'):
        parts.append(str(ctx['demo_notice']))
    if ctx.get('error'):
        parts.append(str(ctx['error']))
    if used_fallback_rows and not (ctx.get('rows') or []):
        parts.append("Aucune ligne à exporter pour ces filtres — contenu d'exemple SOMATRIN.")
    return ' '.join(parts) if parts else ''


def _pointages_rows_for_export(ctx):
    rows = list(ctx.get('rows') or [])
    used_fb = False
    if not rows:
        rows = ProductionService.pointages_operations_fallback_rows()
        used_fb = True
    return rows[:POINTAGES_EXPORT_MAX_ROWS], used_fb


# ── 1. Dashboard Production ────────────────────────────────────────────────────

@login_required
def production_dashboard(request):
    """Tableau de bord production — KPIs + 4 graphiques."""
    context = ProductionService.dashboard_data()
    context['page_title'] = 'Dashboard Production'
    return render(request, 'production/dashboard.html', context)


# ── 2. Gasoil (délégation Odoo live) ─────────────────────────────────────────

@login_required
def production_gasoil(request):
    """Gasoil production — données réelles Odoo."""
    from reporting.views import production_gasoil as _v
    return _v(request)


# ── 3. Production détail (délégation Odoo live) ───────────────────────────────

@login_required
def production_detail(request):
    """Production par site — données réelles Odoo."""
    from reporting.views import production_index as _v
    return _v(request)


# ── 4. Pointages Opérations & Foration (délégation Odoo live) ────────────────

@login_required
def production_pointages(request):
    """Pointages opérations et foration — données réelles Odoo."""
    from reporting.views import production_pointages_operations_foration as _v
    return _v(request)


# ── 5. Production Machines / Heures / Tonnages (délégation Odoo live) ────────

@login_required
def production_machines(request):
    """Machines, heures et tonnages — données réelles Odoo."""
    from reporting.views import production_machines_heures_tonnages as _v
    return _v(request)


# ── 6. Coûts par nature (délégation Odoo live) ────────────────────────────────

@login_required
def production_couts(request):
    """Coûts par nature — données réelles Odoo (account.analytic.line)."""
    from reporting.views import production_couts_nature as _v
    return _v(request)


# ── 7. Ratio Exploitation ─────────────────────────────────────────────────────

@login_required
def production_ratios(request):
    """Ratios exploitation — Ammoniaca, Toner, Amoniaq, Radiateur."""
    context = ProductionService.ratios_data()
    context['page_title'] = 'Ratios Exploitation'
    return render(request, 'production/ratios.html', context)


# ── 8. IPC ────────────────────────────────────────────────────────────────────

@login_required
def production_ipc(request):
    """Index Performance Cost par site."""
    context = ProductionService.ipc_data()
    context['page_title'] = 'IPC'
    return render(request, 'production/ipc.html', context)


# ── 9. Facturation / Ventes (délégation Odoo live) ────────────────────────────

@login_required
def production_ventes(request):
    """Facturation & ventes — données réelles Odoo (account.move)."""
    from reporting.views import production_facturation_ventes as _v
    return _v(request)


# ── 10. Rentabilité (délégation Odoo live) ────────────────────────────────────

@login_required
def production_rentabilite(request):
    """Analyse rentabilité — données réelles Odoo."""
    from reporting.views import production_rentabilite as _v
    return _v(request)


# ── 11. Rapports & Analyses ────────────────────────────────────────────────────

@login_required
def production_rapports(request):
    """Rapports & analyses — génération PDF/Excel."""
    from reporting.views import production_rapports as _v
    return _v(request)


# ── 12. Sites ─────────────────────────────────────────────────────────────────

@login_required
def production_sites(request):
    """Gestion et comparaison des 3 sites de production."""
    from reporting.views import production_sites as _v
    return _v(request)


# ── Index (redirect) ──────────────────────────────────────────────────────────

@login_required
def production_index(request):
    """Page d'accueil module production."""
    from reporting.views import production_index as _v
    return _v(request)


# ── API Endpoints ──────────────────────────────────────────────────────────────

@login_required
@require_GET
def api_production_kpis(request):
    """GET /production/api/kpis/ — KPI dashboard JSON."""
    data = ProductionService.get_dashboard_kpis()
    return JsonResponse(data)


@login_required
@require_GET
def api_production_charts(request):
    """GET /production/api/charts/ — données graphiques JSON."""
    dashboard = ProductionService.dashboard_data()
    sites = ProductionService.get_sites()
    rentabilite = ProductionService.get_rentabilite_data()
    return JsonResponse({
        'dashboard': dashboard.get('charts', {}),
        'sites_noms': [s['name'] for s in sites],
        'sites_tonnages': [s['tonnage'] for s in sites],
        'ca_monthly': rentabilite['charts']['ca_values'],
        'ca_labels': rentabilite['charts']['labels'],
        'marge_monthly': rentabilite['charts']['marge_values'],
    })


# ── Exports Excel / CSV / PDF ──────────────────────────────────────────────────

@login_required
@require_http_methods(["GET"])
def export_gasoil_excel(request):
    data = [{'Date': '2026-05-13', 'Véhicule': 'VOLVO-001', 'Chauffeur': 'Ahmed Bennani', 'Litres': 150.50, 'Montant': 3011.00, 'Statut': 'OK'}]
    cols = ['Date', 'Véhicule', 'Chauffeur', 'Litres', 'Montant', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Gasoil_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Sorties Gasoil")


@login_required
@require_http_methods(["GET"])
def export_gasoil_csv(request):
    data = [{'Date': '2026-05-13', 'Véhicule': 'VOLVO-001', 'Chauffeur': 'Ahmed Bennani', 'Litres': 150.50, 'Montant': 3011.00, 'Statut': 'OK'}]
    cols = ['Date', 'Véhicule', 'Chauffeur', 'Litres', 'Montant', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Gasoil_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_gasoil_pdf(request):
    data = [{'Date': '2026-05-13', 'Véhicule': 'VOLVO-001', 'Chauffeur': 'Ahmed Bennani', 'Litres': 150.50, 'Montant': 3011.00, 'Statut': 'OK'}]
    cols = ['Date', 'Véhicule', 'Chauffeur', 'Litres', 'Montant', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Gasoil_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Sorties Gasoil")


@login_required
@require_http_methods(["GET"])
def export_production_excel(request):
    data = [{'Date': '2026-05-13', 'Machine': 'M001', 'Heures': 8.5, 'Tonnages': 250.00, 'Rendement': 29.41, 'Statut': 'OK'}]
    cols = ['Date', 'Machine', 'Heures', 'Tonnages', 'Rendement', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Production_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Production")


@login_required
@require_http_methods(["GET"])
def export_production_csv(request):
    data = [{'Date': '2026-05-13', 'Machine': 'M001', 'Heures': 8.5, 'Tonnages': 250.00, 'Rendement': 29.41, 'Statut': 'OK'}]
    cols = ['Date', 'Machine', 'Heures', 'Tonnages', 'Rendement', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Production_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_production_pdf(request):
    data = [{'Date': '2026-05-13', 'Machine': 'M001', 'Heures': 8.5, 'Tonnages': 250.00, 'Rendement': 29.41, 'Statut': 'OK'}]
    cols = ['Date', 'Machine', 'Heures', 'Tonnages', 'Rendement', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Production_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Production")


@login_required
@require_http_methods(["GET"])
def export_pointages_excel(request):
    ctx = pointages_operations_request_context(request)
    rows, _ = _pointages_rows_for_export(ctx)
    data = pointages_rows_to_export_dicts(rows)
    title = "SOMATRIN — Pointages opérations & foration"
    return export_to_excel(
        data,
        f"SOMATRIN_Pointages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        list(POINTAGES_EXPORT_COLUMNS),
        title,
    )


@login_required
@require_http_methods(["GET"])
def export_pointages_csv(request):
    ctx = pointages_operations_request_context(request)
    rows, _ = _pointages_rows_for_export(ctx)
    data = pointages_rows_to_export_dicts(rows)
    return export_to_csv(
        data,
        f"SOMATRIN_Pointages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        list(POINTAGES_EXPORT_COLUMNS),
    )


@login_required
@require_http_methods(["GET"])
def export_pointages_pdf(request):
    ctx = pointages_operations_request_context(request)
    rows, used_fb = _pointages_rows_for_export(ctx)
    data = pointages_rows_to_export_dicts(rows)
    subtitle = _pointages_export_subtitle_lines(ctx)
    source = _pointages_export_source_note(ctx, used_fallback_rows=used_fb)
    return export_pointages_operations_pdf(
        data,
        f"SOMATRIN_Pointages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        title="SOMATRIN — Pointages opérations & foration",
        subtitle_lines=subtitle,
        source_note=source,
    )


@login_required
@require_http_methods(["GET"])
def export_machines_excel(request):
    data = [{'Machine': 'M001', 'Heures': 240.5, 'Tonnages': 7500.00, 'Ratio': 31.17, 'Statut': 'Actif'}]
    cols = ['Machine', 'Heures', 'Tonnages', 'Ratio', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Machines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Machines/Heures")


@login_required
@require_http_methods(["GET"])
def export_machines_csv(request):
    data = [{'Machine': 'M001', 'Heures': 240.5, 'Tonnages': 7500.00, 'Ratio': 31.17, 'Statut': 'Actif'}]
    cols = ['Machine', 'Heures', 'Tonnages', 'Ratio', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Machines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_machines_pdf(request):
    data = [{'Machine': 'M001', 'Heures': 240.5, 'Tonnages': 7500.00, 'Ratio': 31.17, 'Statut': 'Actif'}]
    cols = ['Machine', 'Heures', 'Tonnages', 'Ratio', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Machines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Machines/Heures")


@login_required
@require_http_methods(["GET"])
def export_couts_excel(request):
    data = [{'Nature': 'Matière Première', 'Montant': 150000.00, 'Pourcentage': 45.0, 'Statut': 'OK'}]
    cols = ['Nature', 'Montant', 'Pourcentage', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Couts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Coûts par Nature")


@login_required
@require_http_methods(["GET"])
def export_couts_csv(request):
    data = [{'Nature': 'Matière Première', 'Montant': 150000.00, 'Pourcentage': 45.0, 'Statut': 'OK'}]
    cols = ['Nature', 'Montant', 'Pourcentage', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Couts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_couts_pdf(request):
    data = [{'Nature': 'Matière Première', 'Montant': 150000.00, 'Pourcentage': 45.0, 'Statut': 'OK'}]
    cols = ['Nature', 'Montant', 'Pourcentage', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Couts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Coûts par Nature")


@login_required
@require_http_methods(["GET"])
def export_ratios_excel(request):
    data = [{'Ratio': 'OEE', 'Valeur': 85.5, 'Cible': 90.0, 'Écart': -4.5, 'Statut': 'Alerte'}]
    cols = ['Ratio', 'Valeur', 'Cible', 'Écart', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Ratios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Ratios d'Exploitation")


@login_required
@require_http_methods(["GET"])
def export_ratios_csv(request):
    data = [{'Ratio': 'OEE', 'Valeur': 85.5, 'Cible': 90.0, 'Écart': -4.5, 'Statut': 'Alerte'}]
    cols = ['Ratio', 'Valeur', 'Cible', 'Écart', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Ratios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_ratios_pdf(request):
    data = [{'Ratio': 'OEE', 'Valeur': 85.5, 'Cible': 90.0, 'Écart': -4.5, 'Statut': 'Alerte'}]
    cols = ['Ratio', 'Valeur', 'Cible', 'Écart', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Ratios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Ratios d'Exploitation")


@login_required
@require_http_methods(["GET"])
def export_ipc_excel(request):
    data = [{'Date': '2026-05-13', 'Tonnage': 250.0, 'Heures': 8.5, 'IPC': 29.41, 'Tendance': '+'}]
    cols = ['Date', 'Tonnage', 'Heures', 'IPC', 'Tendance']
    return export_to_excel(data, f"SOMATRIN_IPC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - IPC (Indice Productivité)")


@login_required
@require_http_methods(["GET"])
def export_ipc_csv(request):
    data = [{'Date': '2026-05-13', 'Tonnage': 250.0, 'Heures': 8.5, 'IPC': 29.41, 'Tendance': '+'}]
    cols = ['Date', 'Tonnage', 'Heures', 'IPC', 'Tendance']
    return export_to_csv(data, f"SOMATRIN_IPC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_ipc_pdf(request):
    data = [{'Date': '2026-05-13', 'Tonnage': 250.0, 'Heures': 8.5, 'IPC': 29.41, 'Tendance': '+'}]
    cols = ['Date', 'Tonnage', 'Heures', 'IPC', 'Tendance']
    return export_to_pdf(data, f"SOMATRIN_IPC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - IPC (Indice Productivité)")


@login_required
@require_http_methods(["GET"])
def export_facturation_excel(request):
    data = [{'Date': '2026-05-13', 'Client': 'Client A', 'Montant HT': 10000.00, 'Montant TTC': 12000.00, 'Statut': 'Payée'}]
    cols = ['Date', 'Client', 'Montant HT', 'Montant TTC', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Facturation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Facturation (Ventes)")


@login_required
@require_http_methods(["GET"])
def export_facturation_csv(request):
    data = [{'Date': '2026-05-13', 'Client': 'Client A', 'Montant HT': 10000.00, 'Montant TTC': 12000.00, 'Statut': 'Payée'}]
    cols = ['Date', 'Client', 'Montant HT', 'Montant TTC', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Facturation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_facturation_pdf(request):
    data = [{'Date': '2026-05-13', 'Client': 'Client A', 'Montant HT': 10000.00, 'Montant TTC': 12000.00, 'Statut': 'Payée'}]
    cols = ['Date', 'Client', 'Montant HT', 'Montant TTC', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Facturation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Facturation (Ventes)")


@login_required
@require_http_methods(["GET"])
def export_rentabilite_excel(request):
    data = [{'Site': 'Settat', 'Chiffre Affaires': 500000.00, 'Charges': 300000.00, 'Marge': 200000.00, 'Statut': 'Bon'}]
    cols = ['Site', 'Chiffre Affaires', 'Charges', 'Marge', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Rentabilite_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Rentabilité")


@login_required
@require_http_methods(["GET"])
def export_rentabilite_csv(request):
    data = [{'Site': 'Settat', 'Chiffre Affaires': 500000.00, 'Charges': 300000.00, 'Marge': 200000.00, 'Statut': 'Bon'}]
    cols = ['Site', 'Chiffre Affaires', 'Charges', 'Marge', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Rentabilite_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_rentabilite_pdf(request):
    data = [{'Site': 'Settat', 'Chiffre Affaires': 500000.00, 'Charges': 300000.00, 'Marge': 200000.00, 'Statut': 'Bon'}]
    cols = ['Site', 'Chiffre Affaires', 'Charges', 'Marge', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Rentabilite_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Rentabilité")


@login_required
@require_http_methods(["GET"])
def export_rapports_excel(request):
    data = [{'Date': '2026-05-13', 'Type': 'Hebdomadaire', 'Sujet': 'Performance Production', 'Statut': 'Complété'}]
    cols = ['Date', 'Type', 'Sujet', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Rapports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Rapports & Analyses")


@login_required
@require_http_methods(["GET"])
def export_rapports_csv(request):
    data = [{'Date': '2026-05-13', 'Type': 'Hebdomadaire', 'Sujet': 'Performance Production', 'Statut': 'Complété'}]
    cols = ['Date', 'Type', 'Sujet', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Rapports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_rapports_pdf(request):
    data = [{'Date': '2026-05-13', 'Type': 'Hebdomadaire', 'Sujet': 'Performance Production', 'Statut': 'Complété'}]
    cols = ['Date', 'Type', 'Sujet', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Rapports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Rapports & Analyses")


@login_required
@require_http_methods(["GET"])
def export_sites_excel(request):
    data = [
        {'Site': 'Settat', 'Production': 5000.00, 'Charges': 300000.00, 'Rendement': 92.5, 'Statut': 'Actif'},
        {'Site': 'Casablanca', 'Production': 4500.00, 'Charges': 280000.00, 'Rendement': 88.3, 'Statut': 'Actif'},
        {'Site': 'Marrakech', 'Production': 3500.00, 'Charges': 200000.00, 'Rendement': 85.0, 'Statut': 'Actif'},
    ]
    cols = ['Site', 'Production', 'Charges', 'Rendement', 'Statut']
    return export_to_excel(data, f"SOMATRIN_Sites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", cols, "SOMATRIN - Sites")


@login_required
@require_http_methods(["GET"])
def export_sites_csv(request):
    data = [
        {'Site': 'Settat', 'Production': 5000.00, 'Charges': 300000.00, 'Rendement': 92.5, 'Statut': 'Actif'},
        {'Site': 'Casablanca', 'Production': 4500.00, 'Charges': 280000.00, 'Rendement': 88.3, 'Statut': 'Actif'},
        {'Site': 'Marrakech', 'Production': 3500.00, 'Charges': 200000.00, 'Rendement': 85.0, 'Statut': 'Actif'},
    ]
    cols = ['Site', 'Production', 'Charges', 'Rendement', 'Statut']
    return export_to_csv(data, f"SOMATRIN_Sites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", cols)


@login_required
@require_http_methods(["GET"])
def export_sites_pdf(request):
    data = [
        {'Site': 'Settat', 'Production': 5000.00, 'Charges': 300000.00, 'Rendement': 92.5, 'Statut': 'Actif'},
        {'Site': 'Casablanca', 'Production': 4500.00, 'Charges': 280000.00, 'Rendement': 88.3, 'Statut': 'Actif'},
        {'Site': 'Marrakech', 'Production': 3500.00, 'Charges': 200000.00, 'Rendement': 85.0, 'Statut': 'Actif'},
    ]
    cols = ['Site', 'Production', 'Charges', 'Rendement', 'Statut']
    return export_to_pdf(data, f"SOMATRIN_Sites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", cols, "SOMATRIN - Sites")
