"""
Vues QHSE — module autonome Odoo 16.
Toutes les routes /qhse/ sont implémentées ici sans dépendance sur reporting.views.
"""
from __future__ import annotations

import logging
from datetime import datetime

from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_GET, require_http_methods

from reporting.services.odoo_service import OdooServiceManager
from reporting.services.qhse_service import QHSEService

logger = logging.getLogger(__name__)


def _fmt_number(value) -> str:
    """Formate un nombre en français (séparateur milliers espace insécable)."""
    try:
        v = float(value or 0)
        s = f'{v:,.2f}'
        integer, decimal = s.split('.')
        integer = integer.replace(',', ' ')
        return f'{integer},{decimal}'
    except (TypeError, ValueError):
        return str(value or 0)


# ── Dashboard ──────────────────────────────────────────────────────────────────

@login_required
def qhse_dashboard(request):
    """Tableau de bord QHSE — données Odoo réelles via QHSEService."""
    svc = QHSEService()
    data = svc.get_dashboard_data()
    return render(request, 'qhse/dashboard.html', {
        'page_title': 'Dashboard QHSE',
        'annee': datetime.now().year,
        **data,
    })


# ── Incidents & Accidents ──────────────────────────────────────────────────────

@login_required
def qhse_incidents(request):
    """Incidents & Accidents — synoptique corporel, graphiques, tableau Odoo."""
    site = (request.GET.get('site') or '').strip()
    mois = (request.GET.get('mois') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(site_filter=site, mois_filter=mois)
    chart = QHSEService.to_chart_json(acc)
    selected_month = mois or (chart['months_list'][-1] if chart['months_list'] else '')
    return render(request, 'qhse/incidents_accidents.html', {
        'page_title': 'Incidents & Accidents',
        'page_subtitle': "Gestion et analyse des incidents, accidents de travail et maladies professionnelles.",
        'selected_site': site,
        'selected_month': selected_month,
        'months_list': chart['months_list'],
        'kpi_avec_arret': acc.get('kpi_avec_arret', 0),
        'kpi_sans_arret': acc.get('kpi_sans_arret', 0),
        'kpi_incidents': acc.get('kpi_incidents', 0),
        'kpi_jours_arret': acc.get('kpi_jours_arret', 0),
        'body_zones': acc.get('body_zones', {}),
        'body_total': acc.get('body_total', 0),
        'rows': acc.get('rows', []),
        'error': acc.get('error'),
        **chart,
    })


# ── Plan d'actions JAM ─────────────────────────────────────────────────────────

@login_required
def qhse_plan_actions(request):
    """Plan d'actions JAM — project.task / action.schedule.plan."""
    state_filter = (request.GET.get('state') or '').strip()
    assignee_filter = (request.GET.get('assignee') or '').strip()
    project_filter = (request.GET.get('project') or '').strip()
    search_query = (request.GET.get('q') or '').strip()
    svc = QHSEService()
    data = svc.get_actions(
        state_filter=state_filter,
        assignee_filter=assignee_filter,
        project_filter=project_filter,
        search_query=search_query,
        limit=800,
    )
    return render(request, 'qhse/plan_actions.html', {
        'page_title': "Plan d'actions",
        'page_subtitle': "Suivi des actions correctives, préventives et d'amélioration (JAM).",
        'rows': data.get('rows', []),
        'kpi_total': data.get('kpi_total', 0),
        'kpi_open': data.get('kpi_open', 0),
        'kpi_closed': data.get('kpi_closed', 0),
        'kpi_taux_cloture': data.get('kpi_taux_cloture', 0),
        'kpi_en_retard': data.get('kpi_en_retard', 0),
        'kpi_efficacite': data.get('kpi_efficacite', 0),
        'by_project': data.get('by_project', []),
        'by_type': data.get('by_type', []),
        'top_assignees': data.get('top_assignees', []),
        'selected_state': state_filter,
        'selected_assignee': assignee_filter,
        'selected_project': project_filter,
        'search_query': search_query,
        'error': data.get('error'),
    })


# ── Détail action JAM ──────────────────────────────────────────────────────────

@login_required
def qhse_action_detail(request, action_id: int):
    """Détail d'une action JAM."""
    from django.http import Http404
    svc = QHSEService()
    action = svc.get_action_detail(action_id)
    if action is None:
        raise Http404('Action introuvable')
    return render(request, 'qhse/action_detail.html', {
        'page_title': f"Action #{action_id}",
        'action': action,
        'error': svc._error,
    })


# ── Achats QHSE ───────────────────────────────────────────────────────────────

@login_required
def qhse_achats(request):
    """Commandes d'achat QHSE — purchase.order depuis Odoo."""
    statut = (request.GET.get('statut') or 'all').strip().lower()
    fournisseur = (request.GET.get('fournisseur') or '').strip()
    svc = QHSEService()
    data = svc.get_achats_qhse(status_filter=statut, supplier_filter=fournisseur)

    return render(request, 'qhse/achats.html', {
        'page_title': 'Achats QHSE',
        'page_subtitle': 'Suivi des achats de catégorie QHSE récupérés depuis Odoo.',
        'error': data.get('error'),
        'rows': data.get('rows', []),
        'fournisseurs': data.get('fournisseurs', []),
        'selected_fournisseur': fournisseur,
        'selected_statut': statut,
        'kpi_total_commandes': data.get('kpi_total_commandes', 0),
        'kpi_fournisseurs': data.get('kpi_fournisseurs', 0),
        'kpi_en_attente': data.get('kpi_en_attente', 0),
        'kpi_montant': _fmt_number(data.get('kpi_montant', 0)),
    })


# ── Consommations QHSE ────────────────────────────────────────────────────────

@login_required
def qhse_consommations(request):
    """Sorties EPI — stock.move.line depuis Odoo."""
    site = (request.GET.get('site') or '').strip()
    personne = (request.GET.get('personne') or '').strip()
    svc = QHSEService()
    data = svc.get_consommations_qhse(site_filter=site, person_filter=personne)
    return render(request, 'qhse/consommations.html', {
        'page_title': 'Consommations QHSE',
        'page_subtitle': 'Sorties EPI détaillées : affectation, transfert site et suivi par personne.',
        'rows': data.get('rows', []),
        'sites': data.get('sites', []),
        'personnes': data.get('personnes', []),
        'selected_site': site,
        'selected_personne': personne,
        'kpis': data.get('kpis', {}),
        'error': data.get('error'),
    })


# ── Produits HSE ──────────────────────────────────────────────────────────────

@login_required
def qhse_produits_hse(request):
    """Catalogue produits HSE/EPI — product.product depuis Odoo."""
    statut = (request.GET.get('statut') or 'all').strip().lower()
    recherche = (request.GET.get('q') or '').strip()
    svc = QHSEService()
    data = svc.get_produits_hse(status_filter=statut, search_term=recherche)

    return render(request, 'qhse/produits_hse.html', {
        'page_title': 'Produits HSE',
        'page_subtitle': 'Catalogue des équipements de protection individuelle (EPI) et produits HSE.',
        'error': data.get('error'),
        'rows': data.get('rows', []),
        'selected_statut': statut,
        'search_term': recherche,
        'kpi_total': data.get('kpi_total', 0),
        'kpi_en_stock': data.get('kpi_en_stock', 0),
        'kpi_achat': data.get('kpi_achat', 0),
        'kpi_stock_total': _fmt_number(data.get('kpi_stock_total', 0)),
    })


# ── Indicateurs HSE ───────────────────────────────────────────────────────────

@login_required
def qhse_indicateurs(request):
    """Extincteurs, TF, TG — maintenance.equipment depuis Odoo."""
    site = (request.GET.get('site') or '').strip()
    svc = QHSEService()
    data = svc.get_indicateurs_hse(site_filter=site)
    return render(request, 'qhse/indicateurs.html', {
        'page_title': 'Indicateurs HSE (TF, TG)',
        'page_subtitle': 'Suivi extincteurs, contrôle réglementaire et alertes J-7.',
        'selected_site': site,
        'rows': data.get('rows', []),
        'alertes': data.get('alertes', []),
        'kpi_total': data.get('kpi_total', 0),
        'kpi_alertes': data.get('kpi_alertes', 0),
        'kpi_echus': data.get('kpi_echus', 0),
        'kpi_j7': data.get('kpi_j7', 0),
        'tf_score': data.get('tf_score', 0.0),
        'tg_score': data.get('tg_score', 0.0),
        'error': data.get('error'),
    })


# ── Audits Qualité ────────────────────────────────────────────────────────────

@login_required
def qhse_audits(request):
    """Audits qualité — synthèse incidents par site + stock EPI."""
    svc = QHSEService()
    acc = svc.get_accidents(limit=2000)
    prod_data = svc.get_produits_hse()
    from collections import defaultdict
    by_cat: dict = defaultdict(float)
    for p in prod_data.get('rows', []):
        by_cat[p.get('categorie') or 'Non renseigné'] += float(p.get('stock') or 0)
    return render(request, 'qhse/audits.html', {
        'page_title': 'Audits qualité',
        'page_subtitle': 'Rapports par site et stock détail EPI.',
        'incidents_by_site': acc.get('sites', []),
        'stock_by_category': sorted(by_cat.items(), key=lambda x: x[1], reverse=True),
        'error': acc.get('error') or prod_data.get('error'),
    })


# ── Factures QHSE (stub) ──────────────────────────────────────────────────────

@login_required
def qhse_factures(request):
    """Factures QHSE — page d'information."""
    return render(request, 'qhse/bilan.html', {
        'page_title': 'Factures QHSE',
        'page_subtitle': 'Facturation des dépenses et services QHSE.',
        'alerts': [], 'summary': {}, 'sites': [], 'error': None,
    })


# ── Bilan QHSE ────────────────────────────────────────────────────────────────

@login_required
def qhse_bilan(request):
    """Bilan QHSE — résumé des alertes qualité."""
    site = (request.GET.get('site') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(site_filter=site, limit=500)
    rows = acc.get('rows', [])
    summary = {
        'total': acc.get('kpi_total', 0),
        'open': acc.get('kpi_open', 0),
        'in_progress': 0,
        'done': acc.get('kpi_closed', 0),
    }
    sites_list = acc.get('sites', [])
    return render(request, 'qhse/bilan.html', {
        'page_title': 'Bilan QHSE',
        'page_subtitle': 'Indicateurs QHSE récupérés depuis Odoo.',
        'site': site,
        'sites': [{'name': s[0]} for s in sites_list],
        'alerts': rows[:50],
        'summary': summary,
        'error': acc.get('error'),
    })


# ── Entrées QHSE ─────────────────────────────────────────────────────────────

@login_required
def qhse_entrees(request):
    """Entrées QHSE — conformité et contrôles entrants."""
    site = (request.GET.get('site') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(site_filter=site, limit=500)
    rows = [r for r in acc.get('rows', []) if r.get('state') not in ('done', 'cancel')]
    sites_list = acc.get('sites', [])
    return render(request, 'qhse/entrees.html', {
        'page_title': 'Entrées QHSE',
        'page_subtitle': 'Données de conformité et incidents entrants depuis Odoo.',
        'site': site,
        'sites': [{'name': s[0]} for s in sites_list],
        'alerts': rows[:50],
        'summary': {
            'total': len(rows), 'open': len(rows),
            'in_progress': 0, 'done': 0,
        },
        'error': acc.get('error'),
    })


# ── Sorties QHSE ─────────────────────────────────────────────────────────────

@login_required
def qhse_sorties(request):
    """Sorties QHSE — actions et rapports clôturés."""
    site = (request.GET.get('site') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(site_filter=site, limit=500)
    rows = [r for r in acc.get('rows', []) if r.get('state') == 'done']
    sites_list = acc.get('sites', [])
    return render(request, 'qhse/sorties.html', {
        'page_title': 'Sorties QHSE',
        'page_subtitle': "Rapports de sortie et actions QHSE provenant d'Odoo.",
        'site': site,
        'sites': [{'name': s[0]} for s in sites_list],
        'alerts': rows[:50],
        'summary': {
            'total': len(rows), 'open': 0,
            'in_progress': 0, 'done': len(rows),
        },
        'error': acc.get('error'),
    })


# ── Configuration QHSE ────────────────────────────────────────────────────────

@login_required
@require_http_methods(['GET', 'POST'])
def configuration_qhse(request):
    """Préférences affichage QHSE (stockage session)."""
    message = ''
    if request.method == 'POST':
        request.session['qhse_dashboard_compact'] = bool(request.POST.get('compact'))
        request.session['qhse_default_site'] = (request.POST.get('default_site') or '').strip()
        message = 'Préférences enregistrées pour cette session.'
    return render(request, 'qhse/configuration.html', {
        'page_title': 'Configuration QHSE',
        'message': message,
        'compact': request.session.get('qhse_dashboard_compact', False),
        'default_site': request.session.get('qhse_default_site', ''),
        'odoo_ok': OdooServiceManager().test_connection(),
    })


# ── API Endpoints ──────────────────────────────────────────────────────────────

@login_required
@require_GET
def api_qhse_kpis(request):
    """GET /qhse/api/kpis/ — KPI dashboard JSON."""
    mgr = OdooServiceManager()
    return JsonResponse(mgr.get_dashboard_kpis())


@login_required
@require_GET
def api_incidents_month(request):
    """GET /qhse/api/incidents/?month=YYYY-MM — incidents du mois JSON."""
    mois = (request.GET.get('month') or request.GET.get('mois') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(mois_filter=mois, limit=2500)
    chart = QHSEService.to_chart_json(acc)
    return JsonResponse({
        'month': mois,
        'incidents_arret': acc.get('kpi_avec_arret', 0),
        'incidents_sans_arret': acc.get('kpi_sans_arret', 0),
        'incidents_autre': acc.get('kpi_incidents', 0),
        'total_jours_arret': acc.get('kpi_jours_arret', 0),
        'rows_sample': acc.get('rows', [])[:50],
        'chart': chart,
        'error': acc.get('error'),
    })
