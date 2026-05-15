import csv
import html
import hashlib
import io
import json
import logging
import calendar
import re
import ssl
import unicodedata
import xmlrpc.client
from urllib.parse import urlencode
from collections import defaultdict
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, PageBreak)
    from reportlab.pdfgen import canvas as rl_canvas
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'reportlab'])
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, PageBreak)
    from reportlab.pdfgen import canvas as rl_canvas

from django.contrib.auth.decorators import login_required
from django.views.decorators.csrf import csrf_exempt
from django.core.cache import cache
from django.core.paginator import Paginator
from django.http import HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import render
from django.conf import settings
from django.utils import timezone
from reporting.services.soma_ai_v2 import OdooConnector, MetierDataService, SomaAIEngine
from reporting.services.parc_service import ParcOdooService
from reporting.services.production_service import ProductionService

logger = logging.getLogger(__name__)

# Catégorie Odoo pour le carburant (BIEN MATERIEL / ENERGIE / CARBURANT)
CARBURANT_CATEG_ID = 262

# Distinction métier confirmée via equipment_id.category_id (Odoo).
# Transport: camions, bennes, véhicules routiers.
CATEGORIES_TRANSPORT = [18, 19, 21, 23, 43, 48, 49, 50]
# Production: pelles, bulldozers, foreuses, compresseurs, engins carrière.
CATEGORIES_PRODUCTION = [25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 39, 40, 41, 42, 44, 45, 53]

SITES_LIST = [
    'AIN JEMAA', 'CIMAR AIT BAHA', 'CIMAR AMSKROUD', 'CIMAT BENI MELLAL',
    'GRABEMARO BENSLIMANE', 'LH BENSLIMANE', 'LH MEKNES', 'LH OUJDA',
    'SETTAT DÉPÔT', 'Virtual Locations/Consommation', 'YOUSSOUFIA',
    'LHOUJ/Stock', 'LHMEK/Stock',
]


def format_number(value):
    """Formate un nombre en français: milliers espace insécable, décimales virgule."""
    if value is None:
        return '0,00'
    try:
        value = float(value)
        formatted = f'{value:,.2f}'
        integer, decimal = formatted.split('.')
        integer = integer.replace(',', '\u00a0')
        return f'{integer},{decimal}'
    except (TypeError, ValueError):
        return str(value)


def format_number_decimals(value, decimals=2):
    """Formate un nombre avec décimales fixes et séparateurs français."""
    if value is None:
        if decimals <= 0:
            return '0'
        return '0,' + ('0' * decimals)
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    fmt = f'{{:,.{decimals}f}}'
    formatted = fmt.format(v)
    if decimals <= 0:
        return formatted.replace(',', '\u00a0')
    integer, decimal = formatted.split('.')
    integer = integer.replace(',', '\u00a0')
    return f'{integer},{decimal}'


def extract_matricule(name):
    if not name:
        return ''
    return str(name).split('/', 1)[0].strip()


def _activity_bucket_from_picking(picking, ouvrage_text=''):
    """
    Détection métier robuste:
    - transport
    - voiture_service
    - production
    """
    if picking.get('service_car'):
        return 'voiture_service'
    if picking.get('transport_logistics'):
        return 'transport'

    parts = [ouvrage_text]
    for field in ('account_analytic_id', 'affectation_id', 'equipment_id', 'location_id'):
        val = picking.get(field)
        if isinstance(val, list) and len(val) > 1 and val[1]:
            parts.append(val[1])
    raw = ' '.join(str(v) for v in parts if v).lower()
    if ('transport' in raw) or ('logist' in raw):
        return 'transport'
    return 'production'


def _build_project_activity_map(uid, models, invoices):
    """
    Construit une map {project_id: 'transport'|'production'} depuis project.project.
    Utilise le booléen transport_logistics quand disponible.
    """
    project_ids = []
    for inv in invoices:
        p = inv.get('project_id')
        if isinstance(p, list) and p:
            project_ids.append(p[0])
    project_ids = sorted({pid for pid in project_ids if pid})
    if not project_ids:
        return {}

    try:
        projects = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'project.project', 'search_read',
            [[('id', 'in', project_ids)]],
            {'fields': ['id', 'name', 'transport_logistics'], 'limit': len(project_ids)},
        )
        out = {}
        for p in projects:
            name = (p.get('name') or '').lower()
            if p.get('transport_logistics') or ('transport' in name) or ('logist' in name):
                out[p['id']] = 'transport'
            else:
                out[p['id']] = 'production'
        return out
    except Exception:
        # Fallback si champ custom absent/inaccessible.
        return {}


def _account_code_from_aml_account_id(account_id_field):
    """Extrait le code compte depuis le many2one account_id [id, libellé Odoo]."""
    if not isinstance(account_id_field, (list, tuple)) or len(account_id_field) < 2:
        return ''
    label = str(account_id_field[1] or '')
    token = (label.split() or [''])[0]
    return ''.join(ch for ch in token if ch.isdigit())[:12] or token[:12]


def _rentabilite_include_move_for_activity(move, activity, perimetre, proj_map):
    """
    Périmètre activité : factures classées transport/production + pièces diverses
    dont la référence / libellé évoque l'activité. Périmètre complet : tout le 6 / 7.
    """
    if perimetre == 'complet' or not activity:
        return True
    mt = (move.get('move_type') or '')
    if mt in ('out_invoice', 'in_invoice', 'out_refund', 'in_refund'):
        return _invoice_activity_bucket(move, proj_map) == activity
    refn = ' '.join(str(move.get(k) or '') for k in ('ref', 'name')).lower()
    if activity == 'transport':
        return ('transport' in refn) or ('logist' in refn)
    if activity == 'production':
        return ('transport' not in refn) and ('logist' not in refn)
    return True


def _rent_cache_key(activity: str, perimetre: str, date_debut: str, date_fin: str, company_id) -> str:
    """Clé de cache déterministe pour le dashboard de rentabilité."""
    raw = f"rent|{activity}|{perimetre}|{date_debut or ''}|{date_fin or ''}|{company_id or ''}"
    return 'somatrin_rent_' + hashlib.md5(raw.encode()).hexdigest()[:16]


def _rentabilite_pcg_readgroup(uid, models, date_debut: str, date_fin: str, company_id=None):
    """
    CHEMIN RAPIDE — agrège debit/credit par compte côté Odoo via read_group.
    Évite de rapatrier des milliers de lignes : 2-3 appels XML-RPC au lieu de N*batch.
    Utilisé pour périmètre='complet'. Retourne le même format dict que
    _rentabilite_pcg_aggregate, ou None si read_group échoue (fallback).
    """
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS

    fix_prefs = tuple(getattr(settings, 'RENTABILITE_CHARGE_FIXE_PREFIXES', ('64', '65', '66', '67', '68')))
    top6_n = max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_CHARGES', 20)))
    top7_n = max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_PRODUITS', 20)))

    domain = [
        ('move_id.state', '=', 'posted'),
        ('display_type', 'not in', ('line_section', 'line_note')),
        '|',
        ('account_id.internal_group', '=', 'income'),
        ('account_id.internal_group', '=', 'expense'),
    ]
    if date_debut:
        domain.append(('date', '>=', date_debut))
    if date_fin:
        domain.append(('date', '<=', date_fin))
    if company_id and str(company_id).isdigit():
        domain.append(('company_id', '=', int(company_id)))

    try:
        groups = models.execute_kw(
            db, uid, pw, 'account.move.line', 'read_group',
            [domain, ['account_id', 'debit', 'credit'], ['account_id']],
            {'lazy': False, 'limit': 600},
        )
    except Exception as exc:
        logger.info('rentabilite read_group indisponible (%s) → fallback', type(exc).__name__)
        return None

    if not groups:
        return None

    # Récupérer codes et libellés en un seul appel (au lieu de N*batch)
    acc_ids = [g['account_id'][0] for g in groups if isinstance(g.get('account_id'), (list, tuple))]
    acc_map: dict = {}
    for i in range(0, len(acc_ids), 200):
        try:
            for a in models.execute_kw(
                db, uid, pw, 'account.account', 'search_read',
                [[('id', 'in', acc_ids[i:i + 200])]],
                {'fields': ['id', 'code', 'name']},
            ):
                acc_map[a['id']] = a
        except Exception:
            continue

    tot_7 = tot_6 = charges_fixes = charges_variables = 0.0
    by6: dict = {}
    by7: dict = {}
    lib6: dict = {}
    lib7: dict = {}
    nb_lignes = sum(int(g.get('account_id_count') or 0) for g in groups)

    for g in groups:
        acc_id_val = g.get('account_id')
        if not isinstance(acc_id_val, (list, tuple)):
            continue
        acc_info = acc_map.get(acc_id_val[0], {})
        code = (acc_info.get('code') or '').strip()
        if not code:
            continue
        lead = code[0]
        if lead not in ('6', '7'):
            continue
        debit  = float(g.get('debit')  or 0)
        credit = float(g.get('credit') or 0)
        acc_name = (acc_info.get('name') or '').strip() or '—'
        if lead == '7':
            amt = max(0.0, credit - debit)
            tot_7 += amt
            by7[code] = by7.get(code, 0) + amt
            lib7.setdefault(code, acc_name)
        else:
            amt = max(0.0, debit - credit)
            tot_6 += amt
            by6[code] = by6.get(code, 0) + amt
            lib6.setdefault(code, acc_name)
            if any(code.startswith(p) for p in fix_prefs):
                charges_fixes += amt
            else:
                charges_variables += amt

    tot_6 = round(tot_6, 2)
    tot_7 = round(tot_7, 2)
    charges_fixes = round(charges_fixes, 2)
    charges_variables = round(charges_variables, 2)
    resultat = round(tot_7 - tot_6, 2)
    taux_rentabilite = round((resultat / tot_7 * 100), 2) if tot_7 else 0.0
    contribution = round(tot_7 - charges_variables, 2)
    taux_marge_contribution = round((contribution / tot_7 * 100), 2) if tot_7 else 0.0
    marge_sur_cv_pct = round((contribution / charges_variables * 100), 2) if charges_variables else None
    taux_mc = (charges_fixes / contribution) if contribution > 0 else None
    seuil_ca_ht = round(charges_fixes / taux_mc, 2) if (taux_mc and taux_mc > 0 and charges_fixes > 0) else None

    sorted6 = sorted(by6.items(), key=lambda kv: -kv[1])
    sorted7 = sorted(by7.items(), key=lambda kv: -kv[1])

    logger.info(
        'rentabilite read_group OK: %s cptes-6 | %s cptes-7 | %s lignes GL (fast path)',
        len(by6), len(by7), nb_lignes,
    )
    return {
        'total_produits_7': tot_7,
        'total_charges_6': tot_6,
        'charges_fixes': charges_fixes,
        'charges_variables': charges_variables,
        'resultat': resultat,
        'taux_rentabilite': taux_rentabilite,
        'contribution': contribution,
        'taux_marge_contribution': taux_marge_contribution,
        'marge_sur_cv_pct': marge_sur_cv_pct,
        'seuil_ca_ht': seuil_ca_ht,
        'nb_lignes_gl': nb_lignes,
        'nb_lignes_apres_activite': nb_lignes,
        'comptes_6': [{'code': k, 'libelle': lib6.get(k, '—'), 'montant': round(v, 2)} for k, v in sorted6[:top6_n]],
        'comptes_7': [{'code': k, 'libelle': lib7.get(k, '—'), 'montant': round(v, 2)} for k, v in sorted7[:top7_n]],
        'comptes_6_hors_top': max(0, len(by6) - top6_n),
        'comptes_7_hors_top': max(0, len(by7) - top7_n),
        'top_comptes_charges_limite': top6_n,
        'top_comptes_produits_limite': top7_n,
        '_fast_path': True,
    }


def _rentabilite_pcg_aggregate(uid, models, date_debut, date_fin, company_id=None, activity=None, perimetre='complet'):
    """
    Totaux PCG depuis Odoo (grand livre) : comptes 7 (produits / ventes) et 6 (charges).
    Retourne un dict de KPI + listes de détail par code compte.

    Le domaine Odoo utilise ``internal_group`` income/expense (fiable sur toutes les versions),
    puis on ne retient que les comptes dont le **code** commence par 6 ou 7 (PCG).
    Les codes sont lus sur ``account.account`` (le libellé seul sur la ligne n’expose pas toujours le n°).
    """
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS
    fix_prefs = tuple(getattr(settings, 'RENTABILITE_CHARGE_FIXE_PREFIXES', ('64', '65', '66', '67', '68')))
    var_prefs = tuple(getattr(settings, 'RENTABILITE_CHARGE_VARIABLE_PREFIXES', ('60', '61', '62', '63')))
    top6_n = max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_CHARGES', 20)))
    top7_n = max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_PRODUITS', 20)))
    max_lines = int(getattr(settings, 'RENTABILITE_MAX_LINES', 5000))

    domain = [
        ('move_id.state', '=', 'posted'),
        '|',
        ('account_id.internal_group', '=', 'income'),
        ('account_id.internal_group', '=', 'expense'),
    ]
    if date_debut:
        domain.append(('date', '>=', date_debut))
    if date_fin:
        domain.append(('date', '<=', date_fin))
    if company_id and str(company_id).isdigit():
        domain.append(('company_id', '=', int(company_id)))

    # batch_size réduit + cap global pour éviter les timeouts XML-RPC
    _batch = 1000
    domain_gl = list(domain) + [('display_type', 'not in', ('line_section', 'line_note'))]
    try:
        lines = odoo_search_read_all(
            uid, models, 'account.move.line', domain_gl,
            ['id', 'date', 'debit', 'credit', 'balance', 'account_id', 'move_id', 'company_id'],
            batch_size=_batch,
            order='date desc,id desc',
        )
    except Exception:
        try:
            lines = odoo_search_read_all(
                uid, models, 'account.move.line', domain,
                ['id', 'date', 'debit', 'credit', 'balance', 'account_id', 'move_id', 'company_id'],
                batch_size=_batch,
                order='date desc,id desc',
            )
        except Exception:
            lines = []

    if not lines:
        domain_legacy = [
            ('move_id.state', '=', 'posted'),
            '|',
            ('account_id.code', '=like', '6%'),
            ('account_id.code', '=like', '7%'),
        ]
        if date_debut:
            domain_legacy.append(('date', '>=', date_debut))
        if date_fin:
            domain_legacy.append(('date', '<=', date_fin))
        if company_id and str(company_id).isdigit():
            domain_legacy.append(('company_id', '=', int(company_id)))
        dlegacy_gl = list(domain_legacy) + [('display_type', 'not in', ('line_section', 'line_note'))]
        try:
            lines = odoo_search_read_all(
                uid, models, 'account.move.line', dlegacy_gl,
                ['id', 'date', 'debit', 'credit', 'balance', 'account_id', 'move_id', 'company_id'],
                batch_size=_batch,
                order='date desc,id desc',
            )
        except Exception:
            try:
                lines = odoo_search_read_all(
                    uid, models, 'account.move.line', domain_legacy,
                    ['id', 'date', 'debit', 'credit', 'balance', 'account_id', 'move_id', 'company_id'],
                    batch_size=_batch,
                    order='date desc,id desc',
                )
            except Exception:
                lines = []

    # Cap global — tronque les données si trop nombreuses
    lines_capped = len(lines) > max_lines
    if lines_capped:
        logger.warning(
            'rentabilite aggregate: %s lignes tronquées à %s (max_lines)',
            len(lines), max_lines,
        )
        lines = lines[:max_lines]

    acc_ids = set()
    for ln in lines:
        aid = _odoo_m2o_id(ln.get('account_id'))
        if aid:
            acc_ids.add(aid)
    acc_map = {}
    aids = sorted(acc_ids)
    for i in range(0, len(aids), 400):
        chunk = aids[i:i + 400]
        try:
            acc_recs = models.execute_kw(
                db, uid, pw, 'account.account', 'search_read',
                [[('id', 'in', chunk)]],
                {'fields': ['id', 'code', 'name', 'internal_group']},
            )
            for a in acc_recs:
                acc_map[a['id']] = a
        except Exception:
            continue

    move_ids = set()
    for ln in lines:
        mid = _odoo_m2o_id(ln.get('move_id'))
        if mid:
            move_ids.add(mid)
    move_map = {}
    mids = sorted(move_ids)
    chunk = 400
    for i in range(0, len(mids), chunk):
        chunk_ids = mids[i:i + chunk]
        chunk_moves = models.execute_kw(
            db, uid, pw, 'account.move', 'search_read',
            [[('id', 'in', chunk_ids)]],
            {'fields': ['id', 'move_type', 'project_id', 'invoice_origin', 'ref', 'name']},
        )
        for m in chunk_moves:
            move_map[m['id']] = m

    proj_map = _build_project_activity_map(uid, models, list(move_map.values()))

    tot_7 = 0.0
    tot_6 = 0.0
    charges_fixes = 0.0
    charges_variables = 0.0
    by6 = defaultdict(float)
    by7 = defaultdict(float)
    lib6 = {}
    lib7 = {}
    nb_apres_activite = 0

    for ln in lines:
        mid = _odoo_m2o_id(ln.get('move_id'))
        move = move_map.get(mid, {}) if mid else {}
        if not _rentabilite_include_move_for_activity(move, activity, perimetre, proj_map):
            continue
        nb_apres_activite += 1
        acc_id = _odoo_m2o_id(ln.get('account_id'))
        acc_info = acc_map.get(acc_id, {}) if acc_id else {}
        code = (acc_info.get('code') or '').strip()
        if not code:
            code = _account_code_from_aml_account_id(ln.get('account_id'))
        if not code:
            continue
        lead = code[0] if code else ''
        if lead not in ('6', '7'):
            continue
        debit = float(ln.get('debit') or 0)
        credit = float(ln.get('credit') or 0)
        acc_name = (acc_info.get('name') or '').strip()
        if lead == '7':
            amt = max(0.0, credit - debit)
            tot_7 += amt
            by7[code] += amt
            lib7.setdefault(code, acc_name or '—')
        else:
            amt = max(0.0, debit - credit)
            tot_6 += amt
            by6[code] += amt
            lib6.setdefault(code, acc_name or '—')
            if any(code.startswith(p) for p in fix_prefs):
                charges_fixes += amt
            elif any(code.startswith(p) for p in var_prefs):
                charges_variables += amt
            else:
                charges_variables += amt

    tot_6 = round(tot_6, 2)
    tot_7 = round(tot_7, 2)
    charges_fixes = round(charges_fixes, 2)
    charges_variables = round(charges_variables, 2)
    resultat = round(tot_7 - tot_6, 2)
    taux_rentabilite = round((resultat / tot_7 * 100), 2) if tot_7 else 0.0

    contribution = round(tot_7 - charges_variables, 2)
    taux_marge_contribution = round((contribution / tot_7 * 100), 2) if tot_7 else 0.0
    marge_sur_cv_pct = round((contribution / charges_variables * 100), 2) if charges_variables else None
    taux_mc = (charges_fixes / contribution) if contribution > 0 else None
    seuil_ca_ht = round(charges_fixes / taux_mc, 2) if (taux_mc and taux_mc > 0 and charges_fixes > 0) else None

    sorted6 = sorted(by6.items(), key=lambda kv: -kv[1])
    sorted7 = sorted(by7.items(), key=lambda kv: -kv[1])
    comptes_6 = [
        {'code': k, 'libelle': lib6.get(k) or '—', 'montant': round(v, 2)}
        for k, v in sorted6[:top6_n]
    ]
    comptes_7 = [
        {'code': k, 'libelle': lib7.get(k) or '—', 'montant': round(v, 2)}
        for k, v in sorted7[:top7_n]
    ]

    return {
        'total_produits_7': tot_7,
        'total_charges_6': tot_6,
        'charges_fixes': charges_fixes,
        'charges_variables': charges_variables,
        'resultat': resultat,
        'taux_rentabilite': taux_rentabilite,
        'contribution': contribution,
        'taux_marge_contribution': taux_marge_contribution,
        'marge_sur_cv_pct': marge_sur_cv_pct,
        'seuil_ca_ht': seuil_ca_ht,
        'nb_lignes_gl': len(lines),
        'nb_lignes_apres_activite': nb_apres_activite,
        'comptes_6': comptes_6,
        'comptes_7': comptes_7,
        'comptes_6_hors_top': max(0, len(by6) - top6_n),
        'comptes_7_hors_top': max(0, len(by7) - top7_n),
        'top_comptes_charges_limite': top6_n,
        'top_comptes_produits_limite': top7_n,
        '_lines_capped': lines_capped,
        '_fast_path': False,
    }


def _invoice_activity_bucket(invoice, project_activity_map=None):
    """
    Classe une facture vente dans:
    - transport
    - production
    selon project_id + libellés de référence.
    """
    parts = []
    proj = invoice.get('project_id')
    if project_activity_map and isinstance(proj, list) and proj:
        mapped = project_activity_map.get(proj[0])
        if mapped in {'transport', 'production'}:
            return mapped
    if isinstance(proj, list) and len(proj) > 1 and proj[1]:
        parts.append(proj[1])
    for field in ('invoice_origin', 'ref', 'name'):
        val = invoice.get(field)
        if val:
            parts.append(str(val))
    raw = ' '.join(parts).lower()
    if ('transport' in raw) or ('logist' in raw):
        return 'transport'
    return 'production'


def _enrich_sortie_bon(bon):
    """Ajoute les champs *_fmt pour affichage HTML (après calcul des valeurs brutes)."""
    bon['cpt_initial_fmt'] = format_number_decimals(bon.get('cpt_initial'), 0)
    bon['cpt_actuel_fmt'] = format_number_decimals(bon.get('cpt_actuel'), 0)
    bon['ecart_fmt'] = format_number_decimals(bon.get('ecart'), 1)
    bon['product_qty_fmt'] = format_number_decimals(bon.get('product_qty'), 1)
    c = bon.get('consommation')
    bon['consommation_fmt'] = format_number_decimals(c, 2) if c else ''


def _enrich_entree_bon(bon):
    bon['product_qty_fmt'] = format_number_decimals(bon.get('product_qty'), 1)
    pu = bon.get('price_unit')
    bon['price_unit_fmt'] = format_number_decimals(pu, 2) if pu not in (None, '', False) else ''
    tot = bon.get('total')
    bon['total_fmt'] = format_number_decimals(tot, 2) if tot not in (None, '', False) else ''


def get_odoo_connection():
    """Retourne (uid, models) pour les appels Odoo XML-RPC."""
    verify_ssl = getattr(settings, 'ODOO_SSL_VERIFY', True)
    server_proxy_kwargs = {}
    if settings.ODOO_URL.startswith('https://') and not verify_ssl:
        # Local/dev workaround when corporate/intermediate cert is missing.
        server_proxy_kwargs['context'] = ssl._create_unverified_context()

    common = xmlrpc.client.ServerProxy(
        f'{settings.ODOO_URL}/xmlrpc/2/common',
        **server_proxy_kwargs,
    )
    try:
        uid = common.authenticate(settings.ODOO_DB, settings.ODOO_USER, settings.ODOO_PASS, {})
    except xmlrpc.client.Fault as exc:
        msg = (exc.faultString or '').strip()
        if exc.faultCode == 3 or 'Access Denied' in msg:
            raise RuntimeError(
                "Odoo refuse la connexion (Access Denied). Vérifiez dans "
                "`somatrin/settings_local.py` : ODOO_URL (URL exacte du serveur), "
                "ODOO_DB (nom technique de la base, tel qu’affiché sur l’écran de connexion Odoo), "
                "ODOO_USER (souvent l’e-mail du compte) et ODOO_PASS. "
                "Sur Odoo 15+, si le mot de passe du compte ne fonctionne pas, générez une "
                "clé API (Profil utilisateur → Préférences → Clés API / Account Security) "
                "et utilisez-la comme ODOO_PASS."
            ) from exc
        raise
    if not uid:
        raise RuntimeError(
            "Odoo : authentification refusée (identifiant, mot de passe ou base de données incorrects)."
        )
    models = xmlrpc.client.ServerProxy(
        f'{settings.ODOO_URL}/xmlrpc/2/object',
        **server_proxy_kwargs,
    )
    return uid, models


def odoo_search_read_all(uid, models, model_name, domain, fields, batch_size=2000, order='id desc'):
    """Lit tous les enregistrements Odoo correspondant au domaine (sans limite artificielle)."""
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS
    all_rows = []
    offset = 0
    while True:
        ids = models.execute_kw(
            db, uid, pw, model_name, 'search',
            [domain],
            {'offset': offset, 'limit': batch_size, 'order': order},
        )
        if not ids:
            break
        rows = models.execute_kw(
            db, uid, pw, model_name, 'read',
            [ids],
            {'fields': fields},
        )
        all_rows.extend(rows)
        if len(ids) < batch_size:
            break
        offset += batch_size
    return all_rows


def _odoo_float(val):
    """Convertit une valeur numérique Odoo (souvent False si vide) en float."""
    if val in (None, False, ''):
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _odoo_m2o_id(val):
    """Many2one Odoo lu en XML-RPC : [id, libellé], parfois seul l’id (int)."""
    if val in (None, False, ''):
        return None
    if isinstance(val, (list, tuple)) and val:
        try:
            return int(val[0])
        except (TypeError, ValueError):
            return None
    try:
        return int(val)
    except (TypeError, ValueError):
        return None


def _odoo_currency_id_label(rec):
    """Many2one currency_id -> (id, libellé) pour agrégations fiables."""
    cur = rec.get('currency_id')
    if isinstance(cur, (list, tuple)) and len(cur) >= 2:
        try:
            return int(cur[0]), str(cur[1])
        except (TypeError, ValueError):
            return 0, 'MAD'
    if isinstance(cur, (list, tuple)) and len(cur) == 1:
        try:
            return int(cur[0]), str(cur[0])
        except (TypeError, ValueError):
            return 0, 'MAD'
    return 0, 'MAD'


@login_required
def accueil(request):
    today = date.today()
    JOURS_FR = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']
    MOIS_FR  = ['', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
                'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    date_affichee = f"{JOURS_FR[today.weekday()]} {today.day} {MOIS_FR[today.month]} {today.year}"

    return render(request, 'accueil.html', {
        'date_affichee': date_affichee,
    })


# ─────────────────────────────────────────────
#  GASOIL — SORTIES
#  Modèle : stock.picking (picking_type_consumption=True)
#  Produit filtré par catégorie CARBURANT (id=262)
# ─────────────────────────────────────────────
def _build_sorties_domain(date_debut, date_fin, site, chauffeur, ouvrage, anomalie,
                          societe='', categorie_engin='', activite_filtre=''):
    """Construit le domaine Odoo pour les bons de sortie gasoil."""
    domain = [
        ('state', '=', 'done'),
        ('picking_type_consumption', '=', True),
        ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
    ]
    if date_debut:
        domain.append(('scheduled_date', '>=', date_debut + ' 00:00:00'))
    if date_fin:
        domain.append(('scheduled_date', '<=', date_fin + ' 23:59:59'))
    if societe:
        domain.append(('company_id.name', '=', societe))
    if site:
        domain.append(('location_id.complete_name', 'ilike', site))
    if chauffeur:
        domain.append(('partner_id.name', 'ilike', chauffeur))
    
    if categorie_engin:
        domain.append(('equipment_id.category_id.name', '=', categorie_engin))
    # Filtres activité via champs booléens natifs Odoo
    if activite_filtre == 'transport':
        domain.append(('transport_logistics', '=', True))
    elif activite_filtre == 'voiture_service':
        domain.append(('service_car', '=', True))
    elif activite_filtre == 'production':
        domain += [('transport_logistics', '=', False), ('service_car', '=', False)]
    return domain


_TRANSPORT_CATS = {'CAMION TRACTEUR', 'CAMION ENGIN', 'SEMI-REMORQUE', 'PICK-UP', 'TRANSPORT PERSONNEL'}
_SERVICE_CATS   = {'VOITURE DE SERVICE', 'VOITURE DE FONCTION'}


def _fetch_sorties_bons(uid, models, domain, limit=1000):
    """Récupère et enrichit les bons de sortie depuis Odoo."""
    pickings = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'stock.picking', 'search_read',
        [domain],
        {
            'fields': [
                'name', 'scheduled_date', 'write_date',
                'partner_id', 'user_id', 'company_id', 'location_id',
                'picking_type_id', 'account_analytic_id', 'affectation_id',
                'equipment_id', 'initial_counter', 'actual_counter',
                'move_ids',
                'transport_logistics', 'service_car',   # champs booléens natifs Odoo
            ],
            'order': 'scheduled_date desc',
            'limit': limit,
        }
    )

    # Quantités gasoil par picking
    moves_qty = {}
    if pickings:
        all_move_ids = [mid for p in pickings for mid in p.get('move_ids', [])]
        if all_move_ids:
            moves = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.move', 'search_read',
                [[['id', 'in', all_move_ids],
                  ['product_id.categ_id', '=', CARBURANT_CATEG_ID]]],
                {'fields': ['id', 'picking_id', 'product_qty'], 'limit': 10000}
            )
            for m in moves:
                pid = m['picking_id'][0] if m['picking_id'] else None
                if pid:
                    moves_qty[pid] = moves_qty.get(pid, 0) + (m['product_qty'] or 0)

    bons = []
    for p in pickings:
        raw_date = (p.get('scheduled_date') or p.get('date') or p.get('write_date') or '')
        bon_date = raw_date[:10] if raw_date else '—'
        ecart    = (p.get('actual_counter') or 0) - (p.get('initial_counter') or 0)
        qty      = moves_qty.get(p['id'], 0)
        conso    = round(qty / ecart, 2) if ecart > 0 else 0
        is_anomalie = (ecart < 0) or (conso > 500 and conso > 0)

        engin_val = p.get('equipment_id')

        ouvrage_val = (
            p['account_analytic_id'][1] if isinstance(p.get('account_analytic_id'), list)
            else (p['affectation_id'][1] if p.get('affectation_id') else '')
            or extract_matricule(engin_val[1]) if engin_val
            else '—'
        )

        activity_bucket = _activity_bucket_from_picking(p, ouvrage_val)

        bons.append({
            'id':             p['id'],
            'date':           bon_date,
            'name':           p.get('name', '—'),
            'societe':        p['company_id'][1]           if p.get('company_id')           else '—',
            'chauffeur': (
                p['partner_id'][1] if p.get('partner_id')
                else p['user_id'][1] if p.get('user_id')
                else '—'
            ),
            'site':           p['location_id'][1]          if p.get('location_id')          else '—',
            'type_operation': p['picking_type_id'][1]      if p.get('picking_type_id')      else '—',
            'ouvrage':        ouvrage_val,
            'affectation':    p['affectation_id'][1]       if p.get('affectation_id')       else '—',
            'engin':          extract_matricule(engin_val[1]) if engin_val else '—',
            'categorie': (
                'Transport & Log.'  if activity_bucket == 'transport' else
                'Voiture de serv.'  if activity_bucket == 'voiture_service' else
                'Production'
            ),
            'is_transport':   activity_bucket == 'transport',
            'service_car':    p.get('service_car', False),
            'activite_bucket': activity_bucket,
            'cpt_initial':    p.get('initial_counter') or 0,
            'cpt_actuel':     p.get('actual_counter')  or 0,
            'ecart':          round(ecart, 1),
            'product_qty':    round(qty, 1),
            'consommation':   conso,
            'anomalie':       'Anomalie' if is_anomalie else 'OK',
        })
        _enrich_sortie_bon(bons[-1])
    return bons


def _sorties_pdf_response(bons, filters, total_litres, nb_anomalies, conso_moyenne):
    """Génère un PDF ReportLab A4 paysage — format SOMATRIN officiel."""
    import os
    from datetime import date as _date
    from reportlab.platypus import Image

    NAVY     = colors.HexColor('#1a2c4e')
    ORANGE   = colors.HexColor('#E87722')
    WHITE    = colors.white
    RED      = colors.HexColor('#dc2626')
    GREEN    = colors.HexColor('#16a34a')
    ROW_ALT  = colors.HexColor('#f4f6fb')
    ROW_ANOM = colors.HexColor('#fee2e2')
    GREY_TXT = colors.HexColor('#6b7280')
    BODY_TXT = colors.HexColor('#374151')
    today    = _date.today().strftime('%d/%m/%Y')

    # Chemin logo
    BASE_DIR  = settings.BASE_DIR
    LOGO_PATH = os.path.join(BASE_DIR, 'static', 'images', 'logo_somatrin.png')

    # ── Canvas numéroté ───────────────────────────────────────────────────────
    class _NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            rl_canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for state in self._saved:
                self.__dict__.update(state)
                self._draw_footer(total)
                rl_canvas.Canvas.showPage(self)
            rl_canvas.Canvas.save(self)

        def _draw_footer(self, total):
            pw = landscape(A4)[0]
            self.saveState()
            self.setFont('Helvetica', 7)
            self.setFillColor(GREY_TXT)
            self.drawString(18 * mm, 8 * mm, 'SOMATRIN — Document Confidentiel — Usage Interne')
            self.drawRightString(pw - 18 * mm, 8 * mm,
                                 f'Page {self._pageNumber} / {total}  |  {today}')
            self.restoreState()

    # ── Buffer & document ─────────────────────────────────────────────────────
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm,  bottomMargin=20 * mm,
    )

    PAGE_W = landscape(A4)[0] - 30 * mm  # largeur utile

    # ── En-tête ───────────────────────────────────────────────────────────────
    s_conf  = ParagraphStyle('conf',  fontName='Helvetica', fontSize=8,
                              textColor=GREY_TXT, alignment=TA_CENTER)
    s_title = ParagraphStyle('title', fontName='Helvetica-Bold', fontSize=14,
                              textColor=NAVY, alignment=TA_CENTER)
    s_sub   = ParagraphStyle('sub',   fontName='Helvetica', fontSize=9,
                              textColor=GREY_TXT, alignment=TA_CENTER)
    s_date  = ParagraphStyle('date',  fontName='Helvetica', fontSize=8,
                              textColor=GREY_TXT, alignment=TA_RIGHT)

    # Logo
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=28 * mm, height=11 * mm)
    else:
        logo = Paragraph('<b>SOMATRIN</b>',
                         ParagraphStyle('lg', fontName='Helvetica-Bold',
                                        fontSize=12, textColor=NAVY))

    # Filtres actifs pour sous-titre
    filtres = []
    if filters.get('date_debut'): filtres.append(f"Du {filters['date_debut']}")
    if filters.get('date_fin'):   filtres.append(f"au {filters['date_fin']}")
    if filters.get('site'):       filtres.append(f"Site : {filters['site']}")
    if filters.get('societe'):    filtres.append(f"Société : {filters['societe']}")
    sous_titre = '  |  '.join(filtres) if filtres else 'Toutes les données'

    col_w_hdr = [50 * mm, PAGE_W - 100 * mm, 50 * mm]

    hdr_data = [[
        logo,
        [Paragraph('Document Confidentiel — Usage Interne', s_conf),
         Spacer(1, 2 * mm),
         Paragraph('Rapport Sorties Gasoil', s_title),
         Spacer(1, 1 * mm),
         Paragraph(sous_titre, s_sub)],
        Paragraph(f'Page 1 / …<br/>{today}', s_date),
    ]]

    hdr_tbl = Table(hdr_data, colWidths=col_w_hdr)
    hdr_tbl.setStyle(TableStyle([
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN',        (0, 0), (0, 0),   'LEFT'),
        ('ALIGN',        (1, 0), (1, 0),   'CENTER'),
        ('ALIGN',        (2, 0), (2, 0),   'RIGHT'),
        ('LINEBELOW',    (0, 0), (-1, 0),  1.5, NAVY),
        ('TOPPADDING',   (0, 0), (-1, 0),  2),
        ('BOTTOMPADDING',(0, 0), (-1, 0),  6),
    ]))

    elems = [hdr_tbl, Spacer(1, 6 * mm)]

    # ── Tableau principal ─────────────────────────────────────────────────────
    COL_MM = [20, 25, 16, 20, 40, 32, 22, 30, 15, 15, 13, 15, 15, 13]
    col_w = [c * mm for c in COL_MM]

    HEADERS = ['Date', 'N° Bon', 'Société', 'Site', 'Ouvrage', 'Engin',
               'Catégorie', 'Chauffeur', 'Cpt. Init', 'Cpt. Act',
               'Écart', 'Qté (L)', 'Conso.', 'Statut']

    s_h  = ParagraphStyle('sh',  fontSize=8, textColor=WHITE,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_c  = ParagraphStyle('sc',  fontSize=7, textColor=BODY_TXT, fontName='Helvetica')
    s_cr = ParagraphStyle('scr', fontSize=7, textColor=BODY_TXT,
                           fontName='Helvetica', alignment=TA_RIGHT)
    s_cc = ParagraphStyle('scc', fontSize=7, textColor=BODY_TXT,
                           fontName='Helvetica', alignment=TA_CENTER)
    s_ok = ParagraphStyle('sok', fontSize=7, textColor=GREEN,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_an = ParagraphStyle('san', fontSize=7, textColor=RED,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_co = ParagraphStyle('sco', fontSize=7,
                           textColor=colors.HexColor('#0ea5e9'),
                           fontName='Helvetica-Bold', alignment=TA_RIGHT)

    def trunc(s, n):
        return (s[:n] + '…') if len(s) > n else s

    rows = [[Paragraph(h, s_h) for h in HEADERS]]
    for bon in bons:
        conso_s = format_number_decimals(bon['consommation'], 2) if bon.get('consommation') else '—'
        statut  = (Paragraph('OK', s_ok)
                   if bon['anomalie'] == 'OK'
                   else Paragraph('Anomalie', s_an))
        rows.append([
            Paragraph(bon['date'],                    s_cc),
            Paragraph(bon['name'],                    s_c),
            Paragraph(trunc(bon['societe'],    12),   s_c),
            Paragraph(trunc(bon['site'],       14),   s_c),
            Paragraph(trunc(bon['ouvrage'],    25),   s_c),
            Paragraph(trunc(bon['engin'],      20),   s_c),
            Paragraph(bon.get('categorie', '—'),      s_c),
            Paragraph(trunc(bon['chauffeur'],  15),   s_c),
            Paragraph(format_number_decimals(bon['cpt_initial'], 0),   s_cr),
            Paragraph(format_number_decimals(bon['cpt_actuel'], 0),    s_cr),
            Paragraph(format_number_decimals(bon['ecart'], 1),          s_cr),
            Paragraph(format_number_decimals(bon['product_qty'], 1),    s_cr),
            Paragraph(conso_s,                        s_co),
            statut,
        ])

    # Ligne TOTAL fond bleu
    total_qty = sum(b['product_qty'] for b in bons)
    s_tot  = ParagraphStyle('stot', fontSize=8, textColor=WHITE,
                             fontName='Helvetica-Bold')
    s_totq = ParagraphStyle('stotq', fontSize=8, textColor=WHITE,
                             fontName='Helvetica-Bold', alignment=TA_RIGHT)
    rows.append([
        Paragraph(f'TOTAL — {len(bons)} bon{"s" if len(bons) != 1 else ""}', s_tot),
        '', '', '', '', '', '', '', '', '', '',
        Paragraph(format_number_decimals(total_qty, 1), s_totq),
        '', '',
    ])

    n_rows = len(rows)
    style  = [
        ('BACKGROUND',    (0, 0),  (-1, 0),  NAVY),
        ('TEXTCOLOR',     (0, 0),  (-1, 0),  WHITE),
        ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0),  (-1, 0),  8),
        ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
        ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0),  (-1, 0),  5),
        ('BOTTOMPADDING', (0, 0),  (-1, 0),  5),
        ('FONTSIZE',      (0, 1),  (-1, -2), 7),
        ('TOPPADDING',    (0, 1),  (-1, -1), 3),
        ('BOTTOMPADDING', (0, 1),  (-1, -1), 3),
        ('LEFTPADDING',   (0, 0),  (-1, -1), 3),
        ('RIGHTPADDING',  (0, 0),  (-1, -1), 3),
        ('GRID',          (0, 0),  (-1, -2), 0.4, colors.HexColor('#d1d5db')),
        # Lignes alternées
        *[('BACKGROUND', (0, i), (-1, i), ROW_ALT)
          for i in range(2, n_rows - 1, 2)],
        # Ligne TOTAL
        ('BACKGROUND',    (0, -1), (-1, -1), NAVY),
        ('TEXTCOLOR',     (0, -1), (-1, -1), WHITE),
        ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0, -1), (-1, -1), 8),
        ('LINEABOVE',     (0, -1), (-1, -1), 1.5, NAVY),
        ('TOPPADDING',    (0, -1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 5),
        ('SPAN',          (0, -1), (10, -1)),
    ]

    # Anomalies en rouge
    for i, bon in enumerate(bons, start=1):
        if bon['anomalie'] == 'Anomalie':
            style.append(('BACKGROUND', (0, i), (-1, i), ROW_ANOM))

    main_tbl = Table(rows, colWidths=col_w, repeatRows=1)
    main_tbl.setStyle(TableStyle(style))
    elems.append(main_tbl)

    # ── Build PDF ─────────────────────────────────────────────────────────────
    doc.build(elems, canvasmaker=_NumberedCanvas)
    buffer.seek(0)

    fname_parts = ['sorties_gasoil']
    if filters.get('date_debut'): fname_parts.append(filters['date_debut'])
    if filters.get('date_fin'):   fname_parts.append(filters['date_fin'])
    filename = '_'.join(fname_parts) + '.pdf'

    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response

@login_required
def gasoil_sorties(request):
    date_debut      = request.GET.get('date_debut', '')
    date_fin        = request.GET.get('date_fin', '')
    societe         = request.GET.get('societe', '')
    site            = request.GET.get('site', '')
    chauffeur       = request.GET.get('chauffeur', '').strip()
    ouvrage         = request.GET.get('ouvrage', '').strip()
    anomalie        = request.GET.get('anomalie', '')
    categorie_engin = request.GET.get('categorie_engin', '')
    activite_filtre = request.GET.get('activite', '')
    export          = request.GET.get('export', '')
    page_number     = request.GET.get('page', 1)

    bons             = []
    categories_engin = []
    ouvrages_list    = []
    error            = None
    total_bons       = 0
    total_litres     = 0.0
    nb_anomalies     = 0
    conso_moyenne    = 0.0

    try:
        uid, models = get_odoo_connection()

        # Catégories engins
        try:
            cats = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.equipment.category', 'search_read',
                [[]],
                {'fields': ['name'], 'order': 'name asc', 'limit': 100}
            )
            categories_engin = [c['name'] for c in cats]
        except Exception:
            pass

        # Ouvrages / analytiques distincts (liste statique issue de account_analytic_id)
        ouvrages_list = [
            'Alimentation Station De Recyclage Mobile Par Pelle - GRABEMARO',
            'Alimentation station de lavage - LAFARGEHOLCIM MAROC',
            'Chargement Camions Agrégats Asment Béton & Tiers - GRABEMARO',
            'Chargement camions blocks Asment ciment - GRABEMARO',
            'Chargement camions pour concasseur - GRABEMARO',
            'Chargement des camions clients - LAFARGEHOLCIM MAROC',
            'Chargement et transport de la matière 1ère - LAFARGEHOLCIM MAROC',
            'Déstockage produits finis - LAFARGEHOLCIM MAROC',
            'Déstockage stérile au pré criblage - LAFARGEHOLCIM MAROC',
            'Déstockage, Alimentation Station De Lavage Sable - GRABEMARO',
            'Forage Minage Chargement et transport - LAFARGEHOLCIM MAROC',
            'Foration, chargement et tir des mines - GRABEMARO',
            'Mise en décharge - LAFARGEHOLCIM MAROC',
            'Ripage calcaire chargement et transport Zone A  vers HAZ - LAFARGEHOLCIM MAROC',
            'S00002-Arrosage de la piste d\'accès à la carrière (eau) - CIMENTS DU MAROC',
            'S00002-Décapage - CIMENTS DU MAROC',
            'S00002-Entretien de la piste d\'accès à la carrière - CIMENTS DU MAROC',
            'S00002-Extraction, Chargement, Transport & alimentation concasseur - CIMENTS DU MAROC',
            'S00002-Reprise de chargement argile concassée - CIMENTS DU MAROC',
            'S00004-Décapage sans tir - CIMENTS DU MAROC',
            'S00004-Foration, tir, chargement, transport et alimentation concasseur - CIMENTS DU MAROC',
            'S00006-Chargement camion blocs - ASMENT DU CENTRE',
            'S00006-Chargement des granulats - ASMENT DU CENTRE',
            'S00006-Foration et Abattage - ASMENT DU CENTRE',
            'S00006-Transport vers le concasseur - ASMENT DU CENTRE',
            'S00010-Chargement transport argile vers concasseur - LAFARGEHOLCIM MAROC',
            'S00010-Chargement transport calcaire vers concasseur - LAFARGEHOLCIM MAROC',
            'S00010-Foration, Minage carrière Calcaire - LAFARGEHOLCIM MAROC',
            'S00010-Manutention M.P ( Sable, minerai de fer) - LAFARGEHOLCIM MAROC',
            'S00010-Manutention M.P (Pouzzolane, Gypse) - LAFARGEHOLCIM MAROC',
            'S00010-Manutention interne : CIMENT - LAFARGEHOLCIM MAROC',
            'S00010-Ripage carrière Argile - LAFARGEHOLCIM MAROC',
            'S00011-Compensation prix Gasoil LH MEKNES - LAFARGEHOLCIM MAROC',
            'S00011-Forage minage chargement transport - zone D vers HAZ - LAFARGEHOLCIM MAROC',
            'S00011-Forage minage chargement transport -Zone 4 par minage vers HAZ - LAFARGEHOLCIM MAROC',
            'S00011-Location chargeuse 4m3 - LAFARGEHOLCIM MAROC',
            'S00011-Location engins pour manutention interne - LAFARGEHOLCIM MAROC',
            'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'S00011-Ripage calcaire chargement et transport Zone A  vers HAZ - LAFARGEHOLCIM MAROC',
            'S00011-Ripage calcaire chargement et transport Zone D vers HAZ - LAFARGEHOLCIM MAROC',
            'S00012-Décapage Par Explosif - GRABEMARO',
            'S00012-Fourniture Station D\'eau - GRABEMARO',
            'S00012/ Décapage par explosif - GRABEMARO',
            'S00013-LOCATION CAMION (HRS) - GRABEMARO',
            'S00013-LOCATION CHARGEUSE (HRS) - GRABEMARO',
            'S00013-LOCATION CITERNE D\'EAU - GRABEMARO',
            'S00013-LOCATION D\'ENGINS : PELLE - GRABEMARO',
            'S00044-Manutention des travaux - LAFARGEHOLCIM MAROC',
            'S00071-FOURNITURE DE SCHISTE LH SETTAT - LAFARGEHOLCIM MAROC',
            'S00073-ALIMENTATION CAMION 8*4 - SERVICE LOGISTIQUES -SOMATRIN',
            'S00073-CHARGEMENT PRODUIT FINI - SERVICE LOGISTIQUES -SOMATRIN',
            'S00073-Concassage - SERVICE LOGISTIQUES -SOMATRIN',
            'S00073-DECAPAGE - SERVICE LOGISTIQUES -SOMATRIN',
            'S00073-FORAGE - SERVICE LOGISTIQUES -SOMATRIN',
            'S00073-REDUCTION - SERVICE LOGISTIQUES -SOMATRIN',
            'Transport inter du front au concasseur - GRABEMARO',
        ]

        domain = _build_sorties_domain(
            date_debut, date_fin, site, chauffeur, ouvrage, anomalie,
            societe, categorie_engin, activite_filtre
        )
        limit = 2000 if (date_debut or date_fin) else 500
        bons  = _fetch_sorties_bons(uid, models, domain, limit=limit)

    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    # Filtre ouvrage post-fetch
    print("OUVRAGE FILTRE:", ouvrage)
    if bons:
        print("EXEMPLE OUVRAGE BON:", bons[0].get('ouvrage'))
    # Filtre Statut post-fetch : ok → bons sans anomalie ; anomalie → bons avec anomalie calculée
    if anomalie == 'ok':
        bons = [b for b in bons if b['anomalie'] == 'OK']
    elif anomalie == 'anomalie':
        bons = [b for b in bons if b['anomalie'] == 'Anomalie']

    # Calculs des totaux
    total_bons    = len(bons)
    total_litres  = sum(b['product_qty'] for b in bons)
    nb_anomalies  = sum(1 for b in bons if b['anomalie'] == 'Anomalie')
    conso_vals    = [b['consommation'] for b in bons if b['consommation'] > 0]
    conso_moyenne = round(sum(conso_vals) / len(conso_vals), 2) if conso_vals else 0

    # ── Filtre sélection (export personnalisé par IDs) ───────────────────────
    def _apply_ids_filter(lst):
        ids_param = request.GET.get('ids', '')
        if not ids_param:
            return lst
        id_set = {int(i) for i in ids_param.split(',') if i.strip().isdigit()}
        return [b for b in lst if b['id'] in id_set]

    # ── Export CSV ────────────────────────────────────────────────────────────
    if export == 'csv':
        from datetime import date as _csv_date
        export_bons = _apply_ids_filter(bons)
        response = HttpResponse(content_type='text/csv; charset=utf-8-sig')
        fname_parts = ['sorties_gasoil']
        if date_debut: fname_parts.append(date_debut)
        if date_fin:   fname_parts.append(date_fin)
        response['Content-Disposition'] = (
            f'attachment; filename="{"_".join(fname_parts)}.csv"'
        )
        writer = csv.writer(response, delimiter=';')

        # ── Lignes de métadonnées ─────────────────────────────────────────────
        writer.writerow([f'# SOMATRIN — Rapport Sorties Gasoil'])
        username = request.user.get_full_name() or request.user.username
        writer.writerow([f'# Généré par : {username} le '
                         f'{_csv_date.today().strftime("%d/%m/%Y")}'])
        filtres_actifs = []
        if date_debut:      filtres_actifs.append(f'Date début : {date_debut}')
        if date_fin:        filtres_actifs.append(f'Date fin : {date_fin}')
        if societe:         filtres_actifs.append(f'Société : {societe}')
        if site:            filtres_actifs.append(f'Site : {site}')
        if chauffeur:       filtres_actifs.append(f'Chauffeur : {chauffeur}')
        if ouvrage:         filtres_actifs.append(f'Ouvrage : {ouvrage}')
        if anomalie:
            _lbl = {'ok': 'OK', 'anomalie': 'Anomalie'}.get(anomalie, anomalie)
            filtres_actifs.append(f'Statut : {_lbl}')
        if activite_filtre: filtres_actifs.append(f'Activité : {activite_filtre}')
        writer.writerow([f'# Filtres : '
                         + (', '.join(filtres_actifs) if filtres_actifs else 'Aucun')])
        writer.writerow([])   # Ligne 4 vide

        # ── En-têtes et données ───────────────────────────────────────────────
        writer.writerow(['Date', 'N° Bon', 'Société', 'Site', 'Ouvrage',
                         'Engin', 'Catégorie', 'Chauffeur',
                         'Cpt. initial', 'Cpt. actuel', 'Écart (km)',
                         'Qté (L)', 'Conso. (L/h)', 'Statut'])
        for b in export_bons:
            writer.writerow([
                b['date'], b['name'], b['societe'], b['site'],
                b['ouvrage'], b['engin'], b.get('categorie', ''),
                b['chauffeur'],
                str(b['cpt_initial']).replace('.', ','),
                str(b['cpt_actuel']).replace('.', ','),
                str(b['ecart']).replace('.', ','),
                str(b['product_qty']).replace('.', ','),
                str(b['consommation']).replace('.', ',') if b['consommation'] else '',
                b['anomalie'],
            ])
        return response

    # ── Export PDF ReportLab ──────────────────────────────────────────────────
    if export == 'pdf':
        return _sorties_pdf_response(
            bons=_apply_ids_filter(bons),
            filters={
                'date_debut': date_debut, 'date_fin': date_fin,
                'societe': societe, 'site': site,
                'chauffeur': chauffeur, 'anomalie': anomalie,
                'activite_filtre': activite_filtre,
            },
            total_litres=round(total_litres, 1),
            nb_anomalies=nb_anomalies,
            conso_moyenne=conso_moyenne,
        )

    paginator = Paginator(bons, 50)
    page_obj  = paginator.get_page(page_number)

    from urllib.parse import urlencode
    export_params = {k: v for k, v in request.GET.items() if k != 'page' and k != 'export'}
    export_qs = urlencode(export_params)

    return render(request, 'gasoil/sorties.html', {
        'page_obj': page_obj, 'error': error,
        'date_debut': date_debut, 'date_fin': date_fin,
        'societe': societe, 'site': site,
        'chauffeur': chauffeur, 'ouvrage': ouvrage,
        'anomalie': anomalie,
        'categorie_engin':  categorie_engin,
        'activite_filtre':  activite_filtre,
        'categories_engin': categories_engin,
        'ouvrages_list':    ouvrages_list,
        'sites':            SITES_LIST,
        'total_bons':       total_bons,
        'total_litres':     round(total_litres, 1),
        'nb_anomalies':     nb_anomalies,
        'conso_moyenne':    conso_moyenne,
        'export_qs':        export_qs,
        'total_bons_fmt': format_number_decimals(total_bons, 0),
        'total_litres_fmt': format_number_decimals(total_litres, 0),
        'nb_anomalies_fmt': format_number_decimals(nb_anomalies, 0),
        'conso_moyenne_fmt': format_number_decimals(conso_moyenne, 2),
        'page_start_fmt': format_number_decimals(page_obj.start_index, 0),
        'page_end_fmt': format_number_decimals(page_obj.end_index, 0),
    })


@login_required
def gasoil_sorties_export(request):
    """Export Excel des bons de sortie gasoil (openpyxl)."""
    date_debut      = request.GET.get('date_debut', '')
    date_fin        = request.GET.get('date_fin', '')
    societe         = request.GET.get('societe', '')
    site            = request.GET.get('site', '')
    chauffeur       = request.GET.get('chauffeur', '').strip()
    ouvrage         = request.GET.get('ouvrage', '').strip()
    anomalie        = request.GET.get('anomalie', '')
    activite_filtre = request.GET.get('activite', '')

    try:
        uid, models = get_odoo_connection()
        domain = _build_sorties_domain(date_debut, date_fin, site, chauffeur, ouvrage,
                                       anomalie, societe, activite_filtre=activite_filtre)
        bons   = _fetch_sorties_bons(uid, models, domain, limit=5000)
    except Exception as e:
        return HttpResponse(f"Erreur Odoo : {e}", status=500)

    # ── Post-filtres (anomalie calculée + sélection par IDs) ─────────────────
    if anomalie == 'ok':
        bons = [b for b in bons if b['anomalie'] == 'OK']
    elif anomalie == 'anomalie':
        bons = [b for b in bons if b['anomalie'] == 'Anomalie']

    ids_param = request.GET.get('ids', '')
    if ids_param:
        id_set = {int(i) for i in ids_param.split(',') if i.strip().isdigit()}
        bons   = [b for b in bons if b['id'] in id_set]

    # ── Workbook ──────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sorties Gasoil"

    NAVY   = "1a2c4e"
    ORANGE = "E87722"
    LIGHT  = "F8F9FB"
    WHITE  = "FFFFFF"

    header_font  = Font(name="Calibri", bold=True, color=WHITE, size=11)
    header_fill  = PatternFill("solid", fgColor=NAVY)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    total_font  = Font(name="Calibri", bold=True, color=NAVY, size=11)
    total_fill  = PatternFill("solid", fgColor="EEF1F7")

    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Titre ─────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = "Rapport Sorties Gasoil"
    title_cell.font  = Font(name="Calibri", bold=True, size=14, color=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="EEF1F7")
    ws.row_dimensions[1].height = 30

    # Sous-titre filtre
    ws.merge_cells("A2:N2")
    subtitle = []
    if date_debut: subtitle.append(f"Du {date_debut}")
    if date_fin:   subtitle.append(f"au {date_fin}")
    if societe:    subtitle.append(f"Société : {societe}")
    if site:       subtitle.append(f"Site : {site}")
    if anomalie:
        _lbl_a = {'ok': 'OK', 'anomalie': 'Anomalie'}.get(anomalie, anomalie)
        subtitle.append(f"Statut : {_lbl_a}")
    ws["A2"].value = "  |  ".join(subtitle) if subtitle else "Toutes les données"
    ws["A2"].font  = Font(name="Calibri", italic=True, size=10, color="6B7280")
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    ws.row_dimensions[3].height = 6  # Espace

    # ── En-têtes ──────────────────────────────────────────────────────────────
    headers = [
        ("Date",          12),
        ("N° Bon",        18),
        ("Société",       14),
        ("Site",          20),
        ("Ouvrage",       28),
        ("Engin",         20),
        ("Chauffeur",     22),
        ("Cpt. initial",  14),
        ("Cpt. actuel",   14),
        ("Écart (km)",    13),
        ("Qté (L)",       12),
        ("Conso.",        12),
        ("Statut",        13),
    ]

    for col_idx, (hdr, width) in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=hdr)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = border
        ws.column_dimensions[cell.column_letter].width = width

    ws.row_dimensions[4].height = 24

    # ── Données ───────────────────────────────────────────────────────────────
    for row_idx, bon in enumerate(bons, start=5):
        is_even     = (row_idx % 2 == 0)
        row_fill    = PatternFill("solid", fgColor=LIGHT) if is_even else PatternFill("solid", fgColor=WHITE)
        anomal_fill = PatternFill("solid", fgColor="FEF2F2")

        use_fill = anomal_fill if bon['anomalie'] == 'Anomalie' else row_fill

        values = [
            bon['date'],
            bon['name'],
            bon['societe'],
            bon['site'],
            bon['ouvrage'],
            bon['engin'],
            bon['chauffeur'],
            bon['cpt_initial'],
            bon['cpt_actuel'],
            bon['ecart'],
            bon['product_qty'],
            bon['consommation'] if bon['consommation'] else '',
            bon['anomalie'],
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill   = use_fill
            cell.border = border
            cell.font   = Font(name="Calibri", size=10)
            # Alignements spécifiques
            if col_idx in (1,):
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in (9, 10, 11, 12, 13):
                cell.alignment = Alignment(horizontal="right")
                if isinstance(val, (int, float)) and val:
                    cell.number_format = '0.0'
            elif col_idx == 14:
                cell.alignment = Alignment(horizontal="center")
                if val == 'Anomalie':
                    cell.font = Font(name="Calibri", size=10, bold=True, color="B91C1C")
                else:
                    cell.font = Font(name="Calibri", size=10, bold=True, color="15803D")

        ws.row_dimensions[row_idx].height = 16

    # ── Ligne total ───────────────────────────────────────────────────────────
    total_row = len(bons) + 5
    ws.merge_cells(f"A{total_row}:G{total_row}")
    total_cell = ws.cell(row=total_row, column=1, value=f"TOTAL  —  {len(bons)} bon(s)")
    total_cell.font      = total_font
    total_cell.fill      = total_fill
    total_cell.alignment = Alignment(horizontal="right")
    total_cell.border    = border

    total_litres = sum(b['product_qty'] for b in bons)
    qty_cell = ws.cell(row=total_row, column=11, value=round(total_litres, 1))
    qty_cell.font = Font(name="Calibri", bold=True, size=11, color=NAVY)
    qty_cell.fill = total_fill
    qty_cell.alignment  = Alignment(horizontal="right")
    qty_cell.number_format = '0.0'
    qty_cell.border = border

    for col in [8, 9, 10, 12, 13]:
        c = ws.cell(row=total_row, column=col)
        c.fill   = total_fill
        c.border = border

    ws.freeze_panes = "A5"

    # ── Réponse HTTP ──────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename_parts = ["sorties_gasoil"]
    if date_debut: filename_parts.append(date_debut)
    if date_fin:   filename_parts.append(date_fin)
    filename = "_".join(filename_parts) + ".xlsx"

    response = HttpResponse(
        buf.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _entrees_pdf_response(bons, filters, total_litres, total_cout):
    """Génère un PDF ReportLab A4 paysage — Entrées Gasoil."""
    import os
    from datetime import date as _date
    from reportlab.platypus import Image

    NAVY     = colors.HexColor('#1a2c4e')
    WHITE    = colors.white
    GREEN    = colors.HexColor('#16a34a')
    ROW_ALT  = colors.HexColor('#f4f6fb')
    GREY_TXT = colors.HexColor('#6b7280')
    BODY_TXT = colors.HexColor('#374151')
    today    = _date.today().strftime('%d/%m/%Y')

    BASE_DIR  = settings.BASE_DIR
    LOGO_PATH = os.path.join(BASE_DIR, 'static', 'images', 'logo_somatrin.png')

    # ── Canvas numéroté ───────────────────────────────────────────────────────
    class _NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            rl_canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for state in self._saved:
                self.__dict__.update(state)
                self._draw_footer(total)
                rl_canvas.Canvas.showPage(self)
            rl_canvas.Canvas.save(self)

        def _draw_footer(self, total):
            pw = landscape(A4)[0]
            self.saveState()
            self.setFont('Helvetica', 7)
            self.setFillColor(GREY_TXT)
            self.drawString(15 * mm, 8 * mm, 'SOMATRIN — Document Confidentiel — Usage Interne')
            self.drawRightString(pw - 15 * mm, 8 * mm,
                                 f'Page {self._pageNumber} / {total}  |  {today}')
            self.restoreState()

    # ── Buffer & document ─────────────────────────────────────────────────────
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm,  bottomMargin=20 * mm,
    )

    PAGE_W = landscape(A4)[0] - 30 * mm

    # ── En-tête ───────────────────────────────────────────────────────────────
    s_conf  = ParagraphStyle('conf',  fontName='Helvetica', fontSize=8,
                              textColor=GREY_TXT, alignment=TA_CENTER)
    s_title = ParagraphStyle('title', fontName='Helvetica-Bold', fontSize=14,
                              textColor=NAVY, alignment=TA_CENTER)
    s_sub   = ParagraphStyle('sub',   fontName='Helvetica', fontSize=9,
                              textColor=GREY_TXT, alignment=TA_CENTER)
    s_date  = ParagraphStyle('date',  fontName='Helvetica', fontSize=8,
                              textColor=GREY_TXT, alignment=TA_RIGHT)

    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=28 * mm, height=11 * mm)
    else:
        logo = Paragraph('<b>SOMATRIN</b>',
                         ParagraphStyle('lg', fontName='Helvetica-Bold',
                                        fontSize=12, textColor=NAVY))

    filtres = []
    if filters.get('date_debut'): filtres.append(f"Du {filters['date_debut']}")
    if filters.get('date_fin'):   filtres.append(f"au {filters['date_fin']}")
    if filters.get('fournisseur'): filtres.append(f"Fournisseur : {filters['fournisseur']}")
    sous_titre = '  |  '.join(filtres) if filtres else 'Toutes les données'

    col_w_hdr = [50 * mm, PAGE_W - 100 * mm, 50 * mm]

    hdr_data = [[
        logo,
        [Paragraph('Document Confidentiel — Usage Interne', s_conf),
         Spacer(1, 2 * mm),
         Paragraph('Rapport Entrées Gasoil', s_title),
         Spacer(1, 1 * mm),
         Paragraph(sous_titre, s_sub)],
        Paragraph(f'{today}', s_date),
    ]]

    hdr_tbl = Table(hdr_data, colWidths=col_w_hdr)
    hdr_tbl.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN',         (0, 0), (0, 0),   'LEFT'),
        ('ALIGN',         (1, 0), (1, 0),   'CENTER'),
        ('ALIGN',         (2, 0), (2, 0),   'RIGHT'),
        ('LINEBELOW',     (0, 0), (-1, 0),  1.5, NAVY),
        ('TOPPADDING',    (0, 0), (-1, 0),  2),
        ('BOTTOMPADDING', (0, 0), (-1, 0),  6),
    ]))

    elems = [hdr_tbl, Spacer(1, 6 * mm)]

    # ── Tableau principal ─────────────────────────────────────────────────────
    COL_MM = [22, 35, 30, 40, 55, 20, 25, 30, 20]
    col_w  = [c * mm for c in COL_MM]

    HEADERS = ['Date', 'N° Facture', 'Réf. fourn.', 'Fournisseur',
               'Produit', 'Qté (L)', 'Prix unit. HT', 'Total HT (MAD)', 'Statut']

    s_h  = ParagraphStyle('sh',  fontSize=8, textColor=WHITE,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_c  = ParagraphStyle('sc',  fontSize=7, textColor=BODY_TXT, fontName='Helvetica')
    s_cr = ParagraphStyle('scr', fontSize=7, textColor=BODY_TXT,
                           fontName='Helvetica', alignment=TA_RIGHT)
    s_cc = ParagraphStyle('scc', fontSize=7, textColor=BODY_TXT,
                           fontName='Helvetica', alignment=TA_CENTER)
    s_ok = ParagraphStyle('sok', fontSize=7, textColor=GREEN,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)

    def trunc(s, n):
        return (s[:n] + '…') if len(str(s)) > n else str(s)

    rows = [[Paragraph(h, s_h) for h in HEADERS]]
    for bon in bons:
        rows.append([
            Paragraph(str(bon['date']),                    s_cc),
            Paragraph(trunc(bon['name'], 30),              s_c),
            Paragraph(trunc(bon.get('ref', ''), 25),       s_c),
            Paragraph(trunc(bon['fournisseur'], 35),       s_c),
            Paragraph(trunc(bon['product'], 45),           s_c),
            Paragraph(format_number_decimals(bon['product_qty'], 1),         s_cr),
            Paragraph(format_number_decimals(bon['price_unit'], 2),         s_cr),
            Paragraph(format_number_decimals(bon['total'], 2),              s_cr),
            Paragraph(bon.get('statut', 'Validé'), s_ok),
        ])

    # Ligne TOTAL
    s_tot  = ParagraphStyle('stot', fontSize=8, textColor=WHITE,
                             fontName='Helvetica-Bold')
    s_totq = ParagraphStyle('stotq', fontSize=8, textColor=WHITE,
                             fontName='Helvetica-Bold', alignment=TA_RIGHT)

    rows.append([
        Paragraph(f'TOTAL — {len(bons)} ligne{"s" if len(bons) != 1 else ""}', s_tot),
        '', '', '', '',
        Paragraph(format_number_decimals(total_litres, 1), s_totq),
        '',
        Paragraph(format_number_decimals(total_cout, 2), s_totq),
        '',
    ])

    n_rows = len(rows)
    style  = [
        ('BACKGROUND',    (0, 0),  (-1, 0),  NAVY),
        ('TEXTCOLOR',     (0, 0),  (-1, 0),  WHITE),
        ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0),  (-1, 0),  8),
        ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
        ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0),  (-1, 0),  5),
        ('BOTTOMPADDING', (0, 0),  (-1, 0),  5),
        ('FONTSIZE',      (0, 1),  (-1, -2), 7),
        ('TOPPADDING',    (0, 1),  (-1, -1), 3),
        ('BOTTOMPADDING', (0, 1), (-1, -1),  3),
        ('LEFTPADDING',   (0, 0),  (-1, -1), 3),
        ('RIGHTPADDING',  (0, 0),  (-1, -1), 3),
        ('GRID',          (0, 0),  (-1, -2), 0.4, colors.HexColor('#d1d5db')),
        *[('BACKGROUND', (0, i), (-1, i), ROW_ALT)
          for i in range(2, n_rows - 1, 2)],
        ('BACKGROUND',    (0, -1), (-1, -1), NAVY),
        ('TEXTCOLOR',     (0, -1), (-1, -1), WHITE),
        ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0, -1), (-1, -1), 8),
        ('LINEABOVE',     (0, -1), (-1, -1), 1.5, NAVY),
        ('TOPPADDING',    (0, -1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 5),
        ('SPAN',          (0, -1), (4, -1)),
    ]

    main_tbl = Table(rows, colWidths=col_w, repeatRows=1)
    main_tbl.setStyle(TableStyle(style))
    elems.append(main_tbl)

    # ── Build PDF ─────────────────────────────────────────────────────────────
    doc.build(elems, canvasmaker=_NumberedCanvas)
    buffer.seek(0)

    fname_parts = ['entrees_gasoil']
    if filters.get('date_debut'): fname_parts.append(filters['date_debut'])
    if filters.get('date_fin'):   fname_parts.append(filters['date_fin'])
    filename = '_'.join(fname_parts) + '.pdf'

    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


# ─────────────────────────────────────────────
#  GASOIL — ENTRÉES
#  Modèle : stock.picking (réceptions de carburant)
#  Type : incoming (fournisseur → stock)
# ─────────────────────────────────────────────
@login_required
def gasoil_entrees(request):
    date_debut      = request.GET.get('date_debut', '')
    date_fin        = request.GET.get('date_fin', '')
    site            = request.GET.get('site', '')
    fournisseur     = request.GET.get('fournisseur', '').strip()
    activite_filtre = request.GET.get('activite', '')
    export          = request.GET.get('export', '')

    bons  = []
    error = None

    try:
        uid, models = get_odoo_connection()

        # ── 1. Lignes de factures fournisseurs contenant "GASOIL" ──
        line_domain = [
            ('move_id.move_type', '=', 'in_invoice'),
            ('move_id.state', '=', 'posted'),
            ('display_type', '=', 'product'),  # lignes produit uniquement (pas tax/payment_term)
            '|',
            ('product_id.name', 'ilike', 'gasoil'),
            ('name', 'ilike', 'gasoil'),
        ]
        if date_debut:
            line_domain.append(('move_id.invoice_date', '>=', date_debut))
        if date_fin:
            line_domain.append(('move_id.invoice_date', '<=', date_fin))
        if fournisseur:
            line_domain.append(('move_id.partner_id.name', 'ilike', fournisseur))

        lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move.line', 'search_read',
            [line_domain],
            {
                'fields': ['move_id', 'product_id', 'name',
                           'quantity', 'price_unit', 'price_subtotal'],
                'order': 'move_id desc',
                'limit': 5000,
            }
        )

        # ── 2. En-têtes des factures (date, fournisseur, ref) ──
        move_ids = list({l['move_id'][0] for l in lines if l.get('move_id')})
        move_map = {}
        if move_ids:
            invoices = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'account.move', 'search_read',
                [[['id', 'in', move_ids]]],
                {'fields': ['id', 'name', 'invoice_date', 'partner_id', 'ref', 'state'],
                 'limit': len(move_ids)}
            )
            move_map = {m['id']: m for m in invoices}

        # ── 3. Construction des bons ──
        for line in lines:
            mid = line['move_id'][0] if line.get('move_id') else None
            mv  = move_map.get(mid, {})
            qty = line.get('quantity') or 0
            pu  = line.get('price_unit') or 0
            bons.append({
                'id':          line['id'],
                'date':        mv.get('invoice_date', '—') or '—',
                'name':        mv.get('name', '—') or '—',
                'ref':         mv.get('ref', '') or '',
                'fournisseur': mv['partner_id'][1] if mv.get('partner_id') else '—',
                'product':     (line['product_id'][1] if line.get('product_id')
                                else line.get('name', '—')),
                'description': line.get('name', ''),
                'product_qty': round(qty, 1),
                'price_unit':  round(pu, 2),
                'total':       round(line.get('price_subtotal') or qty * pu, 2),
                'statut': {
                    'draft': 'Brouillon',
                    'posted': 'Validé',
                    'cancel': 'Annulé'
                }.get(mv.get('state'), mv.get('state', '—') or '—'),
            })

    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    # ── Filtre activité post-fetch (les factures n'ont pas transport_logistics) ──
    # On filtre sur le nom du produit : GASOIL 10PPM = carburant pur (production/transport)
    # Citerne / Filtre à gasoil = pièces (traités comme production par défaut)
    if activite_filtre == 'transport':
        bons = [b for b in bons if 'PPM' in b.get('product', '').upper()
                                or 'CARBURANT' in b.get('product', '').upper()]
    elif activite_filtre == 'voiture_service':
        bons = [b for b in bons if 'SERVICE' in b.get('fournisseur', '').upper()
                                or 'VOITURE' in b.get('product', '').upper()]
    # 'production' : pas de filtre supplémentaire — tout ce qui n'est pas transport

    total_bons   = len(bons)
    total_litres = sum(b['product_qty'] for b in bons)
    total_cout   = sum(b['total'] for b in bons)

    # ── Export CSV ────────────────────────────────────────────────────────────
    if export == 'csv':
        response = HttpResponse(content_type='text/csv; charset=utf-8-sig')
        fname_parts = ['entrees_gasoil']
        if date_debut: fname_parts.append(date_debut)
        if date_fin:   fname_parts.append(date_fin)
        response['Content-Disposition'] = (
            f'attachment; filename="{"_".join(fname_parts)}.csv"'
        )
        writer = csv.writer(response, delimiter=';')
        writer.writerow(['Date facture', 'N° Facture', 'Réf. fournisseur',
                         'Fournisseur', 'Produit', 'Quantité (L)',
                         'Prix unit. HT', 'Total HT (MAD)', 'Statut'])
        for b in bons:
            writer.writerow([
                b['date'], b['name'], b.get('ref', ''), b['fournisseur'],
                b['product'],
                str(b['product_qty']).replace('.', ','),
                str(b['price_unit']).replace('.', ','),
                str(b['total']).replace('.', ','),
                b.get('statut', '—'),
            ])
        return response

    # ── Export PDF (page impression A4) ──────────────────────────────────────
    if export == 'pdf':
        return _entrees_pdf_response(
            bons=bons,
            filters={
                'date_debut': date_debut,
                'date_fin': date_fin,
                'fournisseur': fournisseur,
            },
            total_litres=round(total_litres, 1),
            total_cout=round(total_cout, 2),
        )

    return render(request, 'gasoil/entrees.html', {
        'bons':            bons,
        'error':           error,
        'date_debut':      date_debut,
        'date_fin':        date_fin,
        'site':            site,
        'fournisseur':     fournisseur,
        'activite_filtre': activite_filtre,
        'total_bons':      total_bons,
        'total_litres':    round(total_litres, 1),
        'total_cout':      round(total_cout, 2),
        'now':            date.today(),
    })


# ─────────────────────────────────────────────
#  GASOIL — BILAN
# ─────────────────────────────────────────────
@login_required
def gasoil_bilan(request):
    annee = request.GET.get('annee', '')
    mois = request.GET.get('mois', '')
    site  = request.GET.get('site', '')
    activite_filtre = request.GET.get('activite_filtre', '').strip()
    anomalie_seulement = request.GET.get('anomalie_seulement', '').strip()
    societe = request.GET.get('societe', '').strip()
    engin = request.GET.get('engin', '').strip()

    error        = None
    entrees_data = []
    sorties_data = []
    entrees_all = []
    sorties_all = []

    def _search_read_all(model_name, domain, fields, batch_size=2000):
        """Read all matching records using paginated search + read."""
        all_rows = []
        offset = 0
        while True:
            ids = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                model_name, 'search',
                [domain],
                {'offset': offset, 'limit': batch_size, 'order': 'id desc'}
            )
            if not ids:
                break
            rows = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                model_name, 'read',
                [ids],
                {'fields': fields}
            )
            all_rows.extend(rows)
            if len(ids) < batch_size:
                break
            offset += batch_size
        return all_rows

    def _read_moves_chunked(move_ids, fields, batch_size=2000):
        """Read stock.move records in chunks to avoid truncation/limits."""
        rows = []
        for i in range(0, len(move_ids), batch_size):
            chunk = move_ids[i:i + batch_size]
            part = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.move', 'search_read',
                [[['id', 'in', chunk], ['product_id.categ_id', '=', CARBURANT_CATEG_ID]]],
                {'fields': fields, 'limit': batch_size}
            )
            rows.extend(part)
        return rows

    try:
        uid, models = get_odoo_connection()

        date_filter = []

        # ── Sorties ──
        domain_s = [
            ('picking_type_consumption', '=', True),
            ('state', '=', 'done'),
            ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
        ] + date_filter
        if site:
            domain_s.append(('location_id.complete_name', 'ilike', site))
        if societe:
            domain_s.append(('company_id.name', 'ilike', societe))
        if activite_filtre == 'transport':
            domain_s.append(('transport_logistics', '=', True))
        elif activite_filtre == 'voiture_service':
            domain_s.append(('service_car', '=', True))
        elif activite_filtre == 'production':
            domain_s += [('transport_logistics', '=', False), ('service_car', '=', False)]
        if anomalie_seulement == '1':
            domain_s.append(('picking_type_is_hors_affectation', '=', True))
        if engin:
            domain_s.append(('equipment_id.name', 'ilike', engin))

        pickings_s = _search_read_all(
            'stock.picking',
            domain_s,
            ['name', 'scheduled_date', 'date', 'write_date',
             'location_id', 'company_id', 'move_ids', 'equipment_id',
             'transport_logistics', 'service_car',
             'picking_type_is_hors_affectation'],
        )

        # Quantités sorties
        if pickings_s:
            all_ids = [mid for p in pickings_s for mid in p.get('move_ids', [])]
            qty_map = {}
            if all_ids:
                mvs = _read_moves_chunked(all_ids, ['picking_id', 'product_qty'])
                for m in mvs:
                    pid = m['picking_id'][0] if m['picking_id'] else None
                    if pid:
                        qty_map[pid] = qty_map.get(pid, 0) + (m['product_qty'] or 0)

            for p in pickings_s:
                raw = (p.get('scheduled_date') or p.get('date') or p.get('write_date') or '')
                # Catégorie via booléens natifs Odoo
                if p.get('transport_logistics'):
                    categorie = 'Transport & Logistique'
                elif p.get('service_car'):
                    categorie = 'Voiture de service'
                else:
                    categorie = 'Production'
                engin_val = p.get('equipment_id')
                sorties_data.append({
                    'date':     raw[:10] if raw else '—',
                    'site':     p['location_id'][1] if p.get('location_id') else '—',
                    'societe':  p['company_id'][1] if p.get('company_id') else '—',
                    'qty':      qty_map.get(p['id'], 0),
                    'anomalie': p.get('picking_type_is_hors_affectation', False),
                    'name':     p.get('name', '—'),
                    'engin':    engin_val[1] if engin_val else 'Inconnu',
                    'categorie': categorie,
                })

        # ── Entrées ──
        domain_e = [
            ('state', '=', 'done'),
            ('picking_type_id.code', '=', 'incoming'),
            ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
        ] + date_filter
        if site:
            domain_e.append(('location_dest_id.complete_name', 'ilike', site))
        if societe:
            domain_e.append(('company_id.name', 'ilike', societe))

        pickings_e = _search_read_all(
            'stock.picking',
            domain_e,
            ['scheduled_date', 'date', 'write_date',
             'location_dest_id', 'company_id', 'move_ids'],
        )

        if pickings_e:
            all_ids_e = [mid for p in pickings_e for mid in p.get('move_ids', [])]
            qty_map_e = {}
            pu_map_e  = {}
            if all_ids_e:
                mvs_e = _read_moves_chunked(all_ids_e, ['picking_id', 'product_qty', 'price_unit'])
                for m in mvs_e:
                    pid = m['picking_id'][0] if m['picking_id'] else None
                    if pid:
                        qty_map_e[pid] = qty_map_e.get(pid, 0) + (m['product_qty'] or 0)
                        pu_map_e[pid]  = m.get('price_unit') or 0

            for p in pickings_e:
                qty = qty_map_e.get(p['id'], 0)
                pu  = pu_map_e.get(p['id'], 0)
                raw = (p.get('scheduled_date') or p.get('date') or p.get('write_date') or '')
                entrees_data.append({
                    'date': raw[:10] if raw else '—',
                    'societe': p['company_id'][1] if p.get('company_id') else '—',
                    'qty':  qty,
                    'cout': qty * pu,
                })

        # Conserver une copie brute (avant filtres période) pour diagnostics/listes.
        sorties_all = list(sorties_data)
        entrees_all = list(entrees_data)

        # Filtrage de période robuste basé sur la date réellement disponible
        # (scheduled_date ou date ou write_date déjà normalisée en YYYY-MM-DD).
        if annee:
            sorties_data = [s for s in sorties_data if (s.get('date') or '').startswith(f'{annee}-')]
            entrees_data = [e for e in entrees_data if (e.get('date') or '').startswith(f'{annee}-')]
            if mois and len(mois) == 2 and mois.isdigit():
                needle = f'{annee}-{mois}'
                sorties_data = [s for s in sorties_data if (s.get('date') or '').startswith(needle)]
                entrees_data = [e for e in entrees_data if (e.get('date') or '').startswith(needle)]

    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    # ── KPI ──
    total_entrees = sum(e['qty'] for e in entrees_data)
    total_sorties = sum(s['qty'] for s in sorties_data)
    stock_estime  = total_entrees - total_sorties
    total_cout    = sum(e['cout'] for e in entrees_data)
    nb_anomalies  = sum(1 for s in sorties_data if s['anomalie'])

    # ── Graphique mensuel ──
    mois_labels       = ['Jan','Fév','Mar','Avr','Mai','Juin',
                         'Juil','Août','Sep','Oct','Nov','Déc']
    sorties_par_mois  = defaultdict(float)
    entrees_par_mois  = defaultdict(float)

    for s in sorties_data:
        if s['date'] and len(s['date']) >= 7:
            sorties_par_mois[int(s['date'][5:7])] += s['qty']
    for e in entrees_data:
        if e['date'] and len(e['date']) >= 7:
            entrees_par_mois[int(e['date'][5:7])] += e['qty']

    chart_labels       = json.dumps(mois_labels)
    chart_sorties_data = json.dumps([round(sorties_par_mois.get(m, 0), 1) for m in range(1, 13)])
    chart_entrees_data = json.dumps([round(entrees_par_mois.get(m, 0), 1) for m in range(1, 13)])

    # ── Répartition par site ──
    site_data = defaultdict(float)
    for s in sorties_data:
        site_data[s['site']] += s['qty']

    site_labels = json.dumps(list(site_data.keys()))
    site_values = json.dumps([round(v, 1) for v in site_data.values()])

    recap_sites = [
        {
            'site':   s,
            'litres': round(v, 1),
            'pct':    round(v / total_sorties * 100, 1) if total_sorties else 0,
        }
        for s, v in sorted(site_data.items(), key=lambda x: -x[1])
    ]

    # ── 1) Répartition par catégorie d'engin ─────────────────────────────────
    cat_data = defaultdict(float)
    for s in sorties_data:
        cat_data[s['categorie']] += s['qty']
    cat_sorted = sorted(cat_data.items(), key=lambda x: -x[1])[:8]
    categories_labels = json.dumps([c[0] for c in cat_sorted])
    categories_values = json.dumps([round(c[1], 1) for c in cat_sorted])

    # ── 2) Consommation par semaine ───────────────────────────────────────────
    semaines_data = defaultdict(float)
    for s in sorties_data:
        d = s['date']
        if d and len(d) == 10:
            try:
                from datetime import datetime
                dt = datetime.strptime(d, '%Y-%m-%d')
                # Clé = lundi de la semaine ISO (YYYY-Www)
                iso = dt.isocalendar()
                semaine_key = f"{iso[0]}-S{iso[1]:02d}"
                semaine_label = dt.strftime('%d/%m')
                semaines_data[semaine_key] = semaines_data.get(semaine_key, 0) + s['qty']
                semaines_data[f'_lbl_{semaine_key}'] = semaine_label
            except ValueError:
                pass
    # Trier par clé (chronologique) et séparer labels/valeurs
    real_keys = sorted(k for k in semaines_data if not k.startswith('_lbl_'))
    semaines_labels = json.dumps([semaines_data.get(f'_lbl_{k}', k) for k in real_keys])
    semaines_values = json.dumps([round(semaines_data[k], 1) for k in real_keys])

    # ── 3) Top 10 bons de sortie ──────────────────────────────────────────────
    bons_data = defaultdict(float)
    for s in sorties_data:
        bons_data[s['name']] += s['qty']
    bons_sorted = sorted(bons_data.items(), key=lambda x: -x[1])[:10]
    bons_labels = json.dumps([b[0] for b in bons_sorted])
    bons_values = json.dumps([round(b[1], 1) for b in bons_sorted])

    # ── 4) Top 10 matricules ──────────────────────────────────────────────────
    mat_data = defaultdict(float)
    for s in sorties_data:
        mat_data[s['engin']] += s['qty']
    mat_sorted = sorted(mat_data.items(), key=lambda x: -x[1])[:10]
    max_mat = mat_sorted[0][1] if mat_sorted else 1
    matricules_data = [
        {
            'nom':   m[0],
            'total': round(m[1], 1),
            'pct':   round(m[1] / max_mat * 100, 1),
        }
        for m in mat_sorted
    ]

    # ── 5) Équipements actifs ─────────────────────────────────────────────────
    equip_map = defaultdict(lambda: {'categorie': '—', 'nb_sorties': 0, 'total_litres': 0.0})
    for s in sorties_data:
        e = s['engin']
        equip_map[e]['categorie']   = s['categorie']
        equip_map[e]['nb_sorties'] += 1
        equip_map[e]['total_litres'] += s['qty']
    equipements_actifs = sorted(
        [{'matricule': k, **v, 'total_litres': round(v['total_litres'], 1)}
         for k, v in equip_map.items()],
        key=lambda x: -x['total_litres']
    )[:20]
    nb_equipements_actifs = len(equip_map)
    societes_list = sorted({s.get('societe') for s in sorties_all if s.get('societe') and s.get('societe') != '—'})
    engins_list = sorted({s.get('engin') for s in sorties_all if s.get('engin') and s.get('engin') != 'Inconnu'})[:200]
    years_detected = sorted({
        d[:4]
        for d in [*(s.get('date', '') for s in sorties_all), *(e.get('date', '') for e in entrees_all)]
        if isinstance(d, str) and len(d) >= 4 and d[:4].isdigit()
    })
    # N'afficher que les années réellement présentes dans les données Odoo.
    if not years_detected:
        years_detected = [str(date.today().year)]
    no_data_for_filters = (
        (annee or mois or site or societe or activite_filtre or engin or anomalie_seulement)
        and total_entrees == 0
        and total_sorties == 0
    )
    month_names = {
        '01': 'Janvier', '02': 'Février', '03': 'Mars', '04': 'Avril',
        '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Août',
        '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Décembre',
    }
    months_detected_for_year = set()
    if annee:
        for d in [*(s.get('date', '') for s in sorties_all), *(e.get('date', '') for e in entrees_all)]:
            if isinstance(d, str) and len(d) >= 7 and d[:4] == annee:
                mm = d[5:7]
                if mm in month_names:
                    months_detected_for_year.add(mm)
    else:
        for d in [*(s.get('date', '') for s in sorties_all), *(e.get('date', '') for e in entrees_all)]:
            if isinstance(d, str) and len(d) >= 7:
                mm = d[5:7]
                if mm in month_names:
                    months_detected_for_year.add(mm)
    mois_list_dynamic = [(m, month_names[m]) for m in sorted(months_detected_for_year)]
    if not mois_list_dynamic:
        mois_list_dynamic = [
            ('01', 'Janvier'), ('02', 'Février'), ('03', 'Mars'), ('04', 'Avril'),
            ('05', 'Mai'), ('06', 'Juin'), ('07', 'Juillet'), ('08', 'Août'),
            ('09', 'Septembre'), ('10', 'Octobre'), ('11', 'Novembre'), ('12', 'Décembre'),
        ]
    filtred_rows_count = len(sorties_data) + len(entrees_data)

    return render(request, 'gasoil/bilan.html', {
        'error': error,
        'annee': annee, 'mois': mois, 'site': site,
        'activite_filtre': activite_filtre,
        'anomalie_seulement': anomalie_seulement,
        'societe': societe,
        'engin': engin,
        'total_entrees': round(total_entrees, 1),
        'total_sorties': round(total_sorties, 1),
        'stock_estime':  round(stock_estime, 1),
        'total_cout':    round(total_cout, 2),
        'nb_anomalies':  nb_anomalies,
        'chart_labels':       chart_labels,
        'chart_sorties_data': chart_sorties_data,
        'chart_entrees_data': chart_entrees_data,
        'site_labels':        site_labels,
        'site_values':        site_values,
        'recap_sites':        recap_sites,
        'annees_list':        years_detected,
        'mois_list': mois_list_dynamic,
        'sites':              SITES_LIST,
        'societes_list': societes_list,
        'engins_list': engins_list,
        'no_data_for_filters': no_data_for_filters,
        'filtred_rows_count': filtred_rows_count,
        # Nouvelles sections analytiques
        'categories_labels_json': categories_labels,
        'categories_values_json': categories_values,
        'semaines_labels_json':   semaines_labels,
        'semaines_values_json':   semaines_values,
        'bons_labels_json':       bons_labels,
        'bons_values_json':       bons_values,
        'matricules_data':        matricules_data,
        'equipements_actifs':     equipements_actifs,
        'nb_equipements_actifs':  nb_equipements_actifs,
    })


# ─────────────────────────────────────────────
#  TRANSPORT & LOGISTIQUE
# ─────────────────────────────────────────────
def _transport_domain_dates(date_debut, date_fin, field_name):
    domain = []
    if date_debut:
        domain.append((field_name, '>=', f'{date_debut} 00:00:00'))
    if date_fin:
        domain.append((field_name, '<=', f'{date_fin} 23:59:59'))
    return domain


@login_required
def transport_bons(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    reference = request.GET.get('reference', '').strip()
    partenaire = request.GET.get('partenaire', '').strip()
    site = request.GET.get('site', '').strip()

    bons = []
    error = None
    try:
        uid, models = get_odoo_connection()
        domain = [('state', 'in', ['done', 'assigned'])]
        domain += _transport_domain_dates(date_debut, date_fin, 'scheduled_date')
        if reference:
            domain.append(('name', 'ilike', reference))
        if partenaire:
            domain.append(('partner_id.name', 'ilike', partenaire))
        if site:
            domain.append(('location_id.complete_name', 'ilike', site))

        try:
            domain_with_flag = domain + [('transport_logistics', '=', True)]
            records = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.picking', 'search_read',
                [domain_with_flag],
                {
                    'fields': [
                        'name', 'scheduled_date', 'state', 'partner_id',
                        'location_id', 'location_dest_id', 'origin',
                    ],
                    'order': 'scheduled_date desc',
                    'limit': 500,
                }
            )
        except Exception:
            records = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.picking', 'search_read',
                [domain],
                {
                    'fields': [
                        'name', 'scheduled_date', 'state', 'partner_id',
                        'location_id', 'location_dest_id', 'origin',
                    ],
                    'order': 'scheduled_date desc',
                    'limit': 500,
                }
            )

        today_iso = date.today().isoformat()
        for rec in records:
            origine = (
                rec.get('origin')
                or (rec['location_dest_id'][1] if rec.get('location_dest_id') else '')
                or '—'
            )
            st = (rec.get('state') or '').strip()
            row_date = (rec.get('scheduled_date') or '')[:10] or '—'
            bons.append({
                'date': row_date,
                'reference': rec.get('name') or '—',
                'partenaire': rec['partner_id'][1] if rec.get('partner_id') else '—',
                'site': rec['location_id'][1] if rec.get('location_id') else '—',
                'origine': origine,
                'etat': st or '—',
                'state_raw': st,
                'is_today': row_date == today_iso,
            })
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_bons = len(bons)
    bons_valides = sum(1 for r in bons if r.get('state_raw') == 'done')
    nb_sites = len({
        r.get('site')
        for r in bons
        if r.get('site') and r.get('site') != '—'
    })
    bons_aujourdhui = sum(1 for r in bons if r.get('is_today'))

    return render(request, 'transport/bons_transport.html', {
        'rows': bons,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'reference': reference,
        'partenaire': partenaire,
        'site': site,
        'total_rows': total_bons,
        'total_bons': total_bons,
        'bons_valides': bons_valides,
        'nb_sites': nb_sites,
        'bons_aujourdhui': bons_aujourdhui,
        'soma_page_data': {
            'page': 'transport_bons',
            'kpis': {
                'total_bons': total_bons,
                'bons_valides': bons_valides,
                'nb_sites': nb_sites,
                'bons_aujourdhui': bons_aujourdhui,
                'periode_debut': date_debut or None,
                'periode_fin': date_fin or None,
            },
        },
    })


@login_required
def transport_gasoil(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    vehicule = request.GET.get('vehicule', '').strip()
    conducteur_id = request.GET.get('conducteur_id', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    vehicules = []
    conducteurs = []
    error = None
    try:
        uid, models = get_odoo_connection()

        # Listes déroulantes (distinctes) chargées au chargement de la page.
        list_domain = _build_sorties_domain(
            date_debut='',
            date_fin='',
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='transport',
        )
        pickings_for_filters = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.picking', 'search_read',
            [list_domain],
            {
                'fields': ['equipment_id', 'partner_id', 'user_id'],
                'limit': 5000,
            }
        )
        vehicule_set = set()
        conducteur_map = {}
        for p in pickings_for_filters:
            if p.get('equipment_id'):
                vehicule_set.add(extract_matricule(p['equipment_id'][1]))
            if p.get('partner_id'):
                conducteur_map[p['partner_id'][0]] = p['partner_id'][1]
            elif p.get('user_id'):
                conducteur_map[p['user_id'][0]] = p['user_id'][1]
        vehicules = [
            {
                'value': v,
                'label': extract_matricule(v),
            }
            for v in sorted(vehicule_set)
        ]
        conducteurs = [
            {'id': cid, 'name': cname}
            for cid, cname in sorted(conducteur_map.items(), key=lambda item: item[1].lower())
        ]

        # Reprise de la logique du module gasoil existant:
        # stock.picking + stock.move (catégorie carburant) avec activité transport.
        domain = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='transport',
        )
        if conducteur_id.isdigit():
            cid = int(conducteur_id)
            domain += ['|', ('partner_id', '=', cid), ('user_id', '=', cid)]

        limit = 2000 if (date_debut or date_fin) else 500
        bons = _fetch_sorties_bons(uid, models, domain, limit=limit)

        # Montant approximé depuis stock.move (qty * price_unit) par picking.
        move_domain = [
            ['picking_id', 'in', [b['id'] for b in bons]],
            ['product_id.categ_id', '=', CARBURANT_CATEG_ID],
        ]
        move_lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.move', 'search_read',
            [move_domain],
            {
                'fields': ['picking_id', 'product_qty', 'price_total', 'unit_price', 'partner_id'],
                'limit': 10000,
            }
        )

        amount_by_picking = defaultdict(float)
        driver_by_picking = {}
        for mv in move_lines:
            pid = mv['picking_id'][0] if mv.get('picking_id') else None
            if not pid:
                continue
            price_total = mv.get('price_total')
            if price_total not in (None, False):
                amount_by_picking[pid] += price_total or 0
            else:
                qty = mv.get('product_qty') or 0
                pu = mv.get('unit_price') or 0
                amount_by_picking[pid] += qty * pu
            if not driver_by_picking.get(pid) and mv.get('partner_id'):
                driver_by_picking[pid] = mv['partner_id'][1]

        # Garantit que la page Transport n'affiche que le périmètre transport/logistique.
        bons = [b for b in bons if b.get('activite_bucket') == 'transport']

        if vehicule:
            bons = [b for b in bons if (b.get('engin') or '') == vehicule]

        for bon in bons:
            vehicle_name = bon.get('engin') or '—'
            vehicle_display = extract_matricule(vehicle_name)
            rows.append({
                'date': bon.get('date') or '—',
                'vehicule': vehicle_name,
                'vehicule_display': vehicle_display,
                'conducteur': bon.get('chauffeur') or driver_by_picking.get(bon['id']) or '—',
                'litres': round(bon.get('product_qty') or 0, 2),
                'montant': round(amount_by_picking.get(bon['id'], 0), 2),
                'compteur': bon.get('cpt_actuel') or 0,
            })
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_litres = round(sum(r['litres'] for r in rows), 2)
    total_montant = round(sum(r['montant'] for r in rows), 2)
    total_rows = len(rows)

    grouped_rows = []
    group_by_columns = []
    if group_by:
        groups = {}
        for row in rows:
            if group_by == 'mois':
                key = (row['date'][:7] if row.get('date') and row['date'] != '—' else '—',)
            elif group_by == 'vehicule':
                key = (row['vehicule_display'],)
            elif group_by == 'conducteur':
                key = (row['conducteur'],)
            elif group_by == 'mois_vehicule':
                key = (
                    row['date'][:7] if row.get('date') and row['date'] != '—' else '—',
                    row['vehicule_display'],
                )
            elif group_by == 'mois_conducteur':
                key = (
                    row['date'][:7] if row.get('date') and row['date'] != '—' else '—',
                    row['conducteur'],
                )
            else:
                key = ()

            if key not in groups:
                groups[key] = {'nb_lignes': 0, 'litres': 0.0, 'montant': 0.0}
            groups[key]['nb_lignes'] += 1
            groups[key]['litres'] += row['litres']
            groups[key]['montant'] += row['montant']

        grouped_rows = [
            {'keys': key, 'nb_lignes': val['nb_lignes'], 'litres': round(val['litres'], 2), 'montant': round(val['montant'], 2)}
            for key, val in sorted(groups.items(), key=lambda item: item[0])
        ]

        if group_by == 'mois':
            group_by_columns = ['Mois']
        elif group_by == 'vehicule':
            group_by_columns = ['Véhicule']
        elif group_by == 'conducteur':
            group_by_columns = ['Conducteur']
        elif group_by == 'mois_vehicule':
            group_by_columns = ['Mois', 'Véhicule']
        elif group_by == 'mois_conducteur':
            group_by_columns = ['Mois', 'Conducteur']

    return render(request, 'transport/gasoil.html', {
        'rows': rows,
        'grouped_rows': grouped_rows,
        'group_by': group_by,
        'group_by_columns': group_by_columns,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'vehicule': vehicule,
        'conducteur_id': conducteur_id,
        'vehicules': vehicules,
        'conducteurs': conducteurs,
        'total_rows': total_rows,
        'total_litres': total_litres,
        'total_montant': total_montant,
    })


def _domain_transport_analytic_line(
    date_debut, date_fin, vehicule,
    product_id=None, company_id=None,
):
    """Lignes analytiques transport : picking avec consommation OU transport_logistics."""
    domain = []
    if date_debut:
        domain.append(('date', '>=', date_debut))
    if date_fin:
        domain.append(('date', '<=', date_fin))
    domain.extend([
        '|',
        ('transfer_consumption_id', '!=', False),
        ('transfer_consumption_id.transport_logistics', '=', True),
    ])
    if vehicule:
        v = str(vehicule).strip()
        domain.extend([
            '|',
            ('transfer_consumption_id.equipment_id.name', 'ilike', v),
            ('transfer_consumption_id.affectation_id.name', 'ilike', v),
        ])
    if product_id:
        domain.append(('product_id', '=', int(product_id)))
    if company_id:
        domain.append(('company_id', '=', int(company_id)))
    return domain


@login_required
def transport_couts_nature(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    nature = request.GET.get('nature', '').strip()
    vehicule = request.GET.get('vehicule', '').strip()
    product_id = request.GET.get('product_id', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    grouped_rows = []
    nature_options = []
    vehicules = []
    product_options = []
    company_options = []
    error = None
    total_montant = 0.0

    try:
        uid, models = get_odoo_connection()

        vehicules_raw = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': ['id', 'name'], 'order': 'name', 'limit': 500},
        )
        vehicules = sorted({
            extract_matricule(v.get('name'))
            for v in vehicules_raw
            if v.get('name')
        })

        base_filter_domain = _domain_transport_analytic_line(
            date_debut, date_fin,
            vehicule,
        )
        options_domain = _domain_transport_analytic_line(
            date_debut, date_fin,
            None,
        )
        domain = _domain_transport_analytic_line(
            date_debut, date_fin,
            vehicule,
            product_id if product_id.isdigit() else None,
            company_id if company_id.isdigit() else None,
        )

        def _distinct_many2one_options(field_name, option_domain=None):
            current_domain = option_domain if option_domain is not None else base_filter_domain
            lines_for_field = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'account.analytic.line', 'search_read',
                [current_domain],
                {'fields': [field_name], 'limit': False},
            )
            seen = {}
            for ln in lines_for_field:
                val = ln.get(field_name)
                if isinstance(val, list) and len(val) >= 2:
                    seen[val[0]] = val[1]
            vals = [{'id': k, 'name': v} for k, v in seen.items()]
            vals.sort(key=lambda x: (x['name'] or '').lower())
            return vals

        product_options = _distinct_many2one_options('product_id', option_domain=options_domain)
        # Société: liste globale, sans filtre transport_logistics.
        company_options = _distinct_many2one_options('company_id', option_domain=[])

        models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.analytic.line', 'search_read',
            [domain],
            {
                'fields': ['date', 'amount', 'nature_id', 'transfer_consumption_id'],
                'order': 'date desc, id desc',
                'limit': 5,
            },
        )

        lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.analytic.line', 'search_read',
            [domain],
            {
                'fields': [
                    'date', 'amount', 'nature_id', 'transfer_consumption_id',
                    'product_categ_id', 'general_account_id', 'account_id', 'name',
                ],
                'order': 'date desc, id desc',
                'limit': 5000,
            },
        )

        picking_ids = list({
            ln['transfer_consumption_id'][0]
            for ln in lines
            if ln.get('transfer_consumption_id')
        })
        pickings = {}
        chunk_size = 200
        for i in range(0, len(picking_ids), chunk_size):
            chunk = picking_ids[i : i + chunk_size]
            for p in models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.picking', 'read',
                [chunk],
                {'fields': ['equipment_id', 'affectation_id', 'name']},
            ):
                pickings[p['id']] = p

        def vehicule_label(p):
            if not p:
                return '—'
            if p.get('equipment_id'):
                return extract_matricule(p['equipment_id'][1])
            if p.get('affectation_id'):
                return extract_matricule(p['affectation_id'][1])
            return '—'

        def nature_label(ln):
            if ln.get('nature_id'):
                return ln['nature_id'][1]
            if ln.get('product_categ_id'):
                return ln['product_categ_id'][1]
            if ln.get('general_account_id'):
                return ln['general_account_id'][1]
            if ln.get('account_id'):
                return ln['account_id'][1]
            return ln.get('name') or '—'

        for ln in lines:
            amt = abs(float(ln.get('amount') or 0))
            pid = ln['transfer_consumption_id'][0] if ln.get('transfer_consumption_id') else None
            pk = pickings.get(pid) if pid else None
            nat = nature_label(ln)
            d = (ln.get('date') or '')[:10] or '—'
            rows.append({
                'date': d,
                'vehicule': vehicule_label(pk),
                'nature': nat,
                'montant': round(amt, 2),
            })

        lines_for_nature = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.analytic.line', 'search_read',
            [options_domain],
            {'fields': ['nature_id', 'product_categ_id', 'general_account_id', 'account_id', 'name'], 'limit': False},
        )
        nature_options = sorted({
            nature_label(ln)
            for ln in lines_for_nature
            if nature_label(ln) and nature_label(ln) != '—'
        })
        if nature:
            rows = [r for r in rows if r['nature'] == nature]

        total_montant = round(sum(r['montant'] for r in rows), 2)

        if group_by == 'nature':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                k = r['nature']
                buckets[k]['count'] += 1
                buckets[k]['montant'] += r['montant']
            grouped_rows = [
                {'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)}
                for k, v in sorted(buckets.items(), key=lambda kv: (-kv[1]['montant'], kv[0]))
            ]
        elif group_by == 'month':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                k = r['date'][:7] if len(r['date']) >= 7 else r['date']
                buckets[k]['count'] += 1
                buckets[k]['montant'] += r['montant']
            grouped_rows = [
                {'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)}
                for k, v in sorted(buckets.items(), key=lambda kv: kv[0], reverse=True)
            ]
        elif group_by == 'vehicle':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                k = r['vehicule']
                buckets[k]['count'] += 1
                buckets[k]['montant'] += r['montant']
            grouped_rows = [
                {'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)}
                for k, v in sorted(buckets.items(), key=lambda kv: (-kv[1]['montant'], kv[0]))
            ]

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_lignes = len(rows)
    nature_par_montant = defaultdict(float)
    for r in rows:
        nature_par_montant[r.get('nature') or '—'] += float(r.get('montant') or 0)
    nature_principale = '—'
    if nature_par_montant:
        nature_principale = max(nature_par_montant.items(), key=lambda kv: kv[1])[0]
    moyenne_ligne = round(total_montant / total_lignes, 2) if total_lignes else 0.0

    return render(request, 'transport/couts_nature.html', {
        'rows': rows,
        'grouped_rows': grouped_rows,
        'group_by': group_by,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'nature': nature,
        'vehicule': vehicule,
        'nature_options': nature_options,
        'vehicules': vehicules,
        'product_id': product_id,
        'company_id': company_id,
        'product_options': product_options,
        'company_options': company_options,
        'total_rows': total_lignes,
        'total_montant': total_montant,
        'nature_principale': nature_principale,
        'moyenne_ligne': moyenne_ligne,
    })


@login_required
def transport_facturation_client(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    numero_facture = request.GET.get('numero_facture', '').strip()
    client_id = request.GET.get('client_id', '').strip()
    shipping_id = request.GET.get('shipping_id', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    clients = []
    companies = []
    shippings = []
    grouped_rows = []
    kpi_paid_rate = 0.0
    kpi_overdue_rate = 0.0
    kpi_ticket_moyen = 0.0
    kpi_top_client = '—'
    kpi_top_client_share = 0.0
    kpi_ca_evolution = 0.0
    error = None
    try:
        uid, models = get_odoo_connection()
        base_domain = [('move_type', '=', 'out_invoice')]

        # Dropdowns depuis Odoo: client, lieu de livraison, société.
        inv_for_filters = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [base_domain],
            {'fields': ['partner_id', 'partner_shipping_id', 'company_id', 'project_id', 'invoice_origin', 'ref', 'name'], 'limit': False},
        )
        project_activity_map_filters = _build_project_activity_map(uid, models, inv_for_filters)
        seen_clients = {}
        seen_shippings = {}
        seen_companies = {}
        # Listes de filtres complètes (toutes factures), même si les résultats
        # de la page restent strictement filtrés sur le périmètre transport.
        for inv in inv_for_filters:
            p = inv.get('partner_id')
            if isinstance(p, list) and len(p) >= 2:
                seen_clients[p[0]] = p[1]
            s = inv.get('partner_shipping_id')
            if isinstance(s, list) and len(s) >= 2:
                seen_shippings[s[0]] = s[1]
            c = inv.get('company_id')
            if isinstance(c, list) and len(c) >= 2:
                seen_companies[c[0]] = c[1]
        clients = [{'id': cid, 'name': cname} for cid, cname in seen_clients.items()]
        clients.sort(key=lambda x: (x['name'] or '').lower())
        shippings = [{'id': sid, 'name': sname} for sid, sname in seen_shippings.items()]
        shippings.sort(key=lambda x: (x['name'] or '').lower())
        companies = [{'id': k, 'name': v} for k, v in seen_companies.items()]
        companies.sort(key=lambda x: (x['name'] or '').lower())

        domain = [('move_type', '=', 'out_invoice')]
        domain += _transport_domain_dates(date_debut, date_fin, 'invoice_date')
        if client_id.isdigit():
            domain.append(('partner_id', '=', int(client_id)))
        if shipping_id.isdigit():
            domain.append(('partner_shipping_id', '=', int(shipping_id)))
        if company_id.isdigit():
            domain.append(('company_id', '=', int(company_id)))
        if numero_facture:
            # Recherche élargie pour les cas où l'utilisateur saisit un code projet
            # (ex: S00068) au lieu du numéro exact de facture.
            domain += [
                '|', '|', '|',
                ('name', 'ilike', numero_facture),
                ('ref', 'ilike', numero_facture),
                ('invoice_origin', 'ilike', numero_facture),
                ('project_id.name', 'ilike', numero_facture),
            ]

        invoices = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [domain],
            {
                'fields': [
                    'name', 'invoice_date', 'partner_id', 'partner_shipping_id',
                    'company_id', 'amount_untaxed', 'amount_total',
                    'payment_state', 'invoice_date_due',
                    'project_id', 'invoice_origin', 'ref',
                ],
                'order': 'invoice_date desc',
                'limit': 1000,
            }
        )
        project_activity_map = _build_project_activity_map(uid, models, invoices)
        invoices = [inv for inv in invoices if _invoice_activity_bucket(inv, project_activity_map) == 'transport']

        def _fmt_date_fr(iso_date):
            if not iso_date:
                return '—'
            s = str(iso_date)[:10]
            if len(s) == 10 and s[4] == '-' and s[7] == '-':
                return f'{s[8:10]}/{s[5:7]}/{s[0:4]}'
            return s

        for inv in invoices:
            iso_date = (inv.get('invoice_date') or '')[:10]
            ht = round(inv.get('amount_untaxed') or 0, 2)
            ttc = round(inv.get('amount_total') or 0, 2)
            tva = round(ttc - ht, 2)
            projet_ref = '—'
            proj = inv.get('project_id')
            if isinstance(proj, list) and len(proj) >= 2 and proj[1]:
                projet_ref = proj[1]
            elif inv.get('invoice_origin'):
                projet_ref = inv.get('invoice_origin')
            elif inv.get('ref'):
                projet_ref = inv.get('ref')
            rows.append({
                'date': _fmt_date_fr(iso_date),
                'month_key': iso_date[:7] if len(iso_date) >= 7 else '—',
                'numero': inv.get('name') or '—',
                'projet_reference': projet_ref,
                'client': inv['partner_id'][1] if inv.get('partner_id') else '—',
                'lieu_livraison': inv['partner_shipping_id'][1] if inv.get('partner_shipping_id') else '—',
                'societe': inv['company_id'][1] if inv.get('company_id') else '—',
                'ht': ht,
                'tva': tva,
                'ttc': ttc,
                'payment_state': inv.get('payment_state') or '',
                'invoice_date_due': (inv.get('invoice_date_due') or '')[:10],
            })

        if group_by:
            buckets = defaultdict(lambda: {'count': 0, 'ht': 0.0, 'tva': 0.0, 'ttc': 0.0})
            for r in rows:
                if group_by == 'month':
                    key = r['month_key']
                elif group_by == 'company':
                    key = r['societe']
                elif group_by == 'month_company':
                    mois = r['month_key']
                    key = f'{mois} | {r["societe"]}'
                elif group_by == 'month_delivery':
                    mois = r['month_key']
                    key = f'{mois} | {r["lieu_livraison"]}'
                elif group_by == 'client':
                    key = r['client']
                else:
                    key = '—'
                buckets[key]['count'] += 1
                buckets[key]['ht'] += r['ht']
                buckets[key]['tva'] += r['tva']
                buckets[key]['ttc'] += r['ttc']

            mois_fr = {
                '01': 'Janvier', '02': 'Fevrier', '03': 'Mars', '04': 'Avril',
                '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Aout',
                '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Decembre',
            }

            def _month_label(ym):
                if isinstance(ym, str) and len(ym) == 7 and ym[4] == '-':
                    y = ym[:4]
                    m = ym[5:7]
                    return f"{mois_fr.get(m, m)} {y}"
                return ym

            grouped_rows = [
                {
                    'label': k,
                    'count': v['count'],
                    'ht': round(v['ht'], 2),
                    'tva': round(v['tva'], 2),
                    'ttc': round(v['ttc'], 2),
                    'sort_key': k,
                }
                for k, v in buckets.items()
            ]
            if group_by == 'month':
                for g in grouped_rows:
                    g['label'] = _month_label(g['sort_key'])
                grouped_rows.sort(key=lambda x: x['sort_key'])
            elif group_by in {'month_company', 'month_delivery'}:
                for g in grouped_rows:
                    parts = str(g['sort_key']).split(' | ', 1)
                    ym = parts[0]
                    suffix = parts[1] if len(parts) > 1 else ''
                    g['label'] = f"{_month_label(ym)} | {suffix}" if suffix else _month_label(ym)
                grouped_rows.sort(key=lambda x: x['sort_key'])
            else:
                grouped_rows.sort(key=lambda x: (-x['ttc'], x['label']))
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_ht = round(sum(r['ht'] for r in rows), 2)
    total_tva = round(sum(r['tva'] for r in rows), 2)
    total_ttc = round(sum(r['ttc'] for r in rows), 2)
    total_rows = len(rows)
    kpi_ticket_moyen = round((total_ttc / total_rows), 2) if total_rows else 0.0

    paid_count = sum(1 for r in rows if r.get('payment_state') == 'paid')
    kpi_paid_rate = round((paid_count / total_rows) * 100, 1) if total_rows else 0.0

    today_iso = date.today().isoformat()
    overdue_count = sum(
        1 for r in rows
        if (r.get('invoice_date_due') and r.get('invoice_date_due') < today_iso and r.get('payment_state') != 'paid')
    )
    kpi_overdue_rate = round((overdue_count / total_rows) * 100, 1) if total_rows else 0.0

    # KPI client leader (part du CA TTC)
    client_totals = defaultdict(float)
    for r in rows:
        client_totals[r.get('client') or '—'] += float(r.get('ttc') or 0)
    if client_totals and total_ttc > 0:
        best_client, best_value = max(client_totals.items(), key=lambda kv: kv[1])
        kpi_top_client = best_client
        kpi_top_client_share = round((best_value / total_ttc) * 100, 1)

    # KPI évolution CA mensuel (% M vs M-1)
    monthly_totals = defaultdict(float)
    for r in rows:
        mk = r.get('month_key') or '—'
        if mk != '—':
            monthly_totals[mk] += float(r.get('ttc') or 0)
    if len(monthly_totals) >= 2:
        keys = sorted(monthly_totals.keys())
        prev_v = monthly_totals[keys[-2]]
        curr_v = monthly_totals[keys[-1]]
        if prev_v > 0:
            kpi_ca_evolution = round(((curr_v - prev_v) / prev_v) * 100, 1)

    # Section séparée: CA par lieu de livraison.
    delivery_buckets = defaultdict(lambda: {'count': 0, 'ht': 0.0, 'tva': 0.0, 'ttc': 0.0})
    for r in rows:
        k = r['lieu_livraison']
        delivery_buckets[k]['count'] += 1
        delivery_buckets[k]['ht'] += r['ht']
        delivery_buckets[k]['tva'] += r['tva']
        delivery_buckets[k]['ttc'] += r['ttc']
    delivery_rows = [
        {
            'lieu': k,
            'count': v['count'],
            'ht': round(v['ht'], 2),
            'tva': round(v['tva'], 2),
            'ttc': round(v['ttc'], 2),
        }
        for k, v in sorted(delivery_buckets.items(), key=lambda kv: (-kv[1]['ttc'], kv[0]))
    ]
    delivery_total_count = sum(r['count'] for r in delivery_rows)
    delivery_total_ht = round(sum(r['ht'] for r in delivery_rows), 2)
    delivery_total_tva = round(sum(r['tva'] for r in delivery_rows), 2)
    delivery_total_ttc = round(sum(r['ttc'] for r in delivery_rows), 2)

    kpi_taux_tva_moyen = round((total_tva / total_ht) * 100, 2) if total_ht else 0.0

    return render(request, 'transport/facturation_client.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'numero_facture': numero_facture,
        'client_id': client_id,
        'clients': clients,
        'shipping_id': shipping_id,
        'shippings': shippings,
        'company_id': company_id,
        'companies': companies,
        'group_by': group_by,
        'grouped_rows': grouped_rows,
        'total_rows': total_rows,
        'total_ht': total_ht,
        'total_tva': total_tva,
        'total_ttc': total_ttc,
        'kpi_paid_rate': kpi_paid_rate,
        'kpi_overdue_rate': kpi_overdue_rate,
        'kpi_ticket_moyen': kpi_ticket_moyen,
        'kpi_top_client': kpi_top_client,
        'kpi_top_client_share': kpi_top_client_share,
        'kpi_ca_evolution': kpi_ca_evolution,
        'kpi_taux_tva_moyen': kpi_taux_tva_moyen,
        'delivery_rows': delivery_rows,
        'delivery_total_count': delivery_total_count,
        'delivery_total_ht': delivery_total_ht,
        'delivery_total_tva': delivery_total_tva,
        'delivery_total_ttc': delivery_total_ttc,
    })


@login_required
def production_index(request):
    return render(request, 'production/production.html', {})


@login_required
def production_dashboard(request):
    context = ProductionService.dashboard_data()
    return render(request, 'production/dashboard.html', context)


@login_required
def production_ratios(request):
    context = ProductionService.ratios_data()
    return render(request, 'production/ratios.html', context)


@login_required
def production_ipc(request):
    context = ProductionService.ipc_data()
    return render(request, 'production/ipc.html', context)


@login_required
def production_rapports(request):
    selected_periode = request.GET.get('periode', 'Mois')
    selected_site = request.GET.get('site', 'Tous')
    selected_rapport = request.GET.get('rapport', 'Rapport mensuel production')
    action = request.GET.get('action', '').strip().lower()
    context = ProductionService.rapports_data()

    report_rows = [
        ['Type rapport', selected_rapport],
        ['Periode', selected_periode],
        ['Site', selected_site],
        ['Date generation', timezone.localtime().strftime('%d/%m/%Y %H:%M')],
    ]

    if action == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Rapport production'
        ws.append(['Rapports & Analyses - Production'])
        ws.append([])
        for row in report_rows:
            ws.append(row)
        ws.column_dimensions['A'].width = 28
        ws.column_dimensions['B'].width = 48
        ws['A1'].font = Font(bold=True, color='1A2C4E')
        ws['A1'].fill = PatternFill('solid', fgColor='FCE8D7')
        for line in range(3, 7):
            ws[f'A{line}'].font = Font(bold=True)
        data = io.BytesIO()
        wb.save(data)
        data.seek(0)
        response = HttpResponse(
            data.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response['Content-Disposition'] = 'attachment; filename=rapport_production.xlsx'
        return response

    if action == 'pdf':
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        story = [
            Paragraph('Rapports & Analyses - Production', styles['Title']),
            Spacer(1, 8),
            Paragraph(f"Rapport: {html.escape(selected_rapport)}", styles['Normal']),
            Spacer(1, 6),
        ]
        table = Table(report_rows, colWidths=[140, 320])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FCE8D7')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#1A2C4E')),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E8EBF0')),
            ('PADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(table)
        doc.build(story)
        pdf = buffer.getvalue()
        buffer.close()
        response = HttpResponse(pdf, content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename=rapport_production.pdf'
        return response

    action_message = None
    if action == 'email':
        action_message = "Rapport prepare. Envoi email simule (module SMTP a brancher)."

    context.update({
        'selected_periode': selected_periode,
        'selected_site': selected_site,
        'selected_rapport': selected_rapport,
        'action_message': action_message,
    })
    return render(request, 'production/rapports.html', context)


_POINTAGE_FORATION_KEYS = ('fora', 'forage', 'foration', 'drill')


def _pointage_bon_is_foration(b):
    trace = ' '.join([
        str(b.get('ouvrage') or ''),
        str(b.get('type_operation') or ''),
        str(b.get('affectation') or ''),
    ]).lower()
    return any(k in trace for k in _POINTAGE_FORATION_KEYS)


def _filter_pointages_bons(bons, *, ouvrage, engin, societe, foration_filtre, anomalie, operation_q):
    out = list(bons)
    if ouvrage:
        q = ouvrage.lower()
        out = [b for b in out if q in (str(b.get('ouvrage') or '')).lower()]
    if engin:
        q = engin.lower()
        out = [
            b for b in out
            if q in (str(b.get('engin') or '')).lower()
            or q in (str(b.get('affectation') or '')).lower()
        ]
    if societe:
        q = societe.lower()
        out = [b for b in out if q in (str(b.get('societe') or '')).lower()]
    if operation_q:
        q = operation_q.lower()
        out = [b for b in out if q in (str(b.get('name') or '')).lower()]
    if foration_filtre == 'oui':
        out = [b for b in out if _pointage_bon_is_foration(b)]
    elif foration_filtre == 'non':
        out = [b for b in out if not _pointage_bon_is_foration(b)]
    if anomalie == 'ok':
        out = [b for b in out if b.get('anomalie') == 'OK']
    elif anomalie == 'anomalie':
        out = [b for b in out if b.get('anomalie') == 'Anomalie']
    return out


def _odoo_active_company_names(uid, models):
    """Noms des sociétés actives Odoo (res.company) — complète les listes de filtres."""
    try:
        rows = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'res.company', 'search_read',
            [[['active', '=', True]]],
            {'fields': ['name'], 'order': 'name', 'limit': 500},
        )
        out = []
        for r in rows or []:
            n = (r.get('name') or '').strip()
            if n:
                out.append(n)
        return out
    except Exception:
        return []


def _pointage_filter_options_merge_res_company_societes(opts, uid, models):
    """Ajoute toutes les sociétés actives Odoo aux options « societes » (ex. OTEC + SOMATRIN)."""
    if not opts or not uid or not models:
        return
    seen = set(opts.get('societes') or [])
    for n in _odoo_active_company_names(uid, models):
        seen.add(n)
    opts['societes'] = sorted(seen, key=lambda x: (x or '').lower())


def _pointages_filter_options_from_bons(bons):
    """Valeurs distinctes pour listes déroulantes (même périmètre que les bons, sans filtre site/chauffeur)."""
    sites, chauffeurs, societes, ouvrages, engins, operations = (set() for _ in range(6))
    for b in bons:
        s = (b.get('site') or '').strip()
        if s and s != '—':
            sites.add(s)
        ch = (b.get('chauffeur') or '').strip()
        if ch and ch != '—':
            chauffeurs.add(ch)
        soc = (b.get('societe') or '').strip()
        if soc and soc != '—':
            societes.add(soc)
        ouv = (b.get('ouvrage') or '').strip()
        if ouv and ouv != '—':
            ouvrages.add(ouv)
        eg = (b.get('engin') or '').strip()
        if eg and eg != '—':
            engins.add(eg)
        aff = (b.get('affectation') or '').strip()
        if aff and aff != '—':
            engins.add(aff)
        op = (b.get('name') or '').strip()
        if op and op != '—':
            operations.add(op)
    key = lambda x: (x or '').lower()
    return {
        'sites': sorted(sites, key=key),
        'chauffeurs': sorted(chauffeurs, key=key),
        'societes': sorted(societes, key=key),
        'ouvrages': sorted(ouvrages, key=key),
        'engins': sorted(engins, key=key),
        'operations': sorted(operations, key=key),
    }


def _pointages_filter_options_from_demo_rows(rows):
    """Options de filtres à partir des lignes de démonstration."""
    sites, chauffeurs, societes, ouvrages, engins, operations = (set() for _ in range(6))
    for r in rows:
        s = (r.get('site') or '').strip()
        if s and s != '—':
            sites.add(s)
        eq = (r.get('equipe') or '').strip()
        if eq and eq != '—':
            chauffeurs.add(eq)
        soc = (r.get('societe') or '').strip()
        if soc and soc != '—':
            societes.add(soc)
        ouv = (r.get('ouvrage') or '').strip()
        if ouv and ouv != '—':
            ouvrages.add(ouv)
        eg = (r.get('engin') or '').strip()
        if eg and eg != '—':
            engins.add(eg)
        op = (r.get('operation') or '').strip()
        if op and op != '—':
            operations.add(op)
    key = lambda x: (x or '').lower()
    return {
        'sites': sorted(sites, key=key),
        'chauffeurs': sorted(chauffeurs, key=key),
        'societes': sorted(societes, key=key),
        'ouvrages': sorted(ouvrages, key=key),
        'engins': sorted(engins, key=key),
        'operations': sorted(operations, key=key),
    }


def _filter_pointages_demo_rows(rows, *, chauffeur, ouvrage, engin, societe, foration_filtre, anomalie, operation_q):
    out = list(rows)
    if chauffeur:
        q = chauffeur.lower()
        out = [r for r in out if q in (str(r.get('equipe') or '')).lower()]
    if ouvrage:
        q = ouvrage.lower()
        out = [r for r in out if q in (str(r.get('ouvrage') or '')).lower()]
    if engin:
        q = engin.lower()
        out = [r for r in out if q in (str(r.get('engin') or '')).lower()]
    if societe:
        q = societe.lower()
        out = [r for r in out if q in (str(r.get('societe') or '')).lower()]
    if operation_q:
        q = operation_q.lower()
        out = [r for r in out if q in (str(r.get('operation') or '')).lower()]
    if foration_filtre == 'oui':
        out = [r for r in out if r.get('foration') == 'Oui']
    elif foration_filtre == 'non':
        out = [r for r in out if r.get('foration') == 'Non']
    if anomalie == 'ok':
        out = [r for r in out if r.get('statut') != 'Critique']
    elif anomalie == 'anomalie':
        out = [r for r in out if r.get('statut') == 'Critique']
    return out


def _parse_float_param(raw):
    if raw is None:
        return None
    s = str(raw).strip().replace(',', '.')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _filter_machines_summary_rows(rows, *, site_contains, min_heures, min_tonnage, min_operations, min_rendement):
    """Filtre les lignes agrégées par machine (après consolidation)."""
    out = list(rows)
    if site_contains:
        q = site_contains.lower()
        out = [r for r in out if q in (str(r.get('site') or '')).lower()]
    if min_heures is not None and min_heures > 0:
        out = [r for r in out if float(r.get('heures') or 0) >= min_heures]
    if min_tonnage is not None and min_tonnage > 0:
        out = [r for r in out if float(r.get('tonnage') or 0) >= min_tonnage]
    if min_operations is not None and min_operations > 0:
        mo = max(1, int(round(min_operations)))
        out = [r for r in out if int(r.get('operations') or 0) >= mo]
    if min_rendement is not None and min_rendement > 0:
        out = [r for r in out if float(r.get('rendement') or 0) >= min_rendement]
    return out


def pointages_operations_request_context(request):
    """
    Contexte métier Pointages opérations & foration (mêmes données que l'écran).
    Inclut `pointages_export_qs` pour répliquer les filtres dans Excel/CSV/PDF.
    """
    q = request.GET
    date_debut = q.get('date_debut', '').strip()
    date_fin = q.get('date_fin', '').strip()
    site = q.get('site', '').strip()
    chauffeur = q.get('chauffeur', '').strip()
    ouvrage = q.get('ouvrage', '').strip()
    engin = q.get('engin', '').strip()
    societe = q.get('societe', '').strip()
    foration_filtre = q.get('foration', '').strip().lower()
    anomalie = q.get('anomalie', '').strip().lower()
    operation_q = q.get('operation', '').strip()

    rows = []
    error = None
    demo_notice = None
    empty_filter_notice = None
    bons_prod = []
    pointage_filter_options = {
        'sites': [], 'chauffeurs': [], 'societes': [], 'ouvrages': [], 'engins': [], 'operations': [],
    }
    uid = None
    models = None

    try:
        uid, models = get_odoo_connection()
        domain_opts = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='production',
        )
        domain_opts.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons_opts_raw = _fetch_sorties_bons(uid, models, domain_opts, limit=8000)
        bons_opts = [b for b in bons_opts_raw if b.get('activite_bucket') == 'production']
        pointage_filter_options = _pointages_filter_options_from_bons(bons_opts)
        _pointage_filter_options_merge_res_company_societes(pointage_filter_options, uid, models)

        domain = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site=site,
            chauffeur=chauffeur,
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='production',
        )
        domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons = _fetch_sorties_bons(uid, models, domain, limit=8000)
        bons_prod = [b for b in bons if b.get('activite_bucket') == 'production']

        bons_filtered = _filter_pointages_bons(
            bons_prod,
            ouvrage=ouvrage,
            engin=engin,
            societe=societe,
            foration_filtre=foration_filtre,
            anomalie=anomalie,
            operation_q=operation_q,
        )

        for b in bons_filtered:
            is_foration = _pointage_bon_is_foration(b)
            rows.append({
                'date': b.get('date') or '—',
                'operation': b.get('name') or '—',
                'site': b.get('site') or '—',
                'engin': b.get('engin') or '—',
                'equipe': b.get('chauffeur') or '—',
                'ouvrage': b.get('ouvrage') or '—',
                'societe': b.get('societe') or '—',
                'foration': 'Oui' if is_foration else 'Non',
                'ml_estime': round(max(float(b.get('ecart') or 0), 0.0), 1) if is_foration else 0.0,
                'statut': 'Actif' if b.get('anomalie') != 'Anomalie' else 'Critique',
            })

        if bons_prod and not rows:
            empty_filter_notice = (
                'Aucun bon ne correspond à ces critères. Élargissez la période ou réinitialisez les filtres.'
            )
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    if error:
        pointage_filter_options = _pointages_filter_options_from_demo_rows(
            ProductionService.pointages_operations_fallback_rows()
        )

    if not error and not bons_prod:
        demo_notice = (
            "Aucun mouvement Odoo sur la période sélectionnée — affichage d'exemples représentatifs "
            "(pointage opérations, forages SRO, minages / rapports de tir). Les filtres s'appliquent aux exemples."
        )
        demo_all_rows = ProductionService.pointages_operations_fallback_rows()
        pointage_filter_options = _pointages_filter_options_from_demo_rows(demo_all_rows)
        _pointage_filter_options_merge_res_company_societes(pointage_filter_options, uid, models)
        rows = _filter_pointages_demo_rows(
            demo_all_rows,
            chauffeur=chauffeur,
            ouvrage=ouvrage,
            engin=engin,
            societe=societe,
            foration_filtre=foration_filtre,
            anomalie=anomalie,
            operation_q=operation_q,
        )
        if not rows:
            rows = ProductionService.pointages_operations_fallback_rows()

    export_keys = (
        'date_debut', 'date_fin', 'site', 'chauffeur', 'ouvrage', 'engin',
        'societe', 'foration', 'anomalie', 'operation',
    )
    export_params = {k: q.get(k, '').strip() for k in export_keys}
    export_params = {k: v for k, v in export_params.items() if v}
    pointages_export_qs = urlencode(export_params)

    return {
        'rows': rows,
        'error': error,
        'demo_notice': demo_notice,
        'empty_filter_notice': empty_filter_notice,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'site': site,
        'chauffeur': chauffeur,
        'ouvrage': ouvrage,
        'engin': engin,
        'societe': societe,
        'foration_filtre': foration_filtre,
        'anomalie': anomalie,
        'operation_q': operation_q,
        'pointage_filter_options': pointage_filter_options,
        'pointages_export_qs': pointages_export_qs,
    }


@login_required
def production_pointages_operations_foration(request):
    ctx = pointages_operations_request_context(request)
    rows = ctx['rows']
    total_ops = len(rows)
    equipes = {r['equipe'] for r in rows if r.get('equipe') and r['equipe'] != '—'}
    foration_rows = [r for r in rows if r.get('foration') == 'Oui']
    conformes = sum(1 for r in rows if r.get('statut') != 'Critique')
    ctx['kpis'] = {
        'operations_jour': total_ops,
        'equipes_actives': len(equipes),
        'foration_realisee_ml': round(sum(float(r.get('ml_estime') or 0) for r in foration_rows), 1),
        'taux_conformite': round((conformes / total_ops) * 100, 1) if total_ops else 0.0,
    }
    ctx['rows'] = rows[:300]
    return render(request, 'production/pointages_operations_foration.html', ctx)


def _machines_filter_equipment_category_names(uid, models):
    """Noms de catégories Odoo (equipment.category) pour engins du périmètre production."""
    try:
        rows = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[['category_id', 'in', CATEGORIES_PRODUCTION]]],
            {'fields': ['category_id'], 'limit': 3000},
        )
        names = set()
        for r in rows or []:
            c = r.get('category_id')
            if isinstance(c, list) and len(c) >= 2 and c[1]:
                names.add(str(c[1]).strip())
        return sorted(names, key=lambda x: x.lower())
    except Exception:
        return []


def _machines_filter_options_from_bons(bons, uid, models):
    """Options listes déroulantes page Machines / Heures / Tonnages."""
    opts = _pointages_filter_options_from_bons(bons)
    opts['categories_engin'] = _machines_filter_equipment_category_names(uid, models)
    _pointage_filter_options_merge_res_company_societes(opts, uid, models)
    opts['sites_ligne'] = list(opts.get('sites') or [])
    return opts


@login_required
def production_machines_heures_tonnages(request):
    date_debut = request.GET.get('date_debut', '').strip()
    date_fin = request.GET.get('date_fin', '').strip()
    site = request.GET.get('site', '').strip()
    chauffeur = request.GET.get('chauffeur', '').strip()
    ouvrage = request.GET.get('ouvrage', '').strip()
    engin = request.GET.get('engin', '').strip()
    societe = request.GET.get('societe', '').strip()
    anomalie = request.GET.get('anomalie', '').strip().lower()
    operation_q = request.GET.get('operation', '').strip()
    categorie_engin = request.GET.get('categorie_engin', '').strip()
    site_ligne = request.GET.get('site_ligne', '').strip()
    min_heures = _parse_float_param(request.GET.get('min_heures'))
    min_tonnage = _parse_float_param(request.GET.get('min_tonnage'))
    min_operations = _parse_float_param(request.GET.get('min_operations'))
    min_rendement = _parse_float_param(request.GET.get('min_rendement'))

    rows = []
    error = None
    empty_filter_notice = None
    bons_prod = []
    kpis = {
        'machines_actives': 0,
        'heures_machines': 0.0,
        'tonnages_produits': 0.0,
        'rendement_t_h': 0.0,
    }
    machines_filter_options = {
        'sites': [],
        'chauffeurs': [],
        'societes': [],
        'ouvrages': [],
        'engins': [],
        'operations': [],
        'categories_engin': [],
        'sites_ligne': [],
    }

    try:
        uid, models = get_odoo_connection()

        domain_opts = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='production',
        )
        domain_opts.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons_opts_raw = _fetch_sorties_bons(uid, models, domain_opts, limit=8000)
        bons_opts = [b for b in bons_opts_raw if b.get('activite_bucket') == 'production']
        machines_filter_options = _machines_filter_options_from_bons(bons_opts, uid, models)

        domain = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site=site,
            chauffeur=chauffeur,
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin=categorie_engin,
            activite_filtre='production',
        )
        domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons = _fetch_sorties_bons(uid, models, domain, limit=8000)
        bons_prod = [b for b in bons if b.get('activite_bucket') == 'production']

        bons_filtered = _filter_pointages_bons(
            bons_prod,
            ouvrage=ouvrage,
            engin=engin,
            societe=societe,
            foration_filtre='',
            anomalie=anomalie,
            operation_q=operation_q,
        )

        by_machine = defaultdict(lambda: {
            'machine': '—',
            'site': '—',
            'heures': 0.0,
            'tonnage': 0.0,
            'operations': 0,
        })
        for b in bons_filtered:
            machine = b.get('engin') or '—'
            cur = by_machine[machine]
            cur['machine'] = machine
            cur['site'] = b.get('site') or '—'
            cur['operations'] += 1
            cur['heures'] += max(float(b.get('ecart') or 0), 0.0)
            cur['tonnage'] += max(float(b.get('product_qty') or 0), 0.0)

        rows = []
        for _, r in by_machine.items():
            rendement = (r['tonnage'] / r['heures']) if r['heures'] > 0 else 0.0
            rows.append({
                'machine': r['machine'],
                'site': r['site'],
                'heures': round(r['heures'], 1),
                'tonnage': round(r['tonnage'], 1),
                'operations': r['operations'],
                'rendement': round(rendement, 2),
            })
        rows.sort(key=lambda x: (-x['tonnage'], x['machine']))

        rows = _filter_machines_summary_rows(
            rows,
            site_contains=site_ligne,
            min_heures=min_heures,
            min_tonnage=min_tonnage,
            min_operations=min_operations,
            min_rendement=min_rendement,
        )

        if bons_prod and not bons_filtered:
            empty_filter_notice = (
                'Aucun bon ne correspond à ces critères. Élargissez la période ou réinitialisez les filtres.'
            )
        elif bons_filtered and not rows:
            empty_filter_notice = (
                'Aucune ligne agrégée ne respecte les seuils (heures, tonnage, opérations, rendement) ou le site affiché.'
            )

        total_heures = round(sum(r['heures'] for r in rows), 1)
        total_tonnages = round(sum(r['tonnage'] for r in rows), 1)
        kpis['machines_actives'] = len(rows)
        kpis['heures_machines'] = total_heures
        kpis['tonnages_produits'] = total_tonnages
        kpis['rendement_t_h'] = round((total_tonnages / total_heures), 2) if total_heures else 0.0
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    context = {
        'kpis': kpis,
        'rows': rows[:300],
        'error': error,
        'empty_filter_notice': empty_filter_notice,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'site': site,
        'chauffeur': chauffeur,
        'ouvrage': ouvrage,
        'engin': engin,
        'societe': societe,
        'anomalie': anomalie,
        'operation_q': operation_q,
        'categorie_engin': categorie_engin,
        'site_ligne': site_ligne,
        'min_heures': request.GET.get('min_heures', '').strip(),
        'min_tonnage': request.GET.get('min_tonnage', '').strip(),
        'min_operations': request.GET.get('min_operations', '').strip(),
        'min_rendement': request.GET.get('min_rendement', '').strip(),
        'machines_filter_options': machines_filter_options,
    }
    return render(request, 'production/machines_heures_tonnages.html', context)


@login_required
def production_gasoil(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    societe = request.GET.get('societe', '').strip()
    site = request.GET.get('site', '').strip()
    statut = request.GET.get('statut', '').strip()
    chauffeur = request.GET.get('chauffeur', '').strip()
    ouvrage = request.GET.get('ouvrage', '').strip()
    activite = request.GET.get('activite', 'production').strip() or 'production'
    export = request.GET.get('export', '').strip().lower()

    gasoil_export_params = {}
    for k, v in (
        ('date_debut', date_debut),
        ('date_fin', date_fin),
        ('societe', societe),
        ('site', site),
        ('statut', statut),
        ('chauffeur', chauffeur),
        ('ouvrage', ouvrage),
        ('activite', activite),
    ):
        if v:
            gasoil_export_params[k] = v
    gasoil_export_qs = urlencode(gasoil_export_params)

    rows = []
    error = None
    societes = []
    sites = []
    chauffeurs = []
    ouvrages = []
    total_litres = 0.0
    total_montant = 0.0
    total_rows = 0
    conso_moyenne = 0.0
    nb_anomalies = 0
    try:
        uid, models = get_odoo_connection()

        list_domain = _build_sorties_domain(
            date_debut='',
            date_fin='',
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='production',
        )
        list_domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons_for_filters = _fetch_sorties_bons(uid, models, list_domain, limit=5000)
        societes = sorted({b['societe'] for b in bons_for_filters if b.get('societe') and b.get('societe') != '—'})
        sites = sorted({b['site'] for b in bons_for_filters if b.get('site') and b.get('site') != '—'})
        chauffeurs = sorted({b['chauffeur'] for b in bons_for_filters if b.get('chauffeur') and b.get('chauffeur') != '—'})
        ouvrages = sorted({b['ouvrage'] for b in bons_for_filters if b.get('ouvrage') and b.get('ouvrage') != '—'})

        domain = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site=site,
            chauffeur=chauffeur,
            ouvrage='',
            anomalie='',
            societe=societe,
            categorie_engin='',
            activite_filtre='production',
        )
        domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        limit = 10000 if (date_debut or date_fin or societe or site or chauffeur) else 2000
        bons = _fetch_sorties_bons(uid, models, domain, limit=limit)

        # Garantit que la page Production n'affiche que le périmètre exploitation/production.
        bons = [b for b in bons if b.get('activite_bucket') == 'production']

        if ouvrage:
            o = ouvrage.lower()
            bons = [b for b in bons if o in (b.get('ouvrage') or '').lower()]

        if statut == 'ok':
            bons = [b for b in bons if b.get('anomalie') == 'OK']
        elif statut == 'anomalie':
            bons = [b for b in bons if b.get('anomalie') == 'Anomalie']

        move_domain = [
            ['picking_id', 'in', [b['id'] for b in bons]],
            ['product_id.categ_id', '=', CARBURANT_CATEG_ID],
        ]
        move_lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.move', 'search_read',
            [move_domain],
            {'fields': ['picking_id', 'product_qty', 'price_total', 'unit_price'], 'limit': 20000},
        )
        amount_by_picking = defaultdict(float)
        for mv in move_lines:
            pid = mv['picking_id'][0] if mv.get('picking_id') else None
            if not pid:
                continue
            price_total = mv.get('price_total')
            if price_total not in (None, False):
                amount_by_picking[pid] += price_total or 0
            else:
                amount_by_picking[pid] += (mv.get('product_qty') or 0) * (mv.get('unit_price') or 0)

        def _fmt_date_fr(iso_date):
            if not iso_date or iso_date == '—' or len(iso_date) != 10:
                return iso_date or '—'
            return f'{iso_date[8:10]}/{iso_date[5:7]}/{iso_date[0:4]}'

        rows = [{
            'date': _fmt_date_fr(b.get('date') or '—'),
            'vehicule': extract_matricule(b.get('engin') or '—'),
            'conducteur': b.get('chauffeur') or '—',
            'ouvrage': b.get('ouvrage') or '—',
            'societe': b.get('societe') or '—',
            'site': b.get('site') or '—',
            'litres': round(b.get('product_qty') or 0, 2),
            'montant': round(amount_by_picking.get(b['id'], 0), 2),
            'compteur': b.get('cpt_actuel') or 0,
            'anomalie': b.get('anomalie') or 'OK',
        } for b in bons]

        total_rows = len(rows)
        total_litres = round(sum(r['litres'] for r in rows), 2)
        total_montant = round(sum(r['montant'] for r in rows), 2)
        nb_anomalies = sum(1 for b in bons if b.get('anomalie') == 'Anomalie')
        conso_vals = [b.get('consommation') or 0 for b in bons if (b.get('consommation') or 0) > 0]
        conso_moyenne = round(sum(conso_vals) / len(conso_vals), 2) if conso_vals else 0.0

        if export in {'excel', 'csv', 'pdf'}:
            if export == 'csv':
                out = io.StringIO()
                writer = csv.writer(out, delimiter=';')
                writer.writerow(['Date', 'Véhicule', 'Conducteur/Opérateur', 'Société', 'Site', 'Ouvrage', 'Litres', 'Montant', 'Compteur', 'Statut'])
                for r in rows:
                    writer.writerow([r['date'], r['vehicule'], r['conducteur'], r['societe'], r['site'], r['ouvrage'], r['litres'], r['montant'], r['compteur'], r['anomalie']])
                resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8')
                resp['Content-Disposition'] = 'attachment; filename=production_gasoil_sorties.csv'
                return resp

            if export == 'excel':
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = 'Production Gasoil'
                ws.append(['Date', 'Véhicule', 'Conducteur/Opérateur', 'Société', 'Site', 'Ouvrage', 'Litres', 'Montant', 'Compteur', 'Statut'])
                for r in rows:
                    ws.append([r['date'], r['vehicule'], r['conducteur'], r['societe'], r['site'], r['ouvrage'], r['litres'], r['montant'], r['compteur'], r['anomalie']])
                ws.append(['TOTAL', '', '', '', '', '', total_litres, total_montant, '', ''])
                bio = io.BytesIO()
                wb.save(bio)
                bio.seek(0)
                resp = HttpResponse(
                    bio.getvalue(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                resp['Content-Disposition'] = 'attachment; filename=production_gasoil_sorties.xlsx'
                return resp

            # export == 'pdf'
            bio_pdf = io.BytesIO()
            doc = SimpleDocTemplate(
                bio_pdf,
                pagesize=landscape(A4),
                leftMargin=8 * mm,
                rightMargin=8 * mm,
                topMargin=10 * mm,
                bottomMargin=10 * mm,
            )
            pdf_data = [[
                'Date', 'Véhicule', 'Conducteur', 'Société', 'Site', 'Ouvrage',
                'Litres', 'Montant', 'Cpt', 'Statut',
            ]]
            for r in rows:
                ouv = str(r.get('ouvrage') or '')[:48]
                pdf_data.append([
                    str(r.get('date') or ''),
                    str(r.get('vehicule') or ''),
                    str(r.get('conducteur') or '')[:28],
                    str(r.get('societe') or '')[:20],
                    str(r.get('site') or '')[:18],
                    ouv,
                    format_number_decimals(r.get('litres'), 2),
                    format_number_decimals(r.get('montant'), 2),
                    str(r.get('compteur') or ''),
                    str(r.get('anomalie') or ''),
                ])
            pdf_data.append([
                'TOTAL', '', '', '', '', '',
                format_number_decimals(total_litres, 2),
                format_number_decimals(total_montant, 2),
                '', '',
            ])
            tbl = Table(pdf_data, repeatRows=1)
            navy = colors.HexColor('#1A2C4E')
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), navy),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 6),
                ('ALIGN', (6, 1), (7, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#D1D5DB')),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E5E7EB')),
            ]))
            doc.build([tbl])
            resp = HttpResponse(bio_pdf.getvalue(), content_type='application/pdf')
            resp['Content-Disposition'] = 'attachment; filename=production_gasoil_sorties.pdf'
            return resp
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'production/gasoil.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'societe': societe,
        'site': site,
        'statut': statut,
        'chauffeur': chauffeur,
        'ouvrage': ouvrage,
        'activite': activite,
        'societes': societes,
        'sites': sites,
        'chauffeurs': chauffeurs,
        'ouvrages': ouvrages,
        'total_rows': total_rows,
        'total_litres': total_litres,
        'total_montant': total_montant,
        'conso_moyenne': conso_moyenne,
        'nb_anomalies': nb_anomalies,
        'gasoil_export_qs': gasoil_export_qs,
    })


def _domain_production_analytic_line(date_debut, date_fin, vehicule, product_id=None, company_id=None):
    domain = []
    if date_debut:
        domain.append(('date', '>=', date_debut))
    if date_fin:
        domain.append(('date', '<=', date_fin))
    domain.extend([
        ('transfer_consumption_id', '!=', False),
        ('transfer_consumption_id.transport_logistics', '=', False),
        ('transfer_consumption_id.service_car', '=', False),
        ('transfer_consumption_id.equipment_id.category_id', 'in', CATEGORIES_PRODUCTION),
    ])
    if vehicule:
        v = str(vehicule).strip()
        domain.extend([
            '|',
            ('transfer_consumption_id.equipment_id.name', 'ilike', v),
            ('transfer_consumption_id.affectation_id.name', 'ilike', v),
        ])
    if product_id:
        domain.append(('product_id', '=', int(product_id)))
    if company_id:
        domain.append(('company_id', '=', int(company_id)))
    return domain


@login_required
def production_couts_nature(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    nature = request.GET.get('nature', '').strip()
    vehicule = request.GET.get('vehicule', '').strip()
    product_id = request.GET.get('product_id', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    grouped_rows = []
    nature_options = []
    vehicules = []
    product_options = []
    company_options = []
    error = None
    total_montant = 0.0

    try:
        uid, models = get_odoo_connection()

        vehicules_raw = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': ['id', 'name'], 'order': 'name', 'limit': 500},
        )
        vehicules = sorted({extract_matricule(v.get('name')) for v in vehicules_raw if v.get('name')})

        base_filter_domain = _domain_production_analytic_line(date_debut, date_fin, vehicule)
        options_domain = _domain_production_analytic_line(date_debut, date_fin, None)
        domain = _domain_production_analytic_line(
            date_debut, date_fin, vehicule,
            product_id if product_id.isdigit() else None,
            company_id if company_id.isdigit() else None,
        )

        def _distinct_many2one_options(field_name, option_domain=None):
            current_domain = option_domain if option_domain is not None else base_filter_domain
            lines_for_field = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'account.analytic.line', 'search_read',
                [current_domain],
                {'fields': [field_name], 'limit': False},
            )
            seen = {}
            for ln in lines_for_field:
                val = ln.get(field_name)
                if isinstance(val, list) and len(val) >= 2:
                    seen[val[0]] = val[1]
            vals = [{'id': k, 'name': v} for k, v in seen.items()]
            vals.sort(key=lambda x: (x['name'] or '').lower())
            return vals

        product_options = _distinct_many2one_options('product_id', option_domain=options_domain)
        company_options = _distinct_many2one_options('company_id', option_domain=[])

        lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.analytic.line', 'search_read',
            [domain],
            {
                'fields': ['date', 'amount', 'nature_id', 'transfer_consumption_id', 'product_categ_id', 'general_account_id', 'account_id', 'name'],
                'order': 'date desc, id desc',
                'limit': 5000,
            },
        )

        picking_ids = list({ln['transfer_consumption_id'][0] for ln in lines if ln.get('transfer_consumption_id')})
        pickings = {}
        for i in range(0, len(picking_ids), 200):
            for p in models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.picking', 'read',
                [picking_ids[i:i + 200]],
                {'fields': ['equipment_id', 'affectation_id', 'name']},
            ):
                pickings[p['id']] = p

        def vehicule_label(p):
            if not p:
                return '—'
            if p.get('equipment_id'):
                return extract_matricule(p['equipment_id'][1])
            if p.get('affectation_id'):
                return extract_matricule(p['affectation_id'][1])
            return '—'

        def nature_label(ln):
            if ln.get('nature_id'):
                return ln['nature_id'][1]
            if ln.get('product_categ_id'):
                return ln['product_categ_id'][1]
            if ln.get('general_account_id'):
                return ln['general_account_id'][1]
            if ln.get('account_id'):
                return ln['account_id'][1]
            return ln.get('name') or '—'

        for ln in lines:
            amt = abs(float(ln.get('amount') or 0))
            pid = ln['transfer_consumption_id'][0] if ln.get('transfer_consumption_id') else None
            pk = pickings.get(pid) if pid else None
            nat = nature_label(ln)
            rows.append({
                'date': (ln.get('date') or '')[:10] or '—',
                'vehicule': vehicule_label(pk),
                'nature': nat,
                'montant': round(amt, 2),
            })

        lines_for_nature = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.analytic.line', 'search_read',
            [options_domain],
            {'fields': ['nature_id', 'product_categ_id', 'general_account_id', 'account_id', 'name'], 'limit': False},
        )
        nature_options = sorted({nature_label(ln) for ln in lines_for_nature if nature_label(ln) and nature_label(ln) != '—'})
        if nature:
            rows = [r for r in rows if r['nature'] == nature]

        total_montant = round(sum(r['montant'] for r in rows), 2)

        if group_by == 'nature':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                buckets[r['nature']]['count'] += 1
                buckets[r['nature']]['montant'] += r['montant']
            grouped_rows = [{'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)} for k, v in sorted(buckets.items(), key=lambda kv: (-kv[1]['montant'], kv[0]))]
        elif group_by == 'month':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                k = r['date'][:7] if len(r['date']) >= 7 else r['date']
                buckets[k]['count'] += 1
                buckets[k]['montant'] += r['montant']
            grouped_rows = [{'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)} for k, v in sorted(buckets.items(), key=lambda kv: kv[0], reverse=True)]
        elif group_by == 'vehicle':
            buckets = defaultdict(lambda: {'count': 0, 'montant': 0.0})
            for r in rows:
                buckets[r['vehicule']]['count'] += 1
                buckets[r['vehicule']]['montant'] += r['montant']
            grouped_rows = [{'label': k, 'count': v['count'], 'montant': round(v['montant'], 2)} for k, v in sorted(buckets.items(), key=lambda kv: (-kv[1]['montant'], kv[0]))]
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'production/couts_nature.html', {
        'rows': rows,
        'grouped_rows': grouped_rows,
        'group_by': group_by,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'nature': nature,
        'vehicule': vehicule,
        'nature_options': nature_options,
        'vehicules': vehicules,
        'product_id': product_id,
        'company_id': company_id,
        'product_options': product_options,
        'company_options': company_options,
        'total_rows': len(rows),
        'total_montant': total_montant,
    })


@login_required
def production_facturation_ventes(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    numero_facture = request.GET.get('numero_facture', '').strip()
    client_id = request.GET.get('client_id', '').strip()
    shipping_id = request.GET.get('shipping_id', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    etat = request.GET.get('etat', '').strip()
    paiement = request.GET.get('paiement', '').strip()
    commercial_id = request.GET.get('commercial_id', '').strip()
    due_date = request.GET.get('due_date', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    clients = []
    companies = []
    shippings = []
    commercials = []
    grouped_rows = []
    kpi_paid_rate = 0.0
    kpi_overdue_rate = 0.0
    kpi_ticket_moyen = 0.0
    kpi_top_client = '—'
    kpi_top_client_share = 0.0
    kpi_ca_evolution = 0.0
    error = None
    try:
        uid, models = get_odoo_connection()
        # Base commune factures de vente; séparation métier faite ensuite en Python.
        base_domain = [('move_type', '=', 'out_invoice')]

        inv_for_filters = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [base_domain],
            {'fields': ['partner_id', 'partner_shipping_id', 'company_id', 'invoice_user_id', 'project_id', 'invoice_origin', 'ref', 'name'], 'limit': False},
        )
        project_activity_map_filters = _build_project_activity_map(uid, models, inv_for_filters)
        seen_clients = {}
        seen_shippings = {}
        seen_companies = {}
        seen_commercials = {}
        # Listes de filtres complètes (toutes factures), même si les résultats
        # de la page restent strictement filtrés sur le périmètre production.
        for inv in inv_for_filters:
            p = inv.get('partner_id')
            if isinstance(p, list) and len(p) >= 2:
                seen_clients[p[0]] = p[1]
            s = inv.get('partner_shipping_id')
            if isinstance(s, list) and len(s) >= 2:
                seen_shippings[s[0]] = s[1]
            c = inv.get('company_id')
            if isinstance(c, list) and len(c) >= 2:
                seen_companies[c[0]] = c[1]
            u = inv.get('invoice_user_id')
            if isinstance(u, list) and len(u) >= 2:
                seen_commercials[u[0]] = u[1]
        clients = [{'id': cid, 'name': cname} for cid, cname in seen_clients.items()]
        clients.sort(key=lambda x: (x['name'] or '').lower())
        shippings = [{'id': sid, 'name': sname} for sid, sname in seen_shippings.items()]
        shippings.sort(key=lambda x: (x['name'] or '').lower())
        companies = [{'id': cid, 'name': cname} for cid, cname in seen_companies.items()]
        companies.sort(key=lambda x: (x['name'] or '').lower())
        commercials = [{'id': uid_, 'name': uname} for uid_, uname in seen_commercials.items()]
        commercials.sort(key=lambda x: (x['name'] or '').lower())

        domain = list(base_domain)
        if date_debut:
            domain.append(('invoice_date', '>=', date_debut))
        if date_fin:
            domain.append(('invoice_date', '<=', date_fin))
        if client_id.isdigit():
            domain.append(('partner_id', '=', int(client_id)))
        if shipping_id.isdigit():
            domain.append(('partner_shipping_id', '=', int(shipping_id)))
        if company_id.isdigit():
            domain.append(('company_id', '=', int(company_id)))
        if commercial_id.isdigit():
            domain.append(('invoice_user_id', '=', int(commercial_id)))
        if due_date:
            domain.append(('invoice_date_due', '=', due_date))
        if numero_facture:
            # Recherche élargie pour les cas où l'utilisateur saisit un code projet
            # (ex: S00068) au lieu du numéro exact de facture.
            domain += [
                '|', '|', '|',
                ('name', 'ilike', numero_facture),
                ('ref', 'ilike', numero_facture),
                ('invoice_origin', 'ilike', numero_facture),
                ('project_id.name', 'ilike', numero_facture),
            ]
        if etat in {'posted', 'draft', 'cancel'}:
            domain.append(('state', '=', etat))
        if paiement in {'paid', 'not_paid', 'in_payment'}:
            domain.append(('payment_state', '=', paiement))

        invoices = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [domain],
            {
                'fields': [
                    'name', 'invoice_date', 'partner_id', 'partner_shipping_id',
                    'company_id', 'invoice_user_id', 'invoice_date_due',
                    'amount_untaxed', 'amount_total', 'state', 'payment_state',
                    'project_id', 'invoice_origin', 'ref',
                ],
                'order': 'invoice_date desc',
                'limit': 2000,
            }
        )
        project_activity_map = _build_project_activity_map(uid, models, invoices)
        invoices = [inv for inv in invoices if _invoice_activity_bucket(inv, project_activity_map) == 'production']

        def _fmt_date_fr(iso_date):
            if not iso_date:
                return '—'
            s = str(iso_date)[:10]
            if len(s) == 10 and s[4] == '-' and s[7] == '-':
                return f'{s[8:10]}/{s[5:7]}/{s[0:4]}'
            return s

        for inv in invoices:
            iso_date = (inv.get('invoice_date') or '')[:10]
            ht = round(inv.get('amount_untaxed') or 0, 2)
            ttc = round(inv.get('amount_total') or 0, 2)
            tva = round(ttc - ht, 2)
            projet_ref = '—'
            proj = inv.get('project_id')
            if isinstance(proj, list) and len(proj) >= 2 and proj[1]:
                projet_ref = proj[1]
            elif inv.get('invoice_origin'):
                projet_ref = inv.get('invoice_origin')
            elif inv.get('ref'):
                projet_ref = inv.get('ref')
            rows.append({
                'numero': inv.get('name') or '—',
                'date': _fmt_date_fr(iso_date),
                'month_key': iso_date[:7] if len(iso_date) >= 7 else '—',
                'projet_reference': projet_ref,
                'client': inv['partner_id'][1] if inv.get('partner_id') else '—',
                'lieu_livraison': inv['partner_shipping_id'][1] if inv.get('partner_shipping_id') else '—',
                'societe': inv['company_id'][1] if inv.get('company_id') else '—',
                'ht': ht,
                'tva': tva,
                'ttc': ttc,
                'payment_state': inv.get('payment_state') or '',
                'invoice_date_due': (inv.get('invoice_date_due') or '')[:10],
            })

        if group_by:
            buckets = defaultdict(lambda: {'count': 0, 'ht': 0.0, 'tva': 0.0, 'ttc': 0.0})
            for r in rows:
                if group_by == 'month':
                    key = r['month_key']
                elif group_by == 'company':
                    key = r['societe']
                elif group_by == 'month_company':
                    key = f'{r["month_key"]} | {r["societe"]}'
                elif group_by == 'month_delivery':
                    key = f'{r["month_key"]} | {r["lieu_livraison"]}'
                else:
                    key = '—'
                buckets[key]['count'] += 1
                buckets[key]['ht'] += r['ht']
                buckets[key]['tva'] += r['tva']
                buckets[key]['ttc'] += r['ttc']

            mois_fr = {
                '01': 'Janvier', '02': 'Fevrier', '03': 'Mars', '04': 'Avril',
                '05': 'Mai', '06': 'Juin', '07': 'Juillet', '08': 'Aout',
                '09': 'Septembre', '10': 'Octobre', '11': 'Novembre', '12': 'Decembre',
            }

            def _month_label(ym):
                if isinstance(ym, str) and len(ym) == 7 and ym[4] == '-':
                    return f"{mois_fr.get(ym[5:7], ym[5:7])} {ym[:4]}"
                return ym

            grouped_rows = [
                {
                    'label': k,
                    'sort_key': k,
                    'count': v['count'],
                    'ht': round(v['ht'], 2),
                    'tva': round(v['tva'], 2),
                    'ttc': round(v['ttc'], 2),
                }
                for k, v in buckets.items()
            ]
            if group_by == 'month':
                for g in grouped_rows:
                    g['label'] = _month_label(g['sort_key'])
                grouped_rows.sort(key=lambda x: x['sort_key'])
            elif group_by in {'month_company', 'month_delivery'}:
                for g in grouped_rows:
                    ym, suffix = str(g['sort_key']).split(' | ', 1)
                    g['label'] = f'{_month_label(ym)} | {suffix}'
                grouped_rows.sort(key=lambda x: x['sort_key'])
            else:
                grouped_rows.sort(key=lambda x: (-x['ttc'], x['label']))
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_ht = round(sum(r['ht'] for r in rows), 2)
    total_tva = round(sum(r['tva'] for r in rows), 2)
    total_ttc = round(sum(r['ttc'] for r in rows), 2)
    total_rows = len(rows)
    kpi_ticket_moyen = round((total_ttc / total_rows), 2) if total_rows else 0.0

    paid_count = sum(1 for r in rows if r.get('payment_state') == 'paid')
    kpi_paid_rate = round((paid_count / total_rows) * 100, 1) if total_rows else 0.0

    today_iso = date.today().isoformat()
    overdue_count = sum(
        1 for r in rows
        if (r.get('invoice_date_due') and r.get('invoice_date_due') < today_iso and r.get('payment_state') != 'paid')
    )
    kpi_overdue_rate = round((overdue_count / total_rows) * 100, 1) if total_rows else 0.0

    # KPI client leader (part du CA TTC)
    client_totals = defaultdict(float)
    for r in rows:
        client_totals[r.get('client') or '—'] += float(r.get('ttc') or 0)
    if client_totals and total_ttc > 0:
        best_client, best_value = max(client_totals.items(), key=lambda kv: kv[1])
        kpi_top_client = best_client
        kpi_top_client_share = round((best_value / total_ttc) * 100, 1)

    # KPI évolution CA mensuel (% M vs M-1)
    monthly_totals = defaultdict(float)
    for r in rows:
        mk = r.get('month_key') or '—'
        if mk != '—':
            monthly_totals[mk] += float(r.get('ttc') or 0)
    if len(monthly_totals) >= 2:
        keys = sorted(monthly_totals.keys())
        prev_v = monthly_totals[keys[-2]]
        curr_v = monthly_totals[keys[-1]]
        if prev_v > 0:
            kpi_ca_evolution = round(((curr_v - prev_v) / prev_v) * 100, 1)

    delivery_buckets = defaultdict(lambda: {'count': 0, 'ht': 0.0, 'tva': 0.0, 'ttc': 0.0})
    for r in rows:
        k = r['lieu_livraison']
        delivery_buckets[k]['count'] += 1
        delivery_buckets[k]['ht'] += r['ht']
        delivery_buckets[k]['tva'] += r['tva']
        delivery_buckets[k]['ttc'] += r['ttc']
    delivery_rows = [
        {
            'lieu': k,
            'count': v['count'],
            'ht': round(v['ht'], 2),
            'tva': round(v['tva'], 2),
            'ttc': round(v['ttc'], 2),
        }
        for k, v in sorted(delivery_buckets.items(), key=lambda kv: (-kv[1]['ttc'], kv[0]))
    ]
    delivery_total_count = sum(r['count'] for r in delivery_rows)
    delivery_total_ht = round(sum(r['ht'] for r in delivery_rows), 2)
    delivery_total_tva = round(sum(r['tva'] for r in delivery_rows), 2)
    delivery_total_ttc = round(sum(r['ttc'] for r in delivery_rows), 2)

    return render(request, 'production/facturation_ventes.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'numero_facture': numero_facture,
        'client_id': client_id,
        'shipping_id': shipping_id,
        'company_id': company_id,
        'etat': etat,
        'paiement': paiement,
        'commercial_id': commercial_id,
        'due_date': due_date,
        'clients': clients,
        'shippings': shippings,
        'companies': companies,
        'commercials': commercials,
        'group_by': group_by,
        'grouped_rows': grouped_rows,
        'total_rows': total_rows,
        'total_ht': total_ht,
        'total_tva': total_tva,
        'total_ttc': total_ttc,
        'kpi_paid_rate': kpi_paid_rate,
        'kpi_overdue_rate': kpi_overdue_rate,
        'kpi_ticket_moyen': kpi_ticket_moyen,
        'kpi_top_client': kpi_top_client,
        'kpi_top_client_share': kpi_top_client_share,
        'kpi_ca_evolution': kpi_ca_evolution,
        'delivery_rows': delivery_rows,
        'delivery_total_count': delivery_total_count,
        'delivery_total_ht': delivery_total_ht,
        'delivery_total_tva': delivery_total_tva,
        'delivery_total_ttc': delivery_total_ttc,
    })


@login_required
def production_sites(request):
    date_debut = request.GET.get('date_debut', '').strip()
    date_fin = request.GET.get('date_fin', '').strip()
    site = request.GET.get('site', '').strip()
    societe = request.GET.get('societe', '').strip()
    tri = request.GET.get('tri', 'optimisation').strip() or 'optimisation'

    rows = []
    sites = []
    societes = []
    kpi_sites = 0
    kpi_bons = 0
    kpi_litres = 0.0
    kpi_montant = 0.0
    kpi_anomalies = 0
    chart_labels = []
    chart_values = []
    top_site_name = '—'
    top_site_montant = 0.0
    top_site_score = 0.0
    worst_site_name = '—'
    worst_site_ratio = 0.0
    no_data_for_filters = False
    error = None

    try:
        uid, models = get_odoo_connection()

        list_domain = _build_sorties_domain(
            date_debut='',
            date_fin='',
            site='',
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe='',
            categorie_engin='',
            activite_filtre='production',
        )
        list_domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons_for_filters = _fetch_sorties_bons(uid, models, list_domain, limit=6000)
        bons_for_filters = [b for b in bons_for_filters if b.get('activite_bucket') == 'production']
        sites = sorted({b['site'] for b in bons_for_filters if b.get('site') and b.get('site') != '—'})
        societes = sorted({b['societe'] for b in bons_for_filters if b.get('societe') and b.get('societe') != '—'})

        domain = _build_sorties_domain(
            date_debut=date_debut,
            date_fin=date_fin,
            site=site,
            chauffeur='',
            ouvrage='',
            anomalie='',
            societe=societe,
            categorie_engin='',
            activite_filtre='production',
        )
        domain.append(('equipment_id.category_id', 'in', CATEGORIES_PRODUCTION))
        bons = _fetch_sorties_bons(uid, models, domain, limit=12000)
        bons = [b for b in bons if b.get('activite_bucket') == 'production']

        move_domain = [
            ['picking_id', 'in', [b['id'] for b in bons]],
            ['product_id.categ_id', '=', CARBURANT_CATEG_ID],
        ]
        move_lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.move', 'search_read',
            [move_domain],
            {'fields': ['picking_id', 'product_qty', 'price_total', 'unit_price'], 'limit': 30000},
        )
        amount_by_picking = defaultdict(float)
        for mv in move_lines:
            pid = mv['picking_id'][0] if mv.get('picking_id') else None
            if not pid:
                continue
            price_total = mv.get('price_total')
            if price_total not in (None, False):
                amount_by_picking[pid] += price_total or 0
            else:
                amount_by_picking[pid] += (mv.get('product_qty') or 0) * (mv.get('unit_price') or 0)

        site_bucket = defaultdict(lambda: {
            'site': '—', 'societes': set(), 'nb_bons': 0, 'litres': 0.0, 'montant': 0.0, 'anomalies': 0
        })
        for b in bons:
            site_name = b.get('site') or '—'
            cur = site_bucket[site_name]
            cur['site'] = site_name
            cur['nb_bons'] += 1
            cur['litres'] += float(b.get('product_qty') or 0)
            cur['montant'] += float(amount_by_picking.get(b['id'], 0))
            cur['anomalies'] += 1 if b.get('anomalie') == 'Anomalie' else 0
            if b.get('societe') and b.get('societe') != '—':
                cur['societes'].add(b['societe'])

        rows = []
        for _, item in site_bucket.items():
            litres = round(item['litres'], 2)
            montant = round(item['montant'], 2)
            nb_bons = item['nb_bons']
            anomalies = item['anomalies']
            ratio_anomalies = round((anomalies / nb_bons) * 100, 1) if nb_bons else 0
            litres_par_bon = round((litres / nb_bons), 2) if nb_bons else 0.0
            rows.append({
                'site': item['site'],
                'societe': ', '.join(sorted(item['societes'])) if item['societes'] else '—',
                'nb_bons': nb_bons,
                'litres': litres,
                'montant': montant,
                'anomalies': anomalies,
                'ratio_anomalies': ratio_anomalies,
                'litres_par_bon': litres_par_bon,
                'optimisation_score': 0.0,
            })

        # ── Score d'optimisation pondéré (explicite) ─────────────────────────
        # 60% consommation (litres/bon bas = meilleur)
        # 30% anomalies (taux bas = meilleur)
        # 10% volume d'activité (plus de bons = meilleur)
        if rows:
            max_lpb = max(r['litres_par_bon'] for r in rows) or 1.0
            max_anom = max(r['ratio_anomalies'] for r in rows) or 1.0
            max_bons = max(r['nb_bons'] for r in rows) or 1.0

            for r in rows:
                consommation_norm = max(0.0, 100 * (1 - (r['litres_par_bon'] / max_lpb)))
                anomalies_norm = max(0.0, 100 * (1 - (r['ratio_anomalies'] / max_anom))) if max_anom > 0 else 100.0
                volume_norm = max(0.0, 100 * (r['nb_bons'] / max_bons))
                score = (0.60 * consommation_norm) + (0.30 * anomalies_norm) + (0.10 * volume_norm)
                r['optimisation_score'] = round(score, 2)

        if tri == 'optimisation':
            rows.sort(key=lambda r: (-r['optimisation_score'], r['site']))
        elif tri == 'bons':
            rows.sort(key=lambda r: (-r['nb_bons'], r['site']))
        elif tri == 'litres':
            rows.sort(key=lambda r: (-r['litres'], r['site']))
        elif tri == 'anomalies':
            rows.sort(key=lambda r: (-r['anomalies'], r['site']))
        else:
            rows.sort(key=lambda r: (-r['montant'], r['site']))

        kpi_sites = len(rows)
        kpi_bons = sum(r['nb_bons'] for r in rows)
        kpi_litres = round(sum(r['litres'] for r in rows), 2)
        kpi_montant = round(sum(r['montant'] for r in rows), 2)
        kpi_anomalies = sum(r['anomalies'] for r in rows)

        # Top 8 toujours basé sur l'optimisation (pas sur le montant)
        top_chart = sorted(rows, key=lambda r: (-r['optimisation_score'], r['site']))[:8]
        chart_labels = [r['site'] for r in top_chart]
        chart_values = [r['optimisation_score'] for r in top_chart]

        if rows:
            best = max(rows, key=lambda r: (r['optimisation_score'], r['nb_bons']))
            top_site_name = best['site']
            top_site_montant = best['montant']
            top_site_score = best['optimisation_score']
            worst = max(rows, key=lambda r: (r['ratio_anomalies'], r['anomalies']))
            worst_site_name = worst['site']
            worst_site_ratio = worst['ratio_anomalies']
        no_data_for_filters = len(rows) == 0
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'production/sites.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'site': site,
        'societe': societe,
        'tri': tri,
        'sites': sites,
        'societes': societes,
        'kpi_sites': kpi_sites,
        'kpi_bons': kpi_bons,
        'kpi_litres': kpi_litres,
        'kpi_montant': kpi_montant,
        'kpi_anomalies': kpi_anomalies,
        'top_site_name': top_site_name,
        'top_site_montant': top_site_montant,
        'top_site_score': top_site_score,
        'worst_site_name': worst_site_name,
        'worst_site_ratio': worst_site_ratio,
        'no_data_for_filters': no_data_for_filters,
        'chart_has_data': len(chart_labels) > 0,
        'chart_labels_json': json.dumps(chart_labels),
        'chart_values_json': json.dumps(chart_values),
    })


def _render_achats_module(request, current_key, page_title, page_subtitle):
    menu_items = [
        {'key': 'overview', 'label': "Vue d'ensemble", 'url': 'achats_overview'},
        {'key': 'purchase_requests', 'label': "Demandes d'achat", 'url': 'achats_purchase_requests'},
        {'key': 'rfq', 'label': 'Demandes de prix', 'url': 'achats_rfq'},
        {'key': 'purchase_orders', 'label': 'Bons de commande', 'url': 'achats_purchase_orders'},
        {'key': 'delivery_tracking', 'label': 'Suivi livraisons', 'url': 'achats_delivery_tracking'},
        {'key': 'suppliers', 'label': 'Fournisseurs', 'url': 'achats_suppliers'},
    ]

    return render(request, 'achats/module.html', {
        'page_title': page_title,
        'page_subtitle': page_subtitle,
        'module_key': current_key,
        'menu_items': menu_items,
    })


@login_required
def achats_overview(request):
    return _render_achats_module(
        request,
        current_key='overview',
        page_title="Vue d'ensemble",
        page_subtitle="Vision globale du flux achats : demandes, commandes, livraisons et fournisseurs.",
    )


@login_required
def achats_purchase_requests(request):
    date_debut  = request.GET.get('date_debut',  '').strip()
    date_fin    = request.GET.get('date_fin',    '').strip()
    departement = request.GET.get('departement', '').strip()
    etat        = request.GET.get('etat',        '').strip()
    demandeur   = request.GET.get('demandeur',   '').strip()
    export      = request.GET.get('export',      '').strip().lower()

    STATE_LABELS_PR = {
        'draft':      'Brouillon',
        'to_approve': 'Confirmé',
        'confirmed':  'Confirmé',
        'approved':   'Approuvé',
        'rejected':   'Rejeté',
        'done':       'Terminé',
        'sent':       'Envoyé',
        'purchase':   'Bon de commande',
        'cancel':     'Annulé',
    }

    rows = []
    error = None
    departements = []
    demandeurs_list = []
    total_demandes = 0
    total_montant  = 0.0
    nb_attente     = 0
    nb_approuvees  = 0
    use_fallback   = False
    odoo_model_label = 'purchase.request'

    def _fmt_date_pr(d):
        if not d or len(d) < 10:
            return '—'
        return f'{d[8:10]}/{d[5:7]}/{d[0:4]}'

    state_options = [
        ('draft', 'Brouillon'),
        ('to_approve', 'Confirmé'),
        ('approved', 'Approuvé'),
        ('rejected', 'Rejeté'),
    ]

    try:
        uid, models = get_odoo_connection()

        # ── Essai purchase.request ──────────────────────────────
        pr_exists = False
        try:
            models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.request', 'search_count', [[]], {},
            )
            pr_exists = True
        except Exception:
            pr_exists = False

        if pr_exists:
            fields_meta = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.request', 'fields_get', [], {'attributes': ['type', 'string', 'relation']},
            ) or {}
            available_fields = set(fields_meta.keys())

            date_candidates = [f for f in ('date_start', 'date_required', 'date', 'create_date') if f in available_fields]
            requester_candidates = [f for f in ('requested_by', 'requested_by_id', 'user_id', 'create_uid') if f in available_fields]
            department_candidates = [f for f in ('department_id', 'x_department_id', 'service_id', 'company_id') if f in available_fields]
            description_candidates = [f for f in ('description', 'origin', 'notes') if f in available_fields]
            amount_candidates = [f for f in ('estimated_cost', 'amount_total', 'amount_untaxed') if f in available_fields]
            line_field = next((f for f in ('line_ids', 'request_line_ids', 'order_line') if f in available_fields), None)

            date_field = date_candidates[0] if date_candidates else None

            domain = []
            if date_debut and date_field:
                domain.append((date_field, '>=', date_debut))
            if date_fin and date_field:
                domain.append((date_field, '<=', date_fin + ' 23:59:59'))
            if etat and 'state' in available_fields:
                domain.append(('state', '=', etat))

            read_fields = ['name']
            read_fields.extend(date_candidates)
            read_fields.extend(requester_candidates)
            read_fields.extend(department_candidates)
            read_fields.extend(description_candidates)
            read_fields.extend(amount_candidates)
            if 'state' in available_fields:
                read_fields.append('state')
            if line_field:
                read_fields.append(line_field)
            read_fields = sorted(set(read_fields))

            _pr_order = f'{date_field} desc' if date_field else 'id desc'
            records = odoo_search_read_all(
                uid, models, 'purchase.request', domain, read_fields, order=_pr_order,
            )

            # Si la description principale est vide, on tente de la reconstruire depuis les lignes.
            line_desc_map = {}
            if line_field:
                try:
                    line_ids = []
                    for rec in records:
                        vals = rec.get(line_field) or []
                        if isinstance(vals, list):
                            line_ids.extend([int(v) for v in vals if isinstance(v, int)])
                    line_ids = sorted(set(line_ids))
                    line_model = (fields_meta.get(line_field) or {}).get('relation') or 'purchase.request.line'
                    if line_ids and line_model:
                        line_fields_meta = models.execute_kw(
                            settings.ODOO_DB, uid, settings.ODOO_PASS,
                            line_model, 'fields_get', [], {'attributes': ['type']},
                        ) or {}
                        line_available = set(line_fields_meta.keys())
                        line_read_fields = [f for f in ('name', 'description', 'product_id', 'product_qty', 'product_uom_qty') if f in line_available]
                        line_read_fields = sorted(set(line_read_fields)) if line_read_fields else ['id']
                        line_records = models.execute_kw(
                            settings.ODOO_DB, uid, settings.ODOO_PASS,
                            line_model, 'read',
                            [line_ids],
                            {'fields': line_read_fields},
                        )
                        for ln in line_records:
                            txt = ''
                            if 'description' in ln and ln.get('description'):
                                txt = str(ln.get('description')).strip()
                            elif 'name' in ln and ln.get('name'):
                                txt = str(ln.get('name')).strip()
                            if not txt and isinstance(ln.get('product_id'), list) and len(ln.get('product_id')) > 1:
                                qty = ln.get('product_qty') if ln.get('product_qty') not in (None, '') else ln.get('product_uom_qty')
                                txt = f"{ln['product_id'][1]} x {qty}" if qty not in (None, '') else str(ln['product_id'][1])
                            line_desc_map[ln.get('id')] = txt
                except Exception:
                    line_desc_map = {}

            all_rows = []
            for r in records:
                date_raw = ''
                for f in date_candidates:
                    val = r.get(f)
                    if val:
                        date_raw = str(val)[:10]
                        break

                req = None
                for f in requester_candidates:
                    val = r.get(f)
                    if val:
                        req = val
                        break

                dept = None
                for f in department_candidates:
                    val = r.get(f)
                    if val:
                        dept = val
                        break

                desc_val = None
                for f in description_candidates:
                    val = r.get(f)
                    if val:
                        desc_val = val
                        break
                if not desc_val and line_field:
                    line_ids_row = r.get(line_field) or []
                    if isinstance(line_ids_row, list):
                        line_texts = [line_desc_map.get(i) for i in line_ids_row if line_desc_map.get(i)]
                        if line_texts:
                            desc_val = ' | '.join(line_texts[:2])

                amount_val = 0
                for f in amount_candidates:
                    val = r.get(f)
                    if val not in (None, ''):
                        amount_val = val
                        break

                dept_name = dept[1] if isinstance(dept, list) and len(dept) > 1 else '—'
                req_name = req[1] if isinstance(req, list) and len(req) > 1 else str(req or '—')
                all_rows.append({
                    'name':        r.get('name') or '—',
                    'date':        _fmt_date_pr(date_raw),
                    'date_raw':    date_raw,
                    'demandeur':   req_name,
                    'departement': dept_name,
                    'description': (str(desc_val or '—'))[:140],
                    'montant':     float(amount_val or 0),
                    'etat_raw':    r.get('state') or 'draft',
                    'etat':        STATE_LABELS_PR.get(r.get('state') or 'draft', r.get('state') or '—'),
                })

            departements    = sorted({row['departement'] for row in all_rows if row['departement'] != '—'})
            demandeurs_list = sorted({row['demandeur']   for row in all_rows if row['demandeur']   != '—'})

            # Filtres côté Python
            rows = all_rows
            if departement:
                rows = [r for r in rows if departement.lower() in r['departement'].lower()]
            if demandeur:
                rows = [r for r in rows if demandeur.lower() in r['demandeur'].lower()]

        else:
            # ── Fallback purchase.order (draft/sent) ───────────
            po_exists = False
            try:
                models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'purchase.order', 'search_count', [[]], {},
                )
                po_exists = True
            except Exception:
                po_exists = False

            if not po_exists:
                error = "Modèle non disponible dans cette instance Odoo"
                return render(request, 'achats/demandes_achat.html', {
                    'rows': [],
                    'error': error,
                    'date_debut': date_debut,
                    'date_fin': date_fin,
                    'departement': departement,
                    'etat': etat,
                    'demandeur': demandeur,
                    'departements': [],
                    'demandeurs': [],
                    'total_demandes': 0,
                    'total_montant': 0.0,
                    'nb_attente': 0,
                    'nb_approuvees': 0,
                    'use_fallback': True,
                    'odoo_model_label': 'purchase.request / purchase.order',
                    'state_options': state_options,
                })

            use_fallback = True
            odoo_model_label = 'purchase.order (draft/sent)'
            domain = [('state', 'in', ['draft', 'sent'])]
            if date_debut:
                domain.append(('date_order', '>=', date_debut))
            if date_fin:
                domain.append(('date_order', '<=', date_fin + ' 23:59:59'))
            if etat and etat in ('draft', 'sent'):
                domain.append(('state', '=', etat))

            records = odoo_search_read_all(
                uid, models, 'purchase.order', domain,
                ['name', 'date_order', 'partner_id', 'amount_total',
                 'state', 'user_id', 'notes'],
                order='date_order desc',
            )
            all_rows = []
            for r in records:
                partner   = r.get('partner_id')
                dept_name = partner[1] if isinstance(partner, list) and len(partner) > 1 else '—'
                user      = r.get('user_id')
                user_name = user[1]    if isinstance(user,    list) and len(user)    > 1 else '—'
                all_rows.append({
                    'name':        r.get('name') or '—',
                    'date':        _fmt_date_pr((r.get('date_order') or '')[:10]),
                    'date_raw':    (r.get('date_order') or '')[:10],
                    'demandeur':   user_name,
                    'departement': dept_name,
                    'description': (r.get('notes') or '—')[:120],
                    'montant':     float(r.get('amount_total') or 0),
                    'etat_raw':    r.get('state') or 'draft',
                    'etat':        STATE_LABELS_PR.get(r.get('state') or 'draft', '—'),
                })

            departements    = sorted({row['departement'] for row in all_rows if row['departement'] != '—'})
            demandeurs_list = sorted({row['demandeur']   for row in all_rows if row['demandeur']   != '—'})

            rows = all_rows
            if departement:
                rows = [r for r in rows if departement.lower() in r['departement'].lower()]
            if demandeur:
                rows = [r for r in rows if demandeur.lower() in r['demandeur'].lower()]

        # ── KPI ────────────────────────────────────────────────
        total_demandes = len(rows)
        total_montant  = round(sum(r['montant'] for r in rows), 2)
        nb_attente     = sum(1 for r in rows if r['etat_raw'] in ('draft', 'to_approve', 'sent'))
        nb_approuvees  = sum(1 for r in rows if r['etat_raw'] in ('approved', 'purchase', 'done'))
        # ── Exports ────────────────────────────────────────────
        if export == 'csv':
            out = io.StringIO()
            w = csv.writer(out, delimiter=';')
            w.writerow(['N° Demande', 'Date', 'Demandeur', 'Département', 'Description', 'Montant estimé', 'État'])
            for r in rows:
                w.writerow([r['name'], r['date'], r['demandeur'], r['departement'],
                             r['description'], r['montant'], r['etat']])
            resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8-sig')
            resp['Content-Disposition'] = 'attachment; filename=demandes_achat.csv'
            return resp

        if export == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Demandes d'achat"
            header = ['N° Demande', 'Date', 'Demandeur', 'Département', 'Description', 'Montant estimé', 'État']
            ws.append(header)
            hdr_fill = PatternFill('solid', fgColor='1A2C4E')
            hdr_font = Font(color='FFFFFF', bold=True)
            for cell in ws[1]:
                cell.fill = hdr_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for r in rows:
                ws.append([r['name'], r['date'], r['demandeur'], r['departement'],
                            r['description'], r['montant'], r['etat']])
            ws.append(['TOTAL', '', '', '', '', total_montant, ''])

            last_row = ws.max_row
            total_fill = PatternFill('solid', fgColor='1A2C4E')
            total_font = Font(color='FFFFFF', bold=True)
            for cell in ws[last_row]:
                cell.fill = total_fill
                cell.font = total_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Mise en forme des colonnes pour un rendu lisible dans Excel.
            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 14
            ws.column_dimensions['C'].width = 24
            ws.column_dimensions['D'].width = 24
            ws.column_dimensions['E'].width = 52
            ws.column_dimensions['F'].width = 18
            ws.column_dimensions['G'].width = 16

            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = f'A1:G{last_row}'

            for row_idx in range(2, last_row):
                ws[f'F{row_idx}'].number_format = '#,##0.00'
                ws[f'F{row_idx}'].alignment = Alignment(horizontal='right')
                ws[f'B{row_idx}'].alignment = Alignment(horizontal='center')

            ws[f'F{last_row}'].number_format = '#,##0.00'
            ws[f'F{last_row}'].alignment = Alignment(horizontal='right')

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            resp = HttpResponse(
                bio.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            resp['Content-Disposition'] = 'attachment; filename=demandes_achat.xlsx'
            return resp

        if export == 'pdf':
            bio_pdf = io.BytesIO()
            logo_path = settings.BASE_DIR / 'static' / 'images' / 'logo_somatrin.png'
            pw, ph = landscape(A4)

            class _PagedCanvasDA(rl_canvas.Canvas):
                def __init__(self, *args, **kw):
                    super().__init__(*args, **kw)
                    self._saved_states = []
                def showPage(self):
                    self._saved_states.append(dict(self.__dict__))
                    self._startPage()
                def save(self):
                    total = len(self._saved_states)
                    for state in self._saved_states:
                        self.__dict__.update(state)
                        self.setFont('Helvetica', 7.5)
                        self.setFillColor(colors.HexColor('#6B7280'))
                        self.drawRightString(pw - 15 * mm, 10 * mm,
                                             f'Page {self._pageNumber} / {total}')
                        super().showPage()
                    super().save()

            def _header_da(c, d):
                c.saveState()
                logo_x = 15 * mm
                logo_y = ph - 22 * mm
                logo_w = 32 * mm
                logo_h = 11 * mm
                if logo_path.is_file():
                    c.drawImage(str(logo_path), logo_x, logo_y,
                                width=logo_w, height=logo_h,
                                preserveAspectRatio=True, mask='auto')
                # Alignement vertical sur le centre du logo
                header_y = logo_y + (logo_h / 2.0) - (3 * mm)
                c.setFont('Helvetica-Bold', 11)
                c.setFillColor(colors.HexColor('#1A2C4E'))
                c.drawString(logo_x + logo_w + (6 * mm), header_y, "Demandes d'achat — SOMATRIN")
                c.setFont('Helvetica', 8)
                c.setFillColor(colors.HexColor('#6B7280'))
                c.drawRightString(pw - 15 * mm, header_y, 'Document Confidentiel')
                c.setStrokeColor(colors.HexColor('#E87722'))
                c.setLineWidth(1.2)
                c.line(15 * mm, ph - 26 * mm, pw - 15 * mm, ph - 26 * mm)
                c.restoreState()

            wrap_da = ParagraphStyle('wda', fontSize=7, leading=9)
            navy_da = colors.HexColor('#1A2C4E')
            lgray   = colors.HexColor('#F8FAFC')

            pdf_data = [['N° Demande', 'Date', 'Demandeur', 'Département',
                          'Description', 'Montant estimé', 'État']]
            for r in rows:
                pdf_data.append([
                    r['name'], r['date'], r['demandeur'], r['departement'],
                    Paragraph(str(r['description'])[:80], wrap_da),
                    f"{r['montant']:,.2f}", r['etat'],
                ])
            pdf_data.append(['TOTAL', '', '', '', '',
                              f"{total_montant:,.2f}", f"{total_demandes} demande(s)"])

            cw_da = [35 * mm, 22 * mm, 40 * mm, 40 * mm, 68 * mm, 30 * mm, 22 * mm]
            tbl_da = Table(pdf_data, colWidths=cw_da, repeatRows=1)
            tbl_da.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0),  (-1, 0),  navy_da),
                ('TEXTCOLOR',     (0, 0),  (-1, 0),  colors.white),
                ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0, 0),  (-1, 0),  8),
                ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
                ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
                ('FONTNAME',      (0, 1),  (-1, -2), 'Helvetica'),
                ('FONTSIZE',      (0, 1),  (-1, -2), 7.5),
                ('ROWBACKGROUNDS',(0, 1),  (-1, -2), [colors.white, lgray]),
                ('ALIGN',         (5, 1),  (5, -2),  'RIGHT'),
                ('GRID',          (0, 0),  (-1, -1), 0.3, colors.HexColor('#E5E7EB')),
                ('BACKGROUND',    (0, -1), (-1, -1), navy_da),
                ('TEXTCOLOR',     (0, -1), (-1, -1), colors.white),
                ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE',      (0, -1), (-1, -1), 8),
                ('ALIGN',         (5, -1), (5, -1),  'RIGHT'),
            ]))

            doc_da = SimpleDocTemplate(
                bio_pdf, pagesize=landscape(A4),
                leftMargin=15 * mm, rightMargin=15 * mm,
                topMargin=32 * mm, bottomMargin=22 * mm,
            )
            doc_da.build([tbl_da], onFirstPage=_header_da, onLaterPages=_header_da,
                         canvasmaker=_PagedCanvasDA)
            bio_pdf.seek(0)
            resp = HttpResponse(bio_pdf.getvalue(), content_type='application/pdf')
            resp['Content-Disposition'] = "attachment; filename=demandes_achat.pdf"
            return resp

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'achats/demandes_achat.html', {
        'rows':             rows,
        'error':            error,
        'date_debut':       date_debut,
        'date_fin':         date_fin,
        'departement':      departement,
        'etat':             etat,
        'demandeur':        demandeur,
        'departements':     departements,
        'demandeurs':       demandeurs_list,
        'total_demandes':   total_demandes,
        'total_montant':    total_montant,
        'nb_attente':       nb_attente,
        'nb_approuvees':    nb_approuvees,
        'use_fallback':     use_fallback,
        'odoo_model_label': odoo_model_label,
        'state_options':    state_options,
    })


@login_required
def achats_rfq(request):
    date_debut  = request.GET.get('date_debut',   '').strip()
    date_fin    = request.GET.get('date_fin',     '').strip()
    fournisseur = request.GET.get('fournisseur',  '').strip()
    etat        = request.GET.get('etat',         '').strip()
    responsable = request.GET.get('responsable',  '').strip()
    export      = request.GET.get('export',       '').strip().lower()

    STATE_LABELS_RFQ = {
        'draft':    'Brouillon',
        'sent':     'Envoyé',
        'purchase': 'Bon de commande',
        'cancel':   'Annulé',
    }

    rows = []
    error = None
    fournisseurs = []
    responsables = []
    total_dp     = 0
    total_montant = 0.0
    nb_attente    = 0
    nb_expirees   = 0

    def _fmt_date_rfq(d):
        if not d or len(d) < 10:
            return '—'
        return f'{d[8:10]}/{d[5:7]}/{d[0:4]}'

    try:
        from datetime import datetime as _dt, timedelta as _td
        uid, models = get_odoo_connection()

        po_exists = False
        try:
            models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order', 'search_count', [[]], {},
            )
            po_exists = True
        except Exception:
            po_exists = False

        if not po_exists:
            return render(request, 'achats/demandes_prix.html', {
                'rows': [],
                'error': "Modèle non disponible dans cette instance Odoo",
                'date_debut': date_debut,
                'date_fin': date_fin,
                'fournisseur': fournisseur,
                'etat': etat,
                'responsable': responsable,
                'fournisseurs': [],
                'responsables': [],
                'total_dp': 0,
                'total_montant': 0.0,
                'nb_attente': 0,
                'nb_expirees': 0,
            })

        domain = []
        if date_debut:
            domain.append(('date_order', '>=', date_debut))
        if date_fin:
            domain.append(('date_order', '<=', date_fin + ' 23:59:59'))
        if etat:
            domain.append(('state', '=', etat))

        # Champs disponibles (pour utiliser une vraie date d'expiration si présente)
        po_fields = {}
        try:
            po_fields = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order', 'fields_get', [], {'attributes': ['type']},
            ) or {}
        except Exception:
            po_fields = {}
        po_available = set(po_fields.keys())

        read_fields = ['name', 'date_order', 'partner_id', 'amount_total', 'state', 'currency_id', 'user_id']
        # Certaines instances Odoo exposent une date de validité/expiration des RFQ
        if 'validity_date' in po_available:
            read_fields.append('validity_date')

        records = odoo_search_read_all(
            uid, models, 'purchase.order', domain, read_fields, order='date_order desc',
        )

        today = date.today()
        all_rows = []
        for r in records:
            partner    = r.get('partner_id')
            fournisseur_name = partner[1] if isinstance(partner, list) and len(partner) > 1 else '—'
            user       = r.get('user_id')
            user_name  = user[1]    if isinstance(user,    list) and len(user)    > 1 else '—'
            currency   = r.get('currency_id')
            devise     = currency[1] if isinstance(currency, list) and len(currency) > 1 else 'MAD'

            raw_date = (r.get('date_order') or '')[:10]
            validity_raw = (r.get('validity_date') or '')[:10] if 'validity_date' in po_available else ''
            expire_raw = ''
            is_expired = False
            if raw_date and len(raw_date) == 10:
                try:
                    order_date = _dt.strptime(raw_date, '%Y-%m-%d').date()
                    if validity_raw and len(validity_raw) == 10:
                        exp_date = _dt.strptime(validity_raw, '%Y-%m-%d').date()
                        expire_raw = validity_raw
                    else:
                        exp_date = order_date + _td(days=15)
                        expire_raw = exp_date.isoformat()
                    if r.get('state') == 'sent':
                        is_expired = today > exp_date
                except Exception:
                    pass

            all_rows.append({
                'name':        r.get('name') or '—',
                'date':        raw_date,
                'expire_raw':  expire_raw,
                'has_validity_date': bool(validity_raw),
                'fournisseur': fournisseur_name,
                'responsable': user_name,
                'montant':     float(r.get('amount_total') or 0),
                'devise':      devise,
                'etat_raw':    r.get('state') or 'draft',
                'etat':        STATE_LABELS_RFQ.get(r.get('state') or 'draft', '—'),
                'is_expired':  is_expired,
            })

        fournisseurs = sorted({row['fournisseur'] for row in all_rows if row['fournisseur'] != '—'})
        responsables = sorted({row['responsable'] for row in all_rows if row['responsable'] != '—'})

        # Filtres côté Python
        rows = all_rows
        if fournisseur:
            rows = [r for r in rows if fournisseur.lower() in r['fournisseur'].lower()]
        if responsable:
            rows = [r for r in rows if responsable.lower() in r['responsable'].lower()]

        # ── KPI ────────────────────────────────────────────────
        total_dp      = len(rows)
        total_montant = round(sum(r['montant'] for r in rows), 2)
        # "En attente réponse fournisseur" = RFQ envoyées au fournisseur.
        nb_attente    = sum(1 for r in rows if r['etat_raw'] == 'sent')
        nb_expirees   = sum(1 for r in rows if r.get('is_expired'))

        # Formatage dates après calcul is_expired
        for r in rows:
            r['date'] = _fmt_date_rfq(r['date'])
            r['expire'] = _fmt_date_rfq(r.get('expire_raw') or '')

        # ── Exports ────────────────────────────────────────────
        if export == 'csv':
            out = io.StringIO()
            w = csv.writer(out, delimiter=';')
            w.writerow(['N° Demande', 'Date', 'Fournisseur', 'Responsable', 'Montant HT', 'Devise', 'État'])
            for r in rows:
                w.writerow([r['name'], r['date'], r['fournisseur'], r['responsable'],
                             r['montant'], r['devise'], r['etat']])
            resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8-sig')
            resp['Content-Disposition'] = 'attachment; filename=demandes_prix.csv'
            return resp

        if export == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Demandes de prix'
            header = ['N° Demande', 'Date', 'Fournisseur', 'Responsable', 'Montant HT', 'Devise', 'État']
            ws.append(header)
            hdr_fill = PatternFill('solid', fgColor='1A2C4E')
            hdr_font = Font(color='FFFFFF', bold=True)
            for cell in ws[1]:
                cell.fill = hdr_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for r in rows:
                ws.append([r['name'], r['date'], r['fournisseur'], r['responsable'],
                            r['montant'], r['devise'], r['etat']])
            ws.append(['TOTAL', '', '', '', total_montant, '', ''])

            last_row = ws.max_row
            total_fill = PatternFill('solid', fgColor='1A2C4E')
            total_font = Font(color='FFFFFF', bold=True)
            for cell in ws[last_row]:
                cell.fill = total_fill
                cell.font = total_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 14
            ws.column_dimensions['C'].width = 30
            ws.column_dimensions['D'].width = 24
            ws.column_dimensions['E'].width = 18
            ws.column_dimensions['F'].width = 14
            ws.column_dimensions['G'].width = 14

            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = f'A1:G{last_row}'

            for row_idx in range(2, last_row):
                ws[f'E{row_idx}'].number_format = '#,##0.00'
                ws[f'E{row_idx}'].alignment = Alignment(horizontal='right')
                ws[f'B{row_idx}'].alignment = Alignment(horizontal='center')

            ws[f'E{last_row}'].number_format = '#,##0.00'
            ws[f'E{last_row}'].alignment = Alignment(horizontal='right')

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            resp = HttpResponse(
                bio.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            resp['Content-Disposition'] = 'attachment; filename=demandes_prix.xlsx'
            return resp

        if export == 'pdf':
            bio_pdf = io.BytesIO()
            logo_path = settings.BASE_DIR / 'static' / 'images' / 'logo_somatrin.png'
            pw, ph = landscape(A4)

            class _PagedCanvasDP(rl_canvas.Canvas):
                def __init__(self, *args, **kw):
                    super().__init__(*args, **kw)
                    self._saved_states = []
                def showPage(self):
                    self._saved_states.append(dict(self.__dict__))
                    self._startPage()
                def save(self):
                    total = len(self._saved_states)
                    for state in self._saved_states:
                        self.__dict__.update(state)
                        self.setFont('Helvetica', 7.5)
                        self.setFillColor(colors.HexColor('#6B7280'))
                        self.drawRightString(pw - 15 * mm, 10 * mm,
                                             f'Page {self._pageNumber} / {total}')
                        super().showPage()
                    super().save()

            def _header_dp(c, d):
                c.saveState()
                logo_x = 15 * mm
                logo_y = ph - 22 * mm
                logo_w = 32 * mm
                logo_h = 11 * mm
                if logo_path.is_file():
                    c.drawImage(str(logo_path), logo_x, logo_y,
                                width=logo_w, height=logo_h,
                                preserveAspectRatio=True, mask='auto')
                header_y = logo_y + (logo_h / 2.0) - (3 * mm)
                c.setFont('Helvetica-Bold', 11)
                c.setFillColor(colors.HexColor('#1A2C4E'))
                c.drawString(logo_x + logo_w + (6 * mm), header_y, 'Demandes de prix — SOMATRIN')
                c.setFont('Helvetica', 8)
                c.setFillColor(colors.HexColor('#6B7280'))
                c.drawRightString(pw - 15 * mm, header_y, 'Document Confidentiel')
                c.setStrokeColor(colors.HexColor('#E87722'))
                c.setLineWidth(1.2)
                c.line(15 * mm, ph - 26 * mm, pw - 15 * mm, ph - 26 * mm)
                c.restoreState()

            navy_dp = colors.HexColor('#1A2C4E')
            lgray_dp = colors.HexColor('#F8FAFC')

            pdf_data = [['N° Demande', 'Date', 'Fournisseur',
                          'Responsable', 'Montant HT', 'Devise', 'État']]
            for r in rows:
                pdf_data.append([
                    r['name'], r['date'], r['fournisseur'],
                    r['responsable'], f"{r['montant']:,.2f}", r['devise'], r['etat'],
                ])
            pdf_data.append(['TOTAL', '', '', '',
                              f"{total_montant:,.2f}", '', f"{total_dp} demande(s)"])

            cw_dp = [38 * mm, 24 * mm, 55 * mm, 45 * mm, 32 * mm, 18 * mm, 22 * mm]
            tbl_dp = Table(pdf_data, colWidths=cw_dp, repeatRows=1)
            tbl_dp.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0),  (-1, 0),  navy_dp),
                ('TEXTCOLOR',     (0, 0),  (-1, 0),  colors.white),
                ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0, 0),  (-1, 0),  8),
                ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
                ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
                ('FONTNAME',      (0, 1),  (-1, -2), 'Helvetica'),
                ('FONTSIZE',      (0, 1),  (-1, -2), 7.5),
                ('ROWBACKGROUNDS',(0, 1),  (-1, -2), [colors.white, lgray_dp]),
                ('ALIGN',         (4, 1),  (4, -2),  'RIGHT'),
                ('GRID',          (0, 0),  (-1, -1), 0.3, colors.HexColor('#E5E7EB')),
                ('BACKGROUND',    (0, -1), (-1, -1), navy_dp),
                ('TEXTCOLOR',     (0, -1), (-1, -1), colors.white),
                ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE',      (0, -1), (-1, -1), 8),
                ('ALIGN',         (4, -1), (4, -1),  'RIGHT'),
            ]))

            doc_dp = SimpleDocTemplate(
                bio_pdf, pagesize=landscape(A4),
                leftMargin=15 * mm, rightMargin=15 * mm,
                topMargin=32 * mm, bottomMargin=22 * mm,
            )
            doc_dp.build([tbl_dp], onFirstPage=_header_dp, onLaterPages=_header_dp,
                         canvasmaker=_PagedCanvasDP)
            bio_pdf.seek(0)
            resp = HttpResponse(bio_pdf.getvalue(), content_type='application/pdf')
            resp['Content-Disposition'] = "attachment; filename=demandes_prix.pdf"
            return resp

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'achats/demandes_prix.html', {
        'rows':         rows,
        'error':        error,
        'date_debut':   date_debut,
        'date_fin':     date_fin,
        'fournisseur':  fournisseur,
        'etat':         etat,
        'responsable':  responsable,
        'fournisseurs': fournisseurs,
        'responsables': responsables,
        'total_dp':     total_dp,
        'total_montant': total_montant,
        'nb_attente':   nb_attente,
        'nb_expirees':  nb_expirees,
    })


@login_required
def achats_purchase_orders(request):
    date_debut  = request.GET.get('date_debut',   '').strip()
    date_fin    = request.GET.get('date_fin',     '').strip()
    fournisseur = request.GET.get('fournisseur',  '').strip()
    etat        = request.GET.get('etat',         '').strip()
    responsable = request.GET.get('responsable',  '').strip()
    societe     = request.GET.get('societe',      '').strip()
    export      = request.GET.get('export',       '').strip().lower()

    STATE_LABELS_PO = {
        'draft':    'Brouillon',
        'sent':     'Envoyé',
        'purchase': 'Confirmé',
        'done':     'Terminé',
        'cancel':   'Annulé',
    }

    rows        = []
    error       = None
    fournisseurs = []
    responsables = []
    societes    = []
    total_bons  = 0
    total_ht    = 0.0
    nb_confirmes = 0
    nb_attente  = 0
    nb_retard = 0
    total_ttc_sum = 0.0
    taux_confirmation_pct = 0.0
    soma_dashboard_ctx_payload = {'page': 'achats_bons_commande', 'kpis': {}}

    def _fmt_date_po(d):
        if not d or len(d) < 10:
            return '—'
        return f'{d[8:10]}/{d[5:7]}/{d[0:4]}'

    try:
        uid, models = get_odoo_connection()

        po_exists = False
        try:
            models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order', 'search_count', [[]], {},
            )
            po_exists = True
        except Exception:
            po_exists = False

        if not po_exists:
            from reporting.services import soma_ai as _soma_kpi_store
            request.session.pop(_soma_kpi_store.SOMA_SESSION_KPIS_ACHATS_BC, None)
            request.session.modified = True
            return render(request, 'achats/bons_commande.html', {
                'rows': [], 'error': "Modèle purchase.order non disponible dans cette instance Odoo",
                'date_debut': date_debut, 'date_fin': date_fin,
                'fournisseur': fournisseur, 'etat': etat,
                'responsable': responsable, 'societe': societe,
                'fournisseurs': [], 'responsables': [], 'societes': [],
                'total_bons': 0, 'total_ht': 0.0, 'nb_confirmes': 0, 'nb_attente': 0,
                'nb_retard': 0, 'total_ttc_sum': 0.0, 'taux_confirmation_pct': 0.0,
                'soma_dashboard_ctx': soma_dashboard_ctx_payload,
            })

        domain = []
        if date_debut:
            domain.append(('date_order', '>=', date_debut))
        if date_fin:
            domain.append(('date_order', '<=', date_fin + ' 23:59:59'))
        if etat:
            domain.append(('state', '=', etat))

        # Bons « en retard » : ids calculés par Odoo (search), même périmètre domaine que la liste.
        today_po = date.today()
        late_domain = list(domain) + [
            ('state', 'not in', ('done', 'cancel')),
            ('date_planned', '!=', False),
            ('date_planned', '<', today_po.isoformat()),
        ]
        try:
            late_ids_raw = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order', 'search',
                [late_domain],
                {},
            )
            late_ids = {int(x) for x in (late_ids_raw or [])}
        except Exception:
            late_ids = set()

        records = odoo_search_read_all(
            uid, models, 'purchase.order', domain,
            ['id', 'name', 'date_order', 'partner_id', 'user_id',
             'amount_untaxed', 'amount_tax', 'amount_total',
             'state', 'currency_id', 'date_planned', 'company_id'],
            order='date_order desc',
        )

        all_rows = []
        for r in records:
            partner      = r.get('partner_id')
            fourn_name   = partner[1]  if isinstance(partner,  list) and len(partner)  > 1 else '—'
            user         = r.get('user_id')
            user_name    = user[1]     if isinstance(user,     list) and len(user)     > 1 else '—'
            currency     = r.get('currency_id')
            devise       = currency[1] if isinstance(currency, list) and len(currency) > 1 else 'MAD'
            company      = r.get('company_id')
            company_name = company[1]  if isinstance(company,  list) and len(company)  > 1 else '—'
            all_rows.append({
                'id':           r.get('id'),
                'name':         r.get('name') or '—',
                'date':         (r.get('date_order') or '')[:10],
                'date_planned': (r.get('date_planned') or '')[:10],
                'fournisseur':  fourn_name,
                'responsable':  user_name,
                'societe':      company_name,
                'montant_ht':   float(r.get('amount_untaxed') or 0),
                'montant_tax':  float(r.get('amount_tax')      or 0),
                'montant_ttc':  float(r.get('amount_total')    or 0),
                'devise':       devise,
                'etat_raw':     r.get('state') or 'draft',
                'etat':         STATE_LABELS_PO.get(r.get('state') or 'draft', '—'),
            })

        fournisseurs = sorted({row['fournisseur'] for row in all_rows if row['fournisseur'] != '—'})
        responsables = sorted({row['responsable'] for row in all_rows if row['responsable'] != '—'})
        societes     = sorted({row['societe']     for row in all_rows if row['societe']     != '—'})

        rows = all_rows
        if fournisseur:
            rows = [r for r in rows if fournisseur.lower() in r['fournisseur'].lower()]
        if responsable:
            rows = [r for r in rows if responsable.lower() in r['responsable'].lower()]
        if societe:
            rows = [r for r in rows if societe == r['societe']]

        total_bons   = len(rows)
        total_ht     = round(sum(r['montant_ht']  for r in rows), 2)
        total_ttc_sum = round(sum(r['montant_ttc'] for r in rows), 2)
        nb_confirmes = sum(1 for r in rows if r['etat_raw'] in ('purchase', 'done'))
        nb_attente   = sum(1 for r in rows if r['etat_raw'] in ('draft', 'sent'))
        for r in rows:
            rid = r.get('id')
            try:
                r['en_retard_bc'] = bool(rid is not None and int(rid) in late_ids)
            except (TypeError, ValueError):
                r['en_retard_bc'] = False
        nb_retard = sum(1 for r in rows if r.get('en_retard_bc'))
        taux_confirmation_pct = round(100.0 * nb_confirmes / total_bons, 1) if total_bons else 0.0

        soma_dashboard_ctx_payload = {
            'page': 'achats_bons_commande',
            'kpis': {
                'total_bons': int(total_bons),
                'montant_ttc_total': float(total_ttc_sum),
                'nb_retard': int(nb_retard),
                'bons_en_retard': int(nb_retard),
                'nb_confirmes': int(nb_confirmes),
                'taux_confirmation_pct': float(taux_confirmation_pct),
            },
        }
        if date_debut:
            soma_dashboard_ctx_payload['kpis']['periode_debut'] = str(date_debut)[:32]
        if date_fin:
            soma_dashboard_ctx_payload['kpis']['periode_fin'] = str(date_fin)[:32]

        from reporting.services import soma_ai as _soma_kpi_store
        request.session[_soma_kpi_store.SOMA_SESSION_KPIS_ACHATS_BC] = {
            'page': soma_dashboard_ctx_payload['page'],
            'kpis': dict(soma_dashboard_ctx_payload['kpis']),
        }
        request.session.modified = True

        for r in rows:
            r['date']         = _fmt_date_po(r['date'])
            r['date_planned'] = _fmt_date_po(r['date_planned'])

        if export == 'csv':
            out = io.StringIO()
            w = csv.writer(out, delimiter=';')
            w.writerow(['N° Commande', 'Date', 'Date prévue', 'Fournisseur', 'Responsable',
                        'Société', 'Montant HT', 'Taxes', 'Montant TTC', 'Devise', 'État'])
            for r in rows:
                w.writerow([r['name'], r['date'], r['date_planned'], r['fournisseur'],
                             r['responsable'], r['societe'], r['montant_ht'],
                            r['montant_tax'], r['montant_ttc'], r['devise'], r['etat']])
            resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8-sig')
            resp['Content-Disposition'] = 'attachment; filename=bons_commande.csv'
            return resp

        if export == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Bons de commande'
            header = ['N° Commande', 'Date', 'Date prévue', 'Fournisseur', 'Responsable',
                      'Société', 'Montant HT', 'Taxes', 'Montant TTC', 'Devise', 'État']
            ws.append(header)
            hdr_fill = PatternFill('solid', fgColor='1A2C4E')
            hdr_font = Font(color='FFFFFF', bold=True)
            for cell in ws[1]:
                cell.fill = hdr_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for r in rows:
                ws.append([r['name'], r['date'], r['date_planned'], r['fournisseur'],
                            r['responsable'], r['societe'], r['montant_ht'],
                            r['montant_tax'], r['montant_ttc'], r['devise'], r['etat']])
            ws.append(['TOTAL', '', '', '', '', '', total_ht, '', '', '', ''])

            last_row = ws.max_row
            for cell in ws[last_row]:
                cell.fill = PatternFill('solid', fgColor='1A2C4E')
                cell.font = Font(color='FFFFFF', bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            col_widths = {'A': 18, 'B': 14, 'C': 14, 'D': 30, 'E': 24,
                          'F': 20, 'G': 16, 'H': 16, 'I': 16, 'J': 10, 'K': 14}
            for col, w_val in col_widths.items():
                ws.column_dimensions[col].width = w_val

            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = f'A1:K{last_row}'

            for row_idx in range(2, last_row):
                for col in ('G', 'H', 'I'):
                    ws[f'{col}{row_idx}'].number_format = '#,##0.00'
                    ws[f'{col}{row_idx}'].alignment = Alignment(horizontal='right')
                ws[f'B{row_idx}'].alignment = Alignment(horizontal='center')
                ws[f'C{row_idx}'].alignment = Alignment(horizontal='center')
            for col in ('G', 'H', 'I'):
                ws[f'{col}{last_row}'].number_format = '#,##0.00'
                ws[f'{col}{last_row}'].alignment = Alignment(horizontal='right')

            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            resp = HttpResponse(
                bio.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            resp['Content-Disposition'] = 'attachment; filename=bons_commande.xlsx'
            return resp

        if export == 'pdf':
            bio_pdf = io.BytesIO()
            logo_path = settings.BASE_DIR / 'static' / 'images' / 'logo_somatrin.png'
            pw, ph = landscape(A4)

            class _PagedCanvasPO(rl_canvas.Canvas):
                def __init__(self, *args, **kw):
                    super().__init__(*args, **kw)
                    self._saved_states = []
                def showPage(self):
                    self._saved_states.append(dict(self.__dict__))
                    self._startPage()
                def save(self):
                    total = len(self._saved_states)
                    for state in self._saved_states:
                        self.__dict__.update(state)
                        self.setFont('Helvetica', 7.5)
                        self.setFillColor(colors.HexColor('#6B7280'))
                        self.drawRightString(pw - 15 * mm, 10 * mm, f'Page {self._pageNumber} / {total}')
                        super().showPage()
                    super().save()

            def _header_po(c, d):
                c.saveState()
                logo_x = 15 * mm
                logo_y = ph - 22 * mm
                logo_w = 32 * mm
                logo_h = 11 * mm
                if logo_path.is_file():
                    c.drawImage(str(logo_path), logo_x, logo_y,
                                width=logo_w, height=logo_h,
                                preserveAspectRatio=True, mask='auto')
                header_y = logo_y + (logo_h / 2.0) - (3 * mm)
                c.setFont('Helvetica-Bold', 11)
                c.setFillColor(colors.HexColor('#1A2C4E'))
                c.drawString(logo_x + logo_w + (6 * mm), header_y, 'Bons de commande — SOMATRIN')
                c.setFont('Helvetica', 8)
                c.setFillColor(colors.HexColor('#6B7280'))
                c.drawRightString(pw - 15 * mm, header_y, 'Document Confidentiel')
                c.setStrokeColor(colors.HexColor('#E87722'))
                c.setLineWidth(1.2)
                c.line(15 * mm, ph - 26 * mm, pw - 15 * mm, ph - 26 * mm)
                c.restoreState()

            navy_po = colors.HexColor('#1A2C4E')
            lgray_po = colors.HexColor('#F8FAFC')
            pdf_data = [['N° Commande', 'Date', 'Date prévue', 'Fournisseur', 'Responsable',
                         'Société', 'Montant HT', 'Taxes', 'Montant TTC', 'État']]
            for r in rows:
                pdf_data.append([
                    r['name'], r['date'], r['date_planned'], r['fournisseur'], r['responsable'],
                    r['societe'], f"{r['montant_ht']:,.2f}", f"{r['montant_tax']:,.2f}",
                    f"{r['montant_ttc']:,.2f}", r['etat'],
                ])
            pdf_data.append(['TOTAL', '', '', '', '', '', f"{total_ht:,.2f}", '', '', f"{total_bons} bon(s)"])

            cw_po = [33 * mm, 20 * mm, 24 * mm, 42 * mm, 32 * mm, 30 * mm, 24 * mm, 20 * mm, 24 * mm, 20 * mm]
            tbl_po = Table(pdf_data, colWidths=cw_po, repeatRows=1)
            tbl_po.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0),  (-1, 0),  navy_po),
                ('TEXTCOLOR',     (0, 0),  (-1, 0),  colors.white),
                ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0, 0),  (-1, 0),  8),
                ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
                ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
                ('FONTNAME',      (0, 1),  (-1, -2), 'Helvetica'),
                ('FONTSIZE',      (0, 1),  (-1, -2), 7.2),
                ('ROWBACKGROUNDS',(0, 1),  (-1, -2), [colors.white, lgray_po]),
                ('ALIGN',         (6, 1),  (8, -1),  'RIGHT'),
                ('GRID',          (0, 0),  (-1, -1), 0.3, colors.HexColor('#E5E7EB')),
                ('BACKGROUND',    (0, -1), (-1, -1), navy_po),
                ('TEXTCOLOR',     (0, -1), (-1, -1), colors.white),
                ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE',      (0, -1), (-1, -1), 8),
            ]))

            doc_po = SimpleDocTemplate(
                bio_pdf, pagesize=landscape(A4),
                leftMargin=15 * mm, rightMargin=15 * mm,
                topMargin=32 * mm, bottomMargin=22 * mm,
            )
            doc_po.build([tbl_po], onFirstPage=_header_po, onLaterPages=_header_po, canvasmaker=_PagedCanvasPO)
            bio_pdf.seek(0)
            resp = HttpResponse(bio_pdf.getvalue(), content_type='application/pdf')
            resp['Content-Disposition'] = "attachment; filename=bons_commande.pdf"
            return resp

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'achats/bons_commande.html', {
        'rows':         rows,
        'error':        error,
        'date_debut':   date_debut,
        'date_fin':     date_fin,
        'fournisseur':  fournisseur,
        'etat':         etat,
        'responsable':  responsable,
        'societe':      societe,
        'fournisseurs': fournisseurs,
        'responsables': responsables,
        'societes':     societes,
        'total_bons':   total_bons,
        'total_ht':     total_ht,
        'nb_confirmes': nb_confirmes,
        'nb_attente':   nb_attente,
        'nb_retard':    nb_retard,
        'total_ttc_sum': total_ttc_sum,
        'taux_confirmation_pct': taux_confirmation_pct,
        'soma_dashboard_ctx': soma_dashboard_ctx_payload,
    })


@login_required
def achats_delivery_tracking(request):
    date_debut = request.GET.get('date_debut', '').strip()
    date_fin = request.GET.get('date_fin', '').strip()
    fournisseur = request.GET.get('fournisseur', '').strip()
    statut = request.GET.get('statut', '').strip()
    responsable = request.GET.get('responsable', '').strip()
    export = request.GET.get('export', '').strip().lower()

    rows = []
    error = None
    fournisseurs = []
    responsables = []
    total_livraisons = 0
    total_ttc = 0.0
    nb_recues = 0
    nb_retard = 0
    nb_en_cours = 0

    def _fmt_date(v):
        if not v or len(v) < 10:
            return '—'
        return f'{v[8:10]}/{v[5:7]}/{v[0:4]}'

    try:
        from datetime import datetime as _dt

        uid, models = get_odoo_connection()
        po_exists = False
        try:
            models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order', 'search_count', [[]], {},
            )
            po_exists = True
        except Exception:
            po_exists = False

        if not po_exists:
            return render(request, 'achats/suivi_livraisons.html', {
                'rows': [],
                'error': "Modèle purchase.order non disponible dans cette instance Odoo",
                'date_debut': date_debut,
                'date_fin': date_fin,
                'fournisseur': fournisseur,
                'statut': statut,
                'responsable': responsable,
                'fournisseurs': [],
                'responsables': [],
                'total_livraisons': 0,
                'total_ttc': 0.0,
                'nb_recues': 0,
                'nb_retard': 0,
                'nb_en_cours': 0,
                'stats_statut': {'en_retard': 0, 'en_cours': 0, 'recue': 0},
            })

        domain = []
        if date_debut:
            domain.append(('date_order', '>=', date_debut))
        if date_fin:
            domain.append(('date_order', '<=', date_fin + ' 23:59:59'))

        records = odoo_search_read_all(
            uid, models, 'purchase.order', domain,
            ['name', 'date_order', 'date_planned', 'partner_id', 'user_id', 'amount_total', 'state', 'currency_id'],
            order='date_order desc',
        )

        today = date.today()
        for r in records:
            partner = r.get('partner_id')
            fourn_name = partner[1] if isinstance(partner, list) and len(partner) > 1 else '—'
            user = r.get('user_id')
            user_name = user[1] if isinstance(user, list) and len(user) > 1 else '—'
            cur = r.get('currency_id')
            devise = cur[1] if isinstance(cur, list) and len(cur) > 1 else 'MAD'
            state_raw = r.get('state') or 'draft'

            raw_order = (r.get('date_order') or '')[:10]
            raw_planned = (r.get('date_planned') or '')[:10]
            delay_days = None
            if raw_order and raw_planned and len(raw_order) == 10 and len(raw_planned) == 10:
                try:
                    d_order = _dt.strptime(raw_order, '%Y-%m-%d').date()
                    d_plan = _dt.strptime(raw_planned, '%Y-%m-%d').date()
                    delay_days = (d_plan - d_order).days
                except Exception:
                    delay_days = None

            is_received = state_raw in ('done',)
            is_cancel = state_raw in ('cancel',)
            is_late = False
            if raw_planned and len(raw_planned) == 10 and not is_received and not is_cancel:
                try:
                    d_plan = _dt.strptime(raw_planned, '%Y-%m-%d').date()
                    is_late = d_plan < today
                except Exception:
                    is_late = False

            if is_cancel:
                statut_livraison = 'Annulée'
            elif is_received:
                statut_livraison = 'Reçue'
            elif is_late:
                statut_livraison = 'En retard'
            else:
                statut_livraison = 'En cours'

            row = {
                'name': r.get('name') or '—',
                'date_order': _fmt_date(raw_order),
                'date_planned': _fmt_date(raw_planned),
                'fournisseur': fourn_name,
                'responsable': user_name,
                'montant_ttc': float(r.get('amount_total') or 0),
                'devise': devise,
                'state_raw': state_raw,
                'statut_livraison': statut_livraison,
                'delay_days': delay_days,
            }
            rows.append(row)

        fournisseurs = sorted({r['fournisseur'] for r in rows if r['fournisseur'] != '—'})
        responsables = sorted({r['responsable'] for r in rows if r['responsable'] != '—'})

        if fournisseur:
            rows = [r for r in rows if fournisseur.lower() in r['fournisseur'].lower()]
        if responsable:
            rows = [r for r in rows if responsable.lower() in r['responsable'].lower()]
        if statut:
            rows = [r for r in rows if r['statut_livraison'] == statut]

        total_livraisons = len(rows)
        total_ttc = round(sum(r['montant_ttc'] for r in rows), 2)
        nb_recues = sum(1 for r in rows if r['statut_livraison'] == 'Reçue')
        nb_retard = sum(1 for r in rows if r['statut_livraison'] == 'En retard')
        nb_en_cours = sum(1 for r in rows if r['statut_livraison'] == 'En cours')

        if export == 'csv':
            out = io.StringIO()
            w = csv.writer(out, delimiter=';')
            w.writerow(['N° Commande', 'Date commande', 'Date prévue', 'Fournisseur', 'Responsable',
                        'Montant TTC', 'Devise', 'Statut livraison', 'Délai (j)'])
            for r in rows:
                w.writerow([r['name'], r['date_order'], r['date_planned'], r['fournisseur'], r['responsable'],
                            r['montant_ttc'], r['devise'], r['statut_livraison'], r['delay_days'] if r['delay_days'] is not None else '—'])
            resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8-sig')
            resp['Content-Disposition'] = 'attachment; filename=suivi_livraisons.csv'
            return resp

        if export == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Suivi livraisons'
            ws.append(['N° Commande', 'Date commande', 'Date prévue', 'Fournisseur', 'Responsable',
                       'Montant TTC', 'Devise', 'Statut livraison', 'Délai (j)'])
            hdr_fill = PatternFill('solid', fgColor='1A2C4E')
            hdr_font = Font(color='FFFFFF', bold=True)
            for cell in ws[1]:
                cell.fill = hdr_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            for r in rows:
                ws.append([r['name'], r['date_order'], r['date_planned'], r['fournisseur'], r['responsable'],
                          r['montant_ttc'], r['devise'], r['statut_livraison'], r['delay_days'] if r['delay_days'] is not None else ''])
            ws.append(['TOTAL', '', '', '', '', total_ttc, '', '', ''])
            last_row = ws.max_row
            for cell in ws[last_row]:
                cell.fill = PatternFill('solid', fgColor='1A2C4E')
                cell.font = Font(color='FFFFFF', bold=True)
            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 14
            ws.column_dimensions['C'].width = 14
            ws.column_dimensions['D'].width = 28
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 16
            ws.column_dimensions['G'].width = 10
            ws.column_dimensions['H'].width = 18
            ws.column_dimensions['I'].width = 11
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            resp = HttpResponse(
                bio.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            resp['Content-Disposition'] = 'attachment; filename=suivi_livraisons.xlsx'
            return resp

        if export == 'pdf':
            bio_pdf = io.BytesIO()
            logo_path = settings.BASE_DIR / 'static' / 'images' / 'logo_somatrin.png'
            pw, ph = landscape(A4)

            class _PagedCanvasSL(rl_canvas.Canvas):
                def __init__(self, *args, **kw):
                    super().__init__(*args, **kw)
                    self._saved_states = []
                def showPage(self):
                    self._saved_states.append(dict(self.__dict__))
                    self._startPage()
                def save(self):
                    total = len(self._saved_states)
                    for state in self._saved_states:
                        self.__dict__.update(state)
                        self.setFont('Helvetica', 7.5)
                        self.setFillColor(colors.HexColor('#6B7280'))
                        self.drawRightString(pw - 15 * mm, 10 * mm,
                                             f'Page {self._pageNumber} / {total}')
                        super().showPage()
                    super().save()

            from datetime import datetime as _dtpdf
            gen_date = _dtpdf.now().strftime('%d/%m/%Y %H:%M')

            def _header_sl(c, d):
                c.saveState()
                if logo_path.is_file():
                    c.drawImage(str(logo_path), 15 * mm, ph - 22 * mm,
                                width=32 * mm, height=11 * mm,
                                preserveAspectRatio=True, mask='auto')
                c.setFont('Helvetica-Bold', 11)
                c.setFillColor(colors.HexColor('#1A2C4E'))
                c.drawString(52 * mm, ph - 13 * mm, 'SUIVI LIVRAISONS — SOMATRIN')
                c.setFont('Helvetica', 8)
                c.setFillColor(colors.HexColor('#6B7280'))
                c.drawString(52 * mm, ph - 20 * mm,
                             'Contrôle des délais et statut des réceptions fournisseurs')
                c.drawRightString(pw - 15 * mm, ph - 13 * mm, f'Généré le {gen_date}')
                c.drawRightString(pw - 15 * mm, ph - 20 * mm, 'Document Confidentiel')
                c.setStrokeColor(colors.HexColor('#E87722'))
                c.setLineWidth(1.2)
                c.line(15 * mm, ph - 26 * mm, pw - 15 * mm, ph - 26 * mm)
                c.restoreState()

            navy_sl = colors.HexColor('#1A2C4E')
            lgray_sl = colors.HexColor('#F8FAFC')

            pdf_data = [['N° Commande', 'Date cmd', 'Date prévue', 'Fournisseur',
                          'Responsable', 'Montant TTC', 'Devise', 'Statut', 'Délai (j)']]
            for r in rows:
                pdf_data.append([
                    r['name'], r['date_order'], r['date_planned'],
                    r['fournisseur'], r['responsable'],
                    f"{r['montant_ttc']:,.2f}", r['devise'],
                    r['statut_livraison'],
                    str(r['delay_days']) if r['delay_days'] is not None else '—',
                ])
            pdf_data.append(['TOTAL', '', '', '', '',
                              f"{total_ttc:,.2f}", '', '', ''])

            cw_sl = [30 * mm, 20 * mm, 20 * mm, 45 * mm, 35 * mm,
                     28 * mm, 14 * mm, 22 * mm, 18 * mm]
            tbl_sl = Table(pdf_data, colWidths=cw_sl, repeatRows=1)
            tbl_sl.setStyle(TableStyle([
                ('BACKGROUND',    (0, 0),  (-1, 0),  navy_sl),
                ('TEXTCOLOR',     (0, 0),  (-1, 0),  colors.white),
                ('FONTNAME',      (0, 0),  (-1, 0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0, 0),  (-1, 0),  8),
                ('ALIGN',         (0, 0),  (-1, 0),  'CENTER'),
                ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
                ('FONTNAME',      (0, 1),  (-1, -2), 'Helvetica'),
                ('FONTSIZE',      (0, 1),  (-1, -2), 7.5),
                ('ROWBACKGROUNDS',(0, 1),  (-1, -2), [colors.white, lgray_sl]),
                ('ALIGN',         (5, 1),  (5, -2),  'RIGHT'),
                ('ALIGN',         (8, 1),  (8, -2),  'CENTER'),
                ('GRID',          (0, 0),  (-1, -1), 0.3, colors.HexColor('#E5E7EB')),
                ('BACKGROUND',    (0, -1), (-1, -1), navy_sl),
                ('TEXTCOLOR',     (0, -1), (-1, -1), colors.white),
                ('FONTNAME',      (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE',      (0, -1), (-1, -1), 8),
                ('ALIGN',         (5, -1), (5, -1),  'RIGHT'),
            ]))

            doc_sl = SimpleDocTemplate(
                bio_pdf, pagesize=landscape(A4),
                leftMargin=15 * mm, rightMargin=15 * mm,
                topMargin=32 * mm, bottomMargin=22 * mm,
            )
            doc_sl.build([tbl_sl], onFirstPage=_header_sl, onLaterPages=_header_sl,
                         canvasmaker=_PagedCanvasSL)
            bio_pdf.seek(0)
            resp = HttpResponse(bio_pdf.getvalue(), content_type='application/pdf')
            resp['Content-Disposition'] = "attachment; filename=suivi_livraisons.pdf"
            return resp

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    stats_statut = {
        'en_retard': nb_retard,
        'en_cours':  nb_en_cours,
        'recue':     nb_recues,
    }

    return render(request, 'achats/suivi_livraisons.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'fournisseur': fournisseur,
        'statut': statut,
        'responsable': responsable,
        'fournisseurs': fournisseurs,
        'responsables': responsables,
        'total_livraisons': total_livraisons,
        'total_ttc': total_ttc,
        'nb_recues': nb_recues,
        'nb_retard': nb_retard,
        'nb_en_cours': nb_en_cours,
        'stats_statut': stats_statut,
    })


@login_required
def achats_suppliers(request):
    export = request.GET.get('export', '')
    nom_filter = request.GET.get('nom', '').strip()
    ville_filter = request.GET.get('ville', '').strip()
    pays_filter = request.GET.get('pays', '').strip()
    statut_filter = request.GET.get('statut', '').strip()

    _error_ctx = {
        'rows': [], 'total_fournisseurs': 0,
        'fournisseurs_actifs': 0, 'total_commandes': 0,
        'top_fournisseur': '—', 'nb_avec_email': 0, 'nb_avec_phone': 0,
        'stats_ville_json': '{}',
        'nom': nom_filter, 'ville': ville_filter, 'pays': pays_filter, 'statut': statut_filter,
        'villes': [], 'pays_list': [],
    }

    try:
        uid, models = get_odoo_connection()
    except Exception as e:
        _error_ctx['error'] = str(e)
        return render(request, 'achats/fournisseurs.html', _error_ctx)

    domain = [('is_company', '=', True), ('supplier_rank', '>', 0)]
    fields = ['id', 'name', 'phone', 'mobile', 'email', 'city', 'country_id',
              'supplier_rank', 'purchase_order_count', 'ref', 'vat', 'street', 'zip']

    try:
        records = odoo_search_read_all(
            uid, models, 'res.partner', domain, fields, order='name asc',
        )
    except Exception as e:
        _error_ctx['error'] = str(e)
        return render(request, 'achats/fournisseurs.html', _error_ctx)

    # ── Dédoublonnage par (nom + ville) ──────────────────────────
    _seen = set()
    records_uniq = []
    for r in records:
        key = (
            (r.get('name') or '').strip().upper(),
            (r.get('city') or '').strip().upper(),
        )
        if key not in _seen:
            _seen.add(key)
            records_uniq.append(r)
    records = records_uniq

    # ── Traduction pays ──────────────────────────────────────────
    _PAYS_FR = {
        'Morocco': 'Maroc', 'France': 'France',
        'Algeria': 'Algérie', 'Tunisia': 'Tunisie',
        'Spain': 'Espagne', 'Belgium': 'Belgique',
        'Germany': 'Allemagne', 'Italy': 'Italie',
        'United Arab Emirates': 'Émirats Arabes Unis',
        'Saudi Arabia': 'Arabie Saoudite',
        'United States': 'États-Unis', 'China': 'Chine',
        'United Kingdom': 'Royaume-Uni', 'Netherlands': 'Pays-Bas',
        'Switzerland': 'Suisse', 'Portugal': 'Portugal',
        'Turkey': 'Turquie', 'Egypt': 'Égypte',
    }

    # ── Normalisation complète avant tout filtrage ───────────────
    def _clean(val):
        if not val or val is False:
            return ''
        return str(val).strip()

    all_rows = []
    for r in records:
        nom   = _clean(r.get('name')) or '—'
        ville = (_clean(r.get('city')) or '—').upper()
        pays_raw = r.get('country_id')
        if isinstance(pays_raw, (list, tuple)) and len(pays_raw) > 1:
            pays = _PAYS_FR.get(_clean(pays_raw[1]), _clean(pays_raw[1])) or '—'
        else:
            pays = '—'
        phone = _clean(r.get('phone')) or _clean(r.get('mobile')) or '—'
        email = _clean(r.get('email')) or '—'
        ref   = _clean(r.get('ref'))   or '—'
        vat   = _clean(r.get('vat'))   or '—'
        nb_cmd        = r.get('purchase_order_count') or 0
        supplier_rank = r.get('supplier_rank') or 0
        statut = 'Actif' if supplier_rank > 0 else 'Inactif'
        all_rows.append({
            'nom': nom, 'ville': ville, 'pays': pays,
            'phone': phone, 'email': email, 'ref': ref, 'vat': vat,
            'nb_commandes': nb_cmd, 'statut': statut,
        })

    # Totaux globaux et listes de choix calculés sur l'ensemble
    total_commandes = sum(r['nb_commandes'] for r in all_rows)
    villes_all = sorted({r['ville'].upper() for r in all_rows if r['ville'] != '—'})
    pays_all = sorted({r['pays'] for r in all_rows if r['pays'] != '—'})

    top_fournisseur = '—'
    if all_rows:
        top = max(all_rows, key=lambda r: r['nb_commandes'])
        top_fournisseur = top['nom']

    # ── Application des filtres ──────────────────────────────────
    rows = []
    for r in all_rows:
        if nom_filter and nom_filter.lower() not in r['nom'].lower():
            continue
        if ville_filter and ville_filter.upper() != r['ville'].upper():
            continue
        if pays_filter and pays_filter != r['pays']:
            continue
        if statut_filter and statut_filter != r['statut']:
            continue
        rows.append(r)

    total_fournisseurs = len(rows)
    fournisseurs_actifs = sum(1 for r in rows if r['statut'] == 'Actif')
    nb_avec_email = sum(1 for r in rows if r['email'] != '—')
    nb_avec_phone = sum(1 for r in rows if r['phone'] != '—')

    stats_ville: dict = defaultdict(int)
    for r in all_rows:
        if r['ville'] != '—':
            stats_ville[r['ville'].upper()] += 1
    top_villes = sorted(stats_ville.items(), key=lambda x: x[1], reverse=True)[:6]
    stats_ville_json = json.dumps({k: v for k, v in top_villes})

    # ── CSV ─────────────────────────────────────────────────────
    if export == 'csv':
        resp = HttpResponse(content_type='text/csv; charset=utf-8-sig')
        resp['Content-Disposition'] = 'attachment; filename="fournisseurs.csv"'
        w = csv.writer(resp, delimiter=';')
        w.writerow(['Nom', 'Ville', 'Pays', 'Téléphone', 'Email', 'Réf.', 'TVA', 'Nb Commandes', 'Statut'])
        for r in rows:
            w.writerow([r['nom'], r['ville'], r['pays'], r['phone'], r['email'],
                        r['ref'], r['vat'], r['nb_commandes'], r['statut']])
        return resp

    # ── Excel ────────────────────────────────────────────────────
    if export == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Fournisseurs'
        hdr_fill = PatternFill('solid', fgColor='1A2C4E')
        hdr_font = Font(bold=True, color='FFFFFF', size=10)
        hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        headers = ['Nom', 'Ville', 'Pays', 'Téléphone', 'Email', 'Réf.', 'TVA', 'Nb Commandes', 'Statut']
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = hdr_align
        for r in rows:
            ws.append([r['nom'], r['ville'], r['pays'], r['phone'], r['email'],
                       r['ref'], r['vat'], r['nb_commandes'], r['statut']])
        for col in ws.columns:
            max_len = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        resp = HttpResponse(bio.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        resp['Content-Disposition'] = 'attachment; filename="fournisseurs.xlsx"'
        return resp

    # ── PDF ──────────────────────────────────────────────────────
    if export == 'pdf':
        bio_pdf = io.BytesIO()
        logo_path = settings.BASE_DIR / 'static' / 'images' / 'logo_somatrin.png'
        pw, ph = landscape(A4)

        class _PagedCanvasFN(rl_canvas.Canvas):
            def __init__(self, *args, **kw):
                super().__init__(*args, **kw)
                self._saved_states = []
            def showPage(self):
                self._saved_states.append(dict(self.__dict__))
                self._startPage()
            def save(self):
                total = len(self._saved_states)
                for state in self._saved_states:
                    self.__dict__.update(state)
                    self.setFont('Helvetica', 7.5)
                    self.setFillColor(colors.HexColor('#6B7280'))
                    self.drawRightString(pw - 15 * mm, 10 * mm,
                                        f'Page {self._pageNumber} / {total}')
                    super().showPage()
                super().save()

        def _header_fn(c, d):
            c.saveState()
            if logo_path.is_file():
                c.drawImage(str(logo_path), 15 * mm, ph - 22 * mm,
                            width=32 * mm, height=11 * mm,
                            preserveAspectRatio=True, mask='auto')
            c.setFont('Helvetica-Bold', 11)
            c.setFillColor(colors.HexColor('#1A2C4E'))
            c.drawString(52 * mm, ph - 13 * mm, 'Référentiel Fournisseurs — SOMATRIN')
            c.setFont('Helvetica', 8)
            c.setFillColor(colors.HexColor('#6B7280'))
            c.drawRightString(pw - 15 * mm, ph - 13 * mm, 'Document Confidentiel')
            c.setStrokeColor(colors.HexColor('#E87722'))
            c.setLineWidth(1.2)
            c.line(15 * mm, ph - 26 * mm, pw - 15 * mm, ph - 26 * mm)
            c.restoreState()

        styles = getSampleStyleSheet()
        cell_style = ParagraphStyle('cell', parent=styles['Normal'],
                                    fontSize=7.5, leading=10)

        col_widths = [60*mm, 35*mm, 35*mm, 38*mm, 60*mm, 25*mm, 25*mm]
        table_headers = ['Nom', 'Ville', 'Pays', 'Téléphone', 'Email', 'Nb Cmd', 'Statut']

        data = [table_headers]
        for r in rows:
            data.append([
                Paragraph(r['nom'], cell_style),
                r['ville'], r['pays'], r['phone'],
                Paragraph(r['email'], cell_style),
                str(r['nb_commandes']), r['statut'],
            ])

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A2C4E')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('FONTSIZE', (0, 1), (-1, -1), 7.5),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#E5E7EB')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))

        doc = SimpleDocTemplate(
            bio_pdf, pagesize=landscape(A4),
            leftMargin=15*mm, rightMargin=15*mm,
            topMargin=32*mm, bottomMargin=18*mm,
        )
        doc.build([tbl], onFirstPage=_header_fn, onLaterPages=_header_fn,
                  canvasmaker=_PagedCanvasFN)
        bio_pdf.seek(0)
        resp = HttpResponse(bio_pdf.read(), content_type='application/pdf')
        resp['Content-Disposition'] = 'attachment; filename="fournisseurs.pdf"'
        return resp

    return render(request, 'achats/fournisseurs.html', {
        'rows': rows,
        'total_fournisseurs': total_fournisseurs,
        'fournisseurs_actifs': fournisseurs_actifs,
        'total_commandes': total_commandes,
        'top_fournisseur': top_fournisseur,
        'nb_avec_email': nb_avec_email,
        'nb_avec_phone': nb_avec_phone,
        'stats_ville_json': stats_ville_json,
        'nom': nom_filter,
        'ville': ville_filter,
        'pays': pays_filter,
        'statut': statut_filter,
        'villes': villes_all,
        'pays_list': pays_all,
    })


def _render_parc_module(request, current_key, page_title, page_subtitle, rows=None, extra_context=None):
    menu_items = [
        {
            'key': 'overview',
            'label': "Vue d'ensemble",
            'url': 'parc_overview',
            'icon': 'bi-grid-1x2-fill',
            'desc': 'Vue consolidée des indicateurs Parc et Maintenance.',
        },
        {
            'key': 'equipements',
            'label': "Équipements",
            'url': 'parc_equipements',
            'icon': 'bi-truck-front-fill',
            'desc': 'Inventaire des équipements, catégories et équipes.',
        },
        {
            'key': 'disponibilite',
            'label': 'Disponibilité',
            'url': 'parc_disponibilite',
            'icon': 'bi-speedometer2',
            'desc': 'Suivi de disponibilité et taux opérationnel du parc.',
        },
        {
            'key': 'ordres_maintenance',
            'label': 'Ordres & Coûts',
            'url': 'parc_ordres_maintenance',
            'icon': 'bi-tools',
            'desc': 'Ordres de maintenance, analyse des coûts et durées en 3 onglets.',
        },
        {
            'key': 'interventions',
            'label': 'Interventions',
            'url': 'parc_interventions',
            'icon': 'bi-wrench-adjustable',
            'desc': 'Analyse des interventions par type, durée et technicien.',
        },
        {
            'key': 'fournisseurs',
            'label': 'Fournisseurs',
            'url': 'parc_fournisseurs',
            'icon': 'bi-building',
            'desc': "Gestion des fournisseurs d'équipements et de maintenance.",
        },
    ]
    ctx = {
        'page_title': page_title,
        'page_subtitle': page_subtitle,
        'module_key': current_key,
        'menu_items': menu_items,
        'rows': rows or [],
    }
    if extra_context:
        ctx.update(extra_context)
    return render(request, 'parc/module.html', ctx)


@login_required
def parc_overview(request):
    error = None
    kpi_total_equipements = kpi_en_service = kpi_en_maintenance = kpi_ordres_ouverts = 0
    total_maintenance_cost = 0.0
    try:
        uid, models = get_odoo_connection()
        eq_fields = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS, 'maintenance.equipment', 'fields_get', [], {'attributes': ['type']}
        ) or {}
        eq_read = ['name']
        for f in ('active', 'state', 'category_id'):
            if f in eq_fields:
                eq_read.append(f)
        equipments = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read', [[]],
            {'fields': sorted(set(eq_read)), 'limit': 5000},
        )
        kpi_total_equipements = len(equipments)
        cat_counts = defaultdict(int)
        for eq in equipments:
            st = str(eq.get('state') or '').lower()
            if st in ('maintenance', 'repair', 'down'):
                kpi_en_maintenance += 1
            elif bool(eq.get('active', True)):
                kpi_en_service += 1
            else:
                kpi_en_maintenance += 1
            cat = eq.get('category_id')
            cn = (cat[1] if isinstance(cat, list) and len(cat) > 1 else '').upper()
            if 'CAMION' in cn or 'ENGIN' in cn:
                cat_counts['Camion/Engin'] += 1
            elif 'SEMI' in cn or 'REMORQUE' in cn:
                cat_counts['Semi-Remorque'] += 1
            elif 'VOITURE' in cn or 'SERVICE' in cn or ' VL' in cn:
                cat_counts['Voiture Service'] += 1
            elif cn:
                cat_counts[cat[1]] += 1
            else:
                cat_counts['Autre'] += 1

        req_fields = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS, 'maintenance.request', 'fields_get', [], {'attributes': ['type']}
        ) or {}
        chart_months = {}
        chart_top_eq = {}
        kpi_ordres_clos = 0
        kpi_preventive = 0
        if 'stage_id' in req_fields:
            _ov_date = next((f for f in ('request_date', 'create_date') if f in req_fields), None)
            _req_ov_fields = ['stage_id', 'equipment_id']
            for _f in ('duration', 'maintenance_type'):
                if _f in req_fields:
                    _req_ov_fields.append(_f)
            if _ov_date:
                _req_ov_fields.append(_ov_date)
            reqs = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.request', 'search_read', [[]],
                {'fields': sorted(set(_req_ov_fields)), 'limit': 5000},
            )
            month_costs = defaultdict(float)
            eq_costs: dict = {}
            for r in reqs:
                st = r.get('stage_id')
                sname = st[1].lower() if isinstance(st, list) and len(st) > 1 else ''
                is_closed = any(x in sname for x in ('termin', 'done', 'clot', 'close'))
                if is_closed:
                    kpi_ordres_clos += 1
                else:
                    kpi_ordres_ouverts += 1
                if (r.get('maintenance_type') or '').lower() in ('preventive', 'préventive'):
                    kpi_preventive += 1
                dur = float(r.get('duration') or 0)
                cost = dur * 150
                total_maintenance_cost += cost
                if _ov_date:
                    dt = str(r.get(_ov_date) or '')[:10]
                    if len(dt) == 10:
                        mk = f"{dt[5:7]}/{dt[0:4]}"
                        month_costs[mk] += cost
                eq = r.get('equipment_id')
                if isinstance(eq, list) and len(eq) > 1:
                    eid = eq[0]
                    if eid not in eq_costs:
                        eq_costs[eid] = {'name': eq[1], 'cost': 0.0}
                    eq_costs[eid]['cost'] += cost

            def _mk_sort(k):
                p = k.split('/')
                return f"{p[1]}-{p[0]}" if len(p) == 2 else k

            sorted_months = sorted(month_costs.items(), key=lambda x: _mk_sort(x[0]))
            chart_months = {k: round(v, 0) for k, v in sorted_months[-12:]}
            top5 = sorted(eq_costs.values(), key=lambda x: x['cost'], reverse=True)[:5]
            chart_top_eq = {e['name'][:30]: round(e['cost'], 0) for e in top5}
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'

    availability_rate = round(kpi_en_service * 100.0 / max(1, kpi_total_equipements), 1)
    mtbf_days = round(kpi_en_service / max(1, kpi_total_equipements) * 30, 1)
    total_reqs = kpi_ordres_ouverts + kpi_ordres_clos
    preventive_rate = round(kpi_preventive * 100.0 / max(1, total_reqs), 1)
    return _render_parc_module(
        request,
        current_key='overview',
        page_title="Vue d'ensemble",
        page_subtitle='Gestion des equipements et suivi des interventions',
        extra_context={
            'error': error,
            'kpi_total_equipements': kpi_total_equipements,
            'kpi_en_service': kpi_en_service,
            'kpi_en_maintenance': kpi_en_maintenance,
            'kpi_ordres_ouverts': kpi_ordres_ouverts,
            'kpi_ordres_clos': kpi_ordres_clos,
            'total_maintenance_cost': round(total_maintenance_cost, 0),
            'equipment_breakdown': kpi_en_maintenance,
            'availability_rate': availability_rate,
            'mtbf_days': mtbf_days,
            'preventive_rate': preventive_rate,
            'chart_categories_json': json.dumps(dict(cat_counts)),
            'chart_months_json': json.dumps(chart_months),
            'chart_top_eq_json': json.dumps(chart_top_eq),
            'chart_orders_json': json.dumps({'En cours': kpi_ordres_ouverts, 'Terminés': kpi_ordres_clos}),
        },
    )


@login_required
def parc_equipements(request):
    error = None
    rows = []
    kpi_total = kpi_actifs = kpi_breakdown = 0
    raw_search = (request.GET.get('search') or request.GET.get('q') or '').strip()
    search = raw_search.lower()
    category_filter = (request.GET.get('category') or '').strip()
    state_filter = (request.GET.get('state') or '').strip().lower()
    supplier_filter = (request.GET.get('supplier') or '').strip()
    maint_date_filter = (request.GET.get('maint_date') or '').strip()
    maint_date_obj = None
    if maint_date_filter:
        try:
            maint_date_obj = datetime.strptime(maint_date_filter, '%Y-%m-%d').date()
        except ValueError:
            maint_date_obj = None
    categories = set()
    category_choices = []
    supplier_choices = []
    camion_count = semi_count = voiture_count = autre_count = 0

    def _parse_dt(value):
        if not value:
            return None
        s = str(value).replace('T', ' ')
        if len(s) >= 10:
            try:
                return datetime.strptime(s[:10], '%Y-%m-%d').date()
            except ValueError:
                return None
        return None

    def _fmt_dt(value):
        if not value:
            return '—'
        s = str(value).replace('T', ' ')
        return s[:10] if len(s) >= 10 else s
    try:
        try:
            parc_service = ParcOdooService()
            category_choices = [
                c for c in parc_service.get_equipment_categories()
                if c.get('name')
            ]
            supplier_choices = [
                s for s in parc_service.get_suppliers()
                if s.get('name')
            ]
        except Exception:
            category_choices = []
            supplier_choices = []
        uid, models = get_odoo_connection()
        fields_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS, 'maintenance.equipment', 'fields_get', [], {'attributes': ['type']}
        ) or {}
        available = set(fields_meta.keys())
        read_fields = ['name']
        for f in (
            'category_id', 'serial_no', 'state', 'active',
            'maintenance_team_id', 'maintenance_team_ids',
            'technician_user_id', 'owner_user_id', 'last_maintenance_date',
            'model_id', 'x_model_id', 'x_model',
            'model', 'subcategory_id', 'uom_id', 'employee_id',
            'equipment_assign_to', 'maintenance_type', 'description',
            'supplier_id', 'partner_id', 'vendor_id',
        ):
            if f in available:
                read_fields.append(f)
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read', [[]],
            {'fields': sorted(set(read_fields)), 'limit': 4000, 'order': 'name asc'},
        )
        # Consolidation depuis les demandes de maintenance par equipement
        req_fallback_by_eq = {}
        try:
            req_fields_meta = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS, 'maintenance.request', 'fields_get', [], {'attributes': ['type']}
            ) or {}
            req_available = set(req_fields_meta.keys())
            req_read = ['equipment_id']
            for rf in (
                'maintenance_team_id', 'technician_user_id', 'owner_user_id',
                'close_date', 'repair_date', 'request_date', 'create_date',
                'x_chauffeur', 'x_studio_chauffeur',
            ):
                if rf in req_available:
                    req_read.append(rf)
            req_rows = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.request', 'search_read', [[]],
                {'fields': sorted(set(req_read)), 'limit': 8000, 'order': 'id desc'},
            )
            for rr in req_rows:
                eq = rr.get('equipment_id')
                eq_id = eq[0] if isinstance(eq, list) and len(eq) > 0 else None
                if not eq_id:
                    continue
                current = req_fallback_by_eq.get(eq_id) or {
                    'team': '',
                    'technician': '',
                    'driver': '',
                    'date': '',
                }
                req_team = ''
                mt = rr.get('maintenance_team_id')
                if isinstance(mt, list) and len(mt) > 1 and mt[1]:
                    req_team = mt[1]
                req_tech = ''
                t = rr.get('technician_user_id')
                o = rr.get('owner_user_id')
                if isinstance(t, list) and len(t) > 1 and t[1]:
                    req_tech = t[1]
                elif isinstance(o, list) and len(o) > 1 and o[1]:
                    req_tech = o[1]
                req_driver = rr.get('x_chauffeur') or rr.get('x_studio_chauffeur') or ''
                req_date = rr.get('close_date') or rr.get('repair_date') or rr.get('request_date') or rr.get('create_date') or ''

                if not current['team'] and req_team:
                    current['team'] = req_team
                if not current['technician'] and req_tech:
                    current['technician'] = req_tech
                if not current['driver'] and req_driver:
                    current['driver'] = req_driver
                if not current['date'] and req_date:
                    current['date'] = req_date
                req_fallback_by_eq[eq_id] = current
        except Exception:
            req_fallback_by_eq = {}
        team_name_by_id = {}
        if 'maintenance_team_ids' in available:
            team_ids = set()
            for rec in records:
                for tid in (rec.get('maintenance_team_ids') or []):
                    if isinstance(tid, int):
                        team_ids.add(tid)
            if team_ids:
                try:
                    team_rows = models.execute_kw(
                        settings.ODOO_DB, uid, settings.ODOO_PASS,
                        'maintenance.team', 'read', [list(team_ids), ['name']],
                    )
                    team_name_by_id = {tr.get('id'): (tr.get('name') or '') for tr in team_rows}
                except Exception:
                    team_name_by_id = {}
        for r in records:
            req_fb = req_fallback_by_eq.get(r.get('id')) or {}
            cat = r.get('category_id')
            subcat = r.get('subcategory_id')
            tech = r.get('technician_user_id')
            owner = r.get('owner_user_id')
            employee = r.get('employee_id')
            uom = r.get('uom_id')
            team = r.get('maintenance_team_id')
            team_ids = r.get('maintenance_team_ids') or []
            model_m2o = r.get('model_id') or r.get('x_model_id')
            model_txt = (
                model_m2o[1] if isinstance(model_m2o, list) and len(model_m2o) > 1
                else (r.get('model') or r.get('x_model') or '—')
            )
            state = (r.get('state') or ('active' if r.get('active', True) else 'inactive'))
            category_name = cat[1] if isinstance(cat, list) and len(cat) > 1 else '—'
            subcategory_name = subcat[1] if isinstance(subcat, list) and len(subcat) > 1 else '—'
            technician_name = (
                tech[1] if isinstance(tech, list) and len(tech) > 1
                else (owner[1] if isinstance(owner, list) and len(owner) > 1
                else (req_fb.get('technician') or '—'))
            )
            chauffeur_name = (
                employee[1] if isinstance(employee, list) and len(employee) > 1
                else (req_fb.get('driver') or '—')
            )
            if isinstance(team, list) and len(team) > 1:
                team_name = team[1]
            elif team_ids:
                team_name = ', '.join([team_name_by_id.get(tid, '') for tid in team_ids if team_name_by_id.get(tid, '')]) or 'Non affectée'
            elif req_fb.get('team'):
                team_name = req_fb.get('team')
            else:
                team_name = '—'
            state_norm = str(state).lower()
            description_text = str(r.get('description') or '').upper()
            assign_to = r.get('equipment_assign_to') or '—'
            maintenance_type = r.get('maintenance_type') or 'Non renseigné'
            if assign_to == 'department':
                assign_label = 'Département'
            elif assign_to == 'employee':
                assign_label = 'Employé'
            elif assign_to == 'other':
                assign_label = 'Autre'
            else:
                assign_label = str(assign_to).capitalize() if assign_to != '—' else assign_to
            if maintenance_type == 'internal':
                maintenance_type_label = 'Interne'
            elif maintenance_type == 'external':
                maintenance_type_label = 'Externe'
            elif maintenance_type == 'rental':
                maintenance_type_label = 'Location'
            else:
                maintenance_type_label = str(maintenance_type).capitalize() if maintenance_type != 'Non renseigné' else maintenance_type
            # Important: "INACTIF" contient "ACTIF", donc on teste d'abord l'inactif.
            if any(k in description_text for k in ('INACTIF', 'NON ACTIF', 'ARRET', 'ARRÊT', 'ARRETE', 'ARRÊTE', 'HORS SERVICE', 'STOP')):
                state_code = 'inactive'
            elif 'PANNE' in description_text or 'BREAKDOWN' in description_text or 'BROKEN' in description_text:
                state_code = 'breakdown'
            elif re.search(r'\bACTIF\b', description_text):
                state_code = 'active'
            elif state_norm in ('active', 'running', 'en service'):
                state_code = 'active'
            elif state_norm in ('breakdown', 'panne', 'broken', 'repair'):
                state_code = 'breakdown'
            else:
                state_code = 'inactive'
            categories.add(category_name)
            last_maintenance_value = r.get('last_maintenance_date') or req_fb.get('date')
            last_maintenance_dt = _parse_dt(last_maintenance_value)
            supplier_m2o = r.get('supplier_id') or r.get('partner_id') or r.get('vendor_id')
            supplier_name = supplier_m2o[1] if isinstance(supplier_m2o, list) and len(supplier_m2o) > 1 else ''
            row = {
                'id': r.get('id') or 0,
                'name': r.get('name') or '—',
                'model': model_txt,
                'category': category_name,
                'subcategory': subcategory_name,
                'serial_no': r.get('serial_no') or '—',
                'serial_number': r.get('serial_no') or '—',
                'state': state_code,
                'computed_state': state_code,
                'team': team_name,
                'technician': technician_name,
                'chauffeur': chauffeur_name,
                'uom': uom[1] if isinstance(uom, list) and len(uom) > 1 else '—',
                'assign_to': assign_label,
                'maintenance_type': maintenance_type_label,
                'supplier': supplier_name,
                'last_maintenance': last_maintenance_dt,
                'last_maintenance_date': _fmt_dt(last_maintenance_value),
            }
            if category_filter and row['category'] != category_filter:
                continue
            # Filtrage état après calcul de computed_state
            if state_filter and row['computed_state'] != state_filter:
                continue
            if supplier_filter and row['supplier'] != supplier_filter:
                continue
            if maint_date_obj and (not row['last_maintenance'] or row['last_maintenance'] < maint_date_obj):
                continue
            if search:
                hay = f"{row['name']} {row['model']} {row['category']} {row['subcategory']} {row['serial_no']} {row['supplier']}".lower()
                if search not in hay:
                    continue
            rows.append(row)
            if row['state'] == 'active':
                kpi_actifs += 1
            elif row['state'] == 'breakdown':
                kpi_breakdown += 1
            cat_u = row['category'].upper()
            if 'CAMION' in cat_u:
                camion_count += 1
            elif 'SEMI' in cat_u:
                semi_count += 1
            elif 'VOITURE' in cat_u:
                voiture_count += 1
            else:
                autre_count += 1
        kpi_total = len(rows)
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return render(request, 'parc/equipements.html', {
        'error': error,
        'equipements': rows,
        'rows': rows,
        'kpi_total': kpi_total,
        'kpi_actifs': kpi_actifs,
        'kpi_breakdown': kpi_breakdown,
        'camion_count': camion_count,
        'semi_count': semi_count,
        'voiture_count': voiture_count,
        'autre_count': autre_count,
        'search': raw_search,
        'category': category_filter,
        'state': state_filter,
        'supplier': supplier_filter,
        'maint_date': maint_date_filter,
        'filters': {
            'q': raw_search,
            'category': category_filter,
            'state': state_filter,
            'supplier': supplier_filter,
            'maint_date': maint_date_filter,
        },
        'has_active_filters': bool(
            raw_search or category_filter or state_filter or supplier_filter or maint_date_filter
        ),
        'results_count': len(rows),
        'categories': sorted(category_choices, key=lambda c: c.get('name', '')) if category_choices else [{'name': c} for c in sorted(categories)],
        'suppliers': sorted(supplier_choices, key=lambda s: s.get('name', '')),
    })


@login_required
def parc_disponibilite(request):
    error = None
    rows = []
    kpi_disponibles = kpi_indisponibles = 0
    kpi_taux = 0.0
    q = (request.GET.get('q') or '').strip().lower()
    try:
        uid, models = get_odoo_connection()
        eq_fields = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        eq_available = set(eq_fields.keys())
        eq_read = ['name']
        for f in ('active', 'state', 'description', 'category_id', 'company_id'):
            if f in eq_available:
                eq_read.append(f)
        equipments = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': sorted(set(eq_read)), 'limit': 2000, 'order': 'name asc'},
        )

        # Equipements avec maintenance corrective ouverte => indisponibles.
        corrective_busy_ids = set()
        try:
            req_fields = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.request', 'fields_get', [], {'attributes': ['type']},
            ) or {}
            req_available = set(req_fields.keys())
            if 'equipment_id' in req_available and 'stage_id' in req_available:
                req_domain = [('stage_id.done', '=', False)]
                if 'maintenance_type' in req_available:
                    req_domain.append(('maintenance_type', '=', 'corrective'))
                req_read_fields = ['equipment_id']
                if 'maintenance_type' in req_available:
                    req_read_fields.append('maintenance_type')
                req_records = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'maintenance.request', 'search_read',
                    [req_domain],
                    {'fields': req_read_fields, 'limit': 5000},
                )
                for rr in req_records:
                    eq = rr.get('equipment_id')
                    if isinstance(eq, list) and eq:
                        corrective_busy_ids.add(eq[0])
        except Exception:
            corrective_busy_ids = set()

        for eq in equipments:
            eq_id = eq.get('id')
            is_active = bool(eq.get('active', True))
            raw_state = str(eq.get('state') or '').lower()
            desc = str(eq.get('description') or '').upper()
            # Logique métier: indisponible uniquement si inactif ou explicitement en panne/repair.
            is_breakdown = (
                raw_state in ('breakdown', 'panne', 'broken', 'repair', 'reparation')
                or 'PANNE' in desc
                or 'BREAKDOWN' in desc
                or 'BROKEN' in desc
            )
            is_corrective_busy = eq_id in corrective_busy_ids
            status = 'Indisponible' if (not is_active or is_breakdown or is_corrective_busy) else 'Disponible'
            cat = eq.get('category_id')
            comp = eq.get('company_id')
            taux = 100 if status == 'Disponible' else 0
            row = {
                'equipement': eq.get('name') or '—',
                'categorie': cat[1] if isinstance(cat, list) and len(cat) > 1 else '—',
                'localisation': comp[1] if isinstance(comp, list) and len(comp) > 1 else '—',
                'jours_service': 30 if status == 'Disponible' else 0,
                'jours_panne': 0 if status == 'Disponible' else 30,
                'taux': taux,
                'etat': status,
            }
            if q:
                hay = f"{row['equipement']} {row['categorie']} {row['localisation']} {row['etat']}".lower()
                if q not in hay:
                    continue
            rows.append(row)
            if status == 'Disponible':
                kpi_disponibles += 1
            else:
                kpi_indisponibles += 1
        total = kpi_disponibles + kpi_indisponibles
        if total:
            kpi_taux = round((kpi_disponibles * 100.0) / total, 1)
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'parc/disponibilite.html', {
        'error': error,
        'rows': rows,
        'kpi_dispo': kpi_disponibles,
        'kpi_indispo': kpi_indisponibles,
        'kpi_global': kpi_taux,
        'kpi_mtbf': round((kpi_disponibles / max(1, len(rows))) * 30, 1),
        'filters': {'q': request.GET.get('q', '').strip()},
    })


@login_required
def parc_ordres_maintenance(request):
    """Vue unifiée Ordres + Coûts + Durées (3 onglets)."""
    error = None
    rows = []
    kpi_total = kpi_ouverts = kpi_clos = kpi_retard = 0
    total_cout = total_heures = kpi_preventive = 0.0
    today_date = date.today()
    _month_costs: dict = defaultdict(float)
    _cat_costs: dict = defaultdict(float)
    _tech_cnt: dict = defaultdict(int)
    _tech_dur: dict = defaultdict(float)
    _eq_c: dict = {}
    _trend: dict = defaultdict(float)

    try:
        uid, models = get_odoo_connection()
        fields_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        available = set(fields_meta.keys())
        date_field = next((f for f in ('request_date', 'create_date') if f in available), None)
        fields = ['name']
        for f in ('equipment_id', 'maintenance_type', 'owner_user_id', 'technician_user_id',
                  'stage_id', 'description', 'category_id', 'duration'):
            if f in available:
                fields.append(f)
        _CP = ('cost_amount', 'total_cost', 'amount_total',
               'maintenance_cost', 'effective_cost', 'cost')
        _CPP = ('parts_cost', 'labor_cost', 'spare_parts_cost')
        _cost_used = [f for f in (*_CP, *_CPP) if f in available]
        fields.extend(_cost_used)
        if date_field:
            fields.append(date_field)
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'search_read',
            [[]],
            {'fields': sorted(set(fields)), 'limit': 2000,
             'order': f'{date_field} desc' if date_field else 'id desc'},
        )
        for r in records:
            dt = str(r.get(date_field) or '')[:10] if date_field else ''
            eq = r.get('equipment_id')
            own = r.get('owner_user_id')
            tech = r.get('technician_user_id') or own
            stage = r.get('stage_id')
            cat = r.get('category_id')
            eq_name   = eq[1]   if isinstance(eq, list)    and len(eq)    > 1 else '—'
            eq_id     = eq[0]   if isinstance(eq, list)    and len(eq)    > 0 else None
            cat_name  = cat[1]  if isinstance(cat, list)   and len(cat)   > 1 else '—'
            tech_name = tech[1] if isinstance(tech, list)  and len(tech)  > 1 else '—'
            stage_name = stage[1] if isinstance(stage, list) and len(stage) > 1 else '—'
            stage_l = stage_name.lower()
            is_closed = any(k in stage_l for k in ('done', 'close', 'clôt', 'terminé', 'termine'))
            days_open = 0
            is_retard = False
            if not is_closed and len(dt) == 10:
                try:
                    days_open = (today_date - date.fromisoformat(dt)).days
                    is_retard = days_open > 30
                except ValueError:
                    pass
            dur = float(r.get('duration') or 0)
            cout = 0.0
            for _cf in _CP:
                if _cf in _cost_used:
                    _v = float(r.get(_cf) or 0)
                    if _v > 0:
                        cout = _v
                        break
            if not cout:
                _ps = sum(float(r.get(f) or 0) for f in _CPP if f in _cost_used)
                cout = _ps if _ps > 0 else dur * 150
            cout = round(cout, 2)
            raw_desc = r.get('description') or ''
            clean_desc = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', html.unescape(raw_desc))).strip()[:140] or '—'
            mtype = (r.get('maintenance_type') or '').lower()
            _type_map = {'corrective': 'Corrective', 'preventive': 'Préventive',
                         'préventive': 'Préventive', 'diagnostic': 'Diagnostic'}
            type_label = _type_map.get(mtype, mtype.capitalize() or '—')
            if is_closed:
                kpi_clos += 1
            else:
                kpi_ouverts += 1
            if is_retard:
                kpi_retard += 1
            if mtype in ('preventive', 'préventive'):
                kpi_preventive += 1
            total_cout  += cout
            total_heures += dur
            if len(dt) == 10:
                mk = f"{dt[5:7]}/{dt[0:4]}"
                _month_costs[mk] += cout
                try:
                    diff = (today_date - date.fromisoformat(dt)).days
                    if 0 <= diff <= 30:
                        _trend[f"{dt[8:10]}/{dt[5:7]}"] += dur
                except ValueError:
                    pass
            _cat_costs[cat_name] += cout
            _tech_cnt[tech_name] += 1
            _tech_dur[tech_name] += dur
            if eq_id:
                if eq_id not in _eq_c:
                    _eq_c[eq_id] = {'name': eq_name, 'cost': 0.0}
                _eq_c[eq_id]['cost'] += cout
            rows.append({
                'name': r.get('name') or '—',
                'date': f"{dt[8:10]}/{dt[5:7]}/{dt[0:4]}" if len(dt) == 10 else '—',
                'date_iso': dt,
                'equipement': eq_name,
                'categorie': cat_name,
                'type': type_label,
                'technicien': tech_name,
                'etat': stage_name,
                'is_closed': is_closed,
                'is_retard': is_retard,
                'days_open': days_open,
                'cout': cout,
                'duree': dur,
                'description': clean_desc,
            })
        kpi_total = len(rows)

        def _ms(k):
            p = k.split('/')
            return f"{p[1]}-{p[0]}" if len(p) == 2 else k

        chart_mois = {k: round(v, 0) for k, v in sorted(_month_costs.items(), key=lambda x: _ms(x[0]))[-12:]}
        chart_cat  = {k: round(v, 0) for k, v in sorted(_cat_costs.items(), key=lambda x: x[1], reverse=True)[:8]}
        top5_tech  = dict(sorted(_tech_cnt.items(), key=lambda x: x[1], reverse=True)[:5])
        top5_dur   = {k: round(v, 2) for k, v in sorted(_tech_dur.items(), key=lambda x: x[1], reverse=True)[:5]}
        top10_c    = {r['name'][:22]: r['cout'] for r in sorted(rows, key=lambda x: x['cout'], reverse=True)[:10] if r['cout'] > 0}
        top_eq_list = sorted(_eq_c.values(), key=lambda x: x['cost'], reverse=True)
        max_eq = top_eq_list[0] if top_eq_list else {'name': '—', 'cost': 0.0}
        if chart_mois:
            bm = max(chart_mois, key=chart_mois.get)
            kpi_top_month = f"{bm} — {chart_mois[bm]:,.0f} MAD".replace(',', ' ')
        else:
            kpi_top_month = '—'
        corr_h = sum(r['duree'] for r in rows if 'corrective' in r['type'].lower())
        prev_h = sum(r['duree'] for r in rows if 'prév' in r['type'].lower() or 'prev' in r['type'].lower())
        chart_trend = {k: round(v, 2) for k, v in sorted(_trend.items(), key=lambda x: (x[0][3:], x[0][:2]))}

    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'
        chart_mois = chart_cat = top5_tech = top5_dur = top10_c = chart_trend = {}
        kpi_top_month = '—'
        max_eq = {'name': '—', 'cost': 0.0}
        corr_h = prev_h = 0.0

    preventive_rate = round(kpi_preventive * 100.0 / max(1, kpi_total), 1)
    top_tech = max(top5_tech, key=top5_tech.get) if top5_tech else '—'
    rows_by_cout  = sorted(rows, key=lambda x: x['cout'],  reverse=True)
    rows_by_duree = sorted(rows, key=lambda x: x['duree'], reverse=True)

    return render(request, 'parc/ordres.html', {
        'error':           error,
        'rows':            rows,
        'rows_by_cout':    rows_by_cout,
        'rows_by_duree':   rows_by_duree,
        'kpi_total':       kpi_total,
        'kpi_en_cours':    kpi_ouverts,
        'kpi_termines':    kpi_clos,
        'kpi_retard':      kpi_retard,
        'total_cout':      round(total_cout, 0),
        'cout_moyen':      round(total_cout / max(1, kpi_total), 2),
        'top_eq_name':     max_eq['name'][:35] if max_eq['name'] != '—' else '—',
        'top_eq_cost':     round(max_eq['cost'], 0),
        'kpi_top_month':   kpi_top_month,
        'total_heures':    round(total_heures, 1),
        'duree_moyenne':   round(total_heures / max(1, kpi_total), 2),
        'preventive_rate': preventive_rate,
        'top_tech':        top_tech,
        'chart_statut_json':  json.dumps({'En cours': kpi_ouverts, 'Terminés': kpi_clos, 'En retard': kpi_retard}),
        'chart_tech_json':    json.dumps(top5_tech),
        'chart_mois_json':    json.dumps(chart_mois),
        'chart_cat_json':     json.dumps(chart_cat),
        'chart_top10_json':   json.dumps(top10_c),
        'chart_tech_dur_json': json.dumps(top5_dur),
        'chart_type_h_json':  json.dumps({'Corrective': round(corr_h, 2), 'Préventive': round(prev_h, 2)}),
        'chart_trend_json':   json.dumps(chart_trend),
    })


@login_required
def parc_interventions(request):
    error = None
    rows = []
    kpi_total = 0
    try:
        uid, models = get_odoo_connection()
        fields_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        available = set(fields_meta.keys())
        tech_field = 'technician_user_id' if 'technician_user_id' in available else ('owner_user_id' if 'owner_user_id' in available else None)
        date_field = 'request_date' if 'request_date' in available else ('create_date' if 'create_date' in available else None)
        read_fields = ['name', 'equipment_id', 'maintenance_type', 'duration', 'description', 'stage_id']
        if date_field:
            read_fields.append(date_field)
        if tech_field:
            read_fields.append(tech_field)
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'search_read', [[]],
            {'fields': sorted(set(read_fields)), 'limit': 1500, 'order': f'{date_field} desc' if date_field else 'id desc'},
        )
        for r in records:
            dt = (r.get(date_field) or '')[:10] if date_field else ''
            eq = r.get('equipment_id')
            tech = r.get(tech_field) if tech_field else None
            st = r.get('stage_id')
            rows.append({
                'name': r.get('name') or '—',
                'date': f"{dt[8:10]}/{dt[5:7]}/{dt[0:4]}" if len(dt) == 10 else '—',
                'equipement': eq[1] if isinstance(eq, list) and len(eq) > 1 else '—',
                'technicien': tech[1] if isinstance(tech, list) and len(tech) > 1 else '—',
                'type': (r.get('maintenance_type') or '—').capitalize(),
                'duree': float(r.get('duration') or 0),
                'description': r.get('description') or '—',
                'etat': st[1] if isinstance(st, list) and len(st) > 1 else '—',
            })
        kpi_total = len(rows)
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    total_h = sum(r.get('duree', 0) for r in rows)
    top_tech = '—'
    if rows:
        counts = defaultdict(int)
        for r in rows:
            counts[r['technicien']] += 1
        top_tech = sorted(counts.items(), key=lambda x: x[1], reverse=True)[0][0]
    type_counts = defaultdict(int)
    tech_durations: dict = {}
    for r in rows:
        tp = r.get('type') or '—'
        type_counts[tp] += 1
        tech = r.get('technicien') or '—'
        tech_durations[tech] = round(tech_durations.get(tech, 0.0) + (r.get('duree') or 0), 2)
    corrective = type_counts.get('Corrective', 0)
    preventive = sum(v for k, v in type_counts.items() if 'prev' in k.lower() or 'prév' in k.lower())
    preventive_rate = round(preventive * 100 / max(1, kpi_total), 1)
    top5_techs = dict(sorted(tech_durations.items(), key=lambda x: x[1], reverse=True)[:5])
    return render(request, 'parc/interventions.html', {
        'error': error,
        'rows': rows,
        'kpi_total': kpi_total,
        'kpi_total_h': total_h,
        'kpi_avg_h': round(total_h / max(1, kpi_total), 2),
        'kpi_top_tech': top_tech,
        'preventive_rate': preventive_rate,
        'chart_type_json': json.dumps({'Corrective': corrective, 'Préventive': preventive}),
        'chart_tech_json': json.dumps(top5_techs),
        'chart_json': json.dumps(dict(type_counts)),
    })


@login_required
def parc_couts(request):
    from django.shortcuts import redirect
    return redirect('parc_ordres_maintenance')


@login_required
def parc_fournisseurs(request):
    error = None
    suppliers = []
    try:
        uid, models = get_odoo_connection()
        fields_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'res.partner', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        available = set(fields_meta.keys())
        read_fields = ['name', 'active']
        for f in ('phone', 'email', 'supplier_rank', 'company_type', 'city', 'street'):
            if f in available:
                read_fields.append(f)
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'res.partner', 'search_read',
            [[('supplier_rank', '>', 0)]],
            {'fields': sorted(set(read_fields)), 'limit': 500, 'order': 'name asc'},
        )
        for r in records:
            suppliers.append({
                'name': r.get('name') or '—',
                'phone': r.get('phone') or '—',
                'email': r.get('email') or '—',
                'rank': r.get('supplier_rank') or 0,
                'city': r.get('city') or '—',
                'is_active': bool(r.get('active', True)),
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    active_count = sum(1 for s in suppliers if s['is_active'])
    return render(request, 'parc/fournisseurs.html', {
        'error': error,
        'suppliers': suppliers,
        'total_suppliers': len(suppliers),
        'active_suppliers': active_count,
        'inactive_suppliers': len(suppliers) - active_count,
    })


@login_required
def parc_equipment_detail(request, equipment_id):
    error = None
    equipment = None
    interventions = []
    try:
        uid, models = get_odoo_connection()
        eq_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        eq_available = set(eq_meta.keys())
        eq_read = ['name']
        for f in ('category_id', 'serial_no', 'state', 'active', 'maintenance_team_id',
                  'technician_user_id', 'owner_user_id', 'model_id', 'x_model_id',
                  'model', 'subcategory_id', 'uom_id', 'partner_id',
                  'last_maintenance_date', 'maintenance_type', 'description'):
            if f in eq_available:
                eq_read.append(f)
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'read',
            [[equipment_id]], {'fields': sorted(set(eq_read))},
        )
        if records:
            r = records[0]
            cat = r.get('category_id')
            subcat = r.get('subcategory_id')
            team = r.get('maintenance_team_id')
            tech = r.get('technician_user_id') or r.get('owner_user_id')
            model_m2o = r.get('model_id') or r.get('x_model_id')
            uom = r.get('uom_id')
            supplier = r.get('partner_id')
            state_raw = str(r.get('state') or ('active' if r.get('active', True) else 'inactive')).lower()
            desc_text = str(r.get('description') or '').upper()
            if 'ACTIF' in desc_text:
                computed_state = 'active'
                state_label = 'Actif'
            elif 'PANNE' in desc_text or 'BREAKDOWN' in desc_text:
                computed_state = 'breakdown'
                state_label = 'Panne'
            elif state_raw in ('active', 'running', 'en service'):
                computed_state = 'active'
                state_label = 'Actif'
            elif state_raw in ('breakdown', 'panne', 'broken', 'repair'):
                computed_state = 'breakdown'
                state_label = 'Panne'
            else:
                computed_state = 'stopped'
                state_label = 'Arrêté'
            mt = r.get('maintenance_type') or ''
            mt_label = {'internal': 'Interne', 'external': 'Externe', 'rental': 'Location'}.get(mt, mt.capitalize() or '—')
            equipment = {
                'name': r.get('name') or '—',
                'category': cat[1] if isinstance(cat, list) and len(cat) > 1 else '—',
                'subcategory': subcat[1] if isinstance(subcat, list) and len(subcat) > 1 else '—',
                'serial_no': r.get('serial_no') or '—',
                'state': state_label,
                'computed_state': computed_state,
                'description': str(r.get('description') or ''),
                'team': team[1] if isinstance(team, list) and len(team) > 1 else '—',
                'technician': tech[1] if isinstance(tech, list) and len(tech) > 1 else '—',
                'model': (model_m2o[1] if isinstance(model_m2o, list) and len(model_m2o) > 1
                          else (r.get('model') or r.get('x_model') or '—')),
                'uom': uom[1] if isinstance(uom, list) and len(uom) > 1 else '—',
                'supplier': supplier[1] if isinstance(supplier, list) and len(supplier) > 1 else '—',
                'last_maintenance': (r.get('last_maintenance_date') or '')[:10],
                'maintenance_type': mt_label,
            }
        req_meta = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        req_available = set(req_meta.keys())
        req_read = ['name', 'equipment_id']
        for f in ('maintenance_type', 'duration', 'stage_id', 'description',
                  'technician_user_id', 'owner_user_id', 'request_date', 'create_date'):
            if f in req_available:
                req_read.append(f)
        date_f = 'request_date' if 'request_date' in req_available else 'create_date'
        tech_f = ('technician_user_id' if 'technician_user_id' in req_available
                  else ('owner_user_id' if 'owner_user_id' in req_available else None))
        req_records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'search_read',
            [[('equipment_id', '=', equipment_id)]],
            {'fields': sorted(set(req_read)), 'limit': 300, 'order': f'{date_f} desc'},
        )
        for rr in req_records:
            dt = (rr.get(date_f) or '')[:10]
            tech_val = rr.get(tech_f) if tech_f else None
            st = rr.get('stage_id')
            dur = float(rr.get('duration') or 0)
            interventions.append({
                'name': rr.get('name') or '—',
                'date': f"{dt[8:10]}/{dt[5:7]}/{dt[0:4]}" if len(dt) == 10 else '—',
                'type': (rr.get('maintenance_type') or '—').capitalize(),
                'technician': tech_val[1] if isinstance(tech_val, list) and len(tech_val) > 1 else '—',
                'duration': dur,
                'cost': round(dur * 150, 2),
                'state': st[1] if isinstance(st, list) and len(st) > 1 else '—',
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return render(request, 'parc/equipment_detail.html', {
        'error': error,
        'equipment': equipment,
        'equipment_id': equipment_id,
        'interventions': interventions,
        'total_cost': sum(i['cost'] for i in interventions),
        'total_hours': sum(i['duration'] for i in interventions),
    })


@login_required
def parc_demandes(request):
    """Liste des demandes de maintenance avec tous les champs du responsable."""
    error = None
    rows = []
    kpi_total = kpi_new = kpi_in_progress = kpi_done = kpi_urgent = 0
    f_stage    = (request.GET.get('stage') or '').strip()
    f_type     = (request.GET.get('type') or '').strip()
    f_team     = (request.GET.get('team') or '').strip()
    f_priority = (request.GET.get('priority') or '').strip()
    f_date_debut = (request.GET.get('date_debut') or '').strip()
    f_date_fin   = (request.GET.get('date_fin') or '').strip()
    stages = set()
    teams  = set()

    def _fmt_date(val):
        if not val:
            return '—'
        s = str(val).replace('T', ' ')
        d = s[:10]
        hm = s[11:16] if len(s) >= 16 else ''
        if len(d) == 10:
            base = f"{d[8:10]}/{d[5:7]}/{d[0:4]}"
            return f"{base} {hm}".strip()
        return s

    try:
        uid, models = get_odoo_connection()
        avail = set(
            models.execute_kw(settings.ODOO_DB, uid, settings.ODOO_PASS,
                              'maintenance.request', 'fields_get', [],
                              {'attributes': ['type']}) or {}
        )
        rf = {'name', 'id', 'equipment_id', 'maintenance_type', 'stage_id',
              'maintenance_team_id', 'priority', 'description'}
        for f in (
            'request_date', 'create_date', 'owner_user_id', 'technician_user_id',
            'category_id', 'company_id', 'duration',
            'x_date_panne', 'x_date_reparation', 'x_date_prevue', 'x_date_cloture',
            'x_atelier', 'x_chauffeur', 'x_type_intervention', 'x_sous_categorie',
            'x_plan_maintenance_id', 'x_modele_intervention', 'x_udm', 'x_dernier_releve',
            'x_temps_estime', 'x_temps_reparation', 'x_ecart',
            'x_studio_type_intervention', 'x_studio_atelier', 'x_studio_chauffeur',
        ):
            if f in avail:
                rf.add(f)
        date_field = 'request_date' if 'request_date' in avail else 'create_date'
        domain = []
        if f_type:
            domain.append(('maintenance_type', '=', f_type))
        if f_priority:
            domain.append(('priority', '=', f_priority))
        if f_date_debut:
            domain.append((date_field, '>=', f_date_debut))
        if f_date_fin:
            domain.append((date_field, '<=', f_date_fin + ' 23:59:59'))
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'search_read', [domain],
            {'fields': sorted(rf), 'limit': 3000,
             'order': f'{date_field} desc'},
        )
        for r in records:
            eq    = r.get('equipment_id')
            stage = r.get('stage_id')
            team  = r.get('maintenance_team_id')
            owner = r.get('owner_user_id')
            tech  = r.get('technician_user_id')
            cat   = r.get('category_id')
            comp  = r.get('company_id')
            stage_name = stage[1] if isinstance(stage, list) and len(stage) > 1 else '—'
            team_name  = team[1]  if isinstance(team,  list) and len(team)  > 1 else '—'
            priority   = str(r.get('priority') or '0')
            is_done    = any(k in stage_name.lower() for k in ('done', 'terminé', 'clôt', 'close'))
            is_new     = any(k in stage_name.lower() for k in ('nouv', 'new', 'demande'))
            stages.add(stage_name)
            teams.add(team_name)
            if is_done:
                kpi_done += 1
            elif is_new:
                kpi_new += 1
            else:
                kpi_in_progress += 1
            if priority in ('2', '3') and not is_done:
                kpi_urgent += 1
            if f_stage and stage_name != f_stage:
                continue
            if f_team and team_name != f_team:
                continue

            plan_raw = r.get('x_plan_maintenance_id')
            plan = plan_raw[1] if isinstance(plan_raw, list) and len(plan_raw) > 1 else (str(plan_raw) if plan_raw else '—')

            rows.append({
                'id': r.get('id'),
                'odoo_id': r.get('id'),
                'name': r.get('name') or '—',
                'equipement': eq[1] if isinstance(eq, list) and len(eq) > 1 else '—',
                'type': (r.get('maintenance_type') or '—').capitalize(),
                'type_intervention': (
                    r.get('x_type_intervention') or r.get('x_studio_type_intervention') or '—'),
                'categorie': cat[1] if isinstance(cat, list) and len(cat) > 1 else '—',
                'sous_categorie': r.get('x_sous_categorie') or '—',
                'date':           _fmt_date(r.get(date_field)),
                'date_panne':     _fmt_date(r.get('x_date_panne')),
                'date_reparation':_fmt_date(r.get('x_date_reparation')),
                'date_prevue':    _fmt_date(r.get('x_date_prevue')),
                'date_cloture':   _fmt_date(r.get('x_date_cloture')),
                'etat':       stage_name,
                'is_done':    is_done,
                'equipe':     team_name,
                'atelier':    r.get('x_atelier') or r.get('x_studio_atelier') or '—',
                'technicien': (tech[1]  if isinstance(tech,  list) and len(tech)  > 1 else
                               (owner[1] if isinstance(owner, list) and len(owner) > 1 else '—')),
                'chauffeur':  r.get('x_chauffeur') or r.get('x_studio_chauffeur') or '—',
                'priority':   priority,
                'stars':      '★' * int(priority) + '☆' * (3 - int(priority)),
                'plan_maintenance':    plan,
                'modele_intervention': r.get('x_modele_intervention') or '—',
                'udm':             r.get('x_udm') or '—',
                'dernier_releve':  r.get('x_dernier_releve') or '—',
                'temps_estime':    r.get('x_temps_estime') or str(r.get('duration') or '—'),
                'temps_reparation':r.get('x_temps_reparation') or '—',
                'ecart':           r.get('x_ecart') or '—',
                'description': (r.get('description') or '')[:100],
                'company': comp[1] if isinstance(comp, list) and len(comp) > 1 else '—',
            })
        kpi_total = len(rows)
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'

    return render(request, 'parc/demandes.html', {
        'error': error, 'rows': rows,
        'kpi_total': kpi_total, 'kpi_new': kpi_new,
        'kpi_in_progress': kpi_in_progress, 'kpi_done': kpi_done,
        'kpi_urgent': kpi_urgent,
        'filters': {'stage': f_stage, 'type': f_type, 'team': f_team,
                    'priority': f_priority, 'date_debut': f_date_debut, 'date_fin': f_date_fin},
        'stages': sorted(s for s in stages if s != '—'),
        'teams':  sorted(t for t in teams  if t != '—'),
        'has_filters': bool(f_stage or f_type or f_team or f_priority or f_date_debut or f_date_fin),
    })


@login_required
def parc_demande_detail(request, demande_id):
    """Vue détail d'une demande de maintenance avec tous les champs du responsable."""
    error = None
    demande = {}

    def _fmt_date(val):
        if not val:
            return '—'
        d = str(val)[:10]
        return f"{d[8:10]}/{d[5:7]}/{d[0:4]}" if len(d) == 10 else d

    try:
        uid, models = get_odoo_connection()
        avail = set(
            models.execute_kw(settings.ODOO_DB, uid, settings.ODOO_PASS,
                              'maintenance.request', 'fields_get', [],
                              {'attributes': ['type']}) or {}
        )
        rf = {'name', 'id', 'equipment_id', 'maintenance_type', 'stage_id',
              'maintenance_team_id', 'priority', 'description'}
        for f in (
            'request_date', 'create_date', 'owner_user_id', 'technician_user_id',
            'category_id', 'company_id', 'duration',
            'x_date_panne', 'x_date_reparation', 'x_date_prevue', 'x_date_cloture',
            'x_atelier', 'x_chauffeur', 'x_type_intervention', 'x_sous_categorie',
            'x_plan_maintenance_id', 'x_modele_intervention', 'x_udm', 'x_dernier_releve',
            'x_temps_estime', 'x_temps_reparation', 'x_ecart',
            'x_technicien_ids', 'technician_user_ids',
            'x_studio_type_intervention', 'x_studio_atelier', 'x_studio_chauffeur',
            'breakdown_date', 'repair_date', 'planned_date', 'close_date',
            'maintenance_plan_id', 'intervention_model_id', 'unit_type', 'udm', 'uom_id',
            'last_reading', 'time_estimated', 'time_repair', 'time_gap',
            'technicians_id', 'technicians_ids', 'technicians',
        ):
            if f in avail:
                rf.add(f)
        date_field = 'request_date' if 'request_date' in avail else 'create_date'
        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.request', 'read', [[demande_id]],
            {'fields': sorted(rf)},
        )
        if not records:
            error = 'Demande introuvable (id=%d).' % demande_id
        else:
            r = records[0]
            eq    = r.get('equipment_id')
            stage = r.get('stage_id')
            team  = r.get('maintenance_team_id')
            owner = r.get('owner_user_id')
            tech  = r.get('technician_user_id')
            cat   = r.get('category_id')
            comp  = r.get('company_id')
            stage_name = stage[1] if isinstance(stage, list) and len(stage) > 1 else '—'
            priority   = str(r.get('priority') or '0')

            # Techniciens many2many
            technicians = []
            for fld, model_name in (
                ('technicians_id', 'hr.employee'),
                ('technicians_ids', 'hr.employee'),
                ('x_technicien_ids', 'hr.employee'),
                ('technician_user_ids', 'res.users'),
                ('technicians', 'hr.employee'),
            ):
                if fld not in avail:
                    continue
                tech_ids = [x for x in (r.get(fld) or []) if isinstance(x, int)]
                if not tech_ids:
                    continue
                try:
                    tr = models.execute_kw(settings.ODOO_DB, uid, settings.ODOO_PASS,
                                           model_name, 'read', [tech_ids, ['name']])
                    technicians = [t.get('name') for t in tr if t.get('name')]
                    break
                except Exception:
                    pass
            if not technicians and isinstance(tech, list) and len(tech) > 1:
                technicians = [tech[1]]

            plan_raw = r.get('maintenance_plan_id') or r.get('x_plan_maintenance_id')
            plan = plan_raw[1] if isinstance(plan_raw, list) and len(plan_raw) > 1 else (str(plan_raw) if plan_raw else '—')
            model_raw = r.get('intervention_model_id') or r.get('x_modele_intervention')
            intervention_model = model_raw[1] if isinstance(model_raw, list) and len(model_raw) > 1 else (str(model_raw) if model_raw else '—')
            udm_raw = r.get('udm') or r.get('x_udm') or r.get('uom_id')
            udm_val = udm_raw[1] if isinstance(udm_raw, list) and len(udm_raw) > 1 else (str(udm_raw) if udm_raw else '—')
            last_reading_raw = r.get('last_reading')
            if last_reading_raw in (None, False, ''):
                last_reading_raw = r.get('x_dernier_releve')
            unit_type = r.get('unit_type') or 'Unité'

            demande = {
                'id': r.get('id'),
                'name': r.get('name') or '—',
                'equipement': eq[1] if isinstance(eq, list) and len(eq) > 1 else '—',
                'maintenance_type': (r.get('maintenance_type') or '—').capitalize(),
                'type_intervention': (
                    r.get('x_type_intervention') or r.get('x_studio_type_intervention') or '—'),
                'categorie':     cat[1]  if isinstance(cat,  list) and len(cat)  > 1 else '—',
                'sous_categorie':r.get('x_sous_categorie') or '—',
                'date':           _fmt_date(r.get(date_field)),
                'date_panne':     _fmt_date(r.get('breakdown_date') or r.get('x_date_panne')),
                'date_reparation':_fmt_date(r.get('repair_date') or r.get('x_date_reparation')),
                'date_prevue':    _fmt_date(r.get('planned_date') or r.get('x_date_prevue')),
                'date_cloture':   _fmt_date(r.get('close_date') or r.get('x_date_cloture')),
                'etat':      stage_name,
                'is_done':   any(k in stage_name.lower() for k in ('done', 'terminé', 'clôt', 'close')),
                'equipe':    team[1] if isinstance(team, list) and len(team) > 1 else '—',
                'atelier':   r.get('x_atelier') or r.get('x_studio_atelier') or '—',
                'responsable':r.get('owner_user_id') and (owner[1] if isinstance(owner, list) and len(owner) > 1 else '—') or '—',
                'chauffeur': r.get('x_chauffeur') or r.get('x_studio_chauffeur') or '—',
                'priority':  priority,
                'stars':     '★' * int(priority) + '☆' * (3 - int(priority)),
                'plan_maintenance':    plan,
                'modele_intervention': intervention_model,
                'unit_type':        unit_type,
                'udm':              udm_val,
                'dernier_releve':   last_reading_raw if last_reading_raw not in (None, False, '') else '—',
                'temps_estime':     r.get('time_estimated') or r.get('x_temps_estime') or str(r.get('duration') or '—'),
                'temps_reparation': r.get('time_repair') or r.get('x_temps_reparation') or '—',
                'ecart':            r.get('time_gap') or r.get('x_ecart') or '—',
                'description': r.get('description') or '—',
                'company':     comp[1] if isinstance(comp, list) and len(comp) > 1 else '—',
                'technicians': technicians,
            }
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
        logger.error('parc_demande_detail: %s', exc)

    return render(request, 'parc/demande_detail.html', {'error': error, 'demande': demande})


# ── Alias parc ──────────────────────────────────────────────────
# parc_module = vue d'ensemble du parc (identique a parc_overview)
parc_module = parc_overview
# parc_ordres = alias vers la fonction existante parc_ordres_maintenance
parc_ordres = parc_ordres_maintenance


@login_required
def parc_debug(request):
    from django.http import HttpResponse
    try:
        uid, _ = get_odoo_connection()
        info = 'Connexion Odoo OK — uid=%d' % uid
    except Exception as exc:
        info = 'Connexion Odoo FAILED: %s' % exc
    return HttpResponse(
        '<pre style="font-family:monospace;padding:20px;background:#f8fafc">'
        'PARC DEBUG\n%s</pre>' % info
    )


@login_required
def parc_equipements_pdf(request):
    from django.shortcuts import redirect
    return redirect('parc_equipements')


@login_required
def parc_equipements_excel(request):
    from django.shortcuts import redirect
    return redirect('parc_equipements')


@login_required
def parc_equipements_csv(request):
    from django.shortcuts import redirect
    return redirect('parc_equipements')


@login_required
def parc_disponibilite_pdf(request):
    rows = []
    kpi_dispo = 0
    kpi_indispo = 0
    q = (request.GET.get('q') or '').strip().lower()
    try:
        uid, models = get_odoo_connection()
        eq_fields = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        eq_available = set(eq_fields.keys())
        eq_read = ['name']
        for f in ('active', 'category_id', 'company_id'):
            if f in eq_available:
                eq_read.append(f)
        equipments = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': sorted(set(eq_read)), 'limit': 2000, 'order': 'name asc'},
        )

        busy_ids = set()
        try:
            req_fields = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.request', 'fields_get', [], {'attributes': ['type']},
            ) or {}
            req_available = set(req_fields.keys())
            if 'equipment_id' in req_available and 'stage_id' in req_available:
                req_records = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'maintenance.request', 'search_read',
                    [[('stage_id.done', '=', False)]],
                    {'fields': ['equipment_id'], 'limit': 5000},
                )
                for rr in req_records:
                    eq = rr.get('equipment_id')
                    if isinstance(eq, list) and eq:
                        busy_ids.add(eq[0])
        except Exception:
            busy_ids = set()

        for eq in equipments:
            eq_id = eq.get('id')
            is_active = bool(eq.get('active', True))
            is_busy = eq_id in busy_ids
            status = 'Indisponible' if (not is_active or is_busy) else 'Disponible'
            cat = eq.get('category_id')
            comp = eq.get('company_id')
            taux = 100 if status == 'Disponible' else 0
            row = {
                'equipement': eq.get('name') or '—',
                'categorie': cat[1] if isinstance(cat, list) and len(cat) > 1 else '—',
                'localisation': comp[1] if isinstance(comp, list) and len(comp) > 1 else '—',
                'taux': taux,
                'etat': status,
            }
            if q:
                hay = f"{row['equipement']} {row['categorie']} {row['localisation']} {row['etat']}".lower()
                if q not in hay:
                    continue
            rows.append(row)
            if status == 'Disponible':
                kpi_dispo += 1
            else:
                kpi_indispo += 1
    except Exception as exc:
        logger.error('parc_disponibilite_pdf: %s', exc)

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="parc_disponibilite.pdf"'

    doc = SimpleDocTemplate(response, pagesize=landscape(A4), leftMargin=12, rightMargin=12, topMargin=16, bottomMargin=16)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleParc', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor('#1a2c4e'))
    sub_style = ParagraphStyle('SubParc', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#64748b'))
    cell_style = ParagraphStyle('CellParc', parent=styles['Normal'], fontSize=8.5, leading=10)
    cell_num_style = ParagraphStyle('CellNumParc', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=TA_RIGHT)
    cell_center_style = ParagraphStyle('CellCenterParc', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=TA_CENTER)
    total = len(rows)
    kpi_taux = round((kpi_dispo * 100.0) / max(1, total), 1)
    story = [
        Paragraph('Parc & Maintenance - Rapport Disponibilite', title_style),
        Spacer(1, 4),
        Paragraph(f'Genere le {datetime.now().strftime("%d/%m/%Y a %H:%M")}', sub_style),
        Spacer(1, 8),
    ]

    summary_style = ParagraphStyle(
        'SummaryParc',
        parent=styles['Normal'],
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#1f2937')
    )
    summary_text = (
        f"<b>Synthese :</b> Total equipements <b>{total}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Disponibles <b>{kpi_dispo}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Indisponibles <b>{kpi_indispo}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Taux global <b>{kpi_taux}%</b>"
    )
    summary_table = Table([[Paragraph(summary_text, summary_style)]], colWidths=[258 * mm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f8fafc')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#cbd5e1')),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.extend([summary_table, Spacer(1, 10)])

    data = [['Equipement', 'Categorie', 'Localisation', 'Taux dispo (%)', 'Etat']]
    for r in rows:
        data.append([
            Paragraph(str(r['equipement']), cell_style),
            Paragraph(str(r['categorie']), cell_style),
            Paragraph(str(r['localisation']), cell_style),
            Paragraph(str(r['taux']), cell_num_style),
            Paragraph(str(r['etat']), cell_center_style),
        ])

    table = Table(data, colWidths=[88 * mm, 52 * mm, 68 * mm, 24 * mm, 26 * mm], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (3, 0), (3, 0), 'CENTER'),
        ('ALIGN', (4, 0), (4, 0), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#d1d5db')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(table)
    doc.build(story)
    return response


@login_required
def parc_disponibilite_excel(request):
    q = (request.GET.get('q') or '').strip().lower()
    rows = []
    try:
        uid, models = get_odoo_connection()
        eq_fields = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'fields_get', [], {'attributes': ['type']},
        ) or {}
        eq_available = set(eq_fields.keys())
        eq_read = ['name']
        for f in ('active', 'category_id', 'company_id'):
            if f in eq_available:
                eq_read.append(f)
        equipments = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': sorted(set(eq_read)), 'limit': 2000, 'order': 'name asc'},
        )

        busy_ids = set()
        try:
            req_fields = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'maintenance.request', 'fields_get', [], {'attributes': ['type']},
            ) or {}
            req_available = set(req_fields.keys())
            if 'equipment_id' in req_available and 'stage_id' in req_available:
                req_records = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'maintenance.request', 'search_read',
                    [[('stage_id.done', '=', False)]],
                    {'fields': ['equipment_id'], 'limit': 5000},
                )
                for rr in req_records:
                    eq = rr.get('equipment_id')
                    if isinstance(eq, list) and eq:
                        busy_ids.add(eq[0])
        except Exception:
            busy_ids = set()

        for eq in equipments:
            eq_id = eq.get('id')
            is_active = bool(eq.get('active', True))
            is_busy = eq_id in busy_ids
            status = 'Indisponible' if (not is_active or is_busy) else 'Disponible'
            cat = eq.get('category_id')
            comp = eq.get('company_id')
            taux = 100 if status == 'Disponible' else 0
            row = {
                'equipement': eq.get('name') or '—',
                'categorie': cat[1] if isinstance(cat, list) and len(cat) > 1 else '—',
                'localisation': comp[1] if isinstance(comp, list) and len(comp) > 1 else '—',
                'jours_service': 30 if status == 'Disponible' else 0,
                'jours_panne': 0 if status == 'Disponible' else 30,
                'taux': taux,
                'etat': status,
            }
            if q:
                hay = f"{row['equipement']} {row['categorie']} {row['localisation']} {row['etat']}".lower()
                if q not in hay:
                    continue
            rows.append(row)
    except Exception as exc:
        logger.error('parc_disponibilite_excel: %s', exc)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Disponibilite'
    headers = ['Equipement', 'Categorie', 'Localisation', 'Jours service', 'Jours panne', 'Taux dispo (%)', 'Etat']
    ws.append(headers)
    for r in rows:
        ws.append([
            r['equipement'], r['categorie'], r['localisation'],
            r['jours_service'], r['jours_panne'], r['taux'], r['etat']
        ])

    for i, width in enumerate((30, 24, 24, 14, 14, 14, 14), start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1A2C4E', end_color='1A2C4E', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="parc_disponibilite.xlsx"'
    wb.save(response)
    return response


@login_required
def parc_disponibilite_csv(request):
    from django.shortcuts import redirect
    return redirect('parc_disponibilite')


@login_required
def parc_ordres_pdf(request):
    from django.shortcuts import redirect
    return redirect('parc_ordres_maintenance')


@login_required
def parc_ordres_excel(request):
    from django.shortcuts import redirect
    return redirect('parc_ordres_maintenance')


@login_required
def parc_ordres_csv(request):
    from django.shortcuts import redirect
    return redirect('parc_ordres_maintenance')


@login_required
def parc_interventions_pdf(request):
    from django.shortcuts import redirect
    return redirect('parc_interventions')


@login_required
def parc_interventions_excel(request):
    from django.shortcuts import redirect
    return redirect('parc_interventions')


@login_required
def parc_interventions_csv(request):
    from django.shortcuts import redirect
    return redirect('parc_interventions')


@login_required
def parc_couts_pdf(request):
    from django.shortcuts import redirect
    return redirect('parc_couts')


@login_required
def parc_couts_excel(request):
    from django.shortcuts import redirect
    return redirect('parc_couts')


@login_required
def parc_couts_csv(request):
    from django.shortcuts import redirect
    return redirect('parc_couts')


def _rent_pdf_esc(text):
    if text is None:
        return ''
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _rent_trunc_pdf_label(text, max_len=54):
    s = (text or '').strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + '…'


def _rentabilite_pdf_response(
    activity,
    date_debut,
    date_fin,
    company_display,
    perimetre_display,
    currency_suffix,
    pcg,
    total_revenus,
    total_couts,
    resultat,
    marge_pct,
):
    activity_title = 'Transport & logistique' if activity == 'transport' else 'Production'
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title=f'Rentabilité {activity}',
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name='RlRentTitle',
        parent=styles['Heading1'],
        fontSize=15,
        spaceAfter=10,
        textColor=colors.HexColor('#1a2c4e'),
    )
    small = ParagraphStyle(
        name='RlRentSmall',
        parent=styles['Normal'],
        fontSize=8.5,
        textColor=colors.HexColor('#525252'),
        leading=11,
    )
    h3 = ParagraphStyle(
        name='RlRentH3',
        parent=styles['Heading3'],
        fontSize=11,
        textColor=colors.HexColor('#1a2c4e'),
        spaceBefore=6,
        spaceAfter=4,
    )
    small_note = ParagraphStyle(
        name='RlRentSmallNote',
        parent=styles['Normal'],
        fontSize=7.8,
        textColor=colors.HexColor('#64748b'),
        leading=10,
    )

    gen_at = timezone.localtime(timezone.now()).strftime('%d/%m/%Y à %H:%M')
    meta_fr = [
        f'<b>Période :</b> {_rent_pdf_esc(date_debut or "—")} → {_rent_pdf_esc(date_fin or "—")}',
        f'<b>Société :</b> {_rent_pdf_esc(company_display)}',
        f'<b>Périmètre :</b> {_rent_pdf_esc(perimetre_display)}',
        f'<b>Généré le</b> {gen_at} — Synthèse indicative (grand livre Odoo, écritures comptabilisées).',
    ]

    story = [Paragraph(_rent_pdf_esc(f'Rentabilité — {activity_title}'), title_style)]
    for line in meta_fr:
        story.append(Paragraph(line, small))
    story.append(Spacer(1, 10))

    cur_e = _rent_pdf_esc(currency_suffix)
    hdr = ['Indicateur', f'Montant ({cur_e})']

    def _fr_num(x):
        v = float(x or 0)
        s = f'{v:,.2f}'
        s = s.replace(',', ' ').replace('.', ',')
        return s

    kpi_rows = [
        ['Produits / ventes (classe 7)', _fr_num(total_revenus)],
        ['Charges (classe 6)', _fr_num(total_couts)],
        ['Charges fixes (PCG)', _fr_num(pcg.get('charges_fixes', 0))],
        ['Charges variables (PCG)', _fr_num(pcg.get('charges_variables', 0))],
        ['Résultat (7 − 6)', _fr_num(resultat)],
        ['Taux de rentabilité %', _fr_num(marge_pct)],
        ['Marge de contribution', _fr_num(pcg.get('contribution', 0))],
    ]
    seuil = pcg.get('seuil_ca_ht')
    if seuil is not None:
        kpi_rows.append(['Seuil de rentabilité CA HT', _fr_num(seuil)])

    tbl = Table([hdr] + kpi_rows, colWidths=[98 * mm, 62 * mm])
    tbl.setStyle(
        TableStyle(
            [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#cbd5e1')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(tbl)
    story.append(Spacer(1, 14))

    pdf_top6 = max(15, int(getattr(settings, 'RENTABILITE_PDF_TOP_CHARGES', 45)))
    pdf_top7 = max(15, int(getattr(settings, 'RENTABILITE_PDF_TOP_PRODUITS', 35)))
    all6 = list(pcg.get('comptes_6') or [])
    all7 = list(pcg.get('comptes_7') or [])
    show6 = all6[:pdf_top6]
    show7 = all7[:pdf_top7]
    hidden6 = max(0, len(all6) - len(show6))
    hidden7 = max(0, len(all7) - len(show7))

    story.append(Paragraph(_rent_pdf_esc('Charges — classe 6 (principaux comptes)'), h3))
    story.append(Spacer(1, 3))
    rows6 = [['Compte', 'Libellé', f'Montant ({cur_e})']]
    for row in show6:
        rows6.append(
            [
                _rent_pdf_esc(row['code']),
                _rent_pdf_esc(_rent_trunc_pdf_label(row.get('libelle'))),
                _fr_num(row['montant']),
            ]
        )
    rows6.append(['TOTAL classe 6', '', _fr_num(total_couts)])
    t6 = Table(rows6, colWidths=[26 * mm, 92 * mm, 42 * mm])
    t6.setStyle(
        TableStyle(
            [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8fafc')]),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e2e8f0')),
                ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                ('SPAN', (0, -1), (1, -1)),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t6)
    if hidden6:
        story.append(Spacer(1, 3))
        story.append(Paragraph(
            _rent_pdf_esc(
                f'Note: {hidden6} compte(s) de charges supplémentaire(s) non affiché(s) dans ce PDF; '
                'les totaux restent complets.'
            ),
            small_note,
        ))

    story.append(Spacer(1, 12))
    story.append(Paragraph(_rent_pdf_esc('Produits / ventes — classe 7 (principaux comptes)'), h3))
    story.append(Spacer(1, 3))
    rows7 = [['Compte', 'Libellé', f'Montant ({cur_e})']]
    for row in show7:
        rows7.append(
            [
                _rent_pdf_esc(row['code']),
                _rent_pdf_esc(_rent_trunc_pdf_label(row.get('libelle'))),
                _fr_num(row['montant']),
            ]
        )
    rows7.append(['TOTAL classe 7', '', _fr_num(total_revenus)])
    t7 = Table(rows7, colWidths=[26 * mm, 92 * mm, 42 * mm])
    t7.setStyle(
        TableStyle(
            [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8fafc')]),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e2e8f0')),
                ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                ('SPAN', (0, -1), (1, -1)),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t7)
    if hidden7:
        story.append(Spacer(1, 3))
        story.append(Paragraph(
            _rent_pdf_esc(
                f'Note: {hidden7} compte(s) de produits supplémentaire(s) non affiché(s) dans ce PDF; '
                'les totaux restent complets.'
            ),
            small_note,
        ))

    story.append(Spacer(1, 12))
    disclaim = (
        'Document non contractuel : agrégation des comptes dont le numéro commence par 6 ou 7 sur la période '
        'et la société sélectionnées. Ne remplace pas les états comptables officiels. '
        'Ventilation fixes / variables selon préfixes PCG configurés sur le serveur.'
    )
    story.append(Paragraph(_rent_pdf_esc(disclaim), small))

    def _draw_footer(canvas_obj, doc_obj):
        canvas_obj.saveState()
        canvas_obj.setFont('Helvetica', 8)
        canvas_obj.setFillColor(colors.HexColor('#64748b'))
        footer = f'SOMATRIN — Rentabilité {activity_title} — Page {canvas_obj.getPageNumber()}'
        canvas_obj.drawString(doc.leftMargin, 8 * mm, footer)
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    buf.seek(0)
    pdf_bytes = buf.read()
    safe_activity = 'transport' if activity == 'transport' else 'production'
    fname = f'Rentabilite_{safe_activity}_{date_debut or "all"}_{date_fin or "all"}.pdf'
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


def _rentabilite_pdf_executive_response(
    activity,
    date_debut,
    date_fin,
    company_display,
    perimetre_display,
    currency_suffix,
    pcg,
    total_revenus,
    total_couts,
    resultat,
    marge_pct,
):
    """Version Executive (condensée) : 1 page orientée direction."""
    activity_title = 'Transport & logistique' if activity == 'transport' else 'Production'
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=f'Rentabilité Executive {activity}',
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        name='ExecTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#0f1e38'),
        spaceAfter=8,
    )
    sub = ParagraphStyle(
        name='ExecSub',
        parent=styles['Normal'],
        fontSize=8.5,
        textColor=colors.HexColor('#64748b'),
        leading=11,
    )
    body = ParagraphStyle(
        name='ExecBody',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#334155'),
        leading=12,
    )
    h = ParagraphStyle(
        name='ExecH',
        parent=styles['Heading3'],
        fontSize=10.5,
        textColor=colors.HexColor('#1a2c4e'),
        spaceBefore=4,
        spaceAfter=4,
    )

    def _fr_num(x):
        v = float(x or 0)
        s = f'{v:,.2f}'.replace(',', ' ').replace('.', ',')
        return s

    contribution = float(pcg.get('contribution') or 0)
    seuil = pcg.get('seuil_ca_ht')
    cf = float(pcg.get('charges_fixes') or 0)
    cv = float(pcg.get('charges_variables') or 0)
    cf_pct = round((cf / max(total_couts, 0.01)) * 100, 1)
    cv_pct = round((cv / max(total_couts, 0.01)) * 100, 1)

    exec_comment = (
        'Rentabilité forte, création de valeur nette.'
        if marge_pct >= 20 else
        'Rentabilité positive, marge à consolider.'
        if marge_pct >= 8 else
        'Rentabilité faible, pilotage des charges recommandé.'
        if marge_pct >= 0 else
        'Déficit constaté, plan de redressement prioritaire.'
    )
    risk_color = (
        colors.HexColor('#16a34a') if marge_pct >= 20 else
        colors.HexColor('#2563eb') if marge_pct >= 8 else
        colors.HexColor('#d97706') if marge_pct >= 0 else
        colors.HexColor('#dc2626')
    )

    story = [
        Paragraph(_rent_pdf_esc(f'Rentabilité Executive — {activity_title}'), title),
        Paragraph(
            _rent_pdf_esc(
                f'Période: {date_debut or "—"} → {date_fin or "—"} | Société: {company_display} | '
                f'Périmètre: {perimetre_display} | Généré le {timezone.localtime(timezone.now()).strftime("%d/%m/%Y %H:%M")}'
            ),
            sub,
        ),
        Spacer(1, 7),
    ]

    kpi = [
        ['Produits (classe 7)', _fr_num(total_revenus)],
        ['Charges (classe 6)', _fr_num(total_couts)],
        ['Résultat (7 − 6)', _fr_num(resultat)],
        ['Taux rentabilité %', _fr_num(marge_pct)],
        ['Marge de contribution', _fr_num(contribution)],
    ]
    if seuil is not None:
        kpi.append(['Seuil de rentabilité', _fr_num(seuil)])
    kpi_tbl = Table([['KPI clés', f'Montant ({_rent_pdf_esc(currency_suffix)})']] + kpi, colWidths=[96 * mm, 74 * mm])
    kpi_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0f1e38')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#cbd5e1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 8))

    story.append(Paragraph(_rent_pdf_esc('Lecture managériale'), h))
    text = (
        f'<b>Diagnostic:</b> <font color="{risk_color.hexval()}">{_rent_pdf_esc(exec_comment)}</font><br/>'
        f'<b>Structure de charges:</b> fixes {cf_pct}% | variables {cv_pct}%.<br/>'
        f'<b>Point mort:</b> '
        + (_rent_pdf_esc(f'{_fr_num(seuil)} {currency_suffix} (calculable)') if seuil is not None else 'non déterminable sur la période.')
    )
    story.append(Paragraph(text, body))
    story.append(Spacer(1, 8))

    top6 = (pcg.get('comptes_6') or [])[:10]
    top7 = (pcg.get('comptes_7') or [])[:10]
    story.append(Paragraph(_rent_pdf_esc('Top 10 comptes charges (classe 6)'), h))
    rows6 = [['Compte', 'Libellé', f'Montant ({_rent_pdf_esc(currency_suffix)})']]
    for row in top6:
        rows6.append([_rent_pdf_esc(row.get('code')), _rent_pdf_esc(_rent_trunc_pdf_label(row.get('libelle'))), _fr_num(row.get('montant'))])
    t6 = Table(rows6, colWidths=[28 * mm, 96 * mm, 46 * mm])
    t6.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    story.append(t6)
    story.append(Spacer(1, 8))

    story.append(Paragraph(_rent_pdf_esc('Top 10 comptes produits (classe 7)'), h))
    rows7 = [['Compte', 'Libellé', f'Montant ({_rent_pdf_esc(currency_suffix)})']]
    for row in top7:
        rows7.append([_rent_pdf_esc(row.get('code')), _rent_pdf_esc(_rent_trunc_pdf_label(row.get('libelle'))), _fr_num(row.get('montant'))])
    t7 = Table(rows7, colWidths=[28 * mm, 96 * mm, 46 * mm])
    t7.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    story.append(t7)
    story.append(Spacer(1, 8))
    story.append(Paragraph(
        _rent_pdf_esc('Document de pilotage interne. Référer aux états comptables officiels pour les usages réglementaires.'),
        sub,
    ))

    doc.build(story)
    buf.seek(0)
    resp = HttpResponse(buf.read(), content_type='application/pdf')
    safe_activity = 'transport' if activity == 'transport' else 'production'
    exec_fname = f'Rentabilite_Executive_{safe_activity}_{date_debut or "all"}_{date_fin or "all"}.pdf'
    resp['Content-Disposition'] = f'attachment; filename="{exec_fname}"'
    return resp


def _render_rentabilite_dashboard(request, activity, template_path):
    """
    Rentabilité à partir du grand livre Odoo (comptes 6 = charges, 7 = produits / ventes),
    avec ventilation fixes / variables (préfixes PCG configurables) et indicateurs de seuil.
    """
    date_debut = request.GET.get('date_debut', '').strip()
    date_fin = request.GET.get('date_fin', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    perimetre = request.GET.get('perimetre', 'complet').strip().lower()
    if perimetre not in ('activite', 'complet'):
        perimetre = 'complet'
    export = request.GET.get('export', '').strip().lower()

    cid = int(company_id) if company_id.isdigit() else None
    companies = []
    pcg = {}
    error = None
    currency_suffix = getattr(settings, 'RENTABILITE_AMOUNT_SUFFIX', None) or 'MAD'

    try:
        uid, models = get_odoo_connection()

        try:
            comp_recs = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'res.company', 'search_read',
                [[]],
                {'fields': ['id', 'name'], 'order': 'name', 'limit': 80},
            )
            companies = [{'id': c['id'], 'name': c.get('name') or '—'} for c in comp_recs]
        except Exception:
            companies = []

        # ── Cache + fast path ──────────────────────────────────────────────
        _cache_ttl = int(getattr(settings, 'RENTABILITE_CACHE_SECONDS', 300))
        _ck = _rent_cache_key(activity, perimetre, date_debut, date_fin, cid)
        pcg = cache.get(_ck)
        if pcg is None:
            import time as _time
            _t0 = _time.monotonic()
            if perimetre == 'complet':
                pcg = _rentabilite_pcg_readgroup(uid, models, date_debut, date_fin, cid)
            if not pcg:
                pcg = _rentabilite_pcg_aggregate(uid, models, date_debut, date_fin, cid, activity, perimetre)
            _elapsed = round(_time.monotonic() - _t0, 2)
            logger.info(
                'rentabilite %s/%s: calculé en %ss (fast=%s) — mis en cache %ss',
                activity, perimetre, _elapsed,
                pcg.get('_fast_path', False) if pcg else 'N/A',
                _cache_ttl,
            )
            if pcg:
                cache.set(_ck, pcg, _cache_ttl)
        else:
            logger.info('rentabilite %s/%s: servi depuis le cache', activity, perimetre)
        # ──────────────────────────────────────────────────────────────────

        if cid:
            try:
                cread = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'res.company', 'read', [[cid]], {'fields': ['currency_id']},
                )
                if cread and cread[0].get('currency_id'):
                    cur_id = cread[0]['currency_id'][0]
                    cur_rec = models.execute_kw(
                        settings.ODOO_DB, uid, settings.ODOO_PASS,
                        'res.currency', 'read', [[cur_id]], {'fields': ['symbol']},
                    )
                    if cur_rec and (cur_rec[0].get('symbol') or '').strip():
                        currency_suffix = cur_rec[0]['symbol'].strip()
            except Exception:
                pass
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'
        pcg = {
            'total_produits_7': 0.0,
            'total_charges_6': 0.0,
            'charges_fixes': 0.0,
            'charges_variables': 0.0,
            'resultat': 0.0,
            'taux_rentabilite': 0.0,
            'contribution': 0.0,
            'taux_marge_contribution': 0.0,
            'marge_sur_cv_pct': None,
            'seuil_ca_ht': None,
            'nb_lignes_gl': 0,
            'nb_lignes_apres_activite': 0,
            'comptes_6': [],
            'comptes_7': [],
            'comptes_6_hors_top': 0,
            'comptes_7_hors_top': 0,
            'top_comptes_charges_limite': max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_CHARGES', 120))),
            'top_comptes_produits_limite': max(10, int(getattr(settings, 'RENTABILITE_TOP_COMPTES_PRODUITS', 60))),
        }

    total_revenus = pcg.get('total_produits_7', 0.0)
    total_couts = pcg.get('total_charges_6', 0.0)
    resultat = pcg.get('resultat', 0.0)
    marge_pct = pcg.get('taux_rentabilite', 0.0)
    reset_url = '/transport/rentabilite/' if activity == 'transport' else '/production/rentabilite/'

    company_display = 'Toutes les sociétés (agrégat)'
    if cid:
        company_display = next((c['name'] for c in companies if c['id'] == cid), f'Société #{cid}')

    if activity == 'transport':
        perimetre_display = (
            'Complet (classes 6 / 7)' if perimetre == 'complet' else 'Activité transport'
        )
    else:
        perimetre_display = (
            'Complet (classes 6 / 7)' if perimetre == 'complet' else 'Activité production'
        )

    report_generated_at = timezone.now()

    if export == 'pdf_exec' and not error:
        try:
            return _rentabilite_pdf_executive_response(
                activity,
                date_debut,
                date_fin,
                company_display,
                perimetre_display,
                currency_suffix,
                pcg,
                total_revenus,
                total_couts,
                resultat,
                marge_pct,
            )
        except Exception as exc:
            error = f'Export PDF Executive indisponible : {exc}'

    if export == 'pdf' and not error:
        try:
            return _rentabilite_pdf_response(
                activity,
                date_debut,
                date_fin,
                company_display,
                perimetre_display,
                currency_suffix,
                pcg,
                total_revenus,
                total_couts,
                resultat,
                marge_pct,
            )
        except Exception as exc:
            error = f'Export PDF indisponible : {exc}'

    if export == 'csv' and not error:
        response = HttpResponse(content_type='text/csv; charset=utf-8-sig')
        fname = f'rentabilite_{activity}_{date_debut or "all"}_{date_fin or "all"}.csv'
        response['Content-Disposition'] = f'attachment; filename="{fname}"'
        w = csv.writer(response, delimiter=';')
        w.writerow(['Rapport', 'Rentabilité — synthèse grand livre Odoo'])
        w.writerow(['Généré le', timezone.localtime(report_generated_at).strftime('%d/%m/%Y %H:%M')])
        w.writerow(['Société', company_display])
        w.writerow(['Période début', date_debut or '—'])
        w.writerow(['Période fin', date_fin or '—'])
        w.writerow(['Périmètre', perimetre_display])
        w.writerow(['Activité', activity])
        w.writerow([])
        w.writerow(['Indicateur', f'Montant ({currency_suffix})'])
        w.writerow(['Produits / ventes (classe 7)', total_revenus])
        w.writerow(['Charges (classe 6)', total_couts])
        w.writerow(['Charges fixes (préfixes PCG)', pcg.get('charges_fixes', 0)])
        w.writerow(['Charges variables (préfixes PCG)', pcg.get('charges_variables', 0)])
        w.writerow(['Résultat (7 − 6)', resultat])
        w.writerow(['Taux rentabilité %', marge_pct])
        w.writerow(['Contribution (7 − var.)', pcg.get('contribution', 0)])
        if pcg.get('seuil_ca_ht') is not None:
            w.writerow(['Seuil de rentabilité CA HT', pcg['seuil_ca_ht']])
        w.writerow([])
        w.writerow(['Compte 6', 'Libellé', f'Montant ({currency_suffix})'])
        for row in pcg.get('comptes_6') or []:
            w.writerow([row['code'], row.get('libelle') or '', row['montant']])
        w.writerow([])
        w.writerow(['Compte 7', 'Libellé', f'Montant ({currency_suffix})'])
        for row in pcg.get('comptes_7') or []:
            w.writerow([row['code'], row.get('libelle') or '', row['montant']])
        return response

    # ── Métriques dérivées pour le dashboard spectaculaire ────────────────
    _cf = float(pcg.get('charges_fixes') or 0)
    _cv = float(pcg.get('charges_variables') or 0)
    _safe_couts = max(float(total_couts), 0.01)
    _safe_rev   = max(float(total_revenus), 0.01)
    charges_fixes_pct     = round(_cf / _safe_couts * 100, 1)
    charges_variables_pct = round(_cv / _safe_couts * 100, 1)
    revenus_restant_pct   = round(float(resultat) / _safe_rev * 100, 1) if total_revenus > 0 else 0.0

    if marge_pct >= 20:
        perf_badge, perf_color, perf_bg = 'excellent', '#16a34a', '#dcfce7'
        perf_label = 'Excellente rentabilité'
    elif marge_pct >= 10:
        perf_badge, perf_color, perf_bg = 'bon', '#2563eb', '#dbeafe'
        perf_label = 'Bonne rentabilité'
    elif marge_pct >= 0:
        perf_badge, perf_color, perf_bg = 'attention', '#d97706', '#fef3c7'
        perf_label = 'Rentabilité à améliorer'
    else:
        perf_badge, perf_color, perf_bg = 'deficitaire', '#dc2626', '#fee2e2'
        perf_label = 'Résultat déficitaire'

    gauge_pct = max(0, min(100, float(marge_pct))) if marge_pct >= 0 else 0
    nb_comptes_charges  = len(pcg.get('comptes_6') or [])
    nb_comptes_produits = len(pcg.get('comptes_7') or [])
    # Jeux de données Chart.js (alimentés par les données Odoo agrégées)
    top_charges = (pcg.get('comptes_6') or [])[:8]
    top_produits = (pcg.get('comptes_7') or [])[:8]
    chart_bar_labels = [str(x.get('code') or '—') for x in top_charges]
    chart_bar_values = [round(float(x.get('montant') or 0), 2) for x in top_charges]
    chart_line_labels = [str(x.get('code') or '—') for x in top_produits]
    chart_line_values = [round(float(x.get('montant') or 0), 2) for x in top_produits]
    non_ventile = max(0.0, float(total_couts) - float(_cf) - float(_cv))
    chart_pie_labels = ['Charges fixes', 'Charges variables']
    chart_pie_values = [round(float(_cf), 2), round(float(_cv), 2)]
    if non_ventile > 0:
        chart_pie_labels.append('Non ventilées')
        chart_pie_values.append(round(non_ventile, 2))
    # ─────────────────────────────────────────────────────────────────────

    return render(request, template_path, {
        'error': error,
        'activity': activity,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'company_id': company_id,
        'companies': companies,
        'perimetre': perimetre,
        'company_display': company_display,
        'perimetre_display': perimetre_display,
        'currency_suffix': currency_suffix,
        'report_generated_at': report_generated_at,
        'total_revenus': total_revenus,
        'total_couts': total_couts,
        'resultat': resultat,
        'marge_pct': marge_pct,
        'reset_url': reset_url,
        'pcg': pcg,
        'charges_fixes': pcg.get('charges_fixes', 0),
        'charges_variables': pcg.get('charges_variables', 0),
        'contribution': pcg.get('contribution', 0),
        'taux_marge_contribution': pcg.get('taux_marge_contribution', 0),
        'marge_sur_cv_pct': pcg.get('marge_sur_cv_pct'),
        'seuil_ca_ht': pcg.get('seuil_ca_ht'),
        'comptes_6': pcg.get('comptes_6') or [],
        'comptes_7': pcg.get('comptes_7') or [],
        # métriques dérivées — dashboard spectaculaire
        'charges_fixes_pct': charges_fixes_pct,
        'charges_variables_pct': charges_variables_pct,
        'revenus_restant_pct': revenus_restant_pct,
        'perf_badge': perf_badge,
        'perf_color': perf_color,
        'perf_bg': perf_bg,
        'perf_label': perf_label,
        'gauge_pct': gauge_pct,
        'nb_comptes_charges': nb_comptes_charges,
        'nb_comptes_produits': nb_comptes_produits,
        'chart_bar_labels_json': json.dumps(chart_bar_labels, ensure_ascii=False),
        'chart_bar_values_json': json.dumps(chart_bar_values),
        'chart_pie_labels_json': json.dumps(chart_pie_labels, ensure_ascii=False),
        'chart_pie_values_json': json.dumps(chart_pie_values),
        'chart_line_labels_json': json.dumps(chart_line_labels, ensure_ascii=False),
        'chart_line_values_json': json.dumps(chart_line_values),
        'soma_page_data': {
            'page': f'{activity}_rentabilite',
            'kpis': {
                'revenus': float(total_revenus),
                'charges': float(total_couts),
                'resultat': float(resultat),
                'marge_pct': float(marge_pct),
                'charges_fixes': float(pcg.get('charges_fixes') or 0),
                'charges_variables': float(pcg.get('charges_variables') or 0),
                'contribution': float(pcg.get('contribution') or 0),
                'seuil_ca_ht': float(pcg.get('seuil_ca_ht') or 0),
                'perf_badge': perf_badge,
                'perf_label': perf_label,
                'currency': currency_suffix,
                'periode_debut': date_debut or None,
                'periode_fin': date_fin or None,
                'nb_comptes_charges': nb_comptes_charges,
                'nb_comptes_produits': nb_comptes_produits,
            },
        },
    })


@login_required
def transport_rentabilite(request):
    # Le transport utilise le dashboard "spectaculaire" avec enrichissement data-viz.
    return _render_rentabilite_dashboard(request, 'transport', 'transport/rentabilite.html')


@login_required
def production_rentabilite(request):
    return _render_rentabilite_dashboard(request, 'production', 'production/rentabilite.html')


# ─────────────────────────────────────────────
#  QHSE — ALERTES QUALITÉ
#  Modèle : quality.alert
# ─────────────────────────────────────────────

def _fetch_qhse_sites(uid, models):
    """Récupère la liste des sites depuis stock.location."""
    sites = []
    try:
        locs = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.location', 'search_read',
            [[['usage', 'in', ['internal', 'customer', 'supplier']]]],
            {'fields': ['complete_name'], 'order': 'complete_name asc', 'limit': 300}
        )
        sites = [loc['complete_name'] for loc in locs if loc.get('complete_name')]
    except Exception:
        pass
    return sites


def _build_qhse_alert_domain(site=None):
    """Construit le domaine pour la recherche d'alertes QHSE."""
    domain = [('state', '!=', 'cancel')]
    if site:
        domain.append(('location_id.complete_name', 'ilike', site))
    return domain


def _fetch_qhse_alerts(uid, models, site=None, limit=200):
    """Récupère les événements QHSE (Odoo 16 : quality.check via QHSEService)."""
    try:
        from reporting.services.qhse_service import QHSEService
        svc = QHSEService()
        data = svc.get_accidents(site_filter=site or '', limit=limit)
        alerts = []
        for r in (data.get('rows') or [])[:limit]:
            alerts.append({
                'name': r.get('name') or '—',
                'date': r.get('date'),
                'state': r.get('state'),
                'location_id': [None, r.get('site') or '—'],
                'team_id': False,
                'category_id': [None, r.get('categorie') or '—'],
                'user_id': [None, r.get('responsable') or '—'],
                'company_id': False,
            })
        return alerts
    except Exception:
        return []


def _summarize_qhse_alerts(alerts):
    """Résume les alertes QHSE par état."""
    summary = {'total': 0, 'open': 0, 'in_progress': 0, 'done': 0}
    for alert in alerts:
        state = (alert.get('state') or '').lower()
        summary['total'] += 1
        if state in ('new', 'draft', 'proposed'):
            summary['open'] += 1
        elif state in ('in_progress', 'progress', 'doing'):
            summary['in_progress'] += 1
        elif state in ('done', 'closed', 'solved'):
            summary['done'] += 1
    return summary


def _load_qhse_context(request):
    """Charge le contexte pour les pages QHSE."""
    site = request.GET.get('site', '')
    error = None
    alerts = []
    sites = []
    summary = {'total': 0, 'open': 0, 'in_progress': 0, 'done': 0}
    try:
        uid, models = get_odoo_connection()
        sites = _fetch_qhse_sites(uid, models)
        alerts = _fetch_qhse_alerts(uid, models, site=site)
        summary = _summarize_qhse_alerts(alerts)
    except Exception as exc:
        error = str(exc)
    return {
        'error': error,
        'site': site,
        'sites': sites,
        'alerts': alerts,
        'summary': summary,
    }


@login_required
def qhse_bilan(request):
    """Bilan des alertes QHSE."""
    context = _load_qhse_context(request)
    context.update({
        'page_title': 'Bilan QHSE',
        'page_subtitle': 'Indicateurs QHSE récupérés depuis Odoo.',
    })
    return render(request, 'qhse/bilan.html', context)


@login_required
def qhse_entrees(request):
    """Entrées QHSE."""
    context = _load_qhse_context(request)
    context.update({
        'page_title': 'Entrées QHSE',
        'page_subtitle': 'Données de conformité et incidents entrés depuis Odoo.',
    })
    return render(request, 'qhse/entrees.html', context)


@login_required
def qhse_sorties(request):
    """Sorties QHSE."""
    context = _load_qhse_context(request)
    context.update({
        'page_title': 'Sorties QHSE',
        'page_subtitle': 'Rapports de sortie et actions QHSE provenant d\'Odoo.',
    })
    return render(request, 'qhse/sorties.html', context)


def _render_qhse_service_page(request, page_title, page_subtitle):
    """Rendu standard des sous-menus service QHSE."""
    context = _load_qhse_context(request)
    context.update({
        'page_title': page_title,
        'page_subtitle': page_subtitle,
    })
    return render(request, 'qhse/bilan.html', context)


def _qhse_keywords():
    return [
        'qhse', 'hse', 'epi', 'e.p.i', 'equipement protection individuelle',
        'equipements de protection individuelle', 'protection individuelle',
        'securite', 'secur', 'qualite', 'hygiene', 'environnement',
        'casque', 'gant', 'gants', 'lunette', 'lunettes', 'masque',
        'chaussure securite', 'chaussures securite', 'gilet',
    ]


def _match_qhse_text(*parts):
    def _norm(text):
        text = unicodedata.normalize('NFKD', str(text or ''))
        text = ''.join(ch for ch in text if not unicodedata.combining(ch))
        text = text.lower()
        text = re.sub(r'[^a-z0-9]+', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()

    raw = _norm(' '.join(str(p or '') for p in parts if p is not None))
    return any(_norm(k) in raw for k in _qhse_keywords())


def _format_date_fr(raw_value):
    raw = str(raw_value or '')[:10]
    if len(raw) == 10 and raw[4] == '-' and raw[7] == '-':
        return f"{raw[8:10]}/{raw[5:7]}/{raw[0:4]}"
    return '—'


def _fetch_qhse_category_ids(uid, models):
    """
    Détermine les catégories QHSE dans Odoo et leurs sous-catégories.
    Règle principale: catégories dont le nom/chemin contient QHSE/HSE/etc.
    """
    categories = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'product.category', 'search_read',
        [[]],
        {'fields': ['id', 'name', 'complete_name', 'parent_path'], 'limit': 5000},
    ) or []

    root_ids = set()
    for cat in categories:
        if _match_qhse_text(cat.get('name'), cat.get('complete_name')):
            root_ids.add(int(cat['id']))

    if not root_ids:
        return set()

    # Inclure descendants via parent_path: "1/4/23/" contient les ancêtres.
    all_ids = set(root_ids)
    for cat in categories:
        parent_path = str(cat.get('parent_path') or '')
        for rid in root_ids:
            token = f'/{rid}/'
            if parent_path.startswith(f'{rid}/') or token in f'/{parent_path}':
                all_ids.add(int(cat['id']))
                break
    return all_ids


def _fetch_qhse_purchase_orders(status_filter='all', supplier_filter=''):
    rows = []
    suppliers = set()
    error = None
    try:
        uid, models = get_odoo_connection()
        qhse_category_ids = _fetch_qhse_category_ids(uid, models)

        line_domain = [('order_id.state', 'in', ['draft', 'sent', 'to approve', 'purchase', 'done'])]
        if status_filter != 'all':
            line_domain.append(('order_id.state', '=', status_filter))

        lines = []
        if qhse_category_ids:
            # Règle métier prioritaire: achat dont la ligne appartient à une catégorie QHSE.
            line_domain_cat = list(line_domain)
            line_domain_cat.append(('product_id.categ_id', 'in', list(qhse_category_ids)))
            lines = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order.line', 'search_read',
                [line_domain_cat],
                {
                    'fields': ['order_id', 'name', 'product_id', 'product_qty', 'price_total'],
                    'order': 'id desc',
                    'limit': 1200,
                },
            ) or []

        if not lines:
            # Fallback sécurité si la structure catégorie QHSE n'existe pas clairement.
            lines = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'purchase.order.line', 'search_read',
                [line_domain],
                {
                    'fields': ['order_id', 'name', 'product_id', 'product_qty', 'price_total'],
                    'order': 'id desc',
                    'limit': 1200,
                },
            ) or []
            lines = [
                ln for ln in lines
                if _match_qhse_text(
                    (ln.get('product_id') or ['', ''])[1] if isinstance(ln.get('product_id'), list) else '',
                    ln.get('name'),
                )
            ]

        # Grouper les lignes QHSE par commande.
        lines_by_order = defaultdict(list)
        order_ids = set()
        for ln in lines:
            order_ref = ln.get('order_id')
            if isinstance(order_ref, list) and order_ref:
                oid = int(order_ref[0])
                lines_by_order[oid].append(ln)
                order_ids.add(oid)

        if not order_ids:
            return [], [], None

        orders = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'purchase.order', 'search_read',
            [[('id', 'in', list(order_ids))]],
            {
                'fields': ['name', 'partner_id', 'user_id', 'company_id', 'state', 'date_order', 'amount_total', 'currency_id', 'order_line'],
                'order': 'date_order desc, id desc',
                'limit': max(220, len(order_ids) + 20),
            },
        ) or []
        for order in orders:
            supplier = order.get('partner_id')
            supplier_name = supplier[1] if isinstance(supplier, list) and len(supplier) > 1 else '—'
            buyer = order.get('user_id')
            buyer_name = buyer[1] if isinstance(buyer, list) and len(buyer) > 1 else '—'
            company = order.get('company_id')
            company_name = company[1] if isinstance(company, list) and len(company) > 1 else '—'
            state = str(order.get('state') or '').strip().lower()
            if supplier_filter and supplier_filter.lower() not in supplier_name.lower():
                continue

            order_lines = lines_by_order.get(int(order.get('id') or 0), [])
            if not order_lines:
                continue

            suppliers.add(supplier_name)
            rows.append({
                'reference': order.get('name') or '—',
                'fournisseur': supplier_name,
                'acheteur': buyer_name,
                'societe': company_name,
                'date_commande': _format_date_fr(order.get('date_order')),
                'etat': state or '—',
                'montant_total': float(order.get('amount_total') or 0),
                'devise': (order.get('currency_id') or ['', 'MAD'])[1] if isinstance(order.get('currency_id'), list) and len(order.get('currency_id')) > 1 else 'MAD',
                'nb_lignes_qhse': len(order_lines),
                'lignes': order_lines,
                'montant_lignes_qhse': sum(float(ln.get('price_total') or 0) for ln in order_lines),
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'

    rows = sorted(rows, key=lambda x: x.get('date_commande') or '', reverse=True)
    return rows, sorted(s for s in suppliers if s and s != '—'), error


def _fetch_qhse_products(status_filter='all', search_term=''):
    rows = []
    error = None
    try:
        uid, models = get_odoo_connection()
        qhse_category_ids = _fetch_qhse_category_ids(uid, models)
        products = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'product.product', 'search_read',
            [[('active', '=', True)]],
            {
                'fields': ['default_code', 'name', 'categ_id', 'qty_available', 'uom_id', 'purchase_ok', 'sale_ok', 'active'],
                'order': 'name asc',
                'limit': 700,
            },
        ) or []

        for p in products:
            code = p.get('default_code') or ''
            name = p.get('name') or ''
            categ = p.get('categ_id')
            categ_id = categ[0] if isinstance(categ, list) and categ else None
            category_name = categ[1] if isinstance(categ, list) and len(categ) > 1 else ''
            is_qhse_category = bool(categ_id and categ_id in qhse_category_ids)
            is_qhse_text = _match_qhse_text(code, name, category_name)
            if not (is_qhse_category or is_qhse_text):
                continue
            if search_term and search_term.lower() not in f'{code} {name} {category_name}'.lower():
                continue
            if status_filter == 'purchase' and not p.get('purchase_ok'):
                continue
            if status_filter == 'stock' and float(p.get('qty_available') or 0) <= 0:
                continue
            uom = p.get('uom_id')
            rows.append({
                'code': code or '—',
                'designation': name or '—',
                'categorie': category_name or '—',
                'stock': float(p.get('qty_available') or 0),
                'uom': uom[1] if isinstance(uom, list) and len(uom) > 1 else '—',
                'achat': bool(p.get('purchase_ok')),
                'vente': bool(p.get('sale_ok')),
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return rows, error


def _safe_model_fields(uid, models, model, candidates):
    try:
        fget = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            model, 'fields_get', [], {'attributes': ['type']}
        ) or {}
        return [f for f in candidates if f in fget]
    except Exception:
        return []


def _parse_date_any(raw):
    s = str(raw or '').strip()[:10]
    if len(s) != 10:
        return None
    try:
        return datetime.strptime(s, '%Y-%m-%d').date()
    except Exception:
        return None


def _fetch_epi_movements(site_filter='', person_filter=''):
    rows = []
    kpis = {'total_sorties': 0, 'total_transferts': 0, 'quantite_totale': 0.0, 'personnes': 0}
    sites = set()
    persons = set()
    error = None
    try:
        uid, models = get_odoo_connection()
        qhse_category_ids = _fetch_qhse_category_ids(uid, models)

        domain = [('qty_done', '>', 0)]
        if qhse_category_ids:
            domain.append(('product_id.categ_id', 'in', list(qhse_category_ids)))

        line_fields = ['id', 'date', 'qty_done', 'reference', 'product_id', 'location_id', 'location_dest_id', 'picking_id']
        move_lines = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.move.line', 'search_read',
            [domain],
            {'fields': line_fields, 'order': 'date desc,id desc', 'limit': 1800},
        ) or []

        if not move_lines:
            return rows, sorted(sites), sorted(persons), kpis, None

        picking_ids = sorted({
            int((ml.get('picking_id') or [0])[0])
            for ml in move_lines if isinstance(ml.get('picking_id'), list) and ml.get('picking_id')
        })
        pickings_map = {}
        if picking_ids:
            picking_candidates = [
                'name', 'origin', 'partner_id', 'employee_id', 'user_id', 'driver_id',
                'site_id', 'location_id', 'location_dest_id', 'service_car', 'equipment_id',
                'vehicle_id', 'scheduled_date', 'date_done',
                'x_tl', 'x_emplacement', 'x_zone', 'x_voiture_service',
                'x_matricule_camion', 'x_tracteur', 'x_cabine'
            ]
            picking_fields = _safe_model_fields(uid, models, 'stock.picking', picking_candidates)
            if 'name' not in picking_fields:
                picking_fields.append('name')
            picks = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'stock.picking', 'read',
                [picking_ids, picking_fields],
            ) or []
            pickings_map = {int(p.get('id')): p for p in picks if p.get('id')}

        for ml in move_lines:
            product = ml.get('product_id')
            product_name = product[1] if isinstance(product, list) and len(product) > 1 else '—'
            qty = float(ml.get('qty_done') or 0)
            if not qhse_category_ids and not _match_qhse_text(product_name, ml.get('reference')):
                continue

            pick_ref = ml.get('picking_id')
            pick = pickings_map.get(int(pick_ref[0])) if isinstance(pick_ref, list) and pick_ref else {}

            person = '—'
            for fld in ('employee_id', 'driver_id', 'user_id', 'partner_id'):
                v = pick.get(fld)
                if isinstance(v, list) and len(v) > 1 and v[1]:
                    person = v[1]
                    break
            src = (ml.get('location_id') or ['', '—'])[1] if isinstance(ml.get('location_id'), list) else '—'
            dst = (ml.get('location_dest_id') or ['', '—'])[1] if isinstance(ml.get('location_dest_id'), list) else '—'
            site_name = '—'
            if isinstance(pick.get('site_id'), list) and len(pick.get('site_id')) > 1:
                site_name = pick['site_id'][1]
            elif dst != '—':
                site_name = dst

            movement_type = 'Transfert site' if src != '—' and dst != '—' and src != dst else 'Sortie'
            row = {
                'date': _format_date_fr(ml.get('date') or pick.get('date_done') or pick.get('scheduled_date')),
                'document': pick.get('name') or ml.get('reference') or '—',
                'produit': product_name,
                'quantite': qty,
                'personne': person,
                'site': site_name,
                'type_mouvement': movement_type,
                'tl': pick.get('x_tl') or (pick.get('origin') or '—'),
                'emplacement': pick.get('x_emplacement') or pick.get('x_zone') or dst,
                'voiture_service': pick.get('x_voiture_service') or ((pick.get('vehicle_id') or ['', '—'])[1] if isinstance(pick.get('vehicle_id'), list) else '—'),
                'matricule_camion': pick.get('x_matricule_camion') or ((pick.get('equipment_id') or ['', '—'])[1] if isinstance(pick.get('equipment_id'), list) else '—'),
                'tracteur': pick.get('x_tracteur') or '—',
                'cabine': pick.get('x_cabine') or '—',
                'source': src,
                'destination': dst,
            }

            if site_filter and site_filter.lower() not in str(row['site']).lower():
                continue
            if person_filter and person_filter.lower() not in str(row['personne']).lower():
                continue

            rows.append(row)
            sites.add(str(row['site']))
            persons.add(str(row['personne']))
            kpis['quantite_totale'] += qty
            if row['type_mouvement'] == 'Transfert site':
                kpis['total_transferts'] += 1
            else:
                kpis['total_sorties'] += 1

        kpis['personnes'] = len({r['personne'] for r in rows if r.get('personne') and r['personne'] != '—'})
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'

    return rows, sorted(s for s in sites if s and s != '—'), sorted(p for p in persons if p and p != '—'), kpis, error


def _fetch_incidents_dashboard(site_filter=''):
    """Agrégats incidents — Odoo 16 quality.check via QHSEService."""
    try:
        from reporting.services.qhse_service import QHSEService
        svc = QHSEService()
        data = svc.get_accidents(site_filter=site_filter, limit=2500)
        rows = []
        for r in data.get('rows') or []:
            site = r.get('site') or 'Non renseigne'
            rows.append({
                'name': r.get('name'),
                'date': r.get('date'),
                'state': r.get('state'),
                'location_id': [None, site],
                'user_id': [None, r.get('responsable')],
                'category_id': [None, r.get('categorie')],
                'team_id': False,
            })
        total = len(rows)
        open_count = sum(
            1 for r in rows
            if str(r.get('state') or '').lower() not in ('done', 'closed', 'solved', 'pass')
        )
        closed_count = max(0, total - open_count)
        return {
            'rows': rows,
            'kpi_total': data.get('kpi_total', total),
            'kpi_open': open_count,
            'kpi_closed': closed_count,
            'kpi_days_lost': data.get('kpi_jours_arret', total * 2),
            'monthly': data.get('monthly', []),
            'top_employees': data.get('top_employees', []),
            'types': data.get('types', []),
            'sites': data.get('sites', []),
            'error': data.get('error'),
        }
    except Exception as exc:
        return {
            'rows': [],
            'kpi_total': 0,
            'kpi_open': 0,
            'kpi_closed': 0,
            'kpi_days_lost': 0,
            'monthly': [],
            'top_employees': [],
            'types': [],
            'sites': [],
            'error': str(exc),
        }


def _maintenance_equipment_is_extincteur(e, name, cat, typ, extra_category_ids):
    """
    Repère un extincteur côté Odoo (libellés TF/TG, feu, champ x_type_extincteur, etc.).
    L'ancien filtre exigeait « extinct » ou un mot-clé QHSE générique — trop strict pour les fiches TF/TG.
    """
    extra_category_ids = extra_category_ids or []
    raw_cat = e.get('category_id')
    cat_id = None
    if isinstance(raw_cat, (list, tuple)) and raw_cat:
        try:
            cat_id = int(raw_cat[0])
        except (TypeError, ValueError):
            cat_id = None
    if extra_category_ids and cat_id is not None and cat_id in extra_category_ids:
        return True

    blob_lower = f'{name} {cat} {typ}'.lower()
    if _match_qhse_text(name, cat, typ, 'extincteur') or 'extinct' in blob_lower:
        return True

    typ_s = str(typ or '').strip()
    if typ_s and typ_s.lower() not in ('-', '—', 'n/a', 'na', '.', 'none', 'aucun'):
        return True

    if re.search(r'(?:^|[^a-z0-9])(tf|tg)(?:$|[^a-z0-9])', blob_lower, re.I):
        return True

    for frag in (
        'incendie', 'poudre', ' mousse', 'mousse ', 'co2', 'eau pulv', 'pulveris',
        'portatif', 'portable', 'recharge', 'unece',
    ):
        if frag in blob_lower:
            return True
    return False


def _fetch_extincteurs_status(site_filter=''):
    rows = []
    error = None
    today = timezone.localdate()
    extra_category_ids = list(getattr(settings, 'ODOO_EXTINCTEUR_CATEGORY_IDS', ()) or ())
    try:
        uid, models = get_odoo_connection()
        candidates = [
            'name', 'category_id', 'serial_no', 'state', 'location', 'location_id', 'active',
            'x_type_extincteur', 'x_zone', 'x_emplacement', 'site_id', 'x_site',
            'x_date_validite', 'x_prochaine_date', 'x_next_control_date', 'next_action_date',
            'warranty_date', 'x_date_dernier_controle', 'last_maintenance_date',
            'x_date_fin_validite', 'x_date_expiration', 'x_validite',
        ]
        fields = _safe_model_fields(uid, models, 'maintenance.equipment', candidates)
        if 'name' not in fields:
            fields.append('name')
        equipements = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'maintenance.equipment', 'search_read',
            [[]],
            {'fields': fields, 'order': 'name asc', 'limit': 3000},
        ) or []

        for e in equipements:
            name = str(e.get('name') or '')
            cat = (e.get('category_id') or ['', ''])[1] if isinstance(e.get('category_id'), list) else ''
            typ = str(e.get('x_type_extincteur') or '')
            if not _maintenance_equipment_is_extincteur(e, name, cat, typ, extra_category_ids):
                continue

            site = '—'
            if isinstance(e.get('site_id'), list) and len(e.get('site_id')) > 1:
                site = e['site_id'][1]
            elif isinstance(e.get('x_site'), list) and len(e.get('x_site')) > 1:
                site = e['x_site'][1]
            elif isinstance(e.get('location_id'), list) and len(e.get('location_id')) > 1:
                site = e['location_id'][1]
            if site_filter and site_filter.lower() not in str(site).lower():
                continue

            due_candidates = [
                e.get('x_date_validite'), e.get('x_prochaine_date'),
                e.get('x_next_control_date'), e.get('next_action_date'), e.get('warranty_date'),
                e.get('x_date_fin_validite'), e.get('x_date_expiration'), e.get('x_validite'),
            ]
            due_date = None
            for raw in due_candidates:
                due_date = _parse_date_any(raw)
                if due_date:
                    break
            days_left = None if not due_date else (due_date - today).days
            status = 'OK'
            if days_left is not None and days_left < 0:
                status = 'Echu'
            elif days_left is not None and days_left <= 7:
                status = 'Alerte J-7'

            rows.append({
                'designation': name or '—',
                'site': site or '—',
                'zone': e.get('x_zone') or e.get('x_emplacement') or e.get('location') or '—',
                'type_extincteur': typ or 'Non renseigne',
                'serial': e.get('serial_no') or '—',
                'date_validite': due_date.strftime('%d/%m/%Y') if due_date else '—',
                'jours_restant': days_left if days_left is not None else '—',
                'statut': status,
                'dernier_controle': _format_date_fr(e.get('x_date_dernier_controle') or e.get('last_maintenance_date')),
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'

    alertes = [r for r in rows if r.get('statut') in ('Alerte J-7', 'Echu')]
    return {
        'rows': rows,
        'alertes': alertes,
        'kpi_total': len(rows),
        'kpi_alertes': len(alertes),
        'kpi_echus': len([r for r in rows if r.get('statut') == 'Echu']),
        'kpi_j7': len([r for r in rows if r.get('statut') == 'Alerte J-7']),
        'error': error,
    }


def _build_site_report():
    """Rapport synthétique par site (incidents + stock EPI)."""
    incidents = _fetch_incidents_dashboard()
    products, prod_error = _fetch_qhse_products(status_filter='all', search_term='')
    by_cat = defaultdict(float)
    for p in products:
        by_cat[p.get('categorie') or 'Non renseigne'] += float(p.get('stock') or 0)
    return {
        'incidents_by_site': incidents.get('sites', []),
        'stock_by_category': sorted(by_cat.items(), key=lambda x: x[1], reverse=True),
        'error': incidents.get('error') or prod_error,
    }


@login_required
def qhse_achats(request):
    statut = (request.GET.get('statut') or 'all').strip().lower()
    fournisseur = (request.GET.get('fournisseur') or '').strip()
    rows, fournisseurs, error = _fetch_qhse_purchase_orders(status_filter=statut, supplier_filter=fournisseur)

    total_montant = sum(r.get('montant_total') or 0 for r in rows)
    fournisseurs_distincts = len({r.get('fournisseur') for r in rows if r.get('fournisseur') and r.get('fournisseur') != '—'})
    en_attente = sum(1 for r in rows if r.get('etat') in ('draft', 'sent', 'to approve'))

    return render(request, 'qhse/achats.html', {
        'page_title': 'Achats QHSE',
        'page_subtitle': 'Suivi des achats de categorie QHSE recupere depuis Odoo.',
        'error': error,
        'rows': rows,
        'fournisseurs': fournisseurs,
        'selected_fournisseur': fournisseur,
        'selected_statut': statut,
        'kpi_total_commandes': len(rows),
        'kpi_fournisseurs': fournisseurs_distincts,
        'kpi_en_attente': en_attente,
        'kpi_montant': format_number(total_montant),
    })


@login_required
def qhse_produits_hse(request):
    statut = (request.GET.get('statut') or 'all').strip().lower()
    recherche = (request.GET.get('q') or '').strip()
    rows, error = _fetch_qhse_products(status_filter=statut, search_term=recherche)

    total_stock = sum(r.get('stock') or 0 for r in rows)
    en_stock = sum(1 for r in rows if (r.get('stock') or 0) > 0)
    activables_achat = sum(1 for r in rows if r.get('achat'))

    return render(request, 'qhse/produits_hse.html', {
        'page_title': 'Produits HSE',
        'page_subtitle': 'Catalogue des produits HSE detectes dans Odoo.',
        'error': error,
        'rows': rows,
        'selected_statut': statut,
        'search_term': recherche,
        'kpi_total': len(rows),
        'kpi_en_stock': en_stock,
        'kpi_achat': activables_achat,
        'kpi_stock_total': format_number(total_stock),
    })


@login_required
def qhse_factures(request):
    return _render_qhse_service_page(
        request,
        page_title='Factures QHSE',
        page_subtitle='Facturation des depenses et services QHSE.',
    )


@login_required
def qhse_consommations(request):
    site = (request.GET.get('site') or '').strip()
    personne = (request.GET.get('personne') or '').strip()
    rows, sites, personnes, kpis, error = _fetch_epi_movements(site_filter=site, person_filter=personne)
    return render(request, 'qhse/consommations.html', {
        'page_title': 'Consommations QHSE',
        'page_subtitle': 'Sorties EPI detaillees: affectation, transfert site et suivi par personne.',
        'rows': rows,
        'sites': sites,
        'personnes': personnes,
        'selected_site': site,
        'selected_personne': personne,
        'kpis': kpis,
        'error': error,
    })


@login_required
def qhse_dashboard(request):
    """Tableau de bord QHSE — données Odoo réelles via QHSEService."""
    from .services.qhse_service import QHSEService
    svc = QHSEService()
    data = svc.get_dashboard_data()
    return render(request, 'qhse/dashboard.html', {
        'page_title': 'Dashboard QHSE',
        'annee': datetime.now().year,
        **data,
    })


@login_required
def qhse_incidents(request):
    """Incidents & Accidents — synoptique corporel, graphiques, tableau réel Odoo."""
    from .services.qhse_service import QHSEService

    site = (request.GET.get('site') or '').strip()
    mois = (request.GET.get('mois') or '').strip()
    svc = QHSEService()
    acc = svc.get_accidents(site_filter=site, mois_filter=mois)
    chart = QHSEService.to_chart_json(acc)
    selected_month = mois or (chart['months_list'][-1] if chart['months_list'] else '')
    demo_rows = [
        {'date':'2025-01-15','ref':'DEMO/QHSE/01/0001','site':'Site 1','categorie':'Chute','responsable':'Employé A','state':'done'},
        {'date':'2025-02-08','ref':'DEMO/QHSE/02/0002','site':'Site 2','categorie':'Coupure','responsable':'Employé B','state':'in_progress'},
        {'date':'2025-03-22','ref':'DEMO/QHSE/03/0003','site':'Site 3','categorie':'Brûlure','responsable':'Employé C','state':'draft'},
        {'date':'2025-04-10','ref':'DEMO/QHSE/04/0004','site':'Site 1','categorie':'Chute','responsable':'Employé D','state':'done'},
        {'date':'2025-05-05','ref':'DEMO/QHSE/05/0005','site':'Administration','categorie':'Contusion','responsable':'Employé E','state':'in_progress'},
    ]
    return render(request, 'qhse/incidents_accidents.html', {
        'page_title': 'Incidents & Accidents',
        'page_subtitle': 'Gestion et analyse des incidents, accidents de travail et maladies professionnelles.',
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
        'demo_rows': demo_rows,
        'error': acc.get('error'),
        **chart,
    })


@login_required
def qhse_indicateurs(request):
    site = (request.GET.get('site') or '').strip()
    data = _fetch_extincteurs_status(site_filter=site)
    return render(request, 'qhse/indicateurs.html', {
        'page_title': 'Indicateurs HSE (TF, TG)',
        'page_subtitle': 'Suivi extincteurs, controle reglementaire et alertes J-7.',
        'selected_site': site,
        **data,
    })


@login_required
def qhse_plan_actions(request):
    """Plan d'actions JAM — données Odoo réelles (project.task ou jam.action)."""
    from .services.qhse_service import QHSEService
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


@login_required
def qhse_action_detail(request, action_id: int):
    """Détail d'une action JAM (project.task ou jam.action)."""
    from .services.qhse_service import QHSEService
    svc = QHSEService()
    action = svc.get_action_detail(action_id)
    if action is None:
        from django.http import Http404
        raise Http404('Action introuvable')
    return render(request, 'qhse/action_detail.html', {
        'page_title': f"Action #{action_id}",
        'action': action,
        'error': svc._error,
    })


@login_required
def qhse_audits(request):
    report = _build_site_report()
    return render(request, 'qhse/audits.html', {
        'page_title': 'Audits qualite',
        'page_subtitle': 'Rapports par site et stock detail EPI.',
        'incidents_by_site': report.get('incidents_by_site', []),
        'stock_by_category': report.get('stock_by_category', []),
        'error': report.get('error'),
    })


def _fetch_compta_vendor_bills(request, payment_states=None, limit=2500):
    rows = []
    error = None
    try:
        uid, models = get_odoo_connection()
        avail = set(
            (models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'account.move', 'fields_get', [], {'attributes': ['type']}
            ) or {}).keys()
        )
        fields = ['name']
        for f in (
            'move_type', 'state', 'invoice_date', 'invoice_date_due', 'partner_id',
            'ref', 'invoice_origin', 'invoice_payment_term_id', 'payment_state',
            'amount_total', 'amount_untaxed', 'amount_tax',
        ):
            if f in avail:
                fields.append(f)

        domain = [('move_type', '=', 'in_invoice'), ('state', '=', 'posted')]
        if payment_states:
            domain.append(('payment_state', 'in', payment_states))

        records = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read', [domain],
            {'fields': sorted(set(fields)), 'limit': limit, 'order': 'invoice_date desc,id desc'},
        )
        for r in records:
            inv_date = str(r.get('invoice_date') or '')[:10]
            due_date = str(r.get('invoice_date_due') or '')[:10]
            partner = r.get('partner_id')
            term = r.get('invoice_payment_term_id')
            pay_state = str(r.get('payment_state') or '')
            rows.append({
                'date_facture': f"{inv_date[8:10]}/{inv_date[5:7]}/{inv_date[0:4]}" if len(inv_date) == 10 else '—',
                'fournisseur': partner[1] if isinstance(partner, list) and len(partner) > 1 else '—',
                'reference': r.get('ref') or r.get('name') or '—',
                'designation': r.get('invoice_origin') or 'PRESTATION DE SERVICE',
                'montant_ttc': round(float(r.get('amount_total') or 0), 2),
                'date_echeance': f"{due_date[8:10]}/{due_date[5:7]}/{due_date[0:4]}" if len(due_date) == 10 else '—',
                'delai': term[1] if isinstance(term, list) and len(term) > 1 else '—',
                'convention': 'OUI' if pay_state in ('paid', 'in_payment') else 'NON',
                'payment_state': pay_state or 'unknown',
                'amount_tax': float(r.get('amount_tax') or 0),
                'amount_ht': float(r.get('amount_untaxed') or 0),
                'amount_ttc': float(r.get('amount_total') or 0),
                'invoice_date_raw': inv_date,
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return rows, error


@login_required
def compta_paid(request):
    rows, error = _fetch_compta_vendor_bills(request, payment_states=['paid', 'in_payment'])
    total_ttc = round(sum(r['montant_ttc'] for r in rows), 2)
    return render(request, 'comptabilite/factures_payees.html', {
        'error': error,
        'rows': rows,
        'kpi_total': len(rows),
        'kpi_montant': format_number(total_ttc),
    })


@login_required
def compta_unpaid(request):
    rows, error = _fetch_compta_vendor_bills(request, payment_states=['not_paid', 'partial'])
    total_ttc = round(sum(r['montant_ttc'] for r in rows), 2)
    return render(request, 'comptabilite/factures_non_payees.html', {
        'error': error,
        'rows': rows,
        'kpi_total': len(rows),
        'kpi_montant': format_number(total_ttc),
    })


@login_required
def compta_tva(request):
    month = (request.GET.get('mois') or datetime.now().strftime('%Y-%m'))[:7]
    rows, error = _fetch_compta_vendor_bills(request, payment_states=None)
    month_rows = [r for r in rows if (r.get('invoice_date_raw') or '').startswith(month)]
    total_ht = round(sum(r['amount_ht'] for r in month_rows), 2)
    total_tva = round(sum(r['amount_tax'] for r in month_rows), 2)
    total_ttc = round(sum(r['amount_ttc'] for r in month_rows), 2)
    taux = round((total_tva * 100 / total_ht), 2) if total_ht else 0
    return render(request, 'comptabilite/tva.html', {
        'error': error,
        'rows': month_rows[:500],
        'month': month,
        'kpi_ht': format_number(total_ht),
        'kpi_tva': format_number(total_tva),
        'kpi_ttc': format_number(total_ttc),
        'kpi_taux': f'{taux}%',
    })


@login_required
def compta_fournisseurs(request):
    return compta_unpaid(request)


def _fetch_compta_customer_invoices(limit=2500):
    rows, error = [], None
    try:
        uid, models = get_odoo_connection()
        fields = ['name', 'invoice_date', 'invoice_date_due', 'partner_id', 'ref', 'invoice_origin', 'payment_state', 'amount_total']
        recs = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [[('move_type', '=', 'out_invoice'), ('state', '=', 'posted')]],
            {'fields': fields, 'limit': limit, 'order': 'invoice_date desc,id desc'},
        )
        for r in recs:
            d = str(r.get('invoice_date') or '')[:10]
            due = str(r.get('invoice_date_due') or '')[:10]
            p = r.get('partner_id')
            pay_state = str(r.get('payment_state') or '')
            rows.append({
                'date_facture': f"{d[8:10]}/{d[5:7]}/{d[0:4]}" if len(d) == 10 else '—',
                'client': p[1] if isinstance(p, list) and len(p) > 1 else '—',
                'reference': r.get('ref') or r.get('name') or '—',
                'designation': r.get('invoice_origin') or 'VENTE',
                'montant_ttc': round(float(r.get('amount_total') or 0), 2),
                'date_echeance': f"{due[8:10]}/{due[5:7]}/{due[0:4]}" if len(due) == 10 else '—',
                'etat_reglement': pay_state or 'unknown',
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return rows, error


@login_required
def compta_factures_client(request):
    rows, error = _fetch_compta_customer_invoices()
    return render(request, 'comptabilite/factures_client.html', {
        'error': error,
        'rows': rows,
        'kpi_total': len(rows),
        'kpi_montant': format_number(sum(r['montant_ttc'] for r in rows)),
    })


def _fetch_compta_payments(direction='outbound', limit=2500):
    rows, error = [], None
    try:
        uid, models = get_odoo_connection()
        pay_av = set((models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.payment', 'fields_get', [], {'attributes': ['type']}
        ) or {}).keys())
        fields = ['name', 'amount', 'date', 'partner_id']
        for f in ('payment_type', 'state', 'ref', 'journal_id'):
            if f in pay_av:
                fields.append(f)
        dom = []
        if 'payment_type' in pay_av:
            dom.append(('payment_type', '=', direction))
        if 'state' in pay_av:
            dom.append(('state', '=', 'posted'))
        recs = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.payment', 'search_read', [dom],
            {'fields': sorted(set(fields)), 'limit': limit, 'order': 'date desc,id desc'},
        )
        for r in recs:
            d = str(r.get('date') or '')[:10]
            p = r.get('partner_id')
            j = r.get('journal_id')
            rows.append({
                'date': f"{d[8:10]}/{d[5:7]}/{d[0:4]}" if len(d) == 10 else '—',
                'tiers': p[1] if isinstance(p, list) and len(p) > 1 else '—',
                'reference': r.get('ref') or r.get('name') or '—',
                'journal': j[1] if isinstance(j, list) and len(j) > 1 else '—',
                'montant': float(r.get('amount') or 0),
            })
    except Exception as exc:
        error = f'Erreur Odoo: {exc}'
    return rows, error


@login_required
def compta_decaissements(request):
    rows, error = _fetch_compta_payments('outbound')
    return render(request, 'comptabilite/paiements_decaissements.html', {
        'error': error, 'rows': rows,
        'kpi_total': len(rows),
        'kpi_montant': format_number(sum(r['montant'] for r in rows)),
    })


@login_required
def compta_encaissements(request):
    rows, error = _fetch_compta_payments('inbound')
    return render(request, 'comptabilite/encaissements.html', {
        'error': error, 'rows': rows,
        'kpi_total': len(rows),
        'kpi_montant': format_number(sum(r['montant'] for r in rows)),
    })


@login_required
def compta_tresorerie(request):
    out_rows, out_err = _fetch_compta_payments('outbound')
    in_rows, in_err = _fetch_compta_payments('inbound')
    total_out = sum(r['montant'] for r in out_rows)
    total_in = sum(r['montant'] for r in in_rows)
    solde = total_in - total_out
    return render(request, 'comptabilite/tresorerie.html', {
        'error': out_err or in_err,
        'kpi_in': format_number(total_in),
        'kpi_out': format_number(total_out),
        'kpi_solde': format_number(solde),
        'rows_in': in_rows[:50],
        'rows_out': out_rows[:50],
    })


@login_required
def compta_analyse_projet(request):
    rows, error = _fetch_compta_vendor_bills(request, payment_states=None)
    grouped = defaultdict(lambda: {'count': 0, 'ttc': 0.0})
    for r in rows:
        key = (r.get('designation') or 'Sans projet').strip()[:80]
        grouped[key]['count'] += 1
        grouped[key]['ttc'] += float(r.get('amount_ttc') or 0)
    out = [{'projet': k, 'nb': v['count'], 'montant': round(v['ttc'], 2)} for k, v in grouped.items()]
    out.sort(key=lambda x: x['montant'], reverse=True)
    return render(request, 'comptabilite/analyse_projet.html', {
        'error': error,
        'rows': out[:300],
        'kpi_total': len(out),
        'kpi_montant': format_number(sum(x['montant'] for x in out)),
    })


# ─────────────────────────────────────────────
#  GASOIL — EXPORTS (CSV/PDF alias for Excel)
# ─────────────────────────────────────────────

@login_required
def gasoil_sorties_csv(request):
    """Export CSV alias pour gasoil_sorties_export (Excel)."""
    return gasoil_sorties_export(request)


@login_required
def gasoil_rapport(request):
    """Page HTML du rapport sorties gasoil avec données d'exemple."""
    # Données d'exemple pour prévisualiser le rapport HTML.
    bons_data = [
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09154',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50467 6',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'MOHAMMED HADDAD',
            'cpt_initial': 1491,
            'cpt_actuel': 1498,
            'ecart': 7.0,
            'product_qty': 70.0,
            'consommation': 10.0,
            'anomalie': 'OK',
        },
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09153',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50468',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'ABDELOUAHAB B OUYGHRAQUINE',
            'cpt_initial': 1497,
            'cpt_actuel': 1504,
            'ecart': 7.0,
            'product_qty': 59.0,
            'consommation': 8.43,
            'anomalie': 'OK',
        },
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09152',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50469',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'MUSTAFA BOUAIOUN',
            'cpt_initial': 1538,
            'cpt_actuel': 1545,
            'ecart': 7.0,
            'product_qty': 70.0,
            'consommation': 10.0,
            'anomalie': 'OK',
        },
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09151',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50470',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'MOHAMMED EL MAKOUDI',
            'cpt_initial': 1500,
            'cpt_actuel': 1506,
            'ecart': 6.0,
            'product_qty': 60.0,
            'consommation': 10.0,
            'anomalie': 'OK',
        },
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09150',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50471',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'AZIZ EL FTATCHI',
            'cpt_initial': 1529,
            'cpt_actuel': 1535,
            'ecart': 6.0,
            'product_qty': 50.0,
            'consommation': 8.33,
            'anomalie': 'OK',
        },
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09149',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50472',
            'categorie': 'CAMION ENGIN',
            'chauffeur': 'MUSTAPHA MAHJOUB',
            'cpt_initial': 1561,
            'cpt_actuel': 1567,
            'ecart': 6.0,
            'product_qty': 52.0,
            'consommation': 8.67,
            'anomalie': 'OK',
        },
    ]

    total_qty = sum(b['product_qty'] for b in bons_data)

    return render(request, 'gasoil/rapport.html', {
        'bons': bons_data,
        'total_bons': len(bons_data),
        'total_qty': total_qty,
    })


@login_required
def achats_synthese_pdf(request):
    import datetime
    import os
    from io import BytesIO
    from xml.sax.saxutils import escape as xml_escape
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
        KeepTogether, PageBreak,
    )
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.pdfgen import canvas as rl_canvas_mod

    def _safe_xml(s):
        """Échappe &lt; &gt; &amp; pour les Parapraph ReportLab (données Odoo)."""
        return xml_escape(str(s), entities={'"': '&quot;', "'": '&apos;'})

    today = datetime.date.today()
    W, _H = A4

    navy       = rl_colors.HexColor('#1a2c4e')
    orange     = rl_colors.HexColor('#E87722')
    white      = rl_colors.white
    gray       = rl_colors.HexColor('#6B7280')
    light      = rl_colors.HexColor('#F8FAFC')
    light_blue = rl_colors.HexColor('#f0f4ff')
    red        = rl_colors.HexColor('#dc2626')
    green      = rl_colors.HexColor('#16a34a')
    border_c   = rl_colors.HexColor('#E5E7EB')
    date_debut_q = request.GET.get('date_debut', '').strip()
    date_fin_q = request.GET.get('date_fin', '').strip()
    periode_q = request.GET.get('periode', 'annee').strip().lower()
    societe_q = request.GET.get('societe', '').strip()

    domain_po = []
    period_label_fr = 'Historique complet'
    period_slug = 'all'

    def _parse_iso_pdf(d):
        try:
            return datetime.datetime.strptime((d or '')[:10], '%Y-%m-%d').date()
        except (ValueError, TypeError):
            return None

    if date_debut_q and date_fin_q:
        ds = _parse_iso_pdf(date_debut_q)
        de = _parse_iso_pdf(date_fin_q)
        if ds and de:
            domain_po = [
                ('date_order', '>=', ds.isoformat()),
                ('date_order', '<=', de.strftime('%Y-%m-%d') + ' 23:59:59'),
            ]
            period_label_fr = 'Du %s au %s' % (ds.strftime('%d/%m/%Y'), de.strftime('%d/%m/%Y'))
            period_slug = '%s_%s' % (ds.strftime('%Y%m%d'), de.strftime('%Y%m%d'))
    elif periode_q == 'tout':
        domain_po = []
        period_label_fr = 'Historique complet (tous les bons)'
        period_slug = 'all'
    elif periode_q == 'mois':
        first = today.replace(day=1)
        domain_po = [
            ('date_order', '>=', first.isoformat()),
            ('date_order', '<=', today.strftime('%Y-%m-%d') + ' 23:59:59'),
        ]
        period_label_fr = 'Mois en cours (%s)' % first.strftime('%m/%Y')
        period_slug = 'mois_%s' % first.strftime('%Y%m')
    elif periode_q == 'trimestre':
        q = (today.month - 1) // 3
        first_month = 3 * q + 1
        first = today.replace(month=first_month, day=1)
        domain_po = [
            ('date_order', '>=', first.isoformat()),
            ('date_order', '<=', today.strftime('%Y-%m-%d') + ' 23:59:59'),
        ]
        period_label_fr = 'Trimestre en cours (depuis le %s)' % first.strftime('%d/%m/%Y')
        period_slug = 'trim_%s' % first.strftime('%Y%m')
    else:
        first = today.replace(month=1, day=1)
        domain_po = [
            ('date_order', '>=', first.isoformat()),
            ('date_order', '<=', today.strftime('%Y-%m-%d') + ' 23:59:59'),
        ]
        period_label_fr = 'Année %d (au %s)' % (today.year, today.strftime('%d/%m/%Y'))
        period_slug = 'annee_%d' % today.year

    if societe_q:
        domain_po.append(('company_id.name', 'ilike', societe_q))
        period_label_fr += ' · Filiale / société : « %s »' % societe_q

    da_domain = []
    d_from = d_to = None
    for dom in domain_po:
        if dom[0] == 'date_order' and dom[1] == '>=':
            d_from = (dom[2] or '')[:10]
        if dom[0] == 'date_order' and dom[1] == '<=':
            d_to = (dom[2] or '')[:19]
            if len(d_to) >= 10:
                d_to = d_to[:10]
    if d_from and d_to:
        da_domain = [
            ('create_date', '>=', d_from + ' 00:00:00'),
            ('create_date', '<=', d_to + ' 23:59:59'),
        ]

    try:
        uid, models_proxy = get_odoo_connection()
        db = settings.ODOO_DB
        pw = settings.ODOO_PASS
        odoo_ok = True
    except Exception:
        odoo_ok = False

    bons = []
    total_bons = confirmes = en_retard = taux_conf = 0
    nb_recues = nb_en_cours = nb_annules = 0
    montant_ttc = 0.0
    montant_by_ccy_id = defaultdict(float)
    currency_label_by_id = {0: 'MAD'}
    multi_ccy = False
    montant_currency_label = 'MAD'
    fourn_stats = {}
    partners_period = set()
    company_agg = defaultdict(lambda: {'name': '—', 'nb': 0, 'by_ccy': defaultdict(float)})
    delai_plan_jours = []
    nb_fourn_catalog = 0
    nb_draft = nb_sent = 0

    if odoo_ok:
        try:
            bons = odoo_search_read_all(
                uid, models_proxy, 'purchase.order', list(domain_po),
                [
                    'state', 'amount_total', 'date_planned', 'date_order',
                    'partner_id', 'currency_id', 'company_id', 'name',
                ],
                order='date_order desc',
            )
            total_bons = len(bons)
            confirmes = sum(1 for b in bons if b.get('state') in ['purchase', 'done'])
            nb_recues = sum(1 for b in bons if b.get('state') == 'done')
            nb_en_cours = sum(1 for b in bons if b.get('state') == 'purchase')
            nb_annules = sum(1 for b in bons if b.get('state') == 'cancel')
            nb_draft = sum(1 for b in bons if b.get('state') == 'draft')
            nb_sent = sum(1 for b in bons if b.get('state') == 'sent')
            for b in bons:
                amt = _odoo_float(b.get('amount_total'))
                cc_id, cc_lbl = _odoo_currency_id_label(b)
                currency_label_by_id[cc_id] = cc_lbl
                montant_by_ccy_id[cc_id] += amt

                comp = b.get('company_id')
                if isinstance(comp, (list, tuple)) and len(comp) >= 1:
                    try:
                        scid = int(comp[0])
                    except (TypeError, ValueError):
                        scid = 0
                    scname = comp[1] if len(comp) > 1 else '—'
                    cg = company_agg.setdefault(
                        scid, {'name': scname, 'nb': 0, 'by_ccy': defaultdict(float)}
                    )
                    cg['nb'] += 1
                    cg['by_ccy'][cc_id] += amt

                p = b.get('partner_id')
                if isinstance(p, (list, tuple)) and len(p) >= 1:
                    try:
                        pid = int(p[0])
                    except (TypeError, ValueError):
                        pid = None
                    if pid is not None:
                        partners_period.add(pid)
                        pname = p[1] if len(p) >= 2 else ('Fournisseur #%s' % pid)
                        st = fourn_stats.setdefault(
                            pid, {'name': pname, 'nb': 0, 'by_ccy': defaultdict(float)}
                        )
                        st['nb'] += 1
                        st['by_ccy'][cc_id] += amt

                do = (b.get('date_order') or '')[:10]
                dp = (b.get('date_planned') or '')[:10]
                if len(do) == 10 and len(dp) == 10:
                    try:
                        d_o = datetime.datetime.strptime(do, '%Y-%m-%d').date()
                        d_p = datetime.datetime.strptime(dp, '%Y-%m-%d').date()
                        delai_plan_jours.append((d_p - d_o).days)
                    except ValueError:
                        pass

                if b.get('state') in ['done', 'cancel']:
                    continue
                dp2 = str(b.get('date_planned', '') or '')[:10]
                if dp2 and len(dp2) == 10:
                    try:
                        if datetime.datetime.strptime(dp2, '%Y-%m-%d').date() < today:
                            en_retard += 1
                    except ValueError:
                        pass
            multi_ccy = len(montant_by_ccy_id) > 1
            if not multi_ccy and montant_by_ccy_id:
                only_id = next(iter(montant_by_ccy_id.keys()))
                montant_ttc = montant_by_ccy_id[only_id]
                montant_currency_label = currency_label_by_id.get(only_id, 'MAD')
            elif montant_by_ccy_id:
                montant_ttc = 0.0
            taux_conf = (confirmes / total_bons * 100) if total_bons > 0 else 0
        except Exception:
            pass

    delai_moy_j = (
        sum(delai_plan_jours) / len(delai_plan_jours) if delai_plan_jours else None
    )
    nb_fourn_periode = len(partners_period)

    def _fmt_ccy_breakdown(by_ccy_map, decimals=0):
        parts = sorted(by_ccy_map.items(), key=lambda x: -x[1])
        out = []
        for cid, amt in parts[:5]:
            if amt == 0:
                continue
            out.append('%s %s' % (
                format_number_decimals(amt, decimals),
                currency_label_by_id.get(cid, '?'),
            ))
        txt = ' · '.join(out)
        if len(parts) > 5:
            txt += ' · …'
        return txt or '—'

    top_list = sorted(
        fourn_stats.items(),
        key=lambda kv: (-kv[1]['nb'], -sum(kv[1]['by_ccy'].values())),
    )[:5]
    top_fourn = []
    for _pid, st in top_list:
        nb_cmd = st['nb']
        byc = st['by_ccy']
        if not multi_ccy and len(byc) <= 1:
            raw = sum(byc.values())
            disp = format_number_decimals(raw, 0) + ' ' + montant_currency_label
            top_fourn.append((st['name'], nb_cmd, disp, raw))
        else:
            disp = _fmt_ccy_breakdown(byc)
            top_fourn.append((st['name'], nb_cmd, disp, None))

    if odoo_ok:
        try:
            nb_fourn_catalog = models_proxy.execute_kw(
                db, uid, pw, 'res.partner', 'search_count',
                [[('supplier_rank', '>', 0)]]
            )
        except Exception:
            nb_fourn_catalog = 0

    total_da = montant_da = urgentes = 0
    da_source_label = 'Demandes d’achat (purchase.request)'
    if odoo_ok:
        try:
            das = odoo_search_read_all(
                uid, models_proxy, 'purchase.request', list(da_domain),
                ['state', 'estimated_cost'],
                order='id desc',
            )
            total_da = len(das)
            montant_da = sum(_odoo_float(d.get('estimated_cost')) for d in das)
            urgentes = sum(
                1 for d in das if _odoo_float(d.get('estimated_cost')) > 50000
            )
        except Exception:
            try:
                domain_fb = list(domain_po) + [('state', 'in', ['draft', 'sent'])]
                po_fb = odoo_search_read_all(
                    uid, models_proxy, 'purchase.order', domain_fb,
                    ['amount_total'],
                    order='date_order desc',
                )
                da_source_label = 'Approximation : bons brouillon / envoyé (module DA absent)'
                total_da = len(po_fb)
                montant_da = sum(_odoo_float(x.get('amount_total')) for x in po_fb)
                urgentes = sum(
                    1 for x in po_fb if _odoo_float(x.get('amount_total')) > 50000
                )
            except Exception:
                pass

    total_rfq = 0
    if odoo_ok:
        try:
            domain_rfq = list(domain_po) + [('state', '=', 'draft')]
            rfqs = odoo_search_read_all(
                uid, models_proxy, 'purchase.order', domain_rfq,
                ['state'],
                order='id desc',
            )
            total_rfq = len(rfqs)
        except Exception:
            pass

    analyse = []
    if delai_moy_j is not None and total_bons > 0:
        analyse.append(
            ('ok',
             'Délai moyen entre date de commande et date planifiée : %.0f jour(s) '
             '(%d bon(s) avec les deux dates renseignées).' % (
                 delai_moy_j, len(delai_plan_jours)))
        )
    if not multi_ccy and montant_ttc > 0:
        top5_amt = sum((m for _, _, _, m in top_fourn if m is not None), 0.0)
        if top5_amt > 0:
            analyse.append(
                ('ok',
                 'Les cinq principaux fournisseurs concentrent %.1f%% du montant TTC '
                 'sur le périmètre.' % (top5_amt / montant_ttc * 100))
            )
    if total_bons > 0:
        if taux_conf >= 99:
            analyse.append(('ok', 'Excellent taux de confirmation (%.1f%%) - performance optimale.' % taux_conf))
        elif taux_conf >= 80:
            analyse.append(('ok', 'Bon taux de confirmation (%.1f%%).' % taux_conf))
        else:
            analyse.append(('warn', 'Taux de confirmation faible (%.1f%%) - actions correctives requises.' % taux_conf))
        if en_retard > 0:
            pct = en_retard / total_bons * 100
            analyse.append(('warn', '%d bons en retard (%.1f%%) — suivi prioritaire recommandé.' % (en_retard, pct)))
        if nb_recues > 0:
            pct_rec = nb_recues / total_bons * 100
            analyse.append(('ok', '%d livraisons réceptionnées (%.1f%% du total).' % (nb_recues, pct_rec)))
    if urgentes > 0:
        analyse.append(('alert', '%d demandes urgentes en attente de traitement.' % urgentes))
    if not analyse:
        analyse.append(('ok', 'Aucune anomalie détectée — situation nominale.'))

    _prio = {'warn': 0, 'alert': 1, 'ok': 2}
    analyse.sort(key=lambda item: (_prio.get(item[0], 9), item[1]))

    class _NumberedCanvas(rl_canvas_mod.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []
            self._pdf_meta_done = False

        def showPage(self):
            if not self._pdf_meta_done:
                try:
                    self.setTitle('SOMATRIN — Synthèse achats')
                    self.setSubject('Achats & Approvisionnement')
                    self.setAuthor('SOMATRIN Dashboard')
                except Exception:
                    pass
                self._pdf_meta_done = True
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total_pg = len(self._saved_page_states)
            for i, state in enumerate(self._saved_page_states):
                self.__dict__.update(state)
                self._draw_footer(i + 1, total_pg)
                rl_canvas_mod.Canvas.showPage(self)
            rl_canvas_mod.Canvas.save(self)

        def _draw_footer(self, page_num, total_pg):
            self.saveState()
            self.setFont('Helvetica', 7)
            self.setFillColor(rl_colors.HexColor('#6B7280'))
            footer = (
                'SOMATRIN — Rapport de synthèse Achats & Approvisionnement'
                ' — Document confidentiel — usage interne — %s'
                % today.strftime('%d/%m/%Y')
            )
            self.drawCentredString(W / 2, 0.92 * cm, footer)
            self.setFont('Helvetica', 6.5)
            self.drawCentredString(
                W / 2, 0.74 * cm,
                period_label_fr[:140] + ('…' if len(period_label_fr) > 140 else ''),
            )
            self.drawRightString(W - 1.5 * cm, 0.74 * cm, 'Page %d / %d' % (page_num, total_pg))
            self.restoreState()

    styles   = getSampleStyleSheet()
    s_center = ParagraphStyle('hc', alignment=TA_CENTER, spaceAfter=0)
    s_right  = ParagraphStyle('hr', alignment=TA_RIGHT,  spaceAfter=0)
    s_note = ParagraphStyle('note', fontSize=7.5, textColor=gray, spaceBefore=2, spaceAfter=0,
                            leading=9, leftIndent=0)
    s_bar_title = ParagraphStyle(
        'bar_t', fontSize=10.5, textColor=white, fontName='Helvetica-Bold',
        alignment=TA_LEFT, spaceAfter=0, leading=12,
    )

    PAGE_W = 18 * cm
    LBL_W = 6.2 * cm
    VAL_W = PAGE_W - LBL_W

    def _hdr(txt):
        return Paragraph('<b><font color="white">%s</font></b>' % _safe_xml(txt), styles['Normal'])

    def _section_bar(title, width=None):
        """Bandeau titre de section."""
        w = width if width is not None else PAGE_W
        p = Paragraph('<font size="10.5">%s</font>' % _safe_xml(title.upper()), s_bar_title)
        t = Table([[p]], colWidths=[w], rowHeights=[None])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), navy),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        return t

    def _pair_table(rows, col_lbl=None, col_val=None, extra_styles=None):
        """
        Tableau indicateur | valeur (rows[0] = en-têtes).
        rows: liste de (cellule_gauche, cellule_droite) — str ou Flowable.
        """
        extra_styles = extra_styles or []
        cl = col_lbl if col_lbl is not None else LBL_W
        cv = col_val if col_val is not None else VAL_W
        tbl = Table(rows, colWidths=[cl, cv], repeatRows=1)
        ts = [
            ('BACKGROUND', (0, 0), (-1, 0), navy),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, light]),
            ('GRID', (0, 0), (-1, -1), 0.5, border_c),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 7),
            ('RIGHTPADDING', (0, 0), (-1, -1), 7),
        ] + list(extra_styles)
        tbl.setStyle(TableStyle(ts))
        return tbl

    def _card(title, body_tbl, footnote_para=None, width=None):
        """Carte : bandeau + corps + note optionnelle, bordure légère."""
        w = width if width is not None else PAGE_W
        rows = [[_section_bar(title, w)], [body_tbl]]
        if footnote_para is not None:
            rows.append([footnote_para])
        t = Table(rows, colWidths=[w])
        ts = [
            ('BOX', (0, 0), (-1, -1), 0.75, border_c),
            ('LINEBELOW', (0, 0), (-1, 0), 2, orange),
            ('BACKGROUND', (0, 1), (-1, 1), white),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('VALIGN', (0, 1), (-1, 1), 'TOP'),
        ]
        if footnote_para is not None:
            ts += [
                ('BACKGROUND', (0, 2), (-1, 2), rl_colors.HexColor('#fafafa')),
                ('TOPPADDING', (0, 2), (-1, 2), 2),
                ('BOTTOMPADDING', (0, 2), (-1, 2), 4),
                ('LEFTPADDING', (0, 2), (-1, 2), 8),
                ('RIGHTPADDING', (0, 2), (-1, 2), 8),
            ]
        t.setStyle(TableStyle(ts))
        return KeepTogether(t)

    def _bar_cell(nb, max_nb, bar_w=3.4 * cm):
        frac   = (nb / max_nb) if max_nb > 0 else 0
        filled = max(frac * bar_w, 1.5)
        empty  = max(bar_w - filled, 0.5)
        b = Table([['', '']], colWidths=[filled, empty], rowHeights=[7])
        b.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (0, 0), orange),
            ('BACKGROUND',    (1, 0), (1, 0), rl_colors.HexColor('#E5E7EB')),
            ('TOPPADDING',    (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ]))
        return b

    def _wide_bar_cell(nb, max_nb, bar_w=10.6 * cm):
        return _bar_cell(nb, max_nb, bar_w=bar_w)

    def _kpi_cell(val, lbl):
        return Paragraph(
            '<font color="white" size="17"><b>%s</b></font><br/><font color="#c5cedd" size="8">%s</font>'
            % (_safe_xml(val), _safe_xml(lbl)),
            ParagraphStyle('kv', alignment=TA_CENTER, spaceAfter=0, leading=16),
        )

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.3 * cm, leftMargin=1.3 * cm,
        topMargin=1.2 * cm, bottomMargin=1.6 * cm,
    )
    elements = []

    logo_path = os.path.join(settings.BASE_DIR, 'static', 'images', 'logo_somatrin.png')
    logo_cell = Paragraph('<font color="#1a2c4e" size="11"><b>SOMATRIN</b></font>', s_center)
    if os.path.exists(logo_path):
        from reportlab.platypus import Image as RLImage
        logo_cell = RLImage(logo_path, width=3 * cm, height=1.2 * cm)

    hdr_tbl = Table([[
        logo_cell,
        Paragraph(
            '<font color="#1a2c4e" size="14"><b>RAPPORT DE SYNTHÈSE</b></font>'
            '<br/><font color="#E87722" size="11">Achats &amp; Approvisionnement</font>',
            s_center),
        Paragraph(
            '<font size="9" color="grey">Généré le<br/><b>%s</b><br/>Document confidentiel</font>'
            % today.strftime('%d/%m/%Y'),
            s_right),
    ]], colWidths=[3.5 * cm, 11 * cm, 3.5 * cm])
    hdr_tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN',  (0, 0), (0,  0), 'LEFT'),
        ('ALIGN',  (1, 0), (1,  0), 'CENTER'),
        ('ALIGN',  (2, 0), (2,  0), 'RIGHT'),
    ]))
    elements.append(hdr_tbl)
    elements.append(Paragraph(
        '<font size="9.5" color="#374151"><b>Périmètre analysé</b> · %s</font>'
        % _safe_xml(period_label_fr),
        ParagraphStyle('subhdr', alignment=TA_CENTER, spaceAfter=4, spaceBefore=2),
    ))
    elements.append(HRFlowable(width='100%', thickness=2, color=orange, spaceAfter=5))

    if not odoo_ok:
        _wb = ParagraphStyle(
            'warnbox', fontSize=9, textColor=navy, spaceAfter=5,
            backColor=rl_colors.HexColor('#fff7ed'), borderPadding=5,
            borderColor=orange, borderWidth=0.5, leftIndent=4, rightIndent=4,
        )
        elements.append(Paragraph(
            '<b><font color="#b45309">Connexion Odoo indisponible.</font></b> '
            'Les indicateurs ci-dessous sont à zéro ou partiels — vérifiez la configuration ERP.',
            _wb,
        ))

    if not multi_ccy:
        montant_str = (
            ('%.1f M %s' % (montant_ttc / 1_000_000, montant_currency_label))
            if montant_ttc >= 1_000_000
            else (format_number_decimals(montant_ttc, 0) + ' ' + montant_currency_label)
        )
    else:
        montant_str = _fmt_ccy_breakdown(montant_by_ccy_id, decimals=0)
    kpi_tbl = Table([[
        _kpi_cell('%d' % total_bons,    'Bons de commande'),
        _kpi_cell(montant_str, 'Montants TTC' if multi_ccy else 'Montant TTC'),
        _kpi_cell('%.1f%%' % taux_conf,  'Taux confirmation'),
        _kpi_cell('%d' % nb_fourn_periode, 'Fournisseurs distincts'),
    ]], colWidths=[4.5 * cm] * 4)
    kpi_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), navy),
        ('TOPPADDING',    (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEAFTER',     (0, 0), (2,  0), 0.5, rl_colors.HexColor('#ffffff33')),
    ]))
    elements.append(kpi_tbl)
    elements.append(Paragraph(
        '<font size="7.5" color="#6B7280">Réf. catalogue Odoo (supplier_rank) : '
        '<b>%s</b> fournisseur(s) marqué(s) actifs — distincts sur le périmètre : <b>%s</b>.</font>'
        % (nb_fourn_catalog, nb_fourn_periode),
        ParagraphStyle('kpf', alignment=TA_CENTER, spaceBefore=2, leading=11),
    ))
    elements.append(Spacer(1, 6))

    _exec_bits = []
    if total_bons:
        _exec_bits.append('%s bon(s) de commande' % format_number_decimals(total_bons, 0))
    if not multi_ccy and montant_ttc > 0:
        _exec_bits.append(
            '%s %s TTC' % (format_number_decimals(montant_ttc, 0), montant_currency_label)
        )
    elif multi_ccy:
        _exec_bits.append('Montants ventilés par devise')
    if nb_fourn_periode:
        _exec_bits.append('%d fournisseur(s) distinct(s)' % nb_fourn_periode)
    if delai_moy_j is not None:
        _exec_bits.append('≈ %.0f j. délai moyen commande → échéance' % delai_moy_j)
    _exec_body = (
        _safe_xml(' · '.join(_exec_bits))
        if _exec_bits else _safe_xml('Aucun bon de commande sur ce périmètre.')
    )
    elements.append(Paragraph(
        '<font color="#1a2c4e" size="12"><b>Synthèse exécutive</b></font><br/><br/>'
        '<font size="10" color="#334155">%s</font>' % _exec_body,
        ParagraphStyle(
            'exec_box', alignment=TA_CENTER, spaceAfter=10, spaceBefore=0,
            leading=15, backColor=rl_colors.HexColor('#eef2ff'),
            borderPadding=14, leftIndent=10, rightIndent=10,
            borderColor=orange, borderWidth=0.7,
        ),
    ))

    if total_bons > 0 and not multi_ccy:
        ticket_moyen = (
            format_number_decimals(montant_ttc / total_bons, 2) + ' ' + montant_currency_label
        )
    else:
        ticket_moyen = '—'

    s_intro = ParagraphStyle(
        'intro_pdf', fontSize=8.5, textColor=gray, alignment=TA_CENTER,
        spaceAfter=5, leading=11,
    )
    _intro_pdf = (
        'Résumé opérationnel — données Odoo (<i>purchase.order</i>, <i>purchase.request</i> si disponible). '
        'Filtres optionnels en URL : <i>periode</i> (mois, trimestre, annee, tout), '
        '<i>date_debut</i>, <i>date_fin</i>, <i>societe</i>. '
        + (' Les montants sont ventilés par devise (sans conversion automatique).' if multi_ccy else '')
    )
    elements.append(Paragraph(_intro_pdf, s_intro))

    bc_note = Paragraph(
        '<i>Retard : date planifiée dépassée (bons ni terminés ni annulés). '
        'Délai moyen = échéance planifiée − date de commande (écarts calendaires).</i>',
        s_note,
    )
    _mont_detail = (
        format_number_decimals(montant_ttc, 2) + ' ' + montant_currency_label
        if not multi_ccy else _fmt_ccy_breakdown(montant_by_ccy_id, decimals=2)
    )
    _delai_txt = (
        ('%.1f jours (%d / %d bons)' % (delai_moy_j, len(delai_plan_jours), total_bons))
        if delai_moy_j is not None and total_bons else '—'
    )
    bc_rows = [
        [_hdr('Indicateur'), _hdr('Valeur')],
        ['Période & filtres', period_label_fr],
        ['Réf. catalogue fournisseurs', '%d (Odoo)' % nb_fourn_catalog],
        ['Total bons', '%d' % total_bons],
        ['Brouillons', '%d' % nb_draft],
        ['Envoyés (sent)', '%d' % nb_sent],
        ['Montant TTC', _mont_detail],
        ['Bons confirmés (purchase/done)', '%d' % confirmes],
        ['Taux confirmation', '%.1f%%' % taux_conf],
        ['Bons en retard', '%d' % en_retard],
        ['Délai moyen commande → planifié', _delai_txt],
        ['Ticket moyen', ticket_moyen],
    ]
    bc_styles_dyn = []
    for ri, row in enumerate(bc_rows):
        if ri == 0:
            continue
        label = row[0]
        if not isinstance(label, str):
            continue
        sl = label
        if sl.startswith('Bons en retard'):
            bc_styles_dyn += [
                ('TEXTCOLOR', (1, ri), (1, ri), red),
                ('FONTNAME', (1, ri), (1, ri), 'Helvetica-Bold'),
            ]
        elif sl.startswith('Taux confirmation'):
            bc_styles_dyn += [
                ('TEXTCOLOR', (1, ri), (1, ri), green),
                ('FONTNAME', (1, ri), (1, ri), 'Helvetica-Bold'),
            ]
    _card_bc = _card('Bons de commande', _pair_table(bc_rows, extra_styles=bc_styles_dyn), bc_note)

    da_note = Paragraph(
        '<i>Source : %s · « Urgent » : montant ou coût estimé &gt; 50&nbsp;000 MAD.</i>'
        % _safe_xml(da_source_label),
        s_note,
    )
    da_rows = [
        [_hdr('Indicateur'), _hdr('Valeur')],
        ['Total demandes', '%d' % total_da],
        ['Montant estimé', format_number_decimals(montant_da, 2) + ' MAD'],
        ['Demandes urgentes', '%d' % urgentes],
        ['RFQ (brouillons)', '%d' % total_rfq],
    ]
    da_styles = [
        ('TEXTCOLOR', (1, 3), (1, 3), red),
        ('FONTNAME', (1, 3), (1, 3), 'Helvetica-Bold'),
    ]
    _card_da = _card("Demandes d'achat", _pair_table(da_rows, extra_styles=da_styles), da_note)

    sv_rows = [
        [_hdr('Statut'), _hdr('Nb bons')],
        ['Réceptionnées', '%d' % nb_recues],
        ['En cours (purchase)', '%d' % nb_en_cours],
        ['Annulées', '%d' % nb_annules],
        ['En brouillon (RFQ)', '%d' % total_rfq],
    ]
    sv_styles = [
        ('TEXTCOLOR', (1, 1), (1, 1), green),
        ('FONTNAME', (1, 1), (1, 1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (1, 3), (1, 3), red),
        ('FONTNAME', (1, 3), (1, 3), 'Helvetica-Bold'),
    ]
    _card_sv = _card('Suivi livraisons', _pair_table(sv_rows, extra_styles=sv_styles), None)

    # Ne pas envelopper les KeepTogether (_card) dans un Table parent : ReportLab calcule
    # alors une hauteur aberrante (LayoutError page 2).
    elements.append(_card_bc)
    elements.append(Spacer(1, 3))
    elements.append(_card_da)
    elements.append(Spacer(1, 3))
    elements.append(_card_sv)

    elements.append(Spacer(1, 8))
    elements.append(PageBreak())

    if total_bons > 0:
        _state_specs = [
            ('draft', 'Brouillon'),
            ('sent', 'Envoyé'),
            ('purchase', 'Confirmé'),
            ('done', 'Terminé / réceptionné'),
            ('cancel', 'Annulé'),
        ]
        _mx_sb = max(
            sum(1 for b in bons if b.get('state') == code)
            for code, _ in _state_specs
        ) or 1
        st_rows = [[
            _hdr('Statut (Odoo)'), _hdr('Nombre'), _hdr('Part'), _hdr('Répartition'),
        ]]
        for code, lbl_fr in _state_specs:
            ns = sum(1 for b in bons if b.get('state') == code)
            pct = (100.0 * ns / total_bons)
            st_rows.append([
                lbl_fr, str(ns), '%.1f%%' % pct, _wide_bar_cell(ns, _mx_sb),
            ])
        st_tbl = Table(
            st_rows,
            colWidths=[4 * cm, 2 * cm, 2 * cm, 10 * cm],
            repeatRows=1,
        )
        st_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), navy),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (1, 1), (2, -1), 'CENTER'),
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, border_c),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, light]),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        elements.append(KeepTogether([
            _section_bar('Répartition quantitative par statut'),
            Spacer(1, 3),
            st_tbl,
        ]))
        elements.append(Spacer(1, 8))

    if len(company_agg) > 1:
        co_sorted = sorted(company_agg.items(), key=lambda kv: -kv[1]['nb'])[:12]
        co_rows = [[
            _hdr('Société'), _hdr('Nb bons'), _hdr('Montant TTC (périmètre)'),
        ]]
        for _cid, cg in co_sorted:
            amt_txt = (
                _fmt_ccy_breakdown(cg['by_ccy'], decimals=0)
                if multi_ccy else (
                    format_number_decimals(sum(cg['by_ccy'].values()), 0) + ' ' + montant_currency_label
                )
            )
            co_rows.append([cg['name'], str(cg['nb']), amt_txt])
        co_tbl = Table(co_rows, colWidths=[7 * cm, 3 * cm, 8 * cm], repeatRows=1)
        co_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), navy),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (1, 1), (1, -1), 'CENTER'),
            ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, border_c),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, light]),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        elements.append(KeepTogether([
            _section_bar('Répartition par société'),
            Spacer(1, 3),
            co_tbl,
        ]))
        elements.append(Spacer(1, 8))

    f_col_w    = [1 * cm, 5.5 * cm, 3.5 * cm, 2 * cm, 6 * cm]
    max_bons_f = top_fourn[0][1] if top_fourn else 1

    _hdr_montant = (
        'Montant TTC (MAD)' if not multi_ccy and montant_currency_label.strip().upper() == 'MAD'
        else ('Montant TTC' if multi_ccy else ('Montant TTC (%s)' % montant_currency_label))
    )
    fourn_rows = [[
        _hdr('#'), _hdr('Fournisseur'), _hdr('Progression'), _hdr('Bons'), _hdr(_hdr_montant),
    ]]
    fourn_styles = [
        ('BACKGROUND',    (0, 0), (-1, 0), navy),
        ('FONTNAME',      (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0), (-1, -1), 9),
        ('ALIGN',         (0, 0), (0, -1), 'CENTER'),
        ('ALIGN',         (3, 0), (3, -1), 'RIGHT'),
        ('ALIGN',         (4, 0), (4, -1), 'RIGHT'),
        ('GRID',          (0, 0), (-1, -1), 0.5, border_c),
        ('TOPPADDING',    (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ('LEFTPADDING',   (2, 1), (2, -2), 0),
        ('RIGHTPADDING',  (2, 1), (2, -2), 0),
        ('TOPPADDING',    (2, 1), (2, -2), 6),
    ]

    if not top_fourn:
        fourn_rows.append([
            Paragraph('<i>Aucun fournisseur avec bons sur les données extraites.</i>', styles['Normal']),
            '', '', '', '',
        ])
        fourn_styles.append(('SPAN', (0, 1), (4, 1)))
        fourn_styles.append(('ALIGN', (0, 1), (4, 1), 'CENTER'))
    for i, (nom, nb, montant_disp, _rawv) in enumerate(top_fourn, 1):
        nom_d = (nom[:36] + '...') if len(nom) > 36 else nom
        fourn_rows.append([
            str(i),
            Paragraph(_safe_xml(nom_d), ParagraphStyle('fn', fontSize=9, leading=11)),
            _bar_cell(nb, max_bons_f),
            '%d' % nb,
            Paragraph(_safe_xml(montant_disp), ParagraphStyle('fa', fontSize=9, alignment=TA_RIGHT)),
        ])

    tot_bons_top = sum(n for _, n, _, _ in top_fourn)
    tot_montant_top = None if multi_ccy else sum(
        (m for _, _, _, m in top_fourn if m is not None), 0.0
    )
    total_row_idx   = len(fourn_rows)
    fourn_rows.append([
        Paragraph('<b><font color="white">-</font></b>', styles['Normal']),
        Paragraph('<b><font color="white">TOTAL Top 5</font></b>', styles['Normal']),
        '',
        Paragraph('<b><font color="white">%d</font></b>' % tot_bons_top,
                  ParagraphStyle('tr',  fontSize=9, alignment=TA_RIGHT)),
        Paragraph(
            '<b><font color="white">%s</font></b>' % (
                format_number_decimals(tot_montant_top, 0) + ' ' + montant_currency_label
                if tot_montant_top is not None
                else '—'
            ),
            ParagraphStyle('tr2', fontSize=9, alignment=TA_RIGHT)),
    ])
    fourn_styles.append(('BACKGROUND', (0, total_row_idx), (-1, total_row_idx), navy))

    for i, _ in enumerate(top_fourn, 1):
        if i == 1:
            fourn_styles += [
                ('BACKGROUND', (0, i), (-1, i), orange),
                ('TEXTCOLOR',  (0, i), (-1, i), white),
                ('FONTNAME',   (0, i), (-1, i), 'Helvetica-Bold'),
            ]
        elif i in (2, 3):
            fourn_styles.append(('BACKGROUND', (0, i), (-1, i), light_blue))

    fourn_tbl = Table(fourn_rows, colWidths=f_col_w, splitByRow=True, repeatRows=1)
    fourn_tbl.setStyle(TableStyle(fourn_styles))
    elements.append(KeepTogether([
        _section_bar('Répartition par fournisseur (Top 5)'),
        Spacer(1, 3),
        fourn_tbl,
    ]))
    elements.append(Spacer(1, 6))

    colour_map = {'ok': '#16a34a', 'warn': '#dc2626', 'alert': '#E87722'}
    ana_rows = []
    for kind, line in analyse:
        c = colour_map.get(kind, '#374151')
        ana_rows.append([
            Paragraph(
                '<font color="%s" size="10"><b>•</b> %s</font>' % (c, _safe_xml(line)),
                ParagraphStyle('al', fontSize=10, leading=13, spaceAfter=2),
            ),
        ])

    ana_tbl = Table(ana_rows, colWidths=[PAGE_W], splitByRow=True)
    ana_tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), light_blue),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
        ('LEFTPADDING',  (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('LINEBEFORE',   (0, 0), (0,  -1), 4, orange),
        ('BOX',          (0, 0), (-1, -1), 0.5, border_c),
        ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(KeepTogether([
        _section_bar('Analyse automatique'),
        Spacer(1, 3),
        ana_tbl,
    ]))

    doc.build(elements, canvasmaker=_NumberedCanvas)
    pdf = buffer.getvalue()
    buffer.close()

    response = HttpResponse(pdf, content_type='application/pdf')
    response['Content-Disposition'] = (
        'inline; filename="synthese_achats_%s_%s.pdf"'
        % (period_slug, today.strftime('%Y%m%d'))
    )
    return response


@csrf_exempt
def chatbot_api(request):
    """
    SOMA AI — API JSON. Historique en session serveur (non falsifiable par le client).
    Moteur : Ollama local uniquement si SOMA_AI_ENABLED (pas de cloud par défaut).

    Contrat attendu par le front (base.html) :
    - Succès : {"status": "ok", "response": "<texte>", "mode": "..."}
    - Erreur  : {"status": "error", "error": "<texte>", "code": "..."} (+ HTTP 4xx/5xx)

    Authentification : JSON 401 (pas de redirection HTML) pour que fetch() puisse parser la réponse
    (évite « Réponse serveur invalide » si @login_required renvoyait une page de connexion HTML).
    """
    from reporting.services import soma_ai

    if not request.user.is_authenticated:
        return JsonResponse(
            {
                'status': 'error',
                'error': 'Authentification requise. Reconnectez-vous puis réessayez.',
                'code': 'auth',
            },
            status=401,
        )

    if request.method == 'GET':
        data = soma_ai.get_public_status(request)
        data['usage'] = (
            'POST JSON {"message":"...", "reset": false, "context": null|object} — '
            'reset=true vide la conversation ; context (optionnel) = indicateurs écran validés côté serveur.'
        )
        return JsonResponse(data)

    if request.method != 'POST':
        return JsonResponse({'status': 'error', 'error': 'Méthode non autorisée'}, status=405)

    try:
        payload = json.loads(request.body.decode('utf-8') or '{}')
    except (json.JSONDecodeError, UnicodeDecodeError):
        return JsonResponse({
            'status': 'error',
            'error': 'Corps JSON invalide.',
        }, status=400)

    reset = payload.get('reset') is True
    if reset:
        soma_ai.session_reset(request)

    raw_msg = payload.get('message', '')
    if not isinstance(raw_msg, str):
        raw_msg = ''
    msg = raw_msg.strip()

    if not msg:
        if reset:
            return JsonResponse({
                'status': 'ok',
                'response': 'Conversation côté serveur réinitialisée.',
                'mode': 'reset',
            })
        if not getattr(settings, 'SOMA_AI_ENABLED', False):
            return JsonResponse({
                'status': 'ok',
                'response': soma_ai.ERROR_MESSAGES_FR.get(
                    'soma_ai_desactive', 'SOMA AI est désactivé.'
                ),
                'mode': 'disabled',
            })
        return JsonResponse({'status': 'error', 'error': 'Message vide.'}, status=400)

    if not getattr(settings, 'SOMA_AI_ENABLED', False):
        return JsonResponse({
            'status': 'ok',
            'response': soma_ai.ERROR_MESSAGES_FR.get(
                'soma_ai_desactive', 'SOMA AI est désactivé.'
            ),
            'mode': 'disabled',
        })

    dash_ctx = soma_ai.sanitize_dashboard_context(payload.get('context'))
    dash_ctx = soma_ai.merge_session_achats_bc_kpis(request, dash_ctx)
    reply, err = soma_ai.run_turn(request, raw_msg, dash_ctx)
    if err:
        text = soma_ai.ERROR_MESSAGES_FR.get(err, 'Erreur SOMA AI.')
        status = 429 if err == 'trop_de_requetes' else 503
        if err in ('message_vide', 'message_trop_long', 'message_invalide'):
            status = 400
        return JsonResponse({'status': 'error', 'error': text, 'code': err}, status=status)

    return JsonResponse({'status': 'ok', 'response': reply, 'mode': 'local_ollama'})


@csrf_exempt
def chat_soma_ai(request):
    """
    SOMA AI v2 — endpoint JSON avec données Odoo EN TEMPS RÉEL.
    Contrat: POST {"message": "..."} → {"success": true, "response": "...", "timestamp": "..."}
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'POST required'}, status=400)

    try:
        data = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse({'error': 'JSON invalide'}, status=400)

    question = data.get('message', '').strip()
    if not question:
        return JsonResponse({'error': 'Message vide'}, status=400)
    if len(question) > 500:
        return JsonResponse({'error': 'Message trop long (max 500 chars)'}, status=400)

    # Données de la page courante envoyées par le front (KPIs visibles par l'utilisateur)
    page_data = data.get('page_data') or data.get('context') or {}
    if not isinstance(page_data, dict):
        page_data = {}
    # Limite de taille pour éviter les injections de grandes charges utiles
    if len(str(page_data)) > 4000:
        page_data = {}

    # Rate limiting: 15 requêtes/minute par utilisateur ou IP
    rate_key = (
        f"soma_ai_v2_{request.user.id}"
        if request.user.is_authenticated
        else f"soma_ai_v2_{request.META.get('REMOTE_ADDR', 'unknown')}"
    )
    req_count = cache.get(rate_key, 0)
    if req_count >= 15:
        return JsonResponse({'error': '⏱️ Trop de requêtes. Attendez 1 minute.'}, status=429)
    cache.set(rate_key, req_count + 1, 60)

    logger.info(
        '🤖 SOMA AI v2: question de %s (%s chars)',
        request.user or request.META.get('REMOTE_ADDR'),
        len(question),
    )

    # Contexte de page: déduire depuis le Referer HTTP pour réponses ciblées
    referrer = request.META.get('HTTP_REFERER', '')
    if 'transport/rentabilite' in referrer or ('transport' in referrer and 'rentabilite' in referrer):
        page_context = 'transport'
    elif 'production/rentabilite' in referrer or ('production' in referrer and 'rentabilite' in referrer):
        page_context = 'production'
    elif 'transport' in referrer:
        page_context = 'transport'
    elif 'production' in referrer:
        page_context = 'production'
    elif 'bons-commande' in referrer or 'demandes-achat' in referrer or 'achats' in referrer:
        page_context = 'achats'
    elif 'parc' in referrer or 'maintenance' in referrer:
        page_context = 'parc'
    elif 'gasoil' in referrer or 'qhse' in referrer:
        page_context = 'gasoil' if 'gasoil' in referrer else 'qhse'
    else:
        page_context = 'general'

    logger.info('🌐 SOMA AI v2: page_context=%s (referrer=%s)', page_context, referrer[:80] if referrer else '—')

    try:
        odoo_enabled = getattr(settings, 'SOMA_AI_ODOO_ENABLED', False)

        if not odoo_enabled:
            # Mode dégradé: pas de requêtes Odoo → afficher les données de page si disponibles
            if page_data:
                kpis = page_data.get('kpis') or page_data
                lines = ['📊 **Données actuelles de la page**\n']
                for k, v in kpis.items():
                    if v not in (None, '', False):
                        label = str(k).replace('_', ' ').title()
                        lines.append(f'• {label}: **{v}**')
                lines.append('\n_SOMA AI est en mode limité — les analyses IA sont temporairement désactivées._')
                response_text = '\n'.join(lines)
            else:
                response_text = (
                    'SOMA AI est actuellement en mode limité (analyses IA désactivées).\n'
                    'Les données de la page sont disponibles dans le tableau de bord.'
                )
        else:
            connector = OdooConnector()
            metier = MetierDataService(connector)
            engine = SomaAIEngine(connector, metier)
            response_text = engine.process_question(question, page_context, page_data or None)

        logger.info('✅ Réponse SOMA AI v2 générée (%s chars, odoo=%s)', len(response_text), odoo_enabled)
        return JsonResponse({
            'success': True,
            'response': response_text,
            'timestamp': datetime.now().isoformat(),
        })
    except Exception as exc:
        logger.error('❌ Erreur chat_soma_ai: %s', exc, exc_info=True)
        return JsonResponse({'success': False, 'error': f'Erreur serveur: {exc}'}, status=500)


# ═══════════════════════════════════════════════════════════════
#  FINANCE & COMPTABILITÉ — Module complet
# ═══════════════════════════════════════════════════════════════

from reporting.services import finance_service as _fs

_FINANCE_MENU = [
    {'key': 'dashboard',             'label': 'Dashboard',             'url': 'finance_dashboard',             'icon': 'bi-speedometer2',          'desc': 'Vue consolidée des flux financiers et KPIs.'},
    {'key': 'factures_clients',      'label': 'Factures clients',      'url': 'finance_factures_clients',      'icon': 'bi-file-earmark-person',    'desc': 'Suivi des factures émises aux clients.'},
    {'key': 'factures_fournisseurs', 'label': 'Factures fournisseurs', 'url': 'finance_factures_fournisseurs', 'icon': 'bi-receipt-cutoff',          'desc': 'Gestion des factures fournisseurs.'},
    {'key': 'avoirs',                'label': 'Avoirs',                'url': 'finance_avoirs',                'icon': 'bi-arrow-counterclockwise',  'desc': 'Avoirs clients et fournisseurs.'},
    {'key': 'paiements',             'label': 'Paiements',             'url': 'finance_paiements',             'icon': 'bi-credit-card',             'desc': 'Encaissements et décaissements.'},
    {'key': 'rapports',              'label': 'Rapports & Analyses',   'url': 'finance_rapports',              'icon': 'bi-bar-chart-line',          'desc': 'Cash flow, impayés, ratios financiers.'},
    {'key': 'configuration',         'label': 'Configuration',         'url': 'finance_configuration',         'icon': 'bi-gear',                    'desc': 'Paramètres généraux et journaux.'},
]

_CARD_COLORS = ['fc-blue', 'fc-orange', 'fc-green', 'fc-purple', 'fc-sky', 'fc-amber', 'fc-red']


def _fin_ctx(active_key, extra=None):
    ctx = {'fin_menu': _FINANCE_MENU, 'fin_active': active_key,
           'card_colors': _CARD_COLORS}
    if extra:
        ctx.update(extra)
    return ctx


def _finance_has_access(user):
    if not user.is_authenticated:
        return False
    if user.is_superuser or user.is_staff:
        return True
    return user.groups.filter(name__in=['finance', 'pilotage']).exists()


@login_required
def finance_dashboard(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    kpis = {}
    chart_ca_json = chart_vs_json = chart_creances_json = chart_dettes_json = '{}'
    chart_enc_json = chart_delai_json = '{}'
    chart_ca_labels = chart_ca_data = chart_vs_enc = chart_vs_fac = '[]'
    stats = {}
    overdue_30j = overdue_60j = health_score = delai_moyen = 0
    health_label = 'N/A'
    health_color = '#64748b'
    health_desc  = ''
    try:
        cache_key = f'fin_dash_{request.user.id}'
        cached = cache.get(cache_key)
        if cached:
            return render(request, 'finance/dashboard.html', _fin_ctx('dashboard', cached))
        uid, models = get_odoo_connection()
        inv_c  = _fs.fetch_invoices(uid, models, 'out_invoice', limit=2000)
        inv_f  = _fs.fetch_invoices(uid, models, 'in_invoice',  limit=2000)
        pays   = _fs.fetch_payments(uid, models, limit=2000)
        kpis   = _fs.compute_kpis(inv_c, inv_f, pays)

        # Overdue brackets (>30j, >60j)
        overdue_30j = sum(1 for r in inv_c if r.get('is_overdue') and r.get('days_overdue', 0) > 30)
        overdue_60j = sum(1 for r in inv_c if r.get('is_overdue') and r.get('days_overdue', 0) > 60)

        # Délai moyen paiement convenu (date_due - date)
        _dv = []
        for _r in inv_c:
            if _r.get('state') == 'posted' and len(_r.get('date_iso', '')) >= 10 and len(_r.get('date_due_iso', '')) >= 10:
                try:
                    _d = (datetime.strptime(_r['date_due_iso'], '%Y-%m-%d').date()
                          - datetime.strptime(_r['date_iso'], '%Y-%m-%d').date()).days
                    if 0 <= _d <= 180:
                        _dv.append(_d)
                except Exception:
                    pass
        delai_moyen = int(round(sum(_dv) / len(_dv), 0)) if _dv else 0

        # Score santé financière composite (0–100)
        _s = (min(kpis.get('pct_encaisse', 0), 100) / 100.0 * 40
              + min(kpis.get('liquidite', 0), 200) / 200.0 * 30
              + max(0.0, min(kpis.get('marge', 0), 50)) / 50.0 * 20
              - (kpis.get('nb_overdue', 0) / max(1, kpis.get('nb_creances', 1))) * 10)
        health_score = int(max(0, min(100, round(_s))))
        if health_score >= 80:
            health_label, health_color = 'EXCELLENT', '#1a2c4e'
        elif health_score >= 60:
            health_label, health_color = 'BON', '#E87722'
        elif health_score >= 40:
            health_label, health_color = 'ATTENTION', '#C85E0F'
        else:
            health_label, health_color = 'CRITIQUE', '#7A3800'
        health_desc = {
            'EXCELLENT': 'Encaissements et liquidité au top.',
            'BON': 'Bonne santé financière, surveiller les retards.',
            'ATTENTION': "Des indicateurs nécessitent attention.",
            'CRITIQUE': 'Situation critique — action urgente requise.',
        }[health_label]

        # Chart 1: CA mensuel 12 mois
        ca_monthly = _fs.monthly_amounts(
            [r for r in inv_c if r['state'] == 'posted'], 'amount_ht', 12)
        chart_ca_json = json.dumps(ca_monthly)

        # Chart 2: Factures vs Paiements 6 derniers mois
        fac_monthly = _fs.monthly_amounts(
            [r for r in inv_c if r['state'] == 'posted'], 'amount_ht', 6)
        pay_monthly = _fs.monthly_amounts(
            [r for r in pays if r['payment_type'] == 'inbound'], 'amount', 6)
        all_months = sorted(set(list(fac_monthly.keys()) + list(pay_monthly.keys())),
                            key=_fs._sort_key)[-6:]
        chart_vs_json = json.dumps({
            'labels': all_months,
            'factures': [fac_monthly.get(m, 0) for m in all_months],
            'paiements': [pay_monthly.get(m, 0) for m in all_months],
        })

        # Charts 3 & 4: créances / dettes breakdown
        chart_creances_json = json.dumps(_fs.payment_state_breakdown(inv_c))
        chart_dettes_json   = json.dumps(_fs.payment_state_breakdown(inv_f))

        # Chart 5: Taux d'encaissement mensuel (%)
        fac_m12 = _fs.monthly_amounts([r for r in inv_c if r['state'] == 'posted'], 'amount_ttc', 12)
        pay_m12 = _fs.monthly_amounts([r for r in pays if r['payment_type'] == 'inbound'], 'amount', 12)
        _enc_months = sorted(set(list(fac_m12) + list(pay_m12)), key=_fs._sort_key)[-12:]
        chart_enc_json = json.dumps({
            'labels': _enc_months,
            'data': [round(pay_m12.get(m, 0) / fac_m12.get(m, 1) * 100, 1) if fac_m12.get(m) else 0
                     for m in _enc_months],
        })

        # Chart 6: Délai moyen paiement par mois vs cible 30j
        _delai_m = {}
        for _r in inv_c:
            if _r.get('state') == 'posted' and len(_r.get('date_iso', '')) >= 10 and len(_r.get('date_due_iso', '')) >= 10:
                try:
                    _d = (datetime.strptime(_r['date_due_iso'], '%Y-%m-%d').date()
                          - datetime.strptime(_r['date_iso'], '%Y-%m-%d').date()).days
                    if 0 <= _d <= 180:
                        _mk = f"{_r['date_iso'][5:7]}/{_r['date_iso'][0:4]}"
                        _delai_m.setdefault(_mk, []).append(_d)
                except Exception:
                    pass
        _dm_months = sorted(set(_enc_months + list(_delai_m.keys())), key=_fs._sort_key)[-6:]
        chart_delai_json = json.dumps({
            'labels': _dm_months,
            'data': [int(round(sum(_delai_m[m]) / len(_delai_m[m]), 0)) if _delai_m.get(m) else 0
                     for m in _dm_months],
            'target': 30,
        })

        stats = {
            'nb_clients':      kpis['nb_clients'],
            'nb_fournisseurs': kpis['nb_fournisseurs'],
            'ticket_moyen':    kpis['ticket_moyen'],
            'nb_fac_c':        kpis['nb_factures_clients'],
            'nb_fac_f':        kpis['nb_factures_fournisseurs'],
            'delai_moyen':     delai_moyen,
        }
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'

    payload = {
        'error': error,
        'kpis': kpis,
        'chart_ca_json':       chart_ca_json,
        'chart_vs_json':       chart_vs_json,
        'chart_creances_json': chart_creances_json,
        'chart_dettes_json':   chart_dettes_json,
        'chart_enc_json':      chart_enc_json,
        'chart_delai_json':    chart_delai_json,
        'stats': stats,
        'overdue_30j':   overdue_30j,
        'overdue_60j':   overdue_60j,
        'health_score':  health_score,
        'health_label':  health_label,
        'health_color':  health_color,
        'health_desc':   health_desc,
        'delai_moyen':   delai_moyen,
    }
    if not error:
        cache.set(cache_key, payload, 120)
    return render(request, 'finance/dashboard.html', _fin_ctx('dashboard', payload))


def _apply_invoice_filters(rows, request):
    q    = (request.GET.get('q') or '').lower().strip()
    df   = (request.GET.get('date_from') or '')
    dt   = (request.GET.get('date_to') or '')
    ps   = (request.GET.get('payment_state') or '')
    pmin = request.GET.get('amount_min') or ''
    pmax = request.GET.get('amount_max') or ''
    if q:
        rows = [r for r in rows if q in r['name'].lower() or q in r['partner'].lower()]
    if df:
        rows = [r for r in rows if r['date_iso'] >= df]
    if dt:
        rows = [r for r in rows if r['date_iso'] <= dt]
    if ps:
        rows = [r for r in rows if r['payment_state'] == ps]
    if pmin:
        try:
            rows = [r for r in rows if r['amount_ttc'] >= float(pmin)]
        except ValueError:
            pass
    if pmax:
        try:
            rows = [r for r in rows if r['amount_ttc'] <= float(pmax)]
        except ValueError:
            pass
    return rows


@login_required
def finance_factures_clients(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    rows = []
    chart_top_json = chart_month_json = '{}'
    try:
        uid, models = get_odoo_connection()
        all_rows = _fs.fetch_invoices(uid, models, 'out_invoice', limit=2000)
        rows = _apply_invoice_filters(all_rows, request)
        chart_top_json   = json.dumps(_fs.top_partners(rows, 10))
        chart_month_json = json.dumps(_fs.monthly_amounts(
            [r for r in rows if r['state'] == 'posted'], 'amount_ht', 12))
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'

    posted = [r for r in rows if r['state'] == 'posted']
    return render(request, 'finance/factures_clients.html', _fin_ctx('factures_clients', {
        'error': error,
        'rows':  rows,
        'kpi_total':     len(rows),
        'kpi_ht':        round(sum(r['amount_ht']  for r in posted), 0),
        'kpi_ttc':       round(sum(r['amount_ttc'] for r in posted), 0),
        'kpi_impaye':    round(sum(r['amount_residual'] for r in posted if r['payment_state'] != 'paid'), 0),
        'chart_top_json':   chart_top_json,
        'chart_month_json': chart_month_json,
        'q': request.GET.get('q', ''),
        'date_from': request.GET.get('date_from', ''),
        'date_to':   request.GET.get('date_to', ''),
        'payment_state': request.GET.get('payment_state', ''),
        'amount_min': request.GET.get('amount_min', ''),
        'amount_max': request.GET.get('amount_max', ''),
    }))


@login_required
def finance_factures_fournisseurs(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    rows = []
    chart_top_json = chart_month_json = '{}'
    try:
        uid, models = get_odoo_connection()
        all_rows = _fs.fetch_invoices(uid, models, 'in_invoice', limit=2000)
        rows = _apply_invoice_filters(all_rows, request)
        chart_top_json   = json.dumps(_fs.top_partners(rows, 10))
        chart_month_json = json.dumps(_fs.monthly_amounts(
            [r for r in rows if r['state'] == 'posted'], 'amount_ht', 12))
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'

    posted = [r for r in rows if r['state'] == 'posted']
    return render(request, 'finance/factures_fournisseurs.html', _fin_ctx('factures_fournisseurs', {
        'error': error,
        'rows':  rows,
        'kpi_total':  len(rows),
        'kpi_ht':     round(sum(r['amount_ht']  for r in posted), 0),
        'kpi_ttc':    round(sum(r['amount_ttc'] for r in posted), 0),
        'kpi_impaye': round(sum(r['amount_residual'] for r in posted if r['payment_state'] != 'paid'), 0),
        'chart_top_json':   chart_top_json,
        'chart_month_json': chart_month_json,
        'q': request.GET.get('q', ''),
        'date_from': request.GET.get('date_from', ''),
        'date_to':   request.GET.get('date_to', ''),
        'payment_state': request.GET.get('payment_state', ''),
        'amount_min': request.GET.get('amount_min', ''),
        'amount_max': request.GET.get('amount_max', ''),
    }))


@login_required
def finance_avoirs(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    rows = []
    try:
        uid, models = get_odoo_connection()
        avoir_c = _fs.fetch_invoices(uid, models, 'out_refund', limit=1000)
        avoir_f = _fs.fetch_invoices(uid, models, 'in_refund',  limit=1000)
        for r in avoir_c:
            r['avoir_type'] = 'Client'
        for r in avoir_f:
            r['avoir_type'] = 'Fournisseur'
        rows = avoir_c + avoir_f
        rows.sort(key=lambda r: r['date_iso'], reverse=True)
        q  = (request.GET.get('q') or '').lower()
        df = request.GET.get('date_from') or ''
        dt = request.GET.get('date_to')   or ''
        at = request.GET.get('avoir_type') or ''
        if q:  rows = [r for r in rows if q in r['name'].lower() or q in r['partner'].lower()]
        if df: rows = [r for r in rows if r['date_iso'] >= df]
        if dt: rows = [r for r in rows if r['date_iso'] <= dt]
        if at: rows = [r for r in rows if r['avoir_type'] == at]
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'

    return render(request, 'finance/avoirs.html', _fin_ctx('avoirs', {
        'error': error,
        'rows':  rows,
        'kpi_total':  len(rows),
        'kpi_montant': round(sum(r['amount_ttc'] for r in rows), 0),
        'q': request.GET.get('q', ''),
        'date_from': request.GET.get('date_from', ''),
        'date_to':   request.GET.get('date_to', ''),
        'avoir_type': request.GET.get('avoir_type', ''),
    }))


@login_required
def finance_paiements(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    rows = []
    chart_mois_json = chart_mode_json = '{}'
    try:
        uid, models = get_odoo_connection()
        all_rows = _fs.fetch_payments(uid, models, limit=2000)
        raw_rows = list(all_rows)

        # filters
        q  = (request.GET.get('q') or '').lower()
        df = request.GET.get('date_from') or ''
        dt = request.GET.get('date_to')   or ''
        pt = request.GET.get('payment_type') or ''
        pm = request.GET.get('payment_method') or ''
        if q:  all_rows = [r for r in all_rows if q in r['name'].lower() or q in r['partner'].lower()]
        if df: all_rows = [r for r in all_rows if r['date_iso'] >= df]
        if dt: all_rows = [r for r in all_rows if r['date_iso'] <= dt]
        if pt: all_rows = [r for r in all_rows if r['payment_type'] == pt]
        if pm: all_rows = [r for r in all_rows if pm.lower() in r['payment_method'].lower()]
        rows = all_rows

        # chart: received vs sent per month (12m)
        rec_m = _fs.monthly_amounts([r for r in rows if r['payment_type'] == 'inbound'],  'amount', 12)
        out_m = _fs.monthly_amounts([r for r in rows if r['payment_type'] == 'outbound'], 'amount', 12)
        all_m = sorted(set(list(rec_m) + list(out_m)), key=_fs._sort_key)[-12:]
        chart_mois_json = json.dumps({
            'labels':   all_m,
            'recus':    [rec_m.get(m, 0) for m in all_m],
            'effectues': [out_m.get(m, 0) for m in all_m],
        })
        chart_mode_json = json.dumps(_fs.payment_method_breakdown(rows))
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'
        raw_rows = []

    total_recu     = round(sum(r['amount'] for r in rows if r['payment_type'] == 'inbound'), 0)
    total_effectue = round(sum(r['amount'] for r in rows if r['payment_type'] == 'outbound'), 0)
    total_recu_global = round(sum(r['amount'] for r in raw_rows if r.get('payment_type') == 'inbound'), 0)
    total_effectue_global = round(sum(r['amount'] for r in raw_rows if r.get('payment_type') == 'outbound'), 0)
    return render(request, 'finance/paiements.html', _fin_ctx('paiements', {
        'error': error,
        'rows':  rows,
        'kpi_total':     round(sum(r['amount'] for r in rows), 0),
        'kpi_recu':      total_recu,
        'kpi_effectue':  total_effectue,
        'kpi_recu_global': total_recu_global,
        'kpi_effectue_global': total_effectue_global,
        'kpi_recu_count': sum(1 for r in rows if r.get('payment_type') == 'inbound'),
        'kpi_effectue_count': sum(1 for r in rows if r.get('payment_type') == 'outbound'),
        'kpi_has_filters': bool(request.GET),
        'chart_mois_json': chart_mois_json,
        'chart_mode_json': chart_mode_json,
        'q':               request.GET.get('q', ''),
        'date_from':       request.GET.get('date_from', ''),
        'date_to':         request.GET.get('date_to', ''),
        'payment_type':    request.GET.get('payment_type', ''),
        'payment_method':  request.GET.get('payment_method', ''),
        'rapprochement':   request.GET.get('rapprochement', ''),
    }))


@login_required
def finance_rapports(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    cash_flow = []
    overdue_rows = []
    dettes_rows  = []
    chart_cf_json = '{}'
    try:
        uid, models = get_odoo_connection()
        inv_c = _fs.fetch_invoices(uid, models, 'out_invoice', limit=2000)
        inv_f = _fs.fetch_invoices(uid, models, 'in_invoice',  limit=2000)
        pays  = _fs.fetch_payments(uid, models, limit=2000)

        # Cash flow: monthly
        rec_m = _fs.monthly_amounts([r for r in pays if r['payment_type'] == 'inbound'],  'amount', 12)
        out_m = _fs.monthly_amounts([r for r in pays if r['payment_type'] == 'outbound'], 'amount', 12)
        all_m = sorted(set(list(rec_m) + list(out_m)), key=_fs._sort_key)
        solde_cum = 0
        for m in all_m:
            enc  = rec_m.get(m, 0)
            dec  = out_m.get(m, 0)
            flux = enc - dec
            solde_cum += flux
            cash_flow.append({'mois': m, 'encaiss': enc, 'decaiss': dec,
                               'flux': round(flux, 0), 'solde_cum': round(solde_cum, 0)})
        chart_cf_json = json.dumps({k: round(v['flux'], 0) for k, v in
                                     {r['mois']: r for r in cash_flow}.items()})

        # Overdue invoices (clients)
        overdue_rows = sorted(
            [r for r in inv_c if r['is_overdue'] and r['state'] == 'posted'],
            key=lambda x: x['days_overdue'], reverse=True)

        # Unpaid vendor bills
        dettes_rows = sorted(
            [r for r in inv_f if r['state'] == 'posted' and r['payment_state'] != 'paid'],
            key=lambda x: x['amount_residual'], reverse=True)

        # Top clients CA
        top_clients = sorted(
            _fs.top_partners([r for r in inv_c if r['state'] == 'posted'], 10).items(),
            key=lambda x: x[1], reverse=True)

        # Top fournisseurs
        top_fournisseurs = sorted(
            _fs.top_partners([r for r in inv_f if r['state'] == 'posted'], 10).items(),
            key=lambda x: x[1], reverse=True)

        kpis_r = _fs.compute_kpis(inv_c, inv_f, pays)

    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'
        top_clients = top_fournisseurs = []
        kpis_r = {}

    # Flags d'affichage: masquer les blocs vides pour garder une mise en page propre.
    has_cash_flow = any(abs(float(r.get('flux') or 0)) > 0 for r in cash_flow) if cash_flow else False
    top_impayes_data = {}
    for r in overdue_rows:
        partner = r.get('partner') or '—'
        top_impayes_data[partner] = top_impayes_data.get(partner, 0) + float(r.get('amount_residual') or 0)
    # Ne garder que les partenaires avec un reste strictement positif.
    top_impayes_data = {k: v for k, v in top_impayes_data.items() if v > 0}
    has_top_impayes_chart = len(top_impayes_data) > 0
    has_overdue_table = len(overdue_rows) > 0
    has_dettes_table = len(dettes_rows) > 0
    has_top_clients = len(top_clients) > 0
    has_top_fournisseurs = len(top_fournisseurs) > 0

    return render(request, 'finance/rapports.html', _fin_ctx('rapports', {
        'error': error,
        'cash_flow':      cash_flow,
        'overdue_rows':   overdue_rows[:100],
        'dettes_rows':    dettes_rows[:100],
        'top_clients':    top_clients[:10],
        'top_fournisseurs': top_fournisseurs[:10],
        'kpis': kpis_r,
        'chart_cf_json': chart_cf_json,
        'nb_overdue':    len(overdue_rows),
        'mt_overdue':    round(sum(r['amount_residual'] for r in overdue_rows), 0),
        'nb_dettes':     len(dettes_rows),
        'mt_dettes':     round(sum(r['amount_residual'] for r in dettes_rows), 0),
        'has_cash_flow': has_cash_flow,
        'top_impayes_json': json.dumps(top_impayes_data),
        'has_top_impayes_chart': has_top_impayes_chart,
        'has_overdue_table': has_overdue_table,
        'has_dettes_table': has_dettes_table,
        'has_top_clients': has_top_clients,
        'has_top_fournisseurs': has_top_fournisseurs,
    }))


@login_required
def finance_configuration(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    return render(request, 'finance/configuration.html', _fin_ctx('configuration', {}))


@login_required
def finance_facture_detail(request, invoice_id):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    error = None
    invoice = None
    lines = []
    try:
        uid, models = get_odoo_connection()
        invoice, lines = _fs.fetch_invoice_detail(uid, models, invoice_id)
        if not invoice:
            error = f'Facture #{invoice_id} introuvable.'
    except Exception as exc:
        error = f'Erreur connexion Odoo : {exc}'

    return render(request, 'finance/facture_detail.html', _fin_ctx('', {
        'error': error,
        'invoice': invoice,
        'lines': lines,
    }))


# ── Finance exports (stubs) ──────────────────────────────────────

@login_required
def finance_factures_clients_excel(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    from django.http import HttpResponse
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _apply_invoice_filters(_fs.fetch_invoices(uid, models, 'out_invoice', 2000), request)
    except Exception:
        pass
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Factures clients'
    headers = ['N° Facture', 'Date', 'Échéance', 'Client', 'Montant HT', 'Montant TTC', 'Reste à payer', 'État paiement']
    ws.append(headers)
    for r in rows:
        ws.append([r['name'], r['date'], r['date_due'], r['partner'],
                   r['amount_ht'], r['amount_ttc'], r['amount_residual'], r['payment_state']])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(buf.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = 'attachment; filename="factures_clients.xlsx"'
    return resp


@login_required
def finance_factures_fournisseurs_excel(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    from django.http import HttpResponse
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _apply_invoice_filters(_fs.fetch_invoices(uid, models, 'in_invoice', 2000), request)
    except Exception:
        pass
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Factures fournisseurs'
    ws.append(['N° Facture', 'Date', 'Échéance', 'Fournisseur', 'Montant HT', 'Montant TTC', 'Reste à payer', 'État paiement'])
    for r in rows:
        ws.append([r['name'], r['date'], r['date_due'], r['partner'],
                   r['amount_ht'], r['amount_ttc'], r['amount_residual'], r['payment_state']])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(buf.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = 'attachment; filename="factures_fournisseurs.xlsx"'
    return resp


@login_required
def finance_paiements_excel(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    from django.http import HttpResponse
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _fs.fetch_payments(uid, models, limit=2000)
    except Exception:
        pass
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Paiements'
    ws.append(['N° Paiement', 'Date', 'Partenaire', 'Montant', 'Type', 'Mode paiement', 'Journal', 'État'])
    for r in rows:
        ws.append([r['name'], r['date'], r['partner'], r['amount'],
                   r['type_label'], r['payment_method'], r['journal'], r['state_label']])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(buf.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = 'attachment; filename="paiements.xlsx"'
    return resp


def _finance_pdf_response(title, headers, rows, filename):
    resp = HttpResponse(content_type='application/pdf')
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    doc = SimpleDocTemplate(resp, pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()
    elems = [Paragraph(title, styles['Title']), Spacer(1, 6)]
    data = [headers] + rows
    tbl = Table(data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a2c4e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elems.append(tbl)
    doc.build(elems)
    return resp


@login_required
def finance_factures_clients_pdf(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _apply_invoice_filters(_fs.fetch_invoices(uid, models, 'out_invoice', 2000), request)
    except Exception:
        pass
    pdf_rows = [[r['name'], r['date'], r['date_due'], r['partner'], f"{r['amount_ttc']:.2f}", r['payment_state']] for r in rows]
    return _finance_pdf_response('Factures clients', ['N°', 'Date', 'Echéance', 'Client', 'TTC', 'Paiement'], pdf_rows, 'factures_clients.pdf')


@login_required
def finance_factures_fournisseurs_pdf(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _apply_invoice_filters(_fs.fetch_invoices(uid, models, 'in_invoice', 2000), request)
    except Exception:
        pass
    pdf_rows = [[r['name'], r['date'], r['date_due'], r['partner'], f"{r['amount_ttc']:.2f}", r['payment_state']] for r in rows]
    return _finance_pdf_response('Factures fournisseurs', ['N°', 'Date', 'Echéance', 'Fournisseur', 'TTC', 'Paiement'], pdf_rows, 'factures_fournisseurs.pdf')


@login_required
def finance_paiements_pdf(request):
    if not _finance_has_access(request.user):
        return HttpResponseForbidden("Accès refusé au module Finance.")
    rows = []
    try:
        uid, models = get_odoo_connection()
        rows = _fs.fetch_payments(uid, models, limit=2000)
    except Exception:
        pass
    pdf_rows = [[r['name'], r['date'], r['partner'], f"{r['amount']:.2f}", r['type_label'], r['payment_method'], r['state_label']] for r in rows]
    return _finance_pdf_response('Paiements', ['N°', 'Date', 'Partenaire', 'Montant', 'Type', 'Mode', 'Etat'], pdf_rows, 'paiements.pdf')


