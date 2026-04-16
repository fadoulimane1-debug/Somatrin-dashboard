import io
import json
import xmlrpc.client
from collections import defaultdict
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.http import HttpResponse
from django.shortcuts import render
from django.conf import settings

# Catégorie Odoo pour le carburant (BIEN MATERIEL / ENERGIE / CARBURANT)
CARBURANT_CATEG_ID = 262


def get_odoo_connection():
    """Retourne (uid, models) pour les appels Odoo XML-RPC."""
    common = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/common')
    uid = common.authenticate(settings.ODOO_DB, settings.ODOO_USER, settings.ODOO_PASS, {})
    models = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/object')
    return uid, models



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
def _build_sorties_domain(date_debut, date_fin, site, chauffeur, ouvrage, anomalie, societe=''):
    """Construit le domaine Odoo pour les bons de sortie gasoil."""
    domain = [
        ('picking_type_consumption', '=', True),
        ('state', '=', 'done'),
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
    if ouvrage:
        domain.append(('account_analytic_id.name', 'ilike', ouvrage))
    if anomalie == 'OK':
        domain.append(('picking_type_is_hors_affectation', '=', False))
    elif anomalie == 'Anomalie':
        domain.append(('picking_type_is_hors_affectation', '=', True))
    return domain


def _fetch_sorties_bons(uid, models, domain, limit=1000):
    """Récupère et enrichit les bons de sortie depuis Odoo."""
    pickings = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'stock.picking', 'search_read',
        [domain],
        {
            'fields': [
                'name',
                'scheduled_date',   # date réelle du bon
                'date',             # fallback
                'write_date',       # fallback
                'date_done',        # toujours récupéré pour info
                'partner_id',
                'company_id',
                'location_id',
                'picking_type_id',
                'account_analytic_id',
                'affectation_id',
                'equipment_id',
                'initial_counter',
                'actual_counter',
                'picking_type_is_hors_affectation',
                'move_ids',
            ],
            'order': 'scheduled_date desc',
            'limit': limit,
        }
    )

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
        # Choisir la meilleure date disponible : scheduled_date > date > write_date
        raw_date = (p.get('scheduled_date')
                    or p.get('date')
                    or p.get('write_date')
                    or '')
        bon_date = raw_date[:10] if raw_date else '—'

        ecart = (p.get('actual_counter') or 0) - (p.get('initial_counter') or 0)
        qty   = moves_qty.get(p['id'], 0)
        conso = round(qty / ecart, 2) if ecart > 0 else 0
        is_anomalie = p.get('picking_type_is_hors_affectation', False)
        bons.append({
            'id':             p['id'],
            'date':           bon_date,
            'name':           p.get('name', '—'),
            'societe':        p['company_id'][1] if p.get('company_id') else '—',
            'chauffeur':      p['partner_id'][1] if p.get('partner_id') else '—',
            'site':           p['location_id'][1] if p.get('location_id') else '—',
            'type_operation': p['picking_type_id'][1] if p.get('picking_type_id') else '—',
            'ouvrage':        p['account_analytic_id'][1] if p.get('account_analytic_id') else '—',
            'affectation':    p['affectation_id'][1] if p.get('affectation_id') else '—',
            'engin':          p['equipment_id'][1] if p.get('equipment_id') else '—',
            'cpt_initial':    p.get('initial_counter') or 0,
            'cpt_actuel':     p.get('actual_counter') or 0,
            'ecart':          round(ecart, 1),
            'product_qty':    round(qty, 1),
            'consommation':   conso,
            'anomalie':       'Anomalie' if is_anomalie else 'OK',
        })
    return bons


@login_required
def gasoil_sorties(request):
    date_debut  = request.GET.get('date_debut', '')
    date_fin    = request.GET.get('date_fin', '')
    societe     = request.GET.get('societe', '')
    site        = request.GET.get('site', '')
    chauffeur   = request.GET.get('chauffeur', '').strip()
    ouvrage     = request.GET.get('ouvrage', '').strip()
    anomalie    = request.GET.get('anomalie', '')
    page_number = request.GET.get('page', 1)

    bons  = []
    error = None

    try:
        uid, models = get_odoo_connection()
        domain = _build_sorties_domain(date_debut, date_fin, site, chauffeur, ouvrage, anomalie, societe)
        # Sans filtre de date : 500 derniers bons. Avec filtre : jusqu'à 2000.
        limit = 2000 if (date_debut or date_fin) else 500
        bons  = _fetch_sorties_bons(uid, models, domain, limit=limit)
    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    total_bons    = len(bons)
    total_litres  = sum(b['product_qty'] for b in bons)
    nb_anomalies  = sum(1 for b in bons if b['anomalie'] == 'Anomalie')
    conso_vals    = [b['consommation'] for b in bons if b['consommation'] > 0]
    conso_moyenne = round(sum(conso_vals) / len(conso_vals), 2) if conso_vals else 0

    paginator = Paginator(bons, 50)
    page_obj  = paginator.get_page(page_number)

    return render(request, 'gasoil/sorties.html', {
        'page_obj': page_obj, 'error': error,
        'date_debut': date_debut, 'date_fin': date_fin,
        'societe': societe, 'site': site,
        'chauffeur': chauffeur,
        'ouvrage': ouvrage, 'anomalie': anomalie,
        'total_bons':    total_bons,
        'total_litres':  round(total_litres, 1),
        'nb_anomalies':  nb_anomalies,
        'conso_moyenne': conso_moyenne,
    })


@login_required
def gasoil_sorties_export(request):
    """Export Excel des bons de sortie gasoil (openpyxl)."""
    date_debut = request.GET.get('date_debut', '')
    date_fin   = request.GET.get('date_fin', '')
    societe    = request.GET.get('societe', '')
    site       = request.GET.get('site', '')
    chauffeur  = request.GET.get('chauffeur', '').strip()
    ouvrage    = request.GET.get('ouvrage', '').strip()
    anomalie   = request.GET.get('anomalie', '')

    try:
        uid, models = get_odoo_connection()
        domain = _build_sorties_domain(date_debut, date_fin, site, chauffeur, ouvrage, anomalie, societe)
        bons   = _fetch_sorties_bons(uid, models, domain, limit=5000)
    except Exception as e:
        return HttpResponse(f"Erreur Odoo : {e}", status=500)

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
    title_cell.value = "SOMATRIN — Rapport Sorties Gasoil"
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
    if anomalie:   subtitle.append(f"Statut : {anomalie}")
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
        ("Affectation",   20),
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
            bon['affectation'],
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
                    cell.number_format = '#,##0.0'
            elif col_idx == 14:
                cell.alignment = Alignment(horizontal="center")
                if val == 'Anomalie':
                    cell.font = Font(name="Calibri", size=10, bold=True, color="B91C1C")
                else:
                    cell.font = Font(name="Calibri", size=10, bold=True, color="15803D")

        ws.row_dimensions[row_idx].height = 16

    # ── Ligne total ───────────────────────────────────────────────────────────
    total_row = len(bons) + 5
    ws.merge_cells(f"A{total_row}:H{total_row}")
    total_cell = ws.cell(row=total_row, column=1, value=f"TOTAL  —  {len(bons)} bon(s)")
    total_cell.font      = total_font
    total_cell.fill      = total_fill
    total_cell.alignment = Alignment(horizontal="right")
    total_cell.border    = border

    total_litres = sum(b['product_qty'] for b in bons)
    qty_cell = ws.cell(row=total_row, column=12, value=round(total_litres, 1))
    qty_cell.font = Font(name="Calibri", bold=True, size=11, color=NAVY)
    qty_cell.fill = total_fill
    qty_cell.alignment  = Alignment(horizontal="right")
    qty_cell.number_format = '#,##0.0'
    qty_cell.border = border

    for col in [9, 10, 11, 13, 14]:
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


# ─────────────────────────────────────────────
#  GASOIL — ENTRÉES
#  Modèle : stock.picking (réceptions de carburant)
#  Type : incoming (fournisseur → stock)
# ─────────────────────────────────────────────
@login_required
def gasoil_entrees(request):
    date_debut  = request.GET.get('date_debut', '')
    date_fin    = request.GET.get('date_fin', '')
    site        = request.GET.get('site', '')
    fournisseur = request.GET.get('fournisseur', '').strip()

    bons  = []
    error = None

    try:
        uid, models = get_odoo_connection()

        domain = [
            ('state', '=', 'done'),
            ('picking_type_id.code', '=', 'incoming'),
            ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
        ]

        if date_debut:
            domain.append(('scheduled_date', '>=', date_debut + ' 00:00:00'))
        if date_fin:
            domain.append(('scheduled_date', '<=', date_fin + ' 23:59:59'))
        if site:
            domain.append(('location_dest_id.complete_name', 'ilike', site))
        if fournisseur:
            domain.append(('partner_id.name', 'ilike', fournisseur))

        pickings = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.picking', 'search_read',
            [domain],
            {
                'fields': [
                    'name', 'scheduled_date', 'date', 'write_date',
                    'partner_id', 'location_dest_id', 'move_ids',
                ],
                'order': 'scheduled_date desc',
                'limit': 500,
            }
        )

        # Quantités et prix gasoil par picking
        moves_data = {}
        if pickings:
            all_move_ids = [mid for p in pickings for mid in p.get('move_ids', [])]
            if all_move_ids:
                moves = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'stock.move', 'search_read',
                    [[['id', 'in', all_move_ids],
                      ['product_id.categ_id', '=', CARBURANT_CATEG_ID]]],
                    {'fields': ['id', 'picking_id', 'product_qty',
                                'product_id', 'price_unit'], 'limit': 5000}
                )
                for m in moves:
                    pid = m['picking_id'][0] if m['picking_id'] else None
                    if pid:
                        if pid not in moves_data:
                            moves_data[pid] = {'qty': 0, 'price_unit': 0,
                                               'product': ''}
                        moves_data[pid]['qty']        += m.get('product_qty') or 0
                        moves_data[pid]['price_unit']  = m.get('price_unit') or 0
                        moves_data[pid]['product']     = (
                            m['product_id'][1] if m.get('product_id') else '—'
                        )

        for p in pickings:
            md  = moves_data.get(p['id'], {})
            qty = md.get('qty', 0)
            pu  = md.get('price_unit', 0)
            raw_date = (p.get('scheduled_date') or p.get('date') or p.get('write_date') or '')
            bons.append({
                'id':          p['id'],
                'date':        raw_date[:10] if raw_date else '—',
                'name':        p.get('name', '—'),
                'fournisseur': p['partner_id'][1] if p.get('partner_id') else '—',
                'site':        p['location_dest_id'][1] if p.get('location_dest_id') else '—',
                'product':     md.get('product', '—'),
                'product_qty': round(qty, 1),
                'price_unit':  round(pu, 2),
                'total':       round(qty * pu, 2),
            })

    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    total_bons   = len(bons)
    total_litres = sum(b['product_qty'] for b in bons)
    total_cout   = sum(b['total'] for b in bons)

    return render(request, 'gasoil/entrees.html', {
        'bons': bons, 'error': error,
        'date_debut': date_debut, 'date_fin': date_fin,
        'site': site, 'fournisseur': fournisseur,
        'total_bons':   total_bons,
        'total_litres': round(total_litres, 1),
        'total_cout':   round(total_cout, 2),
    })


# ─────────────────────────────────────────────
#  GASOIL — BILAN
# ─────────────────────────────────────────────
@login_required
def gasoil_bilan(request):
    annee = request.GET.get('annee', '')
    site  = request.GET.get('site', '')

    error        = None
    entrees_data = []
    sorties_data = []

    try:
        uid, models = get_odoo_connection()

        date_filter = []
        if annee:
            date_filter = [
                ('scheduled_date', '>=', f'{annee}-01-01 00:00:00'),
                ('scheduled_date', '<=', f'{annee}-12-31 23:59:59'),
            ]

        # ── Sorties ──
        domain_s = [
            ('picking_type_consumption', '=', True),
            ('state', '=', 'done'),
            ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
        ] + date_filter
        if site:
            domain_s.append(('location_id.complete_name', 'ilike', site))

        pickings_s = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.picking', 'search_read',
            [domain_s],
            {'fields': ['scheduled_date', 'date', 'write_date',
                        'location_id', 'move_ids',
                        'picking_type_is_hors_affectation'],
             'order': 'scheduled_date asc', 'limit': 2000}
        )

        # Quantités sorties
        if pickings_s:
            all_ids = [mid for p in pickings_s for mid in p.get('move_ids', [])]
            qty_map = {}
            if all_ids:
                mvs = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'stock.move', 'search_read',
                    [[['id', 'in', all_ids],
                      ['product_id.categ_id', '=', CARBURANT_CATEG_ID]]],
                    {'fields': ['picking_id', 'product_qty'], 'limit': 10000}
                )
                for m in mvs:
                    pid = m['picking_id'][0] if m['picking_id'] else None
                    if pid:
                        qty_map[pid] = qty_map.get(pid, 0) + (m['product_qty'] or 0)

            for p in pickings_s:
                raw = (p.get('scheduled_date') or p.get('date') or p.get('write_date') or '')
                sorties_data.append({
                    'date':     raw[:10] if raw else '—',
                    'site':     p['location_id'][1] if p.get('location_id') else '—',
                    'qty':      qty_map.get(p['id'], 0),
                    'anomalie': p.get('picking_type_is_hors_affectation', False),
                })

        # ── Entrées ──
        domain_e = [
            ('state', '=', 'done'),
            ('picking_type_id.code', '=', 'incoming'),
            ('move_ids.product_id.categ_id', '=', CARBURANT_CATEG_ID),
        ] + date_filter
        if site:
            domain_e.append(('location_dest_id.complete_name', 'ilike', site))

        pickings_e = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.picking', 'search_read',
            [domain_e],
            {'fields': ['scheduled_date', 'date', 'write_date',
                        'location_dest_id', 'move_ids'],
             'order': 'scheduled_date asc', 'limit': 2000}
        )

        if pickings_e:
            all_ids_e = [mid for p in pickings_e for mid in p.get('move_ids', [])]
            qty_map_e = {}
            pu_map_e  = {}
            if all_ids_e:
                mvs_e = models.execute_kw(
                    settings.ODOO_DB, uid, settings.ODOO_PASS,
                    'stock.move', 'search_read',
                    [[['id', 'in', all_ids_e],
                      ['product_id.categ_id', '=', CARBURANT_CATEG_ID]]],
                    {'fields': ['picking_id', 'product_qty', 'price_unit'], 'limit': 10000}
                )
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
                    'qty':  qty,
                    'cout': qty * pu,
                })

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

    return render(request, 'gasoil/bilan.html', {
        'error': error,
        'annee': annee, 'site': site,
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
    })
