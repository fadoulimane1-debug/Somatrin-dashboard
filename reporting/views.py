import csv
import io
import json
import xmlrpc.client
from collections import defaultdict
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer)
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
                                     Paragraph, Spacer)
    from reportlab.pdfgen import canvas as rl_canvas

from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.http import HttpResponse
from django.shortcuts import render
from django.conf import settings

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
                'name', 'scheduled_date', 'date', 'write_date',
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
            'ouvrage': (p['account_analytic_id'][1] if isinstance(p.get('account_analytic_id'), list) 
            else p.get('x_affectation', '—') if p.get('x_affectation') 
            else '—'),
            'affectation':    p['affectation_id'][1]       if p.get('affectation_id')       else '—',
            'engin':          extract_matricule(engin_val[1]) if engin_val else '—',
            'categorie': (
                'Transport & Log.'  if p.get('transport_logistics') else
                'Voiture de serv.'  if p.get('service_car')         else
                'Production'
            ),
            'is_transport':   p.get('transport_logistics', False),
            'service_car':    p.get('service_car', False),
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
            {'fields': ['name', 'scheduled_date', 'date', 'write_date',
                        'location_id', 'move_ids', 'equipment_id',
                        'transport_logistics', 'service_car',
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
        'annees_list':        ['2022', '2023', '2024', '2025', '2026'],
        'sites':              SITES_LIST,
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
                    'fields': ['name', 'scheduled_date', 'state', 'partner_id', 'location_id', 'origin'],
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
                    'fields': ['name', 'scheduled_date', 'state', 'partner_id', 'location_id', 'origin'],
                    'order': 'scheduled_date desc',
                    'limit': 500,
                }
            )

        for rec in records:
            bons.append({
                'date': (rec.get('scheduled_date') or '')[:10] or '—',
                'reference': rec.get('name') or '—',
                'partenaire': rec['partner_id'][1] if rec.get('partner_id') else '—',
                'site': rec['location_id'][1] if rec.get('location_id') else '—',
                'origine': rec.get('origin') or '—',
                'etat': rec.get('state') or '—',
            })
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    return render(request, 'transport/bons_transport.html', {
        'rows': bons,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'reference': reference,
        'partenaire': partenaire,
        'site': site,
        'total_rows': len(bons),
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
        'total_rows': len(rows),
        'total_montant': total_montant,
    })


@login_required
def transport_facturation_client(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')
    client_id = request.GET.get('client_id', '').strip()
    shipping_id = request.GET.get('shipping_id', '').strip()
    company_id = request.GET.get('company_id', '').strip()
    group_by = request.GET.get('group_by', '').strip()

    rows = []
    clients = []
    companies = []
    shippings = []
    grouped_rows = []
    error = None
    try:
        uid, models = get_odoo_connection()
        base_domain = [('move_type', '=', 'out_invoice')]

        # Dropdowns depuis Odoo: client, lieu de livraison, société.
        inv_for_filters = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [base_domain],
            {'fields': ['partner_id', 'partner_shipping_id', 'company_id'], 'limit': False},
        )
        seen_clients = {}
        seen_shippings = {}
        seen_companies = {}
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

        invoices = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [domain],
            {
                'fields': [
                    'name', 'invoice_date', 'partner_id', 'partner_shipping_id',
                    'company_id', 'amount_untaxed', 'amount_total',
                ],
                'order': 'invoice_date desc',
                'limit': 1000,
            }
        )

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
            rows.append({
                'date': _fmt_date_fr(iso_date),
                'month_key': iso_date[:7] if len(iso_date) >= 7 else '—',
                'numero': inv.get('name') or '—',
                'client': inv['partner_id'][1] if inv.get('partner_id') else '—',
                'lieu_livraison': inv['partner_shipping_id'][1] if inv.get('partner_shipping_id') else '—',
                'societe': inv['company_id'][1] if inv.get('company_id') else '—',
                'ht': ht,
                'tva': tva,
                'ttc': ttc,
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

    return render(request, 'transport/facturation_client.html', {
        'rows': rows,
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'client_id': client_id,
        'clients': clients,
        'shipping_id': shipping_id,
        'shippings': shippings,
        'company_id': company_id,
        'companies': companies,
        'group_by': group_by,
        'grouped_rows': grouped_rows,
        'total_rows': len(rows),
        'total_ht': total_ht,
        'total_tva': total_tva,
        'total_ttc': total_ttc,
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

        if export in {'excel', 'csv'}:
            if export == 'csv':
                out = io.StringIO()
                writer = csv.writer(out, delimiter=';')
                writer.writerow(['Date', 'Véhicule', 'Conducteur/Opérateur', 'Société', 'Site', 'Ouvrage', 'Litres', 'Montant', 'Compteur', 'Statut'])
                for r in rows:
                    writer.writerow([r['date'], r['vehicule'], r['conducteur'], r['societe'], r['site'], r['ouvrage'], r['litres'], r['montant'], r['compteur'], r['anomalie']])
                resp = HttpResponse(out.getvalue(), content_type='text/csv; charset=utf-8')
                resp['Content-Disposition'] = 'attachment; filename=production_gasoil_sorties.csv'
                return resp

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
    error = None
    try:
        uid, models = get_odoo_connection()
        # Champ discriminant sur account.move: project_id (on exclut transport).
        base_domain = [
            ('move_type', '=', 'out_invoice'),
            ('project_id.name', 'not ilike', 'transport'),
        ]

        inv_for_filters = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [base_domain],
            {'fields': ['partner_id', 'partner_shipping_id', 'company_id', 'invoice_user_id'], 'limit': False},
        )
        seen_clients = {}
        seen_shippings = {}
        seen_companies = {}
        seen_commercials = {}
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
                ],
                'order': 'invoice_date desc',
                'limit': 2000,
            }
        )

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
            rows.append({
                'numero': inv.get('name') or '—',
                'date': _fmt_date_fr(iso_date),
                'month_key': iso_date[:7] if len(iso_date) >= 7 else '—',
                'client': inv['partner_id'][1] if inv.get('partner_id') else '—',
                'lieu_livraison': inv['partner_shipping_id'][1] if inv.get('partner_shipping_id') else '—',
                'societe': inv['company_id'][1] if inv.get('company_id') else '—',
                'ht': ht,
                'tva': tva,
                'ttc': ttc,
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
        'total_rows': len(rows),
        'total_ht': total_ht,
        'total_tva': total_tva,
        'total_ttc': total_ttc,
        'delivery_rows': delivery_rows,
        'delivery_total_count': delivery_total_count,
        'delivery_total_ht': delivery_total_ht,
        'delivery_total_tva': delivery_total_tva,
        'delivery_total_ttc': delivery_total_ttc,
    })


@login_required
def transport_rentabilite(request):
    date_debut = request.GET.get('date_debut', '')
    date_fin = request.GET.get('date_fin', '')

    revenus = []
    couts = []
    error = None
    try:
        uid, models = get_odoo_connection()

        invoice_domain = [('move_type', '=', 'out_invoice')]
        invoice_domain += _transport_domain_dates(date_debut, date_fin, 'invoice_date')
        invoices = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'account.move', 'search_read',
            [invoice_domain],
            {'fields': ['invoice_date', 'amount_untaxed'], 'limit': 5000}
        )
        revenus = [
            {'date': (inv.get('invoice_date') or '')[:10], 'montant': inv.get('amount_untaxed') or 0}
            for inv in invoices
            if inv.get('invoice_date')
        ]

        cost_domain = []
        cost_domain += _transport_domain_dates(date_debut, date_fin, 'date')
        cost_items = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'fleet.vehicle.cost', 'search_read',
            [cost_domain],
            {'fields': ['date', 'amount', 'cost_subtype_id'], 'limit': 5000}
        )
        couts = [
            {
                'date': (item.get('date') or '')[:10],
                'montant': item.get('amount') or 0,
                'nature': item['cost_subtype_id'][1] if item.get('cost_subtype_id') else '—',
            }
            for item in cost_items
            if item.get('date')
        ]
    except Exception as exc:
        error = f'Erreur de connexion Odoo : {exc}'

    total_revenus = round(sum(r['montant'] for r in revenus), 2)
    total_couts = round(sum(c['montant'] for c in couts), 2)
    resultat = round(total_revenus - total_couts, 2)
    marge_pct = round((resultat / total_revenus * 100), 2) if total_revenus else 0

    couts_par_nature = defaultdict(float)
    for cost in couts:
        couts_par_nature[c['nature']] += cost['montant']
    couts_agg = [
        {'nature': k, 'montant': round(v, 2)}
        for k, v in sorted(couts_par_nature.items(), key=lambda kv: -kv[1])
    ]

    return render(request, 'transport/rentabilite.html', {
        'error': error,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'total_revenus': total_revenus,
        'total_couts': total_couts,
        'resultat': resultat,
        'marge_pct': marge_pct,
        'couts_agg': couts_agg,
    })


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
    """Récupère les alertes QHSE depuis Odoo."""
    alerts = []
    try:
        alerts = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'quality.alert', 'search_read',
            [_build_qhse_alert_domain(site)],
            {
                'fields': ['name', 'date', 'state', 'company_id', 'location_id', 'team_id', 'category_id', 'user_id'],
                'order': 'date desc',
                'limit': limit
            }
        )
    except Exception:
        alerts = []
    return alerts


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
    # Données d'exemple (mêmes que dans generate_report_pdf.py)
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
    """Export PDF — redirige vers gasoil_sorties avec export=pdf."""
    # Passe par le GET en le modifiant localement
    request.GET = request.GET.copy()
    request.GET._mutable = True
    request.GET['export'] = 'pdf'
    request.GET._mutable = False
    return gasoil_sorties(request)
