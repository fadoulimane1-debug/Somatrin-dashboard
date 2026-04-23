#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génération du PDF "Rapport Sorties Gasoil" — SOMATRIN
Données d'exemple avec formatage précis selon spécifications
"""

import io
from datetime import date as _date
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image)
from reportlab.pdfgen import canvas as rl_canvas


# Couleurs
NAVY      = colors.HexColor('#1a2c4e')
ORANGE    = colors.HexColor('#E87722')
WHITE     = colors.white
RED       = colors.HexColor('#dc2626')
GREEN     = colors.HexColor('#16a34a')
ROW_ALT   = colors.HexColor('#f4f6fb')
ROW_ANOM  = colors.HexColor('#fee2e2')
GREY_TXT  = colors.HexColor('#6b7280')
GREY_BG   = colors.HexColor('#e5e7eb')
BODY_TXT  = colors.HexColor('#374151')
LIGHT_BG  = colors.HexColor('#f8f9fa')

# Date actuelle
today = '16/04/2026'


class NumberedCanvas(rl_canvas.Canvas):
    """Canvas avec numérotation de page (Page X / Y)."""
    
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
        """Dessine le pied de page avec numéro de page."""
        pw = landscape(A4)[0]
        self.saveState()
        self.setFont('Helvetica', 7)
        self.setFillColor(GREY_TXT)
        # Gauche
        self.drawString(18 * mm, 10 * mm, 'SOMATRIN — Confidentiel')
        # Droite : "Page X / Y | date"
        page_text = f'Page {self._pageNumber} / {total}  |  {today}'
        self.drawRightString(pw - 18 * mm, 10 * mm, page_text)
        self.restoreState()


def generate_gasoil_report(output_path='rapport_sorties_gasoil.pdf'):
    """Génère le PDF du rapport sorties gasoil avec données d'exemple."""
    
    # ── Données d'exemple ─────────────────────────────────────────────────
    bons_data = [
        {
            'date': '2026-04-16',
            'name': 'LHMEK/MOI/09154',
            'societe': 'SOMATRIN',
            'site': 'LHMEK/Stock',
            'ouvrage': 'S00011-Manutention des MP vers concasseurs - Chargement transport Matières premières - LAFARGEHOLCIM MAROC',
            'engin': '59087-B-33/YV2XG30G3SB50467 6',
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
            'chauffeur': 'MUSTAPHA MAHJOUB',
            'cpt_initial': 1561,
            'cpt_actuel': 1567,
            'ecart': 6.0,
            'product_qty': 52.0,
            'consommation': 8.67,
            'anomalie': 'OK',
        },
    ]

    # ── Buffer & document ─────────────────────────────────────────────────
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=20 * mm,
    )

    # ── Styles de paragraphes ─────────────────────────────────────────────
    s_logo = ParagraphStyle(
        'logo',
        fontName='Helvetica-Bold',
        fontSize=11,
        textColor=NAVY,
        alignment=TA_LEFT
    )
    
    s_confidential = ParagraphStyle(
        'confidential',
        fontName='Helvetica',
        fontSize=9,
        textColor=NAVY,
        alignment=TA_CENTER
    )
    
    s_title = ParagraphStyle(
        'title',
        fontName='Helvetica-Bold',
        fontSize=14,
        textColor=NAVY,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    
    s_subtitle = ParagraphStyle(
        'subtitle',
        fontName='Helvetica',
        fontSize=9,
        textColor=GREY_TXT,
        alignment=TA_CENTER
    )

    # ── Éléments du document ──────────────────────────────────────────────
    elems = []

    # En-tête page
    pw = landscape(A4)[0]
    hdr_data = [[
        Image('static/images/logo_somatrin.png', width=60*mm, height=12*mm),
        Paragraph('Document Confidentiel — Usage Interne', s_confidential),
        Paragraph(f'Page 1 / 453  |  {today}', 
                  ParagraphStyle('hdr_right', fontName='Helvetica', fontSize=8, 
                                textColor=GREY_TXT, alignment=TA_RIGHT))
    ]]
    hdr_tbl = Table(hdr_data, colWidths=[70 * mm, 68 * mm, 68 * mm])
    hdr_tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('LINEBELOW', (0, 0), (-1, 0), 1.5, NAVY),
        ('TOPPADDING', (0, 0), (-1, 0), 4),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
    ]))
    elems.append(hdr_tbl)
    elems.append(Spacer(1, 6 * mm))

    # Titre
    elems.append(Paragraph('Rapport Sorties Gasoil', s_title))
    
    # Sous-titre
    elems.append(Paragraph('Toutes les données', s_subtitle))
    elems.append(Spacer(1, 6 * mm))

    # ── Tableau principal ─────────────────────────────────────────────────
    # Largeurs colonnes (en mm)
    COL_MM = [14.3, 19.5, 13.0, 14.3, 28.0, 18.2, 15.6, 18.2, 10.4, 10.4, 9.1, 10.4, 10.4, 9.1]
    col_w = [c * mm for c in COL_MM]

    HEADERS = ['Date', 'N° Bon', 'Société', 'Site', 'Ouvrage', 'Engin',
               'Chauffeur', 'Cpt. Init', 'Cpt. Act',
               'Écart', 'Qté (L)', 'Conso.', 'Statut', '']

    # Styles colonnes
    s_h = ParagraphStyle('sh', fontSize=8, textColor=WHITE, fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_c = ParagraphStyle('sc', fontSize=7, textColor=BODY_TXT, fontName='Helvetica')
    s_cr = ParagraphStyle('scr', fontSize=7, textColor=BODY_TXT, fontName='Helvetica', alignment=TA_RIGHT)
    s_cc = ParagraphStyle('scc', fontSize=7, textColor=BODY_TXT, fontName='Helvetica', alignment=TA_CENTER)
    s_ok = ParagraphStyle('sok', fontSize=7, textColor=GREEN, fontName='Helvetica-Bold', alignment=TA_CENTER)
    s_co = ParagraphStyle('sco', fontSize=7, textColor=colors.HexColor('#0ea5e9'), fontName='Helvetica-Bold', alignment=TA_RIGHT)

    def trunc(s, n):
        """Tronque et ajoute ellipsis si trop long."""
        return (s[:n] + '…') if len(s) > n else s

    # Construction des lignes
    rows = [[Paragraph(h, s_h) for h in HEADERS]]
    
    for bon in bons_data:
        conso_s = f"{bon['consommation']:.2f}" if bon['consommation'] else '—'
        statut = Paragraph('OK', s_ok) if bon['anomalie'] == 'OK' else Paragraph('Anomalie', s_ok)
        
        rows.append([
            Paragraph(bon['date'], s_cc),
            Paragraph(bon['name'], s_c),
            Paragraph(trunc(bon['societe'], 12), s_c),
            Paragraph(trunc(bon['site'], 14), s_c),
            Paragraph(trunc(bon['ouvrage'], 35), s_c),  # Colonne plus large
            Paragraph(trunc(bon['engin'], 20), s_c),
            Paragraph(trunc(bon['chauffeur'], 15), s_c),
            Paragraph(f"{bon['cpt_initial']:,.0f}", s_cr),
            Paragraph(f"{bon['cpt_actuel']:,.0f}", s_cr),
            Paragraph(f"{bon['ecart']:.1f}", s_cr),
            Paragraph(f"{bon['product_qty']:.1f}", s_cr),
            Paragraph(conso_s, s_co),
            statut,
            '',  # Colonne vide pour la 14e
        ])

    # Ligne totaux
    total_qty = sum(b['product_qty'] for b in bons_data)
    rows.append([
        Paragraph(f'TOTAL — {len(bons_data)} bon{"s" if len(bons_data) != 1 else ""}', 
                  ParagraphStyle('stotal', fontSize=8, textColor=WHITE, fontName='Helvetica-Bold')),
        '', '', '', '', '', '',
        '', '',
        '',
        Paragraph(f'{total_qty:.1f}', 
                  ParagraphStyle('stq', fontSize=8, textColor=WHITE, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        '', '', '',
    ])

    # Styles tableau
    n_rows = len(rows)
    style = [
        # En-tête
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 4),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        # Corps
        ('FONTSIZE', (0, 1), (-1, -2), 7),
        ('TOPPADDING', (0, 1), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.HexColor('#d1d5db')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        # Lignes alternées
        *[('BACKGROUND', (0, i), (-1, i), ROW_ALT) for i in range(2, n_rows - 1, 2)],
        # Ligne totaux
        ('BACKGROUND', (0, -1), (-1, -1), NAVY),
        ('TEXTCOLOR', (0, -1), (-1, -1), WHITE),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 8),
        ('ALIGN', (0, -1), (-1, -1), 'RIGHT'),
        ('TOPPADDING', (0, -1), (-1, -1), 4),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 4),
        ('LINEABOVE', (0, -1), (-1, -1), 1.5, NAVY),
    ]

    main_tbl = Table(rows, colWidths=col_w, repeatRows=1)
    main_tbl.setStyle(TableStyle(style))
    elems.append(main_tbl)

    # ── Build du PDF ──────────────────────────────────────────────────────
    doc.build(elems, canvasmaker=NumberedCanvas)
    buffer.seek(0)

    # Écrire dans un fichier
    with open(output_path, 'wb') as f:
        f.write(buffer.read())
    
    print(f"✓ PDF généré : {output_path}")
    return output_path


if __name__ == '__main__':
    generate_gasoil_report()
