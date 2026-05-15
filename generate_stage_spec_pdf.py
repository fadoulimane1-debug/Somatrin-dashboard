from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE_DIR = Path(__file__).resolve().parent
OUT_PDF = BASE_DIR / "Cahier_Charges_Stage_SOMATRIN.pdf"
LOGO = BASE_DIR / "static" / "images" / "logo_somatrin.png"
DIAGRAM_USE_CASE = Path(
    r"C:\Users\SOMATRIN\.cursor\projects\c-Users-SOMATRIN-Desktop-somatrin-dev\assets\c__Users_SOMATRIN_AppData_Roaming_Cursor_User_workspaceStorage_fece56549aaa09def2c72d4202fe07d0_images_Le_Connecter_Flow-2026-04-28-110424-9cfa9284-add7-4b42-a3a5-e2ccab5580c7.png"
)
DIAGRAM_CLASS = Path(
    r"C:\Users\SOMATRIN\.cursor\projects\c-Users-SOMATRIN-Desktop-somatrin-dev\assets\c__Users_SOMATRIN_AppData_Roaming_Cursor_User_workspaceStorage_fece56549aaa09def2c72d4202fe07d0_images_Le_Connecter_Flow-2026-04-28-110749-e7d4fb48-de00-4ed6-920a-c5da639d8678.png"
)
DIAGRAM_SEQ = Path(
    r"C:\Users\SOMATRIN\.cursor\projects\c-Users-SOMATRIN-Desktop-somatrin-dev\assets\c__Users_SOMATRIN_AppData_Roaming_Cursor_User_workspaceStorage_fece56549aaa09def2c72d4202fe07d0_images_Le_Connecter_Flow-2026-04-28-111312-1378d9e5-9722-4055-93e7-d22110c8b45f.png"
)


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#6B7280"))
            self.drawString(15 * mm, 10 * mm, "SOMATRIN — Document Confidentiel")
            self.drawRightString(195 * mm, 10 * mm, f"Page {self._pageNumber} / {total}")
            super().showPage()
        super().save()


def p(text, style):
    return Paragraph(text, style)


def section(story, title, body_style, h1_style, lines):
    story.append(p(title, h1_style))
    for line in lines:
        story.append(p(line, body_style))


def build_pdf():
    styles = getSampleStyleSheet()
    cover_title = ParagraphStyle(
        "cover_title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=21,
        textColor=colors.HexColor("#1A2C4E"),
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    cover_sub = ParagraphStyle(
        "cover_sub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=11,
        textColor=colors.HexColor("#374151"),
        alignment=TA_CENTER,
        leading=16,
        spaceAfter=14,
    )
    h1 = ParagraphStyle(
        "h1",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        textColor=colors.HexColor("#1A2C4E"),
        spaceBefore=8,
        spaceAfter=6,
    )
    body = ParagraphStyle(
        "body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        alignment=TA_JUSTIFY,
        spaceAfter=5,
    )
    toc_style = ParagraphStyle(
        "toc",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#1F2937"),
        spaceAfter=3,
    )

    doc = SimpleDocTemplate(
        str(OUT_PDF),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
        title="Cahier des charges — Stage SOMATRIN",
        author="Imane Fadoul",
    )

    story = []
    if LOGO.exists():
        story.append(Image(str(LOGO), width=44 * mm, height=14 * mm))
        story.append(Spacer(1, 10))

    story.append(p("Cahier des Charges", cover_title))
    story.append(
        p(
            "Conception et Développement d'une Application Web de Reporting Multi-Services "
            "connectée à l'ERP Odoo via API XML-RPC",
            cover_sub,
        )
    )
    story.append(Spacer(1, 3))

    ident = [
        ["Entreprise d'accueil", "SOMATRIN"],
        ["Stagiaire", "Imane Fadoul"],
        ["Filière", "Génie Informatique (GI)"],
        ["Encadrant entreprise", "Mustapha Bouaroua"],
        ["Encadrant pédagogique", "Youssef Baddi"],
        ["Document", "Spécification des besoins et conception"],
    ]
    t = Table(ident, colWidths=[62 * mm, 106 * mm])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F3F4F6")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(t)
    story.append(PageBreak())

    story.append(p("Sommaire", h1))
    toc_items = [
        "1. Contexte, problématique et objectifs du stage",
        "2. Périmètre fonctionnel et modules couverts",
        "3. Acteurs, rôles et gouvernance d'accès",
        "4. Spécification des besoins fonctionnels",
        "5. Spécification des besoins non fonctionnels",
        "6. Conception générale de la solution",
        "7. Conception détaillée par module",
        "8. Modèle de données et intégration Odoo XML-RPC",
        "9. Règles de gestion et KPI métier",
        "10. Stratégie de test, validation et qualité",
        "11. Planning prévisionnel et jalons",
        "12. Risques projet et actions de mitigation",
        "13. Livrables attendus et critères d'acceptation",
        "14. Conclusion et perspectives",
    ]
    for item in toc_items:
        story.append(p(item, toc_style))
    story.append(PageBreak())

    section(
        story,
        "1. Contexte, problématique et objectifs du stage",
        body,
        h1,
        [
            "SOMATRIN dispose d'un ERP Odoo centralisant les données opérationnelles. Le besoin principal identifié "
            "est de disposer d'une couche de pilotage visuelle, orientée décision, sans perturber les processus de saisie Odoo.",
            "Le stage a pour objectif de développer une application web de reporting multi-services permettant la lecture rapide des indicateurs, "
            "la séparation des périmètres métiers et l'export des données pour la communication managériale.",
            "Les objectifs spécifiques sont: améliorer la visibilité des activités, accélérer l'analyse opérationnelle, sécuriser l'accès par rôles "
            "et standardiser la production de rapports.",
        ],
    )

    section(
        story,
        "2. Périmètre fonctionnel et modules couverts",
        body,
        h1,
        [
            "Le périmètre couvre les modules Gasoil, Transport & Logistique, Production, Achats & Approvisionnement et Parc & Maintenance.",
            "Chaque module propose des écrans avec filtres métier, KPI, tableaux analytiques, et fonctions d'export CSV/Excel/PDF.",
            "Le projet conserve Odoo comme source de vérité et implémente une couche de restitution orientée pilotage.",
        ],
    )

    modules = [
        ["Gasoil", "Entrées, Sorties, Bilan, anomalies, indicateurs de consommation"],
        ["Transport & Logistique", "Bons transport, coûts par nature, facturation client, rentabilité"],
        ["Production", "Coûts, facturation ventes, rentabilité, analyse sites"],
        ["Achats & Approvisionnement", "Demandes d'achat, demandes de prix, bons de commande"],
        ["Parc & Maintenance", "Équipements, disponibilité, ordres maintenance, coûts"],
    ]
    mt = Table([["Module", "Fonctions clés"]] + modules, colWidths=[55 * mm, 113 * mm])
    mt.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A2C4E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(mt)

    section(
        story,
        "3. Acteurs, rôles et gouvernance d'accès",
        body,
        h1,
        [
            "Les acteurs sont: Direction/Pilotage, responsables Transport, Production, Achats, Parc & Maintenance, administrateurs SI.",
            "La gouvernance d'accès repose sur des rôles et groupes Django (transport, production, achat, parc_materiel, pilotage, etc.).",
            "Chaque profil accède à son périmètre fonctionnel, ce qui réduit les risques d'erreur d'interprétation et garantit la cohérence métier.",
        ],
    )

    section(
        story,
        "4. Spécification des besoins fonctionnels",
        body,
        h1,
        [
            "<b>Besoins transverses:</b> authentification, filtres dynamiques, KPI, tableaux, indicateurs visuels, recherche rapide, exports.",
            "<b>Achats:</b> suivi demandes d'achat, demandes de prix et bons de commande avec états, montants, responsables et fournisseurs.",
            "<b>Parc & Maintenance:</b> visibilité sur disponibilité des équipements, ordres maintenance et charge opérationnelle.",
            "<b>Production / Transport:</b> séparation stricte des résultats et filtres adaptés aux contextes métier.",
        ],
    )

    section(
        story,
        "5. Spécification des besoins non fonctionnels",
        body,
        h1,
        [
            "<b>Performance:</b> temps de réponse fluide, limitation et pagination lors des lectures Odoo volumineuses.",
            "<b>Sécurité:</b> contrôle d'accès par rôle, limitation de visibilité des modules selon profil utilisateur.",
            "<b>Disponibilité:</b> messages d'erreur explicites en cas de modèle indisponible ou de défaut de connexion Odoo.",
            "<b>Ergonomie:</b> interface homogène, charte graphique professionnelle, lisibilité des états et priorités.",
        ],
    )

    section(
        story,
        "6. Conception générale de la solution",
        body,
        h1,
        [
            "L'architecture repose sur Django en couche applicative et Odoo en backend de données. Les vues Django orchestrent les appels XML-RPC, "
            "appliquent les règles métier et alimentent les templates.",
            "La solution suit une logique en couches: présentation (templates), traitement (views), intégration (services Odoo), export documentaire.",
            "Les composants sont conçus pour être évolutifs et facilement maintenables.",
        ],
    )

    section(
        story,
        "7. Conception détaillée par module",
        body,
        h1,
        [
            "<b>Demandes d'achat:</b> fallback intelligent des champs Odoo, badges d'état, chart de répartition, export multi-format.",
            "<b>Demandes de prix:</b> suivi SLA interne, indicateurs d'attente fournisseur, visualisation d'échéance, export PDF aligné.",
            "<b>Bons de commande:</b> lecture HT/Taxes/TTC, délai livraison calculé, filtres avancés et exports managériaux.",
            "<b>Parc & Maintenance:</b> disponibilité opérationnelle, inventaire équipements et ordres maintenance.",
        ],
    )

    section(
        story,
        "7.1 Conception UML (diagrammes obligatoires)",
        body,
        h1,
        [
            "Les diagrammes suivants matérialisent la conception logicielle de la solution: vision structurelle des composants, "
            "architecture applicative globale et scénario de séquence d'une consultation connectée à Odoo.",
        ],
    )

    diagram_title = ParagraphStyle(
        "diagram_title",
        parent=styles["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=colors.HexColor("#1A2C4E"),
        alignment=TA_CENTER,
        spaceBefore=8,
        spaceAfter=6,
    )

    def add_diagram(title_text, img_path):
        story.append(p(title_text, diagram_title))
        if img_path.exists():
            # Largeur utile ~ 174 mm (A4 - marges). On garde une hauteur lisible.
            story.append(Image(str(img_path), width=170 * mm, height=84 * mm))
            story.append(Spacer(1, 6))
        else:
            story.append(p("Diagramme non trouvé au chemin fourni.", body))
            story.append(Spacer(1, 4))

    add_diagram("Diagramme de cas d'utilisation", DIAGRAM_USE_CASE)
    story.append(PageBreak())
    add_diagram("Diagramme de classes", DIAGRAM_CLASS)
    story.append(PageBreak())
    add_diagram("Diagramme de séquence", DIAGRAM_SEQ)
    story.append(PageBreak())

    section(
        story,
        "8. Modèle de données et intégration Odoo XML-RPC",
        body,
        h1,
        [
            "Les modèles Odoo sollicités incluent principalement: stock.move, stock.picking, account.move, purchase.order, maintenance.equipment, maintenance.request.",
            "Les appels se basent sur search_read avec domaines dynamiques, contrôlés par filtres utilisateurs.",
            "Des mécanismes de fallback sont intégrés pour s'adapter aux différences de schéma selon les instances Odoo.",
        ],
    )

    section(
        story,
        "9. Règles de gestion et KPI métier",
        body,
        h1,
        [
            "Les KPI traduisent la performance opérationnelle: volume, montants, taux d'avancement, états en attente, éléments confirmés et retards.",
            "La distinction des périmètres (Production vs Transport) est une règle prioritaire de cohérence métier.",
            "Les indicateurs de type SLA sont explicitement marqués comme règles internes tant que la date d'échéance native n'est pas activée dans Odoo.",
        ],
    )

    section(
        story,
        "10. Stratégie de test, validation et qualité",
        body,
        h1,
        [
            "La stratégie couvre les tests fonctionnels (filtres, états, KPI), les tests d'intégration Odoo (connectivité, disponibilité modèles), "
            "et les tests de restitution (CSV, Excel, PDF).",
            "Des validations croisées sont prévues avec les utilisateurs métiers pour garantir l'alignement fonctionnel.",
            "La qualité est assurée par des vérifications techniques régulières (cohérence Django, robustesse des vues, lisibilité UI).",
        ],
    )

    section(
        story,
        "11. Planning prévisionnel et jalons",
        body,
        h1,
        [
            "Jalon 1: cadrage besoin et architecture cible.",
            "Jalon 2: implémentation modules principaux (Gasoil, Transport, Production).",
            "Jalon 3: implémentation Achats et Parc & Maintenance.",
            "Jalon 4: enrichissement UI/UX, exports, validation métier et documentation finale.",
        ],
    )

    section(
        story,
        "12. Risques projet et actions de mitigation",
        body,
        h1,
        [
            "<b>Risque schéma Odoo variable:</b> mitigation par fallback de champs et messages explicites.",
            "<b>Risque qualité des données:</b> mitigation par filtres contrôlés et vérification de cohérence.",
            "<b>Risque disponibilité API:</b> mitigation par try/except et gestion des erreurs orientée utilisateur.",
            "<b>Risque de charge:</b> mitigation par limites de lecture, pagination et optimisation des rendus.",
        ],
    )

    section(
        story,
        "13. Livrables attendus et critères d'acceptation",
        body,
        h1,
        [
            "Application web opérationnelle de reporting multi-services.",
            "Écrans livrés avec KPI, filtres, tableaux et exports.",
            "Cahier des charges complet (ce document) et guide de démonstration.",
            "Critères d'acceptation: exactitude des données, stabilité des écrans, conformité besoins métier.",
        ],
    )

    section(
        story,
        "14. Conclusion et perspectives",
        body,
        h1,
        [
            "Le projet fournit une base robuste de pilotage pour SOMATRIN, en valorisant les données Odoo au service de la décision.",
            "Les perspectives incluent l'activation de nouvelles règles natives Odoo (dates d'expiration officielles), "
            "l'ajout d'analyses avancées et l'industrialisation des tableaux de bord stratégiques.",
        ],
    )

    doc.build(story, canvasmaker=NumberedCanvas)


if __name__ == "__main__":
    build_pdf()
    print(f"PDF généré: {OUT_PDF}")
