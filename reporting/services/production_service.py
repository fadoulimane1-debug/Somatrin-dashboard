import calendar
import json
from datetime import date


class ProductionService:
    """Service de calculs pour les nouveaux dashboards Production."""

    # Sites alignés sur les fiches terrain / Odoo (carrières SOMATRIN)
    SITES = ("LH BENSLIMANE", "LH OUJDA", "SME AIT BAHA")

    @staticmethod
    def _last_months(count):
        today = date.today()
        months = []
        year, month = today.year, today.month
        for _ in range(count):
            months.append((year, month))
            month -= 1
            if month == 0:
                month = 12
                year -= 1
        months.reverse()
        return months

    @staticmethod
    def _month_label(year, month):
        return f"{calendar.month_abbr[month]} {str(year)[-2:]}"

    @classmethod
    def dashboard_data(cls):
        months_12 = cls._last_months(12)
        months_6 = months_12[-6:]
        month_labels_12 = [cls._month_label(y, m) for y, m in months_12]
        month_labels_6 = [cls._month_label(y, m) for y, m in months_6]

        production_12 = [8600, 8900, 9100, 9450, 9800, 10050, 10200, 9900, 10150, 10400, 10650, 10900]
        production_6 = production_12[-6:]
        budget_6 = [9900, 9950, 10000, 10200, 10400, 10600]

        site_split = {cls.SITES[0]: 41, cls.SITES[1]: 34, cls.SITES[2]: 25}
        couts_site = {cls.SITES[0]: 2_280_000, cls.SITES[1]: 1_930_000, cls.SITES[2]: 1_460_000}

        prod_mensuelle = production_12[-1]
        cout_total = sum(couts_site.values())
        budget_utilise = round((sum(production_6) / sum(budget_6)) * 100, 1)
        efficacite = 93.6
        ipc = 0.97

        return {
            "kpis": {
                "production_mensuelle": prod_mensuelle,
                "couts_production": cout_total,
                "budget_utilise": budget_utilise,
                "efficacite": efficacite,
                "ipc": ipc,
            },
            "charts": {
                "line_production": {"labels": month_labels_12, "values": production_12},
                "bar_prod_budget": {
                    "labels": month_labels_6,
                    "production": production_6,
                    "budget": budget_6,
                },
                "pie_sites": {"labels": list(site_split.keys()), "values": list(site_split.values())},
                "bar_couts_site": {"labels": list(couts_site.keys()), "values": list(couts_site.values())},
            },
        }

    @classmethod
    def ratios_data(cls):
        ratio_rows = [
            {"ratio": "Rendement forage (ML/h)", "site1": 17.2, "site2": 16.4, "site3": 15.1},
            {"ratio": "Métrage foré cumulé (ML)", "site1": 1240, "site2": 980, "site3": 760},
            {"ratio": "Tonnage traité (kt)", "site1": 4.47, "site2": 3.71, "site3": 2.73},
            {"ratio": "Heures machines / mois", "site1": 720, "site2": 648, "site3": 504},
            {"ratio": "Efficacité opération (%)", "site1": 94.2, "site2": 91.7, "site3": 89.4},
            {"ratio": "Coût unitaire (DH/T)", "site1": 214, "site2": 227, "site3": 239},
        ]
        radar_labels = [r["ratio"] for r in ratio_rows]
        s0, s1, s2 = cls.SITES
        radar_series = {
            s0: [r["site1"] for r in ratio_rows],
            s1: [r["site2"] for r in ratio_rows],
            s2: [r["site3"] for r in ratio_rows],
        }
        heatmap_months = [cls._month_label(y, m) for y, m in cls._last_months(12)]
        heatmap_data = [
            {"month": month, "site1": 88 + i % 7, "site2": 83 + (i * 2) % 8, "site3": 79 + (i * 3) % 9}
            for i, month in enumerate(heatmap_months)
        ]
        return {
            "rows": ratio_rows,
            "radar": {"labels": radar_labels, "series": radar_series},
            "heatmap": heatmap_data,
            "site_labels": list(cls.SITES),
        }

    @classmethod
    def ipc_data(cls):
        rows = [
            {"site": cls.SITES[0], "periode": "Mois courant", "va": 2_180_000, "couts": 2_120_000},
            {"site": cls.SITES[1], "periode": "Mois courant", "va": 1_760_000, "couts": 1_810_000},
            {"site": cls.SITES[2], "periode": "Mois courant", "va": 1_420_000, "couts": 1_510_000},
        ]
        for row in rows:
            row["ipc"] = round(row["va"] / row["couts"], 2) if row["couts"] else 0.0
            if row["ipc"] == 1:
                row["etat"] = "Bon"
            elif row["ipc"] < 1:
                row["etat"] = "Attention"
            else:
                row["etat"] = "Mauvais"

        ipc_global = round(sum(r["ipc"] for r in rows) / len(rows), 2)
        va_total = sum(r["va"] for r in rows)
        couts_total = sum(r["couts"] for r in rows)
        seuil_alerte = 1.0

        months = [cls._month_label(y, m) for y, m in cls._last_months(12)]
        ipc_evolution = [0.92, 0.95, 0.98, 1.03, 1.01, 0.99, 0.97, 0.96, 1.00, 0.98, 0.97, ipc_global]

        return {
            "kpis": {
                "ipc_global": ipc_global,
                "va": va_total,
                "couts_controles": couts_total,
                "seuil_alerte": seuil_alerte,
            },
            "rows": rows,
            "sites": list(cls.SITES),
            "charts": {
                "gauge": ipc_global,
                "bar_site": {
                    "labels": [r["site"] for r in rows],
                    "values": [r["ipc"] for r in rows],
                    "reference": 1.0,
                },
                "line_evolution": {"labels": months, "values": ipc_evolution},
            },
        }

    @classmethod
    def rapports_data(cls):
        return {
            "types": [
                "Rapport mensuel production",
                "Rapport IPC mensuel",
                "Analyse écarts budget",
                "Comparaison sites",
                "Prévisions production",
            ],
            "periodes": ["Mois", "Trimestre", "Année"],
            "sites": ["Tous"] + list(cls.SITES),
            "actions": ["PDF", "Excel", "Email", "Imprimer"],
        }

    @classmethod
    def pointages_operations_fallback_rows(cls) -> list:
        """Exemples calés sur les fiches Odoo (pointage opérations / forages) — maquette hors flux."""
        s0, s1, s2 = cls.SITES
        rows = [
            {
                "date": "26/02/2024",
                "operation": "PO-2024-003 / LH BENSLIMANE",
                "site": s0,
                "engin": "OP 11/VCEC380DL00271204",
                "equipe": "YASSINE ANNAG",
                "ouvrage": "Forage Minage Chargement et transport - LAFARGEHOLCIM MAROC",
                "foration": "Non",
                "ml_estime": 0.0,
                "statut": "Actif",
            },
            {
                "date": "02/01/2026",
                "operation": "FOR-2026-001",
                "site": s1,
                "engin": "SRO 04/AV014A1265",
                "equipe": "MOHAMED EL-OTMANI",
                "ouvrage": "S00010-Foration, Minage carrière Calcaire - LAFARGEHOLCIM MAROC",
                "foration": "Oui",
                "ml_estime": 87.5,
                "statut": "Actif",
            },
            {
                "date": "05/01/2026",
                "operation": "FOR-2026-004",
                "site": s1,
                "engin": "SRO 09/JPS20SED1306",
                "equipe": "RACHID SEGHIR",
                "ouvrage": "Foration front concasseur - GRABEMARO",
                "foration": "Oui",
                "ml_estime": 62.0,
                "statut": "Actif",
            },
            {
                "date": "31/12/2025",
                "operation": "FOR-2025-892",
                "site": s1,
                "engin": "SRO 08/JPS21SED1315",
                "equipe": "FAISSALE ANNAG",
                "ouvrage": "Foration, chargement et tir des mines - GRABEMARO",
                "foration": "Oui",
                "ml_estime": 118.2,
                "statut": "Actif",
            },
            {
                "date": "26/02/2024",
                "operation": "PO-2024-014",
                "site": s0,
                "engin": "SC 19/VCEL150HTF0014002",
                "equipe": "BRAHIM ZAGHLOULI",
                "ouvrage": "Chargement des camions clients - LAFARGEHOLCIM MAROC",
                "foration": "Non",
                "ml_estime": 0.0,
                "statut": "Actif",
            },
            {
                "date": "27/02/2024",
                "operation": "PO-2024-018",
                "site": s0,
                "engin": "2225-B-33/YV2XG10G7HH942241",
                "equipe": "KHALID CHAKIR",
                "ouvrage": "Transport inter du front au concasseur - GRABEMARO",
                "foration": "Non",
                "ml_estime": 0.0,
                "statut": "Actif",
            },
            {
                "date": "04/03/2026",
                "operation": "MIN-2026-127",
                "site": s2,
                "engin": "—",
                "equipe": "Équipe tir SME",
                "ouvrage": "12-2026 RAPPORT DE TIR N° 127 SME AIT BAHA — Ammonix / Tovex",
                "foration": "Non",
                "ml_estime": 0.0,
                "statut": "Actif",
            },
            {
                "date": "06/02/2026",
                "operation": "MIN-2026-40688",
                "site": s2,
                "engin": "—",
                "equipe": "Équipe tir SME",
                "ouvrage": "03/2026 RAPPORT DE TIR N° 40688 — Amorces 15M",
                "foration": "Non",
                "ml_estime": 0.0,
                "statut": "Critique",
            },
        ]
        for r in rows:
            r.setdefault("societe", "SOMATRIN")
        return rows

    # ─── Méthodes supplémentaires ──────────────────────────────────────────────

    @classmethod
    def get_dashboard_kpis(cls) -> dict:
        """KPI consolidés pour le dashboard Production."""
        data = cls.dashboard_data()
        kpis = data["kpis"]
        months = cls._last_months(12)
        labels = [cls._month_label(y, m) for y, m in months]
        prod_values = [8600, 8900, 9100, 9450, 9800, 10050, 10200, 9900, 10150, 10400, 10650, 10900]
        return {
            "production_mensuelle": kpis["production_mensuelle"],
            "couts_production": kpis["couts_production"],
            "budget_utilise": kpis["budget_utilise"],
            "efficacite": kpis["efficacite"],
            "ipc": kpis["ipc"],
            "ca_total": 18_540_000,
            "marge": 15.4,
            "roi": 12.8,
            "rendement_moyen": 93.6,
            "monthly_labels_json": json.dumps(labels),
            "monthly_values_json": json.dumps(prod_values),
            "sites": list(cls.SITES),
        }

    @classmethod
    def get_gasoil_data(cls) -> list:
        """Consommation gasoil par site (données calculées)."""
        site_data = [
            {"site": cls.SITES[0], "consomme": 18420, "cible": 20000, "cout_litre": 12.5},
            {"site": cls.SITES[1], "consomme": 14280, "cible": 15000, "cout_litre": 12.5},
            {"site": cls.SITES[2], "consomme": 9840, "cible": 11000, "cout_litre": 12.5},
        ]
        rows = []
        for s in site_data:
            pct = round((s["consomme"] / s["cible"]) * 100, 1)
            restant = s["cible"] - s["consomme"]
            tendance = "hausse" if pct > 95 else ("stable" if pct > 80 else "baisse")
            rows.append({
                **s,
                "restant": restant,
                "pct_cible": pct,
                "cout_total": round(s["consomme"] * s["cout_litre"], 2),
                "tendance": tendance,
            })
        return rows

    @classmethod
    def get_production_data(cls, site: str = "") -> list:
        """Données production par site."""
        base = [
            {"site": cls.SITES[0], "tonnage": 4470, "rendement": 94.2, "heures": 720, "couts": 2_280_000, "ca": 2_900_000},
            {"site": cls.SITES[1], "tonnage": 3710, "rendement": 91.7, "heures": 648, "couts": 1_930_000, "ca": 2_400_000},
            {"site": cls.SITES[2], "tonnage": 2730, "rendement": 89.4, "heures": 504, "couts": 1_460_000, "ca": 1_780_000},
        ]
        rows = []
        for r in base:
            if site and site.lower() not in r["site"].lower():
                continue
            marge = round(((r["ca"] - r["couts"]) / r["ca"]) * 100, 1)
            rows.append({**r, "marge": marge})
        return rows

    @classmethod
    def get_pointages_data(cls) -> dict:
        """Pointages opérations et formations (données calculées)."""
        months = [cls._month_label(y, m) for y, m in cls._last_months(6)]
        heures_vals = [4820, 5100, 4960, 5320, 5480, 5640]
        return {
            "kpis": {
                "heures_total": 31320,
                "personnes": 148,
                "formations": 24,
                "taux_participation": 87.5,
            },
            "rows": [
                {"site": cls.SITES[0], "heures": 13420, "personnes": 62, "formations": 10, "taux": 91.3},
                {"site": cls.SITES[1], "heures": 11080, "personnes": 51, "formations": 8, "taux": 87.5},
                {"site": cls.SITES[2], "heures": 6820, "personnes": 35, "formations": 6, "taux": 80.0},
            ],
            "charts": {
                "labels": months,
                "heures": heures_vals,
            },
        }

    @classmethod
    def get_machines_data(cls) -> list:
        """Données machines / heures / tonnages (calculées)."""
        machines = [
            ("SRO 04/AV014A1265", cls.SITES[1], "Forage", 510, 2420, 94),
            ("SRO 09/JPS20SED1306", cls.SITES[1], "Forage", 488, 2100, 91),
            ("SC 19/VCEL150HTF0014002", cls.SITES[0], "PELLE HYDRAULIQUE", 612, 1860, 89),
            ("2225-B-33/YV2XG10G7HH942241", cls.SITES[0], "Chargeur", 420, 1540, 85),
            ("SRO 08/JPS21SED1315", cls.SITES[2], "Forage", 530, 0, 96),
            ("VOLVO EC300", cls.SITES[2], "Pelle", 450, 1280, 83),
        ]
        rows = []
        for machine, site, type_, heures, tonnage, util in machines:
            cout = round(heures * 185, 2)
            rendement = round(tonnage / heures, 2) if heures and tonnage else 0.0
            rows.append({
                "machine": machine,
                "site": site,
                "type": type_,
                "heures": heures,
                "tonnage": tonnage,
                "utilisation": util,
                "etat": "Opérationnel" if util > 85 else ("Maintenance" if util > 70 else "En panne"),
                "couts": cout,
                "rendement": rendement,
            })
        return rows

    @classmethod
    def get_couts_par_nature(cls) -> dict:
        """Répartition des coûts par nature."""
        natures = [
            ("Main d'œuvre", 2_180_000, 38.8),
            ("Matières premières", 1_540_000, 27.4),
            ("Énergie & carburant", 1_020_000, 18.2),
            ("Maintenance", 560_000, 9.9),
            ("Autres charges", 320_000, 5.7),
        ]
        months = [cls._month_label(y, m) for y, m in cls._last_months(12)]
        cout_total = sum(n[1] for n in natures)
        rows = []
        for nat, montant, pct in natures:
            rows.append({
                "nature": nat,
                "montant": montant,
                "pct": pct,
                "cout_tonne": round(montant / 10_910, 2),
                "tendance": "stable",
            })
        return {
            "rows": rows,
            "total": cout_total,
            "cout_tonne_global": round(cout_total / 10_910, 2),
            "charts": {
                "pie_labels": [n[0] for n in natures],
                "pie_values": [n[1] for n in natures],
                "evolution_labels": months,
                "evolution_values": [5_020_000 + i * 80_000 for i in range(12)],
            },
        }

    @classmethod
    def get_ventes_data(cls, page: int = 1) -> dict:
        """Facturation & ventes (données calculées) avec pagination."""
        per_page = 50
        total = 74
        ca_total = 499_800_000.0
        months = [cls._month_label(y, m) for y, m in cls._last_months(12)]
        ca_monthly = [38_000_000 + i * 1_200_000 for i in range(12)]

        # Génération de lignes demo
        statuts = ["Payé", "En attente", "Partiel", "En retard"]
        clients = [
            "LAFARGEHOLCIM MAROC",
            "GRABEMARO",
            "HOLCIM MAROC",
            "CIMAR",
            "SOMATRIN (interne)",
        ]
        all_rows = []
        for i in range(1, total + 1):
            mois = (i % 12) + 1
            all_rows.append({
                "date": f"2025-{mois:02d}-{(i % 28) + 1:02d}",
                "reference": f"FACT/2025/{i:04d}",
                "client": clients[i % len(clients)],
                "montant": round(ca_total / total + (i % 5) * 150_000, 2),
                "statut": statuts[i % len(statuts)],
                "delai": f"{15 + (i % 30)} jours",
                "marge": round(12 + (i % 8), 1),
            })

        start = (page - 1) * per_page
        paginated = all_rows[start:start + per_page]
        total_pages = (total + per_page - 1) // per_page

        return {
            "rows": paginated,
            "kpis": {
                "total_factures": total,
                "ca_total": ca_total,
                "taux_paiement": 78.4,
                "delai_moyen": 22,
                "marge_moyenne": 16.2,
            },
            "pagination": {
                "page": page,
                "per_page": per_page,
                "total": total,
                "total_pages": total_pages,
            },
            "charts": {
                "labels": months,
                "ca_values": ca_monthly,
            },
        }

    @classmethod
    def get_rentabilite_data(cls) -> dict:
        """Analyse rentabilité complète (calculée)."""
        months = [cls._month_label(y, m) for y, m in cls._last_months(12)]
        ca_values = [6_200_000 + i * 180_000 for i in range(12)]
        marge_values = [14.2 + i * 0.15 for i in range(12)]
        sites = [
            {"site": cls.SITES[0], "ca": 7_080_000, "couts": 5_940_000, "marge": 16.1, "roi": 19.2, "budget_ca": 7_200_000, "budget_couts": 6_000_000},
            {"site": cls.SITES[1], "ca": 5_860_000, "couts": 5_060_000, "marge": 13.7, "roi": 15.8, "budget_ca": 6_000_000, "budget_couts": 5_100_000},
            {"site": cls.SITES[2], "ca": 4_240_000, "couts": 3_670_000, "marge": 13.4, "roi": 15.5, "budget_ca": 4_500_000, "budget_couts": 3_800_000},
        ]
        ca_total = sum(s["ca"] for s in sites)
        couts_total = sum(s["couts"] for s in sites)
        marge_globale = round(((ca_total - couts_total) / ca_total) * 100, 1)
        return {
            "kpis": {
                "ca_total": ca_total,
                "couts_total": couts_total,
                "marge_brute": marge_globale,
                "roi": 16.8,
                "ebitda": round(ca_total * 0.22, 0),
                "net": round(ca_total * 0.12, 0),
                "budget_utilise": 94.3,
                "ecart_budget": -2.7,
            },
            "sites": sites,
            "charts": {
                "labels": months,
                "ca_values": ca_values,
                "marge_values": marge_values,
            },
        }

    @classmethod
    def get_sites(cls) -> list:
        """Liste des 3 sites de production."""
        return [
            {
                "id": 1, "name": cls.SITES[0], "localisation": "Benslimane",
                "tonnage": 4470, "rendement": 94.2, "heures": 720,
                "ca": 7_080_000, "couts": 5_940_000, "marge": 16.1,
                "machines": 12, "effectif": 62, "statut": "Actif",
            },
            {
                "id": 2, "name": cls.SITES[1], "localisation": "Oujda",
                "tonnage": 3710, "rendement": 91.7, "heures": 648,
                "ca": 5_860_000, "couts": 5_060_000, "marge": 13.7,
                "machines": 9, "effectif": 51, "statut": "Actif",
            },
            {
                "id": 3, "name": cls.SITES[2], "localisation": "Chtouka Aït Baha",
                "tonnage": 2730, "rendement": 89.4, "heures": 504,
                "ca": 4_240_000, "couts": 3_670_000, "marge": 13.4,
                "machines": 7, "effectif": 35, "statut": "Actif",
            },
        ]
