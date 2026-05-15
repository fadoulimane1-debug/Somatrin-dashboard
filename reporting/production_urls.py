"""Routage sous /production/ — Module Production SOMATRIN."""
from django.shortcuts import redirect
from django.urls import path

from . import production_views as pv

urlpatterns = [
    # Index (redirect dashboard)
    path('', lambda request: redirect('production_dashboard'), name='production_redirect'),
    path('production/', pv.production_index, name='production_index'),

    # ── 12 menus principaux ──────────────────────────────────────────────────
    path('dashboard/',    pv.production_dashboard,  name='production_dashboard'),
    path('gasoil/',       pv.production_gasoil,     name='production_gasoil'),
    path('detail/',       pv.production_detail,     name='production_detail'),
    path('pointages/',    pv.production_pointages,  name='production_pointages_operations_foration'),
    path('machines/',     pv.production_machines,   name='production_machines_heures_tonnages'),
    path('couts/',        pv.production_couts,      name='production_couts_nature'),
    path('ratios/',       pv.production_ratios,     name='production_ratios'),
    path('ipc/',          pv.production_ipc,        name='production_ipc'),
    path('ventes/',       pv.production_ventes,     name='production_facturation_ventes'),
    path('rentabilite/',  pv.production_rentabilite, name='production_rentabilite'),
    path('rapports/',     pv.production_rapports,   name='production_rapports'),
    path('sites/',        pv.production_sites,      name='production_sites'),

    # Chemins « lisibles » (menu base.html, cartes) — mêmes vues que ci-dessus
    path('pointages-operations-foration/', pv.production_pointages),
    path('machines-heures-tonnages/', pv.production_machines),
    path('couts-par-nature/', pv.production_couts),
    path('facturation-ventes/', pv.production_ventes),

    # ── API endpoints ────────────────────────────────────────────────────────
    path('api/kpis/',   pv.api_production_kpis,   name='production_api_kpis'),
    path('api/charts/', pv.api_production_charts, name='production_api_charts'),

    # ── Exports Gasoil ───────────────────────────────────────────────────────
    path('gasoil/export/excel/', pv.export_gasoil_excel, name='export_gasoil_excel'),
    path('gasoil/export/csv/',   pv.export_gasoil_csv,   name='export_gasoil_csv'),
    path('gasoil/export/pdf/',   pv.export_gasoil_pdf,   name='export_gasoil_pdf'),

    # ── Exports Production ───────────────────────────────────────────────────
    path('production/export/excel/', pv.export_production_excel, name='export_production_excel'),
    path('production/export/csv/',   pv.export_production_csv,   name='export_production_csv'),
    path('production/export/pdf/',   pv.export_production_pdf,   name='export_production_pdf'),

    # ── Exports Pointages ────────────────────────────────────────────────────
    path('pointages/export/excel/', pv.export_pointages_excel, name='export_pointages_excel'),
    path('pointages/export/csv/',   pv.export_pointages_csv,   name='export_pointages_csv'),
    path('pointages/export/pdf/',   pv.export_pointages_pdf,   name='export_pointages_pdf'),

    # ── Exports Machines ─────────────────────────────────────────────────────
    path('machines/export/excel/', pv.export_machines_excel, name='export_machines_excel'),
    path('machines/export/csv/',   pv.export_machines_csv,   name='export_machines_csv'),
    path('machines/export/pdf/',   pv.export_machines_pdf,   name='export_machines_pdf'),

    # ── Exports Coûts ────────────────────────────────────────────────────────
    path('couts/export/excel/', pv.export_couts_excel, name='export_couts_excel'),
    path('couts/export/csv/',   pv.export_couts_csv,   name='export_couts_csv'),
    path('couts/export/pdf/',   pv.export_couts_pdf,   name='export_couts_pdf'),

    # ── Exports Ratios ───────────────────────────────────────────────────────
    path('ratios/export/excel/', pv.export_ratios_excel, name='export_ratios_excel'),
    path('ratios/export/csv/',   pv.export_ratios_csv,   name='export_ratios_csv'),
    path('ratios/export/pdf/',   pv.export_ratios_pdf,   name='export_ratios_pdf'),

    # ── Exports IPC ──────────────────────────────────────────────────────────
    path('ipc/export/excel/', pv.export_ipc_excel, name='export_ipc_excel'),
    path('ipc/export/csv/',   pv.export_ipc_csv,   name='export_ipc_csv'),
    path('ipc/export/pdf/',   pv.export_ipc_pdf,   name='export_ipc_pdf'),

    # ── Exports Facturation ──────────────────────────────────────────────────
    path('facturation/export/excel/', pv.export_facturation_excel, name='export_facturation_excel'),
    path('facturation/export/csv/',   pv.export_facturation_csv,   name='export_facturation_csv'),
    path('facturation/export/pdf/',   pv.export_facturation_pdf,   name='export_facturation_pdf'),

    # ── Exports Rentabilité ──────────────────────────────────────────────────
    path('rentabilite/export/excel/', pv.export_rentabilite_excel, name='export_rentabilite_excel'),
    path('rentabilite/export/csv/',   pv.export_rentabilite_csv,   name='export_rentabilite_csv'),
    path('rentabilite/export/pdf/',   pv.export_rentabilite_pdf,   name='export_rentabilite_pdf'),

    # ── Exports Rapports ─────────────────────────────────────────────────────
    path('rapports/export/excel/', pv.export_rapports_excel, name='export_rapports_excel'),
    path('rapports/export/csv/',   pv.export_rapports_csv,   name='export_rapports_csv'),
    path('rapports/export/pdf/',   pv.export_rapports_pdf,   name='export_rapports_pdf'),

    # ── Exports Sites ────────────────────────────────────────────────────────
    path('sites/export/excel/', pv.export_sites_excel, name='export_sites_excel'),
    path('sites/export/csv/',   pv.export_sites_csv,   name='export_sites_csv'),
    path('sites/export/pdf/',   pv.export_sites_pdf,   name='export_sites_pdf'),
]
