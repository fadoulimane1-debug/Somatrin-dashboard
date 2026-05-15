"""Routage sous /qhse/ — Odoo 16 QHSE."""
from django.shortcuts import redirect
from django.urls import path

from . import qhse_views

urlpatterns = [
    path('', lambda request: redirect('qhse_dashboard'), name='qhse_index'),
    path('dashboard/', qhse_views.qhse_dashboard, name='qhse_dashboard'),
    path('incidents-accidents/', qhse_views.qhse_incidents, name='qhse_incidents'),
    path('plan-actions/', qhse_views.qhse_plan_actions, name='qhse_plan_actions'),
    path('achats-qhse/', qhse_views.qhse_achats, name='qhse_achats'),
    path('consommations-qhse/', qhse_views.qhse_consommations, name='qhse_consommations'),
    path('produits-hse/', qhse_views.qhse_produits_hse, name='qhse_produits_hse'),
    path('indicateurs-hse/', qhse_views.qhse_indicateurs, name='qhse_indicateurs'),
    path('audits-qualite/', qhse_views.qhse_audits, name='qhse_audits'),
    path('factures-qhse/', qhse_views.qhse_factures, name='qhse_factures'),
    path('entrees/', qhse_views.qhse_entrees, name='qhse_entrees'),
    path('sorties/', qhse_views.qhse_sorties, name='qhse_sorties'),
    path('bilan/', qhse_views.qhse_bilan, name='qhse_bilan'),
    path('action/<int:action_id>/', qhse_views.qhse_action_detail, name='qhse_action_detail'),
    path('configuration/', qhse_views.configuration_qhse, name='qhse_configuration'),
    path('api/kpis/', qhse_views.api_qhse_kpis, name='qhse_api_kpis'),
    path('api/incidents/', qhse_views.api_incidents_month, name='qhse_api_incidents'),
]
