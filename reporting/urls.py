from django.urls import path
from django.shortcuts import redirect
from . import views

urlpatterns = [
    path('',                        views.accueil,               name='accueil'),
    path('accueil/',                views.accueil,               name='accueil_page'),
    path('transport/',              lambda request: redirect('transport_bons'), name='transport_index'),
    path('production/',             lambda request: redirect('production_index'), name='production_redirect'),
    path('production/production/',          views.production_index,              name='production_index'),
    path('production/gasoil/',              views.production_gasoil,             name='production_gasoil'),
    path('production/couts-par-nature/',    views.production_couts_nature,       name='production_couts_nature'),
    path('transport/bons-transport/',      views.transport_bons,               name='transport_bons'),
    path('transport/gasoil/',              views.transport_gasoil,             name='transport_gasoil'),
    path('transport/couts-par-nature/',    views.transport_couts_nature,       name='transport_couts_nature'),
    path('transport/facturation-client/',  views.transport_facturation_client, name='transport_facturation_client'),
    path('transport/rentabilite/',         views.transport_rentabilite,        name='transport_rentabilite'),
    path('production/facturation-ventes/', views.production_facturation_ventes, name='production_facturation_ventes'),
    path('gasoil/sorties/',         views.gasoil_sorties,        name='gasoil_sorties'),
    path('gasoil/sorties/export/',  views.gasoil_sorties_export, name='gasoil_sorties_export'),
    path('gasoil/sorties/export/csv/', views.gasoil_sorties_csv, name='gasoil_sorties_csv'),
    path('gasoil/sorties/rapport/', views.gasoil_rapport, name='gasoil_rapport'),
    path('gasoil/entrees/',         views.gasoil_entrees,        name='gasoil_entrees'),
    path('gasoil/bilan/',           views.gasoil_bilan,          name='gasoil_bilan'),
    path('qhse/entrees/',           views.qhse_entrees,          name='qhse_entrees'),
    path('qhse/sorties/',           views.qhse_sorties,          name='qhse_sorties'),
    path('qhse/bilan/',             views.qhse_bilan,            name='qhse_bilan'),
]
