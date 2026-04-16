from django.urls import path
from . import views

urlpatterns = [
    path('',                        views.accueil,               name='accueil'),
    path('gasoil/sorties/',         views.gasoil_sorties,        name='gasoil_sorties'),
    path('gasoil/sorties/export/',  views.gasoil_sorties_export, name='gasoil_sorties_export'),
    path('gasoil/entrees/',         views.gasoil_entrees,        name='gasoil_entrees'),
    path('gasoil/bilan/',           views.gasoil_bilan,          name='gasoil_bilan'),
]
