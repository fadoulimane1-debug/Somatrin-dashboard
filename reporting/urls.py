from django.urls import path
from . import views

urlpatterns = [
    path('', views.accueil, name='accueil'),
    path('gasoil/sorties/', views.gasoil_sorties, name='gasoil_sorties'),
]
