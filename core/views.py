from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect


@login_required
def redirect_by_role(request):
    """Redirige l'utilisateur vers son espace selon son rôle/groupe."""
    user = request.user

    if user.is_superuser:
        return redirect('/admin/')

    groups = set(user.groups.values_list('name', flat=True))

    if groups & {'pilotage', 'si'} or user.is_staff:
        return redirect('/accueil/')
    if 'transport' in groups:
        return redirect('/transport/')
    if 'production' in groups:
        return redirect('/production/')
    if 'parc_materiel' in groups:
        return redirect('/maintenance/')
    if 'achat' in groups:
        return redirect('/achats/')
    if 'rh' in groups:
        return redirect('/rh/')
    if 'finance' in groups:
        return redirect('/comptabilite/')
    if groups & {'hse', 'smq'}:
        return redirect('/qhse/')

    # Défaut : gasoil
    return redirect('/gasoil/sorties/')
