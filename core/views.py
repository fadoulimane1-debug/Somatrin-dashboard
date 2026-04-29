from django.contrib.auth.decorators import login_required
from django.contrib.auth.views import LoginView
from django.shortcuts import redirect


def _normalize_groups(user):
    return {g.lower().strip() for g in user.groups.values_list('name', flat=True)}


def _has_any(groups, candidates):
    return any(c in groups for c in candidates)


def _redirect_path_for_user(user):
    """
    Route d'atterrissage unifiée selon rôle/groupe.
    Priorité métier demandée :
      1) Transport & Logistique
      2) Production
      puis autres services.
    """
    uname = (user.username or '').strip().lower()
    if uname == 'transport':
        return '/transport/'
    if uname == 'production':
        return '/production/'

    if user.is_superuser:
        return '/admin/'

    groups = _normalize_groups(user)

    if _has_any(groups, {'pilotage', 'si'}) or user.is_staff:
        return '/accueil/'

    # Variantes de nom de groupe côté admin.
    if _has_any(groups, {'transport', 'transport & logistique', 'transport et logistique'}):
        return '/transport/'

    if _has_any(groups, {'production', 'exploitation', 'exploitation des carrières'}):
        return '/production/'

    if 'parc_materiel' in groups:
        return '/maintenance/'
    if 'achat' in groups:
        return '/achats/'
    if 'rh' in groups:
        return '/rh/'
    if 'finance' in groups:
        return '/comptabilite/'
    if _has_any(groups, {'hse', 'smq', 'qhse'}):
        return '/qhse/'

    return '/gasoil/sorties/'


class RoleAwareLoginView(LoginView):
    """
    Redirige l'utilisateur selon son rôle après authentification
    (ignore les `next` qui pointent vers un module non cible).
    """
    def get_success_url(self):
        return _redirect_path_for_user(self.request.user)


@login_required
def redirect_by_role(request):
    """Redirige l'utilisateur vers son espace selon son rôle/groupe."""
    return redirect(_redirect_path_for_user(request.user))
