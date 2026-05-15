def user_groups(request):
    """Expose les noms de groupes de l'utilisateur dans tous les templates."""
    if request.user.is_authenticated:
        groups = {g.lower().strip() for g in request.user.groups.values_list('name', flat=True)}
    else:
        groups = set()
    return {'user_groups': groups}
