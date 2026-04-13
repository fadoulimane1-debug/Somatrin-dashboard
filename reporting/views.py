import xmlrpc.client
from django.shortcuts import render
from django.conf import settings


def get_odoo_connection():
    """Retourne (uid, models) pour les appels Odoo XML-RPC."""
    common = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/common')
    uid = common.authenticate(settings.ODOO_DB, settings.ODOO_USER, settings.ODOO_PASS, {})
    models = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/object')
    return uid, models


def accueil(request):
    return render(request, 'accueil.html')


def gasoil_sorties(request):
    # --- Récupération des filtres depuis GET ---
    date_debut = request.GET.get('date_debut', '')
    date_fin   = request.GET.get('date_fin', '')
    site       = request.GET.get('site', '')
    categorie  = request.GET.get('categorie', '')
    chauffeur  = request.GET.get('chauffeur', '').strip()
    ouvrage    = request.GET.get('ouvrage', '').strip()
    anomalie   = request.GET.get('anomalie', '')

    bons = []
    error = None

    try:
        uid, models = get_odoo_connection()

        # --- Construction du domain Odoo ---
        domain = []

        if date_debut:
            domain.append(('date', '>=', date_debut + ' 00:00:00'))
        if date_fin:
            domain.append(('date', '<=', date_fin + ' 23:59:59'))
        if site:
            domain.append(('location_id.complete_name', 'ilike', site))
        if categorie:
            domain.append(('product_uom_id.name', '=', categorie))
        if chauffeur:
            domain.append(('x_chauffeur', 'ilike', chauffeur))
        if ouvrage:
            domain.append(('x_affectation', 'ilike', ouvrage))

        # --- Appel Odoo ---
        bons = models.execute_kw(
            settings.ODOO_DB, uid, settings.ODOO_PASS,
            'stock.move', 'search_read',
            [domain],
            {
                'fields': [
                    'date', 'name',
                    'location_id',
                    'x_affectation',
                    'x_chauffeur',
                    'product_uom_id',
                    'x_cpt_initial',
                    'x_cpt_actuel',
                    'x_ecart',
                    'product_qty',
                    'x_consommation',
                    'x_anomalie',
                ],
                'order': 'date desc',
                'limit': 500,
            }
        )

        # --- Filtrage anomalie côté Python ---
        if anomalie == 'OK':
            bons = [b for b in bons if str(b.get('x_anomalie', '')).lower() == 'ok']
        elif anomalie == 'Anomalie':
            bons = [b for b in bons if str(b.get('x_anomalie', '')).lower() != 'ok']

    except Exception as e:
        error = f"Erreur de connexion Odoo : {e}"

    # --- KPI calculés ---
    total_bons      = len(bons)
    total_litres    = sum(b.get('product_qty', 0) or 0 for b in bons)
    nb_anomalies    = sum(1 for b in bons if str(b.get('x_anomalie', '')).lower() != 'ok')
    conso_values    = [b.get('x_consommation', 0) or 0 for b in bons if b.get('x_consommation')]
    conso_moyenne   = round(sum(conso_values) / len(conso_values), 2) if conso_values else 0

    context = {
        'bons'          : bons,
        'error'         : error,
        'date_debut'    : date_debut,
        'date_fin'      : date_fin,
        'site'          : site,
        'categorie'     : categorie,
        'chauffeur'     : chauffeur,
        'ouvrage'       : ouvrage,
        'anomalie'      : anomalie,
        # KPI
        'total_bons'    : total_bons,
        'total_litres'  : round(total_litres, 1),
        'nb_anomalies'  : nb_anomalies,
        'conso_moyenne' : conso_moyenne,
    }
    return render(request, 'gasoil/sorties.html', context)
