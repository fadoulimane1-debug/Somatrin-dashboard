import xmlrpc.client

ODOO_URL  = 'https://somatrin.karizma.one'
ODOO_DB   = 'somatrin_PROD'
ODOO_USER = 'bouaraoua.mustapha@somatrin.ma'
ODOO_PASS = 'Bm141174'

print("=" * 60)
print("CONNEXION ODOO")
print("=" * 60)
common = xmlrpc.client.ServerProxy(f'{ODOO_URL}/xmlrpc/2/common')
uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASS, {})
print(f"UID : {uid}")
if not uid:
    print("ERREUR : authentification échouée")
    exit(1)

models = xmlrpc.client.ServerProxy(f'{ODOO_URL}/xmlrpc/2/object')

# ── 1. Tous les champs de stock.move ─────────────────────────────────────────
print("\n" + "=" * 60)
print("CHAMPS DISPONIBLES DANS stock.move")
print("=" * 60)
try:
    fields = models.execute_kw(
        ODOO_DB, uid, ODOO_PASS,
        'stock.move', 'fields_get',
        [],
        {'attributes': ['string', 'type']}
    )
    for fname, finfo in sorted(fields.items()):
        print(f"  {fname:<40} [{finfo['type']:<12}]  {finfo['string']}")
    print(f"\nTotal : {len(fields)} champs")
except Exception as e:
    print(f"ERREUR : {e}")

# ── 2. Valeurs distinctes de x_affectation ───────────────────────────────────
print("\n" + "=" * 60)
print("VALEURS DISTINCTES DE x_affectation")
print("=" * 60)
try:
    recs = models.execute_kw(
        ODOO_DB, uid, ODOO_PASS,
        'stock.move', 'search_read',
        [[['x_affectation', '!=', False]]],
        {'fields': ['x_affectation'], 'limit': 5000}
    )
    print(f"Enregistrements trouvés : {len(recs)}")

    if recs:
        print(f"\nType de x_affectation : {type(recs[0]['x_affectation']).__name__}")
        print(f"Exemple brut          : {recs[0]}")

    ouvrages = sorted({
        r['x_affectation'][1] if isinstance(r['x_affectation'], list) else r['x_affectation']
        for r in recs
        if r.get('x_affectation') and r['x_affectation'] is not False
    })
    print(f"\nValeurs distinctes ({len(ouvrages)}) :")
    for o in ouvrages:
        print(f"  - {o}")
except Exception as e:
    print(f"ERREUR : {e}")
