import xmlrpc.client

ODOO_URL  = 'https://somatrin.karizma.one'
ODOO_DB   = 'somatrin_PROD'
ODOO_USER = 'bouaraoua.mustapha@somatrin.ma'
ODOO_PASS = 'Bm141174'

print("Connexion à Odoo...")
common = xmlrpc.client.ServerProxy(f'{ODOO_URL}/xmlrpc/2/common')
uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASS, {})
print(f"UID: {uid}")

if not uid:
    print("ERREUR : authentification échouée")
    exit(1)

models = xmlrpc.client.ServerProxy(f'{ODOO_URL}/xmlrpc/2/object')

print("\nRecherche des enregistrements stock.move avec x_affectation...")
try:
    recs = models.execute_kw(
        ODOO_DB, uid, ODOO_PASS,
        'stock.move', 'search_read',
        [[['x_affectation', '!=', False]]],
        {'fields': ['x_affectation'], 'limit': 10}
    )
    print(f"Nombre de résultats (limité à 10) : {len(recs)}")
    print("\nExemples de valeurs x_affectation :")
    for r in recs:
        print(f"  id={r['id']} | x_affectation={r['x_affectation']} | type={type(r['x_affectation']).__name__}")
except Exception as e:
    print(f"ERREUR : {e}")
    exit(1)

print("\nExtraction de tous les ouvrages distincts (limit=5000)...")
try:
    all_recs = models.execute_kw(
        ODOO_DB, uid, ODOO_PASS,
        'stock.move', 'search_read',
        [[['x_affectation', '!=', False]]],
        {'fields': ['x_affectation'], 'limit': 5000}
    )
    ouvrages = sorted({
        r['x_affectation'][1] if isinstance(r['x_affectation'], list) else r['x_affectation']
        for r in all_recs
        if r.get('x_affectation') and r['x_affectation'] is not False
    })
    print(f"Nombre d'ouvrages distincts : {len(ouvrages)}")
    print("\nListe complète :")
    for o in ouvrages:
        print(f"  - {o}")
except Exception as e:
    print(f"ERREUR : {e}")
