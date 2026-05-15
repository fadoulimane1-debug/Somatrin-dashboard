"""
SOMA AI v2 — Intégration Odoo en temps réel + LLM local (llama3.2 via Ollama).

Architecture:
    OdooConnector     : connexion XML-RPC sécurisée (lit ODOO_* depuis settings)
    MetierDataService : 20+ méthodes par module métier
    SomaAIEngine      : détecte contexte → données Odoo → prompt LLM → réponse
"""
from __future__ import annotations

import json
import logging
import re
import ssl
import xmlrpc.client
from datetime import date
from typing import Any, Dict, List

import requests
from django.conf import settings

logger = logging.getLogger(__name__)

_VALID_MODEL_RE = re.compile(r'^[a-z][a-z0-9_.]*$')


def _month_range() -> tuple[str, str]:
    """Retourne (premier_jour_mois, aujourd'hui) au format ISO."""
    today = date.today()
    return today.replace(day=1).isoformat(), today.isoformat()


# ─────────────────────────────────────────────────────────────────────────────
class OdooConnector:
    """Connexion XML-RPC sécurisée à Odoo. Lit les credentials depuis settings."""

    def __init__(self) -> None:
        self.odoo_url: str = getattr(settings, 'ODOO_URL', 'http://localhost:8069').rstrip('/')
        self.odoo_db: str = getattr(settings, 'ODOO_DB', 'somatrin_PROD')
        self.odoo_user: str = getattr(settings, 'ODOO_USER', 'admin')
        self.odoo_pwd: str = getattr(settings, 'ODOO_PASS', '')
        self.uid: int | None = None
        self.common: Any = None
        self.models: Any = None
        self._connect()

    # ── Connexion ──────────────────────────────────────────────────────────

    def _make_proxy(self, path: str) -> xmlrpc.client.ServerProxy:
        url = f'{self.odoo_url}{path}'
        ssl_verify = getattr(settings, 'ODOO_SSL_VERIFY', True)
        if not ssl_verify and url.startswith('https://'):
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            return xmlrpc.client.ServerProxy(url, transport=xmlrpc.client.SafeTransport(context=ctx))
        return xmlrpc.client.ServerProxy(url)

    def _connect(self) -> None:
        try:
            self.common = self._make_proxy('/xmlrpc/2/common')
            self.models = self._make_proxy('/xmlrpc/2/object')
            self.uid = self.common.authenticate(
                self.odoo_db, self.odoo_user, self.odoo_pwd, {}
            )
            if not self.uid:
                raise ConnectionError('Authentification Odoo échouée — vérifiez ODOO_USER/ODOO_PASS')
            logger.info('✅ Connexion Odoo établie (uid=%s)', self.uid)
        except Exception as exc:
            logger.error('❌ Connexion Odoo impossible: %s', exc)
            self.uid = None

    # ── Requêtes ───────────────────────────────────────────────────────────

    def search_read(self, model: str, domain: list, fields: list, limit: int = 100) -> List[Dict]:
        """Requête Odoo sécurisée avec validation du nom de modèle."""
        if not _VALID_MODEL_RE.match(model):
            logger.warning('❌ Modèle Odoo invalide: %s', model)
            return []
        if not self.uid:
            logger.error('❌ Odoo non connecté — search_read ignoré (%s)', model)
            return []
        try:
            return self.models.execute_kw(
                self.odoo_db, self.uid, self.odoo_pwd,
                model, 'search_read',
                [domain],
                {'fields': fields, 'limit': limit},
            ) or []
        except Exception as exc:
            logger.error('❌ search_read(%s): %s', model, exc)
            return []

    def get_count(self, model: str, domain: list) -> int:
        """Nombre de records correspondant au domaine."""
        if not _VALID_MODEL_RE.match(model):
            logger.warning('❌ Modèle Odoo invalide: %s', model)
            return 0
        if not self.uid:
            return 0
        try:
            return int(self.models.execute_kw(
                self.odoo_db, self.uid, self.odoo_pwd,
                model, 'search_count',
                [domain],
            ) or 0)
        except Exception as exc:
            logger.error('❌ get_count(%s): %s', model, exc)
            return 0


# ─────────────────────────────────────────────────────────────────────────────
class MetierDataService:
    """Requêtes métier spécifiques par module. Toutes les méthodes retournent
    un dict standardisé: {count, total, items, message}."""

    def __init__(self, conn: OdooConnector) -> None:
        self.conn = conn

    # ── Achats ────────────────────────────────────────────────────────────

    def get_commandes_en_retard(self) -> Dict:
        today = date.today().isoformat()
        items = self.conn.search_read(
            'purchase.order',
            [('state', '=', 'purchase'), ('date_planned', '<', today)],
            ['name', 'partner_id', 'amount_total', 'date_planned', 'date_order'],
            limit=50,
        )
        count = len(items)
        total = sum(float(r.get('amount_total') or 0) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'⚠️ {count} commandes en retard | Montant total: {total:,.2f} DH',
        }

    def get_demandes_prix_en_attente(self) -> Dict:
        items = self.conn.search_read(
            'purchase.request.line',
            [('state', '=', 'to_approve')],
            ['name', 'product_id', 'qty_ordered', 'date_required', 'partner_id'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'📋 {count} demandes de prix en attente',
        }

    def get_bons_reception_attente(self) -> Dict:
        items = self.conn.search_read(
            'stock.picking',
            [('picking_type_code', '=', 'incoming'), ('state', 'in', ['assigned', 'confirmed'])],
            ['name', 'partner_id', 'date_done', 'scheduled_date', 'move_ids_count'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'📦 {count} réceptions en cours',
        }

    def get_fournisseurs_top(self) -> Dict:
        items = self.conn.search_read(
            'res.partner',
            [('supplier_rank', '>', 0)],
            ['name', 'email', 'phone', 'supplier_rank'],
            limit=20,
        )
        count = len(items)
        top5 = sorted(items, key=lambda r: r.get('supplier_rank', 0), reverse=True)[:5]
        return {
            'count': count,
            'total': 0.0,
            'items': top5,
            'message': f'🏢 {count} fournisseurs actifs | Top 5 listés',
        }

    def get_achat_stats_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'purchase.order',
            [('state', '=', 'purchase'), ('date_order', '>=', start), ('date_order', '<=', end)],
            ['name', 'partner_id', 'amount_total', 'date_order'],
            limit=200,
        )
        count = len(items)
        total = sum(float(r.get('amount_total') or 0) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'💰 Dépenses achat mois: {total:,.2f} DH | {count} PO',
        }

    # ── Transport & Logistique ────────────────────────────────────────────

    def get_bons_transport_en_cours(self) -> Dict:
        items = self.conn.search_read(
            'stock.picking',
            [('x_transport_logistics', '=', True), ('state', '=', 'assigned')],
            ['name', 'partner_id', 'date_done', 'move_ids_count'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'🚚 {count} bons de transport en cours',
        }

    def get_bons_transport_livres_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'stock.picking',
            [
                ('x_transport_logistics', '=', True),
                ('state', '=', 'done'),
                ('date_done', '>=', start),
                ('date_done', '<=', end),
            ],
            ['name', 'partner_id', 'date_done', 'move_ids_count'],
            limit=100,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'✅ {count} livraisons ce mois',
        }

    def get_couts_transport_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'account.move.line',
            [
                ('account_id.name', 'ilike', 'transport'),
                ('date', '>=', start),
                ('date', '<=', end),
                ('move_id.state', '=', 'posted'),
            ],
            ['name', 'balance', 'date', 'analytic_account_id'],
            limit=200,
        )
        count = len(items)
        total = sum(abs(float(r.get('balance') or 0)) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'💰 Coûts transport mois: {total:,.2f} DH | {count} lignes',
        }

    def get_revenus_transport_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'account.move',
            [
                ('move_type', 'in', ['out_invoice', 'out_refund']),
                ('state', '=', 'posted'),
                ('invoice_date', '>=', start),
                ('invoice_date', '<=', end),
            ],
            ['name', 'amount_total', 'invoice_date', 'partner_id', 'narration'],
            limit=200,
        )
        transport_items = [
            r for r in items
            if 'transport' in str(r.get('narration', '')).lower()
            or 'transport' in str(r.get('name', '')).lower()
        ] or items
        count = len(transport_items)
        total = sum(float(r.get('amount_total') or 0) for r in transport_items)
        return {
            'count': count,
            'total': total,
            'items': transport_items[:10],
            'message': f'📈 Recettes transport mois: {total:,.2f} DH',
        }

    def get_rentabilite_transport(self) -> Dict:
        revenus = self.get_revenus_transport_mois()
        couts = self.get_couts_transport_mois()
        r, c = revenus['total'], couts['total']
        profit = r - c
        taux = (profit / r * 100) if r > 0 else 0.0
        return {
            'count': 0,
            'total': profit,
            'items': [],
            'message': (
                f'📊 Rentabilité transport | Revenu: {r:,.2f} DH | '
                f'Coûts: {c:,.2f} DH | Profit: {profit:,.2f} DH | Taux: {taux:.1f}%'
            ),
            'revenus': r,
            'couts': c,
            'taux': taux,
        }

    # ── Parc & Maintenance ────────────────────────────────────────────────

    def get_equipements_actifs(self) -> Dict:
        items = self.conn.search_read(
            'maintenance.equipment',
            [('active', '=', True)],
            ['name', 'serial_no', 'category_id', 'technician_user_id', 'maintenance_date'],
            limit=100,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'🔧 {count} équipements en suivi',
        }

    def get_ordres_maintenance_en_cours(self) -> Dict:
        items = self.conn.search_read(
            'maintenance.request',
            [('stage_id.name', 'in', ['En cours', 'Assigné', 'In Progress', 'Assigned'])],
            ['name', 'equipment_id', 'user_id', 'priority', 'schedule_date'],
            limit=50,
        )
        count = len(items)
        urgent_count = sum(1 for r in items if str(r.get('priority', '')) == '3')
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'⚠️ {count} ordres en cours | 🔴 {urgent_count} urgent',
            'urgent_count': urgent_count,
        }

    def get_ordres_maintenance_retard(self) -> Dict:
        today = date.today().isoformat()
        items = self.conn.search_read(
            'maintenance.request',
            [
                ('stage_id.name', 'in', ['En cours', 'Assigné', 'In Progress', 'Assigned']),
                ('schedule_date', '<', today),
            ],
            ['name', 'equipment_id', 'priority', 'schedule_date'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'❌ {count} ordres en RETARD',
        }

    def get_couts_maintenance_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'account.move.line',
            [
                ('account_id.name', 'ilike', 'maintenance'),
                ('date', '>=', start),
                ('date', '<=', end),
                ('move_id.state', '=', 'posted'),
            ],
            ['name', 'balance', 'date'],
            limit=200,
        )
        count = len(items)
        total = sum(abs(float(r.get('balance') or 0)) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'💸 Coûts maintenance mois: {total:,.2f} DH',
        }

    # ── Gasoil ────────────────────────────────────────────────────────────

    def get_stock_gasoil_actuel(self) -> Dict:
        # Priorité au produit carburant réel (catégorie carburant + nom/code gasoil),
        # pour éviter de capter des pièces "filtre gasoil", "bouchon", etc.
        items = self.conn.search_read(
            'product.product',
            [('categ_id', '=', 262), '|', ('default_code', 'ilike', 'A04107'), ('name', 'ilike', 'GASOIL 10PPM')],
            ['name', 'qty_available', 'virtual_available', 'standard_price', 'default_code'],
            limit=1,
        )
        if not items:
            # Fallback plus large mais toujours limité à la catégorie carburant.
            items = self.conn.search_read(
                'product.product',
                [('categ_id', '=', 262), ('name', 'ilike', 'gasoil')],
                ['name', 'qty_available', 'virtual_available', 'standard_price', 'default_code'],
                limit=1,
            )
        if not items:
            return {
                'count': 0,
                'total': 0.0,
                'items': [],
                'message': '⛽ Produit gasoil non trouvé dans Odoo',
            }
        product = items[0]
        qty = float(product.get('qty_available') or 0)
        cost = float(product.get('standard_price') or 0)
        value = qty * cost
        return {
            'count': 1,
            'total': value,
            'items': items,
            'message': f'⛽ Stock gasoil: {qty:,.0f} L | Valeur: {value:,.2f} DH | Produit: {product.get("name", "—")}',
        }

    def get_consommation_gasoil_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'stock.move',
            [
                ('product_id.categ_id', '=', 262),
                ('state', '=', 'done'),
                ('date', '>=', start),
                ('date', '<=', end),
                ('picking_type_id.code', 'in', ['internal', 'outgoing']),
            ],
            ['name', 'quantity_done', 'date', 'picking_id', 'product_id'],
            limit=200,
        )
        count = len(items)
        total_qty = sum(float(r.get('quantity_done') or r.get('quantity', 0)) for r in items)
        return {
            'count': count,
            'total': total_qty,
            'items': items[:10],
            'message': f'📉 Consommation gasoil mois: {total_qty:,.0f} L',
        }

    def get_entrees_gasoil_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'stock.move',
            [
                ('product_id.categ_id', '=', 262),
                ('state', '=', 'done'),
                ('date', '>=', start),
                ('date', '<=', end),
                ('picking_type_id.code', '=', 'incoming'),
            ],
            ['name', 'quantity_done', 'date', 'picking_id', 'product_id'],
            limit=200,
        )
        count = len(items)
        total_qty = sum(float(r.get('quantity_done') or r.get('quantity', 0)) for r in items)
        return {
            'count': count,
            'total': total_qty,
            'items': items[:10],
            'message': f'📈 Entrées gasoil mois: {total_qty:,.0f} L',
        }

    # ── Production ────────────────────────────────────────────────────────

    def get_production_en_cours(self) -> Dict:
        items = self.conn.search_read(
            'stock.picking',
            [
                ('picking_type_code', '=', 'internal'),
                ('state', '=', 'assigned'),
                ('x_transport_logistics', '!=', True),
            ],
            ['name', 'date_done', 'move_ids_count', 'scheduled_date'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'🏭 {count} bons production en cours',
        }

    def get_couts_production_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'account.move.line',
            [
                ('account_id.name', 'ilike', 'production'),
                ('date', '>=', start),
                ('date', '<=', end),
                ('move_id.state', '=', 'posted'),
            ],
            ['name', 'balance', 'date'],
            limit=200,
        )
        count = len(items)
        total = sum(abs(float(r.get('balance') or 0)) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'💰 Coûts production mois: {total:,.2f} DH',
        }

    def get_revenus_production_mois(self) -> Dict:
        start, end = _month_range()
        items = self.conn.search_read(
            'account.move',
            [
                ('move_type', 'in', ['out_invoice', 'out_refund']),
                ('state', '=', 'posted'),
                ('invoice_date', '>=', start),
                ('invoice_date', '<=', end),
            ],
            ['name', 'amount_total', 'invoice_date', 'partner_id'],
            limit=200,
        )
        count = len(items)
        total = sum(float(r.get('amount_total') or 0) for r in items)
        return {
            'count': count,
            'total': total,
            'items': items[:10],
            'message': f'📈 Revenus production mois: {total:,.2f} DH',
        }

    def get_rentabilite_production(self) -> Dict:
        revenus = self.get_revenus_production_mois()
        couts = self.get_couts_production_mois()
        r, c = revenus['total'], couts['total']
        profit = r - c
        taux = (profit / r * 100) if r > 0 else 0.0
        return {
            'count': 0,
            'total': profit,
            'items': [],
            'message': (
                f'📊 Rentabilité production | Revenu: {r:,.2f} DH | '
                f'Coûts: {c:,.2f} DH | Profit: {profit:,.2f} DH | Taux: {taux:.1f}%'
            ),
        }

    # ── QHSE ─────────────────────────────────────────────────────────────

    def get_alertes_qhse_en_cours(self) -> Dict:
        items = self.conn.search_read(
            'quality.alert',
            [('stage_id.name', 'in', ['Ouvert', 'En cours', 'Open', 'In Progress'])],
            ['name', 'stage_id', 'date_time', 'user_id'],
            limit=50,
        )
        count = len(items)
        return {
            'count': count,
            'total': 0.0,
            'items': items[:10],
            'message': f'⚠️ {count} alertes QHSE ouvertes',
        }


# ─────────────────────────────────────────────────────────────────────────────
class SomaAIEngine:
    """Orchestrateur: détecte contexte → requête Odoo → prompt LLM → réponse."""

    _KEYWORDS: Dict[str, List[str]] = {
        'achats': [
            'commande', 'fournisseur', 'achat', ' po ', 'rfq', 'devis', 'prix',
            'réception', 'approvisionnement', 'bon de commande',
        ],
        'transport': [
            'transport', 'livraison', 'bon transport', 'logistique',
            'véhicule', 'camion', 'trajet', 'tournée',
        ],
        'parc': [
            'équipement', 'maintenance', 'panne', 'réparation',
            'ordre maintenance', 'parc', 'engin', 'machine',
            # Peut concerner une fiche équipement (Parc), pas seulement les achats
            'fournisseur',
        ],
        'gasoil': [
            'gasoil', 'carburant', 'essence', 'consommation carburant',
            'plein', 'litre', 'stock carburant',
        ],
        'production': [
            'production', 'fabrication', 'bon travail',
            'carrière', 'extraction', 'chantier',
        ],
        'qhse': [
            'qualité', 'sécurité', 'alerte', 'hygiène',
            'environnement', 'qhse', 'incident', 'non-conformité',
        ],
    }

    def __init__(self, connector: OdooConnector, metier: MetierDataService) -> None:
        self.connector = connector
        self.metier = metier
        self.ollama_url = getattr(settings, 'OLLAMA_BASE_URL', 'http://127.0.0.1:11434').rstrip('/')
        self.ollama_model = getattr(settings, 'OLLAMA_MODEL', 'llama3.2')
        self.ollama_timeout = int(getattr(settings, 'OLLAMA_TIMEOUT_SEC', 60))

    def detect_context(self, question: str) -> str:
        q = question.lower()
        scores = {ctx: sum(1 for kw in kws if kw in q) for ctx, kws in self._KEYWORDS.items()}
        best = max(scores, key=lambda k: scores[k])
        return best if scores[best] > 0 else 'general'

    def get_odoo_context(self, question: str, context: str) -> Dict:
        if context == 'achats':
            return {
                'retard': self.metier.get_commandes_en_retard(),
                'rfq': self.metier.get_demandes_prix_en_attente(),
                'reception': self.metier.get_bons_reception_attente(),
                'fournisseurs': self.metier.get_fournisseurs_top(),
                'stats_mois': self.metier.get_achat_stats_mois(),
            }
        if context == 'transport':
            return {
                'en_cours': self.metier.get_bons_transport_en_cours(),
                'livres_mois': self.metier.get_bons_transport_livres_mois(),
                'couts': self.metier.get_couts_transport_mois(),
                'revenus': self.metier.get_revenus_transport_mois(),
                'rentabilite': self.metier.get_rentabilite_transport(),
            }
        if context == 'parc':
            return {
                'equipements': self.metier.get_equipements_actifs(),
                'ordres_en_cours': self.metier.get_ordres_maintenance_en_cours(),
                'ordres_retard': self.metier.get_ordres_maintenance_retard(),
                'couts': self.metier.get_couts_maintenance_mois(),
            }
        if context == 'gasoil':
            return {
                'stock': self.metier.get_stock_gasoil_actuel(),
                'consommation': self.metier.get_consommation_gasoil_mois(),
                'entrees': self.metier.get_entrees_gasoil_mois(),
            }
        if context == 'production':
            return {
                'en_cours': self.metier.get_production_en_cours(),
                'couts': self.metier.get_couts_production_mois(),
                'rentabilite': self.metier.get_rentabilite_production(),
            }
        if context == 'qhse':
            return {'alertes': self.metier.get_alertes_qhse_en_cours()}
        return {}

    def build_ollama_prompt(self, question: str, odoo_data: Dict) -> str:
        today = date.today().strftime('%d/%m/%Y')
        return (
            f"Tu es SOMA AI, assistant interne pour les collaborateurs SOMATRIN.\n\n"
            f"DATE AUJOURD'HUI: {today}\n\n"
            f"DONNÉES ODOO EN TEMPS RÉEL:\n"
            f"{json.dumps(odoo_data, indent=2, default=str, ensure_ascii=False)}\n\n"
            f"QUESTION: {question}\n\n"
            f"INSTRUCTIONS:\n"
            f"- Réponds UNIQUEMENT en français, de façon concise et professionnelle\n"
            f"- Utilise les CHIFFRES RÉELS ci-dessus (ne les invente pas)\n"
            f"- Formate les montants avec séparateurs: ex. 12 345,50 DH\n"
            f"- Utilise des emojis métier: 💼 📋 ⚠️ 🚚 🔧 ⛽ 🏭 💰 📊 ✅ ❌\n"
            f"- Priorise les éléments urgents (retards, anomalies, alertes)\n"
            f"- Si une section est vide (count=0), mentionne-le clairement\n\n"
            f"RÉPONSE:"
        )

    def _is_top_fournisseur_question(self, question: str) -> bool:
        q = (question or '').lower()
        keys = (
            'top fournisseur',
            'meilleur fournisseur',
            'fournisseur top',
            'classement fournisseur',
            'top 5 fournisseur',
            'top fournisseurs',
        )
        return any(k in q for k in keys)

    def _is_retard_commandes_question(self, question: str) -> bool:
        q = (question or '').lower()
        keys = (
            'commandes en retard',
            'commande en retard',
            'retard commande',
            'bons en retard',
            'bon en retard',
            'po en retard',
            'achat en retard',
        )
        return any(k in q for k in keys)

    def _is_demandes_prix_question(self, question: str) -> bool:
        q = (question or '').lower()
        keys = (
            'demande de prix',
            'demandes de prix',
            'rfq',
            'devis en attente',
            'prix en attente',
            'demandes prix en attente',
        )
        return any(k in q for k in keys)

    def _format_dh(self, val: float) -> str:
        s = f'{float(val or 0):,.2f}'
        return s.replace(',', ' ').replace('.', ',') + ' DH'

    def _build_top_fournisseur_response(self, odoo_data: Dict) -> str:
        """Réponse déterministe pour éviter les contradictions LLM sur ce cas."""
        achats = odoo_data.get('stats_mois') or {}
        retard = odoo_data.get('retard') or {}
        rows = achats.get('items') or []

        spend_by_vendor: Dict[str, float] = {}
        for po in rows:
            partner = po.get('partner_id')
            if isinstance(partner, (list, tuple)) and len(partner) > 1:
                name = str(partner[1] or 'Fournisseur inconnu').strip() or 'Fournisseur inconnu'
            else:
                name = 'Fournisseur inconnu'
            amt = float(po.get('amount_total') or 0)
            spend_by_vendor[name] = spend_by_vendor.get(name, 0.0) + amt

        ordered = sorted(spend_by_vendor.items(), key=lambda kv: kv[1], reverse=True)
        top = ordered[:5]

        lines = []
        retard_count = int(retard.get('count') or 0)
        retard_total = float(retard.get('total') or 0)
        lines.append(
            f'⚠️ Commandes en retard: {retard_count} | Montant total: {self._format_dh(retard_total)}'
        )
        lines.append('')
        lines.append('🏆 Top fournisseurs (dépenses achats du mois, tri décroissant):')

        if not top:
            lines.append('• Aucune dépense achat disponible sur la période.')
        else:
            for idx, (name, amt) in enumerate(top, 1):
                lines.append(f'{idx}. {name} — {self._format_dh(amt)}')

        lines.append('')
        lines.append('✅ Source: agrégation des bons de commande Odoo du mois en cours.')
        return '\n'.join(lines)

    def _build_commandes_retard_response(self, odoo_data: Dict) -> str:
        retard = odoo_data.get('retard') or {}
        items = retard.get('items') or []
        count = int(retard.get('count') or 0)
        total = float(retard.get('total') or 0)

        lines = [
            f'⚠️ Commandes en retard: {count} | Montant total: {self._format_dh(total)}',
            '',
        ]

        if not items:
            lines.append('✅ Aucune commande en retard dans les éléments remontés.')
        else:
            lines.append('📋 Principales commandes en retard:')
            for idx, po in enumerate(items[:10], 1):
                name = str(po.get('name') or '—')
                partner = po.get('partner_id')
                vendor = partner[1] if isinstance(partner, (list, tuple)) and len(partner) > 1 else 'Fournisseur inconnu'
                amt = float(po.get('amount_total') or 0)
                lines.append(f'{idx}. {name} ({vendor}) — {self._format_dh(amt)}')

        lines.append('')
        lines.append('✅ Source: purchase.order (state=purchase, date_planned < aujourd’hui).')
        return '\n'.join(lines)

    def _build_demandes_prix_response(self, odoo_data: Dict) -> str:
        rfq = odoo_data.get('rfq') or {}
        items = rfq.get('items') or []
        count = int(rfq.get('count') or 0)

        lines = [f'📋 Demandes de prix en attente: {count}', '']
        if not items:
            lines.append('✅ Aucune demande de prix en attente actuellement.')
        else:
            lines.append('🧾 Principales lignes en attente:')
            for idx, ln in enumerate(items[:10], 1):
                name = str(ln.get('name') or '—')
                product = ln.get('product_id')
                product_name = product[1] if isinstance(product, (list, tuple)) and len(product) > 1 else 'Produit non renseigné'
                qty = ln.get('qty_ordered')
                try:
                    qty_s = f'{float(qty or 0):,.2f}'.replace(',', ' ').replace('.', ',')
                except Exception:
                    qty_s = str(qty or '0')
                lines.append(f'{idx}. {name} — {product_name} | Qté: {qty_s}')

        lines.append('')
        lines.append('✅ Source: purchase.request.line (state=to_approve).')
        return '\n'.join(lines)

    def generate_direct_response(self, question: str, odoo_data: dict) -> str:
        """Fallback sans LLM: formate les données Odoo directement quand Ollama est indisponible."""
        lines = ['📊 Données temps réel Odoo (réponse directe — IA locale indisponible)\n']
        if not odoo_data:
            lines.append('Aucune donnée Odoo disponible pour cette question.')
            return '\n'.join(lines)
        for section, items in odoo_data.items():
            if not items:
                continue
            lines.append(f'{section.replace("_", " ").title()}:')
            if isinstance(items, dict):
                msg = items.get('message')
                if msg:
                    lines.append(f'- {msg}')
                count = items.get('count')
                total = items.get('total')
                if count not in (None, False, ''):
                    lines.append(f'- Nombre: {count}')
                if total not in (None, False, '') and isinstance(total, (int, float)):
                    lines.append(f'- Total: {total:,.2f}')
                # Montrer au plus 3 lignes lisibles, pas de dump JSON brut.
                sample = items.get('items') if isinstance(items.get('items'), list) else []
                for it in sample[:3]:
                    if not isinstance(it, dict):
                        continue
                    name = it.get('name') or it.get('product_id') or it.get('picking_id') or '—'
                    qty = it.get('quantity_done') or it.get('qty_available') or ''
                    datev = it.get('date') or ''
                    mini = f"- {name}"
                    if qty not in ('', None):
                        mini += f" | Qté: {qty}"
                    if datev:
                        mini += f" | Date: {str(datev)[:10]}"
                    lines.append(mini)
            elif isinstance(items, list):
                lines.append(f'- {len(items)} enregistrement(s)')
            else:
                lines.append(f'- {items}')
            lines.append('')
        lines.append('Pour des analyses détaillées, reposez la question dans quelques instants.')
        return '\n'.join(lines)

    def query_ollama(self, prompt: str, odoo_data: dict | None = None) -> str:
        try:
            resp = requests.post(
                f'{self.ollama_url}/api/generate',
                json={
                    'model': self.ollama_model,
                    'prompt': prompt,
                    'stream': False,
                    'options': {'temperature': 0.1},
                },
                timeout=self.ollama_timeout,
            )
            if resp.status_code == 200:
                return resp.json().get('response', '').strip()
            return f'❌ Erreur Ollama: HTTP {resp.status_code}'
        except requests.Timeout:
            logger.warning('Ollama timeout après %ss — réponse directe depuis données Odoo', self.ollama_timeout)
            return self.generate_direct_response(prompt, odoo_data or {})
        except Exception as exc:
            logger.error('❌ query_ollama: %s', exc)
            return f'❌ Erreur IA locale: {exc}'

    def process_question(self, question: str, page_context: str = 'general', page_data: dict | None = None) -> str:
        """Pipeline complet: contexte → données → LLM → réponse.

        Priorité des sources:
          1. page_data (KPIs de la page courante envoyés par le front) → Odoo ignoré, cohérence garantie
          2. Odoo temps réel via get_odoo_context() → si pas de page_data
        """
        _valid = {'achats', 'transport', 'parc', 'gasoil', 'production', 'qhse'}
        detected_context = self.detect_context(question)

        # ── Contexte prioritaire depuis la page (plus fiable que le simple NLP)
        # Evite de basculer en "achats" sur une fiche parc juste parce que la question contient "fournisseur".
        try:
            if page_data and isinstance(page_data, dict):
                p = page_data.get('page') if isinstance(page_data.get('page'), dict) else {}
                path = p.get('path') or p.get('full_path') or ''
                if isinstance(path, str) and '/parc/' in path:
                    page_context = 'parc'
                    ql = (question or '').lower()
                    explicit_other = any(w in ql for w in (
                        'achats', 'bon de commande', 'bons de commande', 'rfq',
                        'transport', 'production', 'gasoil', 'qhse',
                    ))
                    if not explicit_other:
                        detected_context = 'parc'
        except Exception:
            pass

        # ── Mode inter-services:
        # Si la question cible explicitement un autre service, on ne se limite PAS aux KPI de page.
        cross_service_target = detected_context in _valid and detected_context != page_context

        # ── Chemin rapide local: données de page uniquement quand la question concerne la page courante.
        if page_data and not cross_service_target:
            logger.info(
                '📄 SOMA AI v2: réponse déterministe depuis données de page (%s clés), context page=%s, detect=%s',
                len(page_data), page_context, detected_context
            )
            response = self._build_page_data_response(question, page_data)
            logger.info('✅ Réponse SOMA AI v2 (page déterministe) générée (%s chars)', len(response))
            return response

        # ── Chemin Odoo: pas de page_data → requêter Odoo ───────────────────────────
        if not self.connector.uid:
            return (
                '❌ Odoo indisponible — impossible de récupérer les données en temps réel.\n'
                'Vérifiez la connexion à Odoo (ODOO_URL, ODOO_USER, ODOO_PASS).'
            )

        if detected_context in _valid:
            context = detected_context
            logger.info(
                '🤖 SOMA AI v2: contexte=%s (question detectée, inter-services=%s) | question=%s chars',
                context, cross_service_target, len(question)
            )
        elif page_context in _valid:
            context = page_context
            logger.info('🤖 SOMA AI v2: contexte=%s (via page) | question=%s chars', context, len(question))
        else:
            context = detected_context
            logger.info('🤖 SOMA AI v2: contexte=%s (via NLP) | question=%s chars', context, len(question))

        if context == 'general':
            return (
                "Je n'ai pas pu identifier le module concerné par votre question.\n"
                "Précisez le contexte souhaité : **Achats**, **Transport**, **Parc/Maintenance**, "
                "**Gasoil**, **Production** ou **QHSE**."
            )

        odoo_data = self.get_odoo_context(question, context)
        logger.info('📊 Données Odoo récupérées: %s sections', len(odoo_data))

        # Cas métiers critiques: réponses déterministes (évite contradictions/timeouts LLM).
        if context == 'achats':
            if self._is_top_fournisseur_question(question):
                response = self._build_top_fournisseur_response(odoo_data)
                logger.info('✅ Réponse déterministe top fournisseur (%s chars)', len(response))
                return response
            if self._is_retard_commandes_question(question):
                response = self._build_commandes_retard_response(odoo_data)
                logger.info('✅ Réponse déterministe commandes retard (%s chars)', len(response))
                return response
            if self._is_demandes_prix_question(question):
                response = self._build_demandes_prix_response(odoo_data)
                logger.info('✅ Réponse déterministe demandes prix (%s chars)', len(response))
                return response

        prompt = self.build_ollama_prompt(question, odoo_data)
        response = self.query_ollama(prompt, odoo_data)
        logger.info('✅ Réponse SOMA AI v2 (Odoo) générée (%s chars)', len(response))
        return response

    def _build_page_data_response(self, question: str, page_data: dict) -> str:
        """Réponse déterministe depuis KPI de page (évite hallucinations)."""
        q = (question or '').lower()
        page = page_data.get('page') if isinstance(page_data.get('page'), dict) else {}
        details = page_data.get('page_details') if isinstance(page_data.get('page_details'), dict) else {}
        kpis = page_data.get('kpis') if isinstance(page_data.get('kpis'), dict) else page_data
        if not isinstance(kpis, dict):
            kpis = {}

        def _normalize_key(txt: str) -> str:
            t = (txt or '').lower()
            t = t.replace('é', 'e').replace('è', 'e').replace('ê', 'e').replace('à', 'a').replace('ù', 'u').replace('ô', 'o')
            return t

        def _extract_num(val: Any) -> str:
            s = str(val or '').strip()
            m = re.search(r'[-+]?\d[\d\s.,]*', s)
            return m.group(0).strip() if m else s

        def _find_detail_value(keys: list[str]) -> str | None:
            if not isinstance(details, dict) or not details:
                return None
            target_norm = [_normalize_key(k) for k in keys]
            for dk, dv in details.items():
                ndk = _normalize_key(str(dk))
                if any(t in ndk for t in target_norm):
                    return str(dv).strip()
            return None

        # Cas frequent: question sur filtres actifs
        if 'filtre' in q:
            f = page_data.get('filters_actifs')
            if isinstance(f, dict) and f:
                lines = ['Filtres actuellement appliques :']
                for k, v in f.items():
                    label = str(k).replace('_', ' ').title()
                    lines.append(f'- {label}: {v}')
                return '\n'.join(lines)
            return "Aucun filtre actif detecte sur cette page."

        if 'fournisseur' in q:
            val = _find_detail_value(['fournisseur', 'supplier'])
            if val:
                return f"Fournisseur (selon l'écran courant): {val}."
            return "Je ne vois pas le fournisseur dans les informations affichées sur cette page."

        if any(k in q for k in ('derniere maintenance', 'dernière maintenance', 'last maintenance')):
            val = _find_detail_value(['derniere maintenance', 'maintenance'])
            if val:
                return f"Dernière maintenance (selon l'écran courant): {val}."
            for k, v in kpis.items():
                nk = _normalize_key(str(k))
                if 'derniere' in nk and 'maint' in nk:
                    return f"Dernière maintenance (selon l'écran courant): {v}."
            return "Je ne vois pas la date de dernière maintenance dans les informations affichées sur cette page."

        # Question type: "sur cette page exacte" (contexte uniquement)
        if any(k in q for k in ('page exacte', 'contexte page', 'cette page', 'cet ecran', 'cet écran', 'ecran', 'écran')):
            lines = ['📍 Contexte page:']
            if page.get('url'):
                lines.append(f"- URL: {page.get('url')}")
            if page.get('breadcrumb'):
                lines.append(f"- Navigation: {page.get('breadcrumb')}")
            ent = page.get('entity') if isinstance(page.get('entity'), dict) else None
            if ent:
                et = ent.get('type')
                eid = ent.get('id')
                if et and eid is not None:
                    lines.append(f"- Cible: {et} #{eid}")
            return '\n'.join(lines)

        # Cas fréquent: "Combien d'équipements ?"
        if ('combien' in q or 'nombre' in q or 'total' in q) and ('equip' in q):
            for key in kpis.keys():
                nk = _normalize_key(str(key))
                if 'total' in nk and 'equip' in nk:
                    v = _extract_num(kpis.get(key))
                    return f"Total equipements (page actuelle): {v}."
            for key in kpis.keys():
                nk = _normalize_key(str(key))
                if 'equip' in nk:
                    v = _extract_num(kpis.get(key))
                    return f"{str(key).replace('_', ' ').title()}: {v}."
            return "Je ne trouve pas l'indicateur équipements dans les données visibles de cette page."

        # Réponse générique déterministe depuis les KPI visibles
        lines = ["Donnees visibles de la page (source directe):", ""]
        if page.get('breadcrumb'):
            lines.append(f"- Navigation: {page.get('breadcrumb')}")
        if page.get('url'):
            lines.append(f"- URL: {page.get('url')}")
        if details:
            lines.append("")
            lines.append("Informations detectees sur l'ecran:")
            for dk, dv in list(details.items())[:8]:
                lines.append(f"- {str(dk).strip()}: {str(dv).strip()}")
        shown = 0
        for key, val in kpis.items():
            if val in (None, '', False):
                continue
            label = str(key).replace('_', ' ').title()
            lines.append(f"- {label}: {val}")
            shown += 1
            if shown >= 12:
                break
        if shown == 0:
            return "Aucune donnée KPI exploitable n'est disponible sur la page courante."
        lines.append("")
        lines.append("Reponse deterministe basee uniquement sur les indicateurs affiches.")
        return "\n".join(lines)
