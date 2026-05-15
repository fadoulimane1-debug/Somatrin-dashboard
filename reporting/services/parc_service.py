import logging
import ssl
import xmlrpc.client

from django.conf import settings


logger = logging.getLogger(__name__)


class ParcOdooService:
    """Service Odoo pour le module Parc."""

    def __init__(self):
        transport = getattr(settings, 'ODOO_TRANSPORT', 'xmlrpc')
        context = None
        if transport == 'xmlrpcs':
            context = ssl._create_unverified_context()
        self.common = xmlrpc.client.ServerProxy(
            f"{settings.ODOO_URL}/{transport}/2/common",
            context=context,
            allow_none=True,
        )
        self.models = xmlrpc.client.ServerProxy(
            f"{settings.ODOO_URL}/{transport}/2/object",
            context=context,
            allow_none=True,
        )
        self.odoo_db = settings.ODOO_DB
        self.odoo_user = settings.ODOO_USER
        self.odoo_pwd = settings.ODOO_PASS
        self.uid = self.common.authenticate(
            self.odoo_db,
            self.odoo_user,
            self.odoo_pwd,
            {},
        )

    def get_equipment_categories(self):
        """Recupere toutes les categories d'equipements depuis Odoo."""
        try:
            categories = self.models.execute_kw(
                self.odoo_db,
                self.uid,
                self.odoo_pwd,
                'maintenance.equipment.category',
                'search_read',
                [[]],
                {
                    'fields': ['id', 'name'],
                    'order': 'name',
                },
            )
            return categories or []
        except Exception as e:
            logger.error(f"Erreur categories equipements: {e}")
            return []

    def get_ordres_maintenance(self, limit=500):
        """Récupère les ordres de maintenance avec tous les champs de coût disponibles."""
        try:
            fields_meta = self.models.execute_kw(
                self.odoo_db, self.uid, self.odoo_pwd,
                'maintenance.request', 'fields_get', [], {'attributes': ['type']},
            ) or {}
            available = set(fields_meta.keys())

            base_fields = ['id', 'name', 'equipment_id', 'maintenance_type',
                           'stage_id', 'category_id', 'duration']
            date_field = next((f for f in ('request_date', 'create_date') if f in available), None)
            tech_field = next((f for f in ('technician_user_id', 'owner_user_id') if f in available), None)
            if date_field:
                base_fields.append(date_field)
            if tech_field:
                base_fields.append(tech_field)

            cost_priority = ('cost_amount', 'total_cost', 'amount_total',
                             'maintenance_cost', 'effective_cost', 'cost')
            cost_parts = ('parts_cost', 'labor_cost', 'spare_parts_cost')
            cost_fields_used = [f for f in (*cost_priority, *cost_parts) if f in available]
            read_fields = sorted(set(base_fields + cost_fields_used))

            ordres = self.models.execute_kw(
                self.odoo_db, self.uid, self.odoo_pwd,
                'maintenance.request', 'search_read',
                [[]],
                {'fields': read_fields, 'limit': limit,
                 'order': f'{date_field} desc' if date_field else 'id desc'},
            ) or []

            for ordre in ordres:
                dur = float(ordre.get('duration') or 0)
                final_cost = 0.0
                for cf in cost_priority:
                    if cf in cost_fields_used:
                        v = float(ordre.get(cf) or 0)
                        if v > 0:
                            final_cost = v
                            break
                if not final_cost:
                    parts_sum = sum(float(ordre.get(f) or 0) for f in cost_parts if f in cost_fields_used)
                    final_cost = parts_sum if parts_sum > 0 else dur * 150
                ordre['final_cost'] = round(final_cost, 2)

            return ordres
        except Exception as e:
            logger.error(f"Erreur get_ordres_maintenance: {e}")
            return []

    def get_suppliers(self):
        """Recupere les fournisseurs depuis Odoo."""
        try:
            # Compatibilite Odoo: certains environnements utilisent supplier,
            # d'autres supplier_rank.
            try:
                domain = [['supplier', '=', True]]
                suppliers = self.models.execute_kw(
                    self.odoo_db,
                    self.uid,
                    self.odoo_pwd,
                    'res.partner',
                    'search_read',
                    [domain],
                    {'fields': ['id', 'name'], 'order': 'name', 'limit': 200},
                )
            except Exception:
                domain = [['supplier_rank', '>', 0]]
                suppliers = self.models.execute_kw(
                    self.odoo_db,
                    self.uid,
                    self.odoo_pwd,
                    'res.partner',
                    'search_read',
                    [domain],
                    {'fields': ['id', 'name'], 'order': 'name', 'limit': 200},
                )
            return suppliers or []
        except Exception as e:
            logger.error(f"Erreur fournisseurs: {e}")
            return []
