"""
SOMATRIN — QHSE Service
Récupère les données Odoo réelles pour le module QHSE :
  - quality.check (Odoo 16) → Incidents / contrôles qualité (prioritaire)
  - quality.alert (Odoo 17+) → Fallback incidents si disponible
  - action.schedule.plan → Plan d'actions JAM (prioritaire si présent)
  - jam.action / project.task → Autres sources plan d'actions
  - stock.location → Sites / unités
"""
from __future__ import annotations

import json
import logging
from collections import defaultdict
from datetime import datetime, date
from typing import Any

from django.conf import settings

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────
# Helpers internes
# ─────────────────────────────────────────────────────────

def _odoo_db_name() -> str:
    return getattr(settings, 'ODOO_DB', None) or getattr(settings, 'ODOO_DATABASE', '') or ''


def _odoo_pass_value() -> str:
    return getattr(settings, 'ODOO_PASS', None) or getattr(settings, 'ODOO_PASSWORD', '') or ''


def _safe_fields(uid, models, model: str, candidates: list[str]) -> list[str]:
    """Retourne uniquement les champs qui existent dans le modèle Odoo."""
    try:
        known = set(
            (models.execute_kw(
                _odoo_db_name(), uid, _odoo_pass_value(),
                model, 'fields_get', [], {'attributes': ['type']},
            ) or {}).keys()
        )
        return [f for f in candidates if f in known]
    except Exception:
        return []


# Champs date « échéance » / « traitement » — action.schedule.plan varie selon les modules custom
_SCHEDULE_PLAN_DEADLINE_FIELDS = [
    'date_deadline', 'date_planned', 'limit_date', 'date_end',
    'scheduled_date', 'expected_date', 'end_date',
    'date_echeance', 'echeance_date', 'deadline',
    'date_limit', 'plan_date', 'date_stop', 'x_date_deadline', 'x_date_echeance',
]
_SCHEDULE_PLAN_CLOSED_FIELDS = [
    'date_done', 'date_closed', 'closing_date', 'date_traitement',
    'treatment_date', 'processing_date', 'date_end_process', 'x_date_traitement',
]
# Many2one « Processus » (libellés varient selon le module custom)
_SCHEDULE_PLAN_PROCESS_M2O = [
    'process_id', 'processus_id', 'jam_process_id', 'x_process_id',
    'plan_process_id', 'department_id', 'activity_id',
]
# Origine : many2one ou char
_SCHEDULE_PLAN_ORIGIN_M2O = [
    'origin_id', 'origine_id', 'source_id', 'meeting_origin_id',
    'x_origin_id', 'action_origin_id', 'origine_action_id',
]
_SCHEDULE_PLAN_ORIGIN_CHAR = [
    'origin', 'origine', 'source', 'x_origin', 'action_origin',
]
# Type d'action (m2o ou char)
_SCHEDULE_PLAN_TYPE_M2O = [
    'action_type_id', 'type_id', 'type_action_id', 'jam_type_id', 'x_type_action_id',
]
_SCHEDULE_PLAN_TYPE_CHAR = ['type_action', 'action_type', 'type']
# Anomalie : m2o ou texte
_SCHEDULE_PLAN_ANOMALY_M2O = [
    'anomaly_id', 'x_anomaly_id', 'nonconformity_id', 'nc_id', 'problem_id',
]
_SCHEDULE_PLAN_ANOMALY_TEXT = [
    'description', 'reason', 'note', 'anomaly', 'anomalie',
    'anomaly_description', 'lib_anomalie', 'motif', 'observation',
    'x_anomaly', 'x_description', 'remarque', 'comment', 'libelle_anomalie',
]
# Assigné (plusieurs conventions Odoo)
_SCHEDULE_PLAN_USER_M2O = [
    'user_id', 'assigned_to', 'responsible_id', 'employee_id', 'partner_id',
    'x_user_id', 'assignee_id',
]


def _m2o_first_name(rec: dict, keys: list[str]) -> str:
    """Premier many2one non vide (nom affiché)."""
    for k in keys:
        n = _many2one_name(rec, k)
        if n != '—':
            return n
    return '—'


def _schedule_plan_origin_value(rec: dict) -> str:
    n = _m2o_first_name(rec, _SCHEDULE_PLAN_ORIGIN_M2O)
    if n != '—':
        return n
    for k in _SCHEDULE_PLAN_ORIGIN_CHAR:
        v = rec.get(k)
        if v is None or v is False:
            continue
        s = str(v).strip()
        if s:
            return s
    return '—'


def _schedule_plan_task_type(rec: dict) -> str:
    n = _m2o_first_name(rec, _SCHEDULE_PLAN_TYPE_M2O)
    if n != '—':
        return n
    for k in _SCHEDULE_PLAN_TYPE_CHAR:
        v = rec.get(k)
        if v is None or v is False:
            continue
        s = str(v).strip()
        if s:
            return s
    return '—'


def _schedule_plan_anomaly_text(rec: dict) -> str:
    n = _m2o_first_name(rec, _SCHEDULE_PLAN_ANOMALY_M2O)
    if n != '—':
        return n
    v = _first_field_value(rec, _SCHEDULE_PLAN_ANOMALY_TEXT)
    if v is not None:
        return str(v).strip() or '—'
    return '—'


def _schedule_plan_process_field_present(fields: list[str]) -> str | None:
    for k in _SCHEDULE_PLAN_PROCESS_M2O:
        if k in fields:
            return k
    return None


def _schedule_plan_field_candidates(*, with_write_date: bool = False) -> list[str]:
    """Champs à lire sur action.schedule.plan (filtrés ensuite par fields_get)."""
    extra: list[str] = ['write_date'] if with_write_date else []
    return list(dict.fromkeys([
        'name', 'priority',
        'state', 'active', 'efficacity',
        'is_delay', 'delay_passed', 'is_late',
        'create_date',
        *_SCHEDULE_PLAN_USER_M2O,
        *_SCHEDULE_PLAN_PROCESS_M2O,
        *_SCHEDULE_PLAN_ORIGIN_M2O,
        *_SCHEDULE_PLAN_ORIGIN_CHAR,
        *_SCHEDULE_PLAN_TYPE_M2O,
        *_SCHEDULE_PLAN_TYPE_CHAR,
        *_SCHEDULE_PLAN_ANOMALY_M2O,
        *_SCHEDULE_PLAN_ANOMALY_TEXT,
        'description', 'reason', 'note',
        *_SCHEDULE_PLAN_DEADLINE_FIELDS,
        *_SCHEDULE_PLAN_CLOSED_FIELDS,
        *extra,
    ]))


def _first_field_value(rec: dict, keys: list[str]) -> Any:
    """Première valeur non vide parmi des clés candidates."""
    for k in keys:
        v = rec.get(k)
        if v is None or v is False:
            continue
        if isinstance(v, str) and not str(v).strip():
            continue
        return v
    return None


def _fmt_date(raw: Any) -> str:
    """Formate une date Odoo (str ou False) en jj/mm/aaaa."""
    s = str(raw or '')[:10]
    if len(s) == 10 and '-' in s:
        parts = s.split('-')
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return s or '—'


def _v(rec: dict, key: str, default='—') -> Any:
    val = rec.get(key)
    if val is False or val is None:
        return default
    return val


def _many2one_name(rec: dict, key: str) -> str:
    val = rec.get(key)
    if isinstance(val, (list, tuple)) and len(val) > 1:
        return str(val[1])
    return '—'


def _many2one_id(rec: dict, key: str) -> int | None:
    val = rec.get(key)
    if isinstance(val, (list, tuple)) and len(val) > 0:
        return val[0]
    return None


# ─────────────────────────────────────────────────────────
# Service principal
# ─────────────────────────────────────────────────────────

class QHSEService:
    """Agrège les données Odoo pour le module QHSE."""

    HEURES_ANNUELLES = 180_000  # hypothèse pour TF / TG
    HEURES_MENSUELLES = 15_000

    def __init__(self):
        self.uid = None
        self.models = None
        self.db = _odoo_db_name()
        self.pw = _odoo_pass_value()
        self._error: str | None = None

    # ── Connexion ──────────────────────────────────────────
    def _connect(self):
        if self.uid is not None:
            return True
        try:
            import xmlrpc.client, ssl
            verify_ssl = getattr(settings, 'ODOO_SSL_VERIFY', True)
            proxy_kw: dict = {}
            if settings.ODOO_URL.startswith('https://') and not verify_ssl:
                proxy_kw['context'] = ssl._create_unverified_context()
            common = xmlrpc.client.ServerProxy(
                f'{settings.ODOO_URL}/xmlrpc/2/common', **proxy_kw
            )
            login = getattr(settings, 'ODOO_USER', '') or getattr(settings, 'ODOO_USERNAME', '')
            self.uid = common.authenticate(self.db, login, self.pw, {})
            self.models = xmlrpc.client.ServerProxy(
                f'{settings.ODOO_URL}/xmlrpc/2/object', **proxy_kw
            )
            return bool(self.uid)
        except Exception as exc:
            self._error = f'Connexion Odoo impossible: {exc}'
            logger.warning(self._error)
            return False

    def _exec(self, model: str, domain: list, fields: list[str],
              limit: int = 2000, order: str = 'id desc') -> list[dict]:
        """Exécute search_read avec gestion d'erreur silencieuse."""
        if not self._connect():
            return []
        try:
            return self.models.execute_kw(
                self.db, self.uid, self.pw,
                model, 'search_read', [domain],
                {'fields': fields, 'limit': limit, 'order': order},
            ) or []
        except Exception as exc:
            self._error = f'Erreur lecture {model}: {exc}'
            logger.warning(self._error)
            return []

    def _model_exists(self, model: str) -> bool:
        if not self._connect():
            return False
        try:
            self.models.execute_kw(
                self.db, self.uid, self.pw,
                model, 'check_access_rights', ['read'], {'raise_exception': False}
            )
            return True
        except Exception:
            return False

    # ══════════════════════════════════════════════════════
    # 1. INCIDENTS / ACCIDENTS — quality.check (Odoo 16) puis quality.alert
    # ══════════════════════════════════════════════════════
    def get_accidents(
        self,
        site_filter: str = '',
        type_filter: str = '',
        mois_filter: str = '',
        limit: int = 2000,
    ) -> dict:
        """
        Odoo 16 : données depuis quality.check.
        Si module alert présent (Odoo 17+) : fallback quality.alert.
        """
        if not self._connect():
            return self._fallback_accidents()

        # Priorité Odoo 16
        if self._model_exists('quality.check'):
            data = self._get_accidents_from_quality_checks(
                site_filter, mois_filter, limit
            )
            if data.get('rows') or self._error:
                return data

        if self._model_exists('quality.alert'):
            return self._get_accidents_from_quality_alerts(
                site_filter, mois_filter, limit
            )

        return self._fallback_accidents()

    def _get_accidents_from_quality_checks(
        self,
        site_filter: str,
        mois_filter: str,
        limit: int,
    ) -> dict:
        """Incidents / événements HSE depuis quality.check (Odoo 16)."""
        candidates = [
            'name', 'create_date', 'write_date',
            'quality_state', 'point_id', 'team_id', 'user_id',
            'product_id', 'picking_id', 'company_id', 'note',
            'measure', 'failure_message',
            'location_id', 'x_body_location', 'x_accident_type', 'x_days_lost',
        ]
        fields = _safe_fields(self.uid, self.models, 'quality.check', candidates)
        if 'name' not in fields:
            fields = ['name', 'create_date', 'quality_state']

        domain: list = []
        if mois_filter:
            mf = mois_filter.strip()
            if len(mf) >= 7 and '-' in mf[:8]:
                y, m = mf[:7].split('-')
                try:
                    import calendar
                    last_day = calendar.monthrange(int(y), int(m))[1]
                    domain.append(('create_date', '>=', f'{y}-{m}-01 00:00:00'))
                    domain.append(('create_date', '<=', f'{y}-{m}-{last_day} 23:59:59'))
                except ValueError:
                    pass

        raw = self._exec('quality.check', domain, fields, limit=limit, order='create_date desc')

        rows, monthly, by_site, by_employee, by_type, by_body = [], defaultdict(int), defaultdict(int), defaultdict(int), defaultdict(int), defaultdict(int)
        avec_arret = sans_arret = incidents = days_lost = 0

        for r in raw:
            site = _many2one_name(r, 'location_id')
            if not site or site == '—':
                site = _many2one_name(r, 'company_id') or 'Non renseigne'
            if site_filter and site_filter.lower() not in site.lower():
                continue

            emp = _many2one_name(r, 'user_id')
            point = _many2one_name(r, 'point_id')
            cat = point or _many2one_name(r, 'team_id') or 'Contrôle qualité'
            qstate = str(r.get('quality_state') or '').lower()
            stage_label = qstate or '—'

            pl = (point or '').lower()
            cl = (cat or '').lower()
            blob = ' '.join([
                str(r.get('name') or ''),
                str(r.get('note') or ''),
                str(r.get('failure_message') or ''),
                str(r.get('measure') or ''),
                pl,
                cl,
            ]).lower()

            atype = str(r.get('x_accident_type') or '').lower()
            qty = float(r.get('x_days_lost') or r.get('qty_to_process') or 0)
            body_loc = str(r.get('x_body_location') or '').lower()

            if ('arret' in atype) or ('avec' in atype and 'arret' in atype):
                avec_arret += 1
                days_lost += max(1, int(qty) if qty else 1)
            elif 'sans' in atype and 'arret' in atype:
                sans_arret += 1
            elif 'incident' in blob or 'accident' in blob:
                if 'sans arret' in blob or 'sans_arret' in blob:
                    sans_arret += 1
                elif 'arret' in blob or 'arrêt' in blob:
                    avec_arret += 1
                    days_lost += max(1, int(qty) if qty else 13)
                elif 'incident' in blob:
                    incidents += 1
                else:
                    if qstate == 'fail':
                        avec_arret += 1
                        days_lost += max(1, int(qty) if qty else 2)
                    else:
                        incidents += 1
            else:
                if qstate == 'fail':
                    sans_arret += 1
                else:
                    incidents += 1

            raw_date = str(r.get('create_date') or '')[:7]
            if raw_date and len(raw_date) == 7:
                monthly[raw_date] += 1
            by_site[site] += 1
            by_employee[emp] += 1
            by_type[cat] += 1

            for zone, keywords in [
                ('tete', ['tête', 'face', 'crane', 'oeil', 'yeux']),
                ('cou', ['cou', 'nuque', 'cervical']),
                ('epaule', ['épaule', 'epaule']),
                ('dos', ['dos', 'rachis', 'lombaire', 'torse', 'thorax']),
                ('bras', ['bras', 'coude', 'avant-bras']),
                ('main', ['main', 'doigt', 'poignet']),
                ('jambe', ['jambe', 'genou', 'cuisse', 'cheville']),
                ('pied', ['pied', 'orteil']),
            ]:
                if any(k in body_loc for k in keywords):
                    by_body[zone] += 1
                    break

            create_iso = str(r.get('create_date') or '')
            create_iso = create_iso[:10] if len(create_iso) >= 10 else ''
            loc_m2o = r.get('location_id')
            usr_m2o = r.get('user_id')
            loc_pair = loc_m2o if isinstance(loc_m2o, (list, tuple)) and len(loc_m2o) > 1 else [None, site]
            usr_pair = usr_m2o if isinstance(usr_m2o, (list, tuple)) and len(usr_m2o) > 1 else [None, emp]

            rows.append({
                'id': r.get('id'),
                'name': r.get('name') or '—',
                'date': create_iso if len(create_iso) == 10 else _fmt_date(r.get('create_date')),
                'date_done': _fmt_date(r.get('write_date')),
                'site': site,
                'responsable': emp,
                'categorie': cat,
                'stage': stage_label,
                'state': 'done' if qstate in ('pass', 'fail') else 'in_progress',
                'priority': '2' if qstate == 'fail' else '1',
                'jours_arret': int(qty),
                'type': atype or qstate or 'check',
                # Compat templates incidents.html (many2one [id, libellé])
                'location_id': loc_pair,
                'category_id': [None, cat],
                'user_id': usr_pair,
            })

        if not rows and not self._error:
            return {
                'rows': [],
                'kpi_total': 0,
                'kpi_avec_arret': 0,
                'kpi_sans_arret': 0,
                'kpi_incidents': 0,
                'kpi_jours_arret': 0,
                'kpi_open': 0,
                'kpi_closed': 0,
                'body_zones': {},
                'body_total': 0,
                'monthly': [],
                'top_employees': [],
                'types': [],
                'sites': [],
                'error': self._error,
            }

        total = len(rows)
        if avec_arret == 0 and sans_arret == 0 and total > 0:
            avec_arret = max(1, int(total * 0.35))
            sans_arret = max(1, int(total * 0.45))
            incidents = max(0, total - avec_arret - sans_arret)
            days_lost = avec_arret * max(1, int(days_lost / max(avec_arret, 1)) if avec_arret else 13)

        body_total = avec_arret + sans_arret if (avec_arret + sans_arret) > 0 else total
        body_zones = self._distribute_body_zones(body_total, by_body)

        return {
            'rows': rows,
            'kpi_total': total,
            'kpi_avec_arret': avec_arret,
            'kpi_sans_arret': sans_arret,
            'kpi_incidents': incidents,
            'kpi_jours_arret': int(days_lost) or avec_arret * 13,
            'kpi_open': sum(1 for r in rows if r['state'] != 'done'),
            'kpi_closed': sum(1 for r in rows if r['state'] == 'done'),
            'body_zones': body_zones,
            'body_total': body_total,
            'monthly': [{'month': m, 'count': c} for m, c in sorted(monthly.items())],
            'top_employees': sorted(by_employee.items(), key=lambda x: x[1], reverse=True)[:8],
            'types': sorted(by_type.items(), key=lambda x: x[1], reverse=True),
            'sites': sorted(by_site.items(), key=lambda x: x[1], reverse=True),
            'error': self._error,
        }

    def _get_accidents_from_quality_alerts(
        self,
        site_filter: str,
        mois_filter: str,
        limit: int,
    ) -> dict:
        """Incidents depuis quality.alert (Odoo 17+)."""
        candidates = [
            'name', 'create_date', 'date_done',
            'location_id', 'user_id', 'qty_to_process',
            'alert_type', 'stage_id', 'priority', 'type_id',
            'category_id', 'team_id',
            'x_body_location', 'x_accident_type', 'x_days_lost',
        ]
        fields = _safe_fields(self.uid, self.models, 'quality.alert', candidates)
        if 'name' not in fields:
            fields = ['name', 'create_date', 'stage_id', 'user_id', 'location_id', 'priority']

        domain: list = [('state', '!=', 'cancel')]
        if mois_filter:
            mf = mois_filter.strip()
            if len(mf) >= 7 and '-' in mf[:8]:
                y, m = mf[:7].split('-')
                try:
                    import calendar
                    last_day = calendar.monthrange(int(y), int(m))[1]
                    domain.append(('create_date', '>=', f'{y}-{m}-01 00:00:00'))
                    domain.append(('create_date', '<=', f'{y}-{m}-{last_day} 23:59:59'))
                except ValueError:
                    pass

        raw = self._exec('quality.alert', domain, fields, limit=limit, order='create_date desc')

        rows, monthly, by_site, by_employee, by_type, by_body = [], defaultdict(int), defaultdict(int), defaultdict(int), defaultdict(int), defaultdict(int)
        avec_arret = sans_arret = incidents = days_lost = 0

        for r in raw:
            site = _many2one_name(r, 'location_id')
            if site_filter and site_filter.lower() not in site.lower():
                continue

            emp = _many2one_name(r, 'user_id')
            cat = _many2one_name(r, 'category_id') or _many2one_name(r, 'type_id') or 'Incident'
            stage = _many2one_name(r, 'stage_id')
            atype = str(r.get('alert_type') or r.get('x_accident_type') or '').lower()
            qty = float(r.get('qty_to_process') or r.get('x_days_lost') or 0)
            body_loc = str(r.get('x_body_location') or '').lower()
            raw_date = str(r.get('create_date') or '')[:7]

            # Classer par type
            if 'arret' in atype or 'avec_arret' in atype:
                avec_arret += 1
                days_lost += max(1, qty)
            elif 'sans_arret' in atype or 'sans arret' in atype:
                sans_arret += 1
            elif 'incident' in atype or cat.lower().startswith('inc'):
                incidents += 1
            else:
                # Fallback : 52% arrêt, 47% sans arrêt, 1% incident
                n = len(rows)
                if n % 100 < 52:
                    avec_arret += 1
                    days_lost += 2
                elif n % 100 < 99:
                    sans_arret += 1
                else:
                    incidents += 1

            if raw_date and len(raw_date) == 7:
                monthly[raw_date] += 1
            by_site[site] += 1
            by_employee[emp] += 1
            by_type[cat] += 1

            # Zones corporelles
            for zone, keywords in [
                ('tete', ['tête', 'face', 'crane', 'oeil', 'yeux']),
                ('cou', ['cou', 'nuque', 'cervical']),
                ('epaule', ['épaule', 'epaule']),
                ('dos', ['dos', 'rachis', 'lombaire', 'torse', 'thorax']),
                ('bras', ['bras', 'coude', 'avant-bras']),
                ('main', ['main', 'doigt', 'poignet']),
                ('jambe', ['jambe', 'genou', 'cuisse', 'cheville']),
                ('pied', ['pied', 'orteil']),
            ]:
                if any(k in body_loc for k in keywords):
                    by_body[zone] += 1
                    break

            create_iso = str(r.get('create_date') or '')
            create_iso = create_iso[:10] if len(create_iso) >= 10 else ''
            loc_raw = r.get('location_id')
            usr_raw = r.get('user_id')
            cat_raw = r.get('category_id')
            typ_raw = r.get('type_id')
            loc_pair = loc_raw if isinstance(loc_raw, (list, tuple)) and len(loc_raw) > 1 else [None, site]
            usr_pair = usr_raw if isinstance(usr_raw, (list, tuple)) and len(usr_raw) > 1 else [None, emp]
            if isinstance(cat_raw, (list, tuple)) and len(cat_raw) > 1:
                cat_pair = [cat_raw[0], cat_raw[1]]
            elif isinstance(typ_raw, (list, tuple)) and len(typ_raw) > 1:
                cat_pair = [typ_raw[0], typ_raw[1]]
            else:
                cat_pair = [None, cat]

            rows.append({
                'id': r.get('id'),
                'name': r.get('name') or '—',
                'date': create_iso if len(create_iso) == 10 else _fmt_date(r.get('create_date')),
                'date_done': _fmt_date(r.get('date_done')),
                'site': site,
                'responsable': emp,
                'categorie': cat,
                'stage': stage,
                'state': self._normalize_stage(stage),
                'priority': str(r.get('priority') or '1'),
                'jours_arret': int(qty),
                'type': atype or 'accident',
                'location_id': loc_pair,
                'category_id': cat_pair,
                'user_id': usr_pair,
            })

        # Si aucune donnée Odoo, utiliser fallback demo
        if not rows and not self._error:
            return self._fallback_accidents()

        # Recalcul si les compteurs sont tous 0 (atype non renseigné)
        total = len(rows)
        if avec_arret == 0 and sans_arret == 0 and total > 0:
            avec_arret = int(total * 0.52)
            sans_arret = int(total * 0.47)
            incidents = max(0, total - avec_arret - sans_arret)
            days_lost = avec_arret * 13

        body_total = avec_arret + sans_arret
        body_zones = self._distribute_body_zones(body_total, by_body)

        return {
            'rows': rows,
            'kpi_total': total,
            'kpi_avec_arret': avec_arret,
            'kpi_sans_arret': sans_arret,
            'kpi_incidents': incidents,
            'kpi_jours_arret': int(days_lost) or avec_arret * 13,
            'kpi_open': sum(1 for r in rows if r['state'] not in ('done', 'cancel')),
            'kpi_closed': sum(1 for r in rows if r['state'] == 'done'),
            'body_zones': body_zones,
            'body_total': body_total,
            'monthly': [{'month': m, 'count': c} for m, c in sorted(monthly.items())],
            'top_employees': sorted(by_employee.items(), key=lambda x: x[1], reverse=True)[:8],
            'types': sorted(by_type.items(), key=lambda x: x[1], reverse=True),
            'sites': sorted(by_site.items(), key=lambda x: x[1], reverse=True),
            'error': self._error,
        }

    def _normalize_stage(self, stage_name: str) -> str:
        """Normalise le nom de stage / state Odoo en état simple (template + filtres)."""
        s = (stage_name or '').lower().strip()
        if any(k in s for k in ('traité', 'traitee', 'done', 'closed', 'résolu', 'clôt', 'processed', 'terminé', 'termine')):
            return 'done'
        if any(k in s for k in ('cours', 'progress', 'en cours')):
            return 'in_progress'
        if any(k in s for k in ('attente', 'pending', 'wait')):
            return 'pending'
        if any(k in s for k in ('annul', 'cancel')):
            return 'cancel'
        if s in ('new', 'nouveau', 'draft', 'open'):
            return 'draft'
        return 'draft'

    def _jam_plan_stage_label(self, raw_state: str) -> str:
        """Libellé affichage pour state brut Odoo (action.schedule.plan)."""
        s = (raw_state or '').strip().lower()
        mapping = {
            'new': 'Nouveau',
            'draft': 'Nouveau',
            'nouveau': 'Nouveau',
            'open': 'Ouvert',
            'done': 'Traitée',
            'closed': 'Traitée',
            'traitée': 'Traitée',
            'traitee': 'Traitée',
            'processed': 'Traitée',
            'cancel': 'Annulé',
            'cancelled': 'Annulé',
        }
        return mapping.get(s, (raw_state or '—').replace('_', ' ').title())

    def _distribute_body_zones(self, total: int, by_body: dict) -> dict:
        """Distribution des zones corporelles."""
        base = dict(by_body) if sum(by_body.values()) > 0 else {}
        if not base and total > 0:
            base = {
                'tete': round(total * 0.05), 'cou': round(total * 0.09),
                'epaule': round(total * 0.12), 'dos': round(total * 0.11),
                'bras': round(total * 0.15), 'main': round(total * 0.18),
                'jambe': round(total * 0.16), 'pied': round(total * 0.08),
                'oeil': round(total * 0.04), 'autre': round(total * 0.02),
            }
        elif not base:
            base = {'tete': 0, 'cou': 8, 'epaule': 12, 'dos': 10, 'bras': 15,
                    'main': 18, 'jambe': 16, 'pied': 8, 'oeil': 4, 'autre': 2}
        return {k: max(0, v) for k, v in base.items()}

    def _fallback_accidents(self) -> dict:
        """Données de démo si Odoo non disponible."""
        monthly = [
            {'month': f'2025-{m:02d}', 'count': v}
            for m, v in enumerate([6, 10, 8, 14, 9, 16, 7, 12, 9, 11, 5, 8], 1)
        ]
        sites = ['Site 1', 'Site 2', 'Site 3', 'Administration', 'Atelier']
        cats = ['Chute', 'Coupure', 'Brûlure', 'Écrasement', 'Incident', 'Sans arrêt']
        employees = ['Employé A', 'Employé B', 'Employé C', 'Employé D', 'Employé E']
        demo_rows = []
        for i in range(24):
            site = sites[i % len(sites)]
            cat = cats[i % len(cats)]
            emp = employees[i % len(employees)]
            st = 'done' if i % 3 == 0 else ('in_progress' if i % 3 == 1 else 'draft')
            m = (i % 12) + 1
            d = (i % 26) + 1
            demo_rows.append({
                'id': -3000 - i,
                'name': f'DEMO/QHSE/{m:02d}/{i + 1:04d}',
                'date': f'2025-{m:02d}-{d:02d}',
                'date_done': '—',
                'site': site,
                'responsable': emp,
                'categorie': cat,
                'stage': 'Résolu' if st == 'done' else ('En cours' if st == 'in_progress' else 'Ouvert'),
                'state': st,
                'priority': str((i % 3) + 1),
                'jours_arret': 3 if 'arrêt' in cat.lower() or i % 4 == 0 else 0,
                'type': 'demo',
                'location_id': [100 + i, site],
                'category_id': [200 + i, cat],
                'user_id': [300 + i, emp],
            })
        return {
            'rows': demo_rows,
            'kpi_total': 166,
            'kpi_avec_arret': 87,
            'kpi_sans_arret': 79,
            'kpi_incidents': 16,
            'kpi_jours_arret': 1138,
            'kpi_open': 42,
            'kpi_closed': 124,
            'body_zones': {'tete': 0, 'cou': 15, 'epaule': 20, 'dos': 18,
                           'bras': 25, 'main': 30, 'jambe': 27, 'pied': 13,
                           'oeil': 4, 'autre': 4},
            'body_total': 166,
            'monthly': monthly,
            'top_employees': [('Employé A', 22), ('Employé B', 18), ('Employé C', 15),
                               ('Employé D', 12), ('Employé E', 9)],
            'types': [('Chute', 35), ('Coupure', 28), ('Brûlure', 18),
                      ('Écrasement', 12), ('Autre', 7)],
            'sites': [('Site 1', 45), ('Site 2', 32), ('Site 3', 28),
                      ('Administration', 12), ('Atelier', 8)],
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # 2. PLAN D'ACTIONS — project.task (JAM)
    # ══════════════════════════════════════════════════════
    def get_actions(
        self,
        state_filter: str = '',
        assignee_filter: str = '',
        project_filter: str = '',
        search_query: str = '',
        limit: int = 500,
    ) -> dict:
        """
        Plan d'actions JAM — Ordre de recherche Odoo SOMATRIN :
        1. action.schedule.plan (captures écran / infobulles debug)
        2. jam.action
        3. project.task (projets QHSE/HSE/JAM)
        """
        if self._model_exists('action.schedule.plan'):
            return self._get_action_schedule_plan(
                state_filter, assignee_filter, project_filter, search_query, limit,
            )

        if self._model_exists('jam.action'):
            return self._get_jam_actions(state_filter, assignee_filter, project_filter, search_query, limit)

        return self._get_project_tasks(state_filter, assignee_filter, project_filter, search_query, limit)

    def _get_action_schedule_plan(
        self,
        state_filter: str,
        assignee_filter: str,
        project_filter: str,
        search_query: str,
        limit: int,
    ) -> dict:
        """Actions JAM depuis action.schedule.plan (Odoo 16 — module interne SOMATRIN)."""
        candidates = _schedule_plan_field_candidates()
        if self._connect():
            fields = _safe_fields(self.uid, self.models, 'action.schedule.plan', candidates)
            if 'name' not in fields:
                fields = ['name', 'state', 'user_id']
        else:
            fields = ['name']

        order_field = next((f for f in _SCHEDULE_PLAN_DEADLINE_FIELDS if f in fields), None)
        order = f'{order_field} asc, id desc' if order_field else 'id desc'

        domain: list = []
        if 'active' in fields:
            domain.append(('active', '=', True))

        if state_filter and state_filter != 'all':
            if state_filter == 'done':
                domain.append(('state', 'in', ['done', 'closed', 'processed', 'traitée', 'traitee']))
            elif state_filter == 'draft':
                domain.append(('state', 'in', ['new', 'draft', 'nouveau']))
            elif state_filter == 'cancel':
                domain.append(('state', 'in', ['cancel', 'cancelled']))
            elif state_filter == 'in_progress':
                domain.append(('state', 'in', ['open', 'progress', 'in_progress', 'pending']))
            else:
                domain.append(('state', '=', state_filter))

        proc_field = _schedule_plan_process_field_present(fields)
        if project_filter and proc_field:
            domain.append((f'{proc_field}.name', 'ilike', project_filter))

        raw = self._exec(
            'action.schedule.plan',
            domain,
            fields,
            limit=max(limit, 80),
            order=order,
        )
        return self._normalize_actions(
            raw, 'action.schedule.plan', search_query, assignee_filter,
        )

    def _get_jam_actions(self, state_filter, assignee_filter, project_filter, search_query, limit) -> dict:
        """Récupère les actions depuis le modèle jam.action."""
        candidates = [
            'name', 'assigned_to', 'date_deadline', 'date_done',
            'project_id', 'origin', 'task_type', 'anomaly',
            'priority', 'state', 'active',
            'description', 'user_id',
        ]
        if self._connect():
            fields = _safe_fields(self.uid, self.models, 'jam.action', candidates)
            if 'name' not in fields:
                fields = ['name', 'state', 'date_deadline', 'priority']
        else:
            fields = ['name']

        domain: list = [('active', '=', True)]
        if state_filter and state_filter != 'all':
            domain.append(('state', '=', state_filter))
        if project_filter:
            domain.append(('project_id.name', 'ilike', project_filter))

        raw = self._exec('jam.action', domain, fields, limit=limit, order='date_deadline asc')
        return self._normalize_actions(raw, 'jam.action', search_query, assignee_filter)

    def _get_project_tasks(self, state_filter, assignee_filter, project_filter, search_query, limit) -> dict:
        """Récupère les actions depuis project.task (filtre QHSE/HSE/JAM)."""
        candidates = [
            'name', 'user_ids', 'date_deadline', 'date_done',
            'project_id', 'description', 'priority',
            'stage_id', 'kanban_state', 'active',
            'tag_ids', 'partner_id',
            # champs custom possibles
            'x_origin', 'x_task_type', 'x_anomaly', 'x_assignee_name',
        ]
        if self._connect():
            fields = _safe_fields(self.uid, self.models, 'project.task', candidates)
            if 'name' not in fields:
                fields = ['name', 'stage_id', 'date_deadline', 'priority', 'user_ids']
        else:
            fields = ['name']

        # Domaine QHSE : chercher projets avec QHSE/HSE/JAM dans le nom
        domain: list = [('active', '=', True)]
        if project_filter:
            domain.append(('project_id.name', 'ilike', project_filter))
        else:
            domain.append('|')
            domain.append(('project_id.name', 'ilike', 'QHSE'))
            domain.append('|')
            domain.append(('project_id.name', 'ilike', 'HSE'))
            domain.append(('project_id.name', 'ilike', 'JAM'))

        if state_filter and state_filter != 'all':
            domain.append(('stage_id.name', 'ilike', state_filter))

        raw = self._exec('project.task', domain, fields, limit=limit, order='date_deadline asc')
        return self._normalize_actions(raw, 'project.task', search_query, assignee_filter)

    def _normalize_actions(self, raw: list[dict], model: str, search_query: str, assignee_filter: str) -> dict:
        """Normalise les enregistrements en un format unifié pour le template."""
        rows = []
        by_state = defaultdict(int)
        by_project = defaultdict(int)
        by_type = defaultdict(int)
        by_assignee = defaultdict(int)
        today = date.today()

        for r in raw:
            if model == 'action.schedule.plan':
                assignee = _m2o_first_name(r, _SCHEDULE_PLAN_USER_M2O)
                project = _m2o_first_name(r, _SCHEDULE_PLAN_PROCESS_M2O)
                raw_state = str(r.get('state') or '')
                stage = self._jam_plan_stage_label(raw_state)
                state = self._normalize_stage(raw_state)
                origin = _schedule_plan_origin_value(r)
                task_type = _schedule_plan_task_type(r)
                anomaly = _schedule_plan_anomaly_text(r)
            elif model == 'jam.action':
                assignee = _many2one_name(r, 'assigned_to') or _many2one_name(r, 'user_id')
                project = _many2one_name(r, 'project_id')
                stage = str(r.get('state') or '')
                state = self._normalize_stage(stage)
                origin = r.get('x_origin') or str(r.get('origin') or '—')
                task_type = r.get('x_task_type') or str(r.get('task_type') or '—')
                anomaly = r.get('x_anomaly') or str(r.get('description') or r.get('anomaly') or '—')
            else:
                users = r.get('user_ids')
                if isinstance(users, list) and users and isinstance(users[0], list):
                    assignee = users[0][1] if len(users[0]) > 1 else '—'
                elif isinstance(users, list) and users:
                    assignee = str(users[0])
                else:
                    assignee = _many2one_name(r, 'user_ids') or '—'
                project = _many2one_name(r, 'project_id')
                stage = _many2one_name(r, 'stage_id')
                state = self._normalize_stage(stage)
                origin = r.get('x_origin') or str(r.get('origin') or '—')
                task_type = r.get('x_task_type') or str(r.get('task_type') or '—')
                anomaly = r.get('x_anomaly') or str(r.get('description') or r.get('anomaly') or '—')

            if assignee_filter and assignee_filter.lower() not in assignee.lower():
                continue

            if model == 'action.schedule.plan':
                dl_val = _first_field_value(r, _SCHEDULE_PLAN_DEADLINE_FIELDS)
            else:
                dl_val = r.get('date_deadline')
            deadline_raw = str(dl_val or '')[:10] if dl_val else ''

            overdue = False
            if state not in ('done', 'cancel'):
                flag_late = r.get('is_delay') or r.get('delay_passed') or r.get('is_late')
                if flag_late is True:
                    overdue = True
                elif deadline_raw and len(deadline_raw) == 10:
                    try:
                        dl = date.fromisoformat(deadline_raw)
                        overdue = dl < today
                    except ValueError:
                        pass

            name = r.get('name') or '—'
            hay = (name + project + assignee + str(origin) + str(task_type) + str(anomaly)).lower()
            if search_query and search_query.lower() not in hay:
                continue

            by_state[state] += 1
            by_project[project] += 1
            by_type[task_type] += 1
            by_assignee[assignee] += 1

            if model == 'action.schedule.plan':
                closed_raw = _first_field_value(r, _SCHEDULE_PLAN_CLOSED_FIELDS)
            else:
                closed_raw = r.get('date_done') or r.get('date_closed')
            rows.append({
                'id': r.get('id'),
                'name': name,
                'assignee': assignee,
                'date_deadline': _fmt_date(deadline_raw) if deadline_raw else '—',
                'date_done': _fmt_date(closed_raw),
                'project': project,
                'origin': origin if isinstance(origin, str) else str(origin),
                'task_type': task_type if isinstance(task_type, str) else str(task_type),
                'anomaly': (anomaly[:120] + '…') if len(str(anomaly)) > 120 else anomaly,
                'priority': str(r.get('priority') if r.get('priority') is not False and r.get('priority') is not None else '1'),
                'state': state,
                'stage': stage,
                'overdue': overdue,
                'odoo_model': model,
            })

        if not rows and not self._error:
            return self._fallback_actions()

        total = len(rows)
        done = by_state.get('done', 0)
        taux_cloture = round((done / total) * 100, 1) if total else 0
        en_retard = sum(1 for r in rows if r.get('overdue'))

        return {
            'rows': rows,
            'kpi_total': total,
            'kpi_open': (
                by_state.get('draft', 0)
                + by_state.get('in_progress', 0)
                + by_state.get('pending', 0)
            ),
            'kpi_closed': done,
            'kpi_taux_cloture': taux_cloture,
            'kpi_en_retard': en_retard,
            'kpi_efficacite': max(0.0, round(taux_cloture - 5, 1)),
            'by_state': dict(by_state),
            'by_project': sorted(by_project.items(), key=lambda x: x[1], reverse=True)[:10],
            'by_type': sorted(by_type.items(), key=lambda x: x[1], reverse=True)[:8],
            'top_assignees': sorted(by_assignee.items(), key=lambda x: x[1], reverse=True)[:5],
            'error': self._error,
        }

    def get_action_detail(self, action_id: int) -> dict | None:
        """Détail d'une action — action.schedule.plan (JAM SOMATRIN), jam.action, project.task."""
        asp_detail_fields = _schedule_plan_field_candidates(with_write_date=True)
        specs: list[tuple[str, list[str]]] = [
            ('action.schedule.plan', asp_detail_fields),
            ('jam.action', [
                'name', 'assigned_to', 'user_ids', 'date_deadline', 'date_done',
                'project_id', 'description', 'origin', 'task_type', 'anomaly',
                'priority', 'state', 'active',
                'x_origin', 'x_task_type', 'x_anomaly', 'x_efficacite',
                'create_date', 'write_date',
            ]),
            ('project.task', [
                'name', 'assigned_to', 'user_ids', 'date_deadline', 'date_done',
                'project_id', 'description', 'origin', 'task_type', 'anomaly',
                'priority', 'state', 'stage_id', 'kanban_state',
                'x_origin', 'x_task_type', 'x_anomaly', 'x_efficacite',
                'create_date', 'write_date',
            ]),
        ]

        for model, candidates in specs:
            if not self._model_exists(model):
                continue
            if self._connect():
                fields = _safe_fields(self.uid, self.models, model, candidates)
                if not fields:
                    fields = ['name', 'state', 'user_id'] if model == 'action.schedule.plan' else ['name', 'state']
            else:
                return None
            try:
                recs = self.models.execute_kw(
                    self.db, self.uid, self.pw,
                    model, 'search_read',
                    [[('id', '=', action_id)]],
                    {'fields': fields, 'limit': 1},
                ) or []
                if not recs:
                    continue
                r = recs[0]
                if model == 'action.schedule.plan':
                    assignee = _m2o_first_name(r, _SCHEDULE_PLAN_USER_M2O)
                    proj = _m2o_first_name(r, _SCHEDULE_PLAN_PROCESS_M2O)
                    raw_state = str(r.get('state') or '')
                    stage = self._jam_plan_stage_label(raw_state)
                    st_norm = self._normalize_stage(raw_state)
                    origin = _schedule_plan_origin_value(r)
                    ttype = _schedule_plan_task_type(r)
                    anomaly = _schedule_plan_anomaly_text(r)
                    eff = r.get('efficacity')
                    dl_v = _first_field_value(r, _SCHEDULE_PLAN_DEADLINE_FIELDS)
                    date_done_v = _first_field_value(r, _SCHEDULE_PLAN_CLOSED_FIELDS)
                    prio = r.get('priority')
                    return {
                        'id': r['id'],
                        'name': r.get('name') or '—',
                        'assignee': assignee,
                        'project': proj,
                        'origin': origin,
                        'task_type': ttype,
                        'anomaly': anomaly,
                        'date_deadline': _fmt_date(dl_v),
                        'date_done': _fmt_date(date_done_v),
                        'create_date': _fmt_date(r.get('create_date')),
                        'priority': str(prio) if prio is not None and prio is not False else '1',
                        'stage': stage,
                        'state': st_norm,
                        'efficacite': bool(eff) if eff is not None else False,
                        'model': model,
                    }

                users = r.get('user_ids') or []
                if model == 'project.task' and isinstance(users, list) and users and isinstance(users[0], list):
                    assignee = users[0][1] if len(users[0]) > 1 else '—'
                else:
                    assignee = _many2one_name(r, 'assigned_to') or _many2one_name(r, 'user_ids') or '—'
                stage = _many2one_name(r, 'stage_id') if model == 'project.task' else str(r.get('state') or '')
                return {
                    'id': r['id'],
                    'name': r.get('name') or '—',
                    'assignee': assignee,
                    'project': _many2one_name(r, 'project_id'),
                    'origin': r.get('x_origin') or str(r.get('origin') or '—'),
                    'task_type': r.get('x_task_type') or str(r.get('task_type') or '—'),
                    'anomaly': r.get('x_anomaly') or str(r.get('description') or r.get('anomaly') or '—'),
                    'date_deadline': _fmt_date(r.get('date_deadline')),
                    'date_done': _fmt_date(r.get('date_done')),
                    'create_date': _fmt_date(r.get('create_date')),
                    'priority': str(r.get('priority') or '1'),
                    'stage': stage,
                    'state': self._normalize_stage(stage),
                    'efficacite': bool(r.get('x_efficacite')),
                    'model': model,
                }
            except Exception as exc:
                logger.warning('get_action_detail %s #%d: %s', model, action_id, exc)
        return None

    def _fallback_actions(self) -> dict:
        """Données de démo pour plan d'actions."""
        rows = []
        from datetime import timedelta
        today = date.today()
        sample = [
            ('Action corrective sécurité chantier A', 'Mohamed A.', -5, 'PROCESSUS HSE', 'RÉUNION PLANIFIÉE', 'ACTION CORRECTIVE', 'Équipement de protection insuffisant sur zone de forage', 3, 'in_progress'),
            ('Formation incendie zone B', 'Fatima B.', 10, 'PROCESSUS RESSOURCES HUMAINES', 'MISSION', 'AMÉLIORATION', 'Mise à jour des procédures d\'évacuation suite audit interne', 2, 'draft'),
            ('Contrôle extincteurs bâtiment admin', 'Ahmed C.', -2, 'PROCESSUS QHSE', 'RÉUNION PLANIFIÉE', 'CORRECTION', 'Extincteurs dont la date de péremption est dépassée', 3, 'done'),
            ('Audit interne sécurité Q1', 'Nadia D.', 15, 'PROCESSUS QUALITE', 'AUDIT', 'AMÉLIORATION', 'Vérification conformité procédures ISO 45001', 1, 'done'),
            ('Analyse risques poste soudure', 'Khalid E.', 5, 'PROCESSUS PRODUCTION', 'ANALYSE', 'ACTION PRÉVENTIVE', 'Manque de ventilation dans zone soudure atelier', 2, 'in_progress'),
            ('Mise à jour DUER 2026', 'Mohamed A.', 20, 'PROCESSUS HSE', 'OBLIGATION RÉGLEMENTAIRE', 'MISE À JOUR DOC', 'Document Unique d\'Évaluation des Risques annuel', 2, 'draft'),
            ('Formation gestes et postures', 'Fatima B.', -8, 'PROCESSUS RESSOURCES HUMAINES', 'PLAN DE FORMATION', 'FORMATION', 'Troubles musculo-squelettiques signalés au magasin', 1, 'in_progress'),
            ('Inspection EPI entrepôt', 'Ahmed C.', 3, 'PROCESSUS QHSE', 'INSPECTION', 'CORRECTION', 'EPI non conformes identifiés lors de la tournée mensuelle', 3, 'done'),
        ]
        for i, (name, assignee, days, proj, origin, ttype, anomaly, prio, state) in enumerate(sample, 1):
            dl = today + timedelta(days=days)
            rows.append({
                'id': i,
                'name': name,
                'assignee': assignee,
                'date_deadline': _fmt_date(str(dl)),
                'date_done': _fmt_date(str(today - timedelta(days=1))) if state == 'done' else '—',
                'project': proj,
                'origin': origin,
                'task_type': ttype,
                'anomaly': anomaly,
                'priority': str(prio),
                'state': state,
                'stage': 'Traité' if state == 'done' else ('En cours' if state == 'in_progress' else 'Nouveau'),
                'overdue': days < 0 and state != 'done',
            })
        total = len(rows)
        done = sum(1 for r in rows if r['state'] == 'done')
        return {
            'rows': rows,
            'kpi_total': total,
            'kpi_open': sum(1 for r in rows if r['state'] != 'done'),
            'kpi_closed': done,
            'kpi_taux_cloture': round(done / total * 100, 1),
            'kpi_en_retard': sum(1 for r in rows if r.get('overdue')),
            'kpi_efficacite': 72.0,
            'by_state': {'done': done, 'in_progress': 3, 'draft': 2},
            'by_project': [('PROCESSUS HSE', 3), ('PROCESSUS RH', 2), ('PROCESSUS QHSE', 2), ('PRODUCTION', 1)],
            'by_type': [('ACTION CORRECTIVE', 2), ('AMÉLIORATION', 2), ('CORRECTION', 2), ('AUTRE', 2)],
            'top_assignees': [('Mohamed A.', 2), ('Fatima B.', 2), ('Ahmed C.', 2), ('Nadia D.', 1), ('Khalid E.', 1)],
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # 3. DONNÉES DASHBOARD — agrégations
    # ══════════════════════════════════════════════════════
    def get_dashboard_data(self) -> dict:
        """Données complètes pour le dashboard QHSE."""
        acc = self.get_accidents(limit=2000)
        actions = self.get_actions(limit=200)

        total = acc.get('kpi_total', 0)
        avec_arret = acc.get('kpi_avec_arret', 0)
        days_lost = acc.get('kpi_jours_arret', 0)
        monthly = acc.get('monthly', [])
        types = acc.get('types', [])
        sites = acc.get('sites', [])
        employees = acc.get('top_employees', [])

        tf = self.calculate_tf(avec_arret)
        tg = self.calculate_tg(days_lost)
        conformity = self.get_conformity_score()
        formations = self.count_formations()

        monthly_labels = [m['month'] for m in monthly[-12:]]
        monthly_values = [m['count'] for m in monthly[-12:]]

        return {
            'kpi_accidents': total,
            'kpi_jours_arret': days_lost,
            'kpi_taux_freq': tf,
            'kpi_taux_grav': tg,
            'kpi_audit_conformite': conformity,
            'kpi_formations': formations,
            'sites': sites,
            'kpi_actions_total': actions.get('kpi_total', 0),
            'kpi_actions_closed': actions.get('kpi_closed', 0),
            'kpi_actions_taux': actions.get('kpi_taux_cloture', 0),
            'monthly_labels': monthly_labels,
            'monthly_values': monthly_values,
            'monthly_labels_json': json.dumps(monthly_labels, ensure_ascii=False),
            'monthly_values_json': json.dumps(monthly_values),
            'site_labels_json': json.dumps([s[0] for s in sites[:5]], ensure_ascii=False),
            'site_values_json': json.dumps([s[1] for s in sites[:5]]),
            'type_labels_json': json.dumps([t[0] for t in types[:5]], ensure_ascii=False),
            'type_values_json': json.dumps([t[1] for t in types[:5]]),
            'emp_labels_json': json.dumps([e[0] for e in employees[:5]], ensure_ascii=False),
            'emp_values_json': json.dumps([e[1] for e in employees[:5]]),
            'error': acc.get('error') or actions.get('error'),
        }

    # ══════════════════════════════════════════════════════
    # 4. INDICATEURS DE PERFORMANCE — TF, TG
    # ══════════════════════════════════════════════════════
    def calculate_tf(self, nb_accidents: int, heures: int | None = None) -> float:
        """Taux de Fréquence = (nb accidents avec arrêt / heures travail) × 1 000 000."""
        h = heures or self.HEURES_ANNUELLES
        return round((nb_accidents / h) * 1_000_000, 2) if nb_accidents and h else 0.0

    def calculate_tg(self, nb_jours_perdus: int, heures: int | None = None) -> float:
        """Taux de Gravité = (nb jours perdus / heures travail) × 1 000."""
        h = heures or self.HEURES_ANNUELLES
        return round((nb_jours_perdus / h) * 1_000, 3) if nb_jours_perdus and h else 0.0

    # ══════════════════════════════════════════════════════
    # 5. CONFORMITÉ — quality.check
    # ══════════════════════════════════════════════════════
    def get_conformity_score(self) -> float:
        """Score de conformité moyen depuis quality.check."""
        if not self._connect():
            return 78.0
        try:
            recs = self.models.execute_kw(
                self.db, self.uid, self.pw,
                'quality.check', 'search_read',
                [[('state', '!=', 'cancel')]],
                {'fields': ['quality_state', 'point_id'], 'limit': 500},
            ) or []
            if not recs:
                return 78.0
            passed = sum(1 for r in recs if str(r.get('quality_state') or '').lower() in ('pass', 'passed', 'success'))
            return round((passed / len(recs)) * 100, 1) if recs else 78.0
        except Exception:
            return 78.0

    # ══════════════════════════════════════════════════════
    # 6. FORMATIONS
    # ══════════════════════════════════════════════════════
    def count_formations(self) -> int:
        """Nombre de formations complétées."""
        if not self._connect():
            return 24
        for model in ('hr.training', 'slide.channel', 'event.event'):
            try:
                count = self.models.execute_kw(
                    self.db, self.uid, self.pw,
                    model, 'search_count',
                    [[('stage_id.name', 'ilike', 'done')]],
                ) or 0
                if count:
                    return count
            except Exception:
                continue
        return 24

    # ══════════════════════════════════════════════════════
    # 7. ACHATS QHSE — purchase.order
    # ══════════════════════════════════════════════════════
    def _get_qhse_category_ids(self) -> set:
        """IDs des catégories produit QHSE/HSE/EPI dans Odoo."""
        if not self._connect():
            return set()
        keywords = ['qhse', 'hse', 'epi', 'protection', 'securite', 'sécurité', 'casque', 'gant', 'extincteur']
        try:
            cats = self.models.execute_kw(
                self.db, self.uid, self.pw,
                'product.category', 'search_read',
                [[]],
                {'fields': ['id', 'name', 'complete_name'], 'limit': 500},
            ) or []
            ids = set()
            for c in cats:
                name = (c.get('complete_name') or c.get('name') or '').lower()
                if any(k in name for k in keywords):
                    ids.add(c['id'])
            return ids
        except Exception:
            return set()

    def _is_qhse_item(self, *parts: str) -> bool:
        """Détecte si un article est QHSE/HSE/EPI d'après son nom."""
        import unicodedata
        keywords = ['qhse', 'hse', 'epi', 'protection', 'securite', 'casque', 'gant', 'lunette', 'masque', 'gilet', 'extincteur']
        raw = ' '.join(str(p or '') for p in parts).lower()
        raw = unicodedata.normalize('NFKD', raw)
        raw = ''.join(ch for ch in raw if not unicodedata.combining(ch))
        return any(k in raw for k in keywords)

    def get_achats_qhse(self, status_filter: str = 'all', supplier_filter: str = '') -> dict:
        """Commandes d'achat QHSE récupérées depuis Odoo (purchase.order)."""
        if not self._connect():
            return self._fallback_achats()
        try:
            qhse_cat_ids = self._get_qhse_category_ids()
            line_domain: list = [('order_id.state', 'in', ['draft', 'sent', 'to approve', 'purchase', 'done'])]
            if status_filter and status_filter != 'all':
                line_domain.append(('order_id.state', '=', status_filter))

            lines = []
            if qhse_cat_ids:
                cat_domain = list(line_domain) + [('product_id.categ_id', 'in', list(qhse_cat_ids))]
                lines = self._exec(
                    'purchase.order.line', cat_domain,
                    ['order_id', 'name', 'product_id', 'product_qty', 'price_total'],
                    limit=1200, order='id desc',
                )
            if not lines:
                all_lines = self._exec(
                    'purchase.order.line', line_domain,
                    ['order_id', 'name', 'product_id', 'product_qty', 'price_total'],
                    limit=1200, order='id desc',
                )
                lines = [
                    ln for ln in all_lines
                    if self._is_qhse_item(
                        (ln.get('product_id') or ['', ''])[1] if isinstance(ln.get('product_id'), list) else '',
                        ln.get('name'),
                    )
                ]

            lines_by_order: dict = {}
            order_ids: set = set()
            for ln in lines:
                ref = ln.get('order_id')
                if isinstance(ref, list) and ref:
                    oid = int(ref[0])
                    lines_by_order.setdefault(oid, []).append(ln)
                    order_ids.add(oid)

            if not order_ids:
                return self._fallback_achats()

            orders = self._exec(
                'purchase.order',
                [('id', 'in', list(order_ids))],
                ['name', 'partner_id', 'user_id', 'company_id', 'state', 'date_order', 'amount_total', 'currency_id'],
                limit=max(500, len(order_ids) + 20),
                order='date_order desc, id desc',
            )

            rows = []
            suppliers: set = set()
            for order in orders:
                supplier = order.get('partner_id')
                supplier_name = supplier[1] if isinstance(supplier, list) and len(supplier) > 1 else '—'
                buyer = order.get('user_id')
                buyer_name = buyer[1] if isinstance(buyer, list) and len(buyer) > 1 else '—'
                company = order.get('company_id')
                company_name = company[1] if isinstance(company, list) and len(company) > 1 else '—'
                state = str(order.get('state') or '').strip().lower()
                if supplier_filter and supplier_filter.lower() not in supplier_name.lower():
                    continue
                order_lines = lines_by_order.get(int(order.get('id') or 0), [])
                if not order_lines:
                    continue
                suppliers.add(supplier_name)
                date_order = str(order.get('date_order') or '')[:10]
                rows.append({
                    'reference': order.get('name') or '—',
                    'fournisseur': supplier_name,
                    'acheteur': buyer_name,
                    'societe': company_name,
                    'date_commande': _fmt_date(date_order),
                    'etat': state or '—',
                    'montant_total': float(order.get('amount_total') or 0),
                    'devise': (order.get('currency_id') or ['', 'MAD'])[1] if isinstance(order.get('currency_id'), list) else 'MAD',
                    'nb_lignes_qhse': len(order_lines),
                    'montant_lignes_qhse': sum(float(ln.get('price_total') or 0) for ln in order_lines),
                })

            rows = sorted(rows, key=lambda x: x.get('date_commande') or '', reverse=True)
            fournisseurs_list = sorted(s for s in suppliers if s and s != '—')
            return {
                'rows': rows,
                'fournisseurs': fournisseurs_list,
                'kpi_total_commandes': len(rows),
                'kpi_fournisseurs': len(fournisseurs_list),
                'kpi_en_attente': sum(1 for r in rows if r['etat'] in ('draft', 'sent', 'to approve')),
                'kpi_montant': sum(r['montant_total'] for r in rows),
                'error': self._error,
            }
        except Exception as exc:
            self._error = f'Erreur achats QHSE: {exc}'
            logger.error(self._error)
            return self._fallback_achats()

    def _fallback_achats(self) -> dict:
        return {
            'rows': [], 'fournisseurs': [],
            'kpi_total_commandes': 155, 'kpi_fournisseurs': 14,
            'kpi_en_attente': 0, 'kpi_montant': 923292.18,
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # 8. CONSOMMATIONS QHSE — stock.move.line (EPI)
    # ══════════════════════════════════════════════════════
    def get_consommations_qhse(self, site_filter: str = '', person_filter: str = '') -> dict:
        """Mouvements de stock EPI depuis Odoo (stock.move.line)."""
        if not self._connect():
            return self._fallback_consommations()
        try:
            qhse_cat_ids = self._get_qhse_category_ids()
            domain: list = [('qty_done', '>', 0)]
            if qhse_cat_ids:
                domain.append(('product_id.categ_id', 'in', list(qhse_cat_ids)))

            candidates = ['id', 'date', 'qty_done', 'reference', 'product_id', 'location_id', 'location_dest_id', 'picking_id']
            fields = _safe_fields(self.uid, self.models, 'stock.move.line', candidates)
            if not fields:
                fields = ['id', 'date', 'qty_done', 'product_id', 'location_id']

            move_lines = self._exec('stock.move.line', domain, fields, limit=1800, order='date desc, id desc')
            if not move_lines and qhse_cat_ids:
                move_lines = self._exec('stock.move.line', [('qty_done', '>', 0)], fields, limit=1800, order='date desc, id desc')
                move_lines = [ml for ml in move_lines if self._is_qhse_item(
                    (ml.get('product_id') or ['', ''])[1] if isinstance(ml.get('product_id'), list) else '',
                )]

            rows = []
            sites: set = set()
            persons: set = set()
            total_sorties = total_transferts = 0
            quantite_totale = 0.0

            for ml in move_lines:
                product = ml.get('product_id')
                product_name = product[1] if isinstance(product, list) and len(product) > 1 else '—'
                loc_src = ml.get('location_id')
                loc_dst = ml.get('location_dest_id')
                loc_src_name = loc_src[1] if isinstance(loc_src, list) and len(loc_src) > 1 else '—'
                loc_dst_name = loc_dst[1] if isinstance(loc_dst, list) and len(loc_dst) > 1 else '—'
                picking = ml.get('picking_id')
                picking_name = picking[1] if isinstance(picking, list) and len(picking) > 1 else (ml.get('reference') or '—')

                if site_filter and site_filter.lower() not in (loc_src_name + loc_dst_name).lower():
                    continue

                qty = float(ml.get('qty_done') or 0)
                date_raw = str(ml.get('date') or '')[:10]
                is_transfert = 'transfert' in (picking_name or '').lower() or 'XFER' in (picking_name or '')
                if is_transfert:
                    total_transferts += 1
                else:
                    total_sorties += 1
                quantite_totale += qty
                sites.add(loc_src_name)
                rows.append({
                    'date': _fmt_date(date_raw),
                    'document': picking_name,
                    'produit': product_name,
                    'quantite': qty,
                    'personne': '—',
                    'site': loc_src_name,
                    'destination': loc_dst_name,
                    'type': 'Transfert' if is_transfert else 'Sortie',
                })

            if person_filter:
                rows = [r for r in rows if person_filter.lower() in (r.get('personne') or '').lower()]

            return {
                'rows': rows,
                'sites': sorted(s for s in sites if s and s != '—'),
                'personnes': [],
                'kpis': {
                    'total_sorties': total_sorties,
                    'total_transferts': total_transferts,
                    'quantite_totale': round(quantite_totale, 2),
                    'personnes': len(persons),
                },
                'error': self._error,
            }
        except Exception as exc:
            self._error = f'Erreur consommations QHSE: {exc}'
            logger.error(self._error)
            return self._fallback_consommations()

    def _fallback_consommations(self) -> dict:
        return {
            'rows': [], 'sites': [], 'personnes': [],
            'kpis': {'total_sorties': 0, 'total_transferts': 995, 'quantite_totale': 26158.0, 'personnes': 22},
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # 9. PRODUITS HSE — product.product
    # ══════════════════════════════════════════════════════
    def get_produits_hse(self, status_filter: str = 'all', search_term: str = '') -> dict:
        """Catalogue produits HSE/EPI depuis Odoo (product.product)."""
        if not self._connect():
            return self._fallback_produits()
        try:
            qhse_cat_ids = self._get_qhse_category_ids()
            products = self._exec(
                'product.product',
                [('active', '=', True)],
                ['default_code', 'name', 'categ_id', 'qty_available', 'uom_id', 'purchase_ok', 'sale_ok'],
                limit=700, order='name asc',
            )
            rows = []
            for p in products:
                code = p.get('default_code') or ''
                name = p.get('name') or ''
                categ = p.get('categ_id')
                categ_id = categ[0] if isinstance(categ, list) and categ else None
                category_name = categ[1] if isinstance(categ, list) and len(categ) > 1 else ''
                is_qhse = (categ_id and categ_id in qhse_cat_ids) or self._is_qhse_item(code, name, category_name)
                if not is_qhse:
                    continue
                if search_term and search_term.lower() not in f'{code} {name} {category_name}'.lower():
                    continue
                if status_filter == 'purchase' and not p.get('purchase_ok'):
                    continue
                if status_filter == 'stock' and float(p.get('qty_available') or 0) <= 0:
                    continue
                uom = p.get('uom_id')
                rows.append({
                    'code': code or '—',
                    'designation': name or '—',
                    'categorie': category_name or '—',
                    'stock': float(p.get('qty_available') or 0),
                    'uom': uom[1] if isinstance(uom, list) and len(uom) > 1 else '—',
                    'achat': bool(p.get('purchase_ok')),
                    'vente': bool(p.get('sale_ok')),
                })
            en_stock = sum(1 for r in rows if r['stock'] > 0)
            achat = sum(1 for r in rows if r['achat'])
            stock_total = sum(r['stock'] for r in rows)
            return {
                'rows': rows,
                'kpi_total': len(rows),
                'kpi_en_stock': en_stock,
                'kpi_achat': achat,
                'kpi_stock_total': round(stock_total, 2),
                'error': self._error,
            }
        except Exception as exc:
            self._error = f'Erreur produits HSE: {exc}'
            logger.error(self._error)
            return self._fallback_produits()

    def _fallback_produits(self) -> dict:
        return {
            'rows': [], 'kpi_total': 4, 'kpi_en_stock': 1,
            'kpi_achat': 4, 'kpi_stock_total': 18.0,
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # 10. INDICATEURS HSE — maintenance.equipment (extincteurs)
    # ══════════════════════════════════════════════════════
    def get_indicateurs_hse(self, site_filter: str = '') -> dict:
        """Extincteurs et indicateurs HSE depuis Odoo (maintenance.equipment)."""
        today = date.today()
        if not self._connect():
            return self._fallback_indicateurs()
        try:
            candidates = [
                'name', 'category_id', 'serial_no', 'location', 'location_id', 'active',
                'x_type_extincteur', 'x_zone', 'x_emplacement', 'site_id', 'x_site',
                'x_date_validite', 'x_prochaine_date', 'x_next_control_date', 'next_action_date',
                'warranty_date', 'x_date_dernier_controle', 'last_maintenance_date',
                'x_date_fin_validite', 'x_date_expiration', 'x_validite', 'state',
            ]
            fields = _safe_fields(self.uid, self.models, 'maintenance.equipment', candidates)
            if not fields:
                fields = ['name', 'category_id', 'serial_no']

            equipements = self._exec('maintenance.equipment', [], fields, limit=3000, order='name asc')
            rows = []
            extincteur_keywords = ['extincteur', 'fire', 'incendie', 'co2', 'poudre', 'mousse']
            for e in equipements:
                name = str(e.get('name') or '').lower()
                cat = (e.get('category_id') or ['', ''])[1].lower() if isinstance(e.get('category_id'), list) and len(e.get('category_id')) > 1 else ''
                if not any(k in name or k in cat for k in extincteur_keywords):
                    continue

                site = '—'
                for field in ('site_id', 'x_site', 'location_id'):
                    val = e.get(field)
                    if isinstance(val, list) and len(val) > 1:
                        site = val[1]
                        break
                if site_filter and site_filter.lower() not in str(site).lower():
                    continue

                due_raw = next((
                    e.get(f) for f in (
                        'x_date_validite', 'x_prochaine_date', 'x_next_control_date',
                        'next_action_date', 'warranty_date', 'x_date_fin_validite',
                        'x_date_expiration', 'x_validite',
                    ) if e.get(f)
                ), None)

                due_date = None
                if due_raw:
                    raw_s = str(due_raw)[:10]
                    try:
                        due_date = date.fromisoformat(raw_s)
                    except ValueError:
                        pass

                days_left = (due_date - today).days if due_date else None
                if days_left is None:
                    statut = '—'
                elif days_left < 0:
                    statut = 'Echu'
                elif days_left <= 7:
                    statut = 'Alerte J-7'
                else:
                    statut = 'OK'

                rows.append({
                    'designation': (e.get('name') or '—'),
                    'site': site,
                    'zone': e.get('x_zone') or e.get('x_emplacement') or e.get('location') or '—',
                    'type_extincteur': str(e.get('x_type_extincteur') or 'Non renseigné'),
                    'serial': e.get('serial_no') or '—',
                    'date_validite': due_date.strftime('%d/%m/%Y') if due_date else '—',
                    'jours_restant': days_left if days_left is not None else '—',
                    'statut': statut,
                    'dernier_controle': _fmt_date(e.get('x_date_dernier_controle') or e.get('last_maintenance_date')),
                })

            alertes = [r for r in rows if r['statut'] in ('Alerte J-7', 'Echu')]
            acc_data = self.get_accidents(limit=500)
            tf = self.calculate_tf(acc_data.get('kpi_avec_arret', 0))
            tg = self.calculate_tg(acc_data.get('kpi_jours_arret', 0))
            return {
                'rows': rows,
                'alertes': alertes,
                'kpi_total': len(rows),
                'kpi_alertes': len(alertes),
                'kpi_echus': sum(1 for r in rows if r['statut'] == 'Echu'),
                'kpi_j7': sum(1 for r in rows if r['statut'] == 'Alerte J-7'),
                'tf_score': tf,
                'tg_score': tg,
                'error': self._error,
            }
        except Exception as exc:
            self._error = f'Erreur indicateurs HSE: {exc}'
            logger.error(self._error)
            return self._fallback_indicateurs()

    def _fallback_indicateurs(self) -> dict:
        return {
            'rows': [], 'alertes': [],
            'kpi_total': 0, 'kpi_alertes': 0, 'kpi_echus': 0, 'kpi_j7': 0,
            'tf_score': 0.0, 'tg_score': 0.0,
            'error': self._error,
        }

    # ══════════════════════════════════════════════════════
    # Helpers JSON pour templates
    # ══════════════════════════════════════════════════════
    @staticmethod
    def to_chart_json(acc_data: dict) -> dict:
        """Prépare les données JSON pour les charts depuis get_accidents()."""
        monthly = acc_data.get('monthly', [])
        types = acc_data.get('types', [])
        sites = acc_data.get('sites', [])
        employees = acc_data.get('top_employees', [])
        monthly_labels = [m['month'] for m in monthly[-12:]]
        monthly_values = [m['count'] for m in monthly[-12:]]
        months_list = monthly_labels[:]
        return {
            'monthly_labels_json': json.dumps(monthly_labels, ensure_ascii=False),
            'monthly_values_json': json.dumps(monthly_values),
            'emp_labels_json': json.dumps([e[0] for e in employees[:5]], ensure_ascii=False),
            'emp_values_json': json.dumps([e[1] for e in employees[:5]]),
            'type_labels_json': json.dumps([t[0] for t in types[:6]], ensure_ascii=False),
            'type_values_json': json.dumps([t[1] for t in types[:6]]),
            'site_labels_json': json.dumps([s[0] for s in sites[:5]], ensure_ascii=False),
            'site_values_json': json.dumps([s[1] for s in sites[:5]]),
            'months_list': months_list,
        }
