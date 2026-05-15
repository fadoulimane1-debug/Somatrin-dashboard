"""
SOMATRIN — Odoo 16 JSON-RPC client & agrégats QHSE.

Utilise django.conf.settings (ODOO_URL, ODOO_DB / ODOO_DATABASE, ODOO_USER, ODOO_PASSWORD).
Alternative XML-RPC : reporting.services.qhse_service.QHSEService
"""
from __future__ import annotations

import logging
import random
from datetime import date, datetime
from typing import Any

import requests
from django.conf import settings

logger = logging.getLogger(__name__)


def _odoo_db() -> str:
    return getattr(settings, 'ODOO_DB', None) or getattr(settings, 'ODOO_DATABASE', '') or ''


def _odoo_password() -> str:
    return getattr(settings, 'ODOO_PASS', None) or getattr(settings, 'ODOO_PASSWORD', '') or ''


def _jsonrpc_url() -> str:
    base = (getattr(settings, 'ODOO_URL', '') or '').rstrip('/')
    return f'{base}/jsonrpc'


class OdooServiceManager:
    """Connexion Odoo 16 via JSON-RPC + méthodes métier QHSE."""

    def __init__(self):
        self.db = _odoo_db()
        self.login = getattr(settings, 'ODOO_USER', '') or getattr(settings, 'ODOO_USERNAME', '')
        self.password = _odoo_password()
        self.verify_ssl = getattr(settings, 'ODOO_SSL_VERIFY', True)
        self.uid: int | None = None
        self._session = requests.Session()

    def test_connection(self) -> bool:
        try:
            uid = self._authenticate()
            return uid is not None and uid > 0
        except Exception as exc:
            logger.error('Odoo JSON-RPC test_connection: %s', exc)
            return False

    def _authenticate(self) -> int | None:
        if self.uid is not None:
            return self.uid
        result = self._call(
            'common',
            'authenticate',
            self.db,
            self.login,
            self.password,
            {},
        )
        if isinstance(result, int) and result > 0:
            self.uid = result
            logger.info('Odoo JSON-RPC authentifie uid=%s', self.uid)
            return self.uid
        logger.warning('Odoo JSON-RPC authenticate echoue: %s', result)
        return None

    def _call(self, service: str, method: str, *args) -> Any:
        payload = {
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'service': service,
                'method': method,
                'args': list(args),
            },
            'id': random.randint(1, 2**31 - 1),
        }
        resp = self._session.post(
            _jsonrpc_url(),
            json=payload,
            timeout=120,
            verify=self.verify_ssl,
        )
        resp.raise_for_status()
        body = resp.json()
        if body.get('error'):
            raise RuntimeError(body['error'])
        return body.get('result')

    def execute_kw(self, model: str, method: str, args: list | None = None, kwargs: dict | None = None) -> Any:
        uid = self._authenticate()
        if not uid:
            return []
        args = args if args is not None else []
        kwargs = kwargs if kwargs is not None else {}
        return self._call(
            'object',
            'execute_kw',
            self.db,
            uid,
            self.password,
            model,
            method,
            args,
            kwargs,
        )

    # ── Données agrégées (consommable par API / future migration) ─────────

    def get_sites(self) -> list[dict]:
        """Emplacements internes pour filtres."""
        try:
            recs = self.execute_kw(
                'stock.location',
                'search_read',
                [[['usage', '=', 'internal']]],
                {'fields': ['id', 'complete_name', 'name'], 'limit': 500, 'order': 'complete_name'},
            )
            return [{'id': r['id'], 'name': r.get('complete_name') or r.get('name')} for r in (recs or [])]
        except Exception as exc:
            logger.error('get_sites: %s', exc)
            return []

    def get_dashboard_kpis(self) -> dict:
        """KPI dashboard — s'appuie sur quality.check + project.task (+ formations)."""
        from reporting.services.qhse_service import QHSEService

        svc = QHSEService()
        acc = svc.get_accidents(limit=2000)
        avec_arret = int(acc.get('kpi_avec_arret') or 0)
        jours = int(acc.get('kpi_jours_arret') or 0)
        heures = getattr(svc, 'HEURES_ANNUELLES', 180_000) or 180_000
        tf = round((avec_arret / heures) * 1_000_000, 2) if avec_arret and heures else 0.0
        tg = round((jours / heures) * 1_000, 3) if jours and heures else 0.0
        conformite = float(svc.get_conformity_score())
        formations = int(svc.count_formations())

        src = 'fallback'
        if svc._connect():
            if svc._model_exists('quality.check'):
                src = 'quality.check'
            elif svc._model_exists('quality.alert'):
                src = 'quality.alert'

        return {
            'accidents_year': int(acc.get('kpi_total') or 0),
            'accidents_avec_arret': avec_arret,
            'jours_arret': jours,
            'tf_score': tf,
            'tg_score': tg,
            'conformite': conformite,
            'formations': formations,
            'timestamp': datetime.now().isoformat(timespec='seconds'),
            'source': src,
            'error': acc.get('error'),
        }
