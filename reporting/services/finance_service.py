import logging
import ssl
import xmlrpc.client
from collections import defaultdict
from datetime import date, datetime, timedelta

from django.conf import settings


logger = logging.getLogger(__name__)


class FinanceService:
    """Service Finance connecte a Odoo via XML-RPC."""

    def __init__(self):
        verify_ssl = getattr(settings, 'ODOO_SSL_VERIFY', True)
        kwargs = {}
        if settings.ODOO_URL.startswith('https://') and not verify_ssl:
            kwargs['context'] = ssl._create_unverified_context()
        self.db = settings.ODOO_DB
        self.user = settings.ODOO_USER
        self.pwd = settings.ODOO_PASS
        self.common = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/common', **kwargs)
        self.models = xmlrpc.client.ServerProxy(f'{settings.ODOO_URL}/xmlrpc/2/object', **kwargs)
        self.uid = self.common.authenticate(self.db, self.user, self.pwd, {})

    def _search_read(self, model, domain, fields, limit=2000, order='id desc'):
        if not self.uid:
            return []
        try:
            return self.models.execute_kw(
                self.db, self.uid, self.pwd,
                model, 'search_read',
                [domain],
                {'fields': fields, 'limit': limit, 'order': order},
            ) or []
        except Exception as exc:
            logger.warning('Finance search_read failed (%s): %s', model, exc)
            return []

    def _month_key(self, dt_str):
        if not dt_str:
            return ''
        s = str(dt_str)[:10]
        try:
            d = datetime.strptime(s, '%Y-%m-%d').date()
            return d.strftime('%Y-%m')
        except Exception:
            return ''

    def _format_money(self, val):
        try:
            x = float(val or 0)
        except Exception:
            x = 0.0
        s = f'{x:,.2f}'.replace(',', ' ').replace('.', ',')
        return f'{s} MAD'

    def _invoice_model(self):
        """Priorite account.move (moderne), fallback account.invoice."""
        probe = self._search_read('account.move', [('move_type', 'in', ['out_invoice'])], ['id'], limit=1)
        if probe:
            return 'account.move'
        return 'account.invoice'

    def get_invoices_clients(self, limit=2000):
        model = self._invoice_model()
        if model == 'account.move':
            fields = [
                'name', 'invoice_date', 'invoice_date_due', 'partner_id',
                'amount_untaxed', 'amount_tax', 'amount_total',
                'state', 'payment_state', 'move_type', 'invoice_origin', 'ref',
            ]
            rows = self._search_read(
                'account.move',
                [('move_type', '=', 'out_invoice')],
                fields,
                limit=limit,
                order='invoice_date desc,id desc',
            )
            out = []
            for r in rows:
                p = r.get('partner_id')
                out.append({
                    'id': r.get('id'),
                    'number': r.get('name') or 'N/A',
                    'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                    'invoice_date': (r.get('invoice_date') or '')[:10],
                    'due_date': (r.get('invoice_date_due') or '')[:10],
                    'amount_ht': float(r.get('amount_untaxed') or 0),
                    'amount_tax': float(r.get('amount_tax') or 0),
                    'amount_ttc': float(r.get('amount_total') or 0),
                    'state': r.get('state') or 'draft',
                    'payment_state': r.get('payment_state') or 'not_paid',
                    'description': r.get('invoice_origin') or r.get('ref') or '',
                })
            return out

        fields = [
            'number', 'date_invoice', 'date_due', 'partner_id',
            'amount_untaxed', 'amount_tax', 'amount_total',
            'state', 'type', 'reference',
        ]
        rows = self._search_read(
            'account.invoice',
            [('type', '=', 'out_invoice')],
            fields,
            limit=limit,
            order='date_invoice desc,id desc',
        )
        out = []
        for r in rows:
            p = r.get('partner_id')
            st = r.get('state') or 'draft'
            payment_state = 'paid' if st in ('paid',) else ('partial' if st in ('open',) else 'not_paid')
            out.append({
                'id': r.get('id'),
                'number': r.get('number') or 'N/A',
                'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                'invoice_date': (r.get('date_invoice') or '')[:10],
                'due_date': (r.get('date_due') or '')[:10],
                'amount_ht': float(r.get('amount_untaxed') or 0),
                'amount_tax': float(r.get('amount_tax') or 0),
                'amount_ttc': float(r.get('amount_total') or 0),
                'state': st,
                'payment_state': payment_state,
                'description': r.get('reference') or '',
            })
        return out

    def get_invoices_fournisseurs(self, limit=2000):
        model = self._invoice_model()
        if model == 'account.move':
            fields = [
                'name', 'invoice_date', 'invoice_date_due', 'partner_id',
                'amount_untaxed', 'amount_tax', 'amount_total',
                'state', 'payment_state', 'move_type', 'invoice_origin', 'ref',
            ]
            rows = self._search_read(
                'account.move',
                [('move_type', '=', 'in_invoice')],
                fields,
                limit=limit,
                order='invoice_date desc,id desc',
            )
            out = []
            for r in rows:
                p = r.get('partner_id')
                out.append({
                    'id': r.get('id'),
                    'number': r.get('name') or 'N/A',
                    'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                    'invoice_date': (r.get('invoice_date') or '')[:10],
                    'due_date': (r.get('invoice_date_due') or '')[:10],
                    'amount_ht': float(r.get('amount_untaxed') or 0),
                    'amount_tax': float(r.get('amount_tax') or 0),
                    'amount_ttc': float(r.get('amount_total') or 0),
                    'state': r.get('state') or 'draft',
                    'payment_state': r.get('payment_state') or 'not_paid',
                    'description': r.get('invoice_origin') or r.get('ref') or '',
                })
            return out

        fields = [
            'number', 'date_invoice', 'date_due', 'partner_id',
            'amount_untaxed', 'amount_tax', 'amount_total',
            'state', 'type', 'reference',
        ]
        rows = self._search_read(
            'account.invoice',
            [('type', '=', 'in_invoice')],
            fields,
            limit=limit,
            order='date_invoice desc,id desc',
        )
        out = []
        for r in rows:
            p = r.get('partner_id')
            st = r.get('state') or 'draft'
            payment_state = 'paid' if st in ('paid',) else ('partial' if st in ('open',) else 'not_paid')
            out.append({
                'id': r.get('id'),
                'number': r.get('number') or 'N/A',
                'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                'invoice_date': (r.get('date_invoice') or '')[:10],
                'due_date': (r.get('date_due') or '')[:10],
                'amount_ht': float(r.get('amount_untaxed') or 0),
                'amount_tax': float(r.get('amount_tax') or 0),
                'amount_ttc': float(r.get('amount_total') or 0),
                'state': st,
                'payment_state': payment_state,
                'description': r.get('reference') or '',
            })
        return out

    def get_credit_notes(self, limit=2000):
        model = self._invoice_model()
        if model == 'account.move':
            rows = self._search_read(
                'account.move',
                [('move_type', 'in', ['out_refund', 'in_refund'])],
                ['name', 'invoice_date', 'partner_id', 'amount_untaxed', 'amount_tax', 'amount_total', 'state', 'move_type', 'reversed_entry_id', 'ref'],
                limit=limit,
            )
            out = []
            for r in rows:
                p = r.get('partner_id')
                rev = r.get('reversed_entry_id')
                out.append({
                    'id': r.get('id'),
                    'number': r.get('name') or 'N/A',
                    'date': (r.get('invoice_date') or '')[:10],
                    'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                    'invoice_ref': rev[1] if isinstance(rev, list) and len(rev) > 1 else (r.get('ref') or ''),
                    'amount_ht': float(r.get('amount_untaxed') or 0),
                    'amount_ttc': float(r.get('amount_total') or 0),
                    'reason': r.get('ref') or '',
                    'state': r.get('state') or 'draft',
                    'type': 'client' if r.get('move_type') == 'out_refund' else 'fournisseur',
                })
            return out
        rows = self._search_read(
            'account.invoice',
            [('type', 'in', ['out_refund', 'in_refund'])],
            ['number', 'date_invoice', 'partner_id', 'origin', 'amount_untaxed', 'amount_total', 'comment', 'state', 'type'],
            limit=limit,
        )
        out = []
        for r in rows:
            p = r.get('partner_id')
            out.append({
                'id': r.get('id'),
                'number': r.get('number') or 'N/A',
                'date': (r.get('date_invoice') or '')[:10],
                'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                'invoice_ref': r.get('origin') or '',
                'amount_ht': float(r.get('amount_untaxed') or 0),
                'amount_ttc': float(r.get('amount_total') or 0),
                'reason': r.get('comment') or '',
                'state': r.get('state') or 'draft',
                'type': 'client' if r.get('type') == 'out_refund' else 'fournisseur',
            })
        return out

    def get_payments(self, limit=3000):
        rows = self._search_read(
            'account.payment',
            [],
            ['name', 'payment_date', 'date', 'partner_id', 'amount', 'payment_type', 'payment_method_line_id', 'payment_method_id', 'journal_id', 'state', 'ref'],
            limit=limit,
            order='payment_date desc,id desc',
        )
        out = []
        for r in rows:
            p = r.get('partner_id')
            method = r.get('payment_method_line_id') or r.get('payment_method_id')
            out.append({
                'id': r.get('id'),
                'number': r.get('name') or 'N/A',
                'date': (r.get('payment_date') or r.get('date') or '')[:10],
                'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                'amount': float(r.get('amount') or 0),
                'payment_method': method[1] if isinstance(method, list) and len(method) > 1 else 'N/A',
                'journal': (r.get('journal_id')[1] if isinstance(r.get('journal_id'), list) and len(r.get('journal_id')) > 1 else 'N/A'),
                'value_date': (r.get('date') or r.get('payment_date') or '')[:10],
                'payment_type': r.get('payment_type') or 'outbound',
                'state': r.get('state') or 'draft',
                'memo': r.get('ref') or '',
            })
        return out

    def get_partners_stats(self):
        partners = self._search_read(
            'res.partner',
            [('is_company', '=', True)],
            ['name', 'customer_rank', 'supplier_rank'],
            limit=5000,
            order='name asc',
        )
        clients = [p for p in partners if int(p.get('customer_rank') or 0) > 0]
        suppliers = [p for p in partners if int(p.get('supplier_rank') or 0) > 0]
        return {
            'clients_count': len(clients),
            'suppliers_count': len(suppliers),
        }

    def build_dashboard(self):
        client_invoices = self.get_invoices_clients(limit=5000)
        supplier_invoices = self.get_invoices_fournisseurs(limit=5000)
        payments = self.get_payments(limit=5000)
        partners_stats = self.get_partners_stats()

        ca_ht = sum(x['amount_ht'] for x in client_invoices)
        achats_ht = sum(x['amount_ht'] for x in supplier_invoices)
        encaissements = sum(x['amount'] for x in payments if x['payment_type'] == 'inbound')
        decaissements = sum(x['amount'] for x in payments if x['payment_type'] == 'outbound')
        creances = sum(x['amount_ttc'] for x in client_invoices if x['payment_state'] in ('not_paid', 'partial', 'in_payment'))
        dettes = sum(x['amount_ttc'] for x in supplier_invoices if x['payment_state'] in ('not_paid', 'partial', 'in_payment'))
        ratio_liq = (encaissements / decaissements * 100.0) if decaissements > 0 else 0
        marge_nette = ((ca_ht - achats_ht) / ca_ht * 100.0) if ca_ht > 0 else 0

        overdue_count = 0
        overdue_amount = 0.0
        today = date.today()
        for inv in client_invoices:
            if inv['payment_state'] in ('paid', 'reversed'):
                continue
            dd = inv.get('due_date')
            try:
                d = datetime.strptime(dd, '%Y-%m-%d').date()
                if d < (today - timedelta(days=30)):
                    overdue_count += 1
                    overdue_amount += inv['amount_ttc']
            except Exception:
                continue

        # Series monthly
        monthly_ca = defaultdict(float)
        monthly_pay = defaultdict(float)
        monthly_out = defaultdict(float)
        monthly_count = defaultdict(int)

        for inv in client_invoices:
            mk = self._month_key(inv.get('invoice_date'))
            if mk:
                monthly_ca[mk] += inv['amount_ht']
                monthly_count[mk] += 1
        for p in payments:
            mk = self._month_key(p.get('date'))
            if not mk:
                continue
            if p['payment_type'] == 'inbound':
                monthly_pay[mk] += p['amount']
            else:
                monthly_out[mk] += p['amount']

        months = []
        cur = date.today().replace(day=1)
        for i in range(11, -1, -1):
            m = (cur.replace(day=1) - timedelta(days=1))
            for _ in range(i):
                m = m.replace(day=1) - timedelta(days=1)
            key = m.replace(day=1).strftime('%Y-%m')
            months.append(key)
        months = sorted(set(months))
        labels = [datetime.strptime(m + '-01', '%Y-%m-%d').strftime('%b %Y') for m in months]

        ca_series = [round(monthly_ca.get(m, 0), 2) for m in months]
        pay_series = [round(monthly_pay.get(m, 0), 2) for m in months]
        out_series = [round(monthly_out.get(m, 0), 2) for m in months]
        inv_count_series = [monthly_count.get(m, 0) for m in months]

        top_clients = defaultdict(float)
        for inv in client_invoices:
            top_clients[inv['partner']] += inv['amount_ht']
        top_clients_sorted = sorted(top_clients.items(), key=lambda x: x[1], reverse=True)[:10]

        avg_ticket = (ca_ht / max(1, len(client_invoices)))
        dso = 0
        dpo = 0

        paid_client = [x for x in client_invoices if x['payment_state'] in ('paid', 'reversed')]
        paid_supplier = [x for x in supplier_invoices if x['payment_state'] in ('paid', 'reversed')]
        if paid_client:
            vals = []
            for i in paid_client:
                try:
                    d1 = datetime.strptime(i['invoice_date'], '%Y-%m-%d').date()
                    d2 = datetime.strptime(i['due_date'], '%Y-%m-%d').date() if i.get('due_date') else d1
                    vals.append(max(0, (d2 - d1).days))
                except Exception:
                    pass
            if vals:
                dso = round(sum(vals) / len(vals), 1)
        if paid_supplier:
            vals = []
            for i in paid_supplier:
                try:
                    d1 = datetime.strptime(i['invoice_date'], '%Y-%m-%d').date()
                    d2 = datetime.strptime(i['due_date'], '%Y-%m-%d').date() if i.get('due_date') else d1
                    vals.append(max(0, (d2 - d1).days))
                except Exception:
                    pass
            if vals:
                dpo = round(sum(vals) / len(vals), 1)

        return {
            'kpi': {
                'ca_ht': ca_ht,
                'achats_ht': achats_ht,
                'encaissements': encaissements,
                'dettes': dettes,
                'creances': creances,
                'overdue_count': overdue_count,
                'overdue_amount': overdue_amount,
                'ratio_liq': ratio_liq,
                'marge_nette': marge_nette,
            },
            'stats': {
                'clients_count': partners_stats['clients_count'],
                'suppliers_count': partners_stats['suppliers_count'],
                'ticket_moyen': avg_ticket,
                'dso': dso,
                'dpo': dpo,
            },
            'charts': {
                'labels': labels,
                'ca_series': ca_series,
                'pay_series': pay_series,
                'out_series': out_series,
                'invoice_count_series': inv_count_series,
                'top_clients_labels': [x[0] for x in top_clients_sorted],
                'top_clients_values': [round(x[1], 2) for x in top_clients_sorted],
                'creances_pie': [
                    round(sum(x['amount_ttc'] for x in client_invoices if x['payment_state'] == 'paid'), 2),
                    round(sum(x['amount_ttc'] for x in client_invoices if x['payment_state'] in ('in_payment', 'partial')), 2),
                    round(sum(x['amount_ttc'] for x in client_invoices if x['payment_state'] in ('not_paid',)), 2),
                ],
                'dettes_pie': [
                    round(sum(x['amount_ttc'] for x in supplier_invoices if x['payment_state'] == 'paid'), 2),
                    round(sum(x['amount_ttc'] for x in supplier_invoices if x['payment_state'] in ('in_payment', 'partial')), 2),
                    round(sum(x['amount_ttc'] for x in supplier_invoices if x['payment_state'] in ('not_paid',)), 2),
                ],
            },
            'format_money': self._format_money,
        }


def _sort_key(month_str):
    try:
        return datetime.strptime(month_str, '%Y-%m')
    except Exception:
        return datetime.min


def _fmt_date_iso(v):
    if not v:
        return ''
    s = str(v)[:10]
    try:
        d = datetime.strptime(s, '%Y-%m-%d').date()
        return d.strftime('%d/%m/%Y')
    except Exception:
        return s


def _as_float(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def _safe_invoice_number(primary, secondary, row_id, prefix='FAC'):
    """Retourne un numero facture propre, jamais '/' ou vide."""
    for candidate in (primary, secondary):
        val = str(candidate or '').strip()
        if val and val != '/':
            return val
    return f'{prefix}-{row_id or "N/A"}'


def _detect_invoice_model(models, db, uid, pwd):
    try:
        probe = models.execute_kw(
            db, uid, pwd,
            'account.move', 'search_read',
            [[('move_type', 'in', ['out_invoice', 'in_invoice', 'out_refund', 'in_refund'])]],
            {'fields': ['id'], 'limit': 1},
        ) or []
        if probe:
            return 'account.move'
    except Exception:
        pass
    return 'account.invoice'


def fetch_invoices(uid, models, invoice_type, limit=2000):
    """Retourne une liste normalisée de factures/avoirs."""
    if not uid:
        return []
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS
    model = _detect_invoice_model(models, db, uid, pw)
    rows = []
    if model == 'account.move':
        raw = models.execute_kw(
            db, uid, pw,
            'account.move', 'search_read',
            [[('move_type', '=', invoice_type)]],
            {'fields': ['name', 'invoice_date', 'invoice_date_due', 'partner_id', 'amount_untaxed',
                        'amount_tax', 'amount_total', 'state', 'payment_state', 'amount_residual',
                        'invoice_origin', 'ref'],
             'limit': limit, 'order': 'invoice_date desc,id desc'},
        ) or []
        for r in raw:
            d_iso = (r.get('invoice_date') or '')[:10]
            dd_iso = (r.get('invoice_date_due') or '')[:10]
            p = r.get('partner_id')
            residual = _as_float(r.get('amount_residual'))
            total = _as_float(r.get('amount_total'))
            paid = max(0.0, total - residual)
            is_overdue = False
            days_overdue = 0
            if dd_iso and residual > 0:
                try:
                    due = datetime.strptime(dd_iso, '%Y-%m-%d').date()
                    if due < date.today():
                        is_overdue = True
                        days_overdue = (date.today() - due).days
                except Exception:
                    pass
            rows.append({
                'id': r.get('id'),
                'name': _safe_invoice_number(r.get('name'), r.get('ref'), r.get('id'), prefix='MOV'),
                'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
                'date_iso': d_iso,
                'date_due_iso': dd_iso,
                'date': _fmt_date_iso(d_iso),
                'date_due': _fmt_date_iso(dd_iso),
                'amount_ht': _as_float(r.get('amount_untaxed')),
                'amount_tax': _as_float(r.get('amount_tax')),
                'amount_ttc': total,
                'amount_residual': residual,
                'amount_paid': paid,
                'state': r.get('state') or 'draft',
                'payment_state': r.get('payment_state') or ('paid' if residual <= 0.0001 else 'not_paid'),
                'is_overdue': is_overdue,
                'days_overdue': days_overdue,
                'description': r.get('invoice_origin') or r.get('ref') or '',
            })
        return rows

    # account.invoice fallback
    raw = models.execute_kw(
        db, uid, pw,
        'account.invoice', 'search_read',
        [[('type', '=', invoice_type)]],
        {'fields': ['number', 'date_invoice', 'date_due', 'partner_id', 'amount_untaxed',
                    'amount_tax', 'amount_total', 'residual', 'state', 'reference'],
         'limit': limit, 'order': 'date_invoice desc,id desc'},
    ) or []
    for r in raw:
        d_iso = (r.get('date_invoice') or '')[:10]
        dd_iso = (r.get('date_due') or '')[:10]
        p = r.get('partner_id')
        residual = _as_float(r.get('residual'))
        total = _as_float(r.get('amount_total'))
        paid = max(0.0, total - residual)
        state = r.get('state') or 'draft'
        pstate = 'paid' if residual <= 0.0001 or state == 'paid' else ('partial' if paid > 0 else 'not_paid')
        is_overdue = False
        days_overdue = 0
        if dd_iso and residual > 0:
            try:
                due = datetime.strptime(dd_iso, '%Y-%m-%d').date()
                if due < date.today():
                    is_overdue = True
                    days_overdue = (date.today() - due).days
            except Exception:
                pass
        rows.append({
            'id': r.get('id'),
            'name': _safe_invoice_number(r.get('number'), r.get('reference'), r.get('id')),
            'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
            'date_iso': d_iso,
            'date_due_iso': dd_iso,
            'date': _fmt_date_iso(d_iso),
            'date_due': _fmt_date_iso(dd_iso),
            'amount_ht': _as_float(r.get('amount_untaxed')),
            'amount_tax': _as_float(r.get('amount_tax')),
            'amount_ttc': total,
            'amount_residual': residual,
            'amount_paid': paid,
            'state': state,
            'payment_state': pstate,
            'is_overdue': is_overdue,
            'days_overdue': days_overdue,
            'description': r.get('reference') or '',
        })
    return rows


def fetch_payments(uid, models, limit=2000):
    if not uid:
        return []
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS
    raw = models.execute_kw(
        db, uid, pw,
        'account.payment', 'search_read',
        [[]],
        {'fields': ['name', 'payment_date', 'date', 'partner_id', 'amount', 'payment_type', 'state',
                    'payment_method_line_id', 'payment_method_id', 'journal_id', 'ref'],
         'limit': limit, 'order': 'payment_date desc,id desc'},
    ) or []
    out = []
    for r in raw:
        dt_iso = (r.get('payment_date') or r.get('date') or '')[:10]
        p = r.get('partner_id')
        pm = r.get('payment_method_line_id') or r.get('payment_method_id')
        ptype = r.get('payment_type') or 'outbound'
        out.append({
            'id': r.get('id'),
            'name': r.get('name') or 'N/A',
            'date_iso': dt_iso,
            'date': _fmt_date_iso(dt_iso),
            'partner': p[1] if isinstance(p, list) and len(p) > 1 else 'N/A',
            'amount': _as_float(r.get('amount')),
            'payment_type': ptype,
            'type_label': 'Reçu' if ptype == 'inbound' else 'Envoyé',
            'state': r.get('state') or 'draft',
            'state_label': (r.get('state') or 'draft').capitalize(),
            'payment_method': pm[1] if isinstance(pm, list) and len(pm) > 1 else 'N/A',
            'journal': (r.get('journal_id')[1] if isinstance(r.get('journal_id'), list) and len(r.get('journal_id')) > 1 else 'N/A'),
            'ref': r.get('ref') or '',
        })
    return out


def monthly_amounts(rows, amount_key, n_months=12):
    d = defaultdict(float)
    for r in rows:
        m = ''
        if isinstance(r, dict):
            if r.get('date_iso'):
                m = str(r.get('date_iso'))[:7]
            elif r.get('date'):
                s = str(r.get('date'))
                if len(s) >= 10 and '/' in s:
                    try:
                        dd = datetime.strptime(s[:10], '%d/%m/%Y').date()
                        m = dd.strftime('%Y-%m')
                    except Exception:
                        m = ''
        if m:
            d[m] += _as_float(r.get(amount_key))
    if not d:
        return {}
    keys = sorted(d.keys(), key=_sort_key)[-n_months:]
    return {k: round(d[k], 2) for k in keys}


def top_partners(rows, limit=10):
    d = defaultdict(float)
    for r in rows:
        d[str(r.get('partner') or 'N/A')] += _as_float(r.get('amount_ht'))
    top = sorted(d.items(), key=lambda x: x[1], reverse=True)[:limit]
    return {k: round(v, 2) for k, v in top}


def payment_state_breakdown(rows):
    out = {'Payées': 0, 'Partielles': 0, 'Impayées': 0}
    for r in rows:
        s = (r.get('payment_state') or '').lower()
        if s == 'paid':
            out['Payées'] += _as_float(r.get('amount_ttc'))
        elif s in ('partial', 'in_payment'):
            out['Partielles'] += _as_float(r.get('amount_ttc'))
        else:
            out['Impayées'] += _as_float(r.get('amount_ttc'))
    return {k: round(v, 2) for k, v in out.items()}


def payment_method_breakdown(rows):
    d = defaultdict(float)
    for r in rows:
        d[str(r.get('payment_method') or 'N/A')] += _as_float(r.get('amount'))
    return {k: round(v, 2) for k, v in sorted(d.items(), key=lambda x: x[1], reverse=True)[:8]}


def compute_kpis(inv_c, inv_f, pays):
    ca_total = sum(_as_float(r.get('amount_ht')) for r in inv_c if r.get('state') == 'posted')
    achats = sum(_as_float(r.get('amount_ht')) for r in inv_f if r.get('state') == 'posted')
    encaiss = sum(_as_float(r.get('amount')) for r in pays if r.get('payment_type') == 'inbound')
    decaiss = sum(_as_float(r.get('amount')) for r in pays if r.get('payment_type') == 'outbound')
    creances_rows = [r for r in inv_c if r.get('state') == 'posted' and r.get('payment_state') != 'paid']
    creances = sum(_as_float(r.get('amount_residual')) for r in creances_rows)
    dettes_rows = [r for r in inv_f if r.get('state') == 'posted' and r.get('payment_state') != 'paid']
    dettes = sum(_as_float(r.get('amount_residual')) for r in dettes_rows)
    overdue_rows = [r for r in creances_rows if r.get('is_overdue')]
    mt_overdue = sum(_as_float(r.get('amount_residual')) for r in overdue_rows)
    paid_f = sum(_as_float(r.get('amount_ttc')) for r in inv_f if r.get('state') == 'posted' and r.get('payment_state') == 'paid')
    posted_f = sum(_as_float(r.get('amount_ttc')) for r in inv_f if r.get('state') == 'posted')

    ca_month = 0.0
    cur_m = date.today().strftime('%Y-%m')
    for r in inv_c:
        if r.get('state') == 'posted' and str(r.get('date_iso') or '').startswith(cur_m):
            ca_month += _as_float(r.get('amount_ht'))

    partners_c = {r.get('partner') for r in inv_c if r.get('partner')}
    partners_f = {r.get('partner') for r in inv_f if r.get('partner')}
    ticket_moyen = (ca_total / max(1, len([r for r in inv_c if r.get('state') == 'posted'])))

    liquidite = (encaiss / decaiss * 100.0) if decaiss > 0 else 0.0
    marge = ((ca_total - achats) / ca_total * 100.0) if ca_total > 0 else 0.0
    pct_encaisse = (encaiss / max(1.0, ca_total) * 100.0) if ca_total > 0 else 0.0
    pct_paye_f = (paid_f / max(1.0, posted_f) * 100.0) if posted_f > 0 else 0.0

    return {
        'ca_total': round(ca_total, 2),
        'ca_month': round(ca_month, 2),
        'achats': round(achats, 2),
        'encaiss': round(encaiss, 2),
        'decaiss': round(decaiss, 2),
        'creances': round(creances, 2),
        'dettes': round(dettes, 2),
        'nb_creances': len(creances_rows),
        'nb_overdue': len(overdue_rows),
        'mt_overdue': round(mt_overdue, 2),
        'liquidite': round(liquidite, 1),
        'marge': round(marge, 1),
        'pct_encaisse': round(pct_encaisse, 1),
        'pct_paye_f': round(pct_paye_f, 1),
        'nb_clients': len(partners_c),
        'nb_fournisseurs': len(partners_f),
        'nb_factures_clients': len(inv_c),
        'nb_factures_fournisseurs': len(inv_f),
        'ticket_moyen': round(ticket_moyen, 2),
    }


def fetch_invoice_detail(uid, models, invoice_id):
    if not uid:
        return None, []
    db = settings.ODOO_DB
    pw = settings.ODOO_PASS
    model = _detect_invoice_model(models, db, uid, pw)

    if model == 'account.move':
        inv = models.execute_kw(
            db, uid, pw,
            'account.move', 'search_read',
            [[('id', '=', int(invoice_id))]],
            {'fields': ['name', 'invoice_date', 'invoice_date_due', 'partner_id', 'amount_untaxed',
                        'amount_tax', 'amount_total', 'amount_residual', 'state', 'payment_state',
                        'journal_id', 'payment_reference', 'ref', 'move_type'],
             'limit': 1},
        ) or []
        if not inv:
            return None, []
        inv = inv[0]
        lines = models.execute_kw(
            db, uid, pw,
            'account.move.line', 'search_read',
            [[('move_id', '=', int(invoice_id)), ('display_type', '=', False)]],
            {'fields': ['name', 'quantity', 'price_unit', 'price_subtotal', 'price_total', 'tax_ids'],
             'limit': 200},
        ) or []
        invoice = {
            'id': inv.get('id'),
            'name': _safe_invoice_number(inv.get('name'), inv.get('ref'), inv.get('id'), prefix='MOV'),
            'date': _fmt_date_iso((inv.get('invoice_date') or '')[:10]),
            'date_due': _fmt_date_iso((inv.get('invoice_date_due') or '')[:10]),
            'partner': (inv.get('partner_id')[1] if isinstance(inv.get('partner_id'), list) and len(inv.get('partner_id')) > 1 else 'N/A'),
            'amount_ht': _as_float(inv.get('amount_untaxed')),
            'amount_tax': _as_float(inv.get('amount_tax')),
            'amount_ttc': _as_float(inv.get('amount_total')),
            'amount_residual': _as_float(inv.get('amount_residual')),
            'state': inv.get('state') or 'draft',
            'payment_state': inv.get('payment_state') or 'not_paid',
            'journal': (inv.get('journal_id')[1] if isinstance(inv.get('journal_id'), list) and len(inv.get('journal_id')) > 1 else 'N/A'),
            'memo': inv.get('payment_reference') or inv.get('ref') or '',
            'move_type': inv.get('move_type') or '',
        }
        out_lines = []
        for l in lines:
            out_lines.append({
                'description': l.get('name') or '',
                'quantity': _as_float(l.get('quantity') or 1),
                'price_unit': _as_float(l.get('price_unit')),
                'amount_ht': _as_float(l.get('price_subtotal')),
                'amount_ttc': _as_float(l.get('price_total')),
                'tax': '',
            })
        return invoice, out_lines

    inv = models.execute_kw(
        db, uid, pw,
        'account.invoice', 'search_read',
        [[('id', '=', int(invoice_id))]],
        {'fields': ['number', 'date_invoice', 'date_due', 'partner_id', 'amount_untaxed',
                    'amount_tax', 'amount_total', 'residual', 'state', 'reference', 'type'],
         'limit': 1},
    ) or []
    if not inv:
        return None, []
    inv = inv[0]
    lines = models.execute_kw(
        db, uid, pw,
        'account.invoice.line', 'search_read',
        [[('invoice_id', '=', int(invoice_id))]],
        {'fields': ['name', 'quantity', 'price_unit', 'price_subtotal', 'price_total'],
         'limit': 200},
    ) or []
    invoice = {
        'id': inv.get('id'),
        'name': _safe_invoice_number(inv.get('number'), inv.get('reference'), inv.get('id')),
        'date': _fmt_date_iso((inv.get('date_invoice') or '')[:10]),
        'date_due': _fmt_date_iso((inv.get('date_due') or '')[:10]),
        'partner': (inv.get('partner_id')[1] if isinstance(inv.get('partner_id'), list) and len(inv.get('partner_id')) > 1 else 'N/A'),
        'amount_ht': _as_float(inv.get('amount_untaxed')),
        'amount_tax': _as_float(inv.get('amount_tax')),
        'amount_ttc': _as_float(inv.get('amount_total')),
        'amount_residual': _as_float(inv.get('residual')),
        'state': inv.get('state') or 'draft',
        'payment_state': 'paid' if _as_float(inv.get('residual')) <= 0 else 'not_paid',
        'journal': 'N/A',
        'memo': inv.get('reference') or '',
        'move_type': inv.get('type') or '',
    }
    out_lines = []
    for l in lines:
        out_lines.append({
            'description': l.get('name') or '',
            'quantity': _as_float(l.get('quantity') or 1),
            'price_unit': _as_float(l.get('price_unit')),
            'amount_ht': _as_float(l.get('price_subtotal')),
            'amount_ttc': _as_float(l.get('price_total')),
            'tax': '',
        })
    return invoice, out_lines
"""
Finance Service — Odoo data layer for the Finance & Accounting module.
Queries account.move, account.payment, res.partner via XML-RPC.
"""
from collections import defaultdict
from datetime import date

from django.conf import settings


# ─── helpers ────────────────────────────────────────────────────

def _fmt_d(val):
    s = str(val or '')[:10]
    if len(s) < 10:
        return '—'
    return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"


def _avail(uid, models, model):
    return set((models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        model, 'fields_get', [], {'attributes': ['type']},
    ) or {}).keys())


def _clean_doc_number(primary, secondary, row_id):
    for raw in (primary, secondary):
        val = str(raw or '').strip()
        if val and val != '/':
            return val
    return f'FAC-{row_id or "N/A"}'


# ─── invoices (account.move) ────────────────────────────────────

def fetch_invoices(uid, models, move_type='out_invoice', limit=2000):
    av = _avail(uid, models, 'account.move')
    base = ['name', 'invoice_date', 'partner_id', 'amount_untaxed',
            'amount_tax', 'amount_total', 'state', 'move_type']
    opt = ['invoice_date_due', 'payment_state', 'amount_residual',
           'journal_id', 'ref', 'narration', 'invoice_origin',
           'invoice_line_ids', 'currency_id']
    fields = base + [f for f in opt if f in av]
    records = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'account.move', 'search_read',
        [[('move_type', '=', move_type)]],
        {'fields': sorted(set(fields)), 'limit': limit,
         'order': 'invoice_date desc, id desc'},
    )
    today = date.today()
    rows = []
    for r in records:
        d_raw  = str(r.get('invoice_date') or '')[:10]
        due_raw = str(r.get('invoice_date_due') or '')[:10]
        partner = r.get('partner_id')
        journal = r.get('journal_id')
        ht    = round(float(r.get('amount_untaxed') or 0), 2)
        ttc   = round(float(r.get('amount_total')   or 0), 2)
        resid = round(float(r.get('amount_residual') or 0), 2)
        paid  = round(ttc - resid, 2)
        pstate = str(r.get('payment_state') or 'not_paid')
        state  = str(r.get('state') or 'draft')
        is_overdue, days_overdue = False, 0
        if due_raw and len(due_raw) == 10 and pstate != 'paid':
            try:
                diff = (today - date.fromisoformat(due_raw)).days
                if diff > 0:
                    is_overdue, days_overdue = True, diff
            except ValueError:
                pass
        rows.append({
            'id': r.get('id'),
            'name': _clean_doc_number(r.get('name'), r.get('ref') or r.get('invoice_origin'), r.get('id')),
            'date': _fmt_d(d_raw), 'date_iso': d_raw,
            'date_due': _fmt_d(due_raw), 'date_due_iso': due_raw,
            'partner': partner[1] if isinstance(partner, list) and len(partner) > 1 else '—',
            'partner_id': partner[0] if isinstance(partner, list) else None,
            'journal': journal[1] if isinstance(journal, list) and len(journal) > 1 else '—',
            'amount_ht': ht,
            'amount_tax': round(float(r.get('amount_tax') or 0), 2),
            'amount_ttc': ttc,
            'amount_residual': resid,
            'amount_paid': paid,
            'state': state,
            'payment_state': pstate,
            'ref': r.get('ref') or r.get('invoice_origin') or '—',
            'is_overdue': is_overdue,
            'days_overdue': days_overdue,
        })
    return rows


# ─── payments (account.payment) ─────────────────────────────────

def fetch_payments(uid, models, payment_type=None, limit=2000):
    av = _avail(uid, models, 'account.payment')
    base = ['name', 'date', 'partner_id', 'amount']
    opt  = ['payment_type', 'state', 'journal_id', 'ref', 'memo',
            'payment_method_line_id', 'payment_method_id',
            'partner_type', 'currency_id']
    fields = base + [f for f in opt if f in av]
    dom = []
    if payment_type and 'payment_type' in av:
        dom.append(('payment_type', '=', payment_type))
    if 'state' in av:
        dom.append(('state', '=', 'posted'))
    records = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'account.payment', 'search_read', [dom],
        {'fields': sorted(set(fields)), 'limit': limit,
         'order': 'date desc, id desc'},
    )
    _state_map = {'draft': 'Brouillon', 'posted': 'Validé', 'cancel': 'Annulé',
                  'sent': 'Envoyé', 'reconciled': 'Lettré'}
    rows = []
    for r in records:
        d_raw   = str(r.get('date') or '')[:10]
        partner = r.get('partner_id')
        journal = r.get('journal_id')
        pm_raw  = r.get('payment_method_line_id') or r.get('payment_method_id')
        pm_name = pm_raw[1] if isinstance(pm_raw, list) and len(pm_raw) > 1 else '—'
        ptype   = str(r.get('payment_type') or '')
        ptype_l = 'Reçu' if ptype == 'inbound' else ('Envoyé' if ptype == 'outbound' else ptype.capitalize() or '—')
        sraw    = str(r.get('state') or '')
        rows.append({
            'id': r.get('id'),
            'name': r.get('name') or '—',
            'date': _fmt_d(d_raw), 'date_iso': d_raw,
            'partner': partner[1] if isinstance(partner, list) and len(partner) > 1 else '—',
            'partner_id': partner[0] if isinstance(partner, list) else None,
            'amount': round(float(r.get('amount') or 0), 2),
            'payment_type': ptype,
            'type_label': ptype_l,
            'payment_method': pm_name,
            'journal': journal[1] if isinstance(journal, list) and len(journal) > 1 else '—',
            'state': sraw,
            'state_label': _state_map.get(sraw, sraw.capitalize()),
            'ref': r.get('ref') or r.get('memo') or '—',
        })
    return rows


# ─── KPI aggregations ───────────────────────────────────────────

def compute_kpis(inv_clients, inv_fournisseurs, payments):
    today = date.today()
    cur_ym = today.strftime('%Y-%m')

    def _posted(lst): return [r for r in lst if r['state'] == 'posted']

    pc = _posted(inv_clients)
    pf = _posted(inv_fournisseurs)

    ca_total  = sum(r['amount_ht'] for r in pc)
    ca_month  = sum(r['amount_ht'] for r in pc if r['date_iso'][:7] == cur_ym)
    achats    = sum(r['amount_ht'] for r in pf)
    encaiss   = sum(r['amount'] for r in payments if r['payment_type'] == 'inbound')
    decaiss   = sum(r['amount'] for r in payments if r['payment_type'] == 'outbound')
    creances  = sum(r['amount_residual'] for r in pc if r['payment_state'] != 'paid')
    dettes    = sum(r['amount_residual'] for r in pf if r['payment_state'] != 'paid')
    overdue   = [r for r in pc if r['is_overdue']]
    ca_ttc    = sum(r['amount_ttc'] for r in pc)
    ach_ttc   = sum(r['amount_ttc'] for r in pf)

    return {
        'ca_total':   round(ca_total, 0),
        'ca_month':   round(ca_month, 0),
        'achats':     round(achats, 0),
        'encaiss':    round(encaiss, 0),
        'decaiss':    round(decaiss, 0),
        'creances':   round(creances, 0),
        'dettes':     round(dettes, 0),
        'nb_creances': len([r for r in pc if r['payment_state'] != 'paid']),
        'nb_overdue': len(overdue),
        'mt_overdue': round(sum(r['amount_residual'] for r in overdue), 0),
        'liquidite':  round(encaiss * 100.0 / max(1, decaiss), 1) if decaiss else 0.0,
        'marge':      round((ca_total - achats) * 100.0 / max(1, ca_total), 1) if ca_total else 0.0,
        'pct_encaisse': round(encaiss * 100.0 / max(1, ca_ttc), 1) if ca_ttc else 0.0,
        'pct_paye_f': round(decaiss * 100.0 / max(1, ach_ttc), 1) if ach_ttc else 0.0,
        'nb_clients':      len(set(r['partner_id'] for r in inv_clients if r['partner_id'])),
        'nb_fournisseurs': len(set(r['partner_id'] for r in inv_fournisseurs if r['partner_id'])),
        'ticket_moyen':    round(ca_total / max(1, len(pc)), 0) if pc else 0,
        'nb_factures_clients':      len(pc),
        'nb_factures_fournisseurs': len(pf),
    }


# ─── chart data helpers ──────────────────────────────────────────

def _sort_key(mk):
    p = mk.split('/')
    return f"{p[1]}-{p[0]}" if len(p) == 2 else mk


def monthly_amounts(rows, field='amount_ht', n=12):
    acc = defaultdict(float)
    for r in rows:
        d = r.get('date_iso', '')
        if len(d) >= 7:
            acc[f"{d[5:7]}/{d[0:4]}"] += float(r.get(field) or 0)
    return {k: round(v, 0) for k, v in sorted(acc.items(), key=lambda x: _sort_key(x[0]))[-n:]}


def payment_state_breakdown(rows):
    result = defaultdict(float)
    for r in rows:
        if r['state'] != 'posted':
            continue
        ps = r['payment_state']
        if ps == 'paid':
            result['Payées'] += r['amount_ttc']
        elif ps == 'partial':
            result['Partielles'] += r['amount_ttc']
        elif r['is_overdue']:
            result['En retard'] += r['amount_ttc']
        else:
            result['En cours'] += r['amount_ttc']
    return {k: round(v, 0) for k, v in result.items() if v > 0}


def top_partners(rows, n=10, field='amount_ht'):
    acc = defaultdict(float)
    for r in rows:
        if r['state'] == 'posted':
            acc[r['partner']] += float(r.get(field) or 0)
    return {k: round(v, 0) for k, v in sorted(acc.items(), key=lambda x: x[1], reverse=True)[:n]}


def payment_method_breakdown(payments):
    acc = defaultdict(float)
    for r in payments:
        acc[r['payment_method'] or '—'] += r['amount']
    return {k: round(v, 0) for k, v in sorted(acc.items(), key=lambda x: x[1], reverse=True)}


# ─── invoice detail ──────────────────────────────────────────────

def fetch_invoice_detail(uid, models, invoice_id):
    av = _avail(uid, models, 'account.move')
    fields = ['name', 'invoice_date', 'invoice_date_due', 'partner_id',
              'amount_untaxed', 'amount_tax', 'amount_total', 'amount_residual',
              'state', 'move_type', 'payment_state']
    for f in ('journal_id', 'ref', 'narration', 'invoice_line_ids', 'invoice_origin'):
        if f in av:
            fields.append(f)
    records = models.execute_kw(
        settings.ODOO_DB, uid, settings.ODOO_PASS,
        'account.move', 'read', [[invoice_id]],
        {'fields': sorted(set(fields))},
    )
    if not records:
        return None, []
    r = records[0]
    partner = r.get('partner_id')
    journal = r.get('journal_id')
    ttc   = round(float(r.get('amount_total')   or 0), 2)
    resid = round(float(r.get('amount_residual') or 0), 2)
    inv = {
        'id': r.get('id'),
        'name': _clean_doc_number(r.get('name'), r.get('ref') or r.get('invoice_origin'), r.get('id')),
        'date': _fmt_d(str(r.get('invoice_date') or '')[:10]),
        'date_due': _fmt_d(str(r.get('invoice_date_due') or '')[:10]),
        'partner': partner[1] if isinstance(partner, list) and len(partner) > 1 else '—',
        'journal': journal[1] if isinstance(journal, list) and len(journal) > 1 else '—',
        'amount_ht':  round(float(r.get('amount_untaxed') or 0), 2),
        'amount_tax': round(float(r.get('amount_tax')     or 0), 2),
        'amount_ttc': ttc,
        'amount_residual': resid,
        'amount_paid': round(ttc - resid, 2),
        'state': str(r.get('state') or 'draft'),
        'payment_state': str(r.get('payment_state') or 'not_paid'),
        'move_type': str(r.get('move_type') or ''),
        'ref': r.get('ref') or r.get('invoice_origin') or '—',
        'narration': r.get('narration') or '',
    }
    lines = []
    line_ids = r.get('invoice_line_ids') or []
    if line_ids:
        try:
            lrs = models.execute_kw(
                settings.ODOO_DB, uid, settings.ODOO_PASS,
                'account.move.line', 'read', [line_ids[:60]],
                {'fields': ['name', 'quantity', 'price_unit',
                            'price_subtotal', 'price_total', 'product_id']},
            )
            for lr in lrs:
                prod = lr.get('product_id')
                ht_l  = float(lr.get('price_subtotal') or 0)
                ttc_l = float(lr.get('price_total')    or 0)
                lines.append({
                    'description': lr.get('name') or '—',
                    'product': prod[1] if isinstance(prod, list) and len(prod) > 1 else '',
                    'qty': float(lr.get('quantity') or 0),
                    'prix_unit': round(float(lr.get('price_unit') or 0), 2),
                    'montant_ht': round(ht_l, 2),
                    'montant_ttc': round(ttc_l, 2),
                    'tax_pct': round((ttc_l - ht_l) / max(0.01, ht_l) * 100, 1) if ht_l > 0 else 0,
                })
        except Exception:
            pass
    return inv, lines
