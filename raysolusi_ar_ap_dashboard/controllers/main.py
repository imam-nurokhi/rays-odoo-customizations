import calendar
import io
import json
import logging
from datetime import date
from dateutil.relativedelta import relativedelta

from odoo import http
from odoo.http import request

_logger = logging.getLogger(__name__)


class ArApDashboardController(http.Controller):

    @http.route('/raysolusi/ar_ap_dashboard/data', type='json', auth='user')
    def get_dashboard_data(self, **kwargs):
        """Return aggregated AR & AP data for the dashboard."""
        env = request.env
        today = date.today()

        # ── Optional filters ───────────────────────────────────────────────
        date_from    = kwargs.get('date_from')    # str YYYY-MM-DD or None
        date_to      = kwargs.get('date_to')      # str YYYY-MM-DD or None
        partner_name = kwargs.get('partner_name') # str or None

        # ── Domains ────────────────────────────────────────────────────────
        OUTSTANDING = ['not_paid', 'in_payment', 'partial']

        def build_outstanding_domain(move_type):
            domain = [
                ('move_type', '=', move_type),
                ('state', '=', 'posted'),
                ('payment_state', 'in', OUTSTANDING),
            ]
            if date_from:
                domain.append(('invoice_date', '>=', date_from))
            if date_to:
                domain.append(('invoice_date', '<=', date_to))
            if partner_name:
                domain.append(('partner_id.name', 'ilike', partner_name))
            return domain

        ar_domain = build_outstanding_domain('out_invoice')
        ap_domain = build_outstanding_domain('in_invoice')

        ar_recs = env['account.move'].sudo().search(ar_domain)
        ap_recs = env['account.move'].sudo().search(ap_domain)

        # ── Summary ────────────────────────────────────────────────────────
        total_ar = sum(ar_recs.mapped('amount_residual'))
        total_ap = sum(ap_recs.mapped('amount_residual'))

        overdue_ar_recs = ar_recs.filtered(
            lambda r: r.invoice_date_due and r.invoice_date_due < today)
        overdue_ap_recs = ap_recs.filtered(
            lambda r: r.invoice_date_due and r.invoice_date_due < today)

        overdue_ar = sum(overdue_ar_recs.mapped('amount_residual'))
        overdue_ap = sum(overdue_ap_recs.mapped('amount_residual'))

        # ── Aging Buckets ──────────────────────────────────────────────────
        def aging_buckets(records):
            keys = ['current', 'b1_30', 'b31_60', 'b61_90', 'over90']
            amounts = {k: 0.0 for k in keys}
            counts  = {k: 0   for k in keys}
            for r in records:
                amt = r.amount_residual
                if not r.invoice_date_due:
                    amounts['current'] += amt
                    counts['current']  += 1
                    continue
                days = (today - r.invoice_date_due).days
                if days <= 0:
                    k = 'current'
                elif days <= 30:
                    k = 'b1_30'
                elif days <= 60:
                    k = 'b31_60'
                elif days <= 90:
                    k = 'b61_90'
                else:
                    k = 'over90'
                amounts[k] += amt
                counts[k]  += 1
            return {
                'labels':  ['Belum Jatuh Tempo', '1-30 Hari', '31-60 Hari', '61-90 Hari', '>90 Hari'],
                'amounts': [round(amounts[k]) for k in keys],
                'counts':  [counts[k]         for k in keys],
            }

        # ── Top 10 Partners ────────────────────────────────────────────────
        def top_partners(records, n=10):
            by_partner = {}
            for r in records:
                if not r.partner_id:
                    continue
                pname = r.partner_id.name or 'Unknown'
                entry = by_partner.setdefault(pname, {'amount': 0.0, 'count': 0})
                entry['amount'] += r.amount_residual
                entry['count']  += 1
            sorted_p = sorted(by_partner.items(), key=lambda x: -x[1]['amount'])[:n]
            return [
                {'partner': p, 'amount': round(d['amount']), 'count': d['count']}
                for p, d in sorted_p
            ]

        # ── AR/AP Category Breakdown (PREMI vs DISKON / BROKERAGE) ─────────
        def ar_category_breakdown():
            AML = env['account.move.line']
            aml_domain = [
                ('move_id.state', '=', 'posted'),
                ('account_id.account_type', '=', 'asset_receivable'),
                ('move_id.move_type', 'in', ['out_invoice', 'out_refund']),
            ]
            if date_from:
                aml_domain.append(('date', '>=', date_from))
            if date_to:
                aml_domain.append(('date', '<=', date_to))
            if partner_name:
                aml_domain.append(('partner_id.name', 'ilike', partner_name))
            lines = AML.sudo().search(aml_domain)
            premi  = sum(l.balance for l in lines if l.move_id.move_type == 'out_invoice')
            diskon = abs(sum(l.balance for l in lines if l.move_id.move_type == 'out_refund'))
            return {'premi': round(premi), 'diskon': round(diskon)}

        def ap_category_breakdown():
            AML = env['account.move.line']
            aml_domain = [
                ('move_id.state', '=', 'posted'),
                ('account_id.account_type', '=', 'liability_payable'),
                ('move_id.move_type', 'in', ['in_invoice', 'in_refund']),
            ]
            if date_from:
                aml_domain.append(('date', '>=', date_from))
            if date_to:
                aml_domain.append(('date', '<=', date_to))
            if partner_name:
                aml_domain.append(('partner_id.name', 'ilike', partner_name))
            ap_lines = AML.sudo().search(aml_domain)
            ap_move_ids = ap_lines.mapped('move_id').ids
            brok_lines = AML.sudo().search([
                ('move_id', 'in', ap_move_ids),
                ('move_id.state', '=', 'posted'),
                ('account_id.account_type', '=', 'income'),
                ('product_id.name', 'ilike', 'brokerage'),
            ]) if ap_move_ids else AML.sudo().browse([])
            premi     = round(abs(sum(l.balance for l in ap_lines if l.move_id.move_type == 'in_invoice')))
            brokerage = round(abs(sum(l.balance for l in brok_lines)))
            return {'premi': premi, 'brokerage': brokerage}

                # ── Monthly Trend (last 12 months) — NOT affected by date filter ───
        month_names_id = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
                          'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des']

        months = []
        base = today.replace(day=1)
        for i in range(11, -1, -1):
            m = base - relativedelta(months=i)
            months.append((m.year, m.month))

        monthly_labels = []
        monthly_ar     = []
        monthly_ap     = []

        for year, month in months:
            last_day = calendar.monthrange(year, month)[1]
            d_from   = '%s-%02d-01' % (year, month)
            d_to     = '%s-%02d-%02d' % (year, month, last_day)
            monthly_labels.append('%s %s' % (month_names_id[month - 1], year))

            base_dom = [
                ('state', '=', 'posted'),
                ('date', '>=', d_from),
                ('date', '<=', d_to),
            ]
            ar_m = env['account.move'].sudo().search_read(
                base_dom + [('move_type', '=', 'out_invoice')],
                ['amount_total'],
            )
            ap_m = env['account.move'].sudo().search_read(
                base_dom + [('move_type', '=', 'in_invoice')],
                ['amount_total'],
            )
            monthly_ar.append(round(sum(r['amount_total'] for r in ar_m)))
            monthly_ap.append(round(sum(r['amount_total'] for r in ap_m)))

        # ── Payment Status (last 12 months) — NOT affected by date filter ──
        twelve_ago  = base - relativedelta(months=11)
        date_from12 = '%s-%02d-01' % (twelve_ago.year, twelve_ago.month)

        def payment_status_counts(move_type):
            recs = env['account.move'].sudo().search_read(
                [
                    ('move_type', '=', move_type),
                    ('state', '=', 'posted'),
                    ('date', '>=', date_from12),
                ],
                ['payment_state'],
            )
            d = {'paid': 0, 'not_paid': 0, 'in_payment': 0, 'partial': 0, 'reversed': 0}
            for r in recs:
                ps = r.get('payment_state') or 'not_paid'
                d[ps] = d.get(ps, 0) + 1
            return d

        # ── Top Overdue Partners ───────────────────────────────────────────
        def top_overdue(records, n=5):
            by_partner = {}
            for r in records:
                if not r.partner_id or not r.invoice_date_due:
                    continue
                if r.invoice_date_due >= today:
                    continue
                pname    = r.partner_id.name or 'Unknown'
                days_ov  = (today - r.invoice_date_due).days
                entry    = by_partner.setdefault(pname, {'amount': 0.0, 'max_days': 0, 'count': 0})
                entry['amount']   += r.amount_residual
                entry['max_days']  = max(entry['max_days'], days_ov)
                entry['count']    += 1
            sorted_p = sorted(by_partner.items(), key=lambda x: -x[1]['amount'])[:n]
            return [
                {
                    'partner':  p,
                    'amount':   round(d['amount']),
                    'max_days': d['max_days'],
                    'count':    d['count'],
                }
                for p, d in sorted_p
            ]

        return {
            'summary': {
                'total_ar':         round(total_ar),
                'total_ap':         round(total_ap),
                'overdue_ar':       round(overdue_ar),
                'overdue_ap':       round(overdue_ap),
                'ar_count':         len(ar_recs),
                'ap_count':         len(ap_recs),
                'overdue_ar_count': len(overdue_ar_recs),
                'overdue_ap_count': len(overdue_ap_recs),
                'net_position':     round(total_ar - total_ap),
            },
            'ar_aging':        aging_buckets(ar_recs),
            'ap_aging':        aging_buckets(ap_recs),
            'top_ar':          top_partners(ar_recs),
            'top_ap':          top_partners(ap_recs),
            'top_overdue_ar':  top_overdue(ar_recs),
            'top_overdue_ap':  top_overdue(ap_recs),
            'ar_by_category':  ar_category_breakdown(),
            'ap_by_category':  ap_category_breakdown(),
            'top_customers_ar': [[r['partner'], r['amount']] for r in top_partners(ar_recs)],
            'top_insurance_ap': [[r['partner'], r['amount']] for r in top_partners(ap_recs)],
            'monthly': {
                'labels': monthly_labels,
                'ar':     monthly_ar,
                'ap':     monthly_ap,
            },
            'payment_status': {
                'ar': payment_status_counts('out_invoice'),
                'ap': payment_status_counts('in_invoice'),
            },
        }

    # ──────────────────────────────────────────────────────────────────────────
    # Endpoint 2: Detail drill-down
    # ──────────────────────────────────────────────────────────────────────────

    @http.route('/raysolusi/ar_ap_dashboard/detail', type='json', auth='user')
    def get_dashboard_detail(self, **kwargs):
        """Return individual invoice records for a clicked dashboard segment."""
        _EMPTY = {'records': [], 'total_count': 0, 'total_amount': 0, 'title': ''}

        try:
            env   = request.env
            today = date.today()

            OUTSTANDING = ['not_paid', 'in_payment', 'partial']

            # ── Input parameters ───────────────────────────────────────────
            move_type_raw = (kwargs.get('move_type') or '').strip().lower()   # 'ar' or 'ap'
            filter_type   = (kwargs.get('filter_type') or '').strip().lower() # 'aging'|'month'|'status'|'partner'|'overdue'
            filter_value  = (kwargs.get('filter_value') or '').strip()        # bucket / month / state / name
            date_from     = (kwargs.get('date_from') or '').strip() or None   # YYYY-MM-DD
            date_to       = (kwargs.get('date_to') or '').strip() or None     # YYYY-MM-DD
            partner_name  = (kwargs.get('partner_name') or '').strip() or None

            if move_type_raw not in ('ar', 'ap'):
                return _EMPTY

            move_type_odoo = 'out_invoice' if move_type_raw == 'ar' else 'in_invoice'
            move_type_label = 'AR' if move_type_raw == 'ar' else 'AP'

            # ── Month names (Indonesian) ───────────────────────────────────
            month_names_id = {
                1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
                5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
                9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember',
            }

            # ── Base domain ────────────────────────────────────────────────
            domain = [
                ('move_type', '=', move_type_odoo),
                ('state', '=', 'posted'),
            ]

            # Apply optional date_from / date_to on invoice_date
            if date_from:
                domain.append(('invoice_date', '>=', date_from))
            if date_to:
                domain.append(('invoice_date', '<=', date_to))

            # Apply optional partner_name ilike
            if partner_name:
                domain.append(('partner_id.name', 'ilike', partner_name))

            # ── Filter-type-specific domain additions ──────────────────────
            title = ''
            post_filter = None   # optional Python-level filter applied after search

            if filter_type == 'aging':
                domain.append(('payment_state', 'in', OUTSTANDING))

                bucket_labels = {
                    'current': 'Belum Jatuh Tempo',
                    '1-30':    '1-30 Hari',
                    '31-60':   '31-60 Hari',
                    '61-90':   '61-90 Hari',
                    '>90':     '>90 Hari',
                }
                bucket_label = bucket_labels.get(filter_value, filter_value)
                title = '%s \u2014 Aging %s' % (move_type_label, bucket_label)

                # Define Python-level bucket filter
                def _aging_filter(move, fv=filter_value, td=today):
                    if not move.invoice_date_due:
                        return fv == 'current'
                    days = (td - move.invoice_date_due).days
                    if fv == 'current':
                        return days <= 0
                    elif fv == '1-30':
                        return 1 <= days <= 30
                    elif fv == '31-60':
                        return 31 <= days <= 60
                    elif fv == '61-90':
                        return 61 <= days <= 90
                    elif fv == '>90':
                        return days > 90
                    return False

                post_filter = _aging_filter

            elif filter_type == 'month':
                # filter_value is 'YYYY-MM'
                try:
                    year_s, month_s = filter_value.split('-')
                    year_i  = int(year_s)
                    month_i = int(month_s)
                    last_day = calendar.monthrange(year_i, month_i)[1]
                    first_day_str = '%04d-%02d-01' % (year_i, month_i)
                    last_day_str  = '%04d-%02d-%02d' % (year_i, month_i, last_day)
                    domain.append(('date', '>=', first_day_str))
                    domain.append(('date', '<=', last_day_str))
                    month_label = '%s %s' % (month_names_id.get(month_i, month_s), year_s)
                    title = '%s \u2014 %s' % (move_type_label, month_label)
                except Exception:
                    title = '%s \u2014 Bulan %s' % (move_type_label, filter_value)

            elif filter_type == 'status':
                domain.append(('payment_state', '=', filter_value))

                # Restrict to last 12 months (same as payment_status in /data)
                base        = today.replace(day=1)
                twelve_ago  = base - relativedelta(months=11)
                date_from12 = '%s-%02d-01' % (twelve_ago.year, twelve_ago.month)
                domain.append(('date', '>=', date_from12))

                status_labels = {
                    'paid':       'Lunas',
                    'not_paid':   'Belum Bayar',
                    'in_payment': 'Dalam Pembayaran',
                    'partial':    'Bayar Sebagian',
                    'reversed':   'Dibatalkan',
                }
                status_label = status_labels.get(filter_value, filter_value)
                title = '%s \u2014 Status: %s' % (move_type_label, status_label)

            elif filter_type == 'partner':
                domain.append(('partner_id.name', 'ilike', filter_value))
                domain.append(('payment_state', 'in', OUTSTANDING))
                title = '%s \u2014 Partner: %s' % (move_type_label, filter_value)

            elif filter_type == 'overdue':
                domain.append(('payment_state', 'in', OUTSTANDING))
                domain.append(('invoice_date_due', '<', str(today)))
                title = '%s \u2014 Jatuh Tempo' % move_type_label

            else:
                # Unknown filter_type — return empty
                return _EMPTY

            # ── Search ─────────────────────────────────────────────────────
            records = env['account.move'].sudo().search(domain)

            # Apply Python-level post-filter if needed (aging buckets)
            if post_filter is not None:
                records = records.filtered(post_filter)

            # ── Build return rows (max 200) ────────────────────────────────
            rows = []
            for move in records[:200]:
                if move.invoice_date_due and move.invoice_date_due < today:
                    days_overdue = (today - move.invoice_date_due).days
                else:
                    days_overdue = 0

                rows.append({
                    'id':             move.id,
                    'name':           move.name or '',
                    'partner':        move.partner_id.name or '' if move.partner_id else '',
                    'date':           str(move.invoice_date) if move.invoice_date else '',
                    'due_date':       str(move.invoice_date_due) if move.invoice_date_due else '',
                    'amount_total':   round(move.amount_total),
                    'amount_residual': round(move.amount_residual),
                    'payment_state':  move.payment_state or '',
                    'days_overdue':   days_overdue,
                    'currency':       move.currency_id.symbol or 'Rp' if move.currency_id else 'Rp',
                })

            total_amount = round(sum(records.mapped('amount_residual')))

            return {
                'records':      rows,
                'total_count':  len(records),
                'total_amount': total_amount,
                'title':        title,
            }

        except Exception as exc:
            _logger.exception('ArApDashboard detail error: %s', exc)
            return {'records': [], 'total_count': 0, 'total_amount': 0, 'title': ''}

# ──────────────────────────────────────────────────────────────────────────────
# AR Report endpoints
# ──────────────────────────────────────────────────────────────────────────────

    @http.route('/raysolusi/ar_report/data', type='json', auth='user')
    def get_ar_report_data(self, date_from=None, date_to=None, category='all',
                           partner_name=None, payment_state='all', **kwargs):
        try:
            env = request.env
            today = date.today()

            domain = [('state', '=', 'posted')]
            if category == 'premium':
                domain.append(('move_type', '=', 'out_invoice'))
            elif category == 'discount':
                domain.append(('move_type', '=', 'out_refund'))
            else:
                domain.append(('move_type', 'in', ['out_invoice', 'out_refund']))

            if date_from:
                domain.append(('invoice_date', '>=', date_from))
            if date_to:
                domain.append(('invoice_date', '<=', date_to))
            if partner_name:
                domain.append(('partner_id.name', 'ilike', partner_name))
            if payment_state and payment_state != 'all':
                domain.append(('payment_state', '=', payment_state))

            moves = env['account.move'].sudo().search(domain, order='invoice_date desc, name desc', limit=500)

            premi_moves = moves.filtered(lambda m: m.move_type == 'out_invoice')
            diskon_moves = moves.filtered(lambda m: m.move_type == 'out_refund')

            total_premi = sum(premi_moves.mapped('amount_total'))
            total_diskon = sum(diskon_moves.mapped('amount_total'))

            records = []
            for m in moves[:200]:
                records.append({
                    'id': m.id,
                    'name': m.name or '',
                    'partner': m.partner_id.name or '' if m.partner_id else '',
                    'date': str(m.invoice_date) if m.invoice_date else '',
                    'due_date': str(m.invoice_date_due) if m.invoice_date_due else '',
                    'amount_total': round(m.amount_total),
                    'amount_residual': round(m.amount_residual),
                    'payment_state': m.payment_state or '',
                    'category': 'DISKON' if m.move_type == 'out_refund' else 'PREMI',
                    'currency': m.currency_id.symbol or 'Rp' if m.currency_id else 'Rp',
                })

            # Monthly chart (last 12 months)
            month_names_id = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agt','Sep','Okt','Nov','Des']
            base = today.replace(day=1)
            chart_labels = []
            chart_premi = []
            chart_diskon = []
            for i in range(11, -1, -1):
                m_date = base - relativedelta(months=i)
                last_day = calendar.monthrange(m_date.year, m_date.month)[1]
                d_from = '%s-%02d-01' % (m_date.year, m_date.month)
                d_to   = '%s-%02d-%02d' % (m_date.year, m_date.month, last_day)
                chart_labels.append('%s %s' % (month_names_id[m_date.month - 1], m_date.year))
                p_recs = env['account.move'].sudo().search_read(
                    [('move_type','=','out_invoice'),('state','=','posted'),
                     ('invoice_date','>=',d_from),('invoice_date','<=',d_to)],
                    ['amount_total'])
                d_recs = env['account.move'].sudo().search_read(
                    [('move_type','=','out_refund'),('state','=','posted'),
                     ('invoice_date','>=',d_from),('invoice_date','<=',d_to)],
                    ['amount_total'])
                chart_premi.append(round(sum(r['amount_total'] for r in p_recs)))
                chart_diskon.append(round(sum(r['amount_total'] for r in d_recs)))

            return {
                'records': records,
                'total_count': len(moves),
                'summary': {
                    'total_premi': round(total_premi),
                    'total_diskon': round(total_diskon),
                    'total_net': round(total_premi - total_diskon),
                    'count_premi': len(premi_moves),
                    'count_diskon': len(diskon_moves),
                },
                'chart': {
                    'labels': chart_labels,
                    'premi': chart_premi,
                    'diskon': chart_diskon,
                },
            }
        except Exception as exc:
            _logger.exception('AR report data error: %s', exc)
            return {'records': [], 'total_count': 0,
                    'summary': {'total_premi':0,'total_diskon':0,'total_net':0,'count_premi':0,'count_diskon':0},
                    'chart': {'labels':[],'premi':[],'diskon':[]}}

    @http.route('/raysolusi/ar_report/excel', type='http', auth='user', methods=['POST'], csrf=False)
    def export_ar_excel(self, **post):
        import io as _io
        import json as _json
        try:
            import xlsxwriter
        except ImportError:
            return request.make_response('xlsxwriter not installed', [('Content-Type','text/plain')])

        env = request.env
        params = {}
        raw = post.get('params') or post.get('data') or ''
        if raw:
            try:
                params = _json.loads(raw)
            except Exception:
                pass

        date_from = params.get('date_from') or post.get('date_from')
        date_to   = params.get('date_to')   or post.get('date_to')
        category  = params.get('category')  or post.get('category', 'all')
        partner_name  = params.get('partner_name') or post.get('partner_name')
        payment_state = params.get('payment_state') or post.get('payment_state', 'all')

        domain = [('state','=','posted')]
        if category == 'premium':
            domain.append(('move_type','=','out_invoice'))
        elif category == 'discount':
            domain.append(('move_type','=','out_refund'))
        else:
            domain.append(('move_type','in',['out_invoice','out_refund']))
        if date_from:
            domain.append(('invoice_date','>=', date_from))
        if date_to:
            domain.append(('invoice_date','<=', date_to))
        if partner_name:
            domain.append(('partner_id.name','ilike', partner_name))
        if payment_state and payment_state != 'all':
            domain.append(('payment_state','=', payment_state))

        moves = env['account.move'].sudo().search(domain, order='invoice_date asc, name asc')

        output = _io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('AR Report')

        title_fmt = wb.add_format({'bold':True,'font_size':13,'bg_color':'#1F4E79','font_color':'white','align':'center','valign':'vcenter'})
        hdr_fmt   = wb.add_format({'bold':True,'bg_color':'#2E75B6','font_color':'white','border':1,'align':'center','text_wrap':True})
        txt_fmt   = wb.add_format({'border':1})
        num_fmt   = wb.add_format({'border':1,'num_format':'#,##0'})
        date_fmt  = wb.add_format({'border':1,'num_format':'dd/mm/yyyy'})
        total_fmt = wb.add_format({'bold':True,'bg_color':'#1F4E79','font_color':'white','num_format':'#,##0','border':1})
        total_lbl = wb.add_format({'bold':True,'bg_color':'#1F4E79','font_color':'white','border':1})

        cols = [('No',4),('Tanggal',12),('No Invoice',20),('Partner/Customer',35),
                ('Jatuh Tempo',12),('Kategori',10),('Total',18),('Sisa Tagihan',18),('Status',14),('Mata Uang',10)]
        for c,(h,w) in enumerate(cols):
            ws.set_column(c,c,w)
        ws.set_row(0,28)
        ws.set_row(1,18)
        ws.merge_range(0,0,0,len(cols)-1,'PT RAYSOLUSI - LAPORAN AR (PIUTANG)', title_fmt)
        period_str = 'Periode: %s s/d %s' % (date_from or '-', date_to or '-')
        ws.merge_range(1,0,1,len(cols)-1, period_str, wb.add_format({'align':'center','italic':True}))
        for c,(h,_) in enumerate(cols):
            ws.write(2,c,h,hdr_fmt)

        row = 3
        for idx, m in enumerate(moves, 1):
            cat = 'DISKON' if m.move_type == 'out_refund' else 'PREMI'
            ps_map = {'paid':'Lunas','not_paid':'Belum Bayar','in_payment':'Dalam Proses',
                      'partial':'Sebagian','reversed':'Dibatalkan'}
            ps = ps_map.get(m.payment_state, m.payment_state or '')
            ws.write(row,0,idx,txt_fmt)
            if m.invoice_date:
                ws.write_datetime(row,1,m.invoice_date,date_fmt)
            else:
                ws.write(row,1,'',txt_fmt)
            ws.write(row,2,m.name or '',txt_fmt)
            ws.write(row,3,m.partner_id.name or '' if m.partner_id else '',txt_fmt)
            if m.invoice_date_due:
                ws.write_datetime(row,4,m.invoice_date_due,date_fmt)
            else:
                ws.write(row,4,'',txt_fmt)
            ws.write(row,5,cat,txt_fmt)
            ws.write(row,6,m.amount_total,num_fmt)
            ws.write(row,7,m.amount_residual,num_fmt)
            ws.write(row,8,ps,txt_fmt)
            ws.write(row,9,m.currency_id.name or 'IDR' if m.currency_id else 'IDR',txt_fmt)
            row += 1

        premi_moves  = moves.filtered(lambda x: x.move_type == 'out_invoice')
        diskon_moves = moves.filtered(lambda x: x.move_type == 'out_refund')
        ws.merge_range(row,0,row,5,'TOTAL',total_lbl)
        ws.write(row,6,sum(premi_moves.mapped('amount_total')),total_fmt)
        ws.write(row,7,sum(premi_moves.mapped('amount_residual')),total_fmt)
        ws.write(row,8,'',total_lbl)
        ws.write(row,9,'',total_lbl)
        row += 1
        ws.merge_range(row,0,row,5,'TOTAL DISKON/CN',total_lbl)
        ws.write(row,6,sum(diskon_moves.mapped('amount_total')),total_fmt)
        ws.write(row,7,sum(diskon_moves.mapped('amount_residual')),total_fmt)
        ws.write(row,8,'',total_lbl)
        ws.write(row,9,'',total_lbl)

        wb.close()
        xlsx_data = output.getvalue()
        filename = 'ar_report_%s.xlsx' % date.today().strftime('%Y%m%d')
        headers = [
            ('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition','attachment; filename="%s"' % filename),
            ('Content-Length', str(len(xlsx_data))),
        ]
        return request.make_response(xlsx_data, headers)

# ──────────────────────────────────────────────────────────────────────────────
# AP Report endpoints
# ──────────────────────────────────────────────────────────────────────────────

    @http.route('/raysolusi/ap_report/data', type='json', auth='user')
    def get_ap_report_data(self, date_from=None, date_to=None, category='all',
                           partner_name=None, payment_state='all', **kwargs):
        try:
            env = request.env
            today = date.today()

            domain = [('state','=','posted')]
            if category == 'premi':
                domain.append(('move_type','=','in_invoice'))
            elif category == 'brokerage':
                domain.append(('move_type','=','in_invoice'))
            else:
                domain.append(('move_type','in',['in_invoice','in_refund']))

            if date_from:
                domain.append(('invoice_date','>=', date_from))
            if date_to:
                domain.append(('invoice_date','<=', date_to))
            if partner_name:
                domain.append(('partner_id.name','ilike', partner_name))
            if payment_state and payment_state != 'all':
                domain.append(('payment_state','=', payment_state))

            moves = env['account.move'].sudo().search(domain, order='invoice_date desc, name desc', limit=500)

            premi_moves    = moves.filtered(lambda m: m.move_type == 'in_invoice')
            brokerage_moves = moves.filtered(lambda m: m.move_type == 'in_refund')

            total_premi     = sum(premi_moves.mapped('amount_total'))
            total_brokerage = sum(brokerage_moves.mapped('amount_total'))

            # Brokerage from in_invoice lines
            all_invoice_ids = premi_moves.ids
            brok_lines = env['account.move.line'].sudo().search([
                ('move_id','in', all_invoice_ids),
                ('move_id.state','=','posted'),
                ('product_id.name','ilike','brokerage'),
            ]) if all_invoice_ids else env['account.move.line'].sudo().browse([])
            total_brokerage_lines = abs(sum(brok_lines.mapped('balance')))

            records = []
            for m in moves[:200]:
                cat = 'REFUND' if m.move_type == 'in_refund' else 'PREMI'
                records.append({
                    'id': m.id,
                    'name': m.name or '',
                    'partner': m.partner_id.name or '' if m.partner_id else '',
                    'date': str(m.invoice_date) if m.invoice_date else '',
                    'due_date': str(m.invoice_date_due) if m.invoice_date_due else '',
                    'amount_total': round(m.amount_total),
                    'amount_residual': round(m.amount_residual),
                    'payment_state': m.payment_state or '',
                    'category': cat,
                    'currency': m.currency_id.symbol or 'Rp' if m.currency_id else 'Rp',
                })

            month_names_id = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agt','Sep','Okt','Nov','Des']
            base = today.replace(day=1)
            chart_labels = []
            chart_premi = []
            chart_brokerage = []
            for i in range(11, -1, -1):
                m_date = base - relativedelta(months=i)
                last_day = calendar.monthrange(m_date.year, m_date.month)[1]
                d_from = '%s-%02d-01' % (m_date.year, m_date.month)
                d_to   = '%s-%02d-%02d' % (m_date.year, m_date.month, last_day)
                chart_labels.append('%s %s' % (month_names_id[m_date.month-1], m_date.year))
                p_recs = env['account.move'].sudo().search_read(
                    [('move_type','=','in_invoice'),('state','=','posted'),
                     ('invoice_date','>=',d_from),('invoice_date','<=',d_to)],
                    ['amount_total'])
                b_lines = env['account.move.line'].sudo().search_read(
                    [('move_id.move_type','=','in_invoice'),('move_id.state','=','posted'),
                     ('move_id.invoice_date','>=',d_from),('move_id.invoice_date','<=',d_to),
                     ('product_id.name','ilike','brokerage')],
                    ['balance'])
                chart_premi.append(round(sum(r['amount_total'] for r in p_recs)))
                chart_brokerage.append(round(abs(sum(r['balance'] for r in b_lines))))

            return {
                'records': records,
                'total_count': len(moves),
                'summary': {
                    'total_premi': round(total_premi),
                    'total_brokerage': round(total_brokerage_lines),
                    'total_net': round(total_premi - total_brokerage_lines),
                    'count_premi': len(premi_moves),
                    'count_brokerage': len(brokerage_moves),
                },
                'chart': {
                    'labels': chart_labels,
                    'premi': chart_premi,
                    'brokerage': chart_brokerage,
                },
            }
        except Exception as exc:
            _logger.exception('AP report data error: %s', exc)
            return {'records':[],'total_count':0,
                    'summary':{'total_premi':0,'total_brokerage':0,'total_net':0,'count_premi':0,'count_brokerage':0},
                    'chart':{'labels':[],'premi':[],'brokerage':[]}}

    @http.route('/raysolusi/ap_report/excel', type='http', auth='user', methods=['POST'], csrf=False)
    def export_ap_excel(self, **post):
        import io as _io
        import json as _json
        try:
            import xlsxwriter
        except ImportError:
            return request.make_response('xlsxwriter not installed', [('Content-Type','text/plain')])

        env = request.env
        params = {}
        raw = post.get('params') or post.get('data') or ''
        if raw:
            try:
                params = _json.loads(raw)
            except Exception:
                pass

        date_from = params.get('date_from') or post.get('date_from')
        date_to   = params.get('date_to')   or post.get('date_to')
        category  = params.get('category')  or post.get('category', 'all')
        partner_name  = params.get('partner_name') or post.get('partner_name')
        payment_state = params.get('payment_state') or post.get('payment_state', 'all')

        domain = [('state','=','posted')]
        if category == 'premi':
            domain.append(('move_type','=','in_invoice'))
        else:
            domain.append(('move_type','in',['in_invoice','in_refund']))
        if date_from:
            domain.append(('invoice_date','>=', date_from))
        if date_to:
            domain.append(('invoice_date','<=', date_to))
        if partner_name:
            domain.append(('partner_id.name','ilike', partner_name))
        if payment_state and payment_state != 'all':
            domain.append(('payment_state','=', payment_state))

        moves = env['account.move'].sudo().search(domain, order='invoice_date asc, name asc')

        output = _io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('AP Report')

        title_fmt = wb.add_format({'bold':True,'font_size':13,'bg_color':'#1F4E79','font_color':'white','align':'center','valign':'vcenter'})
        hdr_fmt   = wb.add_format({'bold':True,'bg_color':'#375623','font_color':'white','border':1,'align':'center','text_wrap':True})
        txt_fmt   = wb.add_format({'border':1})
        num_fmt   = wb.add_format({'border':1,'num_format':'#,##0'})
        date_fmt  = wb.add_format({'border':1,'num_format':'dd/mm/yyyy'})
        total_fmt = wb.add_format({'bold':True,'bg_color':'#375623','font_color':'white','num_format':'#,##0','border':1})
        total_lbl = wb.add_format({'bold':True,'bg_color':'#375623','font_color':'white','border':1})

        cols = [('No',4),('Tanggal',12),('No Bill',20),('Vendor/Asuransi',35),
                ('Jatuh Tempo',12),('Kategori',10),('Total',18),('Sisa Hutang',18),('Status',14),('Mata Uang',10)]
        for c,(h,w) in enumerate(cols):
            ws.set_column(c,c,w)
        ws.set_row(0,28)
        ws.set_row(1,18)
        ws.merge_range(0,0,0,len(cols)-1,'PT RAYSOLUSI - LAPORAN AP (HUTANG)', title_fmt)
        period_str = 'Periode: %s s/d %s' % (date_from or '-', date_to or '-')
        ws.merge_range(1,0,1,len(cols)-1, period_str, wb.add_format({'align':'center','italic':True}))
        for c,(h,_) in enumerate(cols):
            ws.write(2,c,h,hdr_fmt)

        row = 3
        for idx, m in enumerate(moves, 1):
            cat = 'REFUND' if m.move_type == 'in_refund' else 'PREMI'
            ps_map = {'paid':'Lunas','not_paid':'Belum Bayar','in_payment':'Dalam Proses',
                      'partial':'Sebagian','reversed':'Dibatalkan'}
            ps = ps_map.get(m.payment_state, m.payment_state or '')
            ws.write(row,0,idx,txt_fmt)
            if m.invoice_date:
                ws.write_datetime(row,1,m.invoice_date,date_fmt)
            else:
                ws.write(row,1,'',txt_fmt)
            ws.write(row,2,m.name or '',txt_fmt)
            ws.write(row,3,m.partner_id.name or '' if m.partner_id else '',txt_fmt)
            if m.invoice_date_due:
                ws.write_datetime(row,4,m.invoice_date_due,date_fmt)
            else:
                ws.write(row,4,'',txt_fmt)
            ws.write(row,5,cat,txt_fmt)
            ws.write(row,6,m.amount_total,num_fmt)
            ws.write(row,7,m.amount_residual,num_fmt)
            ws.write(row,8,ps,txt_fmt)
            ws.write(row,9,m.currency_id.name or 'IDR' if m.currency_id else 'IDR',txt_fmt)
            row += 1

        premi_m = moves.filtered(lambda x: x.move_type == 'in_invoice')
        refund_m = moves.filtered(lambda x: x.move_type == 'in_refund')
        ws.merge_range(row,0,row,5,'TOTAL PREMI',total_lbl)
        ws.write(row,6,sum(premi_m.mapped('amount_total')),total_fmt)
        ws.write(row,7,sum(premi_m.mapped('amount_residual')),total_fmt)
        ws.write(row,8,'',total_lbl); ws.write(row,9,'',total_lbl)
        row += 1
        ws.merge_range(row,0,row,5,'TOTAL REFUND',total_lbl)
        ws.write(row,6,sum(refund_m.mapped('amount_total')),total_fmt)
        ws.write(row,7,sum(refund_m.mapped('amount_residual')),total_fmt)
        ws.write(row,8,'',total_lbl); ws.write(row,9,'',total_lbl)

        wb.close()
        xlsx_data = output.getvalue()
        filename = 'ap_report_%s.xlsx' % date.today().strftime('%Y%m%d')
        headers = [
            ('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition','attachment; filename="%s"' % filename),
            ('Content-Length', str(len(xlsx_data))),
        ]
        return request.make_response(xlsx_data, headers)

# ──────────────────────────────────────────────────────────────────────────────
# Production Report endpoints
# ──────────────────────────────────────────────────────────────────────────────

    @http.route('/raysolusi/production_report/data', type='json', auth='user')
    def get_production_report_data(self, date_from=None, date_to=None,
                                   partner_name=None, asuransi_name=None,
                                   move_type='all', **kwargs):
        try:
            env = request.env
            today = date.today()
            has_policy = 'policy_id' in env['account.move']._fields

            PREMI_KW     = ['premi','premium','insurance premium','hutang kepada insurance']
            DISKON_KW    = ['diskon','discount','insurance discount','brokerage discount']
            BROKERAGE_KW = ['brokerage fee','brokerage','komisi','commission','broker fee']

            def classify_line(name_str):
                if not name_str:
                    return 'other'
                n = name_str.lower()
                for kw in BROKERAGE_KW:
                    if kw in n: return 'brokerage'
                for kw in DISKON_KW:
                    if kw in n: return 'diskon'
                for kw in PREMI_KW:
                    if kw in n: return 'premi'
                return 'other'

            domain = [('state','=','posted'),
                      ('move_type','in',['out_invoice','out_refund','in_invoice','in_refund'])]

            if date_from:
                domain.append(('invoice_date','>=', date_from))
            if date_to:
                domain.append(('invoice_date','<=', date_to))

            if move_type == 'out_invoice':
                domain = [d for d in domain if not (isinstance(d, tuple) and d[0]=='move_type')]
                domain.append(('move_type','in',['out_invoice','out_refund']))
            elif move_type == 'in_invoice':
                domain = [d for d in domain if not (isinstance(d, tuple) and d[0]=='move_type')]
                domain.append(('move_type','in',['in_invoice','in_refund']))

            if has_policy:
                if partner_name:
                    domain.append(('policy_id.customer_id.name','ilike', partner_name))
                if asuransi_name:
                    domain.append(('policy_id.vendor_id.name','ilike', asuransi_name))
            else:
                if partner_name:
                    domain.append(('partner_id.name','ilike', partner_name))

            moves = env['account.move'].sudo().search(domain, order='invoice_date desc, name desc', limit=500)

            records = []
            total_premi = 0.0
            total_diskon = 0.0
            total_brokerage = 0.0

            for m in moves:
                policy = None
                client_name = ''
                asuransi_str = ''
                policy_number = ''
                risk_type = ''

                if has_policy and m.policy_id:
                    policy = m.policy_id
                    client_name = policy.customer_id.name if policy.customer_id else (m.partner_id.name or '')
                    asuransi_str = policy.vendor_id.name if policy.vendor_id else ''
                    policy_number = getattr(policy, 'policy_number', '') or ''
                    if hasattr(policy, 'type_id') and policy.type_id:
                        risk_type = policy.type_id.name or ''
                else:
                    client_name = m.partner_id.name or '' if m.partner_id else ''

                sign = -1 if m.move_type in ('out_refund','in_refund') else 1
                premi_amt = 0.0
                diskon_amt = 0.0
                brok_amt = 0.0

                for line in m.invoice_line_ids.filtered(lambda l: l.display_type == 'product'):
                    pname = (line.product_id.name if line.product_id else '') or line.name or ''
                    cat = classify_line(pname)
                    amt = line.price_subtotal
                    if cat == 'premi':
                        premi_amt += amt
                    elif cat == 'diskon':
                        diskon_amt += amt
                    elif cat == 'brokerage':
                        brok_amt += amt

                total_premi     += premi_amt * sign
                total_diskon    += diskon_amt * sign
                total_brokerage += brok_amt * sign

                if len(records) < 200:
                    records.append({
                        'id': m.id,
                        'name': m.name or '',
                        'date': str(m.invoice_date) if m.invoice_date else '',
                        'ref': m.ref or '',
                        'client': client_name,
                        'asuransi': asuransi_str,
                        'policy_number': policy_number,
                        'risk_type': risk_type,
                        'currency': m.currency_id.name or 'IDR' if m.currency_id else 'IDR',
                        'premi': round(premi_amt * sign),
                        'diskon': round(diskon_amt * sign),
                        'brokerage': round(brok_amt * sign),
                        'amount_total': round(m.amount_total * sign),
                        'move_type': m.move_type,
                        'payment_state': m.payment_state or '',
                    })

            month_names_id = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agt','Sep','Okt','Nov','Des']
            base = today.replace(day=1)
            chart_labels = []
            chart_premi = []
            chart_diskon = []
            chart_brokerage = []

            for i in range(11, -1, -1):
                m_date = base - relativedelta(months=i)
                last_day = calendar.monthrange(m_date.year, m_date.month)[1]
                d_from = '%s-%02d-01' % (m_date.year, m_date.month)
                d_to   = '%s-%02d-%02d' % (m_date.year, m_date.month, last_day)
                chart_labels.append('%s %s' % (month_names_id[m_date.month-1], m_date.year))

                month_domain = [('state','=','posted'),
                                ('move_type','in',['out_invoice','out_refund','in_invoice','in_refund']),
                                ('invoice_date','>=', d_from),('invoice_date','<=', d_to)]
                m_moves = env['account.move'].sudo().search(month_domain)
                mp = md = mb = 0.0
                for mv in m_moves:
                    sg = -1 if mv.move_type in ('out_refund','in_refund') else 1
                    for line in mv.invoice_line_ids.filtered(lambda l: l.display_type == 'product'):
                        pname = (line.product_id.name if line.product_id else '') or line.name or ''
                        cat = classify_line(pname)
                        amt = line.price_subtotal * sg
                        if cat == 'premi': mp += amt
                        elif cat == 'diskon': md += amt
                        elif cat == 'brokerage': mb += amt
                chart_premi.append(round(mp))
                chart_diskon.append(round(md))
                chart_brokerage.append(round(mb))

            return {
                'records': records,
                'total_count': len(moves),
                'summary': {
                    'total_premi': round(total_premi),
                    'total_diskon': round(total_diskon),
                    'total_brokerage': round(total_brokerage),
                    'total_net': round(total_premi - total_diskon - total_brokerage),
                    'count': len(moves),
                },
                'chart': {
                    'labels': chart_labels,
                    'premi': chart_premi,
                    'diskon': chart_diskon,
                    'brokerage': chart_brokerage,
                },
            }
        except Exception as exc:
            _logger.exception('Production report data error: %s', exc)
            return {'records':[],'total_count':0,
                    'summary':{'total_premi':0,'total_diskon':0,'total_brokerage':0,'total_net':0,'count':0},
                    'chart':{'labels':[],'premi':[],'diskon':[],'brokerage':[]}}

    @http.route('/raysolusi/production_report/excel', type='http', auth='user', methods=['POST'], csrf=False)
    def export_production_excel(self, **post):
        import io as _io
        import json as _json
        try:
            import xlsxwriter
        except ImportError:
            return request.make_response('xlsxwriter not installed', [('Content-Type','text/plain')])

        env = request.env
        params = {}
        raw = post.get('params') or post.get('data') or ''
        if raw:
            try:
                params = _json.loads(raw)
            except Exception:
                pass

        date_from    = params.get('date_from')    or post.get('date_from')
        date_to      = params.get('date_to')      or post.get('date_to')
        partner_name = params.get('partner_name') or post.get('partner_name')
        asuransi_name = params.get('asuransi_name') or post.get('asuransi_name')
        move_type    = params.get('move_type')    or post.get('move_type', 'all')

        has_policy = 'policy_id' in env['account.move']._fields
        PREMI_KW     = ['premi','premium','insurance premium','hutang kepada insurance']
        DISKON_KW    = ['diskon','discount','insurance discount','brokerage discount']
        BROKERAGE_KW = ['brokerage fee','brokerage','komisi','commission','broker fee']

        def classify_line(name_str):
            if not name_str: return 'other'
            n = name_str.lower()
            for kw in BROKERAGE_KW:
                if kw in n: return 'brokerage'
            for kw in DISKON_KW:
                if kw in n: return 'diskon'
            for kw in PREMI_KW:
                if kw in n: return 'premi'
            return 'other'

        domain = [('state','=','posted'),
                  ('move_type','in',['out_invoice','out_refund','in_invoice','in_refund'])]
        if date_from: domain.append(('invoice_date','>=', date_from))
        if date_to:   domain.append(('invoice_date','<=', date_to))
        if move_type == 'out_invoice':
            domain = [d for d in domain if not (isinstance(d,tuple) and d[0]=='move_type')]
            domain.append(('move_type','in',['out_invoice','out_refund']))
        elif move_type == 'in_invoice':
            domain = [d for d in domain if not (isinstance(d,tuple) and d[0]=='move_type')]
            domain.append(('move_type','in',['in_invoice','in_refund']))
        if has_policy:
            if partner_name: domain.append(('policy_id.customer_id.name','ilike', partner_name))
            if asuransi_name: domain.append(('policy_id.vendor_id.name','ilike', asuransi_name))
        else:
            if partner_name: domain.append(('partner_id.name','ilike', partner_name))

        moves = env['account.move'].sudo().search(domain, order='invoice_date asc, name asc')

        output = _io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('Laporan Produksi')

        title_fmt = wb.add_format({'bold':True,'font_size':13,'bg_color':'#1F3864','font_color':'white','align':'center','valign':'vcenter'})
        hdr_fmt   = wb.add_format({'bold':True,'bg_color':'#1F3864','font_color':'white','border':1,'align':'center','text_wrap':True,'valign':'vcenter'})
        txt_fmt   = wb.add_format({'border':1})
        num_fmt   = wb.add_format({'border':1,'num_format':'#,##0.00'})
        date_fmt  = wb.add_format({'border':1,'num_format':'dd/mm/yyyy'})
        total_fmt = wb.add_format({'bold':True,'bg_color':'#D9E1F2','num_format':'#,##0.00','border':1})
        total_lbl = wb.add_format({'bold':True,'bg_color':'#D9E1F2','border':1})

        headers = [('No',4),('Tanggal',12),('No DN/CN',20),('Ref',18),
                   ('Client',28),('Asuransi',28),('Risk Type',22),('No Polis',20),
                   ('Mata Uang',10),('Premi',16),('Diskon',14),('Brokerage',14),('Total',16),('Tipe',12)]
        for c,(h,w) in enumerate(headers):
            ws.set_column(c,c,w)
        ws.set_row(0,28); ws.set_row(1,18)
        ws.merge_range(0,0,0,len(headers)-1,'PT RAYSOLUSI - LAPORAN PRODUKSI ASURANSI', title_fmt)
        period_str = 'Periode: %s s/d %s' % (date_from or '-', date_to or '-')
        ws.merge_range(1,0,1,len(headers)-1, period_str, wb.add_format({'align':'center','italic':True}))
        for c,(h,_) in enumerate(headers):
            ws.write(2,c,h,hdr_fmt)

        row_num = 3
        tp = td = tb = 0.0
        for idx, m in enumerate(moves, 1):
            client_name = asuransi_str = policy_number = risk_type = ''
            if has_policy and m.policy_id:
                p = m.policy_id
                client_name   = p.customer_id.name if p.customer_id else (m.partner_id.name or '')
                asuransi_str  = p.vendor_id.name if p.vendor_id else ''
                policy_number = getattr(p,'policy_number','') or ''
                if hasattr(p,'type_id') and p.type_id: risk_type = p.type_id.name or ''
            else:
                client_name = m.partner_id.name or '' if m.partner_id else ''

            sign = -1 if m.move_type in ('out_refund','in_refund') else 1
            premi_amt = diskon_amt = brok_amt = 0.0
            for line in m.invoice_line_ids.filtered(lambda l: l.display_type == 'product'):
                pname = (line.product_id.name if line.product_id else '') or line.name or ''
                cat = classify_line(pname)
                amt = line.price_subtotal
                if cat == 'premi': premi_amt += amt
                elif cat == 'diskon': diskon_amt += amt
                elif cat == 'brokerage': brok_amt += amt

            premi_s  = premi_amt * sign
            diskon_s = diskon_amt * sign
            brok_s   = brok_amt * sign
            tp += premi_s; td += diskon_s; tb += brok_s

            type_map = {'out_invoice':'Invoice(RV)','in_invoice':'Bill(PV)',
                        'out_refund':'CN(RV)','in_refund':'CN(PV)'}
            ws.write(row_num,0,idx,txt_fmt)
            if m.invoice_date: ws.write_datetime(row_num,1,m.invoice_date,date_fmt)
            else: ws.write(row_num,1,'',txt_fmt)
            ws.write(row_num,2,m.name or '',txt_fmt)
            ws.write(row_num,3,m.ref or '',txt_fmt)
            ws.write(row_num,4,client_name,txt_fmt)
            ws.write(row_num,5,asuransi_str,txt_fmt)
            ws.write(row_num,6,risk_type,txt_fmt)
            ws.write(row_num,7,policy_number,txt_fmt)
            ws.write(row_num,8,m.currency_id.name or 'IDR' if m.currency_id else 'IDR',txt_fmt)
            ws.write(row_num,9, premi_s, num_fmt)
            ws.write(row_num,10,diskon_s, num_fmt)
            ws.write(row_num,11,brok_s,  num_fmt)
            ws.write(row_num,12,m.amount_total*sign, num_fmt)
            ws.write(row_num,13,type_map.get(m.move_type, m.move_type), txt_fmt)
            row_num += 1

        ws.merge_range(row_num,0,row_num,8,'TOTAL',total_lbl)
        ws.write(row_num,9,tp,total_fmt)
        ws.write(row_num,10,td,total_fmt)
        ws.write(row_num,11,tb,total_fmt)
        ws.write(row_num,12,tp-td-tb,total_fmt)
        ws.write(row_num,13,'',total_lbl)

        wb.close()
        xlsx_data = output.getvalue()
        filename = 'laporan_produksi_%s.xlsx' % date.today().strftime('%Y%m%d')
        headers_resp = [
            ('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition','attachment; filename="%s"' % filename),
            ('Content-Length', str(len(xlsx_data))),
        ]
        return request.make_response(xlsx_data, headers_resp)
