# -*- coding: utf-8 -*-
from collections import defaultdict
from datetime import date

from odoo import http
from odoo.http import request


class RaysDashboard(http.Controller):

    @http.route('/rays/ar_data', type='json', auth='user', methods=['POST'], csrf=False)
    def get_ar_data(self, date_from=None, date_to=None, **kwargs):
        domain = [
            ('move_id.state', '=', 'posted'),
            ('account_id.account_type', '=', 'asset_receivable'),
            ('move_id.move_type', 'in', ['out_invoice', 'out_refund']),
        ]
        if date_from:
            domain.append(('date', '>=', date_from))
        if date_to:
            domain.append(('date', '<=', date_to))
        AML = request.env['account.move.line']
        all_lines = AML.search(domain)
        premi = all_lines.filtered(lambda l: l.move_id.move_type == 'out_invoice')
        diskon = all_lines.filtered(lambda l: l.move_id.move_type == 'out_refund')
        monthly = defaultdict(lambda: {'premi': 0.0, 'diskon': 0.0})
        for line in premi:
            key = line.date.strftime('%Y-%m') if line.date else 'N/A'
            monthly[key]['premi'] += line.balance
        for line in diskon:
            key = line.date.strftime('%Y-%m') if line.date else 'N/A'
            monthly[key]['diskon'] += abs(line.balance)
        months_sorted = sorted(monthly.keys())
        return {
            'summary': {
                'total_premi': sum(premi.mapped('balance')),
                'total_diskon': abs(sum(diskon.mapped('balance'))),
                'net_ar': sum(all_lines.mapped('balance')),
                'count_premi': len(premi.mapped('move_id')),
                'count_diskon': len(diskon.mapped('move_id')),
            },
            'chart': {
                'labels': months_sorted,
                'premi': [monthly[m]['premi'] for m in months_sorted],
                'diskon': [monthly[m]['diskon'] for m in months_sorted],
            },
        }

    @http.route('/rays/ap_data', type='json', auth='user', methods=['POST'], csrf=False)
    def get_ap_data(self, date_from=None, date_to=None, **kwargs):
        domain_ap = [
            ('move_id.state', '=', 'posted'),
            ('account_id.account_type', '=', 'liability_payable'),
            ('move_id.move_type', 'in', ['in_invoice', 'in_refund']),
        ]
        if date_from:
            domain_ap.append(('date', '>=', date_from))
        if date_to:
            domain_ap.append(('date', '<=', date_to))
        AML = request.env['account.move.line']
        ap_lines = AML.search(domain_ap)
        ap_move_ids = ap_lines.mapped('move_id').ids
        brok_lines = AML.search([
            ('move_id', 'in', ap_move_ids),
            ('move_id.state', '=', 'posted'),
            ('account_id.account_type', '=', 'income'),
            ('product_id.name', 'ilike', 'brokerage'),
        ]) if ap_move_ids else AML
        monthly = defaultdict(lambda: {'premi': 0.0, 'brokerage': 0.0})
        for line in ap_lines:
            key = line.date.strftime('%Y-%m') if line.date else 'N/A'
            monthly[key]['premi'] += abs(line.balance)
        for line in brok_lines:
            key = line.date.strftime('%Y-%m') if line.date else 'N/A'
            monthly[key]['brokerage'] += abs(line.balance)
        months_sorted = sorted(monthly.keys())
        return {
            'summary': {
                'total_premi': abs(sum(ap_lines.mapped('balance'))),
                'total_brokerage': abs(sum(brok_lines.mapped('balance'))),
                'net_ap': abs(sum(ap_lines.mapped('balance'))) - abs(sum(brok_lines.mapped('balance'))),
                'count_bills': len(ap_lines.mapped('move_id')),
            },
            'chart': {
                'labels': months_sorted,
                'premi': [monthly[m]['premi'] for m in months_sorted],
                'brokerage': [monthly[m]['brokerage'] for m in months_sorted],
            },
        }

    @http.route('/rays/dashboard_data', type='json', auth='user', methods=['POST'], csrf=False)
    def dashboard_data(self, date_from=None, date_to=None, currency=None, **kwargs):
        env = request.env
        today = date.today()
        ar_domain = [
            ('account_id.account_type', '=', 'asset_receivable'),
            ('move_id.state', '=', 'posted'),
            ('reconciled', '=', False),
        ]
        ap_domain = [
            ('account_id.account_type', '=', 'liability_payable'),
            ('move_id.state', '=', 'posted'),
            ('reconciled', '=', False),
        ]
        if date_from:
            ar_domain.append(('date', '>=', date_from))
            ap_domain.append(('date', '>=', date_from))
        if date_to:
            ar_domain.append(('date', '<=', date_to))
            ap_domain.append(('date', '<=', date_to))
        AML = env['account.move.line']
        ar_lines = AML.search(ar_domain)
        ap_lines = AML.search(ap_domain)
        monthly_ar = defaultdict(float)
        monthly_ap = defaultdict(float)
        for line in ar_lines:
            key = line.date.strftime('%Y-%m') if line.date else 'unknown'
            monthly_ar[key] += float(line.balance)
        for line in ap_lines:
            key = line.date.strftime('%Y-%m') if line.date else 'unknown'
            monthly_ap[key] += float(abs(line.balance))
        ar_premi = sum(float(l.balance) for l in ar_lines if l.move_id.move_type == 'out_invoice')
        ar_diskon = sum(float(abs(l.balance)) for l in ar_lines if l.move_id.move_type == 'out_refund')
        ap_premi = sum(float(abs(l.balance)) for l in ap_lines if l.move_id.move_type == 'in_invoice')
        ap_broker = sum(float(abs(l.balance)) for l in ap_lines if l.move_id.move_type == 'in_refund')
        partner_ar = defaultdict(float)
        for line in ar_lines:
            if line.partner_id:
                partner_ar[line.partner_id.name] += float(line.balance)
        top_ar = sorted(partner_ar.items(), key=lambda x: -x[1])[:10]
        partner_ap = defaultdict(float)
        for line in ap_lines:
            if line.partner_id:
                partner_ap[line.partner_id.name] += float(abs(line.balance))
        top_ap = sorted(partner_ap.items(), key=lambda x: -x[1])[:10]
        total_ar = sum(float(l.balance) for l in ar_lines)
        total_ap = sum(float(abs(l.balance)) for l in ap_lines)
        overdue_ar = len(ar_lines.filtered(
            lambda l: l.date_maturity and l.date_maturity < today
        ))
        return {
            'kpi': {
                'total_ar': total_ar,
                'total_ap': total_ap,
                'net_position': total_ar - total_ap,
                'overdue_ar_count': overdue_ar,
            },
            'ar_by_category': {'PREMI': ar_premi, 'DISKON': ar_diskon},
            'ap_by_category': {'PREMI': ap_premi, 'BROKERAGE': ap_broker},
            'monthly_ar': dict(list(sorted(monthly_ar.items()))[-12:]),
            'monthly_ap': dict(list(sorted(monthly_ap.items()))[-12:]),
            'top_customers_ar': top_ar,
            'top_insurance_ap': top_ap,
        }
