# -*- coding: utf-8 -*-
import base64
import io
from datetime import date

import xlsxwriter

from odoo import api, fields, models, _
from odoo.exceptions import UserError


PREMI_KEYWORDS = ['hutang kepada insurance', 'premi', 'premium', 'insurance premium']
DISKON_KEYWORDS = ['diskon', 'discount', 'insurance discount', 'brokerage discount']
BROKERAGE_KEYWORDS = ['brokerage fee', 'brokerage', 'komisi', 'commission', 'broker fee']


def _classify_line(product_name):
    """Return one of: premi, diskon, brokerage, other"""
    if not product_name:
        return 'other'
    name_lower = product_name.lower()
    for kw in BROKERAGE_KEYWORDS:
        if kw in name_lower:
            return 'brokerage'
    for kw in DISKON_KEYWORDS:
        if kw in name_lower:
            return 'diskon'
    for kw in PREMI_KEYWORDS:
        if kw in name_lower:
            return 'premi'
    return 'other'


class RaysProductionReport(models.TransientModel):
    _name = 'rays.production.report'
    _description = 'Laporan Produksi Asuransi'

    date_from = fields.Date(string='Dari Tanggal', required=True,
                            default=lambda self: date(date.today().year, 1, 1))
    date_to = fields.Date(string='Sampai Tanggal', required=True,
                          default=fields.Date.context_today)
    partner_id = fields.Many2one('res.partner', string='Client',
                                 domain="[('customer_rank', '>', 0)]")
    asuransi_id = fields.Many2one('res.partner', string='Perusahaan Asuransi',
                                  domain="[('supplier_rank', '>', 0)]")
    currency_id = fields.Many2one('res.currency', string='Mata Uang')
    move_type = fields.Selection([
        ('all', 'Semua'),
        ('out_invoice', 'Invoice (RV)'),
        ('in_invoice', 'Bill (PV)'),
    ], string='Tipe Dokumen', default='all', required=True)

    def _get_invoice_data(self):
        domain = [
            ('state', '=', 'posted'),
            ('move_type', 'in', ['out_invoice', 'out_refund', 'in_invoice', 'in_refund']),
            ('invoice_date', '>=', self.date_from),
            ('invoice_date', '<=', self.date_to),
            ('policy_id', '!=', False),
        ]
        if self.partner_id:
            domain.append(('policy_id.customer_id', '=', self.partner_id.id))
        if self.asuransi_id:
            domain.append(('policy_id.vendor_id', '=', self.asuransi_id.id))
        if self.currency_id:
            domain.append(('currency_id', '=', self.currency_id.id))
        if self.move_type != 'all':
            if self.move_type == 'out_invoice':
                domain.append(('move_type', 'in', ['out_invoice', 'out_refund']))
            else:
                domain.append(('move_type', 'in', ['in_invoice', 'in_refund']))

        moves = self.env['account.move'].search(domain, order='invoice_date asc, name asc')
        return moves

    def action_generate_excel(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('Tanggal awal tidak boleh lebih besar dari tanggal akhir!'))

        moves = self._get_invoice_data()
        if not moves:
            raise UserError(_('Tidak ada data dalam rentang tanggal yang dipilih.'))

        rows = self._build_rows(moves)
        excel_data = self._generate_excel(rows)

        filename = 'laporan_produksi_%s_%s.xlsx' % (
            self.date_from.strftime('%Y%m%d'),
            self.date_to.strftime('%Y%m%d'),
        )
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': excel_data,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%d?download=true' % attachment.id,
            'target': 'self',
        }

    def _build_rows(self, moves):
        rows = []
        for move in moves:
            policy = move.policy_id
            client = policy.customer_id.name or move.partner_id.name or ''
            asuransi = policy.vendor_id.name or ''
            risk_type = ''
            if policy.type_id:
                risk_type = policy.type_id.name or ''
            currency = move.currency_id.name or ''

            premi_total = 0.0
            diskon_total = 0.0
            brokerage_total = 0.0

            product_lines = move.invoice_line_ids.filtered(
                lambda l: l.display_type == 'product'
            )
            for line in product_lines:
                product_name = ''
                if line.product_id and line.product_id.name:
                    product_name = line.product_id.name
                elif line.name:
                    product_name = line.name

                category = _classify_line(product_name)
                amount = line.price_subtotal

                if category == 'premi':
                    premi_total += amount
                elif category == 'diskon':
                    diskon_total += amount
                elif category == 'brokerage':
                    brokerage_total += amount

            sign = -1 if move.move_type in ('out_refund', 'in_refund') else 1
            rows.append({
                'tanggal': move.invoice_date,
                'no_doc': move.name,
                'ref': move.ref or '',
                'client': client,
                'asuransi': asuransi,
                'risk_type': risk_type,
                'policy_number': policy.policy_number or '',
                'currency': currency,
                'premi': premi_total * sign,
                'diskon': diskon_total * sign,
                'brokerage': brokerage_total * sign,
                'total': move.amount_total * sign,
                'move_type': move.move_type,
                'state': move.state,
            })
        return rows

    def _generate_excel(self, rows):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Formats
        fmt_title = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
        })
        fmt_header = workbook.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        fmt_cell = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        fmt_num = workbook.add_format({
            'border': 1, 'num_format': '#,##0.00', 'valign': 'vcenter',
        })
        fmt_date = workbook.add_format({'border': 1, 'num_format': 'dd/mm/yyyy', 'valign': 'vcenter'})
        fmt_total_label = workbook.add_format({
            'bold': True, 'border': 1, 'bg_color': '#D9E1F2',
        })
        fmt_total_num = workbook.add_format({
            'bold': True, 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#D9E1F2',
        })

        headers = [
            ('No', 5),
            ('Tanggal', 12),
            ('No DN/CN', 18),
            ('Ref', 18),
            ('Client', 28),
            ('Asuransi', 28),
            ('Risk Type', 22),
            ('No Polis', 20),
            ('Mata Uang', 10),
            ('Premi', 16),
            ('Diskon', 14),
            ('Brokerage', 14),
            ('Total', 16),
            ('Tipe', 12),
        ]

        sheet_defs = [
            ('Laporan Produksi', rows),
            ('PREMI', [r for r in rows if r['premi'] != 0.0]),
            ('DISKON', [r for r in rows if r['diskon'] != 0.0]),
            ('BROKERAGE', [r for r in rows if r['brokerage'] != 0.0]),
        ]

        period_str = '%s s/d %s' % (
            self.date_from.strftime('%d/%m/%Y'),
            self.date_to.strftime('%d/%m/%Y'),
        )

        for sheet_name, sheet_rows in sheet_defs:
            ws = workbook.add_worksheet(sheet_name)
            ws.set_row(0, 25)
            ws.set_row(1, 18)
            ws.merge_range(0, 0, 0, len(headers) - 1,
                           'LAPORAN PRODUKSI ASURANSI - PT RAYSOLUSI', fmt_title)
            ws.merge_range(1, 0, 1, len(headers) - 1,
                           'Periode: ' + period_str, workbook.add_format({'align': 'center'}))

            # Headers row
            for col, (h, w) in enumerate(headers):
                ws.write(3, col, h, fmt_header)
                ws.set_column(col, col, w)

            # Data rows
            total_premi = total_diskon = total_brokerage = total_amount = 0.0
            for row_idx, row in enumerate(sheet_rows, start=4):
                ws.write(row_idx, 0, row_idx - 3, fmt_cell)
                if row['tanggal']:
                    ws.write(row_idx, 1, row['tanggal'].strftime('%d/%m/%Y'), fmt_cell)
                else:
                    ws.write(row_idx, 1, '', fmt_cell)
                ws.write(row_idx, 2, row['no_doc'], fmt_cell)
                ws.write(row_idx, 3, row['ref'], fmt_cell)
                ws.write(row_idx, 4, row['client'], fmt_cell)
                ws.write(row_idx, 5, row['asuransi'], fmt_cell)
                ws.write(row_idx, 6, row['risk_type'], fmt_cell)
                ws.write(row_idx, 7, row['policy_number'], fmt_cell)
                ws.write(row_idx, 8, row['currency'], fmt_cell)
                ws.write_number(row_idx, 9, row['premi'], fmt_num)
                ws.write_number(row_idx, 10, row['diskon'], fmt_num)
                ws.write_number(row_idx, 11, row['brokerage'], fmt_num)
                ws.write_number(row_idx, 12, row['total'], fmt_num)
                ws.write(row_idx, 13, row['move_type'], fmt_cell)
                total_premi += row['premi']
                total_diskon += row['diskon']
                total_brokerage += row['brokerage']
                total_amount += row['total']

            # Totals row
            total_row = len(sheet_rows) + 4
            ws.merge_range(total_row, 0, total_row, 8, 'TOTAL', fmt_total_label)
            ws.write_number(total_row, 9, total_premi, fmt_total_num)
            ws.write_number(total_row, 10, total_diskon, fmt_total_num)
            ws.write_number(total_row, 11, total_brokerage, fmt_total_num)
            ws.write_number(total_row, 12, total_amount, fmt_total_num)
            ws.write(total_row, 13, '', fmt_total_label)

        # Summary sheet
        ws_sum = workbook.add_worksheet('Summary')
        ws_sum.set_column(0, 0, 30)
        ws_sum.set_column(1, 5, 18)

        ws_sum.merge_range(0, 0, 0, 5, 'SUMMARY LAPORAN PRODUKSI - PT RAYSOLUSI', fmt_title)
        ws_sum.merge_range(1, 0, 1, 5, 'Periode: ' + period_str,
                           workbook.add_format({'align': 'center'}))

        # Summary by currency
        ws_sum.write(3, 0, 'Summary by Mata Uang', fmt_header)
        ws_sum.write(3, 1, 'Total Premi', fmt_header)
        ws_sum.write(3, 2, 'Total Diskon', fmt_header)
        ws_sum.write(3, 3, 'Total Brokerage', fmt_header)
        ws_sum.write(3, 4, 'Total Dokumen', fmt_header)
        ws_sum.write(3, 5, 'Jumlah Invoice', fmt_header)

        currency_summary = {}
        for row in rows:
            cur = row['currency']
            if cur not in currency_summary:
                currency_summary[cur] = {'premi': 0, 'diskon': 0, 'brokerage': 0, 'total': 0, 'count': 0}
            currency_summary[cur]['premi'] += row['premi']
            currency_summary[cur]['diskon'] += row['diskon']
            currency_summary[cur]['brokerage'] += row['brokerage']
            currency_summary[cur]['total'] += row['total']
            currency_summary[cur]['count'] += 1

        cur_row = 4
        for cur, vals in sorted(currency_summary.items()):
            ws_sum.write(cur_row, 0, cur, fmt_cell)
            ws_sum.write_number(cur_row, 1, vals['premi'], fmt_num)
            ws_sum.write_number(cur_row, 2, vals['diskon'], fmt_num)
            ws_sum.write_number(cur_row, 3, vals['brokerage'], fmt_num)
            ws_sum.write_number(cur_row, 4, vals['total'], fmt_num)
            ws_sum.write_number(cur_row, 5, vals['count'], fmt_cell)
            cur_row += 1

        # Summary by client
        cur_row += 2
        ws_sum.write(cur_row, 0, 'Summary by Client', fmt_header)
        ws_sum.write(cur_row, 1, 'Mata Uang', fmt_header)
        ws_sum.write(cur_row, 2, 'Total Premi', fmt_header)
        ws_sum.write(cur_row, 3, 'Total Diskon', fmt_header)
        ws_sum.write(cur_row, 4, 'Total Brokerage', fmt_header)
        ws_sum.write(cur_row, 5, 'Total Dokumen', fmt_header)
        cur_row += 1

        client_summary = {}
        for row in rows:
            key = (row['client'], row['currency'])
            if key not in client_summary:
                client_summary[key] = {'premi': 0, 'diskon': 0, 'brokerage': 0, 'total': 0}
            client_summary[key]['premi'] += row['premi']
            client_summary[key]['diskon'] += row['diskon']
            client_summary[key]['brokerage'] += row['brokerage']
            client_summary[key]['total'] += row['total']

        for (client, cur), vals in sorted(client_summary.items()):
            ws_sum.write(cur_row, 0, client, fmt_cell)
            ws_sum.write(cur_row, 1, cur, fmt_cell)
            ws_sum.write_number(cur_row, 2, vals['premi'], fmt_num)
            ws_sum.write_number(cur_row, 3, vals['diskon'], fmt_num)
            ws_sum.write_number(cur_row, 4, vals['brokerage'], fmt_num)
            ws_sum.write_number(cur_row, 5, vals['total'], fmt_num)
            cur_row += 1

        workbook.close()
        output.seek(0)
        return base64.b64encode(output.read()).decode('utf-8')
