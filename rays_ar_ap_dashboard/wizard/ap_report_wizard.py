# -*- coding: utf-8 -*-
import io
import base64
from datetime import date
import xlsxwriter
from odoo import models, fields, api, _


class RaysAPReport(models.TransientModel):
    _name = 'rays.ap.report'
    _description = 'RAYS AP Report - Hutang Premi & Brokerage'

    date_from = fields.Date('Dari Tanggal', default=lambda self: date(2025, 1, 1))
    date_to = fields.Date('Sampai Tanggal', default=fields.Date.today)
    category = fields.Selection([
        ('all', 'Semua'),
        ('premium', 'PREMI - Hutang ke Asuransi'),
        ('brokerage', 'BROKERAGE - Komisi'),
    ], string='Kategori', default='all')
    partner_id = fields.Many2one('res.partner', 'Vendor/Asuransi')
    currency_id = fields.Many2one('res.currency', 'Mata Uang')
    payment_state = fields.Selection([
        ('all', 'Semua Status'),
        ('not_paid', 'Belum Bayar'),
        ('partial', 'Bayar Sebagian'),
        ('paid', 'Lunas'),
    ], string='Status Pembayaran', default='all')

    # Summary fields
    total_premi = fields.Float('Total PREMI/Hutang Asuransi (IDR)', readonly=True)
    total_brokerage = fields.Float('Total BROKERAGE (IDR)', readonly=True)
    total_net = fields.Float('Total NET AP (IDR)', readonly=True)
    count_premi = fields.Integer('Jumlah Bill PREMI', readonly=True)
    count_brokerage = fields.Integer('Jumlah Bill BROKERAGE', readonly=True)

    # Excel download
    excel_file = fields.Binary('File Excel', readonly=True)
    excel_filename = fields.Char('Nama File', readonly=True)

    def _build_ap_domain(self):
        domain = [
            ('move_id.state', '=', 'posted'),
            ('account_id.account_type', '=', 'liability_payable'),
            ('move_id.move_type', 'in', ['in_invoice', 'in_refund']),
        ]
        if self.date_from:
            domain.append(('date', '>=', self.date_from))
        if self.date_to:
            domain.append(('date', '<=', self.date_to))
        if self.partner_id:
            domain.append(('partner_id', '=', self.partner_id.id))
        if self.currency_id:
            domain.append(('currency_id', '=', self.currency_id.id))
        if self.payment_state and self.payment_state != 'all':
            domain.append(('move_id.payment_state', '=', self.payment_state))
        return domain

    def _get_brokerage_lines(self, ap_move_ids):
        """Get brokerage/income lines linked to the same AP invoices."""
        if not ap_move_ids:
            return self.env['account.move.line']
        return self.env['account.move.line'].search([
            ('move_id', 'in', ap_move_ids),
            ('move_id.state', '=', 'posted'),
            ('account_id.account_type', '=', 'income'),
            ('product_id.name', 'ilike', 'brokerage'),
        ])

    def action_view_report(self):
        ap_lines = self.env['account.move.line'].search(self._build_ap_domain())
        ap_move_ids = ap_lines.mapped('move_id').ids
        brok_lines = self._get_brokerage_lines(ap_move_ids)

        self.write({
            'total_premi': abs(sum(ap_lines.mapped('balance'))),
            'total_brokerage': abs(sum(brok_lines.mapped('balance'))),
            'total_net': abs(sum(ap_lines.mapped('balance'))) - abs(sum(brok_lines.mapped('balance'))),
            'count_premi': len(ap_lines.mapped('move_id')),
            'count_brokerage': len(brok_lines.mapped('move_id')),
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'rays.ap.report',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }

    def action_export_excel(self):
        ap_lines = self.env['account.move.line'].search(
            self._build_ap_domain(), order='date asc')
        ap_move_ids = ap_lines.mapped('move_id').ids
        brok_lines = self._get_brokerage_lines(ap_move_ids)

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # ---- Formats ----
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14,
            'bg_color': '#7B3F00', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter',
        })
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#C0392B', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        subheader_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FADBD8', 'border': 1,
        })
        money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
        money_neg_fmt = workbook.add_format({
            'num_format': '#,##0', 'border': 1, 'font_color': 'red',
        })
        date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        text_fmt = workbook.add_format({'border': 1})
        total_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#7B3F00', 'font_color': 'white',
            'num_format': '#,##0', 'border': 1,
        })
        summary_label = workbook.add_format({
            'bold': True, 'bg_color': '#FDEDEC', 'border': 1,
        })
        summary_value = workbook.add_format({
            'bold': True, 'num_format': '#,##0', 'bg_color': '#FDEDEC', 'border': 1,
        })

        def write_ap_sheet(ws, lines, sheet_title, is_brokerage=False):
            ws.set_zoom(85)
            col_widths = [5, 12, 22, 35, 10, 18, 18, 18, 15, 12]
            for c, w in enumerate(col_widths):
                ws.set_column(c, c, w)
            ws.set_row(0, 30)
            ws.set_row(1, 20)

            ws.merge_range('A1:J1', f'PT RAYSOLUSI - {sheet_title}', title_fmt)
            period = ''
            if self.date_from:
                period += f'Dari: {self.date_from.strftime("%d/%m/%Y")}  '
            if self.date_to:
                period += f'Sampai: {self.date_to.strftime("%d/%m/%Y")}'
            ws.merge_range('A2:J2', period, subheader_fmt)

            headers = ['No', 'Tanggal', 'No Bill/Ref', 'Vendor/Asuransi',
                       'Mata Uang', 'Jumlah Asal', 'IDR Amount', 'Sisa Hutang',
                       'Status', 'Kategori']
            for col, h in enumerate(headers):
                ws.write(2, col, h, header_fmt)

            row = 3
            total_orig = 0.0
            total_idr = 0.0
            total_residual = 0.0
            for idx, line in enumerate(lines, 1):
                move = line.move_id
                kategori = 'BROKERAGE' if is_brokerage else 'PREMI'
                idr_amount = abs(line.balance)
                orig_amount = abs(line.amount_currency) if line.currency_id else idr_amount
                residual = move.amount_residual
                pay_state = dict(
                    move._fields['payment_state'].selection
                ).get(move.payment_state, move.payment_state)

                ws.write(row, 0, idx, text_fmt)
                if move.invoice_date or move.date:
                    ws.write_datetime(row, 1, move.invoice_date or move.date, date_fmt)
                else:
                    ws.write(row, 1, '', text_fmt)
                ws.write(row, 2, move.name or move.ref or '', text_fmt)
                ws.write(row, 3, (line.partner_id.name or '')[:50], text_fmt)
                ws.write(row, 4, line.currency_id.name if line.currency_id else 'IDR', text_fmt)
                ws.write_number(row, 5, float(orig_amount or 0), money_fmt)
                ws.write_number(row, 6, float(idr_amount or 0), money_fmt)
                ws.write_number(row, 7, float(residual or 0),
                                money_neg_fmt if residual < 0 else money_fmt)
                ws.write(row, 8, pay_state, text_fmt)
                ws.write(row, 9, kategori, text_fmt)
                total_orig += float(orig_amount or 0)
                total_idr += float(idr_amount or 0)
                total_residual += float(residual or 0)
                row += 1

            ws.write(row, 0, 'TOTAL', total_fmt)
            ws.merge_range(row, 1, row, 4, '', total_fmt)
            ws.write_number(row, 5, total_orig, total_fmt)
            ws.write_number(row, 6, total_idr, total_fmt)
            ws.write_number(row, 7, total_residual, total_fmt)
            ws.write(row, 8, '', total_fmt)
            ws.write(row, 9, '', total_fmt)
            row += 2

            # Summary
            ws.write(row, 0, 'RINGKASAN', subheader_fmt)
            ws.merge_range(row, 1, row, 3, '', subheader_fmt)
            row += 1
            ws.write(row, 0, 'Total PREMI/Hutang Asuransi (IDR)', summary_label)
            ws.write_number(row, 1, abs(sum(ap_lines.mapped('balance'))), summary_value)
            row += 1
            ws.write(row, 0, 'Total BROKERAGE (IDR)', summary_label)
            ws.write_number(row, 1, abs(sum(brok_lines.mapped('balance'))), summary_value)

        # Sheet 1: AP PREMI (liability_payable lines)
        ws1 = workbook.add_worksheet('AP PREMI')
        write_ap_sheet(ws1, ap_lines, 'AP Outstanding - PREMI (Hutang ke Asuransi)')

        # Sheet 2: BROKERAGE lines
        ws2 = workbook.add_worksheet('AP BROKERAGE')
        write_ap_sheet(ws2, brok_lines, 'AP - BROKERAGE / Komisi', is_brokerage=True)

        # Sheet 3: Summary by Invoice
        ws3 = workbook.add_worksheet('Ringkasan per Invoice')
        ws3.set_column(0, 0, 5)
        ws3.set_column(1, 1, 22)
        ws3.set_column(2, 2, 35)
        ws3.set_column(3, 5, 18)
        ws3.set_column(6, 6, 15)
        ws3.merge_range('A1:G1', 'PT RAYSOLUSI - AP Ringkasan per Invoice', title_fmt)
        headers3 = ['No', 'No Bill/Ref', 'Vendor/Asuransi', 'PREMI (IDR)',
                    'BROKERAGE (IDR)', 'NET AP (IDR)', 'Status']
        for col, h in enumerate(headers3):
            ws3.write(2, col, h, header_fmt)
        row3 = 3
        seen_moves = {}
        for line in ap_lines:
            move_id = line.move_id.id
            if move_id not in seen_moves:
                seen_moves[move_id] = {
                    'move': line.move_id,
                    'premi': 0.0, 'brokerage': 0.0,
                }
            seen_moves[move_id]['premi'] += abs(line.balance)
        for line in brok_lines:
            move_id = line.move_id.id
            if move_id in seen_moves:
                seen_moves[move_id]['brokerage'] += abs(line.balance)
        for idx, (mid, data) in enumerate(seen_moves.items(), 1):
            move = data['move']
            premi = data['premi']
            brok = data['brokerage']
            pay_state = dict(
                move._fields['payment_state'].selection
            ).get(move.payment_state, move.payment_state)
            ws3.write(row3, 0, idx, text_fmt)
            ws3.write(row3, 1, move.name or '', text_fmt)
            partner = move.partner_id.name if move.partner_id else ''
            ws3.write(row3, 2, partner[:50], text_fmt)
            ws3.write_number(row3, 3, premi, money_fmt)
            ws3.write_number(row3, 4, brok, money_fmt)
            ws3.write_number(row3, 5, premi - brok, money_fmt)
            ws3.write(row3, 6, pay_state, text_fmt)
            row3 += 1

        workbook.close()
        output.seek(0)

        fname = f'AP_Report_{fields.Date.today()}.xlsx'
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'excel_filename': fname,
            'total_premi': abs(sum(ap_lines.mapped('balance'))),
            'total_brokerage': abs(sum(brok_lines.mapped('balance'))),
            'total_net': abs(sum(ap_lines.mapped('balance'))) - abs(sum(brok_lines.mapped('balance'))),
            'count_premi': len(ap_lines.mapped('move_id')),
            'count_brokerage': len(brok_lines.mapped('move_id')),
        })

        return {
            'type': 'ir.actions.act_url',
            'url': (f'/web/content/?model=rays.ap.report&id={self.id}'
                    f'&field=excel_file&filename={fname}&download=true'),
            'target': 'new',
        }
