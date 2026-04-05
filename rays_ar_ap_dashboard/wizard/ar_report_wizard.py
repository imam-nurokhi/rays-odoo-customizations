# -*- coding: utf-8 -*-
import io
import base64
from datetime import date
import xlsxwriter
from odoo import models, fields, api, _


class RaysARReport(models.TransientModel):
    _name = 'rays.ar.report'
    _description = 'RAYS AR Report - Piutang Premi & Diskon'

    date_from = fields.Date('Dari Tanggal', default=lambda self: date(2025, 1, 1))
    date_to = fields.Date('Sampai Tanggal', default=fields.Date.today)
    category = fields.Selection([
        ('all', 'Semua'),
        ('premium', 'PREMI'),
        ('discount', 'DISKON'),
    ], string='Kategori', default='all')
    partner_id = fields.Many2one('res.partner', 'Customer')
    currency_id = fields.Many2one('res.currency', 'Mata Uang')
    payment_state = fields.Selection([
        ('all', 'Semua Status'),
        ('not_paid', 'Belum Bayar'),
        ('partial', 'Bayar Sebagian'),
        ('paid', 'Lunas'),
    ], string='Status Pembayaran', default='all')

    # Summary fields (computed)
    total_premi = fields.Float('Total PREMI (IDR)', readonly=True)
    total_diskon = fields.Float('Total DISKON (IDR)', readonly=True)
    total_net = fields.Float('Total NET (IDR)', readonly=True)
    count_premi = fields.Integer('Jumlah Invoice PREMI', readonly=True)
    count_diskon = fields.Integer('Jumlah Credit Note DISKON', readonly=True)

    # Excel download
    excel_file = fields.Binary('File Excel', readonly=True)
    excel_filename = fields.Char('Nama File', readonly=True)

    def _build_domain(self, move_types=None):
        domain = [('move_id.state', '=', 'posted'),
                  ('account_id.account_type', '=', 'asset_receivable')]
        if move_types:
            domain.append(('move_id.move_type', 'in', move_types))
        else:
            domain.append(('move_id.move_type', 'in', ['out_invoice', 'out_refund']))
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

    def _get_category(self, move):
        """PREMI = customer invoice (out_invoice), DISKON = credit note (out_refund)."""
        return 'DISKON' if move.move_type == 'out_refund' else 'PREMI'

    def action_view_report(self):
        """Compute summary and reload wizard form."""
        domain_all = self._build_domain()
        all_lines = self.env['account.move.line'].search(domain_all, order='date asc')

        premi_lines = all_lines.filtered(lambda l: l.move_id.move_type == 'out_invoice')
        diskon_lines = all_lines.filtered(lambda l: l.move_id.move_type == 'out_refund')

        # Filter by category selection
        if self.category == 'premium':
            all_lines = premi_lines
        elif self.category == 'discount':
            all_lines = diskon_lines

        self.write({
            'total_premi': sum(premi_lines.mapped('balance')),
            'total_diskon': abs(sum(diskon_lines.mapped('balance'))),
            'total_net': sum(premi_lines.mapped('balance')) + sum(diskon_lines.mapped('balance')),
            'count_premi': len(premi_lines.mapped('move_id')),
            'count_diskon': len(diskon_lines.mapped('move_id')),
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'rays.ar.report',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }

    def action_export_excel(self):
        """Generate and download Excel report."""
        domain_all = self._build_domain()
        all_lines = self.env['account.move.line'].search(domain_all, order='date asc')

        premi_lines = all_lines.filtered(lambda l: l.move_id.move_type == 'out_invoice')
        diskon_lines = all_lines.filtered(lambda l: l.move_id.move_type == 'out_refund')

        if self.category == 'premium':
            lines_to_export = premi_lines
        elif self.category == 'discount':
            lines_to_export = diskon_lines
        else:
            lines_to_export = all_lines

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # ---- Formats ----
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14,
            'bg_color': '#1F4E79', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter',
        })
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        subheader_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D6E4F0', 'border': 1,
        })
        money_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
        money_red_fmt = workbook.add_format({
            'num_format': '#,##0', 'border': 1, 'font_color': 'red',
        })
        date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        text_fmt = workbook.add_format({'border': 1, 'text_wrap': False})
        total_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white',
            'num_format': '#,##0', 'border': 1,
        })
        summary_label = workbook.add_format({
            'bold': True, 'bg_color': '#E8F4FD', 'border': 1,
        })
        summary_value = workbook.add_format({
            'bold': True, 'num_format': '#,##0', 'bg_color': '#E8F4FD', 'border': 1,
        })

        def write_sheet(ws, lines, sheet_title, kategori_filter=None):
            ws.set_zoom(85)
            col_widths = [5, 12, 22, 35, 10, 18, 18, 18, 15, 10]
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

            headers = ['No', 'Tanggal', 'No Invoice/Ref', 'Customer/Partner',
                       'Mata Uang', 'Jumlah Asal', 'IDR Amount', 'Sisa Tagihan',
                       'Status', 'Kategori']
            for col, h in enumerate(headers):
                ws.write(2, col, h, header_fmt)

            row = 3
            total_orig = 0.0
            total_idr = 0.0
            total_residual = 0.0
            for idx, line in enumerate(lines, 1):
                move = line.move_id
                kategori = self._get_category(move)
                if kategori_filter and kategori != kategori_filter:
                    continue
                idr_amount = line.balance
                orig_amount = line.amount_currency if line.currency_id else idr_amount
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
                ws.write_number(row, 5, float(orig_amount or 0),
                                money_red_fmt if orig_amount < 0 else money_fmt)
                ws.write_number(row, 6, float(idr_amount or 0),
                                money_red_fmt if idr_amount < 0 else money_fmt)
                ws.write_number(row, 7, float(residual or 0),
                                money_red_fmt if residual < 0 else money_fmt)
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

            # Summary box
            ws.write(row, 0, 'RINGKASAN', subheader_fmt)
            ws.merge_range(row, 1, row, 3, '', subheader_fmt)
            row += 1
            ws.write(row, 0, 'Total PREMI (IDR)', summary_label)
            ws.write_number(row, 1,
                sum(l.balance for l in lines if l.move_id.move_type == 'out_invoice'),
                summary_value)
            row += 1
            ws.write(row, 0, 'Total DISKON (IDR)', summary_label)
            ws.write_number(row, 1,
                abs(sum(l.balance for l in lines if l.move_id.move_type == 'out_refund')),
                summary_value)
            row += 1
            ws.write(row, 0, 'NET AR (IDR)', summary_label)
            ws.write_number(row, 1, sum(l.balance for l in lines), summary_value)

        # Sheet 1: All AR
        ws1 = workbook.add_worksheet('AR Semua')
        write_sheet(ws1, lines_to_export, 'AR Outstanding - Semua')

        # Sheet 2: PREMI only
        ws2 = workbook.add_worksheet('AR PREMI')
        write_sheet(ws2, lines_to_export.filtered(
            lambda l: l.move_id.move_type == 'out_invoice'), 'AR Outstanding - PREMI')

        # Sheet 3: DISKON only
        ws3 = workbook.add_worksheet('AR DISKON')
        write_sheet(ws3, lines_to_export.filtered(
            lambda l: l.move_id.move_type == 'out_refund'), 'AR Outstanding - DISKON')

        workbook.close()
        output.seek(0)

        fname = f'AR_Report_{fields.Date.today()}.xlsx'
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'excel_filename': fname,
            'total_premi': sum(
                l.balance for l in lines_to_export if l.move_id.move_type == 'out_invoice'),
            'total_diskon': abs(sum(
                l.balance for l in lines_to_export if l.move_id.move_type == 'out_refund')),
            'total_net': sum(lines_to_export.mapped('balance')),
            'count_premi': len(lines_to_export.filtered(
                lambda l: l.move_id.move_type == 'out_invoice').mapped('move_id')),
            'count_diskon': len(lines_to_export.filtered(
                lambda l: l.move_id.move_type == 'out_refund').mapped('move_id')),
        })

        return {
            'type': 'ir.actions.act_url',
            'url': (f'/web/content/?model=rays.ar.report&id={self.id}'
                    f'&field=excel_file&filename={fname}&download=true'),
            'target': 'new',
        }
