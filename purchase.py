# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2018
#    Stefano Siccardi creativiquadrati snc
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

from odoo import api, models, fields

class PurchaseOrder(models.Model):
    _inherit = "purchase.order"
    
    divisione_fatturazione_line = fields.One2many('divisione.fatturazione.purchase', 'order_id', 'Divisione fatturazione')
    
    @api.one
    def create_div_fatt_line(self):
        if self.divisione_fatturazione_line:
            self.divisione_fatturazione_line.unlink()
        div_fatt_lines = {}
        for line in self.order_line:
            if line.date_planned:
                date = line.date_planned[:10]
                if self.company_id.tax_calculation_rounding_method == 'round_globally':
                    taxes = line.taxes_id.compute_all(line.price_unit, line.order_id.currency_id, line.product_qty, product=line.product_id, partner=line.order_id.partner_id)
                    amount_tax = sum(t.get('amount', 0.0) for t in taxes.get('taxes', []))
                else:
                    amount_tax = line.price_tax
                if date in div_fatt_lines:
                    div_fatt_lines[date] += line.price_subtotal + amount_tax
                else:
                    div_fatt_lines[date] = line.price_subtotal + amount_tax
        self.write({'divisione_fatturazione_line': map(lambda x: (0,0,{'importo': x[1], 'data_prevista': x[0]}), div_fatt_lines.items())})
        return True

class DivisioneFatturazionePurchase(models.Model):
    _name = "divisione.fatturazione.purchase"
    _order = 'data_prevista asc'
    
    order_id = fields.Many2one('purchase.order', 'Ordine', required=True)
    order_currency_id = fields.Many2one('res.currency', 'Currency', related='order_id.currency_id')
    importo =  fields.Monetary('Importo', currency_field='order_currency_id')
    data_prevista = fields.Date('Data Prevista')

