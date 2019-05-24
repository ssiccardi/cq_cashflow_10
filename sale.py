# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2014
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

from odoo import models, fields

class SaleOrder(models.Model):
    _inherit = "sale.order"
    
    divisione_fatturazione_line = fields.One2many('divisione.fatturazione.sale', 'order_id', 'Divisione fatturazione')

class DivisioneFatturazioneSale(models.Model):
    _name = "divisione.fatturazione.sale"
    _order='data_prevista asc'
    
    order_id = fields.Many2one('sale.order', 'Ordine', required=True)
    order_currency_id = fields.Many2one('res.currency', 'Currency', related='order_id.currency_id')
    importo =  fields.Monetary('Importo', currency_field='order_currency_id')
    data_prevista = fields.Date('Data Prevista')

