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
{
    'name': 'Cashflow',
    'version': '1.0',
    'category': 'Accounting',
    'summary': "Cashflow management: creates an excel sheet of incoming / outcoming payments",
    'author': 'Stefano Siccardi @ CQ Creativi Quadrati',
    'license': 'AGPL-3',
    'depends' : ['sale','purchase','account'],
    'update_xml' : [
        'security/ir.model.access.csv',
        'wizard/previsione_in_out_view.xml',
        'config_cashflow_base_view.xml',
        'account_view.xml',
        'sale_view.xml',
        'purchase_view.xml',
    ],
    'installable': True,
    'auto-install': False
}
