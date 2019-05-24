This module works with openerp / odoo rel.10.

It creates an excel sheet with a cashflow, that is a forecast of the income and outcome of money in a time period. 

It takes into account:

- sale and purchase invoices each with its payments terms (delays and installments)

- sale and purchase orders, provided the user fills information about the dates when he thinks that will issue or receive invoces; for purchase orders the default is to receive invoices when the goods are expected to arrive

- moves of accounts selected in the module configuration section

Limitations:

- all incomes and outcomes are put together (no distinction between bank accounts or cash, etc.)

- expected tax payments (included VAT) are not computed

Note: this module is working and has been used in practice, but should not however be considered as a fully finished product. This includes translations: many comments and some terms may be only in Italian!

Please look at it as a starting point and taylor it to your needs.
