from odoo import fields, models


class PosOrder(models.Model):
    _inherit = 'pos.order'

    delivery_mode = fields.Selection(
        selection=[
            ('pickup', 'Store Pickup'),
            ('delivery', 'Delivery'),
        ],
        string='Fulfillment Method',
        default='pickup',
        required=True,
        tracking=True,
    )

