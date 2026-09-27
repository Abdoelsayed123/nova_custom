from odoo import fields, models


class ProductTemplate(models.Model):
    _inherit = 'product.template'

    candle_wick_type = fields.Selection(
        selection=[
            ('wood', 'Wood'),
            ('cotton', 'Cotton'),
        ],
        string='Wick Type',
    )
    candle_burn_time = fields.Float(
        string='Burn Time (Hours)',
        digits=(6, 2),
        help='Expected total burning time of the candle, in hours.',
    )
