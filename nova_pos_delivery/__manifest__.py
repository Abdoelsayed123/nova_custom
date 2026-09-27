{
    'name': 'NOVA POS Delivery Mode',
    'version': '20.0.1.0.0',
    'category': 'Sales/Point of Sale',
    'summary': 'Choose store pickup or delivery on POS orders.',
    'depends': ['point_of_sale'],
    'data': [
        'views/pos_order_views.xml',
    ],
    'assets': {
        'point_of_sale._assets_pos': [
            'nova_pos_delivery/static/src/js/delivery_mode.js',
            'nova_pos_delivery/static/src/xml/delivery_mode.xml',
            'nova_pos_delivery/static/src/scss/delivery_mode.scss',
        ],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
