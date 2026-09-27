{
    'name': 'NOVA System Background',
    'version': '20.0.1.0.0',
    'category': 'Themes/Backend',
    'summary': 'Set NOVA backgrounds in the backend and login page.',
    'author': 'NOVA',
    'description': 'Login photo by David Tomaseti via Unsplash, used under the Unsplash License.',
    'depends': ['web', 'website'],
    'data': [
        'views/login_brand.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'nova_system_background/static/src/scss/background.scss',
        ],
        'web.assets_frontend': [
            'nova_system_background/static/src/scss/login.scss',
        ],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}