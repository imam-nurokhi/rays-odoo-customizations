{
    'name': 'RAYS AR/AP Dashboard',
    'version': '18.0.2.0.0',
    'category': 'Accounting/Reports',
    'summary': 'RAYS Reports top-level menu, AR/AP wizards, and interactive OWL dashboard',
    'author': 'PT RAYSOLUSI',
    'depends': ['account', 'base', 'rays_production_report'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/ar_report_wizard.xml',
        'wizard/ap_report_wizard.xml',
        'views/ar_dashboard.xml',
        'views/ap_dashboard.xml',
        'views/dashboard_action.xml',
        'views/menu.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'web/static/lib/Chart/Chart.js',
            'rays_ar_ap_dashboard/static/src/components/RaysDashboard.xml',
            'rays_ar_ap_dashboard/static/src/components/RaysDashboard.scss',
            'rays_ar_ap_dashboard/static/src/components/RaysDashboard.js',
        ],
    },
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}
