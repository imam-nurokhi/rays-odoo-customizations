{
    'name': 'Raysolusi AR & AP Dashboard',
    'version': '18.0.3.0.0',
    'summary': 'Dashboard AR & AP dengan grafik interaktif untuk PT. Raysolusi Pialang Asuransi',
    'description': """
        Dashboard interaktif menampilkan:
        - KPI Cards: Total AR, AP, Overdue AR, Net Position
        - Aging Analysis AR & AP (Belum JT, 1-30, 31-60, 61-90, >90 hari)
        - Top 10 Debitor & Kreditor
        - Trend Bulanan AR vs AP (12 bulan)
        - Status Pembayaran AR & AP (Doughnut)
        - Top Overdue AR & AP
    """,
    'author': 'Nexora Technology',
    'website': 'https://nexoratech.id',
    'category': 'Accounting/Accounting',
    'depends': ['account'],
    'data': [
        'security/ir.model.access.csv',
        'views/menu_views.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'raysolusi_ar_ap_dashboard/static/lib/chart.js/chart.umd.min.js',
            'raysolusi_ar_ap_dashboard/static/src/css/ar_ap_dashboard.css',
            'raysolusi_ar_ap_dashboard/static/src/xml/ar_ap_dashboard.xml',
            'raysolusi_ar_ap_dashboard/static/src/js/ar_ap_dashboard.js',
        ],
    },
    'installable': True,
    'auto_install': False,
    'application': False,
    'license': 'LGPL-3',
}
