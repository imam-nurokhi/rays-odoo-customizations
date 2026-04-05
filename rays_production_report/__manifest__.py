# -*- coding: utf-8 -*-
{
    "name": "RAYS Production Report",
    "version": "18.0.1.0.0",
    "category": "Accounting/Reports",
    "summary": "Insurance Production Report for PT RAYSOLUSI",
    "author": "PT RAYSOLUSI",
    "depends": ["account", "base", "rays_insurance"],
    "data": [
        "security/ir.model.access.csv",
        "views/production_report_views.xml",
    ],
    "installable": True,
    "auto_install": False,
    "license": "LGPL-3",
}
