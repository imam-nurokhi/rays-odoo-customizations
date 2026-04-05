# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Server & Deploy
- SSH: `theProg@103.196.153.12 -i ~/.ssh/979798.pem`
- Docker container: `odoo18`; addons path: `/home/odoo/odoo18_docker/addons/raysolusi_ar_ap_dashboard/`
- Production DB: `rays_demo`; admin: `admin` / `adm_1`
- Deploy: `scp file theProg@host:/tmp/file && ssh ... 'sudo cp /tmp/file /home/odoo/.../file'`
- After JS/CSS/XML changes: `DELETE FROM ir_attachment WHERE name LIKE '%web.assets_backend%';` then hard refresh (Cmd+Shift+R)
- After Python changes: `sudo docker restart odoo18` (wait ~20s, verify with `curl -o /dev/null -w "%{http_code}" http://localhost:8069/`)
- Avoid `button_immediate_upgrade` via xmlrpc on running instance — it can break Odoo registry; use docker restart instead

## Odoo 18 OWL 2 — Critical Patterns
- RPC calls: `import { rpc } from "@web/core/network/rpc"; await rpc("/url", {})` — `useService("rpc")` does NOT exist in Odoo 18
- Client action components require: `static props = { "*": true };`
- Template expressions are JavaScript, not Python: use `&&` / `||`, never `and` / `or`
- Escape `&` as `&amp;` in XML attributes; `t-if="!x &amp;&amp; y"` is correct
- Use `t-att-value` + `t-on-change`/`t-on-input` for inputs — do NOT use `t-model`
- Modal close-on-overlay: use `t-on-click="closeModal"` on overlay, `t-on-click.stop=""` on inner panel
- Module registration: `registry.category("actions").add("tag_name", ComponentClass);`
- Every JS file must start with `/** @odoo-module **/`

## Architecture
- OWL component: `static/src/js/ar_ap_dashboard.js` — `ArApDashboard` class, Chart.js charts with click handlers
- Template: `static/src/xml/ar_ap_dashboard.xml` — template name `raysolusi_ar_ap_dashboard.Dashboard`
- Chart.js 4.4.3 loaded as UMD via `static/lib/chart.js/chart.umd.min.js` → accessed as `window.Chart`
- Backend: `controllers/main.py` — two JSON routes:
  - `POST /raysolusi/ar_ap_dashboard/data` — dashboard summary + charts (accepts `date_from`, `date_to`, `partner_name` filters)
  - `POST /raysolusi/ar_ap_dashboard/detail` — drill-down records (accepts `move_type`, `filter_type`, `filter_value`)
- Menu: `views/menu_views.xml` — client action tag `raysolusi_ar_ap_dashboard` under `account.menu_finance_reports`

## Known DB & Server Quirks
- `odoo18` database is empty — production is always `rays_demo`
- Set password via psql plain text works: `UPDATE res_users SET password='newpwd' WHERE login='admin';`
- Odoo 18 xmlrpc domain `['name','=',False]` crashes — filter client-side instead
- Multiple databases: `rays_demo` (prod), `rays_backup`, `rays_test` (pwd: `admin`), `demo`
