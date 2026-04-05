# Rays Odoo 18 Customizations

Custom Odoo 18 modules for **PT. Raysolusi Pialang Asuransi** — an insurance brokerage firm using Odoo 18 for accounting, AR/AP management, and production reporting.

---

## 📦 Modules

### 1. `raysolusi_ar_ap_dashboard` *(Primary)*
Interactive AR/AP Dashboard with 4 integrated tabs:

| Tab | Description |
|-----|-------------|
| 📊 **Dashboard** | 9 interactive Chart.js charts — AR/AP trends, category breakdown, top customers/insurance, aging |
| 📋 **AR Report** | Accounts Receivable filtered by PREMI & DISKON, with KPI cards, chart, table, Excel export |
| 📋 **AP Report** | Accounts Payable filtered by PREMI, DISKON Asuransi & BROKERAGE, with KPI cards, chart, table, Excel export |
| 🏭 **Laporan Produksi** | Production report per Client, Insurance, Risk Type, with PREMI/DISKON/BROKERAGE breakdown + Excel export |

**Key features:**
- Filter by date range and partner
- Interactive clickable cards and charts
- One-click Excel export for each report tab
- Real-time data from Odoo `account.move` and `account.move.line`

**Routes:**
```
GET  /raysolusi/ar_ap_dashboard/data
GET  /raysolusi/ar_report/data
GET  /raysolusi/ar_report/excel
GET  /raysolusi/ap_report/data
GET  /raysolusi/ap_report/excel
GET  /raysolusi/production_report/data
GET  /raysolusi/production_report/excel
```

---

### 2. `rays_ar_ap_dashboard`
Supporting wizard module providing:
- `rays.ar.report` — AR Excel report TransientModel
- `rays.ap.report` — AP Excel report TransientModel

---

### 3. `rays_production_report`
Production report module:
- `rays.production.report` — Production Excel wizard
- Classifies journal items as PREMI, DISKON, or BROKERAGE using keyword matching

---

## 🗂️ Repository Structure

```
rays-odoo-customizations/
├── raysolusi_ar_ap_dashboard/        # Primary dashboard addon
│   ├── controllers/
│   │   └── main.py                   # 8 HTTP routes, data queries
│   ├── static/src/
│   │   ├── js/ar_ap_dashboard.js     # OWL component, tabs, charts, export
│   │   ├── xml/ar_ap_dashboard.xml   # QWeb template (4 tabs)
│   │   └── css/ar_ap_dashboard.css   # Dashboard styles
│   ├── views/                        # Menu definitions
│   └── __manifest__.py
├── rays_ar_ap_dashboard/             # AR/AP wizard addon
│   ├── wizard/                       # Excel report wizards
│   └── ...
├── rays_production_report/           # Production report addon
│   ├── models/                       # Production report wizard
│   └── ...
└── docs/
    ├── Rays-Odoo-New-Features-Documentation-v2.docx
    └── Rays-Odoo-New-Features-Presentation-v2.pptx
```

---

## 🚀 Installation

### Requirements
- Odoo 18 (Community or Enterprise)
- PostgreSQL 14+
- Python 3.10+
- `openpyxl` Python package

### Steps

1. **Clone this repository** into your Odoo addons directory:
   ```bash
   git clone https://github.com/imam-nurokhi/rays-odoo-customizations.git /path/to/odoo/addons/
   ```

2. **Install Python dependencies:**
   ```bash
   pip install openpyxl
   ```

3. **Update `odoo.conf`** to include the addons path:
   ```ini
   addons_path = /path/to/odoo/addons
   ```

4. **Restart Odoo** and activate developer mode:
   ```
   Settings → General Settings → Activate Developer Mode
   ```

5. **Install modules** via `Apps` menu:
   - Search for `raysolusi_ar_ap_dashboard`
   - Click Install

6. **Clear browser cache** and reload.

### Docker Deployment (Production)
```bash
# Add volume mount in docker-compose.yml
volumes:
  - ./addons:/mnt/extra-addons

# Restart container
docker restart odoo18

# Clear assets cache (psql)
DELETE FROM ir_attachment WHERE name LIKE '%.assets%' OR url LIKE '/web/assets/%';
```

---

## 📊 AR/AP Classification Logic

### Accounts Receivable (AR)
| Category | Document Type | Description |
|----------|--------------|-------------|
| PREMI | `out_invoice` | Insurance premium invoices |
| DISKON | `out_refund` | Credit notes / discounts |

### Accounts Payable (AP)
| Category | Account / Keyword | Description |
|----------|------------------|-------------|
| PREMI | Liability accounts | Insurance premium payables |
| DISKON Asuransi | `in_refund` | Refunds from insurance companies |
| BROKERAGE | Keywords: brokerage/komisi/commission | Broker fee payables |

### Production Report Keywords
| Category | Keywords |
|----------|---------|
| PREMI | premi, premium, insurance premium, hutang kepada insurance |
| DISKON | diskon, discount, brokerage discount |
| BROKERAGE | brokerage fee, brokerage, komisi, commission, broker fee |

---

## 🔧 Known Issues & Fixes

### `web_responsive` Icon Bug (Odoo 18)
**Error:** `TypeError: iconData.startsWith is not a function`  
**Location:** `web_responsive/static/src/components/apps_menu_tools.esm.js`  
**Fix:**
```js
// Change:
if (!menu.webIcon) return null;
// To:
if (!menu.webIcon || !iconData) return null;
```
> Apply this patch to **both** copies of `web_responsive` if you have multiple addon paths.

---

## 📖 Documentation

Full documentation is in `docs/`:
- **`Rays-Odoo-New-Features-Documentation-v2.docx`** — Step-by-step usage guide including:
  - AR Report guide
  - AP Report guide
  - Laporan Produksi guide
  - Batch Payment guide
  - Excel export instructions
- **`Rays-Odoo-New-Features-Presentation-v2.pptx`** — 10-slide executive presentation

---

## 🏢 Client

**PT. Raysolusi Pialang Asuransi**  
Insurance brokerage firm — Odoo 18 accounting implementation

---

## 👤 Author

**M. Imam Nurokhi**  
imam.nurokhi@cbqaglobal.com

---

## 📄 License

Proprietary — PT. Raysolusi Pialang Asuransi. All rights reserved.
