/** @odoo-module **/

import { Component, onMounted, onWillUnmount, useRef, useState } from "@odoo/owl";
import { registry } from "@web/core/registry";
import { useService } from "@web/core/utils/hooks";
import { rpc } from "@web/core/network/rpc";

// Chart.js is loaded as UMD before this module (via assets bundle)
const getChart = () => window.Chart;

class ArApDashboard extends Component {
    static template = "raysolusi_ar_ap_dashboard.Dashboard";
    static props = { "*": true };

    setup() {
        this.notification = useService("notification");

        // Canvas refs — dashboard charts
        this.arAgingRef  = useRef("arAgingChart");
        this.apAgingRef  = useRef("apAgingChart");
        this.monthlyRef  = useRef("monthlyChart");
        this.arStatusRef = useRef("arStatusChart");
        this.apStatusRef = useRef("apStatusChart");
        this.arCategoryRef      = useRef("arCategoryChart");
        this.apCategoryRef      = useRef("apCategoryChart");
        this.topCustomersArRef  = useRef("topCustomersArChart");
        this.topInsuranceApRef  = useRef("topInsuranceApChart");

        // Canvas refs — report charts
        this.arReportChartRef   = useRef("arReportChart");
        this.apReportChartRef   = useRef("apReportChart");
        this.prodReportChartRef = useRef("prodReportChart");

        this.state = useState({
            loading:     true,
            data:        null,
            lastUpdated: null,
            // Dashboard filters
            filterDateFrom:  '',
            filterDateTo:    '',
            filterPartner:   '',
            // Detail modal
            modal: {
                open:         false,
                loading:      false,
                title:        '',
                records:      [],
                total_count:  0,
                total_amount: 0,
            },
            // Tab management
            activeTab: 'dashboard',
            // AR Report
            arReport: {
                loading: false,
                data: null,
                filters: { dateFrom: '', dateTo: '', category: 'all', partnerName: '', paymentState: 'all' },
            },
            // AP Report
            apReport: {
                loading: false,
                data: null,
                filters: { dateFrom: '', dateTo: '', category: 'all', partnerName: '', paymentState: 'all' },
            },
            // Production Report
            production: {
                loading: false,
                data: null,
                filters: { dateFrom: '', dateTo: '', partnerName: '', asuransiName: '', moveType: 'all' },
            },
        });

        this._charts = {};

        onMounted(() => this.loadData());
        onWillUnmount(() => this._destroyCharts());
    }

    // ── Tab Management ──────────────────────────────────────────────────────

    setTab(tab) {
        if (this.state.activeTab === tab) return;
        this._destroyReportCharts();
        this.state.activeTab = tab;
        if (tab === 'ar_report' && !this.state.arReport.data) {
            setTimeout(() => this.loadArReport(), 50);
        } else if (tab === 'ap_report' && !this.state.apReport.data) {
            setTimeout(() => this.loadApReport(), 50);
        } else if (tab === 'production' && !this.state.production.data) {
            setTimeout(() => this.loadProductionReport(), 50);
        } else if (tab !== 'dashboard') {
            // Re-render chart when switching back to already-loaded tab
            const report = this.state[tab === 'ar_report' ? 'arReport' : tab === 'ap_report' ? 'apReport' : 'production'];
            if (report && report.data) {
                setTimeout(() => this._renderReportChart(tab, report.data), 100);
            }
        }
    }

    // ── AR Report ───────────────────────────────────────────────────────────

    async loadArReport() {
        this.state.arReport.loading = true;
        const f = this.state.arReport.filters;
        try {
            const data = await rpc("/raysolusi/ar_report/data", {
                date_from:     f.dateFrom     || null,
                date_to:       f.dateTo       || null,
                category:      f.category     || 'all',
                partner_name:  f.partnerName  || null,
                payment_state: f.paymentState || 'all',
            });
            this.state.arReport.data    = data;
            this.state.arReport.loading = false;
            setTimeout(() => this._renderReportChart('ar_report', data), 100);
        } catch (err) {
            this.state.arReport.loading = false;
            this.notification.add("Gagal memuat AR Report: " + (err.message || String(err)), { type: "danger" });
        }
    }

    applyArFilter() { this.state.arReport.data = null; this.loadArReport(); }
    resetArFilter() {
        this.state.arReport.filters = { dateFrom: '', dateTo: '', category: 'all', partnerName: '', paymentState: 'all' };
        this.state.arReport.data = null;
        this.loadArReport();
    }

    exportArExcel() {
        const f = this.state.arReport.filters;
        this._postExport('/raysolusi/ar_report/excel', {
            date_from:     f.dateFrom     || '',
            date_to:       f.dateTo       || '',
            category:      f.category     || 'all',
            partner_name:  f.partnerName  || '',
            payment_state: f.paymentState || 'all',
        });
    }

    // ── AP Report ───────────────────────────────────────────────────────────

    async loadApReport() {
        this.state.apReport.loading = true;
        const f = this.state.apReport.filters;
        try {
            const data = await rpc("/raysolusi/ap_report/data", {
                date_from:     f.dateFrom     || null,
                date_to:       f.dateTo       || null,
                category:      f.category     || 'all',
                partner_name:  f.partnerName  || null,
                payment_state: f.paymentState || 'all',
            });
            this.state.apReport.data    = data;
            this.state.apReport.loading = false;
            setTimeout(() => this._renderReportChart('ap_report', data), 100);
        } catch (err) {
            this.state.apReport.loading = false;
            this.notification.add("Gagal memuat AP Report: " + (err.message || String(err)), { type: "danger" });
        }
    }

    applyApFilter() { this.state.apReport.data = null; this.loadApReport(); }
    resetApFilter() {
        this.state.apReport.filters = { dateFrom: '', dateTo: '', category: 'all', partnerName: '', paymentState: 'all' };
        this.state.apReport.data = null;
        this.loadApReport();
    }

    exportApExcel() {
        const f = this.state.apReport.filters;
        this._postExport('/raysolusi/ap_report/excel', {
            date_from:     f.dateFrom     || '',
            date_to:       f.dateTo       || '',
            category:      f.category     || 'all',
            partner_name:  f.partnerName  || '',
            payment_state: f.paymentState || 'all',
        });
    }

    // ── Production Report ───────────────────────────────────────────────────

    async loadProductionReport() {
        this.state.production.loading = true;
        const f = this.state.production.filters;
        try {
            const data = await rpc("/raysolusi/production_report/data", {
                date_from:    f.dateFrom     || null,
                date_to:      f.dateTo       || null,
                partner_name: f.partnerName  || null,
                asuransi_name: f.asuransiName || null,
                move_type:    f.moveType     || 'all',
            });
            this.state.production.data    = data;
            this.state.production.loading = false;
            setTimeout(() => this._renderReportChart('production', data), 100);
        } catch (err) {
            this.state.production.loading = false;
            this.notification.add("Gagal memuat Laporan Produksi: " + (err.message || String(err)), { type: "danger" });
        }
    }

    applyProdFilter() { this.state.production.data = null; this.loadProductionReport(); }
    resetProdFilter() {
        this.state.production.filters = { dateFrom: '', dateTo: '', partnerName: '', asuransiName: '', moveType: 'all' };
        this.state.production.data = null;
        this.loadProductionReport();
    }

    exportProdExcel() {
        const f = this.state.production.filters;
        this._postExport('/raysolusi/production_report/excel', {
            date_from:     f.dateFrom     || '',
            date_to:       f.dateTo       || '',
            partner_name:  f.partnerName  || '',
            asuransi_name: f.asuransiName || '',
            move_type:     f.moveType     || 'all',
        });
    }

    // ── Excel POST helper ───────────────────────────────────────────────────

    _postExport(url, params) {
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = url;
        form.style.display = 'none';
        Object.entries(params).forEach(([k, v]) => {
            const inp = document.createElement('input');
            inp.type  = 'hidden';
            inp.name  = k;
            inp.value = v || '';
            form.appendChild(inp);
        });
        document.body.appendChild(form);
        form.submit();
        setTimeout(() => document.body.removeChild(form), 2000);
    }

    // ── Report Chart rendering ──────────────────────────────────────────────

    _renderReportChart(tab, data) {
        const Chart = getChart();
        if (!Chart || !data || !data.chart) return;

        const chart = data.chart;
        if (!chart.labels || !chart.labels.length) return;

        let ref, key, datasets;
        if (tab === 'ar_report') {
            ref = this.arReportChartRef;
            key = 'arReportChart';
            datasets = [
                { label: 'PREMI', data: chart.premi  || [], backgroundColor: 'rgba(44,95,138,0.8)',  borderColor: '#2c5f8a', borderWidth: 1, borderRadius: 4 },
                { label: 'DISKON', data: chart.diskon || [], backgroundColor: 'rgba(243,156,18,0.8)', borderColor: '#f39c12', borderWidth: 1, borderRadius: 4 },
            ];
        } else if (tab === 'ap_report') {
            ref = this.apReportChartRef;
            key = 'apReportChart';
            datasets = [
                { label: 'PREMI', data: chart.premi     || [], backgroundColor: 'rgba(39,174,96,0.8)',   borderColor: '#27ae60', borderWidth: 1, borderRadius: 4 },
                { label: 'BROKERAGE', data: chart.brokerage || [], backgroundColor: 'rgba(230,126,34,0.8)', borderColor: '#e67e22', borderWidth: 1, borderRadius: 4 },
            ];
        } else {
            ref = this.prodReportChartRef;
            key = 'prodReportChart';
            datasets = [
                { label: 'PREMI',     data: chart.premi     || [], backgroundColor: 'rgba(44,95,138,0.8)',   borderColor: '#2c5f8a', borderWidth: 1, borderRadius: 4 },
                { label: 'DISKON',    data: chart.diskon    || [], backgroundColor: 'rgba(243,156,18,0.8)', borderColor: '#f39c12', borderWidth: 1, borderRadius: 4 },
                { label: 'BROKERAGE', data: chart.brokerage || [], backgroundColor: 'rgba(39,174,96,0.8)',  borderColor: '#27ae60', borderWidth: 1, borderRadius: 4 },
            ];
        }

        if (this._charts[key]) {
            try { this._charts[key].destroy(); } catch (_) {}
            delete this._charts[key];
        }

        const el = ref && ref.el;
        if (!el) return;

        this._charts[key] = new Chart(el, {
            type: 'bar',
            data: { labels: chart.labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: true, position: 'top', labels: { font: { size: 11 }, usePointStyle: true } },
                    tooltip: {
                        callbacks: {
                            label: ctx => ' ' + ctx.dataset.label + ': Rp ' + Math.round(ctx.raw).toLocaleString('id-ID'),
                        },
                    },
                },
                scales: {
                    y: {
                        ticks: {
                            callback: v => {
                                if (v >= 1e9) return 'Rp ' + (v / 1e9).toFixed(1) + 'B';
                                if (v >= 1e6) return 'Rp ' + (v / 1e6).toFixed(0) + 'M';
                                return 'Rp ' + v.toLocaleString('id-ID');
                            },
                            font: { size: 10 },
                        },
                        grid: { color: 'rgba(0,0,0,0.04)' },
                    },
                    x: {
                        ticks: { font: { size: 9 }, maxRotation: 45 },
                        grid: { display: false },
                    },
                },
            },
        });
    }

    _destroyReportCharts() {
        ['arReportChart', 'apReportChart', 'prodReportChart'].forEach(k => {
            if (this._charts[k]) {
                try { this._charts[k].destroy(); } catch (_) {}
                delete this._charts[k];
            }
        });
    }

    // ── Public methods (called from template) ───────────────────────────────

    async loadData() {
        this.state.loading = true;
        this._destroyCharts();
        try {
            const data = await rpc("/raysolusi/ar_ap_dashboard/data", {
                date_from:    this.state.filterDateFrom  || null,
                date_to:      this.state.filterDateTo    || null,
                partner_name: this.state.filterPartner   || null,
            });
            this.state.data        = data;
            this.state.loading     = false;
            this.state.lastUpdated = new Date().toLocaleString("id-ID");
            setTimeout(() => this._renderAllCharts(), 100);
        } catch (err) {
            this.notification.add(
                "Gagal memuat data dashboard: " + (err.message || String(err)),
                { type: "danger" }
            );
            this.state.loading = false;
        }
    }

    applyFilter() { this.loadData(); }

    resetFilter() {
        this.state.filterDateFrom = '';
        this.state.filterDateTo   = '';
        this.state.filterPartner  = '';
        this.loadData();
    }

    closeModal() { this.state.modal.open = false; }

    async openDetail(moveType, filterType, filterValue) {
        this.state.modal.open         = true;
        this.state.modal.loading      = true;
        this.state.modal.records      = [];
        this.state.modal.title        = 'Memuat...';
        this.state.modal.total_count  = 0;
        this.state.modal.total_amount = 0;

        try {
            const result = await rpc("/raysolusi/ar_ap_dashboard/detail", {
                move_type:    moveType,
                filter_type:  filterType,
                filter_value: filterValue,
                date_from:    this.state.filterDateFrom || null,
                date_to:      this.state.filterDateTo   || null,
                partner_name: this.state.filterPartner  || null,
            });
            this.state.modal.loading      = false;
            this.state.modal.records      = result.records;
            this.state.modal.total_count  = result.total_count;
            this.state.modal.total_amount = result.total_amount;
            this.state.modal.title        = result.title;
        } catch (err) {
            this.state.modal.loading = false;
            this.state.modal.title   = 'Gagal memuat detail';
            this.notification.add(
                "Gagal memuat detail: " + (err.message || String(err)),
                { type: "danger" }
            );
        }
    }

    formatIDR(amount) {
        if (amount == null || isNaN(amount)) return "0";
        return Math.round(amount).toLocaleString("id-ID");
    }

    getPercent(amount, list) {
        const total = (list || []).reduce((s, r) => s + (r.amount || 0), 0);
        if (!total) return 0;
        return Math.round((amount / total) * 100);
    }

    getPaymentStateLabel(ps) {
        const m = {
            paid:       'Lunas',
            in_payment: 'Dalam Proses',
            not_paid:   'Belum Bayar',
            partial:    'Sebagian',
            reversed:   'Dibalik',
        };
        return m[ps] || ps;
    }

    getMoveTypeLabel(mt) {
        const m = {
            out_invoice: 'Invoice (RV)',
            in_invoice:  'Bill (PV)',
            out_refund:  'CN (RV)',
            in_refund:   'CN (PV)',
        };
        return m[mt] || mt;
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    _destroyCharts() {
        Object.values(this._charts).forEach(c => {
            try { c.destroy(); } catch (_) {}
        });
        this._charts = {};
    }

    // ── Chart rendering ──────────────────────────────────────────────────────

    _renderAllCharts() {
        const Chart = getChart();
        if (!Chart) {
            console.error("[ArApDashboard] Chart.js (window.Chart) is not available.");
            return;
        }
        const d = this.state.data;
        if (!d) return;

        this._renderAgingChart(
            "arAging", this.arAgingRef, d.ar_aging, "AR Piutang",
            ["rgba(44,95,138,0.85)", "rgba(241,196,15,0.85)",
             "rgba(230,126,34,0.85)", "rgba(231,76,60,0.85)", "rgba(192,57,43,0.85)"]
        );
        this._renderAgingChart(
            "apAging", this.apAgingRef, d.ap_aging, "AP Hutang",
            ["rgba(39,174,96,0.85)", "rgba(241,196,15,0.85)",
             "rgba(230,126,34,0.85)", "rgba(231,76,60,0.85)", "rgba(192,57,43,0.85)"]
        );
        this._renderMonthlyChart(d.monthly);
        this._renderStatusChart("arStatus", this.arStatusRef, d.payment_status.ar);
        this._renderStatusChart("apStatus", this.apStatusRef, d.payment_status.ap);
        if (d.ar_by_category) this._renderCategoryPieChart(
            "arCategory", this.arCategoryRef, d.ar_by_category,
            ["Premi (AR)", "Diskon (AR)"], ["#2c5f8a", "#f39c12"]);
        if (d.ap_by_category) this._renderCategoryPieChart(
            "apCategory", this.apCategoryRef, d.ap_by_category,
            ["Premi (AP)", "Brokerage (AP)"], ["#27ae60", "#e67e22"]);
        if (d.top_customers_ar && d.top_customers_ar.length)
            this._renderHorizontalBarChart(
                "topCustomersAr", this.topCustomersArRef,
                d.top_customers_ar, "Top 10 Customer AR", "#2c5f8a");
        if (d.top_insurance_ap && d.top_insurance_ap.length)
            this._renderHorizontalBarChart(
                "topInsuranceAp", this.topInsuranceApRef,
                d.top_insurance_ap, "Top 10 Asuransi AP", "#e67e22");
    }

    _renderAgingChart(key, ref, data, label, colors) {
        const Chart = getChart();
        const el    = ref.el;
        if (!el) return;

        const buckets = ['current', '1-30', '31-60', '61-90', '>90'];

        this._charts[key] = new Chart(el, {
            type: "bar",
            data: {
                labels:   data.labels,
                datasets: [{
                    label:           label,
                    data:            data.amounts,
                    backgroundColor: colors,
                    borderRadius:    6,
                    borderSkipped:   false,
                }],
            },
            options: {
                responsive:          true,
                maintainAspectRatio: false,
                onClick: (evt, elements) => {
                    if (!elements.length) return;
                    const idx      = elements[0].index;
                    const moveType = key.startsWith('ar') ? 'ar' : 'ap';
                    this.openDetail(moveType, 'aging', buckets[idx]);
                },
                onHover: (evt, elements) => {
                    evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: ctx => [
                                "  Jumlah  : Rp " + ctx.raw.toLocaleString("id-ID"),
                                "  Dokumen : " + data.counts[ctx.dataIndex],
                            ],
                        },
                    },
                },
                scales: {
                    y: {
                        ticks: {
                            callback: v => {
                                if (v >= 1e9) return (v / 1e9).toFixed(1) + "B";
                                if (v >= 1e6) return (v / 1e6).toFixed(0) + "M";
                                return v.toLocaleString("id-ID");
                            },
                            font: { size: 10 },
                        },
                        grid: { color: "rgba(0,0,0,0.04)" },
                    },
                    x: {
                        ticks: { font: { size: 10 } },
                        grid:  { display: false },
                    },
                },
            },
        });
    }

    _renderMonthlyChart(data) {
        const Chart = getChart();
        const el    = this.monthlyRef.el;
        if (!el) return;

        const monthMap = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
            'Mei': '05', 'Jun': '06', 'Jul': '07', 'Agt': '08',
            'Sep': '09', 'Okt': '10', 'Nov': '11', 'Des': '12',
        };

        this._charts.monthly = new Chart(el, {
            type: "line",
            data: {
                labels:   data.labels,
                datasets: [
                    {
                        label:                "AR (Piutang)",
                        data:                 data.ar,
                        borderColor:          "#2c5f8a",
                        backgroundColor:      "rgba(44,95,138,0.08)",
                        borderWidth:          2.5,
                        pointRadius:          4,
                        pointBackgroundColor: "#2c5f8a",
                        fill:                 true,
                        tension:              0.35,
                    },
                    {
                        label:                "AP (Hutang)",
                        data:                 data.ap,
                        borderColor:          "#e67e22",
                        backgroundColor:      "rgba(230,126,34,0.08)",
                        borderWidth:          2.5,
                        pointRadius:          4,
                        pointBackgroundColor: "#e67e22",
                        fill:                 true,
                        tension:              0.35,
                    },
                ],
            },
            options: {
                responsive:          true,
                maintainAspectRatio: false,
                interaction: { mode: "index", intersect: false },
                onClick: (evt, elements) => {
                    if (!elements.length) return;
                    const idx   = elements[0].index;
                    const label = data.labels[idx];
                    const parts = label.split(' ');
                    const monthYear = parts[1] + '-' + (monthMap[parts[0]] || '01');
                    const moveType = elements[0].datasetIndex === 0 ? 'ar' : 'ap';
                    this.openDetail(moveType, 'month', monthYear);
                },
                onHover: (evt, elements) => {
                    evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
                },
                plugins: {
                    legend: {
                        display:  true,
                        position: "top",
                        labels:   { font: { size: 11 }, usePointStyle: true },
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx =>
                                " " + ctx.dataset.label + ": Rp " +
                                ctx.raw.toLocaleString("id-ID"),
                        },
                    },
                },
                scales: {
                    y: {
                        ticks: {
                            callback: v => {
                                if (v >= 1e9) return "Rp " + (v / 1e9).toFixed(1) + "B";
                                if (v >= 1e6) return "Rp " + (v / 1e6).toFixed(0) + "M";
                                return "Rp " + v.toLocaleString("id-ID");
                            },
                            font: { size: 10 },
                        },
                        grid: { color: "rgba(0,0,0,0.04)" },
                    },
                    x: {
                        ticks: { font: { size: 10 } },
                        grid:  { display: false },
                    },
                },
            },
        });
    }

    _renderStatusChart(key, ref, statusData) {
        const Chart = getChart();
        const el    = ref.el;
        if (!el) return;

        const STATUS_MAP = {
            paid:       { label: "Lunas",        color: "#27ae60" },
            in_payment: { label: "Dalam Proses",  color: "#3498db" },
            not_paid:   { label: "Belum Bayar",   color: "#e74c3c" },
            partial:    { label: "Sebagian",      color: "#f39c12" },
            reversed:   { label: "Dibalik",       color: "#95a5a6" },
        };

        const labels     = [];
        const values     = [];
        const colors     = [];
        const statusKeys = [];

        Object.entries(statusData).forEach(([k, v]) => {
            if (v > 0 && STATUS_MAP[k]) {
                labels.push(STATUS_MAP[k].label);
                values.push(v);
                colors.push(STATUS_MAP[k].color);
                statusKeys.push(k);
            }
        });

        this._charts[key] = new Chart(el, {
            type: "doughnut",
            data: {
                labels,
                datasets: [{
                    data:             values,
                    backgroundColor:  colors,
                    borderWidth:      2,
                    borderColor:      "#fff",
                    hoverBorderWidth: 3,
                }],
            },
            options: {
                responsive:          true,
                maintainAspectRatio: false,
                cutout:              "62%",
                onClick: (evt, elements) => {
                    if (!elements.length) return;
                    const idx       = elements[0].index;
                    const statusKey = statusKeys[idx];
                    const moveType  = key.startsWith('ar') ? 'ar' : 'ap';
                    this.openDetail(moveType, 'status', statusKey);
                },
                onHover: (evt, elements) => {
                    evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
                },
                plugins: {
                    legend: {
                        display:  true,
                        position: "right",
                        labels: {
                            font:          { size: 11 },
                            padding:       10,
                            usePointStyle: true,
                            generateLabels: chart => {
                                const ds = chart.data.datasets[0];
                                return chart.data.labels.map((lbl, i) => ({
                                    text:        lbl + " (" + ds.data[i] + ")",
                                    fillStyle:   ds.backgroundColor[i],
                                    strokeStyle: "#fff",
                                    hidden:      false,
                                    index:       i,
                                }));
                            },
                        },
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => " " + ctx.label + ": " + ctx.raw + " transaksi",
                        },
                    },
                },
            },
        });
    }

    _renderCategoryPieChart(key, ref, data, labels, colors) {
        const Chart = getChart();
        const el    = ref.el;
        const values = Object.values(data);
        const filteredLabels = [];
        const filteredValues = [];
        const filteredColors = [];
        labels.forEach((lbl, i) => {
            if (values[i] > 0) {
                filteredLabels.push(lbl);
                filteredValues.push(values[i]);
                filteredColors.push(colors[i]);
            }
        });
        this._charts[key] = new Chart(el, {
            type: "doughnut",
            data: {
                labels:   filteredLabels,
                datasets: [{
                    data:             filteredValues,
                    backgroundColor:  filteredColors,
                    borderWidth:      2,
                    borderColor:      "#fff",
                    hoverBorderWidth: 3,
                }],
            },
            options: {
                responsive:          true,
                maintainAspectRatio: false,
                cutout:              "55%",
                plugins: {
                    legend: {
                        display:  true,
                        position: "right",
                        labels: {
                            font:          { size: 11 },
                            padding:       10,
                            usePointStyle: true,
                            generateLabels: chart => {
                                const ds = chart.data.datasets[0];
                                return chart.data.labels.map((lbl, i) => ({
                                    text:        lbl + ": Rp " + Math.round(ds.data[i]).toLocaleString("id-ID"),
                                    fillStyle:   ds.backgroundColor[i],
                                    strokeStyle: "#fff",
                                    hidden:      false,
                                    index:       i,
                                }));
                            },
                        },
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => " " + ctx.label + ": Rp " + Math.round(ctx.raw).toLocaleString("id-ID"),
                        },
                    },
                },
            },
        });
    }

    _renderHorizontalBarChart(key, ref, dataList, title, color) {
        const Chart = getChart();
        const el    = ref.el;
        const labels  = dataList.map(r => r[0]);
        const amounts = dataList.map(r => r[1]);
        this._charts[key] = new Chart(el, {
            type: "bar",
            data: {
                labels,
                datasets: [{
                    label:           title,
                    data:            amounts,
                    backgroundColor: color + "cc",
                    borderColor:     color,
                    borderWidth:     1,
                    borderRadius:    4,
                    borderSkipped:   false,
                }],
            },
            options: {
                responsive:          true,
                maintainAspectRatio: false,
                indexAxis:           "y",
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: ctx => "  Rp " + Math.round(ctx.raw).toLocaleString("id-ID"),
                        },
                    },
                },
                scales: {
                    x: {
                        ticks: {
                            callback: v => {
                                if (v >= 1e9) return (v / 1e9).toFixed(1) + "B";
                                if (v >= 1e6) return (v / 1e6).toFixed(0) + "M";
                                return v.toLocaleString("id-ID");
                            },
                            font: { size: 10 },
                        },
                        grid: { color: "rgba(0,0,0,0.04)" },
                    },
                    y: {
                        ticks: { font: { size: 10 } },
                        grid:  { display: false },
                    },
                },
            },
        });
    }
}

registry.category("actions").add("raysolusi_ar_ap_dashboard", ArApDashboard);
