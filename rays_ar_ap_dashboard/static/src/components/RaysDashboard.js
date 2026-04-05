/** @odoo-module **/
import { Component, useState, onMounted, onWillUnmount } from "@odoo/owl";
import { registry } from "@web/core/registry";
import { rpc } from "@web/core/network/rpc";

export class RaysDashboard extends Component {
    static template = "rays_ar_ap_dashboard.RaysDashboard";
    static props = { "*": true };

    setup() {
        this.state = useState({
            loading: true,
            error: null,
            kpi: { total_ar: 0, total_ap: 0, net_position: 0, overdue_ar_count: 0 },
            dateFrom: "",
            dateTo: "",
        });
        this._charts = {};
        onMounted(() => this.loadData());
        onWillUnmount(() => this._destroyCharts());
    }

    async loadData() {
        this.state.loading = true;
        this.state.error = null;
        this._destroyCharts();
        try {
            const data = await rpc("/rays/dashboard_data", {
                date_from: this.state.dateFrom || false,
                date_to: this.state.dateTo || false,
            });
            this.state.kpi = data.kpi;
            this.state.loading = false;
            setTimeout(() => this._renderCharts(data), 80);
        } catch (e) {
            console.error("[RaysDashboard] loadData error:", e);
            this.state.error = "Gagal memuat data. Silakan refresh halaman.";
            this.state.loading = false;
        }
    }

    _destroyCharts() {
        Object.values(this._charts).forEach((c) => { if (c) c.destroy(); });
        this._charts = {};
    }

    _renderCharts(data) {
        const C = window.Chart;
        if (!C) { console.warn("[RaysDashboard] Chart.js not loaded"); return; }
        this._charts.arCat = this._pie("ar-category-chart",
            Object.keys(data.ar_by_category), Object.values(data.ar_by_category), "AR by Kategori");
        this._charts.apCat = this._pie("ap-category-chart",
            Object.keys(data.ap_by_category), Object.values(data.ap_by_category), "AP by Kategori");
        this._charts.arTrend = this._line("monthly-ar-chart", data.monthly_ar, "Trend AR Bulanan", "#1F4E79");
        this._charts.apTrend = this._line("monthly-ap-chart", data.monthly_ap, "Trend AP Bulanan", "#A23B72");
        this._charts.topAR = this._bar("top-customers-chart", data.top_customers_ar, "Top 10 Customer AR", "#2E86AB");
        this._charts.topAP = this._bar("top-insurance-chart", data.top_insurance_ap, "Top 10 Asuransi AP", "#F18F01");
    }

    _pie(id, labels, values, title) {
        const ctx = document.getElementById(id);
        if (!ctx) return null;
        return new window.Chart(ctx, {
            type: "pie",
            data: { labels, datasets: [{ data: values, backgroundColor: ["#1F4E79", "#2E86AB", "#A23B72", "#F18F01"], borderWidth: 2 }] },
            options: { responsive: true, plugins: { title: { display: true, text: title, font: { size: 14, weight: "bold" } }, legend: { position: "bottom" } } },
        });
    }

    _line(id, monthlyData, title, color) {
        const ctx = document.getElementById(id);
        if (!ctx) return null;
        return new window.Chart(ctx, {
            type: "line",
            data: { labels: Object.keys(monthlyData), datasets: [{ label: title, data: Object.values(monthlyData), borderColor: color, backgroundColor: color + "22", tension: 0.3, fill: true, pointRadius: 4 }] },
            options: { responsive: true, plugins: { title: { display: true, text: title, font: { size: 14, weight: "bold" } }, legend: { display: false } }, scales: { y: { ticks: { callback: (v) => this._short(v) } } } },
        });
    }

    _bar(id, partnerData, title, color) {
        const ctx = document.getElementById(id);
        if (!ctx) return null;
        const labels = partnerData.map((p) => (p[0].length > 22 ? p[0].substr(0, 22) + "\u2026" : p[0]));
        return new window.Chart(ctx, {
            type: "bar",
            data: { labels, datasets: [{ label: "IDR", data: partnerData.map((p) => p[1]), backgroundColor: color, borderRadius: 4 }] },
            options: { indexAxis: "y", responsive: true, plugins: { title: { display: true, text: title, font: { size: 14, weight: "bold" } }, legend: { display: false } }, scales: { x: { ticks: { callback: (v) => this._short(v) } } } },
        });
    }

    _short(v) {
        const a = Math.abs(v);
        if (a >= 1e9) return (v / 1e9).toFixed(1) + "B";
        if (a >= 1e6) return (v / 1e6).toFixed(1) + "M";
        if (a >= 1e3) return (v / 1e3).toFixed(0) + "K";
        return v;
    }

    formatCurrency(value) {
        return new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(value || 0);
    }

    async onFilterChange() {
        await this.loadData();
    }
}

registry.category("actions").add("rays_ar_ap_dashboard", RaysDashboard);
