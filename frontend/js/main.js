/**
 * frontend/js/main.js — Dashboard SK Frontend Controller
 *
 * Berkomunikasi dengan backend Google Apps Script via Vercel Proxy:
 *   GET /api/proxy?action=getDashboardData
 *   GET /api/proxy?action=getYearlyData&targetYear=YYYY
 *
 * Seluruh fitur lama dipertahankan:
 *   - 5 KPI card dashboard + 2 average card
 *   - 3 chart dashboard (Jenis Keputusan, Status, Tren Tahunan)
 *   - Grid OPD pengusul (dashboard)
 *   - Halaman Data Tahunan (KPI + 3 chart + grid OPD + year picker)
 *   - Tombol Rekap Data (link ke Spreadsheet)
 *   - Reveal animation
 */

(function () {
    'use strict';

    // ═══════════════════════════════════════════════════════════
    //  KONFIGURASI
    // ═══════════════════════════════════════════════════════════
    const API_BASE = '/api/proxy';

    const COLORS = {
        gubernur: '#4f46e5',
        sekda: '#0891b2',
        done: '#059669',
        process: '#d97706',
        pending: '#0284c7',
        rejected: '#e11d48',
        trend: '#4f46e5',
        trendFill: 'rgba(79, 70, 229, 0.12)',
        monthly: '#0891b2',
        palette: ['#4f46e5', '#0891b2', '#7c3aed', '#db2777', '#059669', '#d97706'],
    };

    if (window.Chart) {
        Chart.defaults.font.family = 'Inter, sans-serif';
        Chart.defaults.color = '#475569';
        Chart.defaults.borderColor = 'rgba(203, 213, 225, 0.5)';
    }

    // ═══════════════════════════════════════════════════════════
    //  STATE
    // ═══════════════════════════════════════════════════════════
    const state = {
        dashboardData: null,
        yearlyData: null,
        availableYears: [],
        currentYear: null,
        view: 'dashboard',
    };

    const charts = {
        decision: null,
        status: null,
        trend: null,
        yearlyDecision: null,
        yearlyStatus: null,
        yearlyTrend: null,
    };

    let errorToastTimer = null;

    // ═══════════════════════════════════════════════════════════
    //  API HELPER
    // ═══════════════════════════════════════════════════════════
    async function callApi(action, params) {
        const qs = new URLSearchParams(Object.assign({ action }, params || {}));
        const url = API_BASE + '?' + qs.toString();

        const res = await fetch(url, {
            method: 'GET',
            headers: { Accept: 'application/json' },
            cache: 'no-store',
        });

        if (!res.ok) {
            const text = await res.text().catch(() => '');
            throw new Error('HTTP ' + res.status + ' — ' + (text.slice(0, 200) || res.statusText));
        }

        const data = await res.json();
        if (data && data.error) {
            throw new Error(data.error);
        }
        return data;
    }

    // ═══════════════════════════════════════════════════════════
    //  FORMATTING HELPERS
    // ═══════════════════════════════════════════════════════════
    function formatNumber(n) {
        const num = Number(n);
        return isNaN(num) ? '0' : num.toLocaleString('id-ID');
    }

    function pct(part, total) {
        if (!total || total <= 0) return '0%';
        return Math.round((part / total) * 100) + '%';
    }

    function escapeHtml(str) {
        return String(str || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function setText(id, value) {
        const el = document.getElementById(id);
        if (el) el.textContent = value;
    }

    // ═══════════════════════════════════════════════════════════
    //  VIEW TOGGLING
    // ═══════════════════════════════════════════════════════════
    function showLoading() {
        const loading = document.getElementById('loadingState');
        const dash = document.getElementById('dashboardContent');
        const yearly = document.getElementById('yearlyContent');
        if (loading) loading.classList.remove('hidden');
        if (dash) dash.classList.add('hidden');
        if (yearly) yearly.classList.add('hidden');
    }

    function showView(view) {
        state.view = view;
        const loading = document.getElementById('loadingState');
        const dash = document.getElementById('dashboardContent');
        const yearly = document.getElementById('yearlyContent');

        if (loading) loading.classList.add('hidden');
        if (dash) dash.classList.toggle('hidden', view !== 'dashboard');
        if (yearly) yearly.classList.toggle('hidden', view !== 'yearly');

        window.scrollTo({ top: 0, behavior: 'smooth' });
        syncYearSelectors();
        activateReveal();
    }

    // ═══════════════════════════════════════════════════════════
    //  ERROR TOAST
    // ═══════════════════════════════════════════════════════════
    function showError(message) {
        let toast = document.getElementById('appErrorToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'appErrorToast';
            toast.className =
                'fixed top-24 left-1/2 -translate-x-1/2 z-[100] px-5 py-3 bg-rose-50 border border-rose-300 ' +
                'text-rose-800 rounded-xl shadow-lg font-bold text-sm max-w-md hidden';
            document.body.appendChild(toast);
        }
        toast.innerHTML =
            '<div class="flex items-start gap-2"><i class="ph ph-warning-circle text-lg mt-0.5"></i>' +
            '<span>' + escapeHtml(message) + '</span></div>';
        toast.classList.remove('hidden');
        if (errorToastTimer) clearTimeout(errorToastTimer);
        errorToastTimer = setTimeout(function () {
            toast.classList.add('hidden');
        }, 6500);
    }

    // ═══════════════════════════════════════════════════════════
    //  REVEAL ANIMATION
    // ═══════════════════════════════════════════════════════════
    let revealObserver = null;

    function activateReveal() {
        if (!('IntersectionObserver' in window)) {
            document.querySelectorAll('.reveal').forEach(function (el) {
                el.classList.add('active');
            });
            return;
        }

        if (!revealObserver) {
            revealObserver = new IntersectionObserver(
                function (entries) {
                    entries.forEach(function (entry) {
                        if (entry.isIntersecting) {
                            entry.target.classList.add('active');
                            revealObserver.unobserve(entry.target);
                        }
                    });
                },
                { threshold: 0.08 }
            );
        }

        document.querySelectorAll('.reveal:not(.active)').forEach(function (el) {
            revealObserver.observe(el);
        });
    }

    // ═══════════════════════════════════════════════════════════
    //  YEAR PICKERS
    // ═══════════════════════════════════════════════════════════
    function populateYearPicker(years) {
        state.availableYears = years || [];

        // Navbar dropdown menu
        const menu = document.getElementById('yearPickerMenu');
        if (menu) {
            menu.innerHTML = '';
            if (state.availableYears.length === 0) {
                const empty = document.createElement('div');
                empty.className = 'px-3 py-2 text-xs text-slate-500 text-center';
                empty.textContent = 'Belum ada data tahun';
                menu.appendChild(empty);
            } else {
                state.availableYears.forEach(function (y) {
                    const btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className =
                        'w-full text-left px-3 py-2 rounded-lg text-sm font-bold text-slate-700 ' +
                        'hover:bg-indigo-50 hover:text-indigo-700 transition-colors flex items-center justify-between';
                    btn.innerHTML =
                        '<span>' + y + '</span>' +
                        '<i class="ph ph-arrow-right text-xs opacity-60"></i>';
                    btn.addEventListener('click', function () {
                        closeYearMenu();
                        loadYearlyData(y);
                    });
                    menu.appendChild(btn);
                });
            }
        }

        // Inline select (yearly view)
        const select = document.getElementById('yearlyYearSelect');
        if (select) {
            select.innerHTML = '';
            if (state.availableYears.length === 0) {
                const opt = document.createElement('option');
                opt.textContent = 'Tidak ada tahun';
                opt.disabled = true;
                opt.selected = true;
                select.appendChild(opt);
            } else {
                state.availableYears.forEach(function (y) {
                    const opt = document.createElement('option');
                    opt.value = String(y);
                    opt.textContent = String(y);
                    select.appendChild(opt);
                });
            }
        }
    }

    function syncYearSelectors() {
        const select = document.getElementById('yearlyYearSelect');
        if (select && state.currentYear != null) {
            select.value = String(state.currentYear);
        }

        const btnText = document.getElementById('pickYearBtnText');
        if (btnText) {
            btnText.textContent =
                state.view === 'yearly' && state.currentYear != null
                    ? 'Tahun ' + state.currentYear
                    : 'Pick Year';
        }
    }

    function closeYearMenu() {
        const menu = document.getElementById('yearPickerMenu');
        if (menu) menu.classList.add('hidden');
    }

    // ═══════════════════════════════════════════════════════════
    //  CHART RENDERERS
    // ═══════════════════════════════════════════════════════════
    function destroyChart(key) {
        if (charts[key]) {
            try { charts[key].destroy(); } catch (e) { /* noop */ }
            charts[key] = null;
        }
    }

    function renderDecisionChart(canvasId, decisionTypes, key) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const labels = (decisionTypes && decisionTypes.labels) || [];
        const values = (decisionTypes && decisionTypes.values) || [];

        destroyChart(key);

        const hasData = labels.length > 0 && values.some(function (v) { return v > 0; });

        charts[key] = new Chart(canvas.getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: hasData ? labels : ['Belum ada data'],
                datasets: [
                    {
                        data: hasData ? values : [1],
                        backgroundColor: hasData ? COLORS.palette : ['#cbd5e1'],
                        borderColor: '#ffffff',
                        borderWidth: 3,
                        hoverOffset: 6,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '60%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            font: { family: 'Inter', size: 12, weight: '600' },
                            color: '#0f172a',
                            padding: 14,
                            usePointStyle: true,
                            pointStyle: 'circle',
                            boxWidth: 10,
                            boxHeight: 10,
                        },
                    },
                    tooltip: {
                        callbacks: {
                            label: function (ctx) {
                                if (!hasData) return 'Belum ada data';
                                const total = ctx.dataset.data.reduce(function (a, b) { return a + b; }, 0);
                                const p = total > 0 ? Math.round((ctx.parsed / total) * 100) : 0;
                                return ctx.label + ': ' + formatNumber(ctx.parsed) + ' (' + p + '%)';
                            },
                        },
                    },
                },
            },
        });
    }

    function renderStatusChart(canvasId, statusMonitoring, key) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const sm = statusMonitoring || {};
        const labels = ['Completed', 'On Progress', 'Pending', 'Rejected'];
        const data = [sm.completed || 0, sm.onProgress || 0, sm.pending || 0, sm.rejected || 0];
        const colors = [COLORS.done, COLORS.process, COLORS.pending, COLORS.rejected];
        const hasData = data.some(function (v) { return v > 0; });

        destroyChart(key);

        charts[key] = new Chart(canvas.getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [
                    {
                        data: hasData ? data : [1],
                        backgroundColor: hasData ? colors : ['#cbd5e1'],
                        borderColor: '#ffffff',
                        borderWidth: 3,
                        hoverOffset: 6,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '60%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            font: { family: 'Inter', size: 12, weight: '600' },
                            color: '#0f172a',
                            padding: 14,
                            usePointStyle: true,
                            pointStyle: 'circle',
                            boxWidth: 10,
                            boxHeight: 10,
                        },
                    },
                    tooltip: {
                        callbacks: {
                            label: function (ctx) {
                                if (!hasData) return 'Belum ada data';
                                const total = ctx.dataset.data.reduce(function (a, b) { return a + b; }, 0);
                                const p = total > 0 ? Math.round((ctx.parsed / total) * 100) : 0;
                                return ctx.label + ': ' + formatNumber(ctx.parsed) + ' (' + p + '%)';
                            },
                        },
                    },
                },
            },
        });
    }

    function renderYearlyTrendChart(canvasId, monthlyTrend, key) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const labels = (monthlyTrend && monthlyTrend.labels) || [];
        const values = (monthlyTrend && monthlyTrend.values) || [];

        destroyChart(key);

        charts[key] = new Chart(canvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Jumlah Berkas',
                        data: values,
                        backgroundColor: 'rgba(79, 70, 229, 0.85)',
                        hoverBackgroundColor: '#4f46e5',
                        borderRadius: 8,
                        borderSkipped: false,
                        maxBarThickness: 48,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function (ctx) {
                                return 'Jumlah: ' + formatNumber(ctx.parsed.y) + ' berkas';
                            },
                        },
                    },
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: {
                            font: { family: 'Inter', size: 12, weight: '600' },
                            color: '#0f172a',
                        },
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(203, 213, 225, 0.5)' },
                        ticks: {
                            font: { family: 'Inter', size: 11, weight: '500' },
                            color: '#475569',
                            precision: 0,
                        },
                    },
                },
            },
        });
    }

    function renderMonthlyBarChart(canvasId, monthlyTrend, key) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const labels = (monthlyTrend && monthlyTrend.labels) || [];
        const values = (monthlyTrend && monthlyTrend.values) || [];

        destroyChart(key);

        charts[key] = new Chart(canvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Jumlah Berkas',
                        data: values,
                        backgroundColor: 'rgba(8, 145, 178, 0.85)',
                        hoverBackgroundColor: COLORS.monthly,
                        borderRadius: 8,
                        borderSkipped: false,
                        maxBarThickness: 42,
                    },
                ],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function (ctx) {
                                return ctx.label + ': ' + formatNumber(ctx.parsed.y) + ' berkas';
                            },
                        },
                    },
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: {
                            font: { family: 'Inter', size: 11, weight: '600' },
                            color: '#0f172a',
                        },
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(203, 213, 225, 0.5)' },
                        ticks: {
                            font: { family: 'Inter', size: 11, weight: '500' },
                            color: '#475569',
                            precision: 0,
                        },
                    },
                },
            },
        });
    }

    // ═══════════════════════════════════════════════════════════
    //  OPD GRID RENDERER
    // ═══════════════════════════════════════════════════════════
    function renderOpdGrid(containerId, departmentRanking) {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = '';

        const labels = (departmentRanking && departmentRanking.labels) || [];
        const gubernur = (departmentRanking && departmentRanking.gubernur) || [];
        const sekda = (departmentRanking && departmentRanking.sekda) || [];

        if (labels.length === 0) {
            container.innerHTML =
                '<div class="col-span-full text-center py-12 text-slate-500 font-semibold">' +
                '<i class="ph ph-folder-open text-3xl mb-2 inline-block opacity-60"></i>' +
                '<p>Belum ada data perangkat daerah</p>' +
                '</div>';
            return;
        }

        labels.forEach(function (name, i) {
            const g = gubernur[i] || 0;
            const s = sekda[i] || 0;
            const total = g + s;

            const rankClass =
                i === 0 ? 'opd-rank-1' : i === 1 ? 'opd-rank-2' : i === 2 ? 'opd-rank-3' : '';
            const rankLabel = i === 0 ? 'Top 1' : i === 1 ? 'Top 2' : i === 2 ? 'Top 3' : '#' + (i + 1);

            const box = document.createElement('div');
            box.className = 'opd-box ' + rankClass;
            box.innerHTML =
                '<div class="opd-header">' +
                    '<span class="opd-rank">' + rankLabel + '</span>' +
                    '<span class="opd-total-badge">' + formatNumber(total) + '</span>' +
                '</div>' +
                '<p class="opd-name" title="' + escapeHtml(name) + '">' + escapeHtml(name) + '</p>' +
                '<div class="opd-counts">' +
                    '<span class="opd-count-gov">' +
                        '<span class="opd-dot opd-dot-gov"></span>Gubernur: ' + formatNumber(g) +
                    '</span>' +
                    '<span class="opd-count-sek">' +
                        '<span class="opd-dot opd-dot-sek"></span>Sekda: ' + formatNumber(s) +
                    '</span>' +
                '</div>';

            container.appendChild(box);
        });
    }

    // ═══════════════════════════════════════════════════════════
    //  DASHBOARD RENDER
    // ═══════════════════════════════════════════════════════════
    function renderDashboard(data) {
        const a = data.analytics || {};
        const sc = a.statusCounts || {};
        const total = sc.total || 0;

        // Header
        setText('lastUpdated', data.lastUpdate || '—');

        // KPI
        setText('kpiReceived', formatNumber(total));
        setText('kpiReceivedPct', '100%');

        setText('kpiDone', formatNumber(sc.done || 0));
        setText('kpiDonePct', pct(sc.done, total));

        setText('kpiProcess', formatNumber(sc.proc || 0));
        setText('kpiProcessPct', pct(sc.proc, total));

        setText('kpiPending', formatNumber(sc.pend || 0));
        setText('kpiPendingPct', pct(sc.pend, total));

        setText('kpiRejected', formatNumber(sc.rej || 0));
        setText('kpiRejectedPct', pct(sc.rej, total));

        // Average SK (per tahun)
        const avgSk = (a.monthlyIncoming && a.monthlyIncoming.average) || 0;
        setText('kpiAvgSk', formatNumber(avgSk));

        const trendLabels = (a.monthlyTrend && a.monthlyTrend.labels) || [];
        const avgSub =
            trendLabels.length > 0
                ? 'Rata-rata dari ' + trendLabels.length + ' tahun data (' +
                    trendLabels[0] + ' – ' + trendLabels[trendLabels.length - 1] + ')'
                : 'Belum ada data tren tahunan';
        setText('kpiAvgSkSub', avgSub);

        // Average processing time
        const avgProc = (a.processingTime && a.processingTime.average) || 0;
        const sample = (a.processingTime && a.processingTime.sampleSize) || 0;
        setText('kpiAvgWaktu', formatNumber(avgProc));
        setText(
            'kpiAvgWaktuSub',
            sample > 0 ? 'Berdasarkan ' + formatNumber(sample) + ' SK selesai' : 'Belum ada SK selesai'
        );

        // Rekap link
        const rekap = document.getElementById('rekapDataLink');
        if (rekap && data.spreadsheetUrl) rekap.href = data.spreadsheetUrl;

        // Years
        populateYearPicker(data.availableYears || []);

        // Show view FIRST (agar chart mendapat dimensi yang benar)
        showView('dashboard');

        // Charts
        renderDecisionChart('decisionChart', a.decisionTypes, 'decision');
        renderStatusChart('statusChart', a.statusMonitoring, 'status');
        renderYearlyTrendChart('trendChart', a.monthlyTrend, 'trend');

        // OPD
        renderOpdGrid('deptGrid', a.departmentRanking);
    }

    // ═══════════════════════════════════════════════════════════
    //  YEARLY RENDER
    // ═══════════════════════════════════════════════════════════
    function renderYearly(data) {
        const a = data.analytics || {};
        const sc = a.statusCounts || {};
        const total = sc.total || 0;

        // State
        state.yearlyData = data;
        state.currentYear = data.targetYear != null ? data.targetYear : null;

        // Header
        setText('yearlySubtitleYear', data.targetYear || '—');
        setText('yearlyLastUpdated', data.lastUpdate || '—');

        // KPI
        setText('yearlyKpiReceived', formatNumber(total));
        setText('yearlyKpiReceivedPct', '100%');

        setText('yearlyKpiDone', formatNumber(sc.done || 0));
        setText('yearlyKpiDonePct', pct(sc.done, total));

        setText('yearlyKpiProcess', formatNumber(sc.proc || 0));
        setText('yearlyKpiProcessPct', pct(sc.proc, total));

        setText('yearlyKpiPending', formatNumber(sc.pend || 0));
        setText('yearlyKpiPendingPct', pct(sc.pend, total));

        setText('yearlyKpiRejected', formatNumber(sc.rej || 0));
        setText('yearlyKpiRejectedPct', pct(sc.rej, total));

        // Average per bulan
        const avgSk = (a.monthlyIncoming && a.monthlyIncoming.average) || 0;
        setText('yearlyKpiAvgSk', formatNumber(avgSk));
        const period = (a.monthlyIncoming && a.monthlyIncoming.period) || 'Januari – Desember';
        setText(
            'yearlyKpiAvgSkSub',
            formatNumber(total) + ' SK dalam 12 bulan · ' + period
        );

        // Average processing
        const avgProc = (a.processingTime && a.processingTime.average) || 0;
        const sample = (a.processingTime && a.processingTime.sampleSize) || 0;
        setText('yearlyKpiAvgWaktu', formatNumber(avgProc));
        setText(
            'yearlyKpiAvgWaktuSub',
            sample > 0 ? 'Berdasarkan ' + formatNumber(sample) + ' SK selesai' : 'Belum ada SK selesai'
        );

        // Rekap link
        const rekap = document.getElementById('rekapDataLink');
        if (rekap && data.spreadsheetUrl) rekap.href = data.spreadsheetUrl;

        // Refresh year options if backend provided new list
        if (Array.isArray(data.availableYears) && data.availableYears.length > 0) {
            populateYearPicker(data.availableYears);
        }

        // Show view FIRST
        showView('yearly');

        // Charts
        renderDecisionChart('yearlyDecisionChart', a.decisionTypes, 'yearlyDecision');
        renderStatusChart('yearlyStatusChart', a.statusMonitoring, 'yearlyStatus');
        renderMonthlyBarChart('yearlyTrendChart', a.monthlyTrend, 'yearlyTrend');

        // OPD
        renderOpdGrid('yearlyDeptGrid', a.departmentRanking);
    }

    // ═══════════════════════════════════════════════════════════
    //  LOADERS
    // ═══════════════════════════════════════════════════════════
    async function loadDashboardData() {
        try {
            showLoading();
            const data = await callApi('getDashboardData');
            state.dashboardData = data;
            renderDashboard(data);
        } catch (err) {
            console.error('[Dashboard]', err);
            showError('Gagal memuat dashboard: ' + err.message);
            // Tampilkan dashboard kosong agar UI tidak "nyangkut" di loading
            renderDashboard({
                analytics: {
                    statusCounts: { total: 0, done: 0, proc: 0, pend: 0, rej: 0 },
                    statusMonitoring: { completed: 0, onProgress: 0, pending: 0, rejected: 0 },
                    decisionTypes: { labels: [], values: [], total: 0 },
                    monthlyTrend: { labels: [], values: [] },
                    departmentRanking: { labels: [], gubernur: [], sekda: [] },
                    processingTime: { average: 0, sampleSize: 0 },
                    monthlyIncoming: { average: 0 },
                },
                lastUpdate: '',
                availableYears: [],
                spreadsheetUrl: '',
                error: null,
            });
        }
    }

    async function loadYearlyData(year) {
        try {
            showLoading();
            const data = await callApi('getYearlyData', { targetYear: year });
            state.yearlyData = data;
            renderYearly(data);
        } catch (err) {
            console.error('[Yearly]', err);
            showError('Gagal memuat data tahun ' + year + ': ' + err.message);
            showView('dashboard');
        }
    }

    // ═══════════════════════════════════════════════════════════
    //  EVENT BINDINGS
    // ═══════════════════════════════════════════════════════════
    function bindEvents() {
        const pickBtn = document.getElementById('pickYearBtn');
        const menu = document.getElementById('yearPickerMenu');

        if (pickBtn && menu) {
            pickBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                menu.classList.toggle('hidden');
            });
            document.addEventListener('click', function (e) {
                if (!menu.contains(e.target) && !pickBtn.contains(e.target)) {
                    closeYearMenu();
                }
            });
        }

        const back = document.getElementById('backToDashBtn');
        if (back) {
            back.addEventListener('click', function () {
                showView('dashboard');
            });
        }

        const select = document.getElementById('yearlyYearSelect');
        if (select) {
            select.addEventListener('change', function (e) {
                const y = parseInt(e.target.value, 10);
                if (!isNaN(y)) loadYearlyData(y);
            });
        }

        // Tutup menu saat resize agar posisinya tidak "nyangkut"
        window.addEventListener('resize', closeYearMenu);
    }

    // ═══════════════════════════════════════════════════════════
    //  INIT
    // ═══════════════════════════════════════════════════════════
    function init() {
        bindEvents();
        loadDashboardData();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();