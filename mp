<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Menteng Park Monitoring - Manajemen Listrik & Air</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Font Awesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- html2pdf for PDF Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Cinzel:wght@600;700;800&display=swap" rel="stylesheet">
    
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        gold: {
                            50: '#fffbeb',
                            100: '#fef3c7',
                            300: '#fcd34d',
                            400: '#fbbf24',
                            500: '#f59e0b',
                            600: '#d97706',
                            700: '#b45309',
                            800: '#92400e',
                            goldAccent: '#D4AF37',
                            goldLight: '#F3E5AB',
                            goldDark: '#996515'
                        },
                        darkbg: '#0A0C10',
                        darkcard: '#121620',
                        darkborder: '#1E2433',
                        meterDiamond: '#06B6D4',
                        meterSapphire: '#6366F1',
                        meterEmerald: '#10B981',
                        meterPLN: '#F59E0B'
                    },
                    fontFamily: {
                        sans: ['Plus Jakarta Sans', 'sans-serif'],
                        serif: ['Cinzel', 'serif']
                    }
                }
            }
        }
    </script>
    <style>
        body {
            background-color: #0A0C10;
            color: #E2E8F0;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }
        .gold-gradient-text {
            background: linear-gradient(135deg, #FFF5C0 0%, #D4AF37 50%, #AA7C11 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .gold-border-glow {
            border: 1px solid rgba(212, 175, 55, 0.4);
            box-shadow: 0 0 15px rgba(212, 175, 55, 0.15);
        }
        .gold-btn {
            background: linear-gradient(135deg, #D4AF37 0%, #B8860B 100%);
            color: #0A0C10;
            font-weight: 700;
            transition: all 0.3s ease;
        }
        .gold-btn:hover {
            background: linear-gradient(135deg, #F3E5AB 0%, #D4AF37 100%);
            box-shadow: 0 0 12px rgba(212, 175, 55, 0.4);
        }
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #0A0C10;
        }
        ::-webkit-scrollbar-thumb {
            background: #1E2433;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #D4AF37;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col justify-between selection:bg-gold-500 selection:text-black">

    <!-- HEADER / NAVBAR -->
    <header class="border-b border-gold-goldAccent/30 bg-darkcard/90 backdrop-blur-md sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row items-center justify-between gap-4">
            <div class="flex items-center space-x-3">
                <div class="w-11 h-11 rounded-xl bg-gradient-to-tr from-gold-goldDark via-gold-goldAccent to-yellow-200 flex items-center justify-center shadow-lg shadow-gold-500/10">
                    <i class="fa-solid font-bold text-darkbg text-xl fa-building-circle-check"></i>
                </div>
                <div>
                    <h1 class="text-2xl font-bold tracking-wider font-serif gold-gradient-text uppercase">Menteng Park</h1>
                    <p class="text-xs text-gray-400 flex items-center gap-2">
                        <span>SISTEM MONITORING UTILITAS & AUDIT</span>
                        <span id="sync-dot" class="inline-block w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
                        <span id="sync-status-text" class="text-emerald-400 font-medium">Terhubung Google Sheets Live</span>
                    </p>
                </div>
            </div>

            <!-- NAVIGATION TABS -->
            <nav class="flex flex-wrap gap-1 sm:gap-2 bg-darkbg/80 p-1.5 rounded-xl border border-darkborder">
                <button onclick="switchTab('dashboard')" id="btn-tab-dashboard" class="tab-btn active px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg transition-all text-gold-goldAccent bg-gold-goldAccent/10 border border-gold-goldAccent/40">
                    <i class="fa-solid fa-chart-line mr-1.5"></i> Beranda
                </button>
                <button onclick="switchTab('audit')" id="btn-tab-audit" class="tab-btn px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg text-gray-400 hover:text-gray-200 transition-all">
                    <i class="fa-solid fa-scale-balanced mr-1.5"></i> Audit Selisih (Loss)
                </button>
                <button onclick="switchTab('input')" id="btn-tab-input" class="tab-btn px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg text-gray-400 hover:text-gray-200 transition-all">
                    <i class="fa-solid fa-pen-to-square mr-1.5"></i> Input & Data
                </button>
                <button onclick="switchTab('alarm')" id="btn-tab-alarm" class="tab-btn px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg text-gray-400 hover:text-gray-200 transition-all relative">
                    <i class="fa-solid fa-bell mr-1.5"></i> Alarm Pemakaian
                    <span id="alarm-badge" class="hidden absolute -top-1 -right-1 bg-red-500 text-white text-[10px] w-4 h-4 rounded-full flex items-center justify-center animate-bounce">!</span>
                </button>
                <button onclick="switchTab('tariff')" id="btn-tab-tariff" class="tab-btn px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg text-gray-400 hover:text-gray-200 transition-all">
                    <i class="fa-solid fa-calculator mr-1.5"></i> Tarif Utilitas
                </button>
                <button onclick="switchTab('billing')" id="btn-tab-billing" class="tab-btn px-3 py-2 text-xs sm:text-sm font-semibold rounded-lg text-gray-400 hover:text-gray-200 transition-all">
                    <i class="fa-solid fa-file-invoice-dollar mr-1.5"></i> Billing Bulanan
                </button>
            </nav>
        </div>
    </header>

    <!-- GLOBAL ALERT BANNER -->
    <div id="global-alert-box" class="hidden max-w-7xl mx-auto px-4 mt-4">
        <div class="bg-red-900/40 border border-red-500/60 rounded-xl p-4 flex items-center justify-between text-red-200 shadow-lg">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-triangle-exclamation text-red-400 text-2xl"></i>
                <div>
                    <h4 class="font-bold text-sm text-red-300">Peringatan Pemakaian Utilitas Berlebih!</h4>
                    <p id="global-alert-text" class="text-xs text-red-200/90"></p>
                </div>
            </div>
            <button onclick="dismissAlert()" class="text-xs bg-red-800/60 hover:bg-red-700 px-3 py-1.5 rounded-lg border border-red-500/40">
                Tutup
            </button>
        </div>
    </div>

    <!-- MAIN APP CONTENT -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-grow w-full space-y-6">

        <!-- TAB 1: DASHBOARD -->
        <section id="tab-dashboard" class="space-y-6">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-darkcard p-4 rounded-2xl border border-darkborder">
                <div>
                    <h2 class="text-xl font-bold font-serif gold-gradient-text">Ringkasan Pemakaian Utilitas</h2>
                    <p class="text-xs text-gray-400">Monitoring real-time konsumsi Listrik (kWh) & Air (M³) untuk 23+ User</p>
                </div>
                
                <div class="flex flex-wrap items-center gap-3 w-full md:w-auto">
                    <div class="bg-darkbg px-3 py-1.5 rounded-xl border border-darkborder flex items-center space-x-2">
                        <span class="text-xs text-gray-400">Periode:</span>
                        <select id="chart-period" onchange="updateDashboardCharts()" class="bg-transparent text-xs font-semibold text-gold-goldAccent focus:outline-none cursor-pointer">
                            <option value="daily" class="bg-darkcard text-white">Harian (7 Hari Terakhir)</option>
                            <option value="weekly" class="bg-darkcard text-white">Mingguan (4 Minggu)</option>
                            <option value="monthly" selected class="bg-darkcard text-white">Bulanan (Tahun 2026)</option>
                        </select>
                    </div>

                    <button onclick="refreshCloudData()" class="px-3 py-2 text-xs font-bold rounded-xl bg-gold-goldAccent/20 text-gold-goldAccent border border-gold-goldAccent/40 hover:bg-gold-goldAccent/30 transition-all flex items-center gap-1.5">
                        <i class="fa-solid fa-rotate"></i> Sinkronisasi
                    </button>
                    <button onclick="exportToExcel()" class="px-3 py-2 text-xs font-bold rounded-xl bg-emerald-600/20 text-emerald-400 border border-emerald-500/40 hover:bg-emerald-600/30 transition-all flex items-center gap-1.5">
                        <i class="fa-solid fa-file-excel"></i> Excel
                    </button>
                    <button onclick="exportToPDF()" class="px-3 py-2 text-xs font-bold rounded-xl bg-rose-600/20 text-rose-400 border border-rose-500/40 hover:bg-rose-600/30 transition-all flex items-center gap-1.5">
                        <i class="fa-solid fa-file-pdf"></i> PDF
                    </button>
                </div>
            </div>

            <!-- KPI LISTRIK -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterPLN border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-amber-400">1. Meter PLN (Utama)</span>
                        <span class="p-2 rounded-lg bg-amber-500/10 text-amber-400 text-xs"><i class="fa-solid fa-bolt"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-pln-val">0 kWh</div>
                    <div class="text-[11px] text-gray-400 flex justify-between">
                        <span>Hari ini: <strong class="text-white" id="kpi-pln-daily">0</strong> kWh</span>
                        <span class="text-amber-400" id="kpi-pln-cost">Rp 0</span>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterDiamond border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-cyan-400">2. Listrik Diamond</span>
                        <span class="p-2 rounded-lg bg-cyan-500/10 text-cyan-400 text-xs"><i class="fa-solid fa-gem"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-elec-diamond-val">0 kWh</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-cyan-300" id="kpi-elec-diamond-monthly">0 kWh</strong>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterSapphire border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-indigo-400">3. Listrik Sapphire</span>
                        <span class="p-2 rounded-lg bg-indigo-500/10 text-indigo-400 text-xs"><i class="fa-solid fa-gem"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-elec-sapphire-val">0 kWh</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-indigo-300" id="kpi-elec-sapphire-monthly">0 kWh</strong>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterEmerald border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-400">4. Listrik Emerald</span>
                        <span class="p-2 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs"><i class="fa-solid fa-gem"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-elec-emerald-val">0 kWh</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-emerald-300" id="kpi-elec-emerald-monthly">0 kWh</strong>
                    </div>
                </div>
            </div>

            <!-- KPI AIR -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-blue-500 border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-400">1. Meter PDAM (Utama)</span>
                        <span class="p-2 rounded-lg bg-blue-500/10 text-blue-400 text-xs"><i class="fa-solid fa-droplet"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-pdam-val">0 M³</div>
                    <div class="text-[11px] text-gray-400 flex justify-between">
                        <span>Hari ini: <strong class="text-white" id="kpi-pdam-daily">0</strong> M³</span>
                        <span class="text-blue-400" id="kpi-pdam-cost">Rp 0</span>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterDiamond border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-cyan-400">2. Air Diamond</span>
                        <span class="p-2 rounded-lg bg-cyan-500/10 text-cyan-400 text-xs"><i class="fa-solid fa-faucet-drip"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-water-diamond-val">0 M³</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-cyan-300" id="kpi-water-diamond-monthly">0 M³</strong>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterSapphire border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-indigo-400">3. Air Sapphire</span>
                        <span class="p-2 rounded-lg bg-indigo-500/10 text-indigo-400 text-xs"><i class="fa-solid fa-faucet-drip"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-water-sapphire-val">0 M³</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-indigo-300" id="kpi-water-sapphire-monthly">0 M³</strong>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border-l-4 border-l-meterEmerald border border-darkborder">
                    <div class="flex justify-between items-start mb-2">
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-400">4. Air Emerald</span>
                        <span class="p-2 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs"><i class="fa-solid fa-faucet-drip"></i></span>
                    </div>
                    <div class="text-2xl font-bold text-white mb-1" id="kpi-water-emerald-val">0 M³</div>
                    <div class="text-[11px] text-gray-400">
                        Bulan Ini: <strong class="text-emerald-300" id="kpi-water-emerald-monthly">0 M³</strong>
                    </div>
                </div>
            </div>

            <!-- GRAFIK -->
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-darkcard p-5 rounded-2xl border border-darkborder gold-border-glow">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-sm font-bold uppercase tracking-wider text-gold-goldAccent flex items-center gap-2">
                            <i class="fa-solid fa-bolt text-amber-400"></i> Grafik Pemakaian Listrik (kWh)
                        </h3>
                        <span class="text-xs text-gray-400 bg-darkbg px-2.5 py-1 rounded-md border border-darkborder">PLN vs Sub-Meter</span>
                    </div>
                    <div class="relative h-72 w-full">
                        <canvas id="electricityChart"></canvas>
                    </div>
                </div>

                <div class="bg-darkcard p-5 rounded-2xl border border-darkborder gold-border-glow">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-sm font-bold uppercase tracking-wider text-blue-400 flex items-center gap-2">
                            <i class="fa-solid fa-droplet text-blue-400"></i> Grafik Pemakaian Air (M³)
                        </h3>
                        <span class="text-xs text-gray-400 bg-darkbg px-2.5 py-1 rounded-md border border-darkborder">PDAM vs Sub-Meter</span>
                    </div>
                    <div class="relative h-72 w-full">
                        <canvas id="waterChart"></canvas>
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 2: AUDIT SELISIH (LOSS AUDIT) -->
        <section id="tab-audit" class="hidden space-y-6">
            <div class="bg-darkcard p-6 rounded-2xl border border-darkborder space-y-6">
                <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-darkborder pb-4">
                    <div>
                        <h3 class="text-lg font-bold font-serif gold-gradient-text">Audit Selisih Utilitas (Main vs Sub-Meter)</h3>
                        <p class="text-xs text-gray-400">Membandingkan total pemakaian meter utama dengan gabungan sub-meter untuk mendeteksi kehilangan/kebocoran.</p>
                    </div>
                    <div class="bg-darkbg px-3 py-1.5 rounded-xl border border-darkborder text-xs text-gold-goldAccent font-semibold">
                        Periode Audit Terakhir: <span id="audit-date-label">-</span>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <!-- AUDIT LISTRIK -->
                    <div class="bg-darkbg p-5 rounded-2xl border border-amber-500/30 space-y-4">
                        <div class="flex justify-between items-center">
                            <h4 class="text-xs font-bold uppercase tracking-wider text-amber-400 flex items-center gap-2">
                                <i class="fa-solid fa-bolt"></i> Audit Kelistrikan (kWh)
                            </h4>
                            <span id="audit-elec-status-badge" class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">Normal</span>
                        </div>

                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between py-1 border-b border-darkborder">
                                <span class="text-gray-400">Pemakaian Meter PLN (Utama):</span>
                                <strong class="text-white font-mono" id="audit-main-elec">0 kWh</strong>
                            </div>
                            <div class="flex justify-between py-1 border-b border-darkborder">
                                <span class="text-gray-400">Total Sub-Meter (Diamond+Sapphire+Emerald):</span>
                                <strong class="text-white font-mono" id="audit-sub-elec">0 kWh</strong>
                            </div>
                            <div class="flex justify-between py-2 bg-amber-950/20 px-3 rounded-lg border border-amber-500/20">
                                <span class="font-bold text-amber-300">Selisih Susut (Loss / Unaccounted):</span>
                                <strong class="font-mono text-amber-300" id="audit-diff-elec">0 kWh (0%)</strong>
                            </div>
                        </div>
                        <p class="text-[11px] text-gray-400 italic">* Selisih wajar kelistrikan trafo/jaringan gedung umumnya berkisar 2% - 5%.</p>
                    </div>

                    <!-- AUDIT AIR -->
                    <div class="bg-darkbg p-5 rounded-2xl border border-blue-500/30 space-y-4">
                        <div class="flex justify-between items-center">
                            <h4 class="text-xs font-bold uppercase tracking-wider text-blue-400 flex items-center gap-2">
                                <i class="fa-solid fa-droplet"></i> Audit Air Bersih (M³)
                            </h4>
                            <span id="audit-water-status-badge" class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">Normal</span>
                        </div>

                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between py-1 border-b border-darkborder">
                                <span class="text-gray-400">Pemakaian Meter PDAM (Utama):</span>
                                <strong class="text-white font-mono" id="audit-main-water">0 M³</strong>
                            </div>
                            <div class="flex justify-between py-1 border-b border-darkborder">
                                <span class="text-gray-400">Total Sub-Meter (Diamond+Sapphire+Emerald):</span>
                                <strong class="text-white font-mono" id="audit-sub-water">0 M³</strong>
                            </div>
                            <div class="flex justify-between py-2 bg-blue-950/20 px-3 rounded-lg border border-blue-500/20">
                                <span class="font-bold text-blue-300">Selisih Kehilangan (Leakage / Loss):</span>
                                <strong class="font-mono text-blue-300" id="audit-diff-water">0 M³ (0%)</strong>
                            </div>
                        </div>
                        <p class="text-[11px] text-gray-400 italic">* Selisih air di atas 10% mengindikasikan potensi kebocoran pipa riser utama.</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 3: INPUT & TABEL -->
        <section id="tab-input" class="hidden space-y-6">
            <div class="bg-darkcard p-6 rounded-2xl border border-gold-goldAccent/30 shadow-xl">
                <div class="flex justify-between items-center mb-6 pb-4 border-b border-darkborder">
                    <div>
                        <h3 class="text-lg font-bold font-serif gold-gradient-text" id="form-title">Input Catatan Meteran Terbaru</h3>
                        <p class="text-xs text-gray-400">Dilindungi validasi anti-negatif kumulatif & pemeriksaan tanggal kronologis</p>
                    </div>
                    <button type="button" onclick="resetForm()" id="btn-cancel-edit" class="hidden text-xs bg-gray-700 hover:bg-gray-600 text-white px-3 py-1.5 rounded-lg">
                        Batal Edit
                    </button>
                </div>

                <form id="meter-form" onsubmit="handleFormSubmit(event)" class="space-y-6">
                    <input type="hidden" id="entry-id" value="">
                    
                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                            <label class="block text-xs font-semibold text-gray-300 mb-1">Tanggal Pencatatan</label>
                            <input type="date" id="input-date" required class="w-full bg-darkbg border border-darkborder rounded-xl px-3 py-2 text-sm text-white focus:border-gold-goldAccent focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-gray-300 mb-1">Waktu / Shift</label>
                            <input type="time" id="input-time" required class="w-full bg-darkbg border border-darkborder rounded-xl px-3 py-2 text-sm text-white focus:border-gold-goldAccent focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs font-semibold text-gray-300 mb-1">Petugas / Keterangan</label>
                            <input type="text" id="input-notes" placeholder="Contoh: Shift Pagi / Petugas A" class="w-full bg-darkbg border border-darkborder rounded-xl px-3 py-2 text-sm text-white focus:border-gold-goldAccent focus:outline-none">
                        </div>
                    </div>

                    <!-- LISTRIK -->
                    <div>
                        <h4 class="text-xs font-bold uppercase tracking-widest text-amber-400 mb-3 flex items-center gap-2">
                            <i class="fa-solid fa-bolt"></i> Meter Listrik (Satuan kWh - Kumulatif)
                        </h4>
                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                            <div class="bg-amber-950/20 border border-amber-500/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-amber-400 block">1. Meter PLN (Utama)</span>
                                <div>
                                    <label class="text-[10px] text-amber-300/80 block">PLN LWBP (kWh)</label>
                                    <input type="number" step="0.01" min="0" id="input-pln-lwbp" placeholder="0.00" required class="w-full bg-darkbg border border-amber-500/30 rounded-lg px-2.5 py-1.5 text-xs text-amber-200 focus:outline-none">
                                </div>
                                <div>
                                    <label class="text-[10px] text-amber-300/80 block">PLN WBP (kWh)</label>
                                    <input type="number" step="0.01" min="0" id="input-pln-wbp" placeholder="0.00" required class="w-full bg-darkbg border border-amber-500/30 rounded-lg px-2.5 py-1.5 text-xs text-amber-200 focus:outline-none">
                                </div>
                            </div>

                            <div class="bg-cyan-950/20 border border-meterDiamond/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-cyan-400 block">2. Listrik Diamond</span>
                                <label class="text-[10px] text-cyan-300/80 block">Stand Meter (kWh)</label>
                                <input type="number" step="0.01" min="0" id="input-elec-diamond" placeholder="0.00" required class="w-full bg-darkbg border border-meterDiamond/30 rounded-lg px-2.5 py-1.5 text-xs text-cyan-200 focus:outline-none">
                            </div>

                            <div class="bg-indigo-950/20 border border-meterSapphire/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-indigo-400 block">3. Listrik Sapphire</span>
                                <label class="text-[10px] text-indigo-300/80 block">Stand Meter (kWh)</label>
                                <input type="number" step="0.01" min="0" id="input-elec-sapphire" placeholder="0.00" required class="w-full bg-darkbg border border-meterSapphire/30 rounded-lg px-2.5 py-1.5 text-xs text-indigo-200 focus:outline-none">
                            </div>

                            <div class="bg-emerald-950/20 border border-meterEmerald/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-emerald-400 block">4. Listrik Emerald</span>
                                <label class="text-[10px] text-emerald-300/80 block">Stand Meter (kWh)</label>
                                <input type="number" step="0.01" min="0" id="input-elec-emerald" placeholder="0.00" required class="w-full bg-darkbg border border-meterEmerald/30 rounded-lg px-2.5 py-1.5 text-xs text-emerald-200 focus:outline-none">
                            </div>
                        </div>
                    </div>

                    <!-- AIR -->
                    <div>
                        <h4 class="text-xs font-bold uppercase tracking-widest text-blue-400 mb-3 flex items-center gap-2">
                            <i class="fa-solid fa-droplet"></i> Meter Air (Satuan M³ - Kumulatif)
                        </h4>
                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                            <div class="bg-blue-950/20 border border-blue-500/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-blue-400 block">1. Meter PDAM (Utama)</span>
                                <label class="text-[10px] text-blue-300/80 block">Stand Meter (M³)</label>
                                <input type="number" step="0.01" min="0" id="input-pdam" placeholder="0.00" required class="w-full bg-darkbg border border-blue-500/30 rounded-lg px-2.5 py-1.5 text-xs text-blue-200 focus:outline-none">
                            </div>

                            <div class="bg-cyan-950/20 border border-meterDiamond/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-cyan-400 block">2. Air Diamond</span>
                                <label class="text-[10px] text-cyan-300/80 block">Stand Meter (M³)</label>
                                <input type="number" step="0.01" min="0" id="input-water-diamond" placeholder="0.00" required class="w-full bg-darkbg border border-meterDiamond/30 rounded-lg px-2.5 py-1.5 text-xs text-cyan-200 focus:outline-none">
                            </div>

                            <div class="bg-indigo-950/20 border border-meterSapphire/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-indigo-400 block">3. Air Sapphire</span>
                                <label class="text-[10px] text-indigo-300/80 block">Stand Meter (M³)</label>
                                <input type="number" step="0.01" min="0" id="input-water-sapphire" placeholder="0.00" required class="w-full bg-darkbg border border-meterSapphire/30 rounded-lg px-2.5 py-1.5 text-xs text-indigo-200 focus:outline-none">
                            </div>

                            <div class="bg-emerald-950/20 border border-meterEmerald/40 p-3 rounded-xl space-y-2">
                                <span class="text-xs font-bold text-emerald-400 block">4. Air Emerald</span>
                                <label class="text-[10px] text-emerald-300/80 block">Stand Meter (M³)</label>
                                <input type="number" step="0.01" min="0" id="input-water-emerald" placeholder="0.00" required class="w-full bg-darkbg border border-meterEmerald/30 rounded-lg px-2.5 py-1.5 text-xs text-emerald-200 focus:outline-none">
                            </div>
                        </div>
                    </div>

                    <div class="flex justify-end gap-3 pt-2">
                        <button type="submit" id="submit-btn-element" class="gold-btn px-6 py-2.5 rounded-xl text-sm shadow-md flex items-center gap-2">
                            <i class="fa-solid fa-floppy-disk"></i> <span id="btn-submit-text">Simpan & Sinkron ke Sheets</span>
                        </button>
                    </div>
                </form>
            </div>

            <!-- TABEL DATA -->
            <div class="bg-darkcard p-6 rounded-2xl border border-darkborder shadow-xl">
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-4">
                    <h3 class="text-md font-bold text-gold-goldAccent flex items-center gap-2">
                        <i class="fa-solid fa-table-list"></i> Riwayat Pencatatan Cloud & Lokal
                    </h3>
                    <div class="w-full sm:w-auto flex items-center gap-2">
                        <input type="text" id="search-data" onkeyup="filterDataTable()" placeholder="Cari tanggal / catatan..." class="bg-darkbg border border-darkborder rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none w-full sm:w-64">
                    </div>
                </div>

                <div class="overflow-x-auto rounded-xl border border-darkborder">
                    <table class="w-full text-left text-xs text-gray-300">
                        <thead class="bg-darkbg text-gray-400 uppercase tracking-wider font-semibold border-b border-darkborder">
                            <tr>
                                <th class="p-3">Tanggal & Jam</th>
                                <th class="p-3 text-amber-400 bg-amber-950/20">PLN LWBP/WBP</th>
                                <th class="p-3 text-cyan-400 bg-cyan-950/20">Listrik Diamond</th>
                                <th class="p-3 text-indigo-400 bg-indigo-950/20">Listrik Sapphire</th>
                                <th class="p-3 text-emerald-400 bg-emerald-950/20">Listrik Emerald</th>
                                <th class="p-3 text-blue-400 bg-blue-950/20">Meter PDAM</th>
                                <th class="p-3 text-cyan-400 bg-cyan-950/20">Air Diamond</th>
                                <th class="p-3 text-indigo-400 bg-indigo-950/20">Air Sapphire</th>
                                <th class="p-3 text-emerald-400 bg-emerald-950/20">Air Emerald</th>
                                <th class="p-3 text-center">Aksi</th>
                            </tr>
                        </thead>
                        <tbody id="table-body" class="divide-y divide-darkborder">
                            <!-- Rendered dynamically -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- TAB 4: ALARM -->
        <section id="tab-alarm" class="hidden space-y-6">
            <div class="bg-darkcard p-6 rounded-2xl border border-darkborder">
                <div class="mb-6">
                    <h3 class="text-lg font-bold font-serif gold-gradient-text">Pengaturan Alarm & Batas Wajar Pemakaian</h3>
                    <p class="text-xs text-gray-400">Atur batas harian wajar. Sistem akan mendeteksi anomali pemakaian secara otomatis.</p>
                </div>

                <form id="alarm-settings-form" onsubmit="saveAlarmSettings(event)" class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="space-y-4 bg-darkbg p-4 rounded-xl border border-darkborder">
                        <h4 class="text-xs font-bold uppercase tracking-wider text-amber-400 border-b border-darkborder pb-2">
                            <i class="fa-solid fa-bolt"></i> Batas Listrik Harian (kWh)
                        </h4>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">PLN Utama (kWh/Hari)</label>
                            <input type="number" step="1" id="threshold-pln" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Listrik Diamond (kWh/Hari)</label>
                            <input type="number" step="1" id="threshold-elec-diamond" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Listrik Sapphire (kWh/Hari)</label>
                            <input type="number" step="1" id="threshold-elec-sapphire" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Listrik Emerald (kWh/Hari)</label>
                            <input type="number" step="1" id="threshold-elec-emerald" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                    </div>

                    <div class="space-y-4 bg-darkbg p-4 rounded-xl border border-darkborder">
                        <h4 class="text-xs font-bold uppercase tracking-wider text-blue-400 border-b border-darkborder pb-2">
                            <i class="fa-solid fa-droplet"></i> Batas Air Harian (M³)
                        </h4>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">PDAM Utama (M³/Hari)</label>
                            <input type="number" step="1" id="threshold-pdam" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Air Diamond (M³/Hari)</label>
                            <input type="number" step="1" id="threshold-water-diamond" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Air Sapphire (M³/Hari)</label>
                            <input type="number" step="1" id="threshold-water-sapphire" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Air Emerald (M³/Hari)</label>
                            <input type="number" step="1" id="threshold-water-emerald" class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                    </div>

                    <div class="md:col-span-2 flex justify-end">
                        <button type="submit" class="gold-btn px-6 py-2.5 rounded-xl text-sm flex items-center gap-2">
                            <i class="fa-solid fa-shield-halved"></i> Simpan Batas Alarm
                        </button>
                    </div>
                </form>

                <div class="mt-8">
                    <h4 class="text-xs font-bold uppercase tracking-wider text-gold-goldAccent mb-3">
                        Status Evaluasi Alarm Hari Ini
                    </h4>
                    <div id="alarm-status-list" class="grid grid-cols-1 md:grid-cols-2 gap-3">
                        <!-- Rendered JS -->
                    </div>
                </div>
            </div>
        </section>

        <!-- TAB 5: TARIF -->
        <section id="tab-tariff" class="hidden space-y-6">
            <div class="bg-darkcard p-6 rounded-2xl border border-darkborder">
                <div class="mb-6">
                    <h3 class="text-lg font-bold font-serif gold-gradient-text">Pengaturan Tarif PLN & PDAM</h3>
                    <p class="text-xs text-gray-400">Atur tarif dasar PLN (WBP/LWBP), faktor CT, serta PDAM untuk kalkulasi tagihan bulanan.</p>
                </div>

                <form id="tariff-form" onsubmit="saveTariffSettings(event)" class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="bg-darkbg p-5 rounded-xl border border-amber-500/30 space-y-4">
                        <h4 class="text-xs font-bold uppercase tracking-wider text-amber-400 flex items-center gap-2">
                            <i class="fa-solid fa-bolt"></i> Tarif PLN (Listrik)
                        </h4>
                        <div>
                            <label class="block text-xs font-semibold text-amber-300 mb-1">Faktor Kali CT (Current Transformer)</label>
                            <input type="number" step="0.01" min="0.01" id="tariff-pln-ct" required class="w-full bg-darkcard border border-amber-500/40 rounded-lg px-3 py-2 text-xs text-amber-200 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Tarif LWBP (Rp/kWh)</label>
                            <input type="number" step="1" id="tariff-pln-lwbp" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-amber-200 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Tarif WBP (Rp/kWh)</label>
                            <input type="number" step="1" id="tariff-pln-wbp" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-amber-200 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Pajak / PPJU Listrik (%)</label>
                            <input type="number" step="0.1" id="tariff-pln-tax" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Biaya Admin PLN (Rp)</label>
                            <input type="number" step="1000" id="tariff-pln-admin" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                    </div>

                    <div class="bg-darkbg p-5 rounded-xl border border-blue-500/30 space-y-4">
                        <h4 class="text-xs font-bold uppercase tracking-wider text-blue-400 flex items-center gap-2">
                            <i class="fa-solid fa-droplet"></i> Tarif PDAM (Air)
                        </h4>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Tarif Air PDAM (Rp/M³)</label>
                            <input type="number" step="1" id="tariff-pdam-rate" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-blue-200 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Biaya Pemeliharaan & Admin (Rp)</label>
                            <input type="number" step="1000" id="tariff-pdam-admin" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs text-gray-300 mb-1">Pajak Air (%)</label>
                            <input type="number" step="0.1" id="tariff-pdam-tax" required class="w-full bg-darkcard border border-darkborder rounded-lg px-3 py-2 text-xs text-white focus:outline-none">
                        </div>
                    </div>

                    <div class="md:col-span-2 flex justify-end">
                        <button type="submit" class="gold-btn px-6 py-2.5 rounded-xl text-sm flex items-center gap-2">
                            <i class="fa-solid fa-percent"></i> Simpan Tarif Utama
                        </button>
                    </div>
                </form>
            </div>
        </section>

        <!-- TAB 6: BILLING -->
        <section id="tab-billing" class="hidden space-y-6">
            <div class="bg-darkcard p-6 rounded-2xl border border-darkborder">
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6 pb-4 border-b border-darkborder">
                    <div>
                        <h3 class="text-lg font-bold font-serif gold-gradient-text">Estimasi Billing Tagihan Bulanan</h3>
                        <p class="text-xs text-gray-400">Rincian tagihan komprehensif berdasarkan pencatatan meteran</p>
                    </div>
                    <div class="flex items-center gap-2">
                        <select id="billing-month" onchange="renderBillingReport()" class="bg-darkbg border border-gold-goldAccent/40 text-xs font-semibold text-gold-goldAccent rounded-xl px-3 py-2 focus:outline-none">
                            <!-- Populated JS -->
                        </select>
                        <button onclick="printBillingReport()" class="px-3 py-2 bg-gold-goldAccent text-darkbg font-bold rounded-xl text-xs hover:bg-gold-goldLight transition-all flex items-center gap-1">
                            <i class="fa-solid fa-print"></i> Cetak Invoice PDF
                        </button>
                    </div>
                </div>

                <div id="invoice-print-area" class="bg-darkbg p-6 sm:p-8 rounded-2xl border border-gold-goldAccent/30 space-y-6">
                    <div class="flex flex-col sm:flex-row justify-between border-b border-gold-goldAccent/30 pb-4 gap-4">
                        <div>
                            <h2 class="text-2xl font-bold font-serif gold-gradient-text">MENTENG PARK</h2>
                            <p class="text-xs text-gray-400">MANAGEMENT & BUILDING UTILITY INVOICE</p>
                            <p class="text-xs text-gray-500">Jl. Cikini Raya No. 88, Jakarta Pusat</p>
                        </div>
                        <div class="sm:text-right">
                            <span class="inline-block px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-widest bg-gold-goldAccent/10 text-gold-goldAccent border border-gold-goldAccent/40 mb-1">TAGIHAN RESMI</span>
                            <p class="text-xs text-gray-300">Periode: <strong id="inv-period" class="text-white">September 2026</strong></p>
                            <p class="text-[11px] text-gray-400">Tanggal Cetak: <span id="inv-print-date">17/09/2026</span></p>
                        </div>
                    </div>

                    <!-- PLN INVOICE -->
                    <div>
                        <div class="flex justify-between items-center mb-2">
                            <h4 class="text-xs font-bold uppercase tracking-wider text-amber-400 flex items-center gap-2">
                                <i class="fa-solid fa-bolt"></i> 1. Rincian Tagihan PLN (Listrik)
                            </h4>
                            <span class="text-[11px] text-amber-300/80 font-mono bg-amber-950/40 px-2 py-0.5 rounded border border-amber-500/30" id="inv-pln-ct-display">Faktor Kali CT: x1</span>
                        </div>
                        <div class="overflow-x-auto rounded-xl border border-darkborder">
                            <table class="w-full text-xs text-left">
                                <thead class="bg-darkcard text-gray-400">
                                    <tr>
                                        <th class="p-2.5">Kategori PLN</th>
                                        <th class="p-2.5 text-right">Pemakaian (kWh)</th>
                                        <th class="p-2.5 text-right">Tarif (Rp/kWh)</th>
                                        <th class="p-2.5 text-right">Subtotal Biaya</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-darkborder">
                                    <tr>
                                        <td class="p-2.5 font-medium text-amber-300">Pemakaian LWBP</td>
                                        <td class="p-2.5 text-right font-semibold" id="inv-pln-lwbp-kwh">0 kWh</td>
                                        <td class="p-2.5 text-right" id="inv-pln-lwbp-rate">Rp 1,444</td>
                                        <td class="p-2.5 text-right font-semibold text-white" id="inv-pln-lwbp-cost">Rp 0</td>
                                    </tr>
                                    <tr>
                                        <td class="p-2.5 font-medium text-amber-400">Pemakaian WBP</td>
                                        <td class="p-2.5 text-right font-semibold" id="inv-pln-wbp-kwh">0 kWh</td>
                                        <td class="p-2.5 text-right" id="inv-pln-wbp-rate">Rp 2,166</td>
                                        <td class="p-2.5 text-right font-semibold text-white" id="inv-pln-wbp-cost">Rp 0</td>
                                    </tr>
                                    <tr>
                                        <td class="p-2.5 text-gray-400">Biaya Admin / Beban PLN</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right text-gray-300" id="inv-pln-admin">Rp 50,000</td>
                                    </tr>
                                    <tr>
                                        <td class="p-2.5 text-gray-400">Pajak PJU Listrik (<span id="inv-pln-tax-pct">3</span>%)</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right text-gray-300" id="inv-pln-tax-val">Rp 0</td>
                                    </tr>
                                </tbody>
                                <tfoot class="bg-darkcard font-bold text-amber-400 border-t border-darkborder">
                                    <tr>
                                        <td colspan="3" class="p-2.5 text-right">Total Tagihan PLN Listrik:</td>
                                        <td class="p-2.5 text-right text-sm" id="inv-pln-total">Rp 0</td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>

                    <!-- PDAM INVOICE -->
                    <div>
                        <h4 class="text-xs font-bold uppercase tracking-wider text-blue-400 mb-2 flex items-center gap-2">
                            <i class="fa-solid fa-droplet"></i> 2. Rincian Tagihan PDAM (Air)
                        </h4>
                        <div class="overflow-x-auto rounded-xl border border-darkborder">
                            <table class="w-full text-xs text-left">
                                <thead class="bg-darkcard text-gray-400">
                                    <tr>
                                        <th class="p-2.5">Kategori Air</th>
                                        <th class="p-2.5 text-right">Pemakaian (M³)</th>
                                        <th class="p-2.5 text-right">Tarif (Rp/M³)</th>
                                        <th class="p-2.5 text-right">Subtotal Biaya</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-darkborder">
                                    <tr>
                                        <td class="p-2.5 font-medium text-blue-300">Pemakaian Air PDAM Utama</td>
                                        <td class="p-2.5 text-right font-semibold" id="inv-pdam-m3">0 M³</td>
                                        <td class="p-2.5 text-right" id="inv-pdam-rate">Rp 12,500</td>
                                        <td class="p-2.5 text-right font-semibold text-white" id="inv-pdam-usage-cost">Rp 0</td>
                                    </tr>
                                    <tr>
                                        <td class="p-2.5 text-gray-400">Biaya Pemeliharaan & Admin PDAM</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right text-gray-300" id="inv-pdam-admin">Rp 25,000</td>
                                    </tr>
                                    <tr>
                                        <td class="p-2.5 text-gray-400">Pajak Air (<span id="inv-pdam-tax-pct">10</span>%)</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right">-</td>
                                        <td class="p-2.5 text-right text-gray-300" id="inv-pdam-tax-val">Rp 0</td>
                                    </tr>
                                </tbody>
                                <tfoot class="bg-darkcard font-bold text-blue-400 border-t border-darkborder">
                                    <tr>
                                        <td colspan="3" class="p-2.5 text-right">Total Tagihan PDAM Air:</td>
                                        <td class="p-2.5 text-right text-sm" id="inv-pdam-total">Rp 0</td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>

                    <!-- GRAND TOTAL -->
                    <div class="bg-gradient-to-r from-darkcard via-gold-goldDark/30 to-darkcard p-4 rounded-xl border border-gold-goldAccent/50 flex flex-col sm:flex-row justify-between items-center gap-4">
                        <div>
                            <span class="text-xs uppercase tracking-wider text-gray-400 font-bold block">Grand Total Biaya Utilitas Utama</span>
                            <span class="text-[11px] text-gray-400">(Gabungan PLN Listrik + PDAM Air)</span>
                        </div>
                        <div class="text-2xl font-extrabold gold-gradient-text" id="inv-grand-total">
                            Rp 0
                        </div>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <footer class="border-t border-darkborder bg-darkcard py-4 text-center text-xs text-gray-500">
        <p>&copy; 2026 Menteng Park Utility Monitoring System. Didukung oleh Google Apps Script & Teknologi Web Modern.</p>
    </footer>

    <script>
        let meterEntries = [];
        let tariffSettings = {
            pln_ct_factor: 1,
            pln_lwbp: 1444,
            pln_wbp: 2166,
            pln_tax: 3.0,
            pln_admin: 50000,
            pdam_rate: 12500,
            pdam_admin: 25000,
            pdam_tax: 10.0
        };

        let alarmSettings = {
            pln: 500,
            elec_diamond: 150,
            elec_sapphire: 150,
            elec_emerald: 150,
            pdam: 50,
            water_diamond: 15,
            water_sapphire: 15,
            water_emerald: 15
        };

        let electricityChartInstance = null;
        let waterChartInstance = null;
        let isCloudConnected = false;

        window.onload = function() {
            const today = new Date().toISOString().split('T')[0];
            const timeNow = new Date().toTimeString().slice(0, 5);
            document.getElementById('input-date').value = today;
            document.getElementById('input-time').value = timeNow;

            loadTariffFormValues();
            loadAlarmFormValues();
            refreshCloudData();
        };

        function refreshCloudData() {
            showToast("Menghubungkan ke Google Sheets Cloud...", "info");
            
            if (typeof google !== 'undefined' && google.script && google.script.run) {
                isCloudConnected = true;
                document.getElementById('sync-status-text').innerText = "Terhubung Google Sheets Live";
                document.getElementById('sync-dot').className = "inline-block w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse";

                google.script.run
                    .withSuccessHandler(function(data) {
                        if (data && data.length > 0) {
                            meterEntries = data.map(row => ({
                                id: row.id || Date.now(),
                                date: row.date,
                                time: row.time || '08:00',
                                pln_lwbp: Number(row.pln_lwbp) || 0,
                                pln_wbp: Number(row.pln_wbp) || 0,
                                elec_diamond: Number(row.elec_diamond) || 0,
                                elec_sapphire: Number(row.elec_sapphire) || 0,
                                elec_emerald: Number(row.elec_emerald) || 0,
                                pdam: Number(row.pdam) || 0,
                                water_diamond: Number(row.water_diamond) || 0,
                                water_sapphire: Number(row.water_sapphire) || 0,
                                water_emerald: Number(row.water_emerald) || 0,
                                notes: row.notes || '-'
                            }));
                        } else {
                            meterEntries = getInitialMockData();
                        }
                        finalizeUIRender();
                        showToast("Data berhasil disinkronkan dari Google Sheets!", "success");
                    })
                    .withFailureHandler(function(err) {
                        console.error("GAS Error:", err);
                        fallbackToLocalStorage();
                    })
                    .getDataFromSheet();
            } else {
                fallbackToLocalStorage();
            }
        }

        function fallbackToLocalStorage() {
            isCloudConnected = false;
            document.getElementById('sync-status-text').innerText = "Mode Lokal (LocalStorage)";
            document.getElementById('sync-dot').className = "inline-block w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse";
            
            meterEntries = JSON.parse(localStorage.getItem('mp_meter_data')) || getInitialMockData();
            finalizeUIRender();
            showToast("Berjalan dalam mode lokal (LocalStorage).", "info");
        }

        function finalizeUIRender() {
            renderDataTable();
            updateDashboardCharts();
            populateBillingMonths();
            renderBillingReport();
            updateAuditReport();
            checkAlarms();
        }

        function switchTab(tabId) {
            const tabs = ['dashboard', 'audit', 'input', 'alarm', 'tariff', 'billing'];
            tabs.forEach(t => {
                const el = document.getElementById(`tab-${t}`);
                const btn = document.getElementById(`btn-tab-${t}`);
                if (t === tabId) {
                    el.classList.remove('hidden');
                    btn.classList.add('active', 'text-gold-goldAccent', 'bg-gold-goldAccent/10', 'border', 'border-gold-goldAccent/40');
                    btn.classList.remove('text-gray-400');
                } else {
                    el.classList.add('hidden');
                    btn.classList.remove('active', 'text-gold-goldAccent', 'bg-gold-goldAccent/10', 'border', 'border-gold-goldAccent/40');
                    btn.classList.add('text-gray-400');
                }
            });

            if (tabId === 'dashboard') {
                updateDashboardCharts();
            } else if (tabId === 'billing') {
                renderBillingReport();
            } else if (tabId === 'audit') {
                updateAuditReport();
            }
        }

        function getInitialMockData() {
            const mock = [];
            const baseDate = new Date();
            baseDate.setDate(baseDate.getDate() - 9);

            let plnLwbpAcc = 10000;
            let plnWbpAcc = 5000;
            let eDiaAcc = 2000;
            let eSapAcc = 2100;
            let eEmeAcc = 1900;
            let pdamAcc = 1200;
            let wDiaAcc = 350;
            let wSapAcc = 370;
            let wEmeAcc = 330;

            for (let i = 0; i < 10; i++) {
                const curDate = new Date(baseDate);
                curDate.setDate(baseDate.getDate() + i);
                const dateStr = curDate.toISOString().split('T')[0];

                plnLwbpAcc += Math.floor(250 + Math.random() * 50);
                plnWbpAcc += Math.floor(100 + Math.random() * 30);
                eDiaAcc += Math.floor(90 + Math.random() * 30);
                eSapAcc += Math.floor(95 + Math.random() * 25);
                eEmeAcc += Math.floor(85 + Math.random() * 35);
                pdamAcc += Math.floor(25 + Math.random() * 10);
                wDiaAcc += Math.floor(8 + Math.random() * 3);
                wSapAcc += Math.floor(9 + Math.random() * 3);
                wEmeAcc += Math.floor(7 + Math.random() * 4);

                mock.push({
                    id: Date.now() + i,
                    date: dateStr,
                    time: '08:00',
                    pln_lwbp: plnLwbpAcc,
                    pln_wbp: plnWbpAcc,
                    elec_diamond: eDiaAcc,
                    elec_sapphire: eSapAcc,
                    elec_emerald: eEmeAcc,
                    pdam: pdamAcc,
                    water_diamond: wDiaAcc,
                    water_sapphire: wSapAcc,
                    water_emerald: wEmeAcc,
                    notes: i === 9 ? 'Catatan Pagi Rutin' : 'Pencatatan Harian'
                });
            }
            return mock;
        }

        function handleFormSubmit(e) {
            e.preventDefault();
            const idInput = document.getElementById('entry-id').value;
            
            const newEntry = {
                id: idInput ? parseInt(idInput) : Date.now(),
                date: document.getElementById('input-date').value,
                time: document.getElementById('input-time').value,
                pln_lwbp: parseFloat(document.getElementById('input-pln-lwbp').value) || 0,
                pln_wbp: parseFloat(document.getElementById('input-pln-wbp').value) || 0,
                elec_diamond: parseFloat(document.getElementById('input-elec-diamond').value) || 0,
                elec_sapphire: parseFloat(document.getElementById('input-elec-sapphire').value) || 0,
                elec_emerald: parseFloat(document.getElementById('input-elec-emerald').value) || 0,
                pdam: parseFloat(document.getElementById('input-pdam').value) || 0,
                water_diamond: parseFloat(document.getElementById('input-water-diamond').value) || 0,
                water_sapphire: parseFloat(document.getElementById('input-water-sapphire').value) || 0,
                water_emerald: parseFloat(document.getElementById('input-water-emerald').value) || 0,
                notes: document.getElementById('input-notes').value || '-'
            };

            // VALIDASI #1: ANTI-NEGATIF KUMULATIF
            if (meterEntries.length > 0 && !idInput) {
                const latest = meterEntries[meterEntries.length - 1];
                if (newEntry.pln_lwbp < latest.pln_lwbp || newEntry.pln_wbp < latest.pln_wbp || 
                    newEntry.elec_diamond < latest.elec_diamond || newEntry.elec_sapphire < latest.elec_sapphire ||
                    newEntry.elec_emerald < latest.elec_emerald || newEntry.pdam < latest.pdam ||
                    newEntry.water_diamond < latest.water_diamond || newEntry.water_sapphire < latest.water_sapphire ||
                    newEntry.water_emerald < latest.water_emerald) {
                    showToast("Peringatan: Angka stand meter baru lebih kecil dari pencatatan kumulatif terakhir!", "error");
                    return;
                }
            }

            // VALIDASI #2: CHRONOLOGICAL (URUTAN TANGGAL)
            if (meterEntries.length > 0 && !idInput) {
                const latest = meterEntries[meterEntries.length - 1];
                if (new Date(newEntry.date) < new Date(latest.date)) {
                    showToast("Peringatan: Tanggal pencatatan baru lebih awal dari catatan terakhir yang ada!", "error");
                    return;
                }
            }

            if (idInput) {
                const index = meterEntries.findIndex(item => item.id == idInput);
                if (index !== -1) {
                    meterEntries[index] = newEntry;
                }
            } else {
                meterEntries.push(newEntry);
            }

            meterEntries.sort((a, b) => new Date(a.date + 'T' + a.time) - new Date(b.date + 'T' + b.time));

            if (isCloudConnected) {
                showToast("Menyimpan ke Google Sheets Cloud...", "info");
                google.script.run
                    .withSuccessHandler(function(resp) {
                        showToast(resp.message || "Berhasil disimpan ke Google Sheets!", "success");
                        finalizeUIRender();
                    })
                    .withFailureHandler(function(err) {
                        console.error(err);
                        showToast("Gagal simpan ke cloud. Tersimpan di lokal.", "error");
                        localStorage.setItem('mp_meter_data', JSON.stringify(meterEntries));
                        finalizeUIRender();
                    })
                    .saveDataToSheet(newEntry);
            } else {
                localStorage.setItem('mp_meter_data', JSON.stringify(meterEntries));
                showToast("Data meteran berhasil disimpan lokal!", "success");
                finalizeUIRender();
            }

            resetForm();
        }

        function editEntry(id) {
            const entry = meterEntries.find(item => item.id == id);
            if (!entry) return;

            document.getElementById('entry-id').value = entry.id;
            document.getElementById('input-date').value = entry.date;
            document.getElementById('input-time').value = entry.time || '08:00';
            document.getElementById('input-pln-lwbp').value = entry.pln_lwbp;
            document.getElementById('input-pln-wbp').value = entry.pln_wbp;
            document.getElementById('input-elec-diamond').value = entry.elec_diamond;
            document.getElementById('input-elec-sapphire').value = entry.elec_sapphire;
            document.getElementById('input-elec-emerald').value = entry.elec_emerald;
            document.getElementById('input-pdam').value = entry.pdam;
            document.getElementById('input-water-diamond').value = entry.water_diamond;
            document.getElementById('input-water-sapphire').value = entry.water_sapphire;
            document.getElementById('input-water-emerald').value = entry.water_emerald;
            document.getElementById('input-notes').value = entry.notes;

            document.getElementById('form-title').innerText = 'Edit Data Catatan Meteran';
            document.getElementById('btn-submit-text').innerText = 'Update Data Meteran';
            document.getElementById('btn-cancel-edit').classList.remove('hidden');

            window.scrollTo({ top: 100, behavior: 'smooth' });
        }

        function deleteEntry(id) {
            if (confirm("Apakah Anda yakin ingin menghapus data pencatatan ini?")) {
                meterEntries = meterEntries.filter(item => item.id != id);
                if (!isCloudConnected) {
                    localStorage.setItem('mp_meter_data', JSON.stringify(meterEntries));
                }
                finalizeUIRender();
                showToast("Data berhasil dihapus.", "success");
            }
        }

        function resetForm() {
            document.getElementById('meter-form').reset();
            document.getElementById('entry-id').value = '';
            document.getElementById('form-title').innerText = 'Input Catatan Meteran Terbaru';
            document.getElementById('btn-submit-text').innerText = 'Simpan & Sinkron ke Sheets';
            document.getElementById('btn-cancel-edit').classList.add('hidden');
            
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('input-date').value = today;
            document.getElementById('input-time').value = new Date().toTimeString().slice(0, 5);
        }

        function renderDataTable(filterText = '') {
            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';

            const filtered = meterEntries.filter(item => 
                item.date.includes(filterText) || item.notes.toLowerCase().includes(filterText.toLowerCase())
            );

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" class="p-4 text-center text-gray-500">Tidak ada riwayat pencatatan ditemukan.</td></tr>`;
                return;
            }

            filtered.forEach(item => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-darkbg/50 transition-all border-b border-darkborder/50';
                tr.innerHTML = `
                    <td class="p-3 font-medium text-white">${item.date} <span class="text-[10px] text-gray-400 block">${item.time || '08:00'}</span></td>
                    <td class="p-3 font-mono text-amber-300 bg-amber-950/10">${item.pln_lwbp} / ${item.pln_wbp}</td>
                    <td class="p-3 font-mono text-cyan-300 bg-cyan-950/10">${item.elec_diamond}</td>
                    <td class="p-3 font-mono text-indigo-300 bg-indigo-950/10">${item.elec_sapphire}</td>
                    <td class="p-3 font-mono text-emerald-300 bg-emerald-950/10">${item.elec_emerald}</td>
                    <td class="p-3 font-mono text-blue-300 bg-blue-950/10">${item.pdam}</td>
                    <td class="p-3 font-mono text-cyan-300 bg-cyan-950/10">${item.water_diamond}</td>
                    <td class="p-3 font-mono text-indigo-300 bg-indigo-950/10">${item.water_sapphire}</td>
                    <td class="p-3 font-mono text-emerald-300 bg-emerald-950/10">${item.water_emerald}</td>
                    <td class="p-3 text-center space-x-2">
                        <button onclick="editEntry(${item.id})" class="px-2.5 py-1 bg-amber-600/20 text-amber-400 rounded hover:bg-amber-600/30 text-[11px]"><i class="fa-solid fa-pen"></i></button>
                        <button onclick="deleteEntry(${item.id})" class="px-2.5 py-1 bg-red-600/20 text-red-400 rounded hover:bg-red-600/30 text-[11px]"><i class="fa-solid fa-trash"></i></button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function filterDataTable() {
            const query = document.getElementById('search-data').value;
            renderDataTable(query);
        }

        function updateAuditReport() {
            if (meterEntries.length < 2) return;
            const latest = meterEntries[meterEntries.length - 1];
            const prev = meterEntries[meterEntries.length - 2];

            const mainElecUsage = (latest.pln_lwbp + latest.pln_wbp) - (prev.pln_lwbp + prev.pln_wbp);
            const subElecUsage = (latest.elec_diamond - prev.elec_diamond) + 
                                 (latest.elec_sapphire - prev.elec_sapphire) + 
                                 (latest.elec_emerald - prev.elec_emerald);

            const diffElec = mainElecUsage - subElecUsage;
            const pctElec = mainElecUsage > 0 ? ((diffElec / mainElecUsage) * 100).toFixed(1) : 0;

            const mainWaterUsage = latest.pdam - prev.pdam;
            const subWaterUsage = (latest.water_diamond - prev.water_diamond) + 
                                  (latest.water_sapphire - prev.water_sapphire) + 
                                  (latest.water_emerald - prev.water_emerald);

            const diffWater = mainWaterUsage - subWaterUsage;
            const pctWater = mainWaterUsage > 0 ? ((diffWater / mainWaterUsage) * 100).toFixed(1) : 0;

            document.getElementById('audit-date-label').innerText = `${prev.date} s.d ${latest.date}`;
            document.getElementById('audit-main-elec').innerText = Math.max(0, mainElecUsage).toLocaleString() + ' kWh';
            document.getElementById('audit-sub-elec').innerText = Math.max(0, subElecUsage).toLocaleString() + ' kWh';
            document.getElementById('audit-diff-elec').innerText = `${Math.max(0, diffElec).toLocaleString()} kWh (${pctElec}%)`;

            document.getElementById('audit-main-water').innerText = Math.max(0, mainWaterUsage).toLocaleString() + ' M³';
            document.getElementById('audit-sub-water').innerText = Math.max(0, subWaterUsage).toLocaleString() + ' M³';
            document.getElementById('audit-diff-water').innerText = `${Math.max(0, diffWater).toLocaleString()} M³ (${pctWater}%)`;

            const elecBadge = document.getElementById('audit-elec-status-badge');
            if (pctElec > 10) {
                elecBadge.className = "px-2 py-0.5 rounded text-[10px] font-bold bg-red-500/20 text-red-400 border border-red-500/30";
                elecBadge.innerText = "Susut Tinggi";
            } else {
                elecBadge.className = "px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30";
                elecBadge.innerText = "Normal";
            }

            const waterBadge = document.getElementById('audit-water-status-badge');
            if (pctWater > 10) {
                waterBadge.className = "px-2 py-0.5 rounded text-[10px] font-bold bg-red-500/20 text-red-400 border border-red-500/30";
                waterBadge.innerText = "Potensi Bocor";
            } else {
                waterBadge.className = "px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30";
                waterBadge.innerText = "Normal";
            }
        }

        function updateDashboardCharts() {
            if (meterEntries.length === 0) return;

            const latest = meterEntries[meterEntries.length - 1];
            const prev = meterEntries.length > 1 ? meterEntries[meterEntries.length - 2] : latest;

            const plnUsage = (latest.pln_lwbp + latest.pln_wbp) - (prev.pln_lwbp + prev.pln_wbp);
            const pdamUsage = latest.pdam - prev.pdam;
            const eDiaUsage = latest.elec_diamond - prev.elec_diamond;
            const eSapUsage = latest.elec_sapphire - prev.elec_sapphire;
            const eEmeUsage = latest.elec_emerald - prev.elec_emerald;
            const wDiaUsage = latest.water_diamond - prev.water_diamond;
            const wSapUsage = latest.water_sapphire - prev.water_sapphire;
            const wEmeUsage = latest.water_emerald - prev.water_emerald;

            document.getElementById('kpi-pln-val').innerText = (latest.pln_lwbp + latest.pln_wbp).toLocaleString() + ' kWh';
            document.getElementById('kpi-pln-daily').innerText = Math.max(0, plnUsage).toLocaleString();
            document.getElementById('kpi-pln-cost').innerText = 'Rp ' + Math.round(Math.max(0, plnUsage) * tariffSettings.pln_lwbp).toLocaleString();

            document.getElementById('kpi-elec-diamond-val').innerText = latest.elec_diamond.toLocaleString() + ' kWh';
            document.getElementById('kpi-elec-diamond-monthly').innerText = Math.max(0, eDiaUsage * 30).toLocaleString() + ' kWh';

            document.getElementById('kpi-elec-sapphire-val').innerText = latest.elec_sapphire.toLocaleString() + ' kWh';
            document.getElementById('kpi-elec-sapphire-monthly').innerText = Math.max(0, eSapUsage * 30).toLocaleString() + ' kWh';

            document.getElementById('kpi-elec-emerald-val').innerText = latest.elec_emerald.toLocaleString() + ' kWh';
            document.getElementById('kpi-elec-emerald-monthly').innerText = Math.max(0, eEmeUsage * 30).toLocaleString() + ' kWh';

            document.getElementById('kpi-pdam-val').innerText = latest.pdam.toLocaleString() + ' M³';
            document.getElementById('kpi-pdam-daily').innerText = Math.max(0, pdamUsage).toLocaleString();
            document.getElementById('kpi-pdam-cost').innerText = 'Rp ' + Math.round(Math.max(0, pdamUsage) * tariffSettings.pdam_rate).toLocaleString();

            document.getElementById('kpi-water-diamond-val').innerText = latest.water_diamond.toLocaleString() + ' M³';
            document.getElementById('kpi-water-diamond-monthly').innerText = Math.max(0, wDiaUsage * 30).toLocaleString() + ' M³';

            document.getElementById('kpi-water-sapphire-val').innerText = latest.water_sapphire.toLocaleString() + ' M³';
            document.getElementById('kpi-water-sapphire-monthly').innerText = Math.max(0, wSapUsage * 30).toLocaleString() + ' M³';

            document.getElementById('kpi-water-emerald-val').innerText = latest.water_emerald.toLocaleString() + ' M³';
            document.getElementById('kpi-water-emerald-monthly').innerText = Math.max(0, wEmeUsage * 30).toLocaleString() + ' M³';

            const labels = meterEntries.slice(-7).map(item => item.date);
            const plnData = meterEntries.slice(-7).map(item => (item.pln_lwbp + item.pln_wbp));
            const eDiaData = meterEntries.slice(-7).map(item => item.elec_diamond);
            const eSapData = meterEntries.slice(-7).map(item => item.elec_sapphire);
            const eEmeData = meterEntries.slice(-7).map(item => item.elec_emerald);

            const pdamData = meterEntries.slice(-7).map(item => item.pdam);
            const wDiaData = meterEntries.slice(-7).map(item => item.water_diamond);
            const wSapData = meterEntries.slice(-7).map(item => item.water_sapphire);
            const wEmeData = meterEntries.slice(-7).map(item => item.water_emerald);

            const ctxElec = document.getElementById('electricityChart').getContext('2d');
            if (electricityChartInstance) electricityChartInstance.destroy();
            electricityChartInstance = new Chart(ctxElec, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        { label: 'PLN Utama', data: plnData, borderColor: '#F59E0B', backgroundColor: 'rgba(245, 158, 11, 0.1)', tension: 0.3, fill: true },
                        { label: 'Diamond', data: eDiaData, borderColor: '#06B6D4', tension: 0.3 },
                        { label: 'Sapphire', data: eSapData, borderColor: '#6366F1', tension: 0.3 },
                        { label: 'Emerald', data: eEmeData, borderColor: '#10B981', tension: 0.3 }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 10 } } } },
                    scales: {
                        x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } },
                        y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } }
                    }
                }
            });

            const ctxWater = document.getElementById('waterChart').getContext('2d');
            if (waterChartInstance) waterChartInstance.destroy();
            waterChartInstance = new Chart(ctxWater, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        { label: 'PDAM Utama', data: pdamData, borderColor: '#3b82f6', backgroundColor: 'rgba(59, 130, 246, 0.1)', tension: 0.3, fill: true },
                        { label: 'Air Diamond', data: wDiaData, borderColor: '#06B6D4', tension: 0.3 },
                        { label: 'Air Sapphire', data: wSapData, borderColor: '#6366F1', tension: 0.3 },
                        { label: 'Air Emerald', data: wEmeData, borderColor: '#10B981', tension: 0.3 }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 10 } } } },
                    scales: {
                        x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } },
                        y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } }
                    }
                }
            });
        }

        function checkAlarms() {
            if (meterEntries.length < 2) return;
            const latest = meterEntries[meterEntries.length - 1];
            const prev = meterEntries[meterEntries.length - 2];

            const dPln = (latest.pln_lwbp + latest.pln_wbp) - (prev.pln_lwbp + prev.pln_wbp);
            const dDia = latest.elec_diamond - prev.elec_diamond;
            const dPdam = latest.pdam - prev.pdam;

            const alerts = [];
            if (dPln > alarmSettings.pln) alerts.push(`PLN Utama melampaui batas harian (${dPln} kWh > ${alarmSettings.pln} kWh)`);
            if (dDia > alarmSettings.elec_diamond) alerts.push(`Listrik Diamond melampaui batas (${dDia} kWh)`);
            if (dPdam > alarmSettings.pdam) alerts.push(`PDAM Utama melampaui batas (${dPdam} M³)`);

            const alarmBox = document.getElementById('global-alert-box');
            const alarmBadge = document.getElementById('alarm-badge');

            if (alerts.length > 0) {
                document.getElementById('global-alert-text').innerText = alerts.join(' | ');
                alarmBox.classList.remove('hidden');
                alarmBadge.classList.remove('hidden');
            } else {
                alarmBox.classList.add('hidden');
                alarmBadge.classList.add('hidden');
            }

            const statusList = document.getElementById('alarm-status-list');
            statusList.innerHTML = `
                <div class="bg-darkbg p-3 rounded-xl border ${dPln > alarmSettings.pln ? 'border-red-500/50 bg-red-950/10' : 'border-darkborder'} flex justify-between items-center text-xs">
                    <span>PLN Utama (Aktif: <strong>${Math.max(0, dPln)}</strong> kWh)</span>
                    <span class="${dPln > alarmSettings.pln ? 'text-red-400 font-bold' : 'text-emerald-400'}">${dPln > alarmSettings.pln ? 'PERINGATAN' : 'Normal'}</span>
                </div>
                <div class="bg-darkbg p-3 rounded-xl border ${dPdam > alarmSettings.pdam ? 'border-red-500/50 bg-red-950/10' : 'border-darkborder'} flex justify-between items-center text-xs">
                    <span>PDAM Utama (Aktif: <strong>${Math.max(0, dPdam)}</strong> M³)</span>
                    <span class="${dPdam > alarmSettings.pdam ? 'text-red-400 font-bold' : 'text-emerald-400'}">${dPdam > alarmSettings.pdam ? 'PERINGATAN' : 'Normal'}</span>
                </div>
            `;
        }

        function dismissAlert() {
            document.getElementById('global-alert-box').classList.add('hidden');
        }

        function saveAlarmSettings(e) {
            e.preventDefault();
            alarmSettings.pln = parseFloat(document.getElementById('threshold-pln').value) || 500;
            alarmSettings.elec_diamond = parseFloat(document.getElementById('threshold-elec-diamond').value) || 150;
            alarmSettings.elec_sapphire = parseFloat(document.getElementById('threshold-elec-sapphire').value) || 150;
            alarmSettings.elec_emerald = parseFloat(document.getElementById('threshold-elec-emerald').value) || 150;
            alarmSettings.pdam = parseFloat(document.getElementById('threshold-pdam').value) || 50;

            checkAlarms();
            showToast("Batas alarm berhasil diperbarui!", "success");
        }

        function loadAlarmFormValues() {
            document.getElementById('threshold-pln').value = alarmSettings.pln;
            document.getElementById('threshold-elec-diamond').value = alarmSettings.elec_diamond;
            document.getElementById('threshold-elec-sapphire').value = alarmSettings.elec_sapphire;
            document.getElementById('threshold-elec-emerald').value = alarmSettings.elec_emerald;
            document.getElementById('threshold-pdam').value = alarmSettings.pdam;
            document.getElementById('threshold-water-diamond').value = alarmSettings.water_diamond;
            document.getElementById('threshold-water-sapphire').value = alarmSettings.water_sapphire;
            document.getElementById('threshold-water-emerald').value = alarmSettings.water_emerald;
        }

        function saveTariffSettings(e) {
            e.preventDefault();
            tariffSettings.pln_ct_factor = parseFloat(document.getElementById('tariff-pln-ct').value) || 1;
            tariffSettings.pln_lwbp = parseFloat(document.getElementById('tariff-pln-lwbp').value) || 1444;
            tariffSettings.pln_wbp = parseFloat(document.getElementById('tariff-pln-wbp').value) || 2166;
            tariffSettings.pln_tax = parseFloat(document.getElementById('tariff-pln-tax').value) || 3.0;
            tariffSettings.pln_admin = parseFloat(document.getElementById('tariff-pln-admin').value) || 50000;
            tariffSettings.pdam_rate = parseFloat(document.getElementById('tariff-pdam-rate').value) || 12500;
            tariffSettings.pdam_admin = parseFloat(document.getElementById('tariff-pdam-admin').value) || 25000;
            tariffSettings.pdam_tax = parseFloat(document.getElementById('tariff-pdam-tax').value) || 10.0;

            updateDashboardCharts();
            renderBillingReport();
            showToast("Pengaturan tarif utama berhasil disimpan!", "success");
        }

        function loadTariffFormValues() {
            document.getElementById('tariff-pln-ct').value = tariffSettings.pln_ct_factor;
            document.getElementById('tariff-pln-lwbp').value = tariffSettings.pln_lwbp;
            document.getElementById('tariff-pln-wbp').value = tariffSettings.pln_wbp;
            document.getElementById('tariff-pln-tax').value = tariffSettings.pln_tax;
            document.getElementById('tariff-pln-admin').value = tariffSettings.pln_admin;
            document.getElementById('tariff-pdam-rate').value = tariffSettings.pdam_rate;
            document.getElementById('tariff-pdam-admin').value = tariffSettings.pdam_admin;
            document.getElementById('tariff-pdam-tax').value = tariffSettings.pdam_tax;
        }

        function populateBillingMonths() {
            const select = document.getElementById('billing-month');
            select.innerHTML = '';
            
            const months = ['Januari 2026', 'Februari 2026', 'Maret 2026', 'April 2026', 'Mei 2026', 'Juni 2026', 'Juli 2026', 'Agustus 2026', 'September 2026', 'Oktober 2026', 'November 2026', 'Desember 2026'];
            months.forEach((m, idx) => {
                const opt = document.createElement('option');
                opt.value = idx;
                opt.innerText = m;
                if (idx === 8) opt.selected = true;
                select.appendChild(opt);
            });
        }

        function renderBillingReport() {
            if (meterEntries.length === 0) return;

            const first = meterEntries[0];
            const last = meterEntries[meterEntries.length - 1];

            const lwbpDiff = Math.max(0, last.pln_lwbp - first.pln_lwbp) * tariffSettings.pln_ct_factor;
            const wbpDiff = Math.max(0, last.pln_wbp - first.pln_wbp) * tariffSettings.pln_ct_factor;
            const pdamDiff = Math.max(0, last.pdam - first.pdam);

            const lwbpCost = lwbpDiff * tariffSettings.pln_lwbp;
            const wbpCost = wbpDiff * tariffSettings.pln_wbp;
            const plnSubtotal = lwbpCost + wbpCost + tariffSettings.pln_admin;
            const plnTaxVal = plnSubtotal * (tariffSettings.pln_tax / 100);
            const plnTotal = plnSubtotal + plnTaxVal;

            const pdamUsageCost = pdamDiff * tariffSettings.pdam_rate;
            const pdamSubtotal = pdamUsageCost + tariffSettings.pdam_admin;
            const pdamTaxVal = pdamSubtotal * (tariffSettings.pdam_tax / 100);
            const pdamTotal = pdamSubtotal + pdamTaxVal;

            const grandTotal = plnTotal + pdamTotal;

            document.getElementById('inv-pln-ct-display').innerText = `Faktor Kali CT: x${tariffSettings.pln_ct_factor}`;
            document.getElementById('inv-pln-lwbp-kwh').innerText = lwbpDiff.toLocaleString() + ' kWh';
            document.getElementById('inv-pln-lwbp-rate').innerText = 'Rp ' + tariffSettings.pln_lwbp.toLocaleString();
            document.getElementById('inv-pln-lwbp-cost').innerText = 'Rp ' + Math.round(lwbpCost).toLocaleString();

            document.getElementById('inv-pln-wbp-kwh').innerText = wbpDiff.toLocaleString() + ' kWh';
            document.getElementById('inv-pln-wbp-rate').innerText = 'Rp ' + tariffSettings.pln_wbp.toLocaleString();
            document.getElementById('inv-pln-wbp-cost').innerText = 'Rp ' + Math.round(wbpCost).toLocaleString();

            document.getElementById('inv-pln-admin').innerText = 'Rp ' + tariffSettings.pln_admin.toLocaleString();
            document.getElementById('inv-pln-tax-pct').innerText = tariffSettings.pln_tax;
            document.getElementById('inv-pln-tax-val').innerText = 'Rp ' + Math.round(plnTaxVal).toLocaleString();
            document.getElementById('inv-pln-total').innerText = 'Rp ' + Math.round(plnTotal).toLocaleString();

            document.getElementById('inv-pdam-m3').innerText = pdamDiff.toLocaleString() + ' M³';
            document.getElementById('inv-pdam-rate').innerText = 'Rp ' + tariffSettings.pdam_rate.toLocaleString();
            document.getElementById('inv-pdam-usage-cost').innerText = 'Rp ' + Math.round(pdamUsageCost).toLocaleString();
            document.getElementById('inv-pdam-admin').innerText = 'Rp ' + tariffSettings.pdam_admin.toLocaleString();
            document.getElementById('inv-pdam-tax-pct').innerText = tariffSettings.pdam_tax;
            document.getElementById('inv-pdam-tax-val').innerText = 'Rp ' + Math.round(pdamTaxVal).toLocaleString();
            document.getElementById('inv-pdam-total').innerText = 'Rp ' + Math.round(pdamTotal).toLocaleString();

            document.getElementById('inv-grand-total').innerText = 'Rp ' + Math.round(grandTotal).toLocaleString();
        }

        function printBillingReport() {
            const element = document.getElementById('invoice-print-area');
            const opt = {
                margin: 10,
                filename: 'Invoice_Utilitas_MentengPark.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };
            html2pdf().from(element).set(opt).save();
        }

        function exportToExcel() {
            const ws = XLSX.utils.json_to_sheet(meterEntries);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "DataMeteran");
            XLSX.writeFile(wb, "MentengPark_Data_Utilitas.xlsx");
            showToast("Berhasil mengekspor data ke Excel!", "success");
        }

        function exportToPDF() {
            const element = document.getElementById('tab-dashboard');
            const opt = {
                margin: 5,
                filename: 'MentengPark_Dashboard_Report.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
            };
            html2pdf().from(element).set(opt).save();
        }

        function showToast(message, type = 'success') {
            const toast = document.createElement('div');
            toast.className = `fixed bottom-5 right-5 z-50 px-4 py-3 rounded-xl text-xs font-bold text-white shadow-2xl transition-all duration-300 flex items-center gap-2 ${type === 'success' ? 'bg-emerald-600 border border-emerald-400' : type === 'info' ? 'bg-amber-600 border border-amber-400' : 'bg-red-600 border border-red-400'}`;
            toast.innerHTML = `<i class="fa-solid ${type === 'success' ? 'fa-circle-check' : type === 'info' ? 'fa-circle-info' : 'fa-triangle-exclamation'}"></i> ${message}`;
            document.body.appendChild(toast);
            setTimeout(() => {
                toast.style.opacity = '0';
                setTimeout(() => toast.remove(), 300);
            }, 3500);
        }
    </script>
</body>
</html>
