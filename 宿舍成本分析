<!DOCTYPE html>
<html lang="zh-TW" class="h-full bg-slate-50">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>宿舍智慧收支與累計盈餘管理系統</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS for Excel import/export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            background-color: #f8fafc;
        }
        .tab-btn.active {
            border-bottom: 3px solid #2563eb;
            color: #2563eb;
            font-weight: 600;
        }
        /* Custom Scrollbar for inner tables */
        .custom-scrollbar::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="text-slate-800 min-h-screen flex flex-col">

    <!-- Toast Notification Container -->
    <div id="toast-container" class="fixed top-4 right-4 z-50 flex flex-col gap-2 pointer-events-none"></div>

    <!-- Header Navigation -->
    <header class="bg-slate-900 text-white shadow-lg sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16">
                <div class="flex items-center space-x-3">
                    <div class="bg-blue-600 p-2.5 rounded-xl text-white shadow-md">
                        <i class="fa-solid fa-building-user text-xl"></i>
                    </div>
                    <div>
                        <h1 class="font-bold text-lg leading-tight tracking-wide">宿舍智慧收支與累計盈餘管理系統</h1>
                        <p class="text-xs text-slate-400 hidden sm:block">多宿舍精準對帳 · 民國年相容 · 月度損益對照</p>
                    </div>
                </div>
                <div class="flex items-center space-x-3">
                    <button onclick="downloadExcelTemplate()" class="bg-slate-800 hover:bg-slate-700 text-xs sm:text-sm px-3.5 py-2 rounded-lg border border-slate-700 flex items-center gap-2 transition font-medium">
                        <i class="fa-solid fa-file-excel text-emerald-400"></i>
                        <span>下載 Excel 範本</span>
                    </button>
                    <label class="cursor-pointer bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm px-4 py-2 rounded-lg flex items-center gap-2 font-medium transition shadow-sm">
                        <i class="fa-solid fa-file-import"></i>
                        <span>匯入 Excel 帳單</span>
                        <input type="file" id="excel-file-input" accept=".xlsx, .xls, .csv" class="hidden" onchange="importExcel(event)">
                    </label>
                </div>
            </div>
        </div>
    </header>

    <!-- Main Content Area -->
    <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

        <!-- Multi-dimensional Filter Bar -->
        <section class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200/80">
            <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                <div class="flex items-center gap-2 text-slate-700 font-semibold text-sm shrink-0">
                    <i class="fa-solid fa-sliders text-blue-600"></i>
                    <span>全系統多維度篩選器：</span>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 flex-1">
                    <!-- Target Month Selector (ROC & AD aware) -->
                    <div>
                        <label class="block text-[11px] font-semibold text-slate-500 uppercase tracking-wider mb-1">對帳月份 (年月)</label>
                        <select id="filter-month" onchange="applyFilters()" class="w-full text-xs sm:text-sm border-slate-300 rounded-xl shadow-sm focus:ring-blue-500 focus:border-blue-500 bg-slate-50/50 p-2.5 border font-medium text-slate-800">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                    <!-- Region Filter -->
                    <div>
                        <label class="block text-[11px] font-semibold text-slate-500 uppercase tracking-wider mb-1">區域</label>
                        <select id="filter-region" onchange="applyFilters()" class="w-full text-xs sm:text-sm border-slate-300 rounded-xl shadow-sm focus:ring-blue-500 focus:border-blue-500 bg-slate-50/50 p-2.5 border text-slate-700">
                            <option value="ALL">全部區域</option>
                        </select>
                    </div>
                    <!-- Client/Employer Filter -->
                    <div>
                        <label class="block text-[11px] font-semibold text-slate-500 uppercase tracking-wider mb-1">客戶 / 雇主</label>
                        <select id="filter-client" onchange="applyFilters()" class="w-full text-xs sm:text-sm border-slate-300 rounded-xl shadow-sm focus:ring-blue-500 focus:border-blue-500 bg-slate-50/50 p-2.5 border text-slate-700">
                            <option value="ALL">全部客戶 (雇主)</option>
                        </select>
                    </div>
                    <!-- Dormitory Filter -->
                    <div>
                        <label class="block text-[11px] font-semibold text-slate-500 uppercase tracking-wider mb-1">宿舍名稱</label>
                        <select id="filter-dorm" onchange="applyFilters()" class="w-full text-xs sm:text-sm border-slate-300 rounded-xl shadow-sm focus:ring-blue-500 focus:border-blue-500 bg-slate-50/50 p-2.5 border text-slate-700">
                            <option value="ALL">全部宿舍</option>
                        </select>
                    </div>
                </div>
                <button onclick="resetFilters()" class="text-xs text-slate-500 hover:text-blue-600 underline text-right lg:text-left self-end lg:self-center py-1 font-medium shrink-0">
                    重置所有篩選
                </button>
            </div>
        </section>

        <!-- Navigation Tabs -->
        <div class="border-b border-slate-200 flex space-x-6 text-sm font-medium">
            <button onclick="switchTab('pnl')" id="tab-pnl" class="tab-btn active pb-3 px-1 flex items-center gap-2">
                <i class="fa-solid fa-chart-line"></i> 月度損益儀表板與對帳單
            </button>
            <button onclick="switchTab('records')" id="tab-records" class="tab-btn text-slate-500 hover:text-slate-700 pb-3 px-1 flex items-center gap-2">
                <i class="fa-solid fa-list-check"></i> 全部流水帳與品項管理
            </button>
        </div>

        <!-- TAB 1: P&L Dashboard and Monthly Breakdown -->
        <div id="view-pnl" class="space-y-6">

            <!-- Summary KPI Cards for Selected Month -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
                <!-- Monthly Income -->
                <div class="bg-white p-5 rounded-2xl shadow-sm border border-emerald-100 flex items-center justify-between">
                    <div>
                        <p class="text-xs font-semibold text-emerald-600 uppercase tracking-wider mb-1">
                            <span id="selected-month-label-1">當月</span> 總收入
                        </p>
                        <h3 class="text-2xl font-bold text-slate-800" id="kpi-monthly-income">$0</h3>
                        <p class="text-xs text-slate-400 mt-1" id="kpi-income-count">共 0 筆收款紀錄</p>
                    </div>
                    <div class="w-12 h-12 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-arrow-down-left"></i>
                    </div>
                </div>

                <!-- Monthly Expense -->
                <div class="bg-white p-5 rounded-2xl shadow-sm border border-rose-100 flex items-center justify-between">
                    <div>
                        <p class="text-xs font-semibold text-rose-600 uppercase tracking-wider mb-1">
                            <span id="selected-month-label-2">當月</span> 總支出
                        </p>
                        <h3 class="text-2xl font-bold text-slate-800" id="kpi-monthly-expense">$0</h3>
                        <p class="text-xs text-slate-400 mt-1" id="kpi-expense-count">共 0 筆支出紀錄</p>
                    </div>
                    <div class="w-12 h-12 bg-rose-50 text-rose-600 rounded-2xl flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-arrow-up-right"></i>
                    </div>
                </div>

                <!-- Monthly Net P&L (Profit / Loss) -->
                <div id="kpi-pnl-card" class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between transition-all">
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">
                            <span id="selected-month-label-3">當月</span> 結算損益
                        </p>
                        <h3 class="text-2xl font-bold" id="kpi-monthly-pnl">$0</h3>
                        <p class="text-xs font-medium mt-1" id="kpi-pnl-status">收支打平</p>
                    </div>
                    <div id="kpi-pnl-icon" class="w-12 h-12 bg-slate-100 text-slate-600 rounded-2xl flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-scale-balanced"></i>
                    </div>
                </div>
            </div>

            <!-- Detailed Monthly Breakdown: Side-by-Side Income vs Expense -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                    <div>
                        <h2 class="font-bold text-slate-800 text-base flex items-center gap-2">
                            <i class="fa-solid fa-receipt text-blue-600"></i>
                            <span id="breakdown-title-month">115年1月</span> 當月明細對帳單 (收入 vs 支出)
                        </h2>
                        <p class="text-xs text-slate-500 mt-0.5">即時列出人員繳款（如 A人/B人）與維修採購項目（如 水管維修/洗衣機）</p>
                    </div>
                    <div class="text-xs bg-blue-50 text-blue-700 px-3 py-1.5 rounded-lg border border-blue-100 font-semibold self-start sm:self-auto">
                        雙欄自動彙總模式
                    </div>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-2 divide-y lg:divide-y-0 lg:divide-x divide-slate-200">
                    
                    <!-- LEFT COLUMN: Income Breakdown -->
                    <div class="p-5 flex flex-col h-full">
                        <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
                            <span class="text-sm font-bold text-emerald-700 flex items-center gap-2">
                                <i class="fa-solid fa-circle-arrow-down"></i> 收入明細 (人員 / 各項收費)
                            </span>
                            <span class="text-xs font-semibold bg-emerald-100 text-emerald-800 px-2.5 py-0.5 rounded-full" id="income-item-count">
                                0 筆
                            </span>
                        </div>

                        <div class="flex-1 overflow-y-auto max-h-80 custom-scrollbar pr-1">
                            <table class="w-full text-left text-xs">
                                <thead>
                                    <tr class="text-slate-400 border-b border-slate-100 uppercase">
                                        <th class="pb-2 font-medium">人員 / 內容</th>
                                        <th class="pb-2 font-medium">客戶 / 宿舍</th>
                                        <th class="pb-2 font-medium text-right">金額</th>
                                    </tr>
                                </thead>
                                <tbody id="monthly-income-tbody" class="divide-y divide-slate-100">
                                    <!-- Populated dynamically -->
                                </tbody>
                            </table>
                        </div>

                        <div class="mt-4 pt-3 border-t border-slate-200 flex justify-between items-center bg-emerald-50/60 p-3 rounded-xl">
                            <span class="text-xs font-bold text-emerald-900">收入小計</span>
                            <span class="text-base font-extrabold text-emerald-600" id="income-subtotal">$0</span>
                        </div>
                    </div>

                    <!-- RIGHT COLUMN: Expense Breakdown -->
                    <div class="p-5 flex flex-col h-full">
                        <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
                            <span class="text-sm font-bold text-rose-700 flex items-center gap-2">
                                <i class="fa-solid fa-circle-arrow-up"></i> 支出明細 (維修 / 公用開支品項)
                            </span>
                            <span class="text-xs font-semibold bg-rose-100 text-rose-800 px-2.5 py-0.5 rounded-full" id="expense-item-count">
                                0 筆
                            </span>
                        </div>

                        <div class="flex-1 overflow-y-auto max-h-80 custom-scrollbar pr-1">
                            <table class="w-full text-left text-xs">
                                <thead>
                                    <tr class="text-slate-400 border-b border-slate-100 uppercase">
                                        <th class="pb-2 font-medium">支出品項 / 內容</th>
                                        <th class="pb-2 font-medium">客戶 / 宿舍</th>
                                        <th class="pb-2 font-medium text-right">金額</th>
                                    </tr>
                                </thead>
                                <tbody id="monthly-expense-tbody" class="divide-y divide-slate-100">
                                    <!-- Populated dynamically -->
                                </tbody>
                            </table>
                        </div>

                        <div class="mt-4 pt-3 border-t border-slate-200 flex justify-between items-center bg-rose-50/60 p-3 rounded-xl">
                            <span class="text-xs font-bold text-rose-900">支出小計</span>
                            <span class="text-base font-extrabold text-rose-600" id="expense-subtotal">$0</span>
                        </div>
                    </div>

                </div>

                <!-- Monthly Settlement Banner Footer -->
                <div class="p-4 bg-slate-900 text-white flex flex-col sm:flex-row items-center justify-between gap-3 text-sm">
                    <div class="flex items-center gap-2">
                        <i class="fa-solid fa-calculator text-amber-400 text-lg"></i>
                        <span>當月對帳結算：收入 <span id="summary-inc-text" class="text-emerald-400 font-bold">$0</span> - 支出 <span id="summary-exp-text" class="text-rose-400 font-bold">$0</span> =</span>
                    </div>
                    <div class="text-base font-bold px-4 py-1.5 rounded-xl bg-slate-800 border border-slate-700 shadow-inner" id="summary-result-text">
                        淨利 $0
                    </div>
                </div>
            </div>

            <!-- Visual Charts & Cumulative Analysis Row -->
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- Monthly Trend Chart -->
                <div class="lg:col-span-2 bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                            <i class="fa-solid fa-chart-column text-blue-600"></i> 各月份損益走勢圖
                        </h3>
                    </div>
                    <div class="h-64">
                        <canvas id="pnlTrendChart"></canvas>
                    </div>
                </div>

                <!-- Cumulative Surplus / Dorm Breakdown -->
                <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-col">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                            <i class="fa-solid fa-coins text-amber-500"></i> 篩選範圍累計盈餘
                        </h3>
                    </div>
                    <div class="bg-slate-50 p-4 rounded-xl border border-slate-100 text-center mb-4">
                        <p class="text-xs text-slate-500 font-medium mb-1">歷史累計淨盈餘 (全期總和)</p>
                        <h4 class="text-2xl font-extrabold" id="kpi-cumulative-surplus">$0</h4>
                    </div>
                    <div class="flex-1 h-48 flex items-center justify-center">
                        <canvas id="dormDistributionChart"></canvas>
                    </div>
                </div>
            </div>

        </div>

        <!-- TAB 2: Full Transactions Table & Management -->
        <div id="view-records" class="space-y-4 hidden">
            <!-- Table Action Top Bar -->
            <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap items-center justify-between gap-3">
                <div class="flex items-center gap-2">
                    <button onclick="openTransactionModal()" class="bg-blue-600 hover:bg-blue-500 text-white text-xs sm:text-sm px-4 py-2.5 rounded-xl font-medium flex items-center gap-2 transition shadow-sm">
                        <i class="fa-solid fa-plus"></i> 新增單筆收支
                    </button>
                    <span class="text-xs text-slate-500 ml-2" id="records-count-label">共 0 筆紀錄</span>
                </div>
                <div class="flex items-center gap-2">
                    <input type="text" id="table-search" oninput="renderRecordsTable()" placeholder="搜尋關鍵字 (客戶/宿舍/內容/備註)..." class="text-xs sm:text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500 w-48 sm:w-64 bg-slate-50">
                    <button onclick="exportFilteredExcel()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs sm:text-sm px-3.5 py-2.5 rounded-xl border border-slate-300 font-medium flex items-center gap-1.5 transition">
                        <i class="fa-solid fa-file-export text-slate-500"></i> 匯出 Excel
                    </button>
                </div>
            </div>

            <!-- Full Transactions Data Table -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-50 text-slate-500 font-semibold text-xs uppercase border-b border-slate-200">
                            <tr>
                                <th class="p-3.5">日期 (民國/西元)</th>
                                <th class="p-3.5">區域</th>
                                <th class="p-3.5">客戶 (雇主)</th>
                                <th class="p-3.5">宿舍名稱</th>
                                <th class="p-3.5">類型</th>
                                <th class="p-3.5">內容 (人員 / 支出品項)</th>
                                <th class="p-3.5 text-right">金額</th>
                                <th class="p-3.5">備註</th>
                                <th class="p-3.5 text-center">操作</th>
                            </tr>
                        </thead>
                        <tbody id="full-records-tbody" class="divide-y divide-slate-200 text-xs sm:text-sm">
                            <!-- Populated dynamically -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

    </main>

    <!-- Modal for Adding / Editing Transaction -->
    <div id="tx-modal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-md w-full overflow-hidden border border-slate-100 animate-in fade-in zoom-in duration-200">
            <div class="bg-slate-900 text-white p-4 flex items-center justify-between">
                <h3 class="font-bold text-base" id="modal-title">新增記帳明細</h3>
                <button onclick="closeTransactionModal()" class="text-slate-400 hover:text-white text-lg">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>
            <form id="tx-form" onsubmit="handleFormSubmit(event)" class="p-5 space-y-3.5">
                <input type="hidden" id="tx-id">
                
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">日期 (支援民國年 115/1/1 或西元年 2026/01/01)</label>
                    <input type="text" id="form-date" placeholder="115/01/01" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                </div>

                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 mb-1">收支類型</label>
                        <select id="form-type" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500 bg-white">
                            <option value="收入">收入 (如人員租金/水電)</option>
                            <option value="支出">支出 (如水管/洗衣機/維修)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 mb-1">金額 (NT$)</label>
                        <input type="number" id="form-amount" min="0" step="1" placeholder="1000" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                    </div>
                </div>

                <div class="grid grid-cols-3 gap-2">
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 mb-1">區域</label>
                        <input type="text" id="form-region" placeholder="建國區" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 mb-1">客戶(雇主)</label>
                        <input type="text" id="form-client" placeholder="APPLE" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 mb-1">宿舍名稱</label>
                        <input type="text" id="form-dorm" placeholder="建國" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">內容 (人員姓名 或 支出品項)</label>
                    <input type="text" id="form-item" placeholder="例：A人 (收入) 或 水管維修 (支出)" required class="w-full text-sm border-slate-300 rounded-xl p-2.5 border focus:ring-blue-500">
                </div>

                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">備註說明 (選填)</label>
                    <textarea id="form-note" rows="2" placeholder="相關備註..." class="w-full text-sm border-slate-300 rounded-xl p-2 border focus:ring-blue-500"></textarea>
                </div>

                <div class="pt-2 flex items-center justify-end gap-2">
                    <button type="button" onclick="closeTransactionModal()" class="px-4 py-2 border border-slate-300 rounded-xl text-xs font-medium text-slate-600 hover:bg-slate-50">
                        取消
                    </button>
                    <button type="submit" class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-medium transition shadow-sm">
                        儲存紀錄
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- Application JavaScript Logic -->
    <script>
        // Default Sample Data (Including exact scenario: 115年1月, APPLE/建國宿舍, Income A/B/C vs Expense Water pipe/Washing machine)
        const initialData = [
            // User requested scenario (115/1/1)
            { id: "1", date: "115/1/1", region: "建國區", client: "APPLE", dorm: "建國", type: "收入", item: "A人", amount: 1000, note: "月租金" },
            { id: "2", date: "115/1/1", region: "建國區", client: "APPLE", dorm: "建國", type: "收入", item: "B人", amount: 1000, note: "月租金" },
            { id: "3", date: "115/1/1", region: "建國區", client: "APPLE", dorm: "建國", type: "收入", item: "C人", amount: 1000, note: "月租金" },
            { id: "4", date: "115/1/5", region: "建國區", client: "APPLE", dorm: "建國", type: "支出", item: "水管", amount: 1000, note: "水管堵塞維修" },
            { id: "5", date: "115/1/10", region: "建國區", client: "APPLE", dorm: "建國", type: "支出", item: "洗衣機", amount: 5000, note: "更換全新洗衣機" },
            
            // Previous Month Sample (114/12/01)
            { id: "6", date: "114/12/1", region: "建國區", client: "APPLE", dorm: "建國", type: "收入", item: "A人", amount: 1000, note: "12月租金" },
            { id: "7", date: "114/12/1", region: "建國區", client: "APPLE", dorm: "建國", type: "收入", item: "B人", amount: 1000, note: "12月租金" },
            { id: "8", date: "114/12/15", region: "建國區", client: "APPLE", dorm: "建國", type: "支出", item: "水電費", amount: 800, note: "公共用電" },

            // Another Client / Dormitory
            { id: "9", date: "115/1/8", region: "龜山區", client: "台積電", dorm: "第一宿舍", type: "收入", item: "張先生", amount: 3000, note: "1月房租" },
            { id: "10", date: "115/1/12", region: "龜山區", client: "台積電", dorm: "第一宿舍", type: "支出", item: "冷氣清洗", amount: 1500, note: "定期保養" }
        ];

        // State Management
        let transactions = JSON.parse(localStorage.getItem('dorm_system_records')) || initialData;
        let trendChartInstance = null;
        let pieChartInstance = null;

        // Initialize App
        window.addEventListener('DOMContentLoaded', () => {
            populateFilterOptions();
            applyFilters();
            initCharts();
        });

        // Save State
        function saveData() {
            localStorage.setItem('dorm_system_records', JSON.stringify(transactions));
        }

        // Helper: Convert ROC or AD Date to standard Month Key (e.g., "115年1月")
        function parseYearMonthKey(dateStr) {
            if (!dateStr) return "未知月份";
            let clean = dateStr.toString().trim().replace(/-/g, '/');
            let parts = clean.split('/');
            
            if (parts.length >= 2) {
                let y = parseInt(parts[0], 10);
                let m = parseInt(parts[1], 10);
                
                // If Western Year (e.g., 2026 -> 115)
                if (y > 1900) {
                    y = y - 1911;
                }
                return `${y}年${m}月`;
            }
            return "其他";
        }

        // Toast Notification System
        function showToast(message, type = 'info') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            
            const bgClass = type === 'success' ? 'bg-emerald-600' : type === 'error' ? 'bg-rose-600' : 'bg-slate-800';
            const icon = type === 'success' ? 'fa-circle-check' : type === 'error' ? 'fa-triangle-exclamation' : 'fa-circle-info';
            
            toast.className = `${bgClass} text-white text-xs sm:text-sm px-4 py-3 rounded-xl shadow-lg flex items-center gap-2 pointer-events-auto transition-all duration-300 opacity-0 transform translate-y-2`;
            toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
            
            container.appendChild(toast);
            
            requestAnimationFrame(() => {
                toast.classList.remove('opacity-0', 'translate-y-2');
            });

            setTimeout(() => {
                toast.classList.add('opacity-0', 'translate-y-2');
                setTimeout(() => toast.remove(), 300);
            }, 3000);
        }

        // Switch Navigation Tabs
        function switchTab(tabKey) {
            const btnPnl = document.getElementById('tab-pnl');
            const btnRecords = document.getElementById('tab-records');
            const viewPnl = document.getElementById('view-pnl');
            const viewRecords = document.getElementById('view-records');

            if (tabKey === 'pnl') {
                btnPnl.classList.add('active');
                btnRecords.classList.remove('active');
                btnRecords.classList.add('text-slate-500');
                viewPnl.classList.remove('hidden');
                viewRecords.classList.add('hidden');
            } else {
                btnRecords.classList.add('active');
                btnPnl.classList.remove('active');
                btnPnl.classList.add('text-slate-500');
                viewRecords.classList.remove('hidden');
                viewPnl.classList.add('hidden');
                renderRecordsTable();
            }
        }

        // Populate Dropdown Options
        function populateFilterOptions() {
            const monthSelect = document.getElementById('filter-month');
            const regionSelect = document.getElementById('filter-region');
            const clientSelect = document.getElementById('filter-client');
            const dormSelect = document.getElementById('filter-dorm');

            const monthsSet = new Set();
            const regionsSet = new Set();
            const clientsSet = new Set();
            const dormsSet = new Set();

            transactions.forEach(t => {
                if (t.date) monthsSet.add(parseYearMonthKey(t.date));
                if (t.region) regionsSet.add(t.region);
                if (t.client) clientsSet.add(t.client);
                if (t.dorm) dormsSet.add(t.dorm);
            });

            // Populate Months
            const currentMonthVal = monthSelect.value;
            monthSelect.innerHTML = '';
            
            const sortedMonths = Array.from(monthsSet).sort((a, b) => b.localeCompare(a));
            if (sortedMonths.length === 0) sortedMonths.push("115年1月");

            sortedMonths.forEach(m => {
                const opt = document.createElement('option');
                opt.value = m;
                opt.textContent = m;
                monthSelect.appendChild(opt);
            });

            if (sortedMonths.includes("115年1月")) {
                monthSelect.value = "115年1月";
            } else if (currentMonthVal && sortedMonths.includes(currentMonthVal)) {
                monthSelect.value = currentMonthVal;
            }

            // Populate Regions
            regionSelect.innerHTML = '<option value="ALL">全部區域</option>';
            regionsSet.forEach(r => {
                const opt = document.createElement('option');
                opt.value = r;
                opt.textContent = r;
                regionSelect.appendChild(opt);
            });

            // Populate Clients
            clientSelect.innerHTML = '<option value="ALL">全部客戶 (雇主)</option>';
            clientsSet.forEach(c => {
                const opt = document.createElement('option');
                opt.value = c;
                opt.textContent = c;
                clientSelect.appendChild(opt);
            });

            // Populate Dorms
            dormSelect.innerHTML = '<option value="ALL">全部宿舍</option>';
            dormsSet.forEach(d => {
                const opt = document.createElement('option');
                opt.value = d;
                opt.textContent = d;
                dormSelect.appendChild(opt);
            });
        }

        // Apply Filters & Calculate P&L
        function applyFilters() {
            const selectedMonth = document.getElementById('filter-month').value;
            const selectedRegion = document.getElementById('filter-region').value;
            const selectedClient = document.getElementById('filter-client').value;
            const selectedDorm = document.getElementById('filter-dorm').value;

            // Update Labels
            document.getElementById('selected-month-label-1').textContent = selectedMonth;
            document.getElementById('selected-month-label-2').textContent = selectedMonth;
            document.getElementById('selected-month-label-3').textContent = selectedMonth;
            document.getElementById('breakdown-title-month').textContent = selectedMonth;

            // Filter Dataset for Current Month
            const filteredMonth = transactions.filter(t => {
                const matchMonth = parseYearMonthKey(t.date) === selectedMonth;
                const matchRegion = selectedRegion === 'ALL' || t.region === selectedRegion;
                const matchClient = selectedClient === 'ALL' || t.client === selectedClient;
                const matchDorm = selectedDorm === 'ALL' || t.dorm === selectedDorm;
                return matchMonth && matchRegion && matchClient && matchDorm;
            });

            // Filter Dataset for Cumulative Overall
            const filteredCumulative = transactions.filter(t => {
                const matchRegion = selectedRegion === 'ALL' || t.region === selectedRegion;
                const matchClient = selectedClient === 'ALL' || t.client === selectedClient;
                const matchDorm = selectedDorm === 'ALL' || t.dorm === selectedDorm;
                return matchRegion && matchClient && matchDorm;
            });

            // Calculate Monthly Metrics
            const incomeList = filteredMonth.filter(t => t.type === '收入');
            const expenseList = filteredMonth.filter(t => t.type === '支出');

            const totalIncome = incomeList.reduce((sum, t) => sum + Number(t.amount || 0), 0);
            const totalExpense = expenseList.reduce((sum, t) => sum + Number(t.amount || 0), 0);
            const netPnL = totalIncome - totalExpense;

            // Calculate Cumulative Overall Surplus
            const cumulativeSurplus = filteredCumulative.reduce((sum, t) => {
                return t.type === '收入' ? sum + Number(t.amount || 0) : sum - Number(t.amount || 0);
            }, 0);

            // Update KPI Cards
            document.getElementById('kpi-monthly-income').textContent = `$${totalIncome.toLocaleString()}`;
            document.getElementById('kpi-income-count').textContent = `共 ${incomeList.length} 筆收入項目`;

            document.getElementById('kpi-monthly-expense').textContent = `$${totalExpense.toLocaleString()}`;
            document.getElementById('kpi-expense-count').textContent = `共 ${expenseList.length} 筆支出項目`;

            // Update Cumulative Card
            const cumuElem = document.getElementById('kpi-cumulative-surplus');
            if (cumulativeSurplus < 0) {
                cumuElem.className = "text-2xl font-extrabold text-rose-600";
                cumuElem.textContent = `-$${Math.abs(cumulativeSurplus).toLocaleString()}`;
            } else {
                cumuElem.className = "text-2xl font-extrabold text-emerald-600";
                cumuElem.textContent = `+$${cumulativeSurplus.toLocaleString()}`;
            }

            // P&L KPI Card Dynamics
            const pnlCard = document.getElementById('kpi-pnl-card');
            const pnlVal = document.getElementById('kpi-monthly-pnl');
            const pnlStatus = document.getElementById('kpi-pnl-status');
            const pnlIcon = document.getElementById('kpi-pnl-icon');

            if (netPnL < 0) {
                pnlCard.className = "bg-rose-50/70 p-5 rounded-2xl shadow-sm border border-rose-200 flex items-center justify-between";
                pnlVal.className = "text-2xl font-extrabold text-rose-600";
                pnlVal.textContent = `-$${Math.abs(netPnL).toLocaleString()}`;
                pnlStatus.className = "text-xs font-semibold text-rose-600";
                pnlStatus.textContent = `當月虧損 $${Math.abs(netPnL).toLocaleString()}`;
                pnlIcon.className = "w-12 h-12 bg-rose-100 text-rose-600 rounded-2xl flex items-center justify-center text-xl font-bold";
                pnlIcon.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i>`;
            } else if (netPnL > 0) {
                pnlCard.className = "bg-emerald-50/70 p-5 rounded-2xl shadow-sm border border-emerald-200 flex items-center justify-between";
                pnlVal.className = "text-2xl font-extrabold text-emerald-600";
                pnlVal.textContent = `+$${netPnL.toLocaleString()}`;
                pnlStatus.className = "text-xs font-semibold text-emerald-600";
                pnlStatus.textContent = `當月淨盈餘 $${netPnL.toLocaleString()}`;
                pnlIcon.className = "w-12 h-12 bg-emerald-100 text-emerald-600 rounded-2xl flex items-center justify-center text-xl font-bold";
                pnlIcon.innerHTML = `<i class="fa-solid fa-piggy-bank"></i>`;
            } else {
                pnlCard.className = "bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between";
                pnlVal.className = "text-2xl font-bold text-slate-700";
                pnlVal.textContent = `$0`;
                pnlStatus.className = "text-xs text-slate-500";
                pnlStatus.textContent = "收支完全打平";
                pnlIcon.className = "w-12 h-12 bg-slate-100 text-slate-600 rounded-2xl flex items-center justify-center text-xl font-bold";
                pnlIcon.innerHTML = `<i class="fa-solid fa-scale-balanced"></i>`;
            }

            // Render Detailed Side-by-Side Tables
            renderBreakdownTables(incomeList, expenseList, totalIncome, totalExpense, netPnL);

            // Update Charts
            updateCharts();
        }

        // Render Side-by-Side Income vs Expense Tables
        function renderBreakdownTables(incomeList, expenseList, totalIncome, totalExpense, netPnL) {
            // Income Table
            const incTbody = document.getElementById('monthly-income-tbody');
            document.getElementById('income-item-count').textContent = `${incomeList.length} 筆`;
            document.getElementById('income-subtotal').textContent = `$${totalIncome.toLocaleString()}`;

            if (incomeList.length === 0) {
                incTbody.innerHTML = `<tr><td colspan="3" class="py-6 text-center text-slate-400">當月無收入項目</td></tr>`;
            } else {
                incTbody.innerHTML = incomeList.map(t => `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="py-2.5 font-semibold text-slate-800">
                            ${escapeHtml(t.item)}
                            ${t.note ? `<span class="block text-[10px] text-slate-400 font-normal">${escapeHtml(t.note)}</span>` : ''}
                        </td>
                        <td class="py-2.5 text-slate-500">
                            <span class="bg-slate-100 px-1.5 py-0.5 rounded text-[10px] font-medium">${escapeHtml(t.client || '無')}/${escapeHtml(t.dorm || '全棟')}</span>
                        </td>
                        <td class="py-2.5 text-right font-bold text-emerald-600">
                            +$${Number(t.amount).toLocaleString()}
                        </td>
                    </tr>
                `).join('');
            }

            // Expense Table
            const expTbody = document.getElementById('monthly-expense-tbody');
            document.getElementById('expense-item-count').textContent = `${expenseList.length} 筆`;
            document.getElementById('expense-subtotal').textContent = `$${totalExpense.toLocaleString()}`;

            if (expenseList.length === 0) {
                expTbody.innerHTML = `<tr><td colspan="3" class="py-6 text-center text-slate-400">當月無支出項目</td></tr>`;
            } else {
                expTbody.innerHTML = expenseList.map(t => `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="py-2.5 font-semibold text-slate-800">
                            ${escapeHtml(t.item)}
                            ${t.note ? `<span class="block text-[10px] text-slate-400 font-normal">${escapeHtml(t.note)}</span>` : ''}
                        </td>
                        <td class="py-2.5 text-slate-500">
                            <span class="bg-slate-100 px-1.5 py-0.5 rounded text-[10px] font-medium">${escapeHtml(t.client || '無')}/${escapeHtml(t.dorm || '全棟')}</span>
                        </td>
                        <td class="py-2.5 text-right font-bold text-rose-600">
                            -$${Number(t.amount).toLocaleString()}
                        </td>
                    </tr>
                `).join('');
            }

            // Summary Banner Updates
            document.getElementById('summary-inc-text').textContent = `$${totalIncome.toLocaleString()}`;
            document.getElementById('summary-exp-text').textContent = `$${totalExpense.toLocaleString()}`;
            
            const summaryRes = document.getElementById('summary-result-text');
            if (netPnL < 0) {
                summaryRes.className = "text-base font-bold px-4 py-1.5 rounded-xl bg-rose-950/90 border border-rose-800 text-rose-300";
                summaryRes.textContent = `當月虧損 $${Math.abs(netPnL).toLocaleString()}`;
            } else if (netPnL > 0) {
                summaryRes.className = "text-base font-bold px-4 py-1.5 rounded-xl bg-emerald-950/90 border border-emerald-800 text-emerald-300";
                summaryRes.textContent = `當月淨利 $${netPnL.toLocaleString()}`;
            } else {
                summaryRes.className = "text-base font-bold px-4 py-1.5 rounded-xl bg-slate-800 border border-slate-700 text-slate-300";
                summaryRes.textContent = `收支完全打平 $0`;
            }
        }

        // Render Full Records Data Table
        function renderRecordsTable() {
            const query = document.getElementById('table-search').value.toLowerCase().trim();
            const tbody = document.getElementById('full-records-tbody');
            
            const filtered = transactions.filter(t => {
                const matchQuery = !query || 
                    (t.item && t.item.toLowerCase().includes(query)) ||
                    (t.dorm && t.dorm.toLowerCase().includes(query)) ||
                    (t.client && t.client.toLowerCase().includes(query)) ||
                    (t.region && t.region.toLowerCase().includes(query)) ||
                    (t.note && t.note.toLowerCase().includes(query));
                return matchQuery;
            });

            document.getElementById('records-count-label').textContent = `共 ${filtered.length} 筆紀錄`;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="9" class="p-8 text-center text-slate-400">尚無符合條件的收支明細</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(t => {
                const isInc = t.type === '收入';
                return `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="p-3.5 font-medium text-slate-700">${escapeHtml(t.date)}</td>
                        <td class="p-3.5 text-slate-500">${escapeHtml(t.region || '-')}</td>
                        <td class="p-3.5 text-slate-700 font-medium">${escapeHtml(t.client || '-')}</td>
                        <td class="p-3.5"><span class="bg-slate-100 text-slate-700 px-2 py-1 rounded-md text-xs font-medium">${escapeHtml(t.dorm)}</span></td>
                        <td class="p-3.5">
                            <span class="${isInc ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'} px-2.5 py-1 rounded-full text-xs font-bold">
                                ${t.type}
                            </span>
                        </td>
                        <td class="p-3.5 font-bold text-slate-900">${escapeHtml(t.item)}</td>
                        <td class="p-3.5 text-right font-extrabold ${isInc ? 'text-emerald-600' : 'text-rose-600'}">
                            ${isInc ? '+' : '-'}$${Number(t.amount).toLocaleString()}
                        </td>
                        <td class="p-3.5 text-slate-400 text-xs">${escapeHtml(t.note || '-')}</td>
                        <td class="p-3.5 text-center space-x-1">
                            <button onclick="editTransaction('${t.id}')" class="text-blue-600 hover:text-blue-800 p-1.5 rounded hover:bg-blue-50">
                                <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button onclick="deleteTransaction('${t.id}')" class="text-rose-500 hover:text-rose-700 p-1.5 rounded hover:bg-rose-50">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        // Reset All Dropdown Filters
        function resetFilters() {
            document.getElementById('filter-region').value = 'ALL';
            document.getElementById('filter-client').value = 'ALL';
            document.getElementById('filter-dorm').value = 'ALL';
            applyFilters();
            showToast("已重置所有篩選條件", "info");
        }

        // Transaction Modal Controls
        function openTransactionModal(editId = null) {
            const modal = document.getElementById('tx-modal');
            const title = document.getElementById('modal-title');
            const form = document.getElementById('tx-form');
            form.reset();

            if (editId) {
                const item = transactions.find(t => t.id === editId);
                if (item) {
                    document.getElementById('tx-id').value = item.id;
                    document.getElementById('form-date').value = item.date;
                    document.getElementById('form-type').value = item.type;
                    document.getElementById('form-amount').value = item.amount;
                    document.getElementById('form-region').value = item.region || '建國區';
                    document.getElementById('form-client').value = item.client || 'APPLE';
                    document.getElementById('form-dorm').value = item.dorm || '建國';
                    document.getElementById('form-item').value = item.item;
                    document.getElementById('form-note').value = item.note || '';
                    title.textContent = "編輯記帳明細";
                }
            } else {
                document.getElementById('tx-id').value = '';
                document.getElementById('form-date').value = "115/1/1";
                document.getElementById('form-region').value = "建國區";
                document.getElementById('form-client').value = "APPLE";
                document.getElementById('form-dorm').value = "建國";
                title.textContent = "新增單筆記帳明細";
            }

            modal.classList.remove('hidden');
        }

        function closeTransactionModal() {
            document.getElementById('tx-modal').classList.add('hidden');
        }

        function handleFormSubmit(e) {
            e.preventDefault();
            const id = document.getElementById('tx-id').value;
            const date = document.getElementById('form-date').value.trim();
            const type = document.getElementById('form-type').value;
            const amount = parseFloat(document.getElementById('form-amount').value) || 0;
            const region = document.getElementById('form-region').value.trim();
            const client = document.getElementById('form-client').value.trim();
            const dorm = document.getElementById('form-dorm').value.trim();
            const item = document.getElementById('form-item').value.trim();
            const note = document.getElementById('form-note').value.trim();

            if (id) {
                const idx = transactions.findIndex(t => t.id === id);
                if (idx !== -1) {
                    transactions[idx] = { id, date, type, amount, region, client, dorm, item, note };
                    showToast("紀錄已更新", "success");
                }
            } else {
                const newTx = {
                    id: Date.now().toString(),
                    date, type, amount, region, client, dorm, item, note
                };
                transactions.unshift(newTx);
                showToast("已成功新增記帳明細", "success");
            }

            saveData();
            populateFilterOptions();
            applyFilters();
            renderRecordsTable();
            closeTransactionModal();
        }

        function editTransaction(id) {
            openTransactionModal(id);
        }

        function deleteTransaction(id) {
            if (confirm("確定要刪除這筆收支紀錄嗎？")) {
                transactions = transactions.filter(t => t.id !== id);
                saveData();
                populateFilterOptions();
                applyFilters();
                renderRecordsTable();
                showToast("紀錄已刪除", "info");
            }
        }

        // Render Chart.js Visuals
        function initCharts() {
            const ctxTrend = document.getElementById('pnlTrendChart').getContext('2d');
            const ctxPie = document.getElementById('dormDistributionChart').getContext('2d');

            trendChartInstance = new Chart(ctxTrend, {
                type: 'bar',
                data: { labels: [], datasets: [] },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'top' } },
                    scales: { y: { beginAtZero: true } }
                }
            });

            pieChartInstance = new Chart(ctxPie, {
                type: 'doughnut',
                data: { labels: [], datasets: [] },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom' } }
                }
            });

            updateCharts();
        }

        function updateCharts() {
            if (!trendChartInstance || !pieChartInstance) return;

            // Group by Month
            const monthMap = {};
            transactions.forEach(t => {
                const ym = parseYearMonthKey(t.date);
                if (!monthMap[ym]) monthMap[ym] = { income: 0, expense: 0 };
                if (t.type === '收入') monthMap[ym].income += Number(t.amount || 0);
                if (t.type === '支出') monthMap[ym].expense += Number(t.amount || 0);
            });

            const months = Object.keys(monthMap).reverse();
            const incomeData = months.map(m => monthMap[m].income);
            const expenseData = months.map(m => monthMap[m].expense);

            trendChartInstance.data = {
                labels: months,
                datasets: [
                    { label: '總收入', data: incomeData, backgroundColor: '#10b981', borderRadius: 6 },
                    { label: '總支出', data: expenseData, backgroundColor: '#f43f5e', borderRadius: 6 }
                ]
            };
            trendChartInstance.update();

            // Group by Dormitory
            const dormMap = {};
            transactions.forEach(t => {
                const dorm = t.dorm || '全棟';
                if (!dormMap[dorm]) dormMap[dorm] = 0;
                if (t.type === '收入') dormMap[dorm] += Number(t.amount || 0);
                else dormMap[dorm] -= Number(t.amount || 0);
            });

            const dormLabels = Object.keys(dormMap);
            const dormValues = dormLabels.map(d => dormMap[d]);

            pieChartInstance.data = {
                labels: dormLabels,
                datasets: [{
                    data: dormValues,
                    backgroundColor: ['#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981']
                }]
            };
            pieChartInstance.update();
        }

        // Excel Template Generation and Download
        function downloadExcelTemplate() {
            const templateData = [
                { "日期": "115/1/1", "客戶": "APPLE", "宿舍名稱": "建國", "內容": "A人", "金額": 1000, "類型": "收入", "備註": "租金/水電" },
                { "日期": "115/1/1", "客戶": "APPLE", "宿舍名稱": "建國", "內容": "B人", "金額": 1000, "類型": "收入", "備註": "租金/水電" },
                { "日期": "115/1/1", "客戶": "APPLE", "宿舍名稱": "建國", "內容": "C人", "金額": 1000, "類型": "收入", "備註": "租金/水電" },
                { "日期": "115/1/5", "客戶": "APPLE", "宿舍名稱": "建國", "內容": "水管", "金額": 1000, "類型": "支出", "備註": "水管維修" },
                { "日期": "115/1/10", "客戶": "APPLE", "宿舍名稱": "建國", "內容": "洗衣機", "金額": 5000, "類型": "支出", "備註": "洗衣機維修" }
            ];

            const ws = XLSX.utils.json_to_sheet(templateData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "宿舍收支對帳範本");
            XLSX.writeFile(wb, "宿舍智慧收支對帳範本_民國年.xlsx");
            showToast("已下載 Excel 匯入範本", "success");
        }

        // Excel Import Reader
        function importExcel(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                    let importedCount = 0;
                    jsonData.forEach(row => {
                        const date = row['日期'] || row['Date'] || '115/1/1';
                        const client = row['客戶'] || row['客戶(雇主)'] || row['雇主'] || 'APPLE';
                        const dorm = row['宿舍名稱'] || row['宿舍'] || '建國';
                        const region = row['區域'] || '建國區';
                        const item = row['內容'] || row['內容(人員/品項)'] || row['品項'] || row['人員'] || '未命名項目';
                        const amount = parseFloat(row['金額'] || row['Amount'] || 0);
                        
                        // Infer Type if missing: Default to 支出 if item contains repair words, else check 類型 column
                        let rawType = row['類型'] || row['收支類型'] || '';
                        let type = '收入';
                        if (rawType.includes('支') || item.includes('維修') || item.includes('水管') || item.includes('洗衣機') || item.includes('冷氣')) {
                            type = '支出';
                        }

                        const note = row['備註'] || row['Note'] || '';

                        if (amount > 0) {
                            transactions.push({
                                id: Date.now().toString() + Math.random().toString().substr(2, 4),
                                date: date.toString(),
                                region: region.toString(),
                                client: client.toString(),
                                dorm: dorm.toString(),
                                type,
                                item: item.toString(),
                                amount,
                                note: note.toString()
                            });
                            importedCount++;
                        }
                    });

                    saveData();
                    populateFilterOptions();
                    applyFilters();
                    renderRecordsTable();
                    showToast(`成功匯入 ${importedCount} 筆對帳明細`, "success");
                } catch (err) {
                    console.error(err);
                    showToast("Excel 讀取失敗，請確認欄位格式", "error");
                }
            };
            reader.readAsArrayBuffer(file);
            event.target.value = '';
        }

        // Export Filtered Table to Excel
        function exportFilteredExcel() {
            const wsData = transactions.map(t => ({
                "日期": t.date,
                "區域": t.region,
                "客戶(雇主)": t.client,
                "宿舍名稱": t.dorm,
                "類型": t.type,
                "內容(人員/品項)": t.item,
                "金額": t.amount,
                "備註": t.note
            }));

            const ws = XLSX.utils.json_to_sheet(wsData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "宿舍對帳流水帳");
            XLSX.writeFile(wb, `宿舍對帳明細紀錄_${new Date().toISOString().slice(0,10)}.xlsx`);
            showToast("已成功匯出 Excel 檔", "success");
        }

        // Utility: HTML Escaping
        function escapeHtml(str) {
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }
    </script>
</body>
</html>
