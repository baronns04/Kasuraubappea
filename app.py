import io
import json
import re
from datetime import datetime
import pandas as pd
import requests
from flask import Flask, jsonify, render_template_string, request

app = Flask(__name__)

# ID Google Spreadsheet Laporan Keuangan Surau Asyura Bappeda
DEFAULT_SHEET_ID = "1FYtS4cvEUQpXVYVOmWDRJm0f64KhCWDnDxn_ixPgnaU"
DEFAULT_Sheets_Names = ["Juli", "Agustus", "September"]


def clean_currency(val):
    """Membersihkan format angka/rupiah dari sel Spreadsheet menjadi integer."""
    if pd.isna(val):
        return 0
    s = str(val).strip()
    if not s or s.lower() in ["nan", "none", "-"]:
        return 0
    # Hapus Rp, spasi, dan karakter non-angka kecuali titik/koma desimal
    s = re.sub(r"[^\d,\.]", "", s)
    if not s:
        return 0
    # Jika akhiran ,00 atau .00 (desimal)
    if re.search(r"[,\.]00$", s):
        s = s[:-3]
    s = re.sub(r"[,\.]", "", s)
    try:
        return int(float(s))
    except ValueError:
        return 0


def parse_sheet_dataframe(df, month_name):
    """Mengekstrak baris transaksi dari masing-masing tab/sheet bulanan."""
    header_idx = None
    for idx, row in df.iterrows():
        row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if any("tanggal" in v for v in row_vals) and any("uraian" in v for v in row_vals):
            header_idx = idx
            break

    if header_idx is None:
        return []

    data_df = df.iloc[header_idx + 1 :].copy()
    transactions = []

    for _, row in data_df.iterrows():
        vals = list(row.values)
        if len(vals) < 5:
            continue

        tanggal = str(vals[1]).strip() if pd.notna(vals[1]) else ""
        uraian = str(vals[2]).strip() if pd.notna(vals[2]) else ""

        if not uraian or uraian.lower() in ["nan", "none"]:
            continue

        # Abaikan baris saldo awal bulan lalu atau baris total
        uraian_low = uraian.lower()
        if "saldo bulan lalu" in uraian_low or uraian_low.startswith("jumlah") or uraian_low.startswith("total"):
            continue

        masuk = clean_currency(vals[3]) if len(vals) > 3 else 0
        keluar = clean_currency(vals[4]) if len(vals) > 4 else 0
        saldo = clean_currency(vals[5]) if len(vals) > 5 else 0

        if masuk == 0 and keluar == 0:
            continue

        kategori = "Pengeluaran" if keluar > 0 else (
            "Wirid Mingguan" if "wirid" in uraian_low else "Transfer Umum"
        )

        transactions.append({
            "bulan": f"{month_name} 2026" if "2026" not in month_name else month_name,
            "bulan_short": month_name.replace("2026", "").strip(),
            "tanggal": tanggal,
            "uraian": uraian,
            "masuk": masuk,
            "keluar": keluar,
            "saldo": saldo,
            "kategori": kategori,
        })

    return transactions


def fetch_spreadsheet_data(sheet_id=DEFAULT_SHEET_ID):
    """
    Mengambil data langsung dari Google Spreadsheet.
    Mencoba metode export XLSX (otomatis mendeteksi semua nama sheet/bulan),
    dan fallback ke CSV per-sheet jika diperlukan.
    """
    all_transactions = []
    months_order = []

    # 1. Coba baca seluruh workbook via export XLSX agar tab baru otomatis terbaca
    xlsx_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    try:
        resp = requests.get(xlsx_url, timeout=10)
        if resp.status_code == 200:
            excel_data = pd.read_excel(io.BytesIO(resp.content), sheet_name=None, header=None)
            for sheet_name, df in excel_data.items():
                txs = parse_sheet_dataframe(df, sheet_name)
                if txs:
                    label = f"{sheet_name} 2026" if "2026" not in sheet_name else sheet_name
                    months_order.append(label)
                    all_transactions.extend(txs)
    except Exception:
        pass

    # 2. Fallback ke endpoint gviz CSV jika XLSX gagal
    if not all_transactions:
        for sheet_name in DEFAULT_Sheets_Names:
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            try:
                resp = requests.get(csv_url, timeout=8)
                if resp.status_code == 200:
                    df = pd.read_csv(io.StringIO(resp.text), header=None)
                    txs = parse_sheet_dataframe(df, sheet_name)
                    if txs:
                        label = f"{sheet_name} 2026" if "2026" not in sheet_name else sheet_name
                        months_order.append(label)
                        all_transactions.extend(txs)
            except Exception:
                continue

    # Hitung ulang saldo berjalan (running balance) & ringkasan per bulan
    running_saldo = 0
    for tx in all_transactions:
        running_saldo += tx["masuk"] - tx["keluar"]
        if tx["saldo"] == 0:
            tx["saldo"] = running_saldo
        else:
            running_saldo = tx["saldo"]

    monthly_summary = []
    prev_net_income = None

    for m in months_order:
        m_txs = [t for t in all_transactions if t["bulan"] == m]
        pemasukan = sum(t["masuk"] for t in m_txs)
        wirid = sum(t["masuk"] for t in m_txs if t["kategori"] == "Wirid Mingguan")
        umum = sum(t["masuk"] for t in m_txs if t["kategori"] == "Transfer Umum")
        pengeluaran = sum(t["keluar"] for t in m_txs)
        saldo_akhir = m_txs[-1]["saldo"] if m_txs else 0
        vs_prev = (pemasukan - prev_net_income) if prev_net_income is not None else None
        prev_net_income = pemasukan

        monthly_summary.append({
            "bulan": m,
            "jumlah_tx": len(m_txs),
            "tx_masuk": sum(1 for t in m_txs if t["masuk"] > 0),
            "tx_keluar": sum(1 for t in m_txs if t["keluar"] > 0),
            "pemasukan": pemasukan,
            "wirid": wirid,
            "umum": umum,
            "pengeluaran": pengeluaran,
            "saldo_akhir": saldo_akhir,
            "vs_bulan_lalu": vs_prev,
        })

    total_masuk = sum(t["masuk"] for t in all_transactions)
    total_keluar = sum(t["keluar"] for t in all_transactions)
    total_wirid = sum(t["masuk"] for t in all_transactions if t["kategori"] == "Wirid Mingguan")
    total_umum = sum(t["masuk"] for t in all_transactions if t["kategori"] == "Transfer Umum")
    saldo_akhir_total = all_transactions[-1]["saldo"] if all_transactions else 0

    return {
        "sheet_id": sheet_id,
        "updated_at": datetime.now().strftime("%d %b %Y, %H.%M"),
        "periode": f"{months_order[0].replace(' 2026', '')} - {months_order[-1]}" if months_order else "Juli - September 2026",
        "months": months_order,
        "monthly_summary": monthly_summary,
        "transactions": all_transactions,
        "totals": {
            "saldo_akhir": saldo_akhir_total,
            "pemasukan": total_masuk,
            "pengeluaran": total_keluar,
            "wirid": total_wirid,
            "umum": total_umum,
            "jumlah_tx": len(all_transactions),
            "tx_masuk": sum(1 for t in all_transactions if t["masuk"] > 0),
            "tx_keluar": sum(1 for t in all_transactions if t["keluar"] > 0),
            "persen_wirid": round((total_wirid / total_masuk) * 100) if total_masuk > 0 else 0,
            "tertinggi": max((t["masuk"] for t in all_transactions), default=0),
        },
    }


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Syura Bappeda - Buku Kas & Progres Keuangan</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Inter', sans-serif; background-color: #f8fafc; color: #0f172a; }
    .mono { font-family: 'JetBrains Mono', monospace; }
    @media print {
      .no-print { display: none !important; }
      .print-only { display: block !important; }
      body { background: #ffffff !important; }
      .print-card { box-shadow: none !important; border: 1px solid #cbd5e1 !important; break-inside: avoid; }
    }
  </style>
</head>
<body class="min-h-screen">

  <!-- ==================== MODE 1: DASHBOARD UTAMA ==================== -->
  <div id="dashboardView" class="pb-14">
    <!-- Top Navbar -->
    <header class="bg-white border-b border-slate-200 sticky top-0 z-30 no-print">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl bg-emerald-700 flex items-center justify-center text-white font-bold text-lg shadow-sm">
            🕌
          </div>
          <div>
            <h1 class="font-bold text-slate-900 text-lg leading-tight">Syura Bappeda</h1>
            <p class="text-xs text-slate-500" id="headerPeriode">Buku Kas & Progres Keuangan (Juli - September 2026)</p>
          </div>
        </div>

        <div class="flex flex-wrap items-center gap-2 text-xs font-medium">
          <button onclick="switchTab('dashboard')" class="px-3 py-2 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200 font-semibold">
            📊 Buku Kas & Progress Bulanan
          </button>
          <button onclick="openSheetModal()" class="px-3 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700">
            🔗 Integrasi Spreadsheet
          </button>
          <button onclick="switchTab('tv')" class="px-3 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700">
            📺 Layar TV Jamaah
          </button>
          <button onclick="refreshData()" class="px-3 py-2 rounded-lg bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100">
            🔄 Sinkronkan Sheet
          </button>
          <button onclick="openWaModal()" class="px-3 py-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-50 text-slate-700">
            💬 Format WA
          </button>
          <button onclick="window.print()" class="px-3 py-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-50 text-slate-700">
            🖨️ Cetak
          </button>
          <a id="btnCatatKas" href="#" target="_blank" class="px-3.5 py-2 rounded-lg bg-emerald-800 hover:bg-emerald-900 text-white font-semibold shadow-sm">
            + Catat Kas
          </a>
        </div>
      </div>
    </header>

    <!-- Header Khusus Cetak (Print Only) -->
    <div class="hidden print-only max-w-7xl mx-auto px-6 pt-6 pb-4 border-b-2 border-slate-800 mb-6">
      <h1 class="text-2xl font-bold uppercase">Laporan Keuangan Kas Surau Asy-Syura Bappeda</h1>
      <p class="text-sm text-slate-600" id="printSubHeader">Periode: Juli - September 2026</p>
    </div>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 pt-6 space-y-6">
      <!-- Box 1: Progress & Perkembangan Kas Bulanan -->
      <section class="bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm print-card">
        <div class="flex flex-wrap items-center justify-between gap-4 mb-5">
          <div>
            <h2 class="text-base font-bold text-slate-900 flex items-center gap-2">
              📅 Progress & Perkembangan Kas Bulanan (<span id="labelPeriode1">Juli - September 2026</span>)
            </h2>
            <p class="text-xs text-slate-500 mt-0.5 no-print">
              Klik salah satu bulan di bawah untuk memfilter pembukuan, atau lihat perbandingan pertumbuhan kas.
            </p>
          </div>
          <div id="monthFilterButtons" class="flex flex-wrap gap-1.5 bg-slate-100 p-1 rounded-xl text-xs font-medium no-print">
            <!-- Diisi otomatis oleh JS -->
          </div>
        </div>

        <!-- Grid Kartu Bulanan -->
        <div id="monthlyCardsContainer" class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <!-- Diisi otomatis oleh JS -->
        </div>

        <!-- Bar Perbandingan Bulanan -->
        <div class="pt-4 border-t border-slate-100">
          <h3 class="text-xs font-semibold text-slate-700 mb-3">
            Tren Pemasukan Bulanan (Perbandingan Antar Bulan)
          </h3>
          <div id="monthlyBarsContainer" class="space-y-2.5">
            <!-- Diisi otomatis oleh JS -->
          </div>
        </div>
      </section>

      <!-- Header Baris Status Periode Aktif -->
      <div class="flex flex-wrap items-center justify-between text-xs text-slate-500 px-1">
        <div class="flex items-center gap-2">
          <span class="font-semibold text-slate-800" id="activeFilterLabel">🗓️ Seluruh Periode</span>
          <span>•</span>
          <span id="activeTxCount">0 transaksi tercatat</span>
        </div>
        <div>Sinkronisasi: <span id="syncTimestamp" class="font-medium text-slate-700">-</span></div>
      </div>

      <!-- Box 2: 4 KPI Utama -->
      <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <!-- Saldo Kas -->
        <div class="bg-slate-900 text-white rounded-2xl p-5 shadow-md print-card">
          <div class="flex items-center justify-between text-xs text-slate-300 mb-2">
            <span>Saldo Kas Surau Saat Ini</span>
            <span class="px-2 py-1 rounded-md bg-emerald-500/20 text-emerald-300">💳</span>
          </div>
          <div id="kpiSaldo" class="text-2xl sm:text-3xl font-extrabold text-emerald-400 mono my-1">Rp 0</div>
          <p class="text-xs text-slate-400 mt-2">Total kas bersih siap digunakan</p>
        </div>

        <!-- Total Pemasukan -->
        <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm print-card">
          <div class="flex items-center justify-between text-xs text-slate-500 mb-2">
            <span>Total Pemasukan</span>
            <span class="px-2 py-1 rounded-md bg-emerald-50 text-emerald-700">📈</span>
          </div>
          <div id="kpiPemasukan" class="text-2xl sm:text-3xl font-extrabold text-emerald-700 mono my-1">+Rp 0</div>
          <p id="kpiPemasukanSub" class="text-xs text-slate-500 mt-2">0 kali penerimaan dana</p>
        </div>

        <!-- Infaq Wirid Mingguan -->
        <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm print-card">
          <div class="flex items-center justify-between text-xs text-slate-500 mb-2">
            <span>Infaq Wirid Mingguan</span>
            <span class="px-2 py-1 rounded-md bg-amber-50 text-amber-600">🕌</span>
          </div>
          <div id="kpiWirid" class="text-2xl sm:text-3xl font-extrabold text-amber-600 mono my-1">Rp 0</div>
          <p id="kpiWiridSub" class="text-xs text-slate-500 mt-2">0% dari penerimaan · Umum: Rp 0</p>
        </div>

        <!-- Total Pengeluaran -->
        <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm print-card">
          <div class="flex items-center justify-between text-xs text-slate-500 mb-2">
            <span>Total Pengeluaran</span>
            <span class="px-2 py-1 rounded-md bg-rose-50 text-rose-600">📉</span>
          </div>
          <div id="kpiPengeluaran" class="text-2xl sm:text-3xl font-extrabold text-rose-600 mono my-1">-Rp 0</div>
          <p id="kpiPengeluaranSub" class="text-xs text-slate-500 mt-2">0 kali pengeluaran tercatat</p>
        </div>
      </section>

      <!-- Box 3: Visualisasi Aliran Dana & Komposisi -->
      <section class="grid grid-cols-1 lg:grid-cols-3 gap-4 no-print">
        <div class="lg:col-span-2 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <div class="flex flex-wrap items-center justify-between gap-2 mb-4">
            <div>
              <div class="flex items-center gap-2">
                <h3 class="font-bold text-slate-900 text-sm">Visualisasi Aliran Dana Kas</h3>
                <span id="badgeMasukCount" class="px-2 py-0.5 text-[11px] rounded-full bg-emerald-50 text-emerald-700 font-semibold">0 Pemasukan</span>
              </div>
              <p class="text-xs text-slate-500">Kurva tren fluktuasi nominal per transaksi</p>
            </div>
            <div class="flex items-center gap-2 text-xs">
              <span class="inline-flex items-center gap-1 text-slate-600"><span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span> Wirid Mingguan</span>
              <span class="inline-flex items-center gap-1 text-slate-600"><span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span> Transfer Umum</span>
            </div>
          </div>
          <div class="h-64">
            <canvas id="trendChart"></canvas>
          </div>
        </div>

        <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col justify-between">
          <div>
            <h3 class="font-bold text-slate-900 text-sm">Komposisi Sumber Kas</h3>
            <p class="text-xs text-slate-500 mb-4">Rasio pemasukan berdasarkan kegiatan</p>
          </div>
          <div class="flex items-center gap-5 my-auto">
            <div class="w-36 h-36 relative">
              <canvas id="donutChart"></canvas>
            </div>
            <div>
              <p class="text-xs text-slate-500">Total Terhimpun</p>
              <p id="donutTotal" class="text-xl font-extrabold text-slate-900 mono">Rp 0</p>
              <p id="donutTxCount" class="text-xs text-slate-400 mt-0.5">0 transaksi tercatat</p>
              <div class="mt-3 space-y-1 text-xs">
                <div class="flex items-center gap-1.5 text-slate-600">
                  <span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                  <span>Wirid: <b id="donutWiridPct">0%</b></span>
                </div>
                <div class="flex items-center gap-1.5 text-slate-600">
                  <span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span>
                  <span>Umum: <b id="donutUmumPct">0%</b></span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <!-- Box 4: Tabel Rincian Buku Kas -->
      <section class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm print-card">
        <div class="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div>
            <h3 class="font-bold text-slate-900 text-base">Rincian Mutasi Buku Kas Surau</h3>
            <p class="text-xs text-slate-500">Data tersinkronisasi otomatis dengan tabel Google Spreadsheet</p>
          </div>
          <input
            type="text"
            id="searchInput"
            oninput="renderTable()"
            placeholder="Cari uraian atau tanggal..."
            class="px-3.5 py-2 text-xs rounded-xl border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-600 no-print"
          />
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse text-xs sm:text-sm">
            <thead>
              <tr class="border-b border-slate-200 text-slate-500 bg-slate-50">
                <th class="py-3 px-3 font-semibold">No</th>
                <th class="py-3 px-3 font-semibold">Tanggal</th>
                <th class="py-3 px-3 font-semibold">Bulan</th>
                <th class="py-3 px-3 font-semibold">Uraian Transaksi</th>
                <th class="py-3 px-3 font-semibold text-right">Masuk (Rp)</th>
                <th class="py-3 px-3 font-semibold text-right">Keluar (Rp)</th>
                <th class="py-3 px-3 font-semibold text-right">Saldo (Rp)</th>
              </tr>
            </thead>
            <tbody id="kasTableBody" class="divide-y divide-slate-100">
              <!-- Diisi otomatis oleh JS -->
            </tbody>
          </table>
        </div>
      </section>
    </main>
  </div>

  <!-- ==================== MODE 2: LAYAR TV JAMAAH ==================== -->
  <div id="tvView" class="hidden min-h-screen bg-[#060d1a] text-white p-6 flex flex-col justify-between">
    <!-- Top Bar TV -->
    <div class="flex items-center justify-between border-b border-slate-800 pb-4">
      <div class="flex items-center gap-4">
        <button onclick="switchTab('dashboard')" class="px-3.5 py-2 rounded-xl bg-slate-800/90 hover:bg-slate-700 text-slate-200 text-xs font-medium flex items-center gap-1.5">
          ← Kembali ke Buku Kas
        </button>
        <div>
          <h1 class="text-xl font-extrabold tracking-tight flex items-center gap-2">
            🕌 KaSurau Asy Syura Bappeda
          </h1>
          <p class="text-xs text-emerald-400" id="tvSubPeriode">Transparansi Kas Jamaah Bappeda (Juli - September 2026)</p>
        </div>
      </div>
      <div class="flex items-center gap-4">
        <div class="text-right">
          <div id="tvClock" class="text-2xl font-extrabold mono tracking-wider">00.00.00</div>
          <div id="tvDate" class="text-xs text-slate-400">Minggu, 27 September 2026</div>
        </div>
        <button onclick="toggleFullScreen()" class="p-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300" title="Layar Penuh">
          ⛶
        </button>
      </div>
    </div>

    <!-- Main Grid TV -->
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 my-6 flex-1">
      <!-- Kolom Kiri (2/3) -->
      <div class="lg:col-span-2 flex flex-col justify-between gap-6">
        <!-- Hero Saldo Card -->
        <div class="rounded-3xl bg-gradient-to-br from-[#062c26] via-[#091e26] to-[#0b1528] border border-emerald-500/20 p-8 shadow-2xl">
          <div class="flex items-center justify-between text-xs font-bold uppercase tracking-widest text-emerald-400 mb-3">
            <span>✨ SALDO KAS TERKINI SURAU</span>
            <span id="tvHeroPeriode">PERIODE: JULI - SEPTEMBER 2026</span>
          </div>
          <div id="tvHeroSaldo" class="text-5xl sm:text-6xl font-extrabold text-emerald-400 mono tracking-tight my-4">
            Rp 0
          </div>
          <p class="text-sm text-slate-300">
            Alhamdulillah, kas bersih siap dimanfaatkan untuk operasional, kajian, dan kemakmuran surau.
          </p>
        </div>

        <!-- Ringkasan Bulanan Mini Cards -->
        <div id="tvMonthlyCards" class="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <!-- Diisi otomatis oleh JS -->
        </div>

        <!-- Ayat Kutipan Infaq -->
        <div class="rounded-2xl bg-[#0b1528] border border-slate-800 px-5 py-4 flex items-center gap-3 text-xs text-slate-300 italic">
          <span class="text-emerald-400 text-base not-italic">💚</span>
          <span>"Perumpamaan orang yang menginfakkan hartanya di jalan Allah seperti sebutir biji yang menumbuhkan tujuh tangkai, pada setiap tangkai ada seratus biji." (QS. Al-Baqarah: 261)</span>
        </div>
      </div>

      <!-- Kolom Kanan: Mutasi Kas Terkini -->
      <div class="rounded-3xl bg-[#0b1528] border border-slate-800 p-6 flex flex-col justify-between">
        <div>
          <h2 class="text-xs font-bold uppercase tracking-wider text-emerald-400 mb-4 flex items-center gap-2">
            📋 MUTASI KAS TERKINI (<span id="tvLatestMonthLabel">SEPTEMBER 2026</span>)
          </h2>
          <div id="tvMutationList" class="space-y-3">
            <!-- Diisi otomatis oleh JS -->
          </div>
        </div>
        <p class="text-[11px] text-center text-slate-500 pt-4 border-t border-slate-800/80 mt-4">
          Data terhubung langsung secara real-time dengan Spreadsheet Kas
        </p>
      </div>
    </div>
  </div>

  <!-- ==================== MODAL FORMAT WHATSAPP ==================== -->
  <div id="waModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white rounded-3xl max-w-lg w-full p-6 shadow-2xl border border-slate-200">
      <div class="flex items-start justify-between gap-4 pb-4 border-b border-slate-100">
        <div class="flex items-center gap-3">
          <div class="w-11 h-11 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center text-xl">
            💬
          </div>
          <div>
            <h3 class="font-bold text-slate-900 text-base">Format Laporan WhatsApp Jamaah</h3>
            <p class="text-xs text-slate-500">Siap dibagikan ke grup WhatsApp Bappeda / Jamaah Surau</p>
          </div>
        </div>
        <button onclick="closeWaModal()" class="text-slate-400 hover:text-slate-600 text-lg font-bold">✕</button>
      </div>

      <div class="my-4 flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2.5 text-xs">
        <span class="font-medium text-slate-700">📅 Pilih Bulan Laporan:</span>
        <select id="waMonthSelect" onchange="updateWaPreview()" class="bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-semibold text-slate-800 focus:outline-none">
          <option value="ALL">Kumulatif Seluruh Bulan</option>
        </select>
      </div>

      <pre id="waTextPreview" class="bg-[#0b1324] text-slate-100 rounded-2xl p-4 text-xs mono whitespace-pre-wrap overflow-y-auto max-h-72 leading-relaxed"></pre>

      <div class="flex items-center justify-between mt-5 pt-3 border-t border-slate-100">
        <button onclick="closeWaModal()" class="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900">
          Tutup
        </button>
        <div class="flex items-center gap-2">
          <button onclick="copyWaText()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold">
            📋 Salin Teks WA
          </button>
          <button onclick="sendToWhatsApp()" class="px-4 py-2.5 rounded-xl bg-emerald-800 hover:bg-emerald-900 text-white text-xs font-semibold">
            🟢 Buka WhatsApp
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ==================== MODAL INTEGRASI SPREADSHEET ==================== -->
  <div id="sheetModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs hidden items-center justify-center z-50 p-4 no-print">
    <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
      <h3 class="font-bold text-slate-900 text-base mb-1">🔗 Integrasi Google Spreadsheet</h3>
      <p class="text-xs text-slate-500 mb-4">
        Masukkan ID atau URL Google Spreadsheet Buku Kas Surau Asy-Syura.
      </p>
      <input
        id="inputSheetId"
        type="text"
        class="w-full px-3.5 py-2.5 text-xs border border-slate-300 rounded-xl mb-4 mono"
        placeholder="ID Google Spreadsheet"
      />
      <div class="flex justify-end gap-2">
        <button onclick="closeSheetModal()" class="px-4 py-2 text-xs font-semibold text-slate-600">Batal</button>
        <button onclick="saveSheetId()" class="px-4 py-2 rounded-xl bg-emerald-800 text-white text-xs font-semibold">Simpan & Sinkronkan</button>
      </div>
    </div>
  </div>

  <script>
    let appData = {{ initial_data | safe }};
    let selectedMonth = "ALL";
    let trendChartInstance = null;
    let donutChartInstance = null;

    function formatRp(num) {
      const n = Math.abs(Number(num) || 0);
      return "Rp " + n.toLocaleString("id-ID");
    }

    function initDashboard() {
      document.getElementById("inputSheetId").value = appData.sheet_id;
      document.getElementById("btnCatatKas").href = `https://docs.google.com/spreadsheets/d/${appData.sheet_id}/edit`;
      document.getElementById("headerPeriode").innerText = `Buku Kas & Progres Keuangan (${appData.periode})`;
      document.getElementById("printSubHeader").innerText = `Periode: ${appData.periode} | Dicetak: ${appData.updated_at}`;
      document.getElementById("labelPeriode1").innerText = appData.periode;
      document.getElementById("syncTimestamp").innerText = appData.updated_at;

      renderMonthFilters();
      renderMonthlyCardsAndBars();
      renderFilteredSection();
      renderTvView();
      populateWaSelect();
    }

    function renderMonthFilters() {
      const container = document.getElementById("monthFilterButtons");
      let html = `<button onclick="setMonthFilter('ALL')" class="px-3 py-1.5 rounded-lg transition ${
        selectedMonth === "ALL" ? "bg-white text-slate-900 shadow-xs font-bold" : "text-slate-600 hover:text-slate-900"
      }">Semua Bulan (Kumulatif)</button>`;

      appData.months.forEach((m) => {
        const active = selectedMonth === m;
        html += `<button onclick="setMonthFilter('${m}')" class="px-3 py-1.5 rounded-lg transition ${
          active ? "bg-white text-slate-900 shadow-xs font-bold" : "text-slate-600 hover:text-slate-900"
        }">${m}</button>`;
      });
      container.innerHTML = html;
    }

    function setMonthFilter(m) {
      selectedMonth = m;
      renderMonthFilters();
      renderFilteredSection();
    }

    function renderMonthlyCardsAndBars() {
      const cardsEl = document.getElementById("monthlyCardsContainer");
      const barsEl = document.getElementById("monthlyBarsContainer");
      const maxIncome = Math.max(...appData.monthly_summary.map((m) => m.pemasukan), 1);

      cardsEl.innerHTML = appData.monthly_summary
        .map((m) => {
          let vsHtml = "";
          if (m.vs_bulan_lalu !== null) {
            const sign = m.vs_bulan_lalu >= 0 ? "+" : "-";
            const color = m.vs_bulan_lalu >= 0 ? "text-emerald-700" : "text-rose-600";
            vsHtml = `
              <div class="flex justify-between text-xs pt-2 mt-2 border-t border-slate-200/70">
                <span class="text-slate-500">Vs Bulan Sebelumnya:</span>
                <span class="font-bold mono ${color}">${sign}${formatRp(m.vs_bulan_lalu)}</span>
              </div>`;
          }
          return `
            <div onclick="setMonthFilter('${m.bulan}')" class="cursor-pointer rounded-2xl border ${
            selectedMonth === m.bulan ? "border-emerald-600 ring-2 ring-emerald-500/20" : "border-slate-200"
          } bg-slate-50/70 hover:bg-white p-5 transition">
              <div class="flex items-center justify-between text-xs mb-2">
                <span class="font-bold text-slate-800">🗓️ ${m.bulan}</span>
                <span class="text-slate-400">${m.jumlah_tx} transaksi</span>
              </div>
              <p class="text-[11px] text-slate-500">Saldo Akhir Bulan Ini</p>
              <div class="text-2xl font-extrabold text-slate-900 mono mb-4">${formatRp(m.saldo_akhir)}</div>

              <div class="space-y-1.5 text-xs border-t border-slate-200/70 pt-3">
                <div class="flex justify-between">
                  <span class="text-slate-600">↙ Pemasukan:</span>
                  <span class="font-bold text-emerald-700 mono">+${formatRp(m.pemasukan)}</span>
                </div>
                <div class="flex justify-between text-slate-500 pl-3">
                  <span>└ Wirid Mingguan:</span>
                  <span class="mono">${formatRp(m.wirid)}</span>
                </div>
                <div class="flex justify-between text-slate-500 pl-3">
                  <span>└ Infaq/Transfer:</span>
                  <span class="mono">${formatRp(m.umum)}</span>
                </div>
                <div class="flex justify-between pt-1">
                  <span class="text-slate-600">↗ Pengeluaran:</span>
                  <span class="font-bold text-rose-600 mono">${m.pengeluaran > 0 ? "-" + formatRp(m.pengeluaran) : "Rp 0"}</span>
                </div>
                ${vsHtml}
              </div>
            </div>`;
        })
        .join("");

      barsEl.innerHTML = appData.monthly_summary
        .map((m) => {
          const widthPct = Math.max(Math.round((m.pemasukan / maxIncome) * 100), 5);
          return `
            <div class="grid grid-cols-12 items-center gap-3 text-xs">
              <div class="col-span-3 sm:col-span-2 text-slate-600 font-medium">${m.bulan}</div>
              <div class="col-span-6 sm:col-span-8 bg-slate-100 rounded-full h-3 overflow-hidden">
                <div class="bg-emerald-700 h-full rounded-full" style="width: ${widthPct}%"></div>
              </div>
              <div class="col-span-3 sm:col-span-2 text-right font-bold text-slate-800 mono">${formatRp(m.pemasukan)}</div>
            </div>`;
        })
        .join("");
    }

    function renderFilteredSection() {
      const txs =
        selectedMonth === "ALL"
          ? appData.transactions
          : appData.transactions.filter((t) => t.bulan === selectedMonth);

      const pemasukan = txs.reduce((a, b) => a + b.masuk, 0);
      const pengeluaran = txs.reduce((a, b) => a + b.keluar, 0);
      const wirid = txs.filter((t) => t.kategori === "Wirid Mingguan").reduce((a, b) => a + b.masuk, 0);
      const umum = txs.filter((t) => t.kategori === "Transfer Umum").reduce((a, b) => a + b.masuk, 0);
      const txMasuk = txs.filter((t) => t.masuk > 0).length;
      const txKeluar = txs.filter((t) => t.keluar > 0).length;
      const saldoAkhir = txs.length > 0 ? txs[txs.length - 1].saldo : 0;
      const wiridPct = pemasukan > 0 ? Math.round((wirid / pemasukan) * 100) : 0;
      const umumPct = pemasukan > 0 ? 100 - wiridPct : 0;

      document.getElementById("activeFilterLabel").innerText =
        selectedMonth === "ALL" ? `🗓️ Seluruh Periode (${appData.periode})` : `🗓️ Periode Bulan ${selectedMonth}`;
      document.getElementById("activeTxCount").innerText =
        `${txs.length} transaksi tercatat (${txMasuk} masuk, ${txKeluar} keluar)`;

      document.getElementById("kpiSaldo").innerText = formatRp(saldoAkhir);
      document.getElementById("kpiPemasukan").innerText = "+" + formatRp(pemasukan);
      document.getElementById("kpiPemasukanSub").innerText = `${txMasuk} kali penerimaan dana`;
      document.getElementById("kpiWirid").innerText = formatRp(wirid);
      document.getElementById("kpiWiridSub").innerText = `${wiridPct}% dari penerimaan · Umum: ${formatRp(umum)}`;
      document.getElementById("kpiPengeluaran").innerText = "-" + formatRp(pengeluaran);
      document.getElementById("kpiPengeluaranSub").innerText = `${txKeluar} kali pengeluaran tercatat`;

      document.getElementById("badgeMasukCount").innerText = `${txMasuk} Pemasukan`;
      document.getElementById("donutTotal").innerText = formatRp(pemasukan);
      document.getElementById("donutTxCount").innerText = `${txs.length} transaksi tercatat`;
      document.getElementById("donutWiridPct").innerText = `${wiridPct}%`;
      document.getElementById("donutUmumPct").innerText = `${umumPct}%`;

      renderCharts(txs, wirid, umum);
      renderTable();
    }

    function renderCharts(txs, wirid, umum) {
      const incomeTxs = txs.filter((t) => t.masuk > 0);
      const labels = incomeTxs.map((t) => t.tanggal);
      const values = incomeTxs.map((t) => t.masuk);
      const pointColors = incomeTxs.map((t) => (t.kategori === "Wirid Mingguan" ? "#f59e0b" : "#059669"));

      if (trendChartInstance) trendChartInstance.destroy();
      const ctxTrend = document.getElementById("trendChart").getContext("2d");
      trendChartInstance = new Chart(ctxTrend, {
        type: "line",
        data: {
          labels: labels,
          datasets: [
            {
              label: "Nominal Masuk",
              data: values,
              borderColor: "#059669",
              backgroundColor: "rgba(5, 150, 105, 0.12)",
              pointBackgroundColor: pointColors,
              pointRadius: 5,
              tension: 0.35,
              fill: true,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: {
              ticks: {
                callback: (v) => "Rp " + Number(v).toLocaleString("id-ID"),
              },
            },
          },
        },
      });

      if (donutChartInstance) donutChartInstance.destroy();
      const ctxDonut = document.getElementById("donutChart").getContext("2d");
      donutChartInstance = new Chart(ctxDonut, {
        type: "doughnut",
        data: {
          labels: ["Wirid Mingguan", "Transfer Umum"],
          datasets: [
            {
              data: [wirid, umum],
              backgroundColor: ["#f59e0b", "#059669"],
              borderWidth: 0,
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: "72%",
          plugins: { legend: { display: false } },
        },
      });
    }

    function renderTable() {
      const q = (document.getElementById("searchInput").value || "").toLowerCase();
      const baseTxs =
        selectedMonth === "ALL"
          ? appData.transactions
          : appData.transactions.filter((t) => t.bulan === selectedMonth);

      const filtered = baseTxs.filter(
        (t) => t.uraian.toLowerCase().includes(q) || t.tanggal.toLowerCase().includes(q)
      );

      const tbody = document.getElementById("kasTableBody");
      tbody.innerHTML = filtered
        .map(
          (t, idx) => `
        <tr class="hover:bg-slate-50">
          <td class="py-2.5 px-3 text-slate-400">${idx + 1}</td>
          <td class="py-2.5 px-3 font-medium text-slate-700">${t.tanggal}</td>
          <td class="py-2.5 px-3 text-slate-500">${t.bulan}</td>
          <td class="py-2.5 px-3 font-medium text-slate-800">${t.uraian}</td>
          <td class="py-2.5 px-3 text-right font-semibold text-emerald-700 mono">${
            t.masuk > 0 ? "+" + formatRp(t.masuk) : "-"
          }</td>
          <td class="py-2.5 px-3 text-right font-semibold text-rose-600 mono">${
            t.keluar > 0 ? "-" + formatRp(t.keluar) : "-"
          }</td>
          <td class="py-2.5 px-3 text-right font-bold text-slate-900 mono">${formatRp(t.saldo)}</td>
        </tr>`
        )
        .join("");
    }

    function renderTvView() {
      document.getElementById("tvSubPeriode").innerText = `Transparansi Kas Jamaah Bappeda (${appData.periode})`;
      document.getElementById("tvHeroPeriode").innerText = `PERIODE: ${appData.periode.toUpperCase()}`;
      document.getElementById("tvHeroSaldo").innerText = formatRp(appData.totals.saldo_akhir);

      document.getElementById("tvMonthlyCards").innerHTML = appData.monthly_summary
        .map(
          (m) => `
        <div class="rounded-2xl bg-[#0b1528] border border-slate-800 p-4">
          <div class="flex items-center justify-between text-xs text-emerald-400 mb-1">
            <span>${m.bulan}</span>
            <span class="text-slate-500">${m.jumlah_tx} tx</span>
          </div>
          <div class="text-xl font-extrabold text-white mono">+${formatRp(m.pemasukan)}</div>
          <div class="text-[11px] text-slate-400 mt-1">Saldo: ${formatRp(m.saldo_akhir)}</div>
        </div>`
        )
        .join("");

      const latestMonth = appData.months[appData.months.length - 1] || "September 2026";
      document.getElementById("tvLatestMonthLabel").innerText = latestMonth.toUpperCase();
      const latestTxs = appData.transactions.slice(-6).reverse();

      document.getElementById("tvMutationList").innerHTML = latestTxs
        .map((t) => {
          const isMasuk = t.masuk > 0;
          return `
          <div class="rounded-2xl bg-[#101c33] border border-slate-800/90 px-4 py-3 flex items-center justify-between gap-3">
            <div>
              <div class="text-xs font-bold text-white">${t.tanggal}</div>
              <div class="text-[11px] text-slate-400 truncate max-w-[210px]">${t.uraian}</div>
            </div>
            <div class="text-xs font-extrabold mono ${isMasuk ? "text-emerald-400" : "text-rose-400"}">
              ${isMasuk ? "+" + formatRp(t.masuk) : "-" + formatRp(t.keluar)}
            </div>
          </div>`;
        })
        .join("");
    }

    function populateWaSelect() {
      const sel = document.getElementById("waMonthSelect");
      sel.innerHTML = `<option value="ALL">Kumulatif Seluruh Bulan</option>`;
      appData.months.forEach((m) => {
        sel.innerHTML += `<option value="${m}">${m}</option>`;
      });
      updateWaPreview();
    }

    function updateWaPreview() {
      const val = document.getElementById("waMonthSelect").value;
      const txs = val === "ALL" ? appData.transactions : appData.transactions.filter((t) => t.bulan === val);
      const periodeLabel = val === "ALL" ? appData.periode : val;

      const pemasukan = txs.reduce((a, b) => a + b.masuk, 0);
      const wirid = txs.filter((t) => t.kategori === "Wirid Mingguan").reduce((a, b) => a + b.masuk, 0);
      const umum = txs.filter((t) => t.kategori === "Transfer Umum").reduce((a, b) => a + b.masuk, 0);
      const pengeluaran = txs.reduce((a, b) => a + b.keluar, 0);
      const saldoAkhir = txs.length > 0 ? txs[txs.length - 1].saldo : 0;

      const text = `*LAPORAN KEUANGAN KAS SURAU ASY SYURA BAPPEDA*
🗓️ Periode: *${periodeLabel}*
🕒 Update per: ${new Date().toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}

*RINGKASAN SALDO:*
• Total Pemasukan: *${formatRp(pemasukan)}*
  └ Wirid Mingguan: ${formatRp(wirid)}
  └ Transfer/Infaq Umum: ${formatRp(umum)}
• Total Pengeluaran: *${formatRp(pengeluaran)}*
──────────────────────
💰 *SALDO KAS SURAU: ${formatRp(saldoAkhir)}*
──────────────────────
_Alhamdulillah, terima kasih atas infaq dan sedekah Bapak/Ibu jamaah Bappeda._`;

      document.getElementById("waTextPreview").innerText = text;
    }

    function copyWaText() {
      const text = document.getElementById("waTextPreview").innerText;
      navigator.clipboard.writeText(text);
      alert("Teks laporan WhatsApp berhasil disalin!");
    }

    function sendToWhatsApp() {
      const text = encodeURIComponent(document.getElementById("waTextPreview").innerText);
      window.open(`https://wa.me/?text=${text}`, "_blank");
    }

    function switchTab(mode) {
      if (mode === "tv") {
        document.getElementById("dashboardView").classList.add("hidden");
        document.getElementById("tvView").classList.remove("hidden");
      } else {
        document.getElementById("tvView").classList.add("hidden");
        document.getElementById("dashboardView").classList.remove("hidden");
      }
    }

    function openWaModal() {
      updateWaPreview();
      document.getElementById("waModal").classList.remove("hidden");
      document.getElementById("waModal").classList.add("flex");
    }
    function closeWaModal() {
      document.getElementById("waModal").classList.add("hidden");
      document.getElementById("waModal").classList.remove("flex");
    }

    function openSheetModal() {
      document.getElementById("sheetModal").classList.remove("hidden");
      document.getElementById("sheetModal").classList.add("flex");
    }
    function closeSheetModal() {
      document.getElementById("sheetModal").classList.add("hidden");
      document.getElementById("sheetModal").classList.remove("flex");
    }

    async function refreshData(customSheetId) {
      const sid = customSheetId || appData.sheet_id;
      const res = await fetch(`/api/kas?sheet_id=${encodeURIComponent(sid)}`);
      if (res.ok) {
        appData = await res.json();
        initDashboard();
      }
    }

    function saveSheetId() {
      let raw = document.getElementById("inputSheetId").value.trim();
      const match = raw.match(/\\/d\\/([a-zA-Z0-9-_]+)/);
      if (match) raw = match[1];
      closeSheetModal();
      refreshData(raw);
    }

    function toggleFullScreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
      } else {
        document.exitFullscreen();
      }
    }

    setInterval(() => {
      const now = new Date();
      document.getElementById("tvClock").innerText = now.toLocaleTimeString("id-ID").replace(/:/g, ".");
      document.getElementById("tvDate").innerText = now.toLocaleDateString("id-ID", {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric",
      });
    }, 1000);

    // Auto-refresh data dari Google Sheet setiap 5 menit untuk Layar TV Jamaah
    setInterval(() => refreshData(), 300000);

    initDashboard();
  </script>
</body>
</html>
"""


@app.route("/")
def index():
    sheet_id = request.args.get("sheet_id", DEFAULT_SHEET_ID)
    data = fetch_spreadsheet_data(sheet_id)
    return render_template_string(HTML_TEMPLATE, initial_data=json.dumps(data))


@app.route("/api/kas")
def api_kas():
    sheet_id = request.args.get("sheet_id", DEFAULT_SHEET_ID)
    data = fetch_spreadsheet_data(sheet_id)
    return jsonify(data)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
