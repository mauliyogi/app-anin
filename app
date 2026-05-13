<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SPS - Transfer Antar Gudang</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

  :root {
    --bg: #0d1117;
    --surface: #161b22;
    --surface2: #21262d;
    --surface3: #30363d;
    --border: #30363d;
    --border2: #21262d;
    --text: #e6edf3;
    --text2: #8b949e;
    --text3: #6e7681;
    --accent: #1f6feb;
    --accent2: #388bfd;
    --green: #3fb950;
    --green2: #238636;
    --yellow: #d29922;
    --red: #f85149;
    --orange: #db6d28;
    --purple: #8b5cf6;
    --radius: 8px;
    --radius2: 12px;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  
  body {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    font-size: 14px;
  }

  /* SIDEBAR */
  .sidebar {
    position: fixed; top: 0; left: 0;
    width: 220px; height: 100vh;
    background: var(--surface);
    border-right: 1px solid var(--border);
    display: flex; flex-direction: column;
    z-index: 100;
  }
  .sidebar-logo {
    padding: 20px 16px 16px;
    border-bottom: 1px solid var(--border);
  }
  .sidebar-logo .brand {
    font-size: 20px; font-weight: 800;
    color: var(--accent2);
    letter-spacing: -0.5px;
  }
  .sidebar-logo .sub {
    font-size: 10px; color: var(--text3);
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: 1px; text-transform: uppercase;
    margin-top: 2px;
  }
  .nav-section { padding: 12px 8px 4px; }
  .nav-label {
    font-size: 10px; font-weight: 600;
    color: var(--text3); letter-spacing: 1px;
    text-transform: uppercase; padding: 0 8px 6px;
  }
  .nav-item {
    display: flex; align-items: center; gap: 10px;
    padding: 8px 12px; border-radius: var(--radius);
    cursor: pointer; color: var(--text2);
    font-size: 13px; font-weight: 500;
    transition: all 0.15s; margin-bottom: 2px;
  }
  .nav-item:hover { background: var(--surface2); color: var(--text); }
  .nav-item.active { background: rgba(31,111,235,0.15); color: var(--accent2); }
  .nav-item .icon { font-size: 16px; width: 20px; text-align: center; }
  .sidebar-footer {
    margin-top: auto; padding: 12px 16px;
    border-top: 1px solid var(--border);
    font-size: 11px; color: var(--text3);
  }

  /* MAIN */
  .main {
    margin-left: 220px;
    min-height: 100vh;
    padding: 24px;
  }

  /* PAGE SECTIONS */
  .page { display: none; }
  .page.active { display: block; }

  /* PAGE HEADER */
  .page-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 24px;
  }
  .page-title { font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }
  .page-sub { font-size: 13px; color: var(--text2); margin-top: 2px; }

  /* STATS CARDS */
  .stats-grid {
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px;
    margin-bottom: 24px;
  }
  .stat-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius2); padding: 20px;
  }
  .stat-label { font-size: 12px; color: var(--text2); font-weight: 500; }
  .stat-value { font-size: 28px; font-weight: 800; margin-top: 6px; }
  .stat-tag {
    display: inline-block; font-size: 10px; font-weight: 600;
    padding: 2px 8px; border-radius: 20px; margin-top: 6px;
  }
  .tag-green { background: rgba(63,185,80,0.15); color: var(--green); }
  .tag-blue { background: rgba(56,139,253,0.15); color: var(--accent2); }
  .tag-yellow { background: rgba(210,153,34,0.15); color: var(--yellow); }
  .tag-purple { background: rgba(139,92,246,0.15); color: var(--purple); }

  /* BUTTONS */
  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 8px 16px; border-radius: var(--radius);
    font-size: 13px; font-weight: 600;
    cursor: pointer; border: none; transition: all 0.15s;
    font-family: 'Plus Jakarta Sans', sans-serif;
  }
  .btn-primary { background: var(--accent); color: white; }
  .btn-primary:hover { background: var(--accent2); }
  .btn-success { background: var(--green2); color: white; }
  .btn-success:hover { background: var(--green); }
  .btn-outline {
    background: transparent; color: var(--text2);
    border: 1px solid var(--border);
  }
  .btn-outline:hover { background: var(--surface2); color: var(--text); }
  .btn-danger { background: rgba(248,81,73,0.15); color: var(--red); border: 1px solid rgba(248,81,73,0.3); }
  .btn-danger:hover { background: rgba(248,81,73,0.25); }
  .btn-sm { padding: 5px 10px; font-size: 12px; }
  .btn-lg { padding: 12px 24px; font-size: 15px; }

  /* FORMS */
  .card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius2); padding: 20px; margin-bottom: 16px;
  }
  .card-title {
    font-size: 15px; font-weight: 700; margin-bottom: 16px;
    display: flex; align-items: center; gap: 8px;
    padding-bottom: 12px; border-bottom: 1px solid var(--border);
  }
  .form-row { display: grid; gap: 12px; margin-bottom: 12px; }
  .form-row.col2 { grid-template-columns: 1fr 1fr; }
  .form-row.col3 { grid-template-columns: 1fr 1fr 1fr; }
  .form-row.col4 { grid-template-columns: 1fr 1fr 1fr 1fr; }
  .form-group { display: flex; flex-direction: column; gap: 6px; }
  .form-label {
    font-size: 12px; font-weight: 600; color: var(--text2);
    text-transform: uppercase; letter-spacing: 0.5px;
  }
  .form-input, .form-select {
    background: var(--surface2); border: 1px solid var(--border);
    color: var(--text); border-radius: var(--radius);
    padding: 9px 12px; font-size: 13px;
    font-family: 'Plus Jakarta Sans', sans-serif;
    transition: border 0.15s;
    width: 100%;
  }
  .form-input:focus, .form-select:focus {
    outline: none; border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(31,111,235,0.1);
  }
  .form-input[readonly] { color: var(--text2); cursor: not-allowed; }
  .form-select option { background: var(--surface2); }

  /* TABLE */
  .table-wrap {
    border: 1px solid var(--border); border-radius: var(--radius2);
    overflow: hidden; margin-bottom: 16px;
  }
  table { width: 100%; border-collapse: collapse; }
  thead { background: var(--surface2); }
  th {
    padding: 10px 12px; text-align: left;
    font-size: 11px; font-weight: 600; color: var(--text2);
    text-transform: uppercase; letter-spacing: 0.5px;
    border-bottom: 1px solid var(--border);
  }
  td {
    padding: 10px 12px; border-bottom: 1px solid var(--border2);
    font-size: 13px; color: var(--text);
    vertical-align: middle;
  }
  tr:last-child td { border-bottom: none; }
  tbody tr:hover { background: rgba(255,255,255,0.02); }

  /* PRODUCT TABLE IN FORM */
  .product-table input {
    background: var(--surface3); border: 1px solid transparent;
    color: var(--text); border-radius: 6px;
    padding: 6px 8px; font-size: 13px; width: 90px;
    font-family: 'Plus Jakarta Sans', sans-serif;
    text-align: center;
  }
  .product-table input:focus {
    outline: none; border-color: var(--accent);
    background: var(--surface2);
  }
  .product-table td:first-child { color: var(--text2); font-size: 12px; }

  /* BADGE */
  .badge {
    display: inline-flex; align-items: center;
    padding: 3px 8px; border-radius: 20px;
    font-size: 11px; font-weight: 600;
  }
  .badge-pagi { background: rgba(210,153,34,0.15); color: var(--yellow); }
  .badge-sore { background: rgba(63,185,80,0.15); color: var(--green); }
  .badge-selesai { background: rgba(56,139,253,0.15); color: var(--accent2); }

  /* PHASE INDICATOR */
  .phase-bar {
    display: flex; gap: 0; margin-bottom: 20px;
    border: 1px solid var(--border); border-radius: var(--radius);
    overflow: hidden;
  }
  .phase-step {
    flex: 1; padding: 10px 16px; text-align: center;
    font-size: 12px; font-weight: 600; color: var(--text3);
    background: var(--surface2); cursor: pointer;
    transition: all 0.15s;
    border-right: 1px solid var(--border);
  }
  .phase-step:last-child { border-right: none; }
  .phase-step.done { background: rgba(63,185,80,0.1); color: var(--green); }
  .phase-step.active { background: rgba(31,111,235,0.15); color: var(--accent2); }

  /* MODAL */
  .modal-overlay {
    display: none; position: fixed; inset: 0;
    background: rgba(0,0,0,0.7); z-index: 1000;
    align-items: center; justify-content: center;
  }
  .modal-overlay.open { display: flex; }
  .modal {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius2); padding: 24px;
    width: 480px; max-width: 90vw; max-height: 85vh; overflow-y: auto;
  }
  .modal-title { font-size: 16px; font-weight: 700; margin-bottom: 16px; }
  .modal-footer { display: flex; gap: 8px; justify-content: flex-end; margin-top: 16px; }

  /* ALERT */
  .alert {
    padding: 10px 14px; border-radius: var(--radius);
    font-size: 13px; margin-bottom: 16px; display: flex; gap: 8px;
  }
  .alert-success { background: rgba(63,185,80,0.1); border: 1px solid rgba(63,185,80,0.3); color: var(--green); }
  .alert-info { background: rgba(31,111,235,0.1); border: 1px solid rgba(31,111,235,0.3); color: var(--accent2); }

  /* CHIP LIST */
  .chip-list { display: flex; flex-wrap: wrap; gap: 8px; }
  .chip {
    display: inline-flex; align-items: center; gap: 6px;
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: 20px; padding: 5px 12px;
    font-size: 12px; font-weight: 500; color: var(--text);
  }
  .chip-del {
    cursor: pointer; color: var(--text3);
    font-size: 14px; line-height: 1;
    transition: color 0.1s;
  }
  .chip-del:hover { color: var(--red); }

  /* RECORD TABLE */
  .record-filters {
    display: flex; gap: 8px; margin-bottom: 16px; align-items: center;
  }
  .record-filters .form-input { max-width: 180px; }

  /* NO FORM NUMBER display */
  .form-number-display {
    font-family: 'JetBrains Mono', monospace;
    font-size: 18px; font-weight: 700;
    color: var(--accent2); letter-spacing: 1px;
  }

  /* SPS SECTION */
  .sps-card {
    background: linear-gradient(135deg, rgba(31,111,235,0.1), rgba(139,92,246,0.1));
    border: 1px solid rgba(31,111,235,0.3);
    border-radius: var(--radius2); padding: 20px; margin-bottom: 16px;
  }
  .sps-title { font-size: 14px; font-weight: 700; color: var(--accent2); margin-bottom: 8px; }
  .sps-desc { font-size: 12px; color: var(--text2); line-height: 1.6; }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--surface3); border-radius: 3px; }

  /* Print styles - the PDF preview area */
  #pdf-preview-area { display: none; }

  @media print {
    body * { visibility: hidden; }
    #pdf-preview-area, #pdf-preview-area * { visibility: visible; }
    #pdf-preview-area {
      display: block !important;
      position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    }
  }

  /* PDF Preview */
  #pdf-preview-area {
    font-family: Arial, sans-serif;
    font-size: 11px;
    color: #000;
    background: white;
  }
  .pdf-doc {
    width: 210mm; min-height: 270mm;
    padding: 12mm 14mm;
    background: white; color: black;
  }
  .pdf-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; }
  .pdf-logo { font-size: 22px; font-weight: 900; color: #1a5fba; letter-spacing: -1px; }
  .pdf-logo-sub { font-size: 7px; color: #555; letter-spacing: 1px; }
  .pdf-title { font-size: 14px; font-weight: 900; text-align: center; margin-bottom: 8px; letter-spacing: 1px; }
  .pdf-no { border: 2px solid black; padding: 3px 10px; font-size: 13px; font-weight: 700; }
  .pdf-info-table { width: 100%; margin-bottom: 8px; }
  .pdf-info-table td { padding: 2px 0; font-size: 11px; vertical-align: top; }
  .pdf-main-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 10px; }
  .pdf-main-table th, .pdf-main-table td {
    border: 1px solid black; padding: 4px 6px; text-align: center;
  }
  .pdf-main-table th { background: #f0f0f0; font-weight: 700; }
  .pdf-main-table .product-name { text-align: left; }
  .pdf-sign-section { display: flex; gap: 10px; margin-top: 10px; }
  .pdf-sign-box { flex: 1; border: 1px solid black; padding: 8px; }
  .pdf-sign-title { font-weight: 700; font-size: 10px; text-align: center; border-bottom: 1px solid black; padding-bottom: 4px; margin-bottom: 4px; }
  .pdf-sign-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
  .pdf-sign-item { text-align: center; font-size: 9px; }
  .pdf-sign-line { border-bottom: 1px solid black; height: 35px; margin-bottom: 2px; }
  .pdf-sign-role { font-weight: 700; }
  .pdf-sign-name { font-style: italic; color: #333; }
</style>
</head>
<body>

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sidebar-logo">
    <div class="brand">⚡ SPS</div>
    <div class="sub">Distribution Partner</div>
  </div>
  
  <div class="nav-section">
    <div class="nav-label">Menu Utama</div>
    <div class="nav-item active" onclick="showPage('dashboard')">
      <span class="icon">🏠</span> Dashboard
    </div>
    <div class="nav-item" onclick="showPage('form')">
      <span class="icon">📋</span> Buat Form Baru
    </div>
    <div class="nav-item" onclick="showPage('records')">
      <span class="icon">📊</span> Rekap Data
    </div>
  </div>
  
  <div class="nav-section">
    <div class="nav-label">Pengaturan</div>
    <div class="nav-item" onclick="showPage('master')">
      <span class="icon">⚙️</span> Master Data
    </div>
  </div>
  
  <div class="sidebar-footer">
    PT. Sentralsari Primasentosa<br>
    Cabang: <strong id="sidebar-branch">Karang Tengah</strong>
  </div>
</div>

<!-- MAIN -->
<div class="main">

<!-- ============ DASHBOARD ============ -->
<div id="page-dashboard" class="page active">
  <div class="page-header">
    <div>
      <div class="page-title">Dashboard</div>
      <div class="page-sub" id="today-date"></div>
    </div>
    <button class="btn btn-primary" onclick="showPage('form')">+ Form Baru</button>
  </div>
  
  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-label">Form Hari Ini</div>
      <div class="stat-value" id="stat-today" style="color:var(--accent2)">0</div>
      <div class="stat-tag tag-blue">Aktif</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Total Form Bulan Ini</div>
      <div class="stat-value" id="stat-month" style="color:var(--purple)">0</div>
      <div class="stat-tag tag-purple">Bulan Ini</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Form Belum Selesai</div>
      <div class="stat-value" id="stat-pending" style="color:var(--yellow)">0</div>
      <div class="stat-tag tag-yellow">Pending</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Total Salesman</div>
      <div class="stat-value" id="stat-sales" style="color:var(--green)">0</div>
      <div class="stat-tag tag-green">Aktif</div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">📋 Form Terbaru</div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>No. Form</th>
            <th>Tanggal</th>
            <th>Salesman</th>
            <th>No. Polisi</th>
            <th>Status</th>
            <th>Aksi</th>
          </tr>
        </thead>
        <tbody id="dashboard-table">
          <tr><td colspan="6" style="text-align:center;color:var(--text3);padding:20px">Belum ada data</td></tr>
        </tbody>
      </table>
    </div>
  </div>
</div>

<!-- ============ FORM BARU ============ -->
<div id="page-form" class="page">
  <div class="page-header">
    <div>
      <div class="page-title">Form Transfer Antar Gudang</div>
      <div class="page-sub">PT. Sentralsari Primasentosa</div>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-outline" onclick="resetForm()">↺ Reset</button>
    </div>
  </div>

  <!-- Phase indicator -->
  <div class="phase-bar">
    <div class="phase-step active" id="phase-1">① Pagi — Isi Naik & Tanda Tangan</div>
    <div class="phase-step" id="phase-2">② Sore — Isi Turun & Jual</div>
    <div class="phase-step" id="phase-3">③ Export & Input SPS</div>
  </div>

  <!-- Form Info -->
  <div class="card">
    <div class="card-title">📄 Informasi Form</div>
    <div class="form-row col3">
      <div class="form-group">
        <div class="form-label">No. Form</div>
        <input type="text" class="form-input" id="f-nomer" placeholder="A - 3070355">
      </div>
      <div class="form-group">
        <div class="form-label">Cabang</div>
        <input type="text" class="form-input" id="f-cabang" value="KARANG TENGAH" readonly>
      </div>
      <div class="form-group">
        <div class="form-label">Tanggal</div>
        <input type="date" class="form-input" id="f-tanggal">
      </div>
    </div>
    <div class="form-row col3">
      <div class="form-group">
        <div class="form-label">Nomer TIS</div>
        <input type="text" class="form-input" id="f-tis" placeholder="Nomer TIS">
      </div>
      <div class="form-group">
        <div class="form-label">Sopir / Salesman</div>
        <select class="form-select" id="f-salesman">
          <option value="">-- Pilih Salesman --</option>
        </select>
      </div>
      <div class="form-group">
        <div class="form-label">No. Polisi</div>
        <select class="form-select" id="f-polisi">
          <option value="">-- Pilih Mobil --</option>
        </select>
      </div>
    </div>
  </div>

  <!-- Products Table -->
  <div class="card">
    <div class="card-title">📦 Data Produk</div>
    <div class="table-wrap">
      <table class="product-table">
        <thead>
          <tr>
            <th style="width:40px">No</th>
            <th>Nama Produk</th>
            <th style="width:110px;text-align:center">NAIK</th>
            <th style="width:110px;text-align:center">TURUN</th>
            <th style="width:110px;text-align:center">JUAL</th>
          </tr>
        </thead>
        <tbody id="product-table-body"></tbody>
      </table>
    </div>
  </div>

  <!-- Signatures -->
  <div class="form-row col2">
    <div class="card">
      <div class="card-title">✍️ Otorisasi TAG Naik (Pagi)</div>
      <div class="form-row col2">
        <div class="form-group">
          <div class="form-label">Checker</div>
          <input type="text" class="form-input" id="s-naik-checker" placeholder="Nama checker">
        </div>
        <div class="form-group">
          <div class="form-label">Driver / Salesman</div>
          <input type="text" class="form-input" id="s-naik-driver" placeholder="Otomatis dari salesman" readonly>
        </div>
        <div class="form-group">
          <div class="form-label">Admin Penjualan</div>
          <input type="text" class="form-input" id="s-naik-admin" placeholder="Nama admin">
        </div>
        <div class="form-group">
          <div class="form-label">Ka. Cabang</div>
          <input type="text" class="form-input" id="s-naik-ka" placeholder="Nama Ka.Cabang">
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-title">✍️ Otorisasi TAG Turun (Sore)</div>
      <div class="form-row col2">
        <div class="form-group">
          <div class="form-label">Driver / Salesman</div>
          <input type="text" class="form-input" id="s-turun-driver" placeholder="Otomatis dari salesman" readonly>
        </div>
        <div class="form-group">
          <div class="form-label">Checker</div>
          <input type="text" class="form-input" id="s-turun-checker" placeholder="Nama checker">
        </div>
        <div class="form-group">
          <div class="form-label">Admin Penjualan</div>
          <input type="text" class="form-input" id="s-turun-admin" placeholder="Nama admin">
        </div>
        <div class="form-group">
          <div class="form-label">Ka. Cabang</div>
          <input type="text" class="form-input" id="s-turun-ka" placeholder="Nama Ka.Cabang">
        </div>
      </div>
    </div>
  </div>

  <!-- Catatan -->
  <div class="card">
    <div class="card-title">📝 Catatan</div>
    <div class="form-row">
      <div class="form-group">
        <textarea class="form-input" id="f-catatan" rows="3" placeholder="Catatan tambahan (opsional)..." style="resize:vertical"></textarea>
      </div>
    </div>
  </div>

  <!-- Action Buttons -->
  <div style="display:flex;gap:12px;flex-wrap:wrap">
    <button class="btn btn-success btn-lg" onclick="saveForm()">💾 Simpan Form</button>
    <button class="btn btn-primary btn-lg" onclick="exportPDF()">🖨️ Export PDF</button>
    <button class="btn btn-outline btn-lg" onclick="exportExcel()">📊 Export Excel</button>
    <button class="btn btn-outline btn-lg" onclick="inputSPS()">⚡ Input ke SPS</button>
  </div>
</div>

<!-- ============ RECORDS ============ -->
<div id="page-records" class="page">
  <div class="page-header">
    <div>
      <div class="page-title">Rekap Data</div>
      <div class="page-sub">Semua riwayat form transfer</div>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-outline" onclick="exportAllExcel()">📊 Export Semua ke Excel</button>
    </div>
  </div>

  <div class="record-filters">
    <input type="date" class="form-input" id="filter-dari" placeholder="Dari tanggal">
    <input type="date" class="form-input" id="filter-sampai" placeholder="Sampai tanggal">
    <select class="form-select" id="filter-salesman" style="max-width:180px">
      <option value="">Semua Salesman</option>
    </select>
    <button class="btn btn-outline" onclick="filterRecords()">🔍 Filter</button>
    <button class="btn btn-outline" onclick="clearFilter()">✕ Reset</button>
  </div>
  
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>No. Form</th>
          <th>Tanggal</th>
          <th>Salesman</th>
          <th>No. Polisi</th>
          <th>Total Naik</th>
          <th>Total Turun</th>
          <th>Total Jual</th>
          <th>Status</th>
          <th>Aksi</th>
        </tr>
      </thead>
      <tbody id="records-table"></tbody>
    </table>
  </div>
</div>

<!-- ============ MASTER DATA ============ -->
<div id="page-master" class="page">
  <div class="page-header">
    <div>
      <div class="page-title">Master Data</div>
      <div class="page-sub">Kelola salesman, mobil, produk, dan info cabang</div>
    </div>
  </div>

  <div class="form-row col2">
    <!-- Salesman -->
    <div class="card">
      <div class="card-title">👤 Salesman</div>
      <div class="chip-list" id="salesman-list" style="margin-bottom:12px"></div>
      <div style="display:flex;gap:8px">
        <input type="text" class="form-input" id="new-salesman" placeholder="Nama salesman baru">
        <button class="btn btn-primary" onclick="addSalesman()">+ Tambah</button>
      </div>
    </div>

    <!-- Mobil -->
    <div class="card">
      <div class="card-title">🚐 No. Polisi Mobil</div>
      <div class="chip-list" id="car-list" style="margin-bottom:12px"></div>
      <div style="display:flex;gap:8px">
        <input type="text" class="form-input" id="new-car" placeholder="Nomor polisi baru">
        <button class="btn btn-primary" onclick="addCar()">+ Tambah</button>
      </div>
    </div>
  </div>

  <!-- Produk -->
  <div class="card">
    <div class="card-title">📦 Produk</div>
    <div class="chip-list" id="product-list" style="margin-bottom:12px"></div>
    <div style="display:flex;gap:8px">
      <input type="text" class="form-input" id="new-product" placeholder="Nama produk baru" style="max-width:300px">
      <button class="btn btn-primary" onclick="addProduct()">+ Tambah</button>
    </div>
  </div>

  <!-- Cabang -->
  <div class="card">
    <div class="card-title">🏢 Info Cabang</div>
    <div class="form-row col2">
      <div class="form-group">
        <div class="form-label">Nama Cabang</div>
        <input type="text" class="form-input" id="m-cabang" placeholder="Nama cabang">
      </div>
      <div class="form-group">
        <div class="form-label">Ka. Cabang Default</div>
        <input type="text" class="form-input" id="m-ka" placeholder="Nama kepala cabang">
      </div>
      <div class="form-group">
        <div class="form-label">Admin Penjualan Default</div>
        <input type="text" class="form-input" id="m-admin" placeholder="Nama admin default">
      </div>
    </div>
    <button class="btn btn-success" onclick="saveMaster()">💾 Simpan Pengaturan</button>
  </div>

  <!-- SPS Info -->
  <div class="sps-card">
    <div class="sps-title">⚡ Informasi Input SPS</div>
    <div class="sps-desc">
      Fitur "Input ke SPS" akan mengunduh file Excel yang sudah diformat sesuai template input sistem SPS. 
      File tersebut berisi data dari form yang siap di-copy-paste atau di-import langsung ke sistem SPS tanpa perlu menginput ulang secara manual.
      <br><br>
      Format output: <strong>SPS_Import_[NomorForm]_[Tanggal].xlsx</strong>
    </div>
  </div>
</div>

</div><!-- end main -->

<!-- PDF HIDDEN AREA -->
<div id="pdf-preview-area">
  <div class="pdf-doc" id="pdf-content"></div>
</div>

<!-- MODAL: View Form -->
<div class="modal-overlay" id="modal-view">
  <div class="modal" style="width:600px">
    <div class="modal-title">📋 Detail Form</div>
    <div id="modal-view-content"></div>
    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeModal('modal-view')">Tutup</button>
      <button class="btn btn-primary" onclick="printFromModal()">🖨️ Print PDF</button>
      <button class="btn btn-outline" onclick="exportFromModal()">📊 Excel</button>
    </div>
  </div>
</div>

<!-- MODAL: SPS Input -->
<div class="modal-overlay" id="modal-sps">
  <div class="modal">
    <div class="modal-title">⚡ Input ke SPS</div>
    <div class="alert alert-info">
      File Excel untuk import SPS telah disiapkan. Buka sistem SPS dan import file berikut:
    </div>
    <div id="sps-modal-content"></div>
    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeModal('modal-sps')">Tutup</button>
      <button class="btn btn-success" id="sps-download-btn">⬇️ Unduh File SPS</button>
    </div>
  </div>
</div>

<script>
// ===================== STATE =====================
let state = {
  records: [],
  salesmen: ['Rohman', 'Ahmad', 'Budi'],
  cars: ['T 8941 HH', 'B 1234 XY', 'B 5678 AB'],
  products: [
    'AMDK 19 ltr', 'Galon kosong Cleo', 'AMDK 19 ltr IDM', 'Galon kosong IDM',
    'Cleo Mini 220', 'Cleo Cup 120 extra', 'Cleo Cup 220',
    'Cleo 330 3D', 'Cleo 550 3D', 'Cleo 550 3D SR', 'Cleo 1500 3D'
  ],
  cabang: 'KARANG TENGAH',
  kaCabang: '',
  adminDefault: '',
  currentFormId: null
};

// Load from localStorage
function loadState() {
  const saved = localStorage.getItem('sps_state');
  if (saved) {
    const s = JSON.parse(saved);
    state = { ...state, ...s };
  }
}

function saveState() {
  localStorage.setItem('sps_state', JSON.stringify(state));
}

// ===================== INIT =====================
loadState();

document.addEventListener('DOMContentLoaded', () => {
  // Set today's date
  const today = new Date();
  document.getElementById('today-date').textContent = today.toLocaleDateString('id-ID', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
  document.getElementById('f-tanggal').value = today.toISOString().split('T')[0];
  
  initDropdowns();
  renderProductTable();
  renderMasterData();
  updateDashboard();
  renderRecordsTable();
  
  // Auto-fill salesman into signature field
  document.getElementById('f-salesman').addEventListener('change', function() {
    document.getElementById('s-naik-driver').value = this.value;
    document.getElementById('s-turun-driver').value = this.value;
  });
  
  // Auto-fill default admin/ka
  document.getElementById('s-naik-admin').value = state.adminDefault || '';
  document.getElementById('s-naik-ka').value = state.kaCabang || '';
  document.getElementById('s-turun-admin').value = state.adminDefault || '';
  document.getElementById('s-turun-ka').value = state.kaCabang || '';
  document.getElementById('f-cabang').value = state.cabang;
});

function initDropdowns() {
  const salesSel = document.getElementById('f-salesman');
  const filterSal = document.getElementById('filter-salesman');
  salesSel.innerHTML = '<option value="">-- Pilih Salesman --</option>';
  filterSal.innerHTML = '<option value="">Semua Salesman</option>';
  state.salesmen.forEach(s => {
    salesSel.innerHTML += `<option value="${s}">${s}</option>`;
    filterSal.innerHTML += `<option value="${s}">${s}</option>`;
  });
  
  const carSel = document.getElementById('f-polisi');
  carSel.innerHTML = '<option value="">-- Pilih Mobil --</option>';
  state.cars.forEach(c => {
    carSel.innerHTML += `<option value="${c}">${c}</option>`;
  });
}

function renderProductTable() {
  const tbody = document.getElementById('product-table-body');
  tbody.innerHTML = '';
  state.products.forEach((p, i) => {
    tbody.innerHTML += `
      <tr>
        <td style="text-align:center;color:var(--text3)">${i+1}</td>
        <td>${p}</td>
        <td style="text-align:center"><input type="number" id="naik-${i}" min="0" placeholder="0" oninput="calcJual(${i})"></td>
        <td style="text-align:center"><input type="number" id="turun-${i}" min="0" placeholder="0" oninput="calcJual(${i})"></td>
        <td style="text-align:center"><input type="number" id="jual-${i}" min="0" placeholder="0" style="background:rgba(63,185,80,0.08)"></td>
      </tr>`;
  });
}

function calcJual(i) {
  const naik = parseInt(document.getElementById(`naik-${i}`).value) || 0;
  const turun = parseInt(document.getElementById(`turun-${i}`).value) || 0;
  const jual = naik - turun;
  document.getElementById(`jual-${i}`).value = jual >= 0 ? jual : '';
}

// ===================== NAVIGATION =====================
function showPage(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + page).classList.add('active');
  event?.target?.closest('.nav-item')?.classList.add('active');
  // force nav highlight
  document.querySelectorAll('.nav-item').forEach(n => {
    if (n.getAttribute('onclick')?.includes(page)) n.classList.add('active');
  });
  if (page === 'records') renderRecordsTable();
  if (page === 'dashboard') updateDashboard();
}

// ===================== FORM OPERATIONS =====================
function getFormData() {
  const products = state.products.map((p, i) => ({
    name: p,
    naik: parseInt(document.getElementById(`naik-${i}`).value) || 0,
    turun: parseInt(document.getElementById(`turun-${i}`).value) || 0,
    jual: parseInt(document.getElementById(`jual-${i}`).value) || 0
  }));

  return {
    id: state.currentFormId || Date.now(),
    nomer: document.getElementById('f-nomer').value,
    cabang: document.getElementById('f-cabang').value,
    tanggal: document.getElementById('f-tanggal').value,
    tis: document.getElementById('f-tis').value,
    salesman: document.getElementById('f-salesman').value,
    polisi: document.getElementById('f-polisi').value,
    catatan: document.getElementById('f-catatan').value,
    products,
    signs: {
      naikChecker: document.getElementById('s-naik-checker').value,
      naikDriver: document.getElementById('s-naik-driver').value,
      naikAdmin: document.getElementById('s-naik-admin').value,
      naikKa: document.getElementById('s-naik-ka').value,
      turunDriver: document.getElementById('s-turun-driver').value,
      turunChecker: document.getElementById('s-turun-checker').value,
      turunAdmin: document.getElementById('s-turun-admin').value,
      turunKa: document.getElementById('s-turun-ka').value,
    },
    status: getFormStatus(products),
    savedAt: new Date().toISOString()
  };
}

function getFormStatus(products) {
  const hasNaik = products.some(p => p.naik > 0);
  const hasTurun = products.some(p => p.turun > 0);
  if (!hasNaik) return 'draft';
  if (!hasTurun) return 'pagi';
  return 'selesai';
}

function saveForm() {
  const data = getFormData();
  if (!data.nomer) { alert('Nomor form wajib diisi!'); return; }
  if (!data.salesman) { alert('Salesman wajib dipilih!'); return; }
  
  const idx = state.records.findIndex(r => r.id === data.id);
  if (idx >= 0) state.records[idx] = data;
  else state.records.unshift(data);
  
  state.currentFormId = data.id;
  saveState();
  updateDashboard();
  
  showAlert('Form berhasil disimpan! ✅');
  updatePhaseIndicator(data.status);
}

function showAlert(msg) {
  const a = document.createElement('div');
  a.className = 'alert alert-success';
  a.style.cssText = 'position:fixed;bottom:24px;right:24px;z-index:9999;min-width:250px;animation:fadeIn 0.3s';
  a.textContent = msg;
  document.body.appendChild(a);
  setTimeout(() => a.remove(), 3000);
}

function updatePhaseIndicator(status) {
  document.querySelectorAll('.phase-step').forEach(s => { s.classList.remove('active','done'); });
  if (status === 'draft') {
    document.getElementById('phase-1').classList.add('active');
  } else if (status === 'pagi') {
    document.getElementById('phase-1').classList.add('done');
    document.getElementById('phase-2').classList.add('active');
  } else {
    document.getElementById('phase-1').classList.add('done');
    document.getElementById('phase-2').classList.add('done');
    document.getElementById('phase-3').classList.add('active');
  }
}

function resetForm() {
  if (!confirm('Reset form? Semua input akan dihapus.')) return;
  state.currentFormId = null;
  document.getElementById('f-nomer').value = '';
  document.getElementById('f-tis').value = '';
  document.getElementById('f-salesman').value = '';
  document.getElementById('f-polisi').value = '';
  document.getElementById('f-catatan').value = '';
  document.getElementById('f-tanggal').value = new Date().toISOString().split('T')[0];
  renderProductTable();
  document.getElementById('s-naik-checker').value = '';
  document.getElementById('s-naik-driver').value = '';
  document.getElementById('s-naik-admin').value = state.adminDefault || '';
  document.getElementById('s-naik-ka').value = state.kaCabang || '';
  document.getElementById('s-turun-checker').value = '';
  document.getElementById('s-turun-driver').value = '';
  document.getElementById('s-turun-admin').value = state.adminDefault || '';
  document.getElementById('s-turun-ka').value = state.kaCabang || '';
  document.querySelectorAll('.phase-step').forEach(s => { s.classList.remove('active','done'); });
  document.getElementById('phase-1').classList.add('active');
}

// ===================== EXPORT PDF =====================
function exportPDF() {
  saveForm();
  const data = getFormData();
  buildPDFContent(data);
  setTimeout(() => window.print(), 300);
}

function buildPDFContent(data) {
  const totalNaik = data.products.reduce((s,p) => s+p.naik, 0);
  const totalTurun = data.products.reduce((s,p) => s+p.turun, 0);
  const totalJual = data.products.reduce((s,p) => s+p.jual, 0);
  const tgl = data.tanggal ? new Date(data.tanggal+'T00:00:00').toLocaleDateString('id-ID', {day:'2-digit',month:'2-digit',year:'numeric'}) : '';

  const rows = data.products.map((p,i) => `
    <tr>
      <td>${i+1}</td>
      <td class="product-name">${p.name}</td>
      <td>${p.naik || ''}</td>
      <td>${p.turun || ''}</td>
      <td>${p.jual || ''}</td>
    </tr>`).join('');

  document.getElementById('pdf-content').innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
      <div>
        <div style="font-size:8px;color:#555">PT. SENTRALSARI PRIMASENTOSA</div>
        <div style="font-size:24px;font-weight:900;color:#1a5fba;letter-spacing:-1px">SPS</div>
        <div style="font-size:7px;letter-spacing:2px;color:#888">DISTRIBUTION PARTNER</div>
      </div>
      <div style="text-align:right">
        <div style="border:2px solid black;padding:4px 12px;font-size:14px;font-weight:900">No: ${data.nomer || 'A - _______'}</div>
      </div>
    </div>
    <div style="text-align:center;font-size:16px;font-weight:900;letter-spacing:2px;border-top:2px solid black;border-bottom:1px solid black;padding:4px 0;margin-bottom:8px">
      TRANSFER ANTAR GUDANG
    </div>
    <table style="width:100%;margin-bottom:6px;font-size:10px">
      <tr>
        <td style="width:100px">Cabang</td>
        <td style="width:10px">:</td>
        <td style="font-weight:700">${data.cabang}</td>
        <td style="width:100px">Nomer TIS</td>
        <td style="width:10px">:</td>
        <td>${data.tis || ''}</td>
      </tr>
      <tr>
        <td>Tanggal</td><td>:</td>
        <td>${tgl}</td>
        <td>Sopir / Salesman</td><td>:</td>
        <td style="font-weight:700">${data.salesman}</td>
      </tr>
      <tr>
        <td>Rit</td><td>:</td><td></td>
        <td>No. Polisi</td><td>:</td>
        <td style="font-weight:700">${data.polisi}</td>
      </tr>
    </table>
    <table class="pdf-main-table">
      <thead>
        <tr>
          <th style="width:30px">No</th>
          <th colspan="2">Produk</th>
          <th style="width:60px">Naik</th>
          <th style="width:60px">Turun</th>
          <th style="width:60px">Jual</th>
        </tr>
      </thead>
      <tbody>
        ${data.products.map((p,i) => `
        <tr>
          <td>${i+1}</td>
          <td style="width:60px;border-right:none">${i===0?'Cleo':''}</td>
          <td class="product-name" style="border-left:none">${p.name}</td>
          <td>${p.naik||''}</td>
          <td>${p.turun||''}</td>
          <td>${p.jual||''}</td>
        </tr>`).join('')}
        <tr style="font-weight:700;background:#f5f5f5">
          <td colspan="3" style="text-align:right">TOTAL</td>
          <td>${totalNaik}</td>
          <td>${totalTurun}</td>
          <td>${totalJual}</td>
        </tr>
      </tbody>
    </table>
    ${data.catatan ? `<div style="font-size:9px;margin-bottom:6px;padding:4px;border:1px solid #ccc">Catatan: ${data.catatan}</div>` : ''}
    <div style="display:flex;gap:8px;margin-top:8px">
      <div style="flex:1;border:1px solid black;padding:6px">
        <div style="font-weight:700;font-size:9px;text-align:center;border-bottom:1px solid black;padding-bottom:3px;margin-bottom:4px">Otorisasi TAG Naik</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
          ${['Diserahkan','Diterima','Diperiksa','Mengetahui'].map((r,i) => `
          <div style="text-align:center;font-size:8px">
            <div style="height:30px;border-bottom:1px solid #aaa;margin-bottom:2px"></div>
            <div style="font-weight:700">${r}</div>
            <div style="font-style:italic;color:#444">${i===0?data.signs.naikChecker:i===1?data.signs.naikDriver:i===2?data.signs.naikAdmin:data.signs.naikKa}</div>
            <div style="color:#666">${['Checker','Driver/Salesman','Admin Penjualan','Ka. Cabang'][i]}</div>
          </div>`).join('')}
        </div>
      </div>
      <div style="flex:1;border:1px solid black;padding:6px">
        <div style="font-weight:700;font-size:9px;text-align:center;border-bottom:1px solid black;padding-bottom:3px;margin-bottom:4px">Otorisasi TAG Turun</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
          ${['Diserahkan','Diterima','Diperiksa','Mengetahui'].map((r,i) => `
          <div style="text-align:center;font-size:8px">
            <div style="height:30px;border-bottom:1px solid #aaa;margin-bottom:2px"></div>
            <div style="font-weight:700">${r}</div>
            <div style="font-style:italic;color:#444">${i===0?data.signs.turunDriver:i===1?data.signs.turunChecker:i===2?data.signs.turunAdmin:data.signs.turunKa}</div>
            <div style="color:#666">${['Driver/Salesman','Checker','Admin Penjualan','Ka. Cabang'][i]}</div>
          </div>`).join('')}
        </div>
      </div>
    </div>
  `;
}

// ===================== EXPORT EXCEL =====================
function buildWorkbook(data) {
  const wb = XLSX.utils.book_new();
  
  // Sheet 1: Form Data
  const rows = [
    ['PT. SENTRALSARI PRIMASENTOSA - TRANSFER ANTAR GUDANG'],
    [],
    ['No. Form', data.nomer, '', 'Nomer TIS', data.tis],
    ['Cabang', data.cabang, '', 'Sopir/Salesman', data.salesman],
    ['Tanggal', data.tanggal, '', 'No. Polisi', data.polisi],
    [],
    ['No', 'Nama Produk', 'NAIK', 'TURUN', 'JUAL'],
    ...data.products.map((p, i) => [i+1, p.name, p.naik||0, p.turun||0, p.jual||0]),
    [],
    ['', 'TOTAL',
      data.products.reduce((s,p)=>s+p.naik,0),
      data.products.reduce((s,p)=>s+p.turun,0),
      data.products.reduce((s,p)=>s+p.jual,0)
    ],
    [],
    ['OTORISASI TAG NAIK'],
    ['Checker', data.signs.naikChecker, 'Driver/Salesman', data.signs.naikDriver],
    ['Admin Penjualan', data.signs.naikAdmin, 'Ka. Cabang', data.signs.naikKa],
    [],
    ['OTORISASI TAG TURUN'],
    ['Driver/Salesman', data.signs.turunDriver, 'Checker', data.signs.turunChecker],
    ['Admin Penjualan', data.signs.turunAdmin, 'Ka. Cabang', data.signs.turunKa],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:5},{wch:30},{wch:12},{wch:12},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws, 'Form');
  
  return wb;
}

function exportExcel() {
  saveForm();
  const data = getFormData();
  const wb = buildWorkbook(data);
  const fn = `Transfer_${data.nomer.replace(/\s/g,'_')}_${data.tanggal}.xlsx`;
  XLSX.writeFile(wb, fn);
  showAlert('File Excel berhasil diunduh! 📊');
}

function exportAllExcel() {
  if (state.records.length === 0) { alert('Belum ada data untuk diekspor.'); return; }
  const wb = XLSX.utils.book_new();
  
  // Summary sheet
  const headers = ['No. Form','Tanggal','Salesman','No. Polisi','Total Naik','Total Turun','Total Jual','Status'];
  const rows = state.records.map(r => [
    r.nomer, r.tanggal, r.salesman, r.polisi,
    r.products.reduce((s,p)=>s+p.naik,0),
    r.products.reduce((s,p)=>s+p.turun,0),
    r.products.reduce((s,p)=>s+p.jual,0),
    r.status
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [{wch:15},{wch:12},{wch:15},{wch:12},{wch:12},{wch:12},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws, 'Rekap Semua');
  
  // Per-product pivot
  const pivotHeaders = ['Tanggal','No.Form','Salesman', ...state.products];
  const pivotRows = state.records.map(r => [
    r.tanggal, r.nomer, r.salesman,
    ...state.products.map(pname => {
      const p = r.products.find(x => x.name === pname);
      return p ? p.jual : 0;
    })
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet([pivotHeaders, ...pivotRows]);
  XLSX.utils.book_append_sheet(wb, ws2, 'Pivot Produk Terjual');
  
  XLSX.writeFile(wb, `SPS_Rekap_${new Date().toISOString().split('T')[0]}.xlsx`);
  showAlert('Rekap Excel berhasil diunduh! 📊');
}

// ===================== INPUT SPS =====================
function inputSPS() {
  saveForm();
  const data = getFormData();
  if (!data.nomer || !data.salesman) { alert('Lengkapi form terlebih dahulu!'); return; }
  
  // Build SPS import Excel
  const wb = XLSX.utils.book_new();
  const rows = [
    ['=== FORMAT IMPORT SPS ==='],
    ['Tanggal Input:', new Date().toLocaleString('id-ID')],
    ['No. Form:', data.nomer],
    ['Nomer TIS:', data.tis],
    ['Salesman:', data.salesman],
    ['No. Polisi:', data.polisi],
    ['Tanggal:', data.tanggal],
    ['Cabang:', data.cabang],
    [],
    ['Kode Produk', 'Nama Produk', 'Qty Naik', 'Qty Turun', 'Qty Jual', 'Satuan'],
    ...data.products.filter(p => p.naik > 0 || p.jual > 0).map((p, i) => [
      `PROD-${String(i+1).padStart(3,'0')}`, p.name, p.naik, p.turun, p.jual, 'pcs'
    ])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:15},{wch:30},{wch:12},{wch:12},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws, 'SPS_Import');
  
  const fn = `SPS_Import_${data.nomer.replace(/\s/g,'_')}_${data.tanggal}.xlsx`;
  
  document.getElementById('sps-modal-content').innerHTML = `
    <table style="width:100%;font-size:12px">
      <tr><td style="color:var(--text2);padding:4px 0">No. Form</td><td><strong>${data.nomer}</strong></td></tr>
      <tr><td style="color:var(--text2);padding:4px 0">Salesman</td><td>${data.salesman}</td></tr>
      <tr><td style="color:var(--text2);padding:4px 0">Tanggal</td><td>${data.tanggal}</td></tr>
      <tr><td style="color:var(--text2);padding:4px 0">Total Produk</td><td>${data.products.filter(p=>p.naik>0||p.jual>0).length} item</td></tr>
      <tr><td style="color:var(--text2);padding:4px 0">File</td><td><code style="font-family:monospace;font-size:11px;color:var(--accent2)">${fn}</code></td></tr>
    </table>`;
  
  document.getElementById('sps-download-btn').onclick = () => {
    XLSX.writeFile(wb, fn);
    showAlert('File SPS berhasil diunduh! ⚡');
    closeModal('modal-sps');
  };
  
  document.getElementById('modal-sps').classList.add('open');
}

// ===================== DASHBOARD =====================
function updateDashboard() {
  const today = new Date().toISOString().split('T')[0];
  const thisMonth = today.substring(0, 7);
  
  const todayForms = state.records.filter(r => r.tanggal === today);
  const monthForms = state.records.filter(r => r.tanggal?.startsWith(thisMonth));
  const pending = state.records.filter(r => r.status !== 'selesai');
  
  document.getElementById('stat-today').textContent = todayForms.length;
  document.getElementById('stat-month').textContent = monthForms.length;
  document.getElementById('stat-pending').textContent = pending.length;
  document.getElementById('stat-sales').textContent = state.salesmen.length;
  document.getElementById('sidebar-branch').textContent = state.cabang;
  
  const tbody = document.getElementById('dashboard-table');
  const recent = state.records.slice(0, 8);
  if (recent.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:20px">Belum ada data</td></tr>';
    return;
  }
  tbody.innerHTML = recent.map(r => `
    <tr>
      <td><span style="font-family:monospace;color:var(--accent2)">${r.nomer}</span></td>
      <td>${r.tanggal}</td>
      <td>${r.salesman}</td>
      <td>${r.polisi}</td>
      <td><span class="badge badge-${r.status === 'selesai' ? 'selesai' : r.status === 'pagi' ? 'pagi' : 'pagi'}">${r.status === 'selesai' ? '✅ Selesai' : r.status === 'pagi' ? '🌅 Pagi' : '📝 Draft'}</span></td>
      <td style="display:flex;gap:4px">
        <button class="btn btn-sm btn-outline" onclick="loadFormForEdit(${r.id})">✏️ Edit</button>
        <button class="btn btn-sm btn-outline" onclick="viewRecord(${r.id})">👁️</button>
      </td>
    </tr>`).join('');
}

// ===================== RECORDS =====================
function renderRecordsTable(filtered) {
  const data = filtered || state.records;
  const tbody = document.getElementById('records-table');
  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text3);padding:20px">Belum ada data</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(r => `
    <tr>
      <td><span style="font-family:monospace;font-size:12px;color:var(--accent2)">${r.nomer}</span></td>
      <td>${r.tanggal}</td>
      <td>${r.salesman}</td>
      <td>${r.polisi}</td>
      <td style="text-align:center">${r.products.reduce((s,p)=>s+p.naik,0)}</td>
      <td style="text-align:center">${r.products.reduce((s,p)=>s+p.turun,0)}</td>
      <td style="text-align:center;font-weight:700;color:var(--green)">${r.products.reduce((s,p)=>s+p.jual,0)}</td>
      <td><span class="badge badge-${r.status === 'selesai' ? 'selesai' : 'pagi'}">${r.status === 'selesai' ? '✅ Selesai' : r.status === 'pagi' ? '🌅 Pagi' : '📝 Draft'}</span></td>
      <td style="display:flex;gap:4px">
        <button class="btn btn-sm btn-outline" onclick="loadFormForEdit(${r.id})">✏️ Edit</button>
        <button class="btn btn-sm btn-outline" onclick="viewRecord(${r.id})">👁️ Lihat</button>
        <button class="btn btn-sm btn-danger" onclick="deleteRecord(${r.id})">🗑️</button>
      </td>
    </tr>`).join('');
}

function filterRecords() {
  const dari = document.getElementById('filter-dari').value;
  const sampai = document.getElementById('filter-sampai').value;
  const salesman = document.getElementById('filter-salesman').value;
  
  let filtered = state.records;
  if (dari) filtered = filtered.filter(r => r.tanggal >= dari);
  if (sampai) filtered = filtered.filter(r => r.tanggal <= sampai);
  if (salesman) filtered = filtered.filter(r => r.salesman === salesman);
  renderRecordsTable(filtered);
}

function clearFilter() {
  document.getElementById('filter-dari').value = '';
  document.getElementById('filter-sampai').value = '';
  document.getElementById('filter-salesman').value = '';
  renderRecordsTable();
}

function deleteRecord(id) {
  if (!confirm('Hapus form ini?')) return;
  state.records = state.records.filter(r => r.id !== id);
  saveState();
  renderRecordsTable();
  updateDashboard();
  showAlert('Form berhasil dihapus');
}

function loadFormForEdit(id) {
  const r = state.records.find(x => x.id === id);
  if (!r) return;
  state.currentFormId = id;
  document.getElementById('f-nomer').value = r.nomer || '';
  document.getElementById('f-cabang').value = r.cabang || state.cabang;
  document.getElementById('f-tanggal').value = r.tanggal || '';
  document.getElementById('f-tis').value = r.tis || '';
  document.getElementById('f-salesman').value = r.salesman || '';
  document.getElementById('f-polisi').value = r.polisi || '';
  document.getElementById('f-catatan').value = r.catatan || '';
  
  // Products
  state.products.forEach((p, i) => {
    const prod = r.products.find(x => x.name === p);
    if (prod) {
      document.getElementById(`naik-${i}`).value = prod.naik || '';
      document.getElementById(`turun-${i}`).value = prod.turun || '';
      document.getElementById(`jual-${i}`).value = prod.jual || '';
    }
  });
  
  // Signs
  document.getElementById('s-naik-checker').value = r.signs?.naikChecker || '';
  document.getElementById('s-naik-driver').value = r.signs?.naikDriver || r.salesman || '';
  document.getElementById('s-naik-admin').value = r.signs?.naikAdmin || '';
  document.getElementById('s-naik-ka').value = r.signs?.naikKa || '';
  document.getElementById('s-turun-checker').value = r.signs?.turunChecker || '';
  document.getElementById('s-turun-driver').value = r.signs?.turunDriver || r.salesman || '';
  document.getElementById('s-turun-admin').value = r.signs?.turunAdmin || '';
  document.getElementById('s-turun-ka').value = r.signs?.turunKa || '';
  
  updatePhaseIndicator(r.status);
  showPage('form');
}

// ===================== VIEW MODAL =====================
let viewingId = null;
function viewRecord(id) {
  viewingId = id;
  const r = state.records.find(x => x.id === id);
  if (!r) return;
  
  const totalNaik = r.products.reduce((s,p)=>s+p.naik,0);
  const totalTurun = r.products.reduce((s,p)=>s+p.turun,0);
  const totalJual = r.products.reduce((s,p)=>s+p.jual,0);
  
  document.getElementById('modal-view-content').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;font-size:12px">
      <div><span style="color:var(--text2)">No. Form:</span> <strong style="color:var(--accent2)">${r.nomer}</strong></div>
      <div><span style="color:var(--text2)">Tanggal:</span> ${r.tanggal}</div>
      <div><span style="color:var(--text2)">Salesman:</span> ${r.salesman}</div>
      <div><span style="color:var(--text2)">No. Polisi:</span> ${r.polisi}</div>
    </div>
    <div class="table-wrap" style="margin-bottom:12px">
      <table>
        <thead><tr><th>Produk</th><th>Naik</th><th>Turun</th><th>Jual</th></tr></thead>
        <tbody>
          ${r.products.filter(p=>p.naik||p.turun||p.jual).map(p=>`
            <tr><td>${p.name}</td><td style="text-align:center">${p.naik||'-'}</td><td style="text-align:center">${p.turun||'-'}</td><td style="text-align:center;font-weight:700;color:var(--green)">${p.jual||'-'}</td></tr>`).join('')}
          <tr style="font-weight:700;background:var(--surface2)">
            <td>TOTAL</td><td style="text-align:center">${totalNaik}</td><td style="text-align:center">${totalTurun}</td><td style="text-align:center;color:var(--green)">${totalJual}</td>
          </tr>
        </tbody>
      </table>
    </div>`;
  
  document.getElementById('modal-view').classList.add('open');
}

function printFromModal() {
  const r = state.records.find(x => x.id === viewingId);
  if (!r) return;
  buildPDFContent(r);
  setTimeout(() => window.print(), 300);
}

function exportFromModal() {
  const r = state.records.find(x => x.id === viewingId);
  if (!r) return;
  const wb = buildWorkbook(r);
  XLSX.writeFile(wb, `Transfer_${r.nomer.replace(/\s/g,'_')}_${r.tanggal}.xlsx`);
}

// ===================== MASTER DATA =====================
function renderMasterData() {
  renderChipList('salesman-list', state.salesmen, 'salesman');
  renderChipList('car-list', state.cars, 'car');
  renderChipList('product-list', state.products, 'product');
  document.getElementById('m-cabang').value = state.cabang;
  document.getElementById('m-ka').value = state.kaCabang;
  document.getElementById('m-admin').value = state.adminDefault;
}

function renderChipList(elId, arr, type) {
  document.getElementById(elId).innerHTML = arr.map((item, i) => `
    <div class="chip">
      ${item}
      <span class="chip-del" onclick="removeItem('${type}', ${i})">×</span>
    </div>`).join('');
}

function removeItem(type, idx) {
  if (type === 'salesman') state.salesmen.splice(idx, 1);
  else if (type === 'car') state.cars.splice(idx, 1);
  else if (type === 'product') state.products.splice(idx, 1);
  saveState();
  renderMasterData();
  initDropdowns();
  renderProductTable();
}

function addSalesman() {
  const v = document.getElementById('new-salesman').value.trim();
  if (!v) return;
  state.salesmen.push(v);
  document.getElementById('new-salesman').value = '';
  saveState(); renderMasterData(); initDropdowns();
  showAlert(`Salesman "${v}" ditambahkan ✅`);
}

function addCar() {
  const v = document.getElementById('new-car').value.trim().toUpperCase();
  if (!v) return;
  state.cars.push(v);
  document.getElementById('new-car').value = '';
  saveState(); renderMasterData(); initDropdowns();
  showAlert(`Mobil "${v}" ditambahkan ✅`);
}

function addProduct() {
  const v = document.getElementById('new-product').value.trim();
  if (!v) return;
  state.products.push(v);
  document.getElementById('new-product').value = '';
  saveState(); renderMasterData(); renderProductTable();
  showAlert(`Produk "${v}" ditambahkan ✅`);
}

function saveMaster() {
  state.cabang = document.getElementById('m-cabang').value.trim().toUpperCase() || state.cabang;
  state.kaCabang = document.getElementById('m-ka').value.trim();
  state.adminDefault = document.getElementById('m-admin').value.trim();
  saveState();
  document.getElementById('sidebar-branch').textContent = state.cabang;
  document.getElementById('f-cabang').value = state.cabang;
  showAlert('Pengaturan disimpan ✅');
}

// ===================== MODAL =====================
function closeModal(id) { document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('.modal-overlay').forEach(m => {
  m.addEventListener('click', e => { if (e.target === m) m.classList.remove('open'); });
});
</script>
</body>
</html>
