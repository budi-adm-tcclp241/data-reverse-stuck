"""
Excel Explorer — Single-file Flask App
Deploy ke Railway / GitHub Codespaces:
  1. pip install -r requirements.txt
  2. python app.py
"""

import io, os, traceback
from flask import Flask, request, jsonify, Response

# ─────────────────────────────────────────────────────────────────────────────
#  ⚙️  KONFIGURASI — Edit sesuai kebutuhan Anda
# ─────────────────────────────────────────────────────────────────────────────

# Kolom yang ingin ditampilkan (None = semua kolom)
TARGET_COLUMNS = None
# TARGET_COLUMNS = None

# Kolom & nilai yang digunakan untuk filter
FILTER_COLUMN = 'DP Terjadwal'
FILTER_VALUES = [
    'CILACAP', 'KROYA', 'CLP03A', 'ADIPALA',
    'MERTASINGA', 'SALIWANGI', 'NUSAWUNGU'
]

# Kolom yang digunakan untuk pengurutan
SORT_COLUMNS = ['DP Terjadwal', 'Sumber Order', 'Status Orderan']

# Nama file hasil export CSV
OUTPUT_FILENAME = 'hasil_filter.csv'

# Teks kustom sebelum/sesudah daftar AWB per status
TEXT_DISPATCH_TOP    = ''
TEXT_DISPATCH_BOTTOM = ''
TEXT_PUSAT_TOP       = ''
TEXT_PUSAT_BOTTOM    = ''

# ─────────────────────────────────────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Excel Explorer</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
/* ── Reset & Base ── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#eef2ff;
  --bg2:#f8f9ff;
  --surface:#ffffff;
  --surface2:#f4f6ff;
  --border:#dde3f5;
  --border2:#c7d0ef;
  --ac:#4f46e5;
  --ac-h:#4338ca;
  --ac-lt:#ede9fe;
  --ac-lter:#f5f3ff;
  --tx:#1e2147;
  --tx2:#4a5080;
  --mu:#7c85b3;
  --ok:#059669;
  --ok-bg:#d1fae5;
  --ok-lt:#ecfdf5;
  --wn:#d97706;
  --wn-bg:#fef3c7;
  --er:#dc2626;
  --er-bg:#fee2e2;
  --shadow:0 1px 3px rgba(79,70,229,.08),0 4px 16px rgba(79,70,229,.06);
  --shadow-lg:0 4px 24px rgba(79,70,229,.12),0 1px 4px rgba(79,70,229,.08);
  --radius:14px;
  --radius-sm:8px;
  --radius-xs:5px;
}

html{scroll-behavior:smooth}

body{
  font-family:'Plus Jakarta Sans',system-ui,sans-serif;
  background:var(--bg);
  color:var(--tx);
  min-height:100vh;
  line-height:1.6;
}

/* ── Layout ── */
.wrap{max-width:1100px;margin:0 auto;padding:24px 20px 60px}

/* ── Header ── */
.hdr{
  text-align:center;
  padding:40px 0 32px;
}
.hdr-badge{
  display:inline-flex;align-items:center;gap:6px;
  background:var(--ac-lt);color:var(--ac);
  border:1px solid rgba(79,70,229,.2);
  font-size:.75rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  padding:4px 12px;border-radius:20px;margin-bottom:16px;
}
.hdr h1{
  font-size:2.4rem;font-weight:800;letter-spacing:-.03em;
  color:var(--tx);line-height:1.2;
  background:linear-gradient(135deg,#312e81,#4f46e5,#7c3aed);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
}
.hdr p{color:var(--tx2);margin-top:8px;font-size:1rem;font-weight:500}

/* ── Notice Banner ── */
.notice{
  background:linear-gradient(135deg,var(--ac-lter),#faf5ff);
  border:1px solid rgba(79,70,229,.18);
  border-left:4px solid var(--ac);
  border-radius:var(--radius-sm);
  padding:13px 18px;
  color:var(--tx2);font-size:.875rem;
  margin-bottom:22px;display:flex;align-items:flex-start;gap:10px;
}
.notice-icon{font-size:1.1rem;flex-shrink:0;margin-top:1px}
.notice code{
  background:rgba(79,70,229,.1);color:var(--ac);
  padding:1px 6px;border-radius:4px;
  font-family:'JetBrains Mono',monospace;font-size:.82rem;
}

/* ── Upload Zone ── */
.uz{
  border:2.5px dashed var(--border2);
  border-radius:var(--radius);
  padding:52px 24px;
  text-align:center;
  cursor:pointer;
  transition:all .22s ease;
  background:var(--surface);
  box-shadow:var(--shadow);
  margin-bottom:24px;
  position:relative;overflow:hidden;
}
.uz::before{
  content:'';position:absolute;inset:0;
  background:linear-gradient(135deg,rgba(79,70,229,.03),rgba(124,58,237,.03));
  opacity:0;transition:opacity .22s;
}
.uz:hover,.uz.over{
  border-color:var(--ac);
  background:var(--ac-lter);
  box-shadow:var(--shadow-lg);
}
.uz:hover::before,.uz.over::before{opacity:1}
.uico{font-size:3.5rem;margin-bottom:14px;display:block;filter:drop-shadow(0 4px 8px rgba(79,70,229,.2))}
.uz h3{font-size:1.15rem;font-weight:700;margin-bottom:6px;color:var(--tx)}
.uz .sub{color:var(--mu);font-size:.9rem;margin-bottom:22px}
.btn-up{
  background:var(--ac);color:#fff;border:none;
  padding:11px 28px;border-radius:var(--radius-sm);
  font-size:.95rem;cursor:pointer;
  transition:all .18s;font-weight:700;
  font-family:inherit;
  box-shadow:0 3px 12px rgba(79,70,229,.35);
}
.btn-up:hover{background:var(--ac-h);transform:translateY(-1px);box-shadow:0 6px 20px rgba(79,70,229,.4)}
.btn-up:active{transform:translateY(0)}

/* ── Card ── */
.card{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  padding:22px 26px;
  margin-bottom:18px;
  box-shadow:var(--shadow);
}
.card-title{
  font-size:.95rem;font-weight:700;color:var(--tx);
  margin-bottom:18px;display:flex;align-items:center;gap:8px;
}
.card-title .ic{font-size:1.1rem}

/* ── Info Grid ── */
.ig{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}
.igi{
  background:linear-gradient(135deg,var(--surface2),#f0f4ff);
  border:1px solid var(--border);
  border-radius:var(--radius-sm);padding:13px 15px;
}
.igl{font-size:.7rem;color:var(--mu);text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px;font-weight:600}
.igv{font-size:1.05rem;font-weight:700;color:var(--tx);word-break:break-all}

/* ── Column Tags ── */
.tags{display:flex;flex-wrap:wrap;gap:7px}
.tag{
  background:var(--ac-lt);
  border:1px solid rgba(79,70,229,.2);
  color:var(--ac);
  padding:4px 10px;border-radius:20px;
  font-size:.78rem;font-weight:600;
  font-family:'JetBrains Mono',monospace;
}

/* ── Process Log ── */
.step{display:flex;align-items:flex-start;gap:13px;padding:12px 0;border-bottom:1px solid var(--border)}
.step:last-child{border-bottom:none}
.si{
  width:28px;height:28px;border-radius:50%;
  display:flex;align-items:center;justify-content:center;
  font-size:.8rem;flex-shrink:0;font-weight:700;
}
.si.ok{background:var(--ok-bg);color:var(--ok)}
.si.wn{background:var(--wn-bg);color:var(--wn)}
.sb{flex:1}
.st{font-weight:600;font-size:.9rem;color:var(--tx)}
.sd{color:var(--tx2);font-size:.83rem;margin-top:3px}

/* ── Stats Row ── */
.stats{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.stat{
  flex:1;min-width:110px;text-align:center;
  background:linear-gradient(135deg,#f5f3ff,#ede9fe);
  border:1px solid rgba(79,70,229,.2);
  border-radius:var(--radius-sm);padding:18px 12px;
}
.stn{font-size:2rem;font-weight:800;color:var(--ac);line-height:1;letter-spacing:-.03em}
.stl{font-size:.75rem;color:var(--mu);margin-top:6px;font-weight:600;text-transform:uppercase;letter-spacing:.06em}

/* ── Download Button ── */
.btn-dl{
  display:inline-flex;align-items:center;gap:8px;
  background:linear-gradient(135deg,var(--ok),#10b981);
  color:#fff;border:none;
  padding:13px 32px;border-radius:var(--radius-sm);
  font-size:.97rem;cursor:pointer;font-weight:700;
  transition:all .18s;font-family:inherit;
  box-shadow:0 3px 14px rgba(5,150,105,.35);
}
.btn-dl:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(5,150,105,.4)}
.btn-dl:active{transform:translateY(0)}

/* ── Table Wrapper ── */
.tw{overflow-x:auto;max-height:520px;border:1px solid var(--border);border-radius:var(--radius-sm);margin-top:16px}
table{width:100%;border-collapse:collapse;font-size:.82rem}
thead th{
  background:linear-gradient(135deg,var(--surface2),#eef2ff);
  padding:10px 14px;text-align:left;
  font-weight:700;color:var(--tx2);
  white-space:nowrap;border-bottom:2px solid var(--border2);
  position:sticky;top:0;z-index:1;
}
tbody tr:hover{background:var(--ac-lter)}
tbody td{
  padding:8px 14px;
  border-bottom:1px solid var(--border);
  white-space:nowrap;max-width:220px;
  overflow:hidden;text-overflow:ellipsis;
  color:var(--tx2);
}

/* ── Loading ── */
.loading{text-align:center;padding:64px 0}
.spinner{
  width:48px;height:48px;
  border:3px solid var(--border2);
  border-top-color:var(--ac);
  border-radius:50%;
  animation:spin .75s linear infinite;
  margin:0 auto 16px;
}
@keyframes spin{to{transform:rotate(360deg)}}
.loading p{color:var(--mu);font-weight:500}

/* ── Error Box ── */
.err-box{
  background:var(--er-bg);
  border:1px solid rgba(220,38,38,.25);
  border-left:4px solid var(--er);
  border-radius:var(--radius-sm);
  padding:14px 18px;color:#991b1b;
  margin-top:12px;font-size:.9rem;
}

/* ── Accordion ── */
.acc-hdr{cursor:pointer;user-select:none;display:flex;align-items:center;justify-content:space-between;margin-bottom:0}
.acc-hdr:hover .acc-arrow{color:var(--ac)}
.acc-hdr .card-title{margin-bottom:0;pointer-events:none}
.acc-arrow{font-size:.85rem;color:var(--mu);transition:transform .22s ease;display:inline-block;margin-left:10px;flex-shrink:0;width:20px;height:20px;background:var(--surface2);border:1px solid var(--border);border-radius:50%;display:flex;align-items:center;justify-content:center}
.acc-arrow.open{transform:rotate(90deg)}
.acc-body{overflow:hidden}
.acc-body.closed{display:none}

/* ── DP Cards ── */
.dp-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px;margin-top:10px}
.dp-card{
  background:linear-gradient(135deg,var(--surface),#fafbff);
  border:1px solid var(--border);
  border-radius:var(--radius);
  padding:18px 20px;
  display:flex;flex-direction:column;gap:12px;
  box-shadow:0 2px 8px rgba(79,70,229,.06);
  transition:box-shadow .18s;
}
.dp-card:hover{box-shadow:var(--shadow-lg)}
.dp-card-hdr{display:flex;align-items:center;justify-content:space-between}
.dp-title{font-size:.95rem;font-weight:800;color:var(--ac)}
.dp-badge{
  font-size:.72rem;padding:3px 9px;border-radius:20px;font-weight:700;
  background:var(--ok-bg);border:1px solid rgba(5,150,105,.2);color:var(--ok);
}
.dp-badge.mt{background:var(--surface2);border-color:var(--border2);color:var(--mu)}
.dp-ta{
  width:100%;min-height:160px;
  background:var(--bg2);
  border:1.5px solid var(--border);
  border-radius:var(--radius-sm);
  color:var(--tx);
  font-family:'JetBrains Mono',monospace;font-size:.8rem;
  padding:11px 13px;resize:vertical;
  line-height:1.7;outline:none;
  transition:border-color .18s;
}
.dp-ta:focus{border-color:var(--ac);background:var(--surface)}
.btn-copy{
  background:var(--ac-lt);
  border:1px solid rgba(79,70,229,.2);
  color:var(--ac);
  padding:9px 16px;border-radius:var(--radius-sm);
  font-size:.83rem;cursor:pointer;
  transition:all .18s;font-weight:700;
  width:100%;font-family:inherit;
}
.btn-copy:hover{background:var(--ac);color:#fff}
.btn-copy.copied{background:var(--ok-bg);border-color:rgba(5,150,105,.3);color:var(--ok)}

/* ── Pivot Table ── */
.pv-wrap{overflow-x:auto;margin-top:6px}
.pv-tbl{border-collapse:collapse;font-size:.85rem;min-width:560px;width:100%}
.pv-tbl th,.pv-tbl td{border:1px solid var(--border);padding:9px 16px;white-space:nowrap}
.pv-tbl thead th{
  background:linear-gradient(135deg,var(--surface2),#eef2ff);
  color:var(--tx2);font-weight:700;text-align:center;
}
.pv-tbl thead th:first-child{text-align:left}
.pv-tbl thead th.grp-hdr{
  background:linear-gradient(135deg,#ede9fe,#e0d9fb);
  color:var(--ac);font-size:.78rem;letter-spacing:.06em;text-transform:uppercase;
  border-bottom:2px solid rgba(79,70,229,.25);
}
.pv-tbl thead th.sub-hdr{font-size:.78rem;color:var(--tx2)}
.pv-tbl tbody td:first-child{font-weight:600;color:var(--tx);text-align:left}
.pv-tbl tbody td{text-align:center;color:var(--tx2)}
.pv-tbl tbody td.num{color:var(--ac);font-weight:700}
.pv-tbl tbody td.muted{color:var(--mu);font-weight:400;font-size:.8rem}
.pv-tbl .sep-r{border-right:2.5px solid var(--border2) !important}
.pv-tbl tbody tr:hover td{background:var(--ac-lter)}
.pv-tbl tfoot td{
  background:linear-gradient(135deg,#ede9fe,#ddd6fe);
  font-weight:800;color:var(--tx);
  text-align:center;border-top:2px solid var(--ac);
}
.pv-tbl tfoot td:first-child{text-align:left}
.pv-tbl tfoot td.num{color:var(--ac)}
.pv-tbl tfoot td.muted{color:var(--mu);font-weight:500}


/* ── Modal ── */
.modal-overlay{
  display:none;position:fixed;inset:0;
  background:rgba(15,20,50,.52);backdrop-filter:blur(4px);
  z-index:1000;align-items:center;justify-content:center;padding:20px;
}
.modal-overlay.open{display:flex;}
.modal-box{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:0 8px 56px rgba(79,70,229,.22),0 2px 12px rgba(0,0,0,.12);
  max-width:530px;width:100%;
  max-height:84vh;
  display:flex;flex-direction:column;
  animation:mfade .18s ease;
}
@keyframes mfade{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.modal-hdr{
  padding:15px 20px;border-bottom:1px solid var(--border);
  display:flex;align-items:center;justify-content:space-between;flex-shrink:0;
}
.modal-title{font-weight:800;font-size:.98rem;color:var(--ac);display:flex;align-items:center;gap:7px}
.modal-close{
  cursor:pointer;background:var(--surface2);
  border:1px solid var(--border);border-radius:50%;
  width:30px;height:30px;display:flex;align-items:center;justify-content:center;
  font-size:.85rem;color:var(--mu);transition:all .15s;flex-shrink:0;
}
.modal-close:hover{background:var(--er-bg);color:var(--er);border-color:rgba(220,38,38,.3)}
.modal-body{padding:18px 22px;overflow-y:auto;flex:1;min-height:0}
.modal-footer{
  padding:12px 20px;border-top:1px solid var(--border);
  flex-shrink:0;
}
.modal-raw{display:none}
/* Rendered markdown inside modal */
.md-h1{
  font-size:1rem;font-weight:800;color:var(--ac);
  margin:16px 0 8px;padding-bottom:5px;
  border-bottom:1.5px solid var(--border);
}
.md-h1:first-child{margin-top:0}
.md-bold{
  font-weight:700;font-size:.82rem;color:var(--tx);
  margin:10px 0 4px;
  background:var(--surface2);border:1px solid var(--border);
  padding:3px 10px;border-radius:var(--radius-xs);
  display:inline-block;letter-spacing:.04em;
}
.md-list{list-style:none;padding:0;margin:2px 0 8px 0}
.md-list li{
  font-family:'JetBrains Mono',monospace;font-size:.78rem;
  color:var(--tx2);padding:2px 0 2px 20px;position:relative;
  line-height:1.65;
}
.md-list li::before{content:'•';position:absolute;left:6px;color:var(--ac);font-size:1rem;line-height:1.4}
/* Clickable pivot rows */
.pv-tbl tbody tr.clickable{cursor:pointer}
.pv-tbl tbody tr.clickable:hover td{background:var(--ac-lt)!important}
.pv-tbl tbody tr.clickable td:first-child{
  color:var(--ac);text-decoration:underline dotted 1.5px;
  text-underline-offset:3px;
}
.pv-tbl tbody tr.clickable td:first-child::after{
  content:' 🔍';font-size:.7rem;opacity:.6;
}

/* ── Responsive ── */
@media(max-width:600px){
  .hdr h1{font-size:1.7rem}
  .ig{grid-template-columns:1fr 1fr}
  .stats{flex-direction:column}
  .dp-grid{grid-template-columns:1fr}
}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div class="hdr-badge">📊 Excel Explorer</div>
  <h1>Filter & Analisis Excel</h1>
  <p>Upload file .xlsx → Filter otomatis → Export ke CSV</p>
  <p>18/04/2026 11:09</p>
</div>

<div class="notice">
  <span class="notice-icon">💡</span>
  <span>Upload file <code>.xlsx</code> — sistem otomatis memfilter kolom <code>DP Terjadwal</code>, mengurutkan data, lalu menyiapkan CSV siap unduh.</span>
</div>

<div id="up-sec">
  <div class="uz" id="dz">
    <span class="uico">📂</span>
    <h3>Unggah File Excel</h3>
    <p class="sub">Drag &amp; Drop file .xlsx di sini, atau klik tombol di bawah</p>
    <button class="btn-up" id="pickBtn">📎 Pilih File .xlsx</button>
  </div>
  <input type="file" id="fi" accept=".xlsx" style="display:none">
</div>

<div id="ld" style="display:none" class="loading">
  <div class="spinner"></div>
  <p>Membaca &amp; memproses file Excel…</p>
</div>
<div id="er"></div>

<div id="res" style="display:none">

  <div class="card">
    <div class="card-title"><span class="ic">📄</span> Informasi File</div>
    <div class="ig" id="fi-grid"></div>
  </div>

  <div class="card">
    <div class="card-title"><span class="ic">🗂️</span> Kolom Tersedia <span id="ccnt" style="color:var(--mu);font-weight:500;font-size:.85rem"></span></div>
    <div class="tags" id="ctags"></div>
  </div>

  <div class="card">
    <div class="card-title"><span class="ic">⚙️</span> Pipeline Proses Otomatis</div>
    <div id="plog"></div>
  </div>

  <div class="card">
    <div class="card-title"><span class="ic">📈</span> Hasil &amp; Export</div>
    <div class="stats" id="srow"></div>
    <button class="btn-dl" id="dlBtn">⬇️ Download CSV</button>
  </div>

  <div class="card">
    <div class="acc-hdr" id="pvHdr">
      <div class="card-title"><span class="ic">🔍</span> Preview Data <span id="pvMeta" style="color:var(--mu);font-weight:500;font-size:.85rem"></span></div>
      <span class="acc-arrow" id="pvArrow">▶</span>
    </div>
    <div class="acc-body closed" id="pvBody">
      <div class="tw"><table id="ptbl"></table></div>
    </div>
  </div>

</div>

<div class="card" id="pv-sec" style="display:none">
  <div class="card-title"><span class="ic">📊</span> Pivot: DP Terjadwal × Status Orderan</div>
  <p style="color:var(--mu);font-size:.85rem;margin-bottom:16px">Jumlah AWB per DP berdasarkan status pengiriman &nbsp;·&nbsp; <span style="color:var(--ac);font-weight:600">🔍 Klik baris DP untuk melihat Rekap DP Terjadwal</span></p>
  <div class="pv-wrap">
    <table class="pv-tbl" id="pv-tbl"></table>
  </div>
</div>

<div class="card" id="dp-sec" style="display:none">
  <div class="card-title"><span class="ic">📋</span> Rekap per DP Terjadwal</div>
  <p style="color:var(--mu);font-size:.85rem;margin-bottom:8px">Setiap kartu mewakili satu nilai <strong>DP Terjadwal</strong> — klik tombol untuk menyalin isi textarea</p>
  <div class="dp-grid" id="dp-cards"></div>
</div>

<!-- Modal Rekap DP -->
<div class="modal-overlay" id="dpModal">
  <div class="modal-box">
    <div class="modal-hdr">
      <span class="modal-title" id="modalTitle">📍 —</span>
      <button class="modal-close" id="modalClose" title="Tutup">✕</button>
    </div>
    <div class="modal-body">
      <div id="modalRendered"></div>
    </div>
    <div class="modal-footer">
      <button class="btn-copy" id="modalCopy" style="width:100%">📋 Salin Teks (Markdown)</button>
      <textarea class="modal-raw" id="modalRaw" readonly></textarea>
    </div>
  </div>
</div>

</div>

<script>
(function(){
  var dz=document.getElementById('dz');
  var fileInput=document.getElementById('fi');
  var pickBtn=document.getElementById('pickBtn');
  var dlBtn=document.getElementById('dlBtn');

  // ── DP cards lookup (dp_value → markdown content) ──
  var dpLookup = {};

  pickBtn.addEventListener('click',function(e){
    e.stopPropagation();
    fileInput.value='';   // reset supaya file yang sama bisa di-upload ulang
    fileInput.click();
  });
  dz.addEventListener('click',function(e){
    // Jika klik dari pickBtn (sudah punya handler sendiri), abaikan
    if(e.target===pickBtn||pickBtn.contains(e.target)) return;
    fileInput.value='';
    fileInput.click();
  });
  dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('over');});
  dz.addEventListener('dragleave',function(){dz.classList.remove('over');});
  dz.addEventListener('drop',function(e){
    e.preventDefault();dz.classList.remove('over');
    if(e.dataTransfer.files[0])handle(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change',function(){
    if(fileInput.files[0])handle(fileInput.files[0]);
  });
  dlBtn.addEventListener('click',function(){window.location.href='/api/download';});

  document.getElementById('pvHdr').addEventListener('click',function(){
    var body=document.getElementById('pvBody');
    var arrow=document.getElementById('pvArrow');
    var closed=body.classList.contains('closed');
    body.classList.toggle('closed',!closed);
    arrow.classList.toggle('open',closed);
  });

  // ── Modal logic ──
  var modal      = document.getElementById('dpModal');
  var modalClose = document.getElementById('modalClose');
  var modalCopy  = document.getElementById('modalCopy');
  var modalRaw   = document.getElementById('modalRaw');

  function openModal(dp, content){
    document.getElementById('modalTitle').textContent = '📍 ' + dp;
    document.getElementById('modalRendered').innerHTML = markdownToHtml(content);
    modalRaw.value = content;
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  }
  function closeModal(){
    modal.classList.remove('open');
    document.body.style.overflow = '';
    modalCopy.innerHTML = '📋 Salin Teks (Markdown)';
    modalCopy.classList.remove('copied');
  }
  modalClose.addEventListener('click', closeModal);
  modal.addEventListener('click', function(e){ if(e.target===modal) closeModal(); });
  document.addEventListener('keydown', function(e){ if(e.key==='Escape') closeModal(); });

  modalCopy.addEventListener('click', function(){
    var text = modalRaw.value;
    var ok = function(){
      modalCopy.innerHTML = '✅ Tersalin!';
      modalCopy.classList.add('copied');
      setTimeout(function(){
        modalCopy.innerHTML = '📋 Salin Teks (Markdown)';
        modalCopy.classList.remove('copied');
      }, 2200);
    };
    if(navigator.clipboard && window.isSecureContext){
      navigator.clipboard.writeText(text).then(ok).catch(function(){
        try{ modalRaw.select(); document.execCommand('copy'); }catch(ex){}
        ok();
      });
    } else {
      try{ modalRaw.select(); document.execCommand('copy'); }catch(ex){}
      ok();
    }
  });

  // ── Markdown → HTML renderer (subset: # h1, __bold__, * list) ──
  function markdownToHtml(md){
    var lines = md.split('\n');
    var html  = '';
    var inList = false;
    for(var i=0; i<lines.length; i++){
      var line = lines[i];
      if(line.startsWith('# ')){
        if(inList){ html += '</ul>'; inList = false; }
        html += '<p class="md-h1">' + esc(line.slice(2)) + '</p>';
      } else if(/^__.*__$/.test(line)){
        if(inList){ html += '</ul>'; inList = false; }
        html += '<span class="md-bold">' + esc(line.slice(2,-2)) + '</span>';
      } else if(line.startsWith('* ')){
        if(!inList){ html += '<ul class="md-list">'; inList = true; }
        html += '<li>' + esc(line.slice(2)) + '</li>';
      } else if(line.trim() === ''){
        if(inList){ html += '</ul>'; inList = false; }
      } else {
        if(inList){ html += '</ul>'; inList = false; }
        html += '<p style="font-size:.85rem;color:var(--tx2)">' + esc(line) + '</p>';
      }
    }
    if(inList) html += '</ul>';
    return html;
  }

  function handle(file){
    if(!file.name.toLowerCase().endsWith('.xlsx')){
      showErr('Hanya file <strong>.xlsx</strong> yang didukung!');return;
    }
    document.getElementById('up-sec').style.opacity='.5';
    document.getElementById('ld').style.display='block';
    document.getElementById('res').style.display='none';
    document.getElementById('pv-sec').style.display='none';
    document.getElementById('dp-sec').style.display='none';
    document.getElementById('er').innerHTML='';
    document.getElementById('pvBody').classList.add('closed');
    document.getElementById('pvArrow').classList.remove('open');
    var fd=new FormData();
    fd.append('file',file);
    fetch('/api/upload',{method:'POST',body:fd})
      .then(function(r){return r.json();})
      .then(function(d){
        document.getElementById('ld').style.display='none';
        document.getElementById('up-sec').style.opacity='1';
        if(d.error){showErr(d.error);return;}
        render(d);
      })
      .catch(function(e){
        document.getElementById('ld').style.display='none';
        document.getElementById('up-sec').style.opacity='1';
        showErr('Gagal konek ke server: '+e.message);
      });
  }

  function showErr(msg){
    document.getElementById('er').innerHTML='<div class="err-box">⚠️ '+msg+'</div>';
  }

  function esc(s){
    return String(s)
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;');
  }

  function fmt(n){return Number(n).toLocaleString('id-ID');}

  function render(d){
    var info=d.file_info;
    var sz=info.size_mb>=0.1?info.size_mb+' MB':info.size_kb+' KB';
    var infoItems=[
      {l:'Nama File',v:info.filename},
      {l:'Ukuran File',v:sz},
      {l:'Total Baris',v:fmt(info.rows_total)},
      {l:'Total Kolom',v:info.cols_total},
      {l:'Memori (df)',v:info.memory_mb+' MB'}
    ];
    document.getElementById('fi-grid').innerHTML=infoItems.map(function(i){
      return '<div class="igi"><div class="igl">'+i.l+'</div><div class="igv">'+esc(String(i.v))+'</div></div>';
    }).join('');

    document.getElementById('ccnt').textContent='('+info.cols_total+' kolom)';
    document.getElementById('ctags').innerHTML=info.columns.map(function(c){
      return '<span class="tag">'+esc(c)+'</span>';
    }).join('');

    document.getElementById('plog').innerHTML=d.log.map(function(s){
      var st=s.status||'success';
      var ic=st==='warning'?'!':'✓';
      var cls=st==='warning'?'wn':'ok';
      return '<div class="step"><div class="si '+cls+'">'+ic+'</div>'
        +'<div class="sb"><div class="st">'+esc(s.step)+'</div>'
        +'<div class="sd">'+esc(s.detail)+'</div></div></div>';
    }).join('');

    document.getElementById('srow').innerHTML=[
      {n:fmt(info.rows_total),l:'Baris Awal'},
      {n:fmt(d.rows_out),l:'Baris Hasil Filter'},
      {n:d.cols_out,l:'Kolom Digunakan'},
      {n:fmt(info.rows_total-d.rows_out),l:'Baris Dihapus'}
    ].map(function(s){
      return '<div class="stat"><div class="stn">'+s.n+'</div><div class="stl">'+s.l+'</div></div>';
    }).join('');

    var pv=d.preview;
    var th='<thead><tr>'+pv.columns.map(function(c){return '<th>'+esc(c)+'</th>';}).join('')+'</tr></thead>';
    var tb='<tbody>'+pv.rows.map(function(row){
      return '<tr>'+row.map(function(c){return '<td title="'+esc(c)+'">'+esc(c)+'</td>';}).join('')+'</tr>';
    }).join('')+'</tbody>';
    document.getElementById('ptbl').innerHTML=th+tb;

    document.getElementById('pvMeta').textContent=
      '('+fmt(pv.rows.length)+' baris · '+pv.columns.length+' kolom)';

    document.getElementById('res').style.display='block';
    document.getElementById('res').scrollIntoView({behavior:'smooth'});
    fetchDpCards();
    fetchPivot();
  }

  function fetchPivot(){
    fetch('/api/pivot')
      .then(function(r){return r.json();})
      .then(function(d){if(!d.error)renderPivot(d);})
      .catch(function(){});
  }

  function renderPivot(d){
    var sec=document.getElementById('pv-sec');
    var tbl=document.getElementById('pv-tbl');
    var sbs=d.sumber_cols;
    var sts=d.status_cols;

    // ── Header Row 1 ──
    var hdr1='<tr>'
      +'<th rowspan="2" style="vertical-align:bottom">'+esc(d.col_dp)+'</th>'
      +sbs.map(function(sb,i){
          var cls='grp-hdr'+(i<sbs.length-1?' sep-r':'');
          return '<th colspan="'+sts.length+'" class="'+cls+'">'+esc(sb)+'</th>';
        }).join('')
      +'<th rowspan="2" style="vertical-align:bottom">Grand Total</th>'
      +'</tr>';

    // ── Header Row 2 ──
    var hdr2='<tr>'
      +sbs.map(function(sb,i){
          return sts.map(function(st,j){
            var isLast=(j===sts.length-1);
            var cls='sub-hdr'+(isLast&&i<sbs.length-1?' sep-r':'');
            return '<th class="'+cls+'">'+esc(st)+'</th>';
          }).join('');
        }).join('')
      +'</tr>';

    var th='<thead>'+hdr1+hdr2+'</thead>';

    function cellVal(v){
      return (v===0||v===''||v===null)
        ? '<td class="muted">-</td>'
        : '<td class="num">'+Number(v).toLocaleString('id-ID')+'</td>';
    }
    function cellValSep(v){
      return (v===0||v===''||v===null)
        ? '<td class="muted sep-r">-</td>'
        : '<td class="num sep-r">'+Number(v).toLocaleString('id-ID')+'</td>';
    }

    var tb='<tbody>'+d.rows.map(function(row){
      var cells=sbs.map(function(sb,i){
        var grp=row[sb]||{};
        var isSepGrp=(i<sbs.length-1);
        return sts.map(function(st,j){
          var v=grp[st]||0;
          var isLastSt=(j===sts.length-1);
          return (isLastSt&&isSepGrp)?cellValSep(v):cellVal(v);
        }).join('');
      }).join('');
      var gt=row.grand_total;
      var dpKey=row.dp;
      var hasData=!!(dpLookup[dpKey]&&dpLookup[dpKey].has_content);
      var rowCls=hasData?' class="clickable"':'';
      var dataAttr=hasData?' data-dp="'+esc(dpKey)+'"':'';
      return '<tr'+rowCls+dataAttr+'><td>'+esc(row.dp)+'</td>'+cells
        +'<td class="num">'+Number(gt).toLocaleString('id-ID')+'</td></tr>';
    }).join('')+'</tbody>';

    // ── Footer ──
    var g=d.grand;
    var gf='<tfoot><tr><td>Grand Total</td>'
      +sbs.map(function(sb,i){
          var grp=g[sb]||{};
          var isSepGrp=(i<sbs.length-1);
          return sts.map(function(st,j){
            var v=grp[st]||0;
            var isLastSt=(j===sts.length-1);
            var sepCls=(isLastSt&&isSepGrp)?' sep-r':'';
            return (v===0)
              ?'<td class="muted'+sepCls+'">-</td>'
              :'<td class="num'+sepCls+'">'+Number(v).toLocaleString('id-ID')+'</td>';
          }).join('');
        }).join('')
      +'<td class="num">'+Number(g.grand_total).toLocaleString('id-ID')+'</td>'
      +'</tr></tfoot>';

    tbl.innerHTML=th+tb+gf;

    // ── Click handler on clickable rows ──
    tbl.addEventListener('click', function(e){
      var tr=e.target.closest('tr[data-dp]');
      if(!tr) return;
      var dp=tr.getAttribute('data-dp');
      var card=dpLookup[dp];
      if(card) openModal(card.dp, card.content);
    });

    sec.style.display='block';
  }

  function fetchDpCards(){
    fetch('/api/textboxdata')
      .then(function(r){return r.json();})
      .then(function(d){
        if(!d.error) buildLookup(d.cards);
      })
      .catch(function(){});
  }

  // Build lookup dict and re-render pivot rows if pivot already drawn
  function buildLookup(cards){
    dpLookup = {};
    if(!cards||!cards.length) return;
    cards.forEach(function(c){ dpLookup[c.dp] = c; });
    // Refresh clickable classes on pivot rows (pivot may render before lookup is ready)
    var rows = document.querySelectorAll('#pv-tbl tbody tr[data-dp]');
    rows.forEach(function(tr){
      var dp = tr.getAttribute('data-dp');
      if(dpLookup[dp] && dpLookup[dp].has_content){
        tr.classList.add('clickable');
      }
    });
    // Also patch rows that weren't marked yet (no data-dp means no content)
    var allRows = document.querySelectorAll('#pv-tbl tbody tr:not([data-dp])');
    // nothing to do for these — they have no content
  }

})();
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────────────────────
#  Flask App
# ─────────────────────────────────────────────────────────────────────────────

try:
    from flask import Flask, request, jsonify, Response
    from flask_cors import CORS
    import pandas as pd
except ImportError as e:
    print(f"[ERROR] Dependency missing: {e}")
    print("Install dulu: pip install flask flask-cors pandas openpyxl")
    raise

app = Flask(__name__)
CORS(app)

store = {
    'df_raw' : None,
    'df_proc': None,
    'info'   : {},
    'csv'    : None,
}


def find_col(df, name):
    """Cari kolom secara case/spasi/underscore-insensitive."""
    n = name.lower().replace('_', ' ').replace('-', ' ').strip()
    for c in df.columns:
        if c.lower().replace('_', ' ').replace('-', ' ').strip() == n:
            return c
    return None


@app.route('/')
def index():
    return HTML_TEMPLATE


@app.route('/api/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Tidak ada file yang diunggah'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'Hanya file .xlsx yang didukung!'}), 400

    try:
        import io
        raw = f.read()
        df  = pd.read_excel(io.BytesIO(raw), engine='openpyxl')
        store['df_raw'] = df

        info = {
            'filename'  : f.filename,
            'size_bytes': len(raw),
            'size_kb'   : round(len(raw) / 1024, 2),
            'size_mb'   : round(len(raw) / 1_048_576, 3),
            'rows_total': int(len(df)),
            'cols_total': int(len(df.columns)),
            'columns'   : list(df.columns),
            'memory_mb' : round(df.memory_usage(deep=True).sum() / 1_048_576, 3),
        }
        store['info'] = info
        log = []
        res = df.copy()

        # Step 1: Seleksi Kolom
        if TARGET_COLUMNS:
            found = [c for c in TARGET_COLUMNS if c in res.columns]
            miss  = [c for c in TARGET_COLUMNS if c not in res.columns]
            res   = res[found]
            detail = f'{len(found)} kolom dipilih'
            if miss:
                detail += f', {len(miss)} tidak ditemukan: {miss}'
            log.append({'step': 'Seleksi Kolom', 'detail': detail, 'status': 'success'})
        else:
            log.append({
                'step'  : 'Seleksi Kolom',
                'detail': f'TARGET_COLUMNS = None → semua {len(df.columns)} kolom digunakan',
                'status': 'success',
            })

        # Step 2: Filter
        fc = find_col(res, FILTER_COLUMN)
        rb = len(res)
        if fc:
            res = res[res[fc].isin(FILTER_VALUES)].copy()
            ra  = len(res)
            log.append({
                'step'  : f'Filter "{FILTER_COLUMN}"',
                'detail': f'{rb:,} → {ra:,} baris ({rb - ra:,} baris dihapus)',
                'status': 'success',
            })
        else:
            log.append({
                'step'  : f'Filter "{FILTER_COLUMN}"',
                'detail': 'Kolom tidak ditemukan dalam data — langkah dilewati',
                'status': 'warning',
            })

        # Step 3: Sorting
        sa = [find_col(res, s) for s in SORT_COLUMNS]
        sa = [c for c in sa if c]
        if sa:
            res = res.sort_values(sa).reset_index(drop=True)
            log.append({
                'step'  : 'Pengurutan',
                'detail': 'Diurutkan berdasarkan: ' + str(sa),
                'status': 'success',
            })
        else:
            log.append({
                'step'  : 'Pengurutan',
                'detail': 'Kolom sort tidak ditemukan — dilewati',
                'status': 'warning',
            })

        store['df_proc'] = res

        # Generate CSV
        import io as _io
        buf = _io.StringIO()
        res.to_csv(buf, index=False)
        store['csv'] = buf.getvalue()

        pv = res.fillna('')
        return jsonify({
            'file_info': info,
            'log'      : log,
            'rows_out' : int(len(res)),
            'cols_out' : int(len(res.columns)),
            'preview'  : {
                'columns': list(pv.columns),
                'rows'   : pv.astype(str).values.tolist(),
            },
        })

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/download')
def download():
    if not store['csv']:
        return jsonify({'error': 'Belum ada data yang diproses'}), 400
    return Response(
        store['csv'].encode('utf-8'),
        mimetype='text/csv; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename={OUTPUT_FILENAME}'},
    )


@app.route('/api/textboxdata')
def textboxdata():
    df = store['df_proc']
    if df is None:
        return jsonify({'error': 'Belum ada data'}), 400

    col_dp       = find_col(df, 'DP Terjadwal')
    col_sumber   = find_col(df, 'Sumber Order')
    col_status   = find_col(df, 'Status Orderan')
    col_awb      = find_col(df, 'Nomor AWB')
    col_sprinter = find_col(df, 'Sprinter')

    if not col_dp:
        return jsonify({'error': "Kolom 'DP Terjadwal' tidak ditemukan"}), 400

    VALID_SUMBER  = ['TTREVERSE', 'TOKOREVERSE']
    VALID_STATUS  = ['DISPATCH', 'PUSAT_DISPATCH']
    SUMBER_DISPLAY = {
        'TTREVERSE'  : 'TTREVERSE',
        'TOKOREVERSE': 'TOKO REVERSE',
    }

    cards = []
    for dp_val in df[col_dp].dropna().unique():
        df_dp = df[df[col_dp] == dp_val]
        lines = [f'# {dp_val}']
        has_content = False

        for sumber_val in VALID_SUMBER:
            if col_sumber is None:
                break
            df_s = df_dp[df_dp[col_sumber].astype(str).str.strip().str.upper() == sumber_val]
            if df_s.empty:
                continue

            sumber_lines = []
            sumber_has   = False

            for status_val in VALID_STATUS:
                if col_status is None:
                    break
                df_st = df_s[df_s[col_status].astype(str).str.strip().str.upper() == status_val]
                if df_st.empty:
                    continue

                sumber_lines.append(f'__{status_val}__')
                for _, row in df_st.iterrows():
                    awb = str(row[col_awb]).strip() if col_awb and col_awb in row.index else ''
                    if status_val == 'DISPATCH' and col_sprinter and col_sprinter in row.index:
                        sp = str(row[col_sprinter]).strip()
                        sumber_lines.append(f'* {awb} [{sp}]')
                    else:
                        sumber_lines.append(f'* {awb}')
                sumber_lines.append('')
                sumber_has = True

            if sumber_has:
                lines.append('')
                lines.append(f'# {SUMBER_DISPLAY[sumber_val]}')
                lines.extend(sumber_lines)
                has_content = True

        # Strip trailing empty lines
        while lines and lines[-1] == '':
            lines.pop()

        cards.append({
            'dp'         : str(dp_val),
            'content'    : '\n'.join(lines),
            'has_content': has_content,
        })

    return jsonify({'cards': cards})


@app.route('/api/pivot')
def pivot():
    df = store['df_proc']
    if df is None:
        return jsonify({'error': 'Belum ada data'}), 400

    col_dp     = find_col(df, 'DP Terjadwal')
    col_status = find_col(df, 'Status Orderan')
    col_sumber = find_col(df, 'Sumber Order')

    if not col_dp:
        return jsonify({'error': "Kolom 'DP Terjadwal' tidak ditemukan"}), 400
    if not col_status:
        return jsonify({'error': "Kolom 'Status Orderan' tidak ditemukan"}), 400

    SUMBER_COLS = ['TTREVERSE', 'TOKOREVERSE']
    STATUS_COLS = ['DISPATCH', 'PUSAT_DISPATCH']

    # Normalisasi kolom bantu
    df_work = df.copy()
    df_work['_dp'] = df_work[col_dp]
    df_work['_st'] = df_work[col_status].astype(str).str.strip().str.upper()
    if col_sumber:
        df_work['_sb'] = df_work[col_sumber].astype(str).str.strip().str.upper()
    else:
        df_work['_sb'] = ''

    df_filt = df_work[
        df_work['_st'].isin(STATUS_COLS) &
        df_work['_sb'].isin(SUMBER_COLS)
    ]

    # Build lookup: {dp: {sumber: {status: count}}}
    lookup = {}
    if not df_filt.empty:
        for (dp_val, sb_val, st_val), grp in df_filt.groupby(['_dp', '_sb', '_st']):
            lookup.setdefault(dp_val, {}).setdefault(sb_val, {})[st_val] = len(grp)

    all_dp = list(df[col_dp].dropna().unique())
    ordered = [v for v in FILTER_VALUES if v in all_dp] + \
              sorted([v for v in all_dp if v not in FILTER_VALUES])

    rows = []
    grand_sb = {sb: {st: 0 for st in STATUS_COLS} for sb in SUMBER_COLS}
    grand_total = 0

    for dp_val in ordered:
        row = {'dp': str(dp_val)}
        row_total = 0
        dp_data = lookup.get(dp_val, {})
        for sb in SUMBER_COLS:
            sb_data = dp_data.get(sb, {})
            row[sb] = {}
            for st in STATUS_COLS:
                v = int(sb_data.get(st, 0))
                row[sb][st] = v
                grand_sb[sb][st] += v
                row_total += v
        row['grand_total'] = row_total
        grand_total += row_total
        rows.append(row)

    grand = {sb: {st: grand_sb[sb][st] for st in STATUS_COLS} for sb in SUMBER_COLS}

    return jsonify({
        'col_dp'     : col_dp,
        'sumber_cols': SUMBER_COLS,
        'status_cols': STATUS_COLS,
        'rows'       : rows,
        'grand'      : {**grand, 'grand_total': grand_total},
    })


# ─────────────────────────────────────────────────────────────────────────────
#  Entry Point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    print(f"✅ Excel Explorer berjalan di http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=debug)