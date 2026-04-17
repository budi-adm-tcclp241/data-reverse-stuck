# 📊 Excel Explorer

Upload file `.xlsx` → Filter & Sort otomatis → Export ke CSV.

## 🚀 Deploy ke Railway

1. Push ke GitHub repository
2. Buka [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
3. Pilih repository ini → Railway otomatis detect `Procfile` dan deploy
4. Buka URL yang diberikan Railway

## 💻 Jalankan Lokal

```bash
pip install -r requirements.txt
python app.py
# Buka http://localhost:5000
```

## ⚙️ Konfigurasi

Edit variabel di bagian atas `app.py`:

| Variabel | Fungsi |
|---|---|
| `TARGET_COLUMNS` | Kolom yang ditampilkan (`None` = semua) |
| `FILTER_COLUMN` | Nama kolom yang difilter |
| `FILTER_VALUES` | Daftar nilai yang lolos filter |
| `SORT_COLUMNS` | Urutan kolom untuk sorting |
| `OUTPUT_FILENAME` | Nama file CSV hasil download |
| `TEXT_DISPATCH_TOP/BOTTOM` | Teks kustom sebelum/sesudah AWB DISPATCH |
| `TEXT_PUSAT_TOP/BOTTOM` | Teks kustom sebelum/sesudah AWB PUSAT_DISPATCH |

## 📁 Struktur File

```
app.py            ← Semua kode (backend + frontend)
requirements.txt  ← Python dependencies
Procfile          ← Konfigurasi Railway/Heroku
README.md         ← Dokumentasi ini
```
