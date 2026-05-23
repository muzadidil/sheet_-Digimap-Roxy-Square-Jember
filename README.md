# Digimap Roxy Square Jember — Google Sheets Edition

Versi Google Apps Script dari [Digimap Roxy Square Jember](https://github.com/editormpi/Digimap-Roxy-Square-Jember), hanya **2 file**:

- `Code.gs` — backend (REST API + Sheets sebagai database)
- `Index.html` — frontend (UI iOS-style glassmorphism, semua dalam 1 file)

## Cara Setup

1. Buka Google Sheet baru
2. **Extensions → Apps Script**
3. Hapus `Code.gs` default, paste isi `Code.gs` dari repo ini
4. Klik **+ → HTML** → namai `Index` → paste isi `Index.html`
5. **Deploy → New deployment → Type: Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (atau Anyone with the link)
6. Otorisasi, lalu copy URL web app yang muncul
7. Buka URL tersebut — sheet `Members`, `Sales`, `Settings` akan otomatis dibuat

## Login Default

- Username: `kiki`
- PIN: `1234`

(super user — bisa kelola anggota, theme, export, dll.)
