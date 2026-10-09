# Setup Cloud: Google Sheets + Drive + GitHub Pages

Sekali saja (±10 menit). Setelah ini app jalan penuh di HP/laptop mana pun.

## 1. Pasang backend (Apps Script)

1. Buka Google Sheet ini:
   <https://docs.google.com/spreadsheets/d/1Dex3iv4toDo1gfEs2r5-0LrZyeRp3omDdcKaNZdNmYk/edit>
2. Menu **Extensions → Apps Script** (hapus kode bawaan bila ada).
3. Copy SELURUH isi file `appsscript/Code.gs` dari repo ini → tempel → **Save**
   (ID Sheet & folder Drive sudah terisi otomatis dari link kamu).
4. **Deploy → New deployment** → tipe **Web app**:
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
   - **Deploy** → izinkan akses saat diminta → copy **Web app URL**
     (bentuk `https://script.google.com/macros/s/…/exec`).

## 2. Hubungkan aplikasi

1. Buka aplikasinya (lihat bagian 3).
2. Panel kiri → kartu **Cloud** → tempel Web app URL → **Simpan & Tes**.
   - Dot hijau = tersambung. Tanpa URL pun app tetap jalan lokal.
3. **⬇ Ambil TTD** = tarik database dari Sheets. **⬆ Kirim semua** =
   dorong TTD lokal ke Sheets. Bila *auto-sync* aktif, tiap TTD baru
   terkirim sendiri (yang lebih baru menang bila konflik).
4. Kartu **Drive**: **Muat daftar** = lihat file Excel di folder Drive,
   klik ⬇ untuk memuat ke app; **Upload file ini** = simpan file asli;
   **Simpan hasil signed** = simpan file bertanda tangan.

Struktur yang dibuat otomatis: tab `TTD` di Sheet
(`key | name | fileId | updatedAt`), subfolder `TTD` di Drive berisi
PNG tiap tanda tangan, file laporan di folder utama:
<https://drive.google.com/drive/folders/1TOQ8S1_QaaMBxfOi2xMYYI4E4wtI5Ctw>

## 3. Hosting GitHub Pages

Otomatis via workflow `.github/workflows/deploy.yml` — tiap push ke
`main` langsung build & deploy. Saat deploy pertama, GitHub membuat
sitenya sendiri; bukanya di:

**https://hitographic.github.io/TTDBulanan/**

(cek progres di tab **Actions** repo bila baru pertama kali).
