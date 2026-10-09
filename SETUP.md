# Setup Cloud: Google Sheets + Drive + GitHub Pages

Sekali saja (±10 menit). Setelah ini app jalan penuh di HP/laptop mana pun.

## 0. Tab User (akun login)

Di Sheet ini, tab **`User`** berisi akun dengan header persis:

`NIK | Name | Name in File | password | role | alias`

- **NIK / password**: untuk halaman login (diverifikasi di server,
  password tak pernah dikirim ke HP lain).
- **Name in File**: nama persis/varian yang terdeteksi di Excel.
  Ejaan beda dikit otomatis cocok (`A. Wahid A.W.` = `Abdul Wahid A.W.`,
  `Satria W.K.` = `Satria Wijaya K.`). Varian yang jauh (`Zaidhiya R.`)
  tulis di kolom **alias**, dipisah `;`.
- **role**: isi `admin` untuk akun admin (melihat semua file & boleh
  mengisi TTD siapa pun). Kosongkan = user biasa (hanya file & TTD miliknya).

## 1. Pasang backend (Apps Script)

1. Buka Google Sheet ini:
   <https://docs.google.com/spreadsheets/d/1Dex3iv4toDo1gfEs2r5-0LrZyeRp3omDdcKaNZdNmYk/edit>
2. Menu **Extensions → Apps Script** (hapus kode bawaan bila ada).
3. Copy SELURUH isi file `appsscript/Code.gs` dari repo ini → tempel → **Save**
   (ID Sheet & folder Drive sudah terisi otomatis dari link kamu).
4. **Deploy → Manage deployments → ✏️ → New version** (WAJIB tiap update
   Code.gs, kalau tidak endpoint baru tak aktif) — atau New deployment
   bila pertama kali — tipe **Web app**:
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
   - **Deploy** → izinkan akses saat diminta → copy **Web app URL**
     (bentuk `https://script.google.com/macros/s/…/exec`).

## 2. Hubungkan aplikasi

1. Buka aplikasinya (lihat bagian 3) → halaman **login** → masuk dengan
   NIK & password dari tab User.
2. Saat login, app otomatis: tarik user + TTD dari cloud, lalu unduh
   file Excel dari tiap folder bulan di Drive (**admin**: semua file,
   **user**: hanya file yang memuat namanya).
3. Panel kiri → kartu **Cloud** → tempel Web app URL → **Simpan & Tes**.
   - Dot hijau = tersambung. Tanpa URL pun app tetap jalan lokal.
4. **⬇ Ambil TTD** = tarik database dari Sheets. **⬆ Kirim semua** =
   dorong TTD lokal ke Sheets. Bila *auto-sync* aktif, tiap TTD baru
   terkirim sendiri (yang lebih baru menang bila konflik).
5. Kartu **Drive**: pilih bulan → **Muat daftar** = lihat file Excel di
   folder Drive, klik ⬇ untuk memuat ke app; **Upload file ini** = simpan
   file asli; **Simpan hasil signed** = simpan file bertanda tangan
   (masuk folder bulan yang sama).

Struktur yang dibuat otomatis: tab `TTD` di Sheet
(`key | name | fileId | updatedAt`), subfolder `TTD` di Drive berisi
PNG tiap tanda tangan, file laporan di folder utama & subfolder bulan:
<https://drive.google.com/drive/folders/1TOQ8S1_QaaMBxfOi2xMYYI4E4wtI5Ctw>

## 3. Hosting GitHub Pages

Otomatis via workflow `.github/workflows/deploy.yml` — tiap push ke
`main` langsung build & deploy. Saat deploy pertama, GitHub membuat
sitenya sendiri; bukanya di:

**https://hitographic.github.io/TTDBulanan/**

(cek progres di tab **Actions** repo bila baru pertama kali).
