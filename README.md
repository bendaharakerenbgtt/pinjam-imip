# PinjamIMIP — Sistem Peminjaman Inventaris & Alat Organisasi

Aplikasi web modern, responsif, dan mobile-first untuk layanan peminjaman fasilitas/inventaris organisasi IMIP. Menggunakan nuansa tema yang selaras dengan sistem keuangan (Deep Obsidian & Emerald / Cream & Orange Cabinet).

---

## 🚀 Fitur Utama

### 1. Sisi Publik (Peminjam)
- **Katalog Inventaris Interaktif**: Pencarian real-time & filter kategori (Elektronik, Audio & Sound, Kelistrikan, Logistik).
- **Status Ketersediaan**: Badge otomatis (*Tersedia [X]* / *Stok Habis*) dan kondisi barang.
- **Formulir Peminjaman Lengkap**:
  - Validasi jadwal (tanggal mulai s.d. selesai).
  - Keperluan pinjam & divisi/organisasi.
  - **Upload Foto KTM / Identitas**: Kompresi otomatis di browser (cepat & hemat kuota) lalu disimpan ke Google Drive.
- **Lacak Pengajuan (Cek Status)**: Masukkan Nomor WhatsApp atau ID Pinjam (misal `PJM-...`) untuk melihat apakah disetujui, ditolak, atau sudah selesai.

### 2. Sisi Admin (Pengurus Logistik)
- **Login PIN Cepat**: Default PIN `123456`.
- **Tinjauan Permintaan Masuk**:
  - Lihat foto KTM pemohon via lightbox popup.
  - **Aksi 1-Klik**:
    - **Terima (ACC)**: Otomatis kurangi stok tersedia.
    - **Tolak**: Input alasan penolakan.
    - **Tandai Dikembalikan**: Otomatis kembalikan stok unit.
  - **Chat WhatsApp Otomatis**: Tombol langsung membuka WhatsApp pemohon dengan format pesan otomatis sesuai status.
- **Manajemen Barang (CRUD)**:
  - Tambah barang baru beserta upload foto barang ke Google Drive.
  - Edit data barang (nama, kategori, stok total, kondisi, ganti foto).
  - Hapus barang yang sudah tidak aktif.
- **Koneksi Google Apps Script**: Konfigurasi URL Web App langsung dari antarmuka web.

---

## 🛠️ Panduan Setup Database (Google Sheets & Drive)

1. Buka [Google Drive](https://drive.google.com) lalu buat Spreadsheet baru (beri nama misalnya: `DB Peminjaman Barang IMIP`).
2. Di menu atas Spreadsheet, klik **Extensions** (Ekstensi) > **Apps Script**.
3. Hapus kode default yang ada di editor Apps Script, lalu buka file `code.gs` di folder ini dan **copy seluruh kodenya** ke editor Apps Script.
4. Di bagian atas editor, pada dropdown fungsi pilih **`setupSheetsDanDrive`**, lalu klik tombol **Run** (Jalankan).
   - Akan muncul pop-up izin akses Google Drive dan Spreadsheet -> klik **Review Permissions** -> pilih akun Anda -> klik **Advanced** -> klik **Go to ... (unsafe)** -> klik **Allow**.
   - Otomatis tab `barang`, `peminjaman`, dan folder Google Drive `PinjamIMIP - Foto & Bukti` akan terbuat!
5. Klik tombol biru **Deploy** di pojok kanan atas > **New deployment**:
   - Pilih jenis: **Web app** (ikon bola dunia).
   - Description: `PinjamIMIP Backend v1`.
   - Execute as: **Me (email Anda)**.
   - Who has access: **Anyone (Siapa saja)** *(PENTING agar web publik bisa kirim pengajuan)*.
   - Klik **Deploy** dan copy **Web App URL** yang berakhiran `/exec`.
6. Buka web `index.html`, masuk ke menu **Admin** (PIN `123456`) > klik **Set Google Script** > tempel URL tersebut lalu simpan. Selesai!

---

## 🌐 Panduan Deploy ke Hosting (Netlify / Vercel)

Karena dipisahkan dari web keuangan:
1. Buat repositori baru di GitHub (misal: `pinjam-imip`).
2. Push seluruh file di folder `c:\Users\iqi\Documents\Pinjamip` ke repo tersebut.
3. Buka dashboard **Netlify** atau **Vercel** Anda (akun yang sama).
4. Klik **Add new site** / **Import Project** > pilih repositori `pinjam-imip`.
5. Deploy langsung! File konfigurasi `netlify.toml` dan `vercel.json` sudah disediakan.
