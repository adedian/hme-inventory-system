# HME Inventory Management System

Sebuah sistem manajemen inventaris dan peminjaman alat berbasis *serverless* (tanpa server) yang dibangun untuk mengoptimalkan alur operasional PT Hexa Multi Energi. Aplikasi ini memanfaatkan **Google Apps Script (GAS)** sebagai *backend* dan **Google Sheets** sebagai *database* terpusat, memberikan solusi yang *real-time*, cepat, dan efisien.

## Fitur Unggulan

- **Role-Based Access Control (RBAC):** Otentikasi pengguna dengan pemisahan hak akses yang tegas antara **Admin** (manajemen master data, persetujuan transaksi, dan rekapitulasi) dan **Staff** (pengajuan peminjaman barang).
- **Real-Time Auto-Sync:** Sinkronisasi status inventaris secara dinamis. Barang yang sedang dipinjam akan otomatis terhapus dari opsi peminjaman *staff* untuk mencegah *double-booking*, dan kembali muncul setelah dikonfirmasi "Kembali".
- **Automated Email Notifications:** Terintegrasi langsung dengan ekosistem Google (Gmail API) untuk mengirimkan notifikasi *email* otomatis kepada administrator setiap kali terdapat aktivitas pengajuan atau pengembalian alat.
- **Analytics & Excel Export:** *Dashboard* visual untuk memantau statistik peminjaman bulanan, dilengkapi dengan fitur *generate* laporan transaksi komprehensif yang dapat diunduh dalam format **Excel (.xlsx)**.
- **Modern & Responsive UI:** Antarmuka pengguna (UI) yang bersih dan intuitif, dirancang menggunakan HTML/CSS *native* tanpa *library* eksternal yang berat, memastikan waktu muat (*load time*) yang sangat cepat di berbagai perangkat.

## Tech Stack

- **Frontend:** HTML5, CSS3, Vanilla JavaScript (DOM Manipulation)
- **Backend:** Google Apps Script (ES6)
- **Database:** Google Sheets API
- **Integrations:** Gmail API (MailApp)

## Panduan Instalasi (Setup Guide)

Bagi Anda yang ingin menjalankan aplikasi ini di lingkungan *Google Workspace* Anda sendiri:

1. Buat **Google Sheets** baru.
2. Buka menu **Ekstensi > Apps Script**.
3. Salin dan tempel kode dari `Code.gs` ke editor sisi *server*.
4. Buat file HTML baru bernama `index.html` dan tempelkan kode *frontend*.
5. Sesuaikan variabel berikut di `Code.gs`:
   - Ganti ID Spreadsheet pada `SpreadsheetApp.openById("ID_ANDA_DI_SINI");`
   - Ganti email Admin pada variabel `ADMIN_EMAIL`.
6. Lakukan *deployment* dengan memilih:
   - **Deploy > New Deployment**
   - Pilih jenis **Web App**
   - Eksekusi sebagai: **Me** (Penting agar fitur *email* berjalan dengan otorisasi Anda)
   - Akses: **Anyone**

## Author

**Ade Dian Sukmana, S.Kom** *IT & Web Developer* | PT Hexa Multi Energi
