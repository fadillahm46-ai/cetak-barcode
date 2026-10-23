Barcode Link Generator (HD)

Aplikasi web pembuat QR Barcode & Barcode Garis (1D) resolusi Ultra HD dengan fitur riwayat penyimpanan.

Struktur File untuk GitHub & Vercel

Pastikan file berikut berada dalam satu folder repositori GitHub Anda:

index.html — Aplikasi utama (siap jalan di Vercel maupun Google Apps Script)

vercel.json — Konfigurasi routing otomatis untuk Vercel

Code.gs — (Opsional) Backend Google Apps Script jika ingin menyimpan data ke Google Sheets

README.md — Dokumentasi proyek

Cara Upload ke GitHub & Deploy ke Vercel

Buat Repositori GitHub Baru

Buka github.com/new dan buat repositori baru.

Upload file index.html, vercel.json, dan README.md ke repositori tersebut, lalu klik Commit changes.

Deploy ke Vercel

Buka vercel.com/new dan login menggunakan akun GitHub Anda.

Pilih repositori yang baru saja dibuat, lalu klik Import.

Biarkan pengaturan Framework Preset pada Other, lalu klik Deploy.

Dalam beberapa detik, aplikasi Anda langsung aktif dan memiliki link publik Vercel.

Penyimpanan Data di Vercel

Otomatis (Default): Riwayat tersimpan langsung di penyimpanan browser (localStorage) dan dilengkapi tombol Ekspor / Impor JSON untuk memindahkan data antar perangkat.

Integrasi Google Sheets (Opsional): Jika ingin data di Vercel tetap tersimpan ke Google Sheets, deploy Code.gs sebagai Web App (akses: Anyone), lalu tempel URL Web App tersebut pada variabel GAS_API_URL di dalam index.html.
