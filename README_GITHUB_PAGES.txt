LASGUM V45 - OFFLINE FULL DISTRIBUTION

Isi utama:
- index.html
- lasgum-sw.js
- lasgum.webmanifest
- .nojekyll
- LASGUM_DATABASE_V3.2_SECURE_FIXED_V5.1_FINAL.gs

GitHub Pages:
1. Upload ISI folder ini ke ROOT repository (jangan membuat subfolder).
2. Pastikan index.html berada langsung di root.
3. Settings -> Pages -> Deploy from a branch -> main -> /(root).
4. Buka URL GitHub Pages.

Offline:
- Setelah pertama kali dibuka secara online, Service Worker V45 akan menyimpan shell aplikasi.
- V45 juga mencoba menyimpan Tailwind dan XLSX dari CDN ke cache runtime.
- Kunjungan berikutnya dapat berjalan tanpa koneksi selama browser sudah pernah memuat resource tersebut.
- Sinkronisasi database online tetap membutuhkan internet.

Catatan:
- V45 belum menghilangkan URL CDN dari HTML; pendekatan ini mempertahankan kompatibilitas V44 dan menambah caching offline.
- Untuk benar-benar tanpa ketergantungan CDN sejak kunjungan pertama, library Tailwind/XLSX perlu divendor sebagai file lokal.
- Jangan menyimpan LASGUM_SERVER_KEY di repository publik.
