LASGUM V47 - CENTRAL DATABASE ARCHITECTURE

ARSITEKTUR:
- Google Sheets = Database Pusat / Single Source of Truth.
- Google Apps Script = API Database.
- LASGUM = admin server: setup, master data, akun, dan administrasi.
- GitHub Pages = aplikasi pengguna.
- Siswa/Admin/Guru membaca database pusat langsung; tidak perlu membuka LASGUM terlebih dahulu.

ALUR PERANGKAT BARU:
1. Buka aplikasi GitHub Pages.
2. Aplikasi otomatis GET /exec?action=bootstrap.
3. Data pusat dimuat ke cache localStorage.
4. Jika internet putus, aplikasi memakai cache terakhir.
5. Saat online kembali, aplikasi mengambil database pusat lagi.

KEAMANAN:
- Bootstrap GET hanya READ dan tidak mengirim passwordHash PENGGUNA.
- Operasi CREATE/UPDATE/DELETE tetap memakai token role (LASGUM atau Admin/Guru).
- Siswa hanya dapat menulis hasil ujian sesuai endpoint yang diizinkan server.
- Jangan menyimpan LASGUM_SERVER_KEY di repository GitHub.

DEPLOY GOOGLE APPS SCRIPT:
- Deploy sebagai Web app.
- Execute as: pemilik script.
- Who has access: siapa saja yang dapat mengakses Web App.
- Pastikan URL /exec pada index.html benar.
- Setelah mengubah .gs, buat deployment/version baru dan gunakan URL Web App yang aktif.

DEPLOY GITHUB:
1. Upload isi ZIP ke root repository.
2. Pastikan index.html berada langsung di root.
3. Settings -> Pages -> Deploy from branch -> main -> /(root).
4. Buka URL GitHub Pages pada perangkat lain.

PENTING:
- Tidak ada sinkronisasi baca melalui LASGUM.
- LASGUM tidak perlu dibuka pada perangkat siswa/guru hanya untuk membaca data.
- localStorage hanya CACHE, bukan database utama.
- Jika database pusat kosong, perangkat tidak akan membuat database pusat sendiri. Data harus diisi melalui LASGUM/admin sesuai hak akses.
