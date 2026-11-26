PAKET UJIAN REFLEKSI XI RPL - LOGIN + 50 SOAL + TIMER 30 MENIT

FILE WEBSITE (upload ke root GitHub repository):
- index.html
- style.css
- script.js

FILE BACKEND (tempel ke editor Google Apps Script):
- Code.gs

FITUR:
- Login siswa menggunakan NIS/NISN dan password yang dibuat guru.
- 50 soal pilihan ganda: 15 Word, 15 Excel, 20 Basis Data.
- Timer 30 menit.
- Navigasi soal, penilaian otomatis, hasil, pembahasan.
- Skor dihitung di server Apps Script; kunci jawaban tidak disertakan di file JS publik.
- Rekap tersimpan ke tab RESULTS dan detail jawaban ke ANSWERS.

SETUP GOOGLE SHEETS:
1. Buat Google Spreadsheet.
2. Salin ID spreadsheet dari URL (antara /d/ dan /edit).
3. Buka https://script.google.com dan buat project.
4. Tempel Code.gs, ganti SPREADSHEET_ID.
5. Save, pilih fungsi setupSheets lalu Run dan beri izin.
6. Buat akun siswa dari editor Apps Script dengan fungsi createStudentAccount.
   Contoh pemanggilan:
   createStudentAccount("XI001","Nama Siswa","XI RPL",1,"password123");
   Ganti nilai contoh sesuai siswa. Jalankan sekali untuk setiap siswa.
7. Deploy > New deployment > Web app.
8. Execute as: Me.
9. Who has access: Anyone (jika tersedia dan sesuai kebijakan akun).
10. Deploy, salin URL /exec.

HUBUNGKAN WEBSITE:
1. Buka script.js.
2. Ganti TEMPEL_URL_WEB_APP_GOOGLE_APPS_SCRIPT_DI_SINI pada konstanta API_URL dengan URL /exec.
3. Upload index.html, style.css, script.js ke root repository GitHub.
4. Settings > Pages > Deploy from a branch > main > /(root) > Save.

CATATAN:
- Untuk integrasi GitHub Pages lintas domain, browser harus mengizinkan respons Apps Script dibaca oleh fetch. Jika login menampilkan error CORS, Google Apps Script ContentService tidak dapat dipakai langsung dari origin tersebut pada konfigurasi itu. Solusi paling stabil adalah menyajikan frontend melalui Apps Script HTML Service atau memakai backend/proxy HTTPS yang mendukung CORS. Jangan mengubah fetch menjadi no-cors untuk login karena respons autentikasi tidak akan dapat dibaca.
- Password disimpan sebagai hash SHA-256 dengan salt, bukan teks biasa. Untuk sistem produksi berisiko tinggi, gunakan penyimpanan autentikasi yang lebih kuat.
- Sesi menggunakan CacheService; siswa dapat login kembali jika sesi berakhir.
- Versi ini membatasi satu kali pengumpulan hasil per NIS untuk ujian ini.
