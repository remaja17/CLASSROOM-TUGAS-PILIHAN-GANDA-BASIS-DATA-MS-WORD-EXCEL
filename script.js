/* ==================================================
   REFLEKSI XI RPL - 50 SOAL - LOGIN - TIMER 30 MENIT
   GANTI URL API DI BAWAH DENGAN URL WEB APP /exec
================================================== */
const API_URL = "https://script.google.com/macros/s/AKfycbzgE9U178btpx9cHgz3y0unFpetzajMdHZJ0XsWapStesEMEVFQoUVlzCWTcwn5zhQg/exec";
const QUESTIONS = [{"number":1,"section":"WORD","level":"Medium","question":"Fitur Styles pada Microsoft Word paling tepat digunakan untuk ...","options":["Mengatur format judul dan paragraf secara konsisten","Menghitung nilai numerik dalam dokumen","Menghapus seluruh halaman kosong","Mengubah dokumen menjadi database"],"explanation":"Styles menjaga konsistensi format dan membantu pembuatan daftar isi otomatis."},{"number":2,"section":"WORD","level":"Medium","question":"Sebelum membuat daftar isi otomatis, hal terpenting yang harus dilakukan adalah ...","options":["Memberi warna pada semua judul","Menerapkan Heading Styles pada judul dan subjudul","Mengubah margin menjadi Narrow","Mengetik daftar isi secara manual"],"explanation":"Word mengambil struktur daftar isi dari Heading Styles."},{"number":3,"section":"WORD","level":"Medium","question":"Track Changes berguna ketika ...","options":["Melacak perubahan saat dokumen ditinjau bersama","Mengurutkan data tabel","Menghitung rata-rata nilai","Membuat diagram relasi"],"explanation":"Track Changes merekam penambahan, penghapusan, dan perubahan dokumen."},{"number":4,"section":"WORD","level":"Medium","question":"Agar teks mengalir di sekeliling gambar, pengaturan yang digunakan adalah ...","options":["Text Wrapping","Word Count","Page Color","Mail Merge"],"explanation":"Text Wrapping mengatur hubungan gambar dengan aliran teks."},{"number":5,"section":"WORD","level":"Medium","question":"Section Break digunakan ketika dokumen membutuhkan ...","options":["Pengaturan layout berbeda pada bagian tertentu","Penghapusan semua komentar","Penggabungan dua file Excel","Pemeriksaan ejaan saja"],"explanation":"Section Break membagi dokumen menjadi bagian yang dapat memiliki format halaman berbeda."},{"number":6,"section":"WORD","level":"Medium","question":"Mail Merge paling tepat digunakan untuk ...","options":["Membuat surat massal dengan data penerima berbeda","Membuat animasi presentasi","Menghitung nilai rata-rata","Menggambar ERD"],"explanation":"Mail Merge menggabungkan template dokumen dengan sumber data penerima."},{"number":7,"section":"WORD","level":"Medium","question":"Sampul tidak bernomor, sedangkan BAB I dimulai dari halaman 1. Solusi yang tepat adalah ...","options":["Mengetik nomor secara manual","Menggunakan Section Break lalu mengatur nomor halaman pada section berikutnya","Menghapus sampul","Mengubah ukuran kertas"],"explanation":"Section terpisah memungkinkan penomoran dimulai ulang pada bagian BAB I."},{"number":8,"section":"WORD","level":"Hard","question":"Judul tidak muncul di Navigation Pane. Penyebab paling mungkin adalah ...","options":["Judul belum diberi Heading Style","Warna font terlalu gelap","Dokumen terlalu pendek","Ukuran kertas A4"],"explanation":"Navigation Pane membaca struktur heading, bukan sekadar teks besar atau tebal."},{"number":9,"section":"WORD","level":"Hard","question":"Cara terbaik agar seluruh subjudul memiliki format seragam adalah ...","options":["Mengatur satu per satu","Memodifikasi style Heading yang digunakan","Menambahkan spasi manual","Mengubah subjudul menjadi gambar"],"explanation":"Perubahan style diterapkan konsisten pada semua teks yang menggunakan style tersebut."},{"number":10,"section":"WORD","level":"Hard","question":"Perbedaan Page Break dan Section Break adalah ...","options":["Keduanya selalu sama","Page Break memulai halaman baru, Section Break membagi bagian dengan pengaturan format yang dapat berbeda","Section Break hanya untuk tabel","Page Break menghapus halaman"],"explanation":"Section Break membuat section baru; Page Break hanya memulai halaman baru."},{"number":11,"section":"WORD","level":"Hard","question":"Agar baris judul tabel muncul kembali pada setiap halaman lanjutan, gunakan ...","options":["Repeat Header Rows","Split Cells","AutoCorrect","Format Painter"],"explanation":"Repeat Header Rows mengulang baris header tabel pada halaman berikutnya."},{"number":12,"section":"WORD","level":"Hard","question":"Mail Merge menghasilkan nama penerima yang sama pada semua surat. Hal pertama yang perlu diperiksa adalah ...","options":["Sumber data dan pemetaan field merge","Warna halaman","Ukuran ikon Word","Pengaturan speaker"],"explanation":"Sumber data atau field yang salah dapat menyebabkan data penerima tidak berubah."},{"number":13,"section":"WORD","level":"Expert","question":"Header berbeda untuk halaman ganjil dan genap serta halaman pertama tanpa header. Pengaturan yang diperlukan adalah ...","options":["Different First Page dan Different Odd & Even Pages","Landscape dan Columns","Track Changes dan Comments","Protect Document saja"],"explanation":"Kedua opsi tersebut mengatur halaman pertama dan pola ganjil-genap."},{"number":14,"section":"WORD","level":"Expert","question":"Untuk menjaga tampilan dokumen saat dibuka di komputer lain, tindakan paling tepat adalah ...","options":["Mengandalkan spasi manual","Menggunakan Styles dan membagikan salinan PDF","Mengubah semua teks menjadi gambar","Menghapus seluruh section"],"explanation":"Styles menjaga struktur dokumen dan PDF mempertahankan tata letak."},{"number":15,"section":"WORD","level":"Expert","question":"Daftar isi otomatis tidak memperbarui nomor halaman setelah isi dokumen berubah. Tindakan yang benar adalah ...","options":["Menghapus semua heading","Memilih Update Table lalu Update entire table","Mengetik ulang semua nomor halaman","Mengubah dokumen menjadi TXT"],"explanation":"Update entire table memperbarui teks heading dan nomor halaman."},{"number":16,"section":"EXCEL","level":"Medium","question":"Rumus untuk menjumlahkan nilai pada sel B2 sampai B11 adalah ...","options":["=ADD(B2:B11)","=SUM(B2:B11)","=TOTAL(B2-B11)","=COUNT(B2:B11)"],"explanation":"SUM menjumlahkan nilai dalam rentang sel."},{"number":17,"section":"EXCEL","level":"Medium","question":"Fungsi AVERAGE digunakan untuk ...","options":["Menghitung nilai rata-rata","Menghitung jumlah sel berisi teks","Mencari nilai terbesar","Menggabungkan teks"],"explanation":"AVERAGE menghasilkan rata-rata aritmetika dari data numerik."},{"number":18,"section":"EXCEL","level":"Medium","question":"Jika nilai di C2 minimal 75 maka tampil LULUS, selain itu REMEDIAL. Rumus yang tepat adalah ...","options":["=IF(C2>=75,\"LULUS\",\"REMEDIAL\")","=SUM(C2,75)","=COUNTIF(C2,75)","=MAX(C2,\"LULUS\")"],"explanation":"IF menguji kondisi logika dan mengembalikan hasil sesuai kondisi benar atau salah."},{"number":19,"section":"EXCEL","level":"Medium","question":"Fungsi COUNT digunakan untuk menghitung ...","options":["Semua sel tidak kosong","Sel yang berisi angka","Sel yang berisi teks saja","Jumlah karakter"],"explanation":"COUNT menghitung banyaknya sel yang berisi angka."},{"number":20,"section":"EXCEL","level":"Medium","question":"Untuk mencari nilai tertinggi dari D2 sampai D31 digunakan ...","options":["=MIN(D2:D31)","=HIGH(D2:D31)","=MAX(D2:D31)","=TOP(D2:D31)"],"explanation":"MAX mengembalikan nilai terbesar pada rentang."},{"number":21,"section":"EXCEL","level":"Medium","question":"Referensi absolut pada Excel ditulis dengan tanda ...","options":["#","$","%","&"],"explanation":"Tanda dolar mengunci kolom dan/atau baris, misalnya $A$1."},{"number":22,"section":"EXCEL","level":"Medium","question":"Untuk menghitung jumlah siswa dengan nilai minimal 75 pada E2:E31, gunakan ...","options":["=COUNTIF(E2:E31,\">=75\")","=SUM(E2:E31,75)","=AVERAGE(E2:E31)","=IF(E2:E31>=75)"],"explanation":"COUNTIF menghitung sel yang memenuhi satu kriteria."},{"number":23,"section":"EXCEL","level":"Hard","question":"Rumus di F2 adalah =D2*$B$1 lalu disalin ke F3. Bagian $B$1 tetap ...","options":["Berubah menjadi B2","Tetap mengacu ke B1","Berubah menjadi C1","Terhapus"],"explanation":"Referensi absolut $B$1 tidak berubah saat rumus disalin."},{"number":24,"section":"EXCEL","level":"Hard","question":"Untuk menjumlahkan penjualan pada wilayah \"Timur\" saja dari B2:B100 berdasarkan wilayah A2:A100, gunakan ...","options":["SUMIF","COUNT","LEFT","ROUND"],"explanation":"SUMIF menjumlahkan rentang berdasarkan kriteria pada rentang lain."},{"number":25,"section":"EXCEL","level":"Hard","question":"Apa hasil logika dari =AND(A1>=70,B1=\"Hadir\") jika A1=80 dan B1=Hadir?","options":["TRUE","FALSE","80","#VALUE!"],"explanation":"AND bernilai TRUE jika seluruh kondisi yang diuji bernilai benar."},{"number":26,"section":"EXCEL","level":"Hard","question":"Untuk mengambil empat karakter paling kiri dari kode pada A2, fungsi yang tepat adalah ...","options":["RIGHT(A2,4)","MID(A2,4)","LEFT(A2,4)","LEN(A2,4)"],"explanation":"LEFT mengambil sejumlah karakter dari sisi kiri teks."},{"number":27,"section":"EXCEL","level":"Hard","question":"Mengapa fitur Sort sebaiknya diterapkan pada seluruh tabel, bukan hanya satu kolom nilai?","options":["Agar baris data tetap berpasangan dengan identitas yang benar","Agar warna berubah otomatis","Agar file lebih kecil","Agar rumus dihapus"],"explanation":"Mengurutkan satu kolom saja dapat memisahkan nilai dari identitas siswa."},{"number":28,"section":"EXCEL","level":"Expert","question":"Rumus untuk menghitung rata-rata nilai hanya pada kelas XI RPL A dengan kelas di A2:A50 dan nilai di C2:C50 adalah ...","options":["=AVERAGEIF(A2:A50,\"XI RPL A\",C2:C50)","=COUNTIF(A2:A50,\"XI RPL A\")","=SUM(A2:A50,C2:C50)","=MAXIF(A2:A50,C2:C50)"],"explanation":"AVERAGEIF menghitung rata-rata pada rentang yang memenuhi kriteria."},{"number":29,"section":"EXCEL","level":"Expert","question":"Rumus =IF(AND(B2>=75,C2=\"LENGKAP\"),\"TUNTAS\",\"PERIKSA\") menghasilkan PERIKSA jika ...","options":["B2=80 dan C2=LENGKAP","B2=75 dan C2=LENGKAP","B2=90 dan C2=KURANG","B2=76 dan C2=LENGKAP"],"explanation":"AND mensyaratkan kedua kondisi benar; C2=KURANG membuat kondisi gabungan salah."},{"number":30,"section":"EXCEL","level":"Expert","question":"Ketika rumus VLOOKUP menghasilkan #N/A, pemeriksaan awal yang paling masuk akal adalah ...","options":["Memastikan nilai kunci ditemukan di kolom pertama tabel referensi dan tipe datanya sesuai","Mengubah semua angka menjadi warna merah","Menghapus worksheet","Menambah tinggi baris"],"explanation":"VLOOKUP mencari nilai pada kolom pertama tabel; perbedaan tipe data atau spasi dapat menyebabkan tidak ditemukan."},{"number":31,"section":"DATABASE","level":"Medium","question":"Sebuah tabel SISWA memiliki kolom NIS, Nama, Kelas. Kolom yang paling tepat dijadikan primary key adalah ...","options":["Nama","Kelas","NIS","Nama dan Kelas secara acak"],"explanation":"NIS dapat menjadi identitas unik untuk setiap siswa."},{"number":32,"section":"DATABASE","level":"Medium","question":"Mengapa primary key tidak boleh bernilai sama pada dua baris berbeda?","options":["Agar data dapat diidentifikasi secara unik","Agar tabel lebih berwarna","Agar query lebih panjang","Agar semua kolom wajib teks"],"explanation":"Primary key harus unik sehingga setiap record dapat dibedakan."},{"number":33,"section":"DATABASE","level":"Medium","question":"Tabel PEMINJAMAN memiliki kolom id_pinjam, tanggal, nis_siswa. Kolom nis_siswa sebaiknya menjadi ...","options":["Foreign key yang mengacu ke SISWA.NIS","Primary key pada tabel SISWA sekaligus","Kolom hasil perhitungan","Index gambar"],"explanation":"Foreign key menghubungkan record peminjaman dengan siswa pemilik transaksi."},{"number":34,"section":"DATABASE","level":"Medium","question":"Perintah SQL untuk menampilkan semua kolom dari tabel SISWA adalah ...","options":["SHOW ALL SISWA;","SELECT * FROM SISWA;","GET SISWA ALL;","OPEN TABLE SISWA;"],"explanation":"SELECT * FROM mengambil seluruh kolom dari tabel yang disebutkan."},{"number":35,"section":"DATABASE","level":"Medium","question":"Jika ingin menampilkan siswa dari kelas XI RPL A, klausa SQL yang digunakan adalah ...","options":["ORDER BY","GROUP BY","WHERE","INSERT"],"explanation":"WHERE menyaring baris berdasarkan kondisi."},{"number":36,"section":"DATABASE","level":"Medium","question":"Apa tujuan relasi one-to-many antara tabel KELAS dan SISWA?","options":["Satu kelas dapat memiliki banyak siswa, setiap siswa terkait pada satu kelas","Satu siswa wajib memiliki banyak kelas","Setiap kelas hanya boleh satu kolom","Semua siswa harus memiliki NIS sama"],"explanation":"Relasi one-to-many menghubungkan satu record kelas dengan banyak record siswa."},{"number":37,"section":"DATABASE","level":"Medium","question":"Jika kolom tanggal_lahir menyimpan tanggal, tipe data yang paling sesuai adalah ...","options":["DATE","BOOLEAN","BLOB","VARCHAR(2)"],"explanation":"DATE dirancang untuk menyimpan nilai tanggal."},{"number":38,"section":"DATABASE","level":"Medium","question":"Perintah untuk menambahkan satu record baru ke tabel SISWA adalah ...","options":["UPDATE","INSERT INTO","DELETE","SELECT"],"explanation":"INSERT INTO menambahkan baris baru ke tabel."},{"number":39,"section":"DATABASE","level":"Hard","question":"Tabel NILAI menyimpan nis_siswa, nama_siswa, id_mapel, nama_mapel, nilai. Nama siswa berulang pada banyak baris. Desain lebih baik adalah ...","options":["Pisahkan SISWA dan MAPEL, lalu hubungkan melalui NILAI","Simpan semua data dalam satu kolom","Hapus NIS","Jadikan nilai sebagai primary key"],"explanation":"Pemisahan entitas mengurangi redundansi dan anomali pembaruan."},{"number":40,"section":"DATABASE","level":"Hard","question":"Dalam tabel DETAIL_PESANAN, satu pesanan dapat memiliki banyak produk. Kunci yang wajar untuk tiap detail adalah ...","options":["Nomor urut baris saja tanpa konteks","Kombinasi id_pesanan dan id_produk atau id_detail tersendiri","Nama pelanggan saja","Harga produk saja"],"explanation":"Kunci gabungan atau id detail unik dapat mengidentifikasi setiap baris detail."},{"number":41,"section":"DATABASE","level":"Hard","question":"Apa fungsi FOREIGN KEY?","options":["Menjaga keterkaitan dan integritas referensial antar tabel","Mengurutkan hasil otomatis","Mengenkripsi seluruh database","Mengubah tipe data"],"explanation":"Foreign key memastikan nilai relasi merujuk pada record yang valid di tabel induk."},{"number":42,"section":"DATABASE","level":"Hard","question":"Query untuk mengubah nilai kolom kelas siswa NIS '24011' menjadi 'XI RPL B' adalah ...","options":["UPDATE SISWA SET kelas='XI RPL B' WHERE NIS='24011';","ALTER SISWA kelas='XI RPL B';","INSERT SISWA kelas='XI RPL B';","SELECT kelas='XI RPL B' FROM SISWA;"],"explanation":"UPDATE mengubah data, sedangkan WHERE membatasi baris yang diperbarui."},{"number":43,"section":"DATABASE","level":"Hard","question":"Mengapa perintah UPDATE tanpa WHERE perlu sangat diwaspadai?","options":["Dapat mengubah seluruh baris dalam tabel","Hanya mengubah baris pertama","Tidak pernah berjalan","Menghapus database otomatis"],"explanation":"Tanpa WHERE, semua baris dapat terkena perubahan."},{"number":44,"section":"DATABASE","level":"Hard","question":"Tabel SISWA memiliki kolom NIS, Nama, Nama_Wali_Kelas. Banyak siswa dalam kelas yang sama mengulang nama wali kelas. Risiko desain ini adalah ...","options":["Redundansi data dan anomali pembaruan","Tidak bisa menyimpan angka","Primary key otomatis hilang","Query SELECT tidak bisa digunakan"],"explanation":"Data wali kelas berulang; perubahan nama perlu dilakukan di banyak baris."},{"number":45,"section":"DATABASE","level":"Hard","question":"Untuk menampilkan nama siswa beserta nama kelas dari tabel SISWA dan KELAS yang terhubung melalui id_kelas, operasi yang digunakan adalah ...","options":["JOIN","DROP","TRUNCATE","COMMIT"],"explanation":"JOIN menggabungkan data dari tabel terkait."},{"number":46,"section":"DATABASE","level":"Expert","question":"Diketahui SISWA(NIS, Nama, id_kelas) dan KELAS(id_kelas, nama_kelas). Query yang tepat untuk menampilkan nama siswa dan nama kelas adalah ...","options":["SELECT SISWA.Nama, KELAS.nama_kelas FROM SISWA JOIN KELAS ON SISWA.id_kelas=KELAS.id_kelas;","SELECT Nama, nama_kelas FROM SISWA DELETE KELAS;","SELECT * FROM SISWA WHERE KELAS;","JOIN SISWA AND KELAS;"],"explanation":"JOIN menghubungkan kedua tabel melalui pasangan foreign key dan primary key."},{"number":47,"section":"DATABASE","level":"Expert","question":"Satu transaksi pembayaran terdiri dari header transaksi dan beberapa detail barang. Mengapa tabel TRANSAKSI dan DETAIL_TRANSAKSI dipisah?","options":["Agar satu transaksi dapat memiliki banyak item tanpa mengulang data header di setiap kolom","Agar tidak perlu primary key","Agar harga tidak dapat berubah","Agar semua data menjadi teks"],"explanation":"Pemisahan header-detail merepresentasikan relasi one-to-many dan mengurangi duplikasi."},{"number":48,"section":"DATABASE","level":"Expert","question":"Dua pengguna memperbarui stok barang yang sama hampir bersamaan. Konsep database yang membantu menjaga konsistensi transaksi adalah ...","options":["Transaction dan kontrol konkurensi","Mengubah nama tabel","Menghapus foreign key","Mengganti semua tipe data menjadi TEXT"],"explanation":"Transaksi dan mekanisme konkurensi membantu mencegah pembaruan saling menimpa dan menjaga konsistensi."},{"number":49,"section":"DATABASE","level":"Expert","question":"Dalam normalisasi, tujuan utama 1NF adalah ...","options":["Setiap atribut menyimpan nilai atomik dan tidak ada kelompok berulang","Semua tabel harus memiliki tepat dua kolom","Semua data disimpan dalam satu tabel","Semua kolom harus berupa angka"],"explanation":"1NF mengharuskan nilai atribut bersifat atomik dan struktur tidak memiliki repeating groups."},{"number":50,"section":"DATABASE","level":"Expert","question":"Sebuah aplikasi mencatat detail nilai siswa per mata pelajaran. Ketergantungan yang tepat untuk mengidentifikasi nilai adalah pasangan NIS dan kode_mapel. Kunci yang paling masuk akal adalah ...","options":["NIS saja","kode_mapel saja","Kombinasi NIS dan kode_mapel","Nama mata pelajaran dan nilai"],"explanation":"Kombinasi siswa dan mata pelajaran mengidentifikasi satu nilai per siswa per mapel, jika aturan bisnisnya demikian."}];
const TOTAL = 50;
const DURATION = 30 * 60;
const LETTERS = ["A","B","C","D"];
const $ = id => document.getElementById(id);

let sessionToken = "";
let student = null;
let current = 0;
let answers = Array(TOTAL).fill("");
let seconds = DURATION;
let timerId = null;
let examStarted = false;
let submitting = false;

document.addEventListener("DOMContentLoaded", () => {
  $("loginForm").addEventListener("submit", login);
  $("prevBtn").addEventListener("click", () => navigate(-1));
  $("nextBtn").addEventListener("click", () => navigate(1));
  $("finishBtn").addEventListener("click", () => finish(false));
  $("reviewBtn").addEventListener("click", toggleReview);
  $("printBtn").addEventListener("click", () => window.print());
  $("logoutBtn").addEventListener("click", () => location.reload());
});

async function api(action, data={}) {
  if (!API_URL || API_URL.includes("TEMPEL_URL")) {
    throw new Error("URL Google Apps Script belum dipasang pada API_URL di script.js.");
  }
  // Apps Script menerima POST text/plain agar request sederhana lintas domain.
  // Jika browser menolak CORS, host halaman perlu dipindahkan ke Apps Script HTML Service
  // atau menggunakan endpoint proxy tepercaya.
  const response = await fetch(API_URL, {
    method: "POST",
    headers: {"Content-Type":"text/plain;charset=utf-8"},
    body: JSON.stringify({action, ...data})
  });
  if (!response.ok) throw new Error("Server Google Apps Script tidak merespons.");
  const result = await response.json();
  if (!result.ok) throw new Error(result.error || "Permintaan gagal.");
  return result;
}

async function login(event) {
  event.preventDefault();
  const nis = $("nis").value.trim();
  const password = $("password").value;
  if (!nis || !password) return;
  const btn = $("loginBtn");
  btn.disabled = true;
  $("loginMessage").textContent = "Memeriksa akun siswa...";
  try {
    const result = await api("login", {nis, password});
    sessionToken = result.token;
    student = result.student;
    const started = await api("startExam", {token:sessionToken});
    if (started.alreadySubmitted) throw new Error("Akun ini sudah menyelesaikan ujian.");
    beginExam();
  } catch (err) {
    $("loginMessage").textContent = err.message + (err.message.toLowerCase().includes("cors") ? " Periksa konfigurasi deployment Apps Script." : "");
  } finally {
    btn.disabled = false;
  }
}

function beginExam() {
  current=0; answers=Array(TOTAL).fill(""); seconds=DURATION; examStarted=true;
  $("studentInfo").textContent = `${student.name} | ${student.className} | Absen ${student.number}`;
  $("loginView").classList.add("hidden");
  $("examView").classList.remove("hidden");
  renderQuestion(); renderNavigation(); updateProgress(); startTimer();
}

function renderQuestion() {
  const q=QUESTIONS[current];
  $("sectionLabel").textContent=q.section==="DATABASE"?"BASIS DATA":q.section==="WORD"?"MICROSOFT WORD":"MICROSOFT EXCEL";
  $("levelLabel").textContent=q.level;
  $("questionNumber").textContent=String(current+1).padStart(2,"0");
  $("questionText").textContent=q.question;
  const wrap=$("options"); wrap.replaceChildren();
  q.options.forEach((option,i)=>{
    const letter=LETTERS[i], button=document.createElement("button");
    button.type="button"; button.className="option"+(answers[current]===letter?" selected":"");
    const key=document.createElement("span"); key.className="option-key"; key.textContent=letter;
    const text=document.createElement("span"); text.textContent=option;
    button.append(key,text);
    button.addEventListener("click",()=>{if(submitting)return;answers[current]=letter;renderQuestion();renderNavigation();updateProgress();});
    wrap.appendChild(button);
  });
  $("prevBtn").disabled=current===0;
  $("nextBtn").textContent=current===TOTAL-1?"KEMBALI KE SOAL 1 ↺":"SELANJUTNYA →";
}
function renderNavigation() {
  const grid=$("numberGrid"); grid.replaceChildren();
  QUESTIONS.forEach((q,i)=>{
    const b=document.createElement("button");b.type="button";b.className="number-btn";
    if(answers[i])b.classList.add("done");if(i===current)b.classList.add("current");
    b.textContent=i+1;b.addEventListener("click",()=>{current=i;renderQuestion();renderNavigation();updateProgress();});
    grid.appendChild(b);
  });
}
function navigate(direction) {
  current=direction>0?(current+1)%TOTAL:Math.max(0,current-1);
  renderQuestion();renderNavigation();updateProgress();
}
function updateProgress() {
  const count=answers.filter(Boolean).length;
  $("progressText").textContent=`SOAL ${current+1} DARI ${TOTAL}`;
  $("answeredText").textContent=`${count}/${TOTAL} TERJAWAB`;
  $("progressBar").style.width=(count/TOTAL*100)+"%";
}
function startTimer() {
  clearInterval(timerId);showTime();
  timerId=setInterval(()=>{
    seconds=Math.max(0,seconds-1);showTime();
    if(seconds===600)alert("Waktu tersisa 10 menit.");
    if(seconds===300)alert("Waktu tersisa 5 menit.");
    if(seconds===0)finish(true);
  },1000);
}
function showTime() {
  $("timer").textContent=String(Math.floor(seconds/60)).padStart(2,"0")+":"+String(seconds%60).padStart(2,"0");
  $("timer").style.color=seconds<=300?"#ff5968":"#fff";
}
async function finish(timeUp) {
  if(submitting)return;
  const blank=answers.filter(a=>!a).length;
  if(!timeUp&&!confirm(`Masih ada ${blank} soal belum dijawab. Yakin menyelesaikan ujian?`))return;
  submitting=true;clearInterval(timerId);
  $("finishBtn").disabled=true;
  $("saveMessage").textContent="Mengirim jawaban dan menghitung nilai di server...";
  try {
    const result=await api("submitExam",{
      token:sessionToken,examId:"REFLEKSI-XI-RPL-50",
      answers, status:timeUp?"WAKTU HABIS":"SELESAI"
    });
    showResult(result,timeUp);
  } catch(err) {
    submitting=false;$("finishBtn").disabled=false;
    $("saveMessage").textContent="Gagal menyimpan: "+err.message+". Jangan tutup halaman; coba kirim kembali.";
    $("saveMessage").style.color="#ff858b";
    // Timer has stopped; the student may retry submission.
  }
}
function showResult(data,timeUp) {
  $("examView").classList.add("hidden");$("resultView").classList.remove("hidden");
  $("resultStudent").textContent=`${student.name} | ${student.className} | Absen ${student.number}`+(timeUp?" | WAKTU HABIS":"");
  $("finalScore").textContent=data.score;
  $("category").textContent=data.category;
  $("correctCount").textContent=data.correct;
  $("wrongCount").textContent=data.wrong;
  $("blankCount").textContent=data.blank;
  $("saveMessage").textContent="Berhasil disimpan ke Google Sheets.";
  $("saveMessage").style.color="#6ce0a8";
  buildReview(data.answerKey || []);
}
function buildReview(key) {
  const list=$("reviewList");list.replaceChildren();
  QUESTIONS.forEach((q,i)=>{
    const correct=key[i]||"";
    const item=document.createElement("div");item.className="review-item";
    const h=document.createElement("h3");h.textContent=`${i+1}. [${q.section} • ${q.level}] ${q.question}`;
    const a=document.createElement("p");a.className=answers[i]===correct?"right":"wrong";
    a.textContent="Jawaban Anda: "+(answers[i]?answers[i]+". "+q.options[LETTERS.indexOf(answers[i])]:"Tidak dijawab");
    const k=document.createElement("p");k.className="right";
    k.textContent=`Kunci: ${correct}. ${q.options[LETTERS.indexOf(correct)]}`;
    const x=document.createElement("p");x.textContent="Pembahasan: "+q.explanation;
    item.append(h,a,k,x);list.appendChild(item);
  });
}
function toggleReview() {
  $("reviewSection").classList.toggle("hidden");
  $("reviewBtn").textContent=$("reviewSection").classList.contains("hidden")?"LIHAT PEMBAHASAN":"TUTUP PEMBAHASAN";
}
