const express = require('express');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');
const ejs = require('ejs');
const bodyParser = require('body-parser');

const app = express();
const fs = require('fs');

app.use(bodyParser.urlencoded({ extended: true }));

app.use(express.static(path.join(__dirname, 'public')));

// Mengatur penyimpanan file menggunakan multer
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

// Pengaturan tampilan mesin render
app.set('view engine', 'ejs');

// Menampilkan halaman utama
app.get('/', function (req, res) {
  res.render('upload');
});

// Menangani unggahan file Excel
app.post('/upload', upload.single('excelFile'), function (req, res) {
  const file = req.file;
  const workbook = xlsx.readFile(file.path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);

  res.render('calculate', { data });
});

// Menampilkan hasil perhitungan berdasarkan nama
app.post('/result', function (req, res) {
  console.log(req.body);
  const name = req.body.name;
  const data = JSON.parse(req.body.data);
  
  // Cari data berdasarkan nama yang dimasukkan
  const person = data.find((item) => item.Nama === name);
  
  if (person) {
    // Lakukan perhitungan gaji akhir
    
    const divisi = person['Divisi'];
    const level = person['Level'];
    const jabatan = person['Jabatan'];
    const rekening = person['Rekening'];
    const gaji = person['Gaji Pokok'];
    const tunjangan_pokok = person['Tunjangan Utama'];
    const tunjangan_lain = person['Tunjangan Lain'];
    const total_tunjangan = tunjangan_pokok + tunjangan_lain;
    const gaji_tunjangan = gaji+ total_tunjangan;
    const kpi = person['KPI'];
    const komisi_dsa = person['Komisi DSA'];
    const total_komisi = kpi+komisi_dsa;
    const gaji_komisi = gaji_tunjangan+total_komisi;
    const bpjs_kes = person['Potongan BPJS Kesehatan'];
    const bpjs_ket = person['Potongan BPJS Ketenagakerjaan'];
    const pph_21 = person['Potongan Pph 21'];
    const hp_cicilan = person['Potongan Hp Cicilan 12']
    const absensi = person['Potongan Absensi']
    const potongan = bpjs_kes + bpjs_ket + pph_21 + hp_cicilan + absensi;
    const payout = gaji_komisi - potongan;
    // const gajiAkhir = gaji - potongan + tunjangan;
    
    
    // Tampilkan hasil perhitungan pada halaman hasil
    res.render('result', { name, divisi, level, jabatan, rekening, gaji, tunjangan_pokok, tunjangan_lain,
      total_tunjangan,gaji_tunjangan, kpi, komisi_dsa, total_komisi, gaji_komisi, bpjs_kes, bpjs_ket, pph_21,
      hp_cicilan, absensi, potongan, payout});
  } else {
    res.render('result', { error: 'Nama tidak ditemukan' });
  }
});

app.listen(3000, function () {
  console.log('Server berjalan pada http://localhost:3000');
});
