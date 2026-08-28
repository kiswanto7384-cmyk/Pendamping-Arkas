# Pendamping ARKAS (Tidak Resmi)

Aplikasi bantu untuk mempersiapkan & menganalisis data RKAS/BKU sekolah. **Bukan** aplikasi
resmi Kemendikdasmen dan **tidak terhubung/ber-API** ke Aplikasi ARKAS — ARKAS adalah sistem
tertutup milik pemerintah tanpa akses integrasi publik. Aplikasi ini berfungsi sebagai alat
bantu di LUAR ARKAS: menyiapkan angka sebelum diinput manual, dan menganalisis data setelah
diekspor dari ARKAS.

## Cara Menjalankan
```bash
pip install -r requirements_arkas.txt
streamlit run app_arkas.py
```

## Fitur
1. **🧮 Simulator RKAS** — susun rincian rencana kegiatan per komponen, otomatis dicek terhadap
   batas Juknis BOSP 2026 (minimal 10% buku, maksimal 20%/40% honor, maksimal 20% sarpras),
   plus estimasi PPN 12%. Bisa diunduh sebagai PDF (ringkasan cetak) & Excel (angka siap disalin
   manual ke ARKAS).
2. **📊 Analisis Laporan (BKU/RKAS)** — unggah hasil ekspor Excel/CSV dari ARKAS, petakan
   kolomnya (fleksibel karena format ekspor bisa berbeda-beda), dapatkan rekap otomatis per
   komponen, grafik proporsi, dan cek kepatuhan realisasi terhadap pagu.
3. **🧾 Kalkulator PPN** — hitung DPP & PPN 12% per transaksi atau banyak transaksi sekaligus.
4. **📅 Checklist Tahapan** — pantau progres tahapan RKAS/ARKAS sepanjang tahun (T-1, verifikasi,
   sinkron MARKAS-SIPD, aktivasi & tutup BKU bulanan, dll).
5. **📖 Panduan Singkat** — referensi cepat ketentuan BOSP 2026 tanpa perlu buka Juknis PDF.
6. **📋 SPJ/LPJ BOS (13 Modul)** — draf SPJ/LPJ otomatis dari satu tabel BKU (Buku Kas Umum):
   Isi Data, BKU, Pembantu Kas, Pembantu Bank, Pembantu Pajak, Pembantu Rincian Objek Belanja,
   Register Penutupan Kas, BAP Kas, LRA, SPTM, SPTMH, R PRPD, dan Validasi Data Rekening. Guru/
   bendahara cukup isi profil sekolah + input transaksi di BKU (tabel yang bisa ditambah baris
   sendiri), 12 laporan lainnya otomatis terhitung dan bisa diunduh sekaligus jadi 1 file Excel.
   Surat resmi (BAP Kas, SPTM, SPTMH) tersedia sebagai draf teks siap disalin ke Word untuk
   ditandatangani.

   > ⚠️ **PENTING:** Ini alat bantu menyusun DRAF, bukan pengganti ARKAS resmi. Formatnya
   > disusun berdasarkan pola umum LPJ BOS dan Juknis BOSP 2026, tapi ARKAS tetap satu-satunya
   > sistem pelaporan resmi ke pemerintah — selalu cocokkan angka akhir ke ARKAS dan konsultasikan
   > format final ke Dinas Pendidikan setempat sebelum ditandatangani/diserahkan.

## Catatan Penting
- Semua perhitungan kepatuhan di sini adalah **ringkasan tidak resmi** berdasarkan Permendikdasmen
  No. 8 Tahun 2026. Aturan bisa direvisi — selalu verifikasi ke Juknis resmi terbaru & Dinas
  Pendidikan setempat, terutama untuk kasus khusus (wilayah 3T, BOS Kinerja, dll).
- Data (rincian RKAS, checklist) saat ini tersimpan selama sesi browser terbuka. Kalau nanti mau
  ditambah login multi-pengguna + data tersimpan permanen (seperti pada Generator Perangkat Ajar
  KKG), tinggal minta — pola Supabase yang sama bisa dipakai di sini.
- Untuk deployment online (Streamlit Community Cloud, dll), ikuti pola yang sama seperti pada
  README_DEPLOY.md aplikasi Generator Perangkat Ajar KKG.
