[Uploading README_DEPLOY.md…]()
# Panduan Deploy Pendamping ARKAS (Online + Data Permanen Gratis)

Panduan ini untuk mengonlinekan aplikasi dengan **Streamlit Community Cloud** (hosting gratis)
dan **Supabase** (database gratis) supaya data (profil sekolah, RKAS, BKU/SPJ, checklist) tidak
hilang setiap kali aplikasi di-restart. Aplikasi ini dirancang untuk **1 sekolah, tanpa login** —
semua data disimpan dalam satu tempat yang sama, dilindungi PIN bersama (opsional tapi disarankan
karena datanya online).

Total waktu: sekitar 15–20 menit, semua langkah pakai akun gratis (tidak perlu kartu kredit).

---

## Bagian 1 — Bikin Database di Supabase (gratis)

1. Buka **https://supabase.com** → **Start your project** → daftar/login (bisa pakai akun GitHub).
2. Klik **New Project**:
   - Isi nama project bebas, misalnya `arkas-sekolah`.
   - Buat **Database Password** — simpan baik-baik (tidak dipakai langsung di aplikasi, tapi
     jaga-jaga kalau nanti perlu akses langsung ke database).
   - Pilih region terdekat (misalnya Singapore) supaya lebih cepat.
   - Klik **Create new project** dan tunggu 1–2 menit sampai project siap.
3. Setelah project siap, buka menu **SQL Editor** (ikon di sidebar kiri) → **New query**, lalu
   tempel & jalankan (klik **Run**) SQL berikut untuk membuat tabel penyimpanan:

   ```sql
   create table if not exists arkas_state (
     id bigint primary key,
     data jsonb not null default '{}'::jsonb,
     updated_at timestamptz not null default now()
   );

   insert into arkas_state (id, data)
   values (1, '{}'::jsonb)
   on conflict (id) do nothing;

   alter table arkas_state enable row level security;

   -- Aplikasi ini 1 sekolah tanpa login, jadi kunci akses sebenarnya ada di
   -- PIN aplikasi (APP_PASSWORD) & kerahasiaan SUPABASE_KEY, bukan di RLS.
   -- Policy di bawah mengizinkan akses baca/tulis lewat anon key.
   create policy "izinkan akses anon" on arkas_state
     for all
     using (true)
     with check (true);
   ```

4. Ambil kredensial koneksi: menu **Project Settings** (ikon gerigi) → **Data API**:
   - Salin **Project URL** → ini nilai `SUPABASE_URL`.
   - Salin **anon public key** (di bagian API Keys) → ini nilai `SUPABASE_KEY`.
   - **Jangan** pakai `service_role` key di aplikasi Streamlit publik — itu key rahasia dengan
     akses penuh, cukup pakai `anon public`.

Database gratis Supabase cukup lebih dari cukup untuk 1 sekolah (batas gratis: 500 MB database —
data teks RKAS/BKU 1 sekolah jauh di bawah itu).

---

## Bagian 2 — Siapkan Repo GitHub

1. Buat repo baru di GitHub (bisa privat atau publik), lalu unggah semua file ini:
   - `app_arkas.py`
   - `db_supabase.py`
   - `requirements_arkas.txt`
   - `README_ARKAS.md` (opsional, dokumentasi)
2. **Jangan** pernah commit file `secrets.toml` atau kredensial Supabase ke GitHub — kredensial
   diisi lewat fitur Secrets di Streamlit Cloud (Bagian 3), bukan di dalam file repo.

---

## Bagian 3 — Deploy ke Streamlit Community Cloud (gratis)

1. Buka **https://share.streamlit.io** → login pakai akun GitHub.
2. Klik **Create app** → **Deploy a public app from GitHub** (atau **From existing repo**).
3. Pilih repo, branch, dan set **Main file path** ke `app_arkas.py`.
   - Karena nama file requirements bukan `requirements.txt` standar, buka **Advanced settings**
     sebelum deploy dan tambahkan baris di bagian *Python environment* / atau paling gampang:
     ganti nama file jadi `requirements.txt` di repo (Streamlit Cloud otomatis mendeteksi nama
     standar ini tanpa konfigurasi tambahan).
4. Masih di **Advanced settings**, buka bagian **Secrets** dan isi:

   ```toml
   SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
   SUPABASE_KEY = "isi-anon-public-key-di-sini"
   APP_PASSWORD = "buat-pin-bebas-misalnya-6-digit"
   ```

   - `APP_PASSWORD` opsional — kalau dikosongkan/dihapus, aplikasi bisa dibuka siapa saja yang
     tahu link-nya tanpa PIN. **Disarankan diisi** karena URL Streamlit Cloud bersifat publik dan
     aplikasi ini menyimpan data keuangan sekolah.
5. Klik **Deploy**. Tunggu beberapa menit sampai build selesai — aplikasi akan online di alamat
   seperti `https://nama-app-anda.streamlit.app`.

Setelah online, buka aplikasinya, masukkan PIN (kalau diaktifkan), isi data seperti biasa, lalu
klik tombol **💾 Simpan** di sidebar setiap selesai mengisi/mengubah data. Data akan tetap ada
walau aplikasi di-restart, tidur (sleep karena tidak diakses), atau dibuka dari perangkat lain.

---

## Menjalankan & Menguji di Komputer Sendiri (opsional, sebelum deploy)

1. Install dependency: `pip install -r requirements_arkas.txt` (atau `requirements.txt` kalau
   sudah diganti nama).
2. Buat folder `.streamlit/` di sebelah `app_arkas.py`, lalu buat file `.streamlit/secrets.toml`:

   ```toml
   SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
   SUPABASE_KEY = "isi-anon-public-key-di-sini"
   APP_PASSWORD = "1234"
   ```

3. Jalankan: `streamlit run app_arkas.py`.
4. Kalau file `secrets.toml` belum dibuat / dikosongkan, aplikasi tetap jalan normal tanpa PIN dan
   tanpa penyimpanan permanen (perilaku sama seperti sebelum fitur ini ditambahkan).

---

## Batasan & Catatan Penting

- **1 sekolah, 1 set data.** Semua orang yang tahu PIN akan melihat & mengubah data yang sama
  (tidak ada pemisahan data per pengguna). Kalau nanti butuh banyak sekolah dengan akun masing-
  masing (login), beri tahu — tabelnya tinggal diubah jadi banyak baris + Supabase Auth, pola
  yang sama seperti pada Generator Perangkat Ajar KKG.
- **Simpan manual, bukan otomatis.** Supaya hemat & tidak membebani database, data hanya ditulis
  ke Supabase saat tombol **💾 Simpan** ditekan — perubahan yang belum disimpan akan hilang kalau
  browser ditutup/di-refresh sebelum sempat klik Simpan.
- **Tetap bukan ARKAS resmi.** Fitur online ini hanya membuat *alat bantu* ini bisa diakses dari
  mana saja & datanya tidak hilang — bukan menghubungkan ke sistem ARKAS pemerintah. Angka akhir
  tetap harus diinput manual ke ARKAS resmi.
- **Free tier Streamlit Cloud** akan membuat aplikasi "tidur" kalau tidak diakses beberapa hari —
  tinggal dibuka lagi dan akan otomatis bangun (data tetap aman di Supabase, tidak hilang).
