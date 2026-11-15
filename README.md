# 🏢 Admin MBG — Sistem Administrasi Kepegawaian & Terminal Absensi Wajah

Aplikasi web modern untuk manajemen administrasi staf, absensi selfie biometrik dengan memori wajah, deteksi shift kerja otomatis, pemantauan kontrak, dan pengolahan data Microsoft Excel untuk program **Makan Bergizi Gratis (MBG) Unit SPPG**.

---

## 🌟 Fitur Unggulan

### 1. 📷 Terminal Absensi Wajah (Kiosk Mobile & Tablet)
- **Face Memory Scanner**: Menyimpan memori wajah master staf secara otomatis pada selfie pertama dan memverifikasi saat presensi berikutnya.
- **Otomatisasi Shift Kerja**: Mendeteksi waktu kedatangan berdasarkan shift kerja masing-masing staf (*Reguler, Shift Pagi/Dapur, Shift Distribusi, Shift Siang*) dengan toleransi keterlambatan 10 menit.
- **Smart Presensi (Masuk & Pulang)**: Otomatis mendeteksi saat staf sudah absen masuk dan mengarahkan ke mode **Absen Pulang (Jam Keluar)**.
- **Audio Chime & Haptic Feedback**: Konfirmasi audio 2-nada (Web Audio API) dan getaran haptic mobile saat presensi berhasil dicatat.
- **Glassmorphism UI**: Tampilan kiosk modern dengan live digital clock dan frame panduan wajah.

### 2. 👥 Manajemen Staf & Memori Wajah (Admin Dashboard)
- **Tipe Kepegawaian Fleksibel**: Mendukung status **Kontrak**, **Non-Kontrak (Tetap)**, dan **Relawan/Volunteer**. Validasi tanggal berakhir otomatis menyesuaikan tipe kepegawaian.
- **Pengelolaan Memori Wajah**: Admin dapat melihat foto memori wajah staf, mengunggah foto baru secara manual, atau mereset memori wajah agar staf dapat foto ulang di terminal.
- **Monitoring Kontrak Cerdas**: Deteksi masa berlaku kontrak kerja dengan peringatan dini ($\le 30$ hari dan $\le 90$ hari).
- **Cetak Surat Kontrak**: Generator surat perpanjangan kontrak kerja instan siap cetak / PDF.

### 3. 📊 Presensi Harian & Rekap Bulanan
- **Daftar Absensi Harian**: Dilengkapi thumbnail foto selfie staf (klik untuk memperbesar dalam lightbox), jadwal shift, jam masuk, jam keluar, dan catatan presensi.
- **Rekap Bulanan**: Kalkulasi otomatis persentase kehadiran, jumlah hari kerja efektif, dan performa staf.

### 4. 📗 Ekspor Excel Asli (.xlsx & .xls)
- Menghasilkan file Excel rapi dengan header styling, multi-kolom rapi, dan auto-fit width menggunakan mesin **SheetJS** dan XML generator.
- Ekspor mencakup: Database Seluruh Staf, Absensi Harian, Rekap Bulanan, dan Riwayat Kontrak.

### 5. 🔐 Login Multi-Akun & Sinkronisasi Lintas Perangkat
- Didukung backend sync (`api-users.php` + `users_db.json`) sehingga akun yang didaftarkan di satu perangkat dapat langsung login di perangkat lain dalam jaringan lokal Wi-Fi / LAN.
- Persistent session: Login tetap tersimpan meskipun browser ditutup.

---

## 👥 Akun Bawaan (Default Login)

| Username | Password | Nama | Role / Akses |
|---|---|---|---|
| `admin` | `admin123` | Adli Zain Pribadi | Administrator (Akses Penuh) |
| `operator` | `op123` | Operator SPPG | Operator (Input & Absensi) |
| `viewer` | `view123` | Pengawas Audit | Viewer (Lihat Data & Laporan) |

> *Catatan: Anda juga dapat mendaftarkan akun baru melalui tombol **"Registrasi Akun Baru"** pada halaman login.*

---

## 💻 Cara Menjalankan

### Opsi A: Menggunakan XAMPP (Direkomendasikan untuk Sinkronisasi Lintas Perangkat)
1. Salin folder proyek ke direktori `C:\xampp\htdocs\admin-mbg\`
2. Buka **XAMPP Control Panel** dan nyalakan modul **Apache**.
3. Akses melalui browser:
   - **PC Lokal**: [http://localhost/admin-mbg/](http://localhost/admin-mbg/)
   - **Terminal Absensi HP / Kiosk**: [http://localhost/admin-mbg/absensi-terminal.html](http://localhost/admin-mbg/absensi-terminal.html)
   - **Akses dari HP / Laptop Lain**: Buka `http://<IP-Komputer-Anda>/admin-mbg/` (misal: `http://192.168.1.44/admin-mbg/`)

### Opsi B: Web Server Python
```bash
python -m http.server 3000
```
Buka browser pada `http://localhost:3000`

### Opsi C: Akses Online (GitHub Pages)
- **URL**: [https://adlizainpribadi65-alt.github.io/Aplikasi-Admin-MBG/](https://adlizainpribadi65-alt.github.io/Aplikasi-Admin-MBG/)
- **Terminal Absensi**: [https://adlizainpribadi65-alt.github.io/Aplikasi-Admin-MBG/absensi-terminal.html](https://adlizainpribadi65-alt.github.io/Aplikasi-Admin-MBG/absensi-terminal.html)

---

## 📁 Struktur Berkas

```
admin-mbg/
├── index.html              # Halaman Login & Registrasi Akun
├── dashboard.html          # Panel Kontrol Admin Utama
├── absensi-terminal.html   # Terminal Kiosk Selfie Biometrik & Shift
├── api-users.php           # API Sinkronisasi Pengguna Lintas Perangkat
├── users_db.json           # Database Akun Pengguna
├── assets/
│   ├── css/
│   │   └── style.css       # Tema Light SaaS, Animasi & Glassmorphism
│   └── js/
│       └── app.js          # Logika Bisnis, CRUD Staf, Absensi, Excel
└── README.md               # Dokumentasi Aplikasi
```

---
**Pengembang**: [adlizainpribadi65-alt](https://github.com/adlizainpribadi65-alt)  
**Lisensi**: MIT
