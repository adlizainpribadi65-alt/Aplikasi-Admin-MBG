# Admin MBG — Sistem Administrasi Kepegawaian & Pengelolaan Data Excel

Aplikasi web modern untuk manajemen administrasi kepegawaian dan 10 modul pengolahan data Microsoft Excel pada program Makan Bergizi Gratis (MBG) Unit SPPG.

## 🚀 Fitur Utama

### A. Administrasi Kepegawaian
1. **Database Staf**: Pencatatan lengkap data diri, NIK KTP, kontak, alamat, jabatan, pendidikan, dan gaji pokok.
2. **Absensi Harian**: Form presensi harian (Hadir, Terlambat, Izin, Sakit, Alpha).
3. **Rekapitulasi Bulanan**: Akumulasi presensi, persentase kehadiran, dan kategori performa kerja otomatis.
4. **Manajemen Kontrak**: Monitoring jatuh tempo kontrak kerja (urgensi $\le 30$ hari dan $\le 90$ hari).
5. **Cetak Dokumen**: Pembuatan surat perpanjangan kontrak kerja otomatis siap cetak / simpan PDF.

### B. Pengelolaan Data Excel Terintegrasi
1. Database seluruh staf dengan validasi kolom.
2. Tabel identitas, jabatan, dan masa kerja.
3. Rekapitulasi absensi periodik.
4. Kalkulasi otomatis kehadiran, izin, sakit, dan keterlambatan.
5. Pemantauan berkala kontrak kerja staf.
6. **Simulator Rumus Excel Interaktif**: Mendukung formula `SUM`, `COUNTIF`, `IF`, `VLOOKUP`, dan `XLOOKUP`.
7. Laporan perpanjangan kontrak per semester/tahunan.
8. Pivot pengelompokan divisi & status keaktifan staf.
9. Laporan bulanan eksekutif disertai visualisasi grafik Chart.js.
10. Audit kebersihan, validitas, dan keakuratan data secara otomatis.

## 📗 Ekspor Excel Asli (.xlsx)
Didukung mesin **SheetJS (OpenXML)** lokal untuk menghasilkan dokumen `.xlsx` multi-sheet dengan format tabel rapi, header berwarna, dan auto-fit lebar kolom tanpa risiko data tertumpuk di Kolom A.

## 💻 Cara Menjalankan Lokal
1. Buka terminal pada folder proyek ini.
2. Jalankan perintah:
   ```bash
   python -m http.server 3000
   ```
3. Buka browser pada alamat:
   ```
   http://localhost:3000
   ```
   *(Atau cukup klik ganda file `start-server.bat`)*

---
**Pengembang**: [adlizainpribadi65-alt](https://github.com/adlizainpribadi65-alt)
