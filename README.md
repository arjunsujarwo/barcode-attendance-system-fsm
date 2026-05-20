# Sistem Integrasi HRD & Absensi Barcode (Timing Lockout Version)
### Dokumentasi Proyek Akhir Sistem Informasi - Universitas Bina Sarana Informatika (UBSI)

---

## 1. Deskripsi Projek

Aplikasi **Sistem Integrasi HRD & Absensi Barcode** adalah perangkat lunak berbasis desktop yang dirancang mandiri (*kiosk-mode*) untuk mengotomatisasi pencatatan kehadiran karyawan menggunakan teknologi pemindaian kode batang (*barcode*). Sistem ini mengintegrasikan fungsi manajemen data karyawan tetap dengan pencatatan log kehadiran harian secara *real-time* ke dalam basis data berbentuk spreadsheet.

Aplikasi ini dilengkapi dengan fitur **Timing Lockout (Jeda Waktu Keamanan)** untuk mencegah kecurangan absensi ganda (antrean scan berulang pada detik yang sama) serta memastikan alur kerja harian karyawan tercatat secara valid sesuai kebijakan operasional.

### Teknologi yang Digunakan:
* **Python**: Bahasa pemrograman utama untuk membangun logika bisnis.
* **Tkinter & TTK**: *Library* bawaan Python untuk merancang Antarmuka Pengguna Grafis (GUI) dengan gaya klasik (*Windows Classic Style*).
* **OpenPyXL**: *Library* penanganan I/O berkas untuk membaca, menulis, dan memformat basis data Microsoft Excel (`.xlsx`) secara multi-sheet.
* **OpenCV (`opencv-python`)**: Mengendalikan dan memproses input video digital dari perangkat keras kamera/webcam secara *real-time*.
* **Pyzbar**: Menangani algoritma pemecahan kode (dekode) gambar mentah menjadi string data barcode.
* **Python-Barcode**: Menghasilkan (*generate*) citra barcode baru berstandar `Code 128` berdasarkan Nomor ID Karyawan.

### Fitur Utama:
1. **🖥️ Mesin Absensi Real-Time**: Mendukung input ganda melalui mesin pemindai barcode fisik (perangkat keras USB *scanner*) maupun pemindaian visual menggunakan kamera laptop/webcam.
2. **⚙️ Master Karyawan & Generator Barcode**: Modul pendaftaran karyawan baru yang secara otomatis menggenerasi file gambar barcode `.png` untuk dicetak sebagai kartu ID Card.
3. **📊 Rekap Data Bulanan**: Fitur ekstraksi data otomatis yang memfilter log kehadiran berdasarkan bulan yang dipilih, serta menghasilkan berkas laporan baru dengan format sel yang rapi dan siap cetak.

---

## 2. Prasyarat Sistem (Prerequisites)

Sebelum menjalankan aplikasi, pastikan perangkat keras dan lingkungan sistem operasi Anda telah memenuhi kriteria berikut:

### Lingkungan Perangkat Lunak:
* **Python**: Versi **3.8 hingga 3.11** (Sangat direkomendasikan menggunakan Python 3.10 ke atas).

### Ketergantungan Sistem Operasi (Khusus Windows):
Library `pyzbar` bergantung pada pustaka eksternal sistem. Jika aplikasi mengalami *crash* saat membaca barcode lewat kamera, Anda wajib menginstal komponen berikut:
1. **Visual C++ Redistributable Packages for Visual Studio**: Diperlukan oleh komponen biner OpenCV dan Pyzbar.
2. **ZBar Shared Library**: Jika pustaka `zbar` tidak terbaca di Windows, unduh installer `zbar` Windows dari pustaka resmi, atau gunakan paket instalasi Python yang sudah membundel *DLL dependency*.

---

## 3. Panduan Instalasi Langkah-demi-Langkah

Ikuti instruksi di bawah ini melalui Terminal atau Command Prompt (cmd) untuk melakukan setup lingkungan aplikasi:

### Langkah 1: Setup Folder Projek
Buat sebuah folder baru di komputer Anda (misal di direktori `D:/ProyekAkhirUBSI`) dan letakkan file kode sumber `app_terpadu.py` di dalam folder tersebut.

### Langkah 2: Membuat Virtual Environment (Isolasi Library)
Sangat disarankan untuk membuat *Virtual Environment* agar pustaka yang diinstal tidak bentrok dengan projek Python lainnya di komputer Anda.
```bash
# Buka terminal dan arahkan ke folder projek Anda
cd /d D:\ProyekAkhirUBSI

# Membuat Virtual Environment baru dengan nama 'venv'
python -m venv venv

### Langkah 3: Aktivasi Virtual EnvironmentSebelum menginstal library, aktifkan terlebih dahulu virtual environment yang baru dibuat:Untuk Windows (Command Prompt):DOSvenv\Scripts\activate
Untuk Windows (PowerShell):PowerShell.\venv\Scripts\activate
Untuk Linux/MacOS:Bashsource venv/bin/activate
(Indikator jika berhasil: Akan muncul tanda (venv) di bagian depan baris perintah terminal Anda).
### Langkah 4: Instalasi Library EksternalInstal seluruh dependensi library eksternal yang dibutuhkan oleh aplikasi dengan mengeksekusi perintah tunggal berikut:Bashpip install openpyxl opencv-python pyzbar python-barcode
Tunggu hingga proses unduhan selesai dan pastikan tidak ada indikator error berwarna merah.4. Panduan Menjalankan AplikasiSetelah seluruh pustaka berhasil terinstal, pastikan Virtual Environment Anda masih dalam posisi aktif, lalu jalankan perintah berikut untuk membuka aplikasi:Bashpython app_terpadu.py
Jendela aplikasi "HRD System & Attendance Barcode v2.1" akan muncul di layar dan siap digunakan untuk demonstrasi atau pengujian operasional.

## 5. Kompilasi Menjadi Aplikasi Executable (.exe)

Aplikasi ini dapat dikompilasi menjadi satu file eksekusi `.exe` mandiri menggunakan `PyInstaller` sehingga pengguna akhir (Admin/HRD) tidak perlu menginstal Python di komputer mereka.

### Perintah Kompilasi (Windows 64-bit):
Pastikan Anda berada di direktori projek, lalu jalankan perintah berikut pada PowerShell:
```bash
python -m PyInstaller --onefile --noconsole --add-data "C:\Users\SHAULA\AppData\Roaming\Python\Python312\site-packages\pyzbar;pyzbar" app_terpadu.py

### 6. Aturan Bisnis Aplikasi (Fitur Timing Lockout)Sistem ini menerapkan metode Finite State Machine (FSM) searah yang mengatur status kehadiran karyawan dalam 1 hari kerja (24 jam).
 Untuk menjaga validitas data, terdapat fitur Timing Lockout (pembatasan jeda waktu scan berturut-turut):
 Status Saat IniStatus Target BerikutnyaKetentuan & Jeda Waktu Keamanan (Lockout)BELUM_MASUKSUDAH_MASUKScan pertama di hari kerja. 
 Mencatat jam masuk kantor.SUDAH_MASUKISTIRAHATKaryawan hanya diperbolehkan scan keluar istirahat minimal 1 jam setelah melakukan scan masuk kerja.
 ISTIRAHATKEMBALI_KERJAKaryawan hanya diperbolehkan scan kembali bekerja minimal 5 menit setelah melakukan scan keluar istirahat.
 KEMBALI_KERJAPULANGKaryawan hanya diperbolehkan scan pulang minimal 1 jam terhitung sejak waktu mulai istirahat tercatat.
 PULANG(Ditolak)Sistem akan memblokir scan ke-5 dst.
 Memunculkan notifikasi bahwa absensi hari ini telah selesai.
 Catatan Uji Coba Penguji: 
 Jika Anda melakukan pemindaian ID Card yang sama berulang-ulang dalam waktu singkat, sistem akan otomatis mengeluarkan peringatan berwarna merah dan menolak perintah manipulasi data demi keamanan log database.
 
 ### 7. Struktur File ProjekSetelah aplikasi dijalankan pertama kali, mendaftarkan karyawan, dan mengekspor laporan, sistem akan membentuk struktur pohon folder (tree directory) secara otomatis sebagai berikut:
ProyekAkhirUBSI/
│
├── venv/                           # Folder isolasi library Python (otomatis terbuat)
├── hasil_barcode/                  # Folder penyimpanan berkas citra barcode karyawan
│   ├── 12345_Arjun Sujarwo.png
│   └── 67890_Budi Santoso.png
│
├── app_terpadu.py                  # Berkas kode program utama (GUI & Logika Bisnis)
├── database_absensi.xlsx           # Database utama (Multi-Sheet: Master_Karyawan & Log_Absensi)
└── Laporan_Absensi_Historis_Mei_2026.xlsx  # Hasil ekspor rekapitulasi bulanan siap cetak
Keterangan Berkas Basis Data:database_absensi.xlsx -> Sheet Master_Karyawan: Menyimpan ID unik, nama karyawan, dan tanggal mendaftar secara permanen.database_absensi.xlsx -> Sheet Log_Absensi: Menyimpan riwayat transaksi harian (Tanggal, ID, Nama, Status Terakhir, Jam Masuk, Jam Istirahat, Jam Kembali, Jam Pulang).
