# Dokumentasi Skrip Pembaruan Harga dari Excel

Dokumen ini menjelaskan cara instalasi, konfigurasi, dan penggunaan skrip Python (`update_prices.py`) untuk memperbarui data harga di database MySQL dari sebuah file Excel.

---

## 📝 Deskripsi

Skrip ini dirancang untuk melakukan tugas-tugas berikut:

1.  **Membaca Data**: Mengambil informasi dari file Microsoft Excel (`.xlsx` atau `.xls`).
2.  **Koneksi Database**: Menghubungkan ke server database MySQL.
3.  **Pembaruan Data**: Memperbarui kolom harga (`new_price`) pada tabel `items` berdasarkan kecocokan data antara kolom identifier di Excel (`r56`) dan di database (`item_id`).

---

## 📂 Struktur Folder yang Direkomendasikan

Untuk menjaga kerapian, disarankan untuk menyusun file proyek seperti ini:

proyek-update-harga/
├── venv/ # Folder virtual environment
├── update_prices.py # Skrip utama
├── sumber_data.xlsx # File sumber data Excel
└── CARA_PENGGUNAAN.md # File dokumentasi ini

---

## ⚙️ Instalasi dan Setup

Ikuti langkah-langkah berikut untuk menyiapkan lingkungan kerja Anda.

### 1. Prasyarat

- **Python 3.8** atau versi yang lebih baru.
- **MySQL Server** sudah terinstal dan berjalan.
- Database dan tabel target sudah ada.

### 2. Buat Virtual Environment

Buka terminal di dalam folder proyek Anda dan jalankan perintah berikut untuk membuat lingkungan virtual yang terisolasi. Ini mencegah konflik antar pustaka proyek.

```bash
python -m venv venv
```

### 3. Aktifkan Virtual Environment

Aktivasi diperlukan setiap kali Anda akan bekerja pada proyek ini.

Untuk pengguna Fish Shell:

```bash
source venv/bin/activate.fish
```

Untuk pengguna Bash/Zsh (Linux/macOS):

```bash
source venv/bin/activate
```

Setelah aktif, nama (venv) akan muncul di awal baris terminal Anda.

### 4. Instal Pustaka yang Dibutuhkan

Instal semua pustaka Python yang diperlukan dengan satu perintah:

```bash
(venv) $ pip install pandas mysql-connector-python openpyxl
```

🔧 Konfigurasi Skrip

Buka file update_prices.py dan sesuaikan variabel di dalam bagian --- 1. Konfigurasi --- sesuai dengan kebutuhan Anda.
Python

# --- 1. Konfigurasi (SESUAIKAN BAGIAN INI) ---

### Detail koneksi database MySQL

db_config = {
'host': 'localhost',
'database': 'nama_database_anda',
'user': 'username_anda',
'password': 'password_anda',
'port': 3306
}

### Path ke file Excel Anda (.xls atau .xlsx)

excel_file_path = 'sumber_data.xlsx'

### Nama atau nomor sheet di Excel (0 untuk sheet pertama)

excel_sheet_name = 0

### Nama kolom identifier dan harga di file Excel

excel_id_column = 'r56'
excel_price_column = 'HNA + PPN'

### Nama tabel dan kolom di Database

db_table_name = 'items'
db_id_column = 'item_id'
db_price_column = 'new_price'

Pastikan semua nama kolom, nama tabel, dan detail koneksi sudah benar sebelum melanjutkan.

▶️ Menjalankan Skrip

Setelah semua konfigurasi selesai, jalankan skrip dari terminal dengan perintah berikut (pastikan virtual environment masih aktif).

```bash
(venv) $ python update_prices.py
```

Skrip akan menampilkan log proses di terminal, termasuk status koneksi, jumlah baris yang diproses, dan hasil akhir (berhasil, gagal, atau dilewati).

⚠️ Peringatan Penting

Selalu buat cadangan (backup) database Anda sebelum menjalankan skrip ini. Proses UPDATE akan mengubah data secara permanen. Adanya cadangan akan memastikan Anda dapat mengembalikan data jika terjadi kesalahan yang tidak diinginkan.
