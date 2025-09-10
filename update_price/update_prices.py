import pandas as pd
import mysql.connector
from mysql.connector import Error


def update_database_from_excel():
    """
    Membaca data dari file Excel dan memperbarui kolom harga di tabel database MySQL.
    """
    # --- 1. Konfigurasi (SESUAIKAN BAGIAN INI) ---

    # Detail koneksi database MySQL
    db_config = {
        "host": "localhost",
        "database": "tera_clone",
        "user": "root",
        "password": "pass",
        "port": 3306,  # Port default MySQL, sesuaikan jika perlu
    }

    # Path ke file Excel Anda (.xls atau .xlsx)
    excel_file_path = "daftar_harga.xls"

    # (Opsional) Nama atau nomor sheet di Excel.
    # Gunakan 0 untuk sheet pertama, 1 untuk kedua, dst.
    # Atau gunakan namanya langsung, contoh: 'Laporan Harga'
    excel_sheet_name = 0

    # Nama kolom di Excel
    excel_id_column = "r56"
    excel_price_column = "HNA + PPN"

    # Nama tabel dan kolom di Database
    db_table_name = "items"
    db_id_column = "item_id"
    db_price_column = "new_price"

    # --- 2. Baca Data dari EXCEL (INI BAGIAN YANG BERUBAH) ---
    try:
        print(
            f"Membaca data dari file Excel: {excel_file_path} (Sheet: {excel_sheet_name})..."
        )
        df = pd.read_excel(excel_file_path, sheet_name=excel_sheet_name)

        # Pastikan kolom ID dibaca sebagai string agar tidak ada masalah tipe data
        if excel_id_column in df.columns:
            df[excel_id_column] = df[excel_id_column].astype(str)
        print(f"Berhasil memuat {len(df)} baris data dari Excel.")
    except FileNotFoundError:
        print(f"Error: File Excel tidak ditemukan di '{excel_file_path}'")
        return
    except Exception as e:
        print(f"Error saat membaca file Excel: {e}")
        return

    # --- 3. Proses Pembaruan ke Database (TIDAK ADA PERUBAHAN DI SINI) ---
    conn = None
    updated_rows = 0
    failed_rows = 0

    try:
        conn = mysql.connector.connect(**db_config)

        if conn.is_connected():
            cursor = conn.cursor()
            print("Koneksi ke database MySQL berhasil. Memulai proses update...")

            sql_query = f"""
                UPDATE {db_table_name}
                SET {db_price_column} = %s
                WHERE {db_id_column} = %s;
            """

            for index, row in df.iterrows():
                item_id = row[excel_id_column]
                new_price = row[excel_price_column]

                if pd.notna(item_id) and pd.notna(new_price):
                    try:
                        cursor.execute(sql_query, (new_price, item_id))
                        if cursor.rowcount > 0:
                            updated_rows += 1
                        else:
                            print(
                                f"Peringatan: Tidak ada baris yang cocok untuk item_id {item_id}."
                            )
                            failed_rows += 1
                    except Error as e:
                        print(f"Gagal update untuk item_id {item_id}: {e}")
                        failed_rows += 1
                else:
                    print(
                        f"Melewati baris ke-{index+2} karena item_id atau harga kosong."
                    )
                    failed_rows += 1

            conn.commit()
            print("\n--- Proses Selesai ---")
            print(f"Total baris yang berhasil diupdate: {updated_rows}")
            print(f"Total baris yang gagal/dilewati: {failed_rows}")

    except Error as e:
        print(f"\nTerjadi error pada koneksi atau transaksi database: {e}")
        if conn and conn.is_connected():
            conn.rollback()
            print("Transaksi dibatalkan (rollback).")
    finally:
        if conn and conn.is_connected():
            cursor.close()
            conn.close()
            print("Koneksi ke database ditutup.")


# Jalankan fungsi utama
if __name__ == "__main__":
    # PENTING: Selalu backup database Anda sebelum menjalankan skrip ini!
    update_database_from_excel()
