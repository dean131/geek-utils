import pandas as pd
import mysql.connector
from mysql.connector import Error


def smart_update_from_excel():
    """
    Secara cerdas memperbarui data di database MySQL dari file Excel.
    Logika: Cek by ID -> Cek by Name. Jika tidak cocok, data akan dilewati.
    """
    # --- 1. Konfigurasi (SESUAIKAN BAGIAN INI DENGAN TELITI) ---

    db_config = {
        "host": "10.0.0.192",
        "database": "terrafarma_staging",
        "user": "admintera",
        "password": "1q2wteradb",
        "port": 3306,  # Port default MySQL, sesuaikan jika perlu
    }

    excel_file_path = "daftar_harga.xls"
    excel_sheet_name = 0

    # Nama kolom di Excel
    excel_id_column = "r56"
    excel_name_column = "Nama"  # SESUAIKAN dengan nama kolom di Excel-mu
    excel_price_column = "HNA + PPN"

    # Nama tabel dan kolom di Database
    db_table_name = "items"
    db_id_column = "item_id"
    db_name_column = "name"  # SESUAIKAN dengan nama kolom di database-mu
    db_price_column = "new_price"

    # --- 2. Baca Data dari Excel ---
    try:
        print(f"Membaca data dari file Excel: {excel_file_path}...")
        df = pd.read_excel(excel_file_path, sheet_name=excel_sheet_name)
        df[excel_id_column] = df[excel_id_column].astype(str)
        print(f"Berhasil memuat {len(df)} baris data.")
    except Exception as e:
        print(f"Error saat membaca file Excel: {e}")
        return

    # --- 3. Proses Pembaruan Cerdas ke Database ---
    conn = None
    # Inisialisasi counter untuk summary
    updated_by_id = 0
    updated_by_name = 0
    not_found_rows = 0
    failed_rows = 0
    skipped_rows = 0

    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        print("Koneksi database berhasil. Memulai proses sinkronisasi...")

        for index, row in df.iterrows():
            item_id = row[excel_id_column]
            item_name = row[excel_name_column]
            new_price = row[excel_price_column]

            if pd.isna(item_id) or pd.isna(item_name) or pd.isna(new_price):
                print(
                    f"Melewati baris Excel ke-{index+2} karena ada data penting yang kosong."
                )
                skipped_rows += 1
                continue

            try:
                # LANGKAH 1: Coba cari berdasarkan ID
                cursor.execute(
                    f"SELECT {db_id_column} FROM {db_table_name} WHERE {db_id_column} = %s",
                    (item_id,),
                )
                result = cursor.fetchone()

                if result:
                    # Jika ID ditemukan, UPDATE harga berdasarkan ID
                    cursor.execute(
                        f"UPDATE {db_table_name} SET {db_price_column} = %s WHERE {db_id_column} = %s",
                        (new_price, item_id),
                    )
                    updated_by_id += 1
                    continue  # Lanjut ke baris berikutnya

                # LANGKAH 2: Jika ID tidak ditemukan, coba cari berdasarkan NAMA
                cursor.execute(
                    f"SELECT {db_id_column} FROM {db_table_name} WHERE {db_name_column} = %s",
                    (item_name,),
                )
                result = cursor.fetchone()

                if result:
                    # Jika Nama ditemukan, UPDATE harga berdasarkan NAMA
                    cursor.execute(
                        f"UPDATE {db_table_name} SET {db_price_column} = %s WHERE {db_name_column} = %s",
                        (new_price, item_name),
                    )
                    updated_by_name += 1
                    continue  # Lanjut ke baris berikutnya

                # LANGKAH 3: Jika ID dan Nama tidak ditemukan, lewati dan catat
                print(
                    f"Info: Item ID {item_id} atau Nama '{item_name}' tidak ditemukan. Data dilewati."
                )
                not_found_rows += 1

            except Error as e:
                print(f"Gagal memproses baris Excel ke-{index+2} (ID: {item_id}): {e}")
                failed_rows += 1

        conn.commit()

    except Error as e:
        print(f"\nTerjadi error database: {e}")
        if conn:
            conn.rollback()
    finally:
        if conn and conn.is_connected():
            cursor.close()
            conn.close()
            print("\nKoneksi ke database ditutup.")

    # --- 4. Tampilkan Laporan/Summary Akhir ---
    print("\n--- RINGKASAN PROSES (SUMMARY) ---")
    print(f"✅ Data Diperbarui berdasarkan ID  : {updated_by_id}")
    print(f"✅ Data Diperbarui berdasarkan Nama: {updated_by_name}")
    print(f"🤷 Data Tidak Ditemukan di DB      : {not_found_rows}")
    print(f"⏭️ Data Dilewati (kosong)        : {skipped_rows}")
    print(f"❌ Data Gagal Diproses             : {failed_rows}")
    print("---------------------------------------")
    print(f"Total baris di Excel              : {len(df)}")


if __name__ == "__main__":
    # PERINGATAN: Skrip ini HANYA akan memperbarui data yang ada. Selalu BACKUP database Anda terlebih dahulu!
    smart_update_from_excel()
