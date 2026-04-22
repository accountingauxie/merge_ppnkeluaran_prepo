import pandas as pd
import glob
import os

# ==========================================
# AREA PENGATURAN (SILAKAN UBAH BAGIAN INI)
# ==========================================
periode_check = '31/1/2026'                     # Isi periode check di sini
folder_sumber = '/content/Data_Import/*.xlsx'   # Lokasi folder tempat file Excel diupload
nama_file_hasil = 'Hasil_Merge_Pajak.xlsx'      # Nama file output yang diinginkan
# ==========================================

print("Memulai proses merge data...")

# Mencari semua file Excel di dalam folder
semua_file = glob.glob(folder_sumber)

# Menyiapkan penampung untuk data gabungan
kumpulan_faktur = []
kumpulan_detail = []

# Memproses setiap file satu per satu
for file in semua_file:
    nama_file = os.path.basename(file)
    print(f"-> Memproses file: {nama_file}")

    # ---------------------------------------------------------
    # 1. PROSES SHEET "Faktur" (Ambil Kolom A sampai R)
    # ---------------------------------------------------------
    try:
        df_faktur = pd.read_excel(file, sheet_name='Faktur', usecols="A:R")
        nama_kolom_A_faktur = df_faktur.columns[0] # Mengambil nama kolom pertama (Baris)

        # Mencari baris yang mengandung kata "END" (mengabaikan huruf besar/kecil/spasi)
        batas_end_faktur = df_faktur[df_faktur[nama_kolom_A_faktur].astype(str).str.strip().str.lower() == 'end'].index

        # Jika ketemu "END", potong data sampai sebelum baris "END" tersebut
        if len(batas_end_faktur) > 0:
            df_faktur = df_faktur.iloc[:batas_end_faktur[0]]

        # Sisipkan kolom Nama File dan Periode Check di paling kiri (posisi 0)
        # Urutan insert dibalik agar hasilnya: [Periode Check] [Nama File] [Data Asli...]
        df_faktur.insert(0, 'Nama File', nama_file)
        df_faktur.insert(0, 'Periode Check', periode_check)

        kumpulan_faktur.append(df_faktur)
    except Exception as e:
        print(f"   [!] Gagal memproses sheet Faktur di {nama_file}. Error: {e}")

    # ---------------------------------------------------------
    # 2. PROSES SHEET "DetailFaktur" (Ambil Kolom A sampai N)
    # ---------------------------------------------------------
    try:
        df_detail = pd.read_excel(file, sheet_name='DetailFaktur', usecols="A:N")
        nama_kolom_A_detail = df_detail.columns[0]

        batas_end_detail = df_detail[df_detail[nama_kolom_A_detail].astype(str).str.strip().str.lower() == 'end'].index

        if len(batas_end_detail) > 0:
            df_detail = df_detail.iloc[:batas_end_detail[0]]

        df_detail.insert(0, 'Nama File', nama_file)
        df_detail.insert(0, 'Periode Check', periode_check)

        kumpulan_detail.append(df_detail)
    except Exception as e:
        print(f"   [!] Gagal memproses sheet DetailFaktur di {nama_file}. Error: {e}")

# ==========================================
# GABUNGKAN DAN SIMPAN DATA
# ==========================================
print("\nMenggabungkan semua data...")

# Menggabungkan semua data yang ada di list (jika list tidak kosong)
if kumpulan_faktur and kumpulan_detail:
    df_faktur_final = pd.concat(kumpulan_faktur, ignore_index=True)
    df_detail_final = pd.concat(kumpulan_detail, ignore_index=True)

    # Menyimpan ke dalam satu file Excel baru dengan 2 Sheet
    with pd.ExcelWriter(nama_file_hasil) as writer:
        df_faktur_final.to_excel(writer, sheet_name='Faktur', index=False)
        df_detail_final.to_excel(writer, sheet_name='DetailFaktur', index=False)

    print(f"\n✅ SUKSES! File berhasil digabungkan menjadi '{nama_file_hasil}'.")
    print("Silakan cek di panel folder sebelah kiri (Refresh jika perlu) untuk mengunduhnya.")
else:
    print("\n❌ GAGAL: Tidak ada data yang berhasil diproses. Pastikan folder dan file sudah benar.")
