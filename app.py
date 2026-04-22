import streamlit as st
import pandas as pd
import io
import datetime
# Konfigurasi Halaman
st.set_page_config(page_title="Merge Excel Pajak", page_icon="🧩")

st.title("🧩 Aplikasi Merge Data Pajak Ekspor")
st.markdown("Upload beberapa file **BAGAN IMPORT** (Excel) sekaligus, lalu gabungkan Sheet `Faktur` dan `DetailFaktur` hanya dengan satu klik.")

# ==========================================
# INPUT DARI USER
# ==========================================
st.subheader("1. Pengaturan Data")
periode_check = st.text_input("Masukkan Periode Check (Contoh: 31/1/2026)", value="31/1/2026")

st.subheader("2. Upload File Excel")
# Allow multiple files
uploaded_files = st.file_uploader("Pilih file Excel (Bisa pilih lebih dari satu file sekaligus)", type=['xlsx'], accept_multiple_files=True)

# ==========================================
# PROSES TOMBOL DIKLIK
# ==========================================
if st.button("🚀 Proses & Merge Data"):
    if not uploaded_files:
        st.warning("⚠️ Silakan upload minimal 1 file Excel terlebih dahulu!")
    elif not periode_check:
        st.warning("⚠️ Kolom Periode Check tidak boleh kosong!")
    else:
        with st.spinner("Sedang memproses data... Mohon tunggu."):
            kumpulan_faktur = []
            kumpulan_detail = []
            
            # Memproses setiap file yang diupload
            for file in uploaded_files:
                nama_file = file.name
                
                # -- PROSES SHEET FAKTUR --
                try:
                    df_faktur = pd.read_excel(file, sheet_name='Faktur', usecols="A:R")
                    nama_kolom_A_faktur = df_faktur.columns[0]
                    
                    batas_end_faktur = df_faktur[df_faktur[nama_kolom_A_faktur].astype(str).str.strip().str.lower() == 'end'].index
                    if len(batas_end_faktur) > 0:
                        df_faktur = df_faktur.iloc[:batas_end_faktur[0]]
                        
                    df_faktur.insert(0, 'Nama File', nama_file)
                    df_faktur.insert(0, 'Periode Check', periode_check)
                    kumpulan_faktur.append(df_faktur)
                except Exception as e:
                    st.error(f"Gagal memproses sheet 'Faktur' di file {nama_file}. Error: {e}")

                # -- PROSES SHEET DETAIL FAKTUR --
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
                    st.error(f"Gagal memproses sheet 'DetailFaktur' di file {nama_file}. Error: {e}")
            
            # -- GABUNGKAN SEMUA --
            if kumpulan_faktur and kumpulan_detail:
                df_faktur_final = pd.concat(kumpulan_faktur, ignore_index=True)
                df_detail_final = pd.concat(kumpulan_detail, ignore_index=True)
                
                # Tulis ke dalam memory buffer (agar tidak perlu save ke hardisk lokal)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_faktur_final.to_excel(writer, sheet_name='Faktur', index=False)
                    df_detail_final.to_excel(writer, sheet_name='DetailFaktur', index=False)
                
                output.seek(0)
                
                st.success("✅ File berhasil digabungkan!")
                
                # Menampilkan tombol Download
                st.download_button(
                    label="⬇️ Download Hasil Merge (Excel)",
                    data=output,
                    file_name="Hasil_Merge_Pajak.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
