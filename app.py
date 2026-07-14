import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Auto-Generate Laporan Tekanan Uap", layout="wide")

st.title("🌪️ Pembuat Laporan Tekanan Uap Air Otomatis")
st.write("Unggah file CSV raw data (seperti format `job_3062.csv`). Aplikasi akan otomatis menghitung Tekanan Uap dan menyusunnya ke dalam format tabel laporan 24 jam yang bisa diunduh langsung sebagai Excel.")

# --- RUMUS METEOROLOGI ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
        
    # Rumus Magnus-Tetens untuk Tekanan Uap Jenuh (e_s) dalam hPa
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    
    # Tekanan Uap Aktual (e)
    e_actual = (rh / 100.0) * es
    
    # Dikali 10 sesuai format tabel laporan Excel Anda
    return round(e_actual * 10, 2)

# --- UPLOAD FILE ---
uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. Baca data mentah
        df_raw = pd.read_csv(uploaded_file)
        
        # 2. Validasi apakah kolom yang dibutuhkan ada
        if 'data_timestamp' not in df_raw.columns or 'temp_drybulb_c_tttttt' not in df_raw.columns or 'relative_humidity_pc' not in df_raw.columns:
            st.error("File CSV tidak valid. Pastikan file memiliki kolom 'data_timestamp', 'temp_drybulb_c_tttttt', dan 'relative_humidity_pc'.")
        else:
            # 3. Proses Data Waktu (Ambil Tanggal dan Jam)
            df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
            df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
            df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
            
            # 4. Hitung Tekanan Uap Air
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
            
            # 5. Susun menjadi tabel Matriks (Baris = Tanggal, Kolom = Jam)
            pivot_table = df_raw.pivot_table(
                index='Tanggal', 
                columns='Jam', 
                values='Tekanan_Uap_x10', 
                aggfunc='first'
            )
            
            # 6. Rapikan format tabel agar persis seperti Excel template
            for hour in range(24):
                if hour not in pivot_table.columns:
                    pivot_table[hour] = np.nan
            pivot_table = pivot_table[list(range(24))] 
            
            pivot_table.columns = [f"{float(h)}" for h in range(24)]
            semua_tanggal = pd.Index(range(1, 32), name='NO.')
            pivot_table = pivot_table.reindex(semua_tanggal)
            
            # Tambahkan kolom Rata-rata Harian
            pivot_table['R A T A   2'] = pivot_table.iloc[:, 0:24].mean(axis=1) / 10
            
            st.success("✅ Data berhasil diproses dan disusun secara otomatis!")
            
            st.subheader("📊 Preview Laporan Tekanan Uap Air")
            st.dataframe(pivot_table.style.format(precision=2), use_container_width=True)
            
            # --- FITUR DOWNLOAD EXCEL ---
            st.markdown("### 📥 Unduh Hasil Akhir")
            st.write("Klik tombol di bawah ini untuk mengunduh laporan dalam format **.xlsx**.")
            
            # Membuat file Excel di dalam memori
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pivot_table.to_excel(writer, sheet_name='TEKANAN UAP AIR')
            
            # Tombol unduh untuk Excel
            st.download_button(
                label="Unduh Tabel Format Laporan (.xlsx)",
                data=buffer.getvalue(),
                file_name="LAPORAN_TEKANAN_UAP_AIR_READY.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
else:
    st.info("Silakan unggah file CSV Anda untuk melihat keajaibannya.")
