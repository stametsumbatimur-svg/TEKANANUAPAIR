import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Auto-Generate Laporan Tekanan Uap", layout="wide")

st.title("🌪️ Pembuat Laporan Tekanan Uap Air Otomatis")
st.write("Unggah file CSV raw data. Aplikasi otomatis menghitung Tekanan Uap dan menyusunnya ke dalam format tabel Laporan Excel yang RAPI dan SIAP CETAK.")

# --- RUMUS METEOROLOGI ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
        
    # Rumus Magnus-Tetens untuk Tekanan Uap Jenuh (e_s) dalam hPa
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    
    # Tekanan Uap Aktual (e)
    e_actual = (rh / 100.0) * es
    
    return round(e_actual * 10, 2)

# --- UPLOAD FILE ---
uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file)
        
        if 'data_timestamp' not in df_raw.columns or 'temp_drybulb_c_tttttt' not in df_raw.columns or 'relative_humidity_pc' not in df_raw.columns:
            st.error("File CSV tidak valid. Pastikan formatnya sesuai tarikan sistem.")
        else:
            df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
            df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
            df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
            
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
            
            pivot_table = df_raw.pivot_table(
                index='Tanggal', columns='Jam', values='Tekanan_Uap_x10', aggfunc='first'
            )
            
            for hour in range(24):
                if hour not in pivot_table.columns:
                    pivot_table[hour] = np.nan
            pivot_table = pivot_table[list(range(24))] 
            
            # Format penamaan kolom jam (0.0 - 23.0)
            pivot_table.columns = [f"{float(h)}" for h in range(24)]
            semua_tanggal = pd.Index(range(1, 32), name='NO.')
            pivot_table = pivot_table.reindex(semua_tanggal)
            
            pivot_table['R A T A   2'] = pivot_table.iloc[:, 0:24].mean(axis=1) / 10
            
            st.success("✅ Data berhasil diproses!")
            st.dataframe(pivot_table.style.format(precision=2), use_container_width=True)
            
            # ==========================================
            # PROSES PEMBUATAN EXCEL YANG RAPI & MENARIK
            # ==========================================
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Tulis dataframe ke Excel
                pivot_table.to_excel(writer, sheet_name='TEKANAN UAP AIR')
                
                # Ambil objek workbook dan worksheet
                workbook  = writer.book
                worksheet = writer.sheets['TEKANAN UAP AIR']
                
                # 1. Buat Format Desain Sel
                format_judul = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#8DB4E2', # Warna Biru Muda
                    'border': 1
                })
                
                format_data = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1,
                    'num_format': '0.00' # Format 2 angka di belakang koma
                })
                
                format_tanggal = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'bg_color': '#EBF1DE', # Warna Hijau Muda
                    'border': 1
                })

                # 2. Atur Lebar Kolom
                worksheet.set_column('A:A', 8, format_tanggal) # Kolom NO. (Tanggal)
                worksheet.set_column('B:Y', 7, format_data)    # Kolom Jam 00-23
                worksheet.set_column('Z:Z', 12, format_data)   # Kolom Rata-rata
                
                # 3. Timpa Header dengan Desain Judul
                worksheet.write(0, 0, "NO.", format_judul)
                for col_num, value in enumerate(pivot_table.columns.values):
                    worksheet.write(0, col_num + 1, value, format_judul)
                    
            # ==========================================
            
            st.markdown("### 📥 Unduh Hasil Akhir")
            st.write("Klik tombol di bawah ini untuk mengunduh laporan dalam format **.xlsx** yang sudah diformat rapi.")
            
            st.download_button(
                label="Unduh Tabel Excel Siap Cetak (.xlsx)",
                data=buffer.getvalue(),
                file_name="LAPORAN_TEKANAN_UAP_AIR_RAPI.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
