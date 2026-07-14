import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Auto-Generate Laporan Tekanan Uap", layout="wide")

st.title("🌪️ Laporan Tekanan Uap Air Otomatis (Multi-Bulan)")
st.write("Unggah CSV mentah. Jika data berisi berbulan-bulan, sistem akan **memisahkannya ke dalam Sheet berbeda per bulan** secara otomatis tanpa tertimpa.")

def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2)

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file)
        
        if 'data_timestamp' not in df_raw.columns or 'temp_drybulb_c_tttttt' not in df_raw.columns or 'relative_humidity_pc' not in df_raw.columns:
            st.error("File CSV tidak valid.")
        else:
            # Parse waktu
            df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
            df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
            df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
            df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
            df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
            
            # Hitung data utama
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
            
            # Siapkan sistem pembuatan file Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                workbook  = writer.book
                
                # Format Tampilan
                format_judul = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#8DB4E2', 'border': 1})
                format_data = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00'})
                format_tanggal = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#EBF1DE', 'border': 1})
                
                # Cek daftar bulan yang ada di dalam data secara unik
                df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
                daftar_bulan_unik = sorted(df_raw['Bulan_Tahun'].unique())
                
                # Loop dan pisahkan data per bulan
                for bulan_tahun in daftar_bulan_unik:
                    tahun_val = int(bulan_tahun.split('-')[0])
                    bulan_val = int(bulan_tahun.split('-')[1])
                    
                    nama_bulan = calendar.month_name[bulan_val]
                    nama_sheet = f"{nama_bulan[:3].upper()} {tahun_val}" # Contoh nama sheet: "JUN 2026"
                    
                    # Filter data HANYA untuk bulan dan tahun yang sedang diproses
                    df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_tahun]
                    
                    # Pivot data untuk bulan ini
                    pivot_table = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values='Tekanan_Uap_x10', aggfunc='first')
                    
                    # Rapikan ukuran tabel
                    for hour in range(24):
                        if hour not in pivot_table.columns:
                            pivot_table[hour] = np.nan
                    pivot_table = pivot_table[list(range(24))] 
                    pivot_table.columns = [f"{float(h)}" for h in range(24)]
                    
                    # Pastikan tanggal sesuai dengan jumlah hari dalam bulan tersebut (bisa 28, 30, atau 31)
                    jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
                    semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')
                    pivot_table = pivot_table.reindex(semua_tanggal)
                    
                    pivot_table['R A T A   2'] = pivot_table.iloc[:, 0:24].mean(axis=1) / 10
                    
                    # Tulis hasil bulan ini ke Sheet tersendiri
                    pivot_table.to_excel(writer, sheet_name=nama_sheet)
                    worksheet = writer.sheets[nama_sheet]
                    
                    # Terapkan gaya lebar kolom dan warna pada Sheet tersebut
                    worksheet.set_column('A:A', 8, format_tanggal)
                    worksheet.set_column('B:Y', 7, format_data)
                    worksheet.set_column('Z:Z', 12, format_data)
                    worksheet.write(0, 0, "NO.", format_judul)
                    for col_num, value in enumerate(pivot_table.columns.values):
                        worksheet.write(0, col_num + 1, value, format_judul)
                        
            st.success("✅ File berisi multi-bulan berhasil dibuat, dipisah per Sheet, dan siap diunduh!")
            
            st.download_button(
                label="Unduh Rekap Lengkap (.xlsx)",
                data=buffer.getvalue(),
                file_name="LAPORAN_TEKANAN_UAP_LENGKAP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
