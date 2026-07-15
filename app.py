import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Sistem ME-45 & Rekap Cuaca", layout="wide")

st.title("🌦️ Auto-Rekapitulasi Data Meteorologi & ME-45")
st.write("Aplikasi ini mengolah data CSV mentah menjadi 2 jenis output: **Laporan Matriks Per Jam** dan **Laporan Rekap Harian (Format ME-45)**.")

# --- RUMUS MENGHITUNG TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2)

# =====================================================================
# --- DAFTAR PARAMETER (MATRIKS JAM) ---
# =====================================================================
parameter_mapping = {
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR',
    'temp_drybulb_c_tttttt': 'SUHU BOLA KERING',
    'temp_wetbulb_c': 'SUHU BOLA BASAH',
    'relative_humidity_pc': 'KELEMBAPAN (RH)',
    'temp_dewpoint_c_tdtdtd': 'TITIK EMBUN (DEW)',
    'pressure_qfe_mb_derived': 'TEKANAN QFE',
    'pressure_qff_mb_derived': 'TEKANAN QFF',
    'pressure_reading_mb': 'PRESSURE READING',
    'temp_max_c_txtxtx': 'SUHU MAKSIMUM',
    'temp_min_c_tntntn': 'SUHU MINIMUM',
    'wind_speed_ff': 'KECEPATAN ANGIN (FF)',
    'wind_dir_deg_dd': 'ARAH ANGIN (DD)'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG (Misal: job_3071.csv)", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. Baca Data
        df_raw = pd.read_csv(uploaded_file)
        
        # 2. Parsing Waktu
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        # 3. Hitung Parameter Tambahan
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
        
        # Filter pilihan bulan
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        daftar_bulan_unik = sorted(df_raw['Bulan_Tahun'].unique())
        
        bulan_dipilih = st.selectbox("Pilih Bulan untuk Preview & Download Excel:", daftar_bulan_unik)
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
        semua_tanggal = pd.Index(range(1, jml_hari + 1), name='TANGGAL')
        
        # ==========================================
        # FUNGSI A: BUAT EXCEL MATRIKS PER JAM (AWAL)
        # ==========================================
        buffer_jam = io.BytesIO()
        with pd.ExcelWriter(buffer_jam, engine='xlsxwriter') as writer:
            workbook = writer.book
            format_judul = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#8DB4E2', 'border': 1})
            format_data = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            format_tanggal = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#EBF1DE', 'border': 1})
            
            for kolom_csv, nama_sheet in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    for hour in range(24):
                        if hour not in pivot.columns: pivot[hour] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot.columns = [f"{float(h)}" for h in range(24)]
                    pivot = pivot.reindex(semua_tanggal)
                    
                    if kolom_csv == 'Tekanan_Uap_x10':
                        pivot['RATA-RATA'] = pivot.iloc[:, 0:24].mean(axis=1) / 10
                    else:
                        pivot['RATA-RATA'] = pivot.iloc[:, 0:24].mean(axis=1)
                        
                    safe_sheet_name = nama_sheet[:31]
                    pivot.to_excel(writer, sheet_name=safe_sheet_name)
                    ws = writer.sheets[safe_sheet_name]
                    ws.set_column('A:A', 8, format_tanggal)
                    ws.set_column('B:Y', 7, format_data)
                    ws.set_column('Z:Z', 12, format_data)
                    ws.write(0, 0, "NO.", format_judul)
                    for col_num, value in enumerate(pivot.columns.values):
                        ws.write(0, col_num + 1, value, format_judul)
                        
        # ==========================================
        # FUNGSI B: BUAT EXCEL REKAP HARIAN (ME-45)
        # ==========================================
        # Agregasi data per tanggal untuk mendapatkan Max, Min, dan Mean
        df_me45 = pd.DataFrame(index=semua_tanggal)
        
        if 'temp_max_c_txtxtx' in df_bulan_ini.columns:
            df_me45['Tx (Suhu Max)'] = df_bulan_ini.groupby('Tanggal')['temp_max_c_txtxtx'].max()
        if 'temp_min_c_tntntn' in df_bulan_ini.columns:
            df_me45['Tn (Suhu Min)'] = df_bulan_ini.groupby('Tanggal')['temp_min_c_tntntn'].min()
        if 'temp_drybulb_c_tttttt' in df_bulan_ini.columns:
            df_me45['T (Suhu Rata-rata)'] = df_bulan_ini.groupby('Tanggal')['temp_drybulb_c_tttttt'].mean()
        if 'relative_humidity_pc' in df_bulan_ini.columns:
            df_me45['RH (%) Rata-rata'] = df_bulan_ini.groupby('Tanggal')['relative_humidity_pc'].mean()
        if 'pressure_qfe_mb_derived' in df_bulan_ini.columns:
            df_me45['QFE (mb) Rata-rata'] = df_bulan_ini.groupby('Tanggal')['pressure_qfe_mb_derived'].mean()
        if 'pressure_qff_mb_derived' in df_bulan_ini.columns:
            df_me45['QFF (mb) Rata-rata'] = df_bulan_ini.groupby('Tanggal')['pressure_qff_mb_derived'].mean()
        if 'wind_speed_ff' in df_bulan_ini.columns:
            df_me45['Angin Max (Knot)'] = df_bulan_ini.groupby('Tanggal')['wind_speed_ff'].max()
            df_me45['Angin Rata-rata (Knot)'] = df_bulan_ini.groupby('Tanggal')['wind_speed_ff'].mean()
        
        buffer_me45 = io.BytesIO()
        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer2:
            workbook2 = writer2.book
            format_header = workbook2.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFC000', 'border': 1, 'text_wrap': True})
            format_isi = workbook2.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            format_tgl = workbook2.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D9D9', 'border': 1})
            
            df_me45.to_excel(writer2, sheet_name='REKAP HARIAN ME-45')
            ws2 = writer2.sheets['REKAP HARIAN ME-45']
            
            # Styling Excel ME45
            ws2.set_column('A:A', 10, format_tgl)
            ws2.set_column('B:Z', 15, format_isi)
            ws2.set_row(0, 30) # Tinggi baris judul
            
            ws2.write(0, 0, "TANGGAL", format_header)
            for col_num, value in enumerate(df_me45.columns):
                ws2.write(0, col_num + 1, value, format_header)
        
        # ==========================================
        # TAMPILAN ANTARMUKA APLIKASI
        # ==========================================
        st.write("---")
        st.subheader("📑 Preview Rekap Harian ME-45 (Agregasi Max/Min/Rata-rata)")
        st.dataframe(df_me45.style.format(precision=1), use_container_width=True)
        
        st.write("---")
        st.success(f"✅ Seluruh Data Bulan {bulan_dipilih} Berhasil Diproses!")
        st.write("Silakan pilih format laporan yang ingin Anda unduh di bawah ini:")
        
        # Tombol Download Berjejer (Side-by-side)
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 1. Unduh Laporan Per Jam (Multi-Sheet)",
                data=buffer_jam.getvalue(),
                file_name=f"LAPORAN_JAM_BMKG_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.info("Berisi 12 Sheet Excel yang mendetail dari Jam 00.00 - 23.00 untuk setiap parameter.")
            
        with col2:
            st.download_button(
                label="📥 2. Unduh Rekap ME-45 (Tabel Harian)",
                data=buffer_me45.getvalue(),
                file_name=f"REKAP_ME45_BMKG_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.warning("Berisi 1 Sheet Excel yang merangkum data per tanggal (Maksimum, Minimum, dan Rata-rata Harian).")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
