import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Laporan Meteorologi Lengkap", layout="wide")

st.title("🌦️ Auto-Rekapitulasi Data Meteorologi")
st.write("Aplikasi ini akan memilah file CSV mentah (seperti `job_3069.csv`) menjadi berbagai Sheet Excel yang rapi per parameter (Suhu, RH, QFE, Tekanan Uap, dll).")

# --- RUMUS MENGHITUNG TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2)  # Sesuai format Excel Anda (ratusan)

# --- DAFTAR PARAMETER YANG AKAN DIBUATKAN SHEET-NYA ---
# Format: {"Nama Kolom di CSV": "NAMA SHEET DI EXCEL"}
parameter_mapping = {
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR',
    'temp_drybulb_c_tttttt': 'SUHU BOLA KERING',
    'temp_wetbulb_c': 'SUHU BOLA BASAH',
    'relative_humidity_pc': 'KELEMBAPAN (RH)',
    'temp_dewpoint_c_tdtdtd': 'TITIK EMBUN (DEW)',
    'pressure_qfe_mb_derived': 'TEKANAN QFE',
    'pressure_qff_mb_derived': 'TEKANAN QFF'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

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
        
        # 3. Hitung Parameter Tambahan (Tekanan Uap Air)
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
        
        # Filter pilihan bulan jika datanya berbulan-bulan (agar tidak kepanjangan saat preview)
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        daftar_bulan_unik = sorted(df_raw['Bulan_Tahun'].unique())
        
        bulan_dipilih = st.selectbox("Pilih Bulan untuk Preview & Download Excel:", daftar_bulan_unik)
        
        # Filter data hanya untuk bulan yang dipilih
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
        
        # ==========================================
        # PROSES PEMBUATAN EXCEL MULTI-SHEET
        # ==========================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            workbook  = writer.book
            
            # Format Desain Excel
            format_judul = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#8DB4E2', 'border': 1})
            format_data = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            format_tanggal = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#EBF1DE', 'border': 1})
            
            st.write("### 📑 Preview Data (Sampel)")
            
            # 4. Looping untuk membuat Sheet per Parameter
            for kolom_csv, nama_sheet in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    
                    # Buat Pivot Table per parameter
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    # Rapikan struktur 24 jam
                    for hour in range(24):
                        if hour not in pivot.columns:
                            pivot[hour] = np.nan
                    pivot = pivot[list(range(24))] 
                    pivot.columns = [f"{float(h)}" for h in range(24)]
                    
                    # Rapikan struktur tanggal
                    semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')
                    pivot = pivot.reindex(semua_tanggal)
                    
                    # Rata-rata khusus: Jika Tekanan Uap dibagi 10, lainnya biasa
                    if kolom_csv == 'Tekanan_Uap_x10':
                        pivot['R A T A   2'] = pivot.iloc[:, 0:24].mean(axis=1) / 10
                    else:
                        pivot['R A T A   2'] = pivot.iloc[:, 0:24].mean(axis=1)
                        
                    # Tampilkan sedikit preview di layar aplikasi (menggunakan Expander)
                    with st.expander(f"Preview Tabel: {nama_sheet}"):
                        st.dataframe(pivot.style.format(precision=1), use_container_width=True)
                    
                    # 5. Tulis ke dalam Sheet Excel
                    pivot.to_excel(writer, sheet_name=nama_sheet[:31]) # Excel membatasi nama sheet max 31 karakter
                    worksheet = writer.sheets[nama_sheet[:31]]
                    
                    # 6. Percantik Sheet-nya
                    worksheet.set_column('A:A', 8, format_tanggal)
                    worksheet.set_column('B:Y', 7, format_data)
                    worksheet.set_column('Z:Z', 12, format_data)
                    worksheet.write(0, 0, "NO.", format_judul)
                    for col_num, value in enumerate(pivot.columns.values):
                        worksheet.write(0, col_num + 1, value, format_judul)

        # ==========================================
        st.success("✅ File Excel Rapi dengan Multi-Parameter siap diunduh!")
        
        st.download_button(
            label=f"Unduh Excel Laporan {bulan_dipilih} (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"LAPORAN_LENGKAP_BMKG_{bulan_dipilih}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
