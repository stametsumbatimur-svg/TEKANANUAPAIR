import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Generator Matriks Fklim BMKG", layout="wide")

st.title("📊 Auto-Generate Excel Matriks Fklim (1 Sheet)")
st.write("Aplikasi ini telah dikalibrasi sesuai standar Fklim terbaru: Formatting SYNOP pada jam tertentu, penambahan summary MAX/MIN bulanan, dan penghilangan parameter yang tidak diperlukan.")

# --- RUMUS TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2)

# --- FUNGSI FORMATTING KHUSUS ---
def fmt_q_str(val):
    """Untuk QFF/QFE 0-23: 1011.1 -> 0111"""
    if pd.isna(val): return ""
    try: return f"{(int(round(float(val) * 10)) % 10000):04d}"
    except: return ""

def fmt_t_str(val):
    """Untuk Suhu 0-23: 23.4 -> 234"""
    if pd.isna(val): return ""
    try: return str(int(round(float(val) * 10)))
    except: return ""

def fmt_int_str(val):
    """Untuk RH/Uap Air: Bulat tanpa koma"""
    if pd.isna(val): return ""
    try: return str(int(round(float(val))))
    except: return ""

# --- PARAMETER MAPPING (Suhu Max/Min & Arah Angin Dihilangkan) ---
parameter_mapping = {
    'pressure_qff_mb_derived': 'QFF RATA-2 HARIAN',
    'pressure_qfe_mb_derived': 'QFE RATA-2 HARIAN',
    'temp_drybulb_c_tttttt': 'SUHU UDARA RATA-2 HARIAN',
    'relative_humidity_pc': 'KELEMBABAN UDARA RATA-2 HARIAN',
    'wind_speed_ff': 'KECEPATAN ANGIN RATA-2 HARIAN',
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR RATA-2 HARIAN'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file)
        
        # Hapus nilai error sensor
        NA_VALUES = [9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A']
        df_raw.replace(NA_VALUES, np.nan, inplace=True)
        
        # Parsing Waktu
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
        
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        
        st.markdown("---")
        bulan_dipilih = st.selectbox("Pilih Bulan untuk di-Generate:", sorted(df_raw['Bulan_Tahun'].unique()))
        
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        nama_bulan = calendar.month_name[bulan_val].upper()
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
        semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')

        # =====================================================================
        # PROSES PEMBUATAN EXCEL
        # =====================================================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('MATRIKS FKLIM')
            
            fmt_teks = wb.add_format({'bold': True, 'align': 'left'})
            fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            # Format Data
            fmt_str = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}) # Teks Tanpa Koma
            fmt_float = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'}) # Desimal Koma
            fmt_summary = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#EBF1DE'})
            
            ws.set_column('A:A', 8)
            ws.set_column('B:Y', 5)
            ws.set_column('Z:Z', 12)
            ws.set_column('AA:AC', 8)
            
            start_row = 0
            
            for kolom_csv, judul_param in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    
                    df_bulan_ini.loc[:, kolom_csv] = pd.to_numeric(df_bulan_ini[kolom_csv], errors='coerce')
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    for h in range(24):
                        if h not in pivot.columns: pivot[h] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot = pivot.reindex(semua_tanggal)
                    
                    rata_harian = pivot.mean(axis=1)
                    if kolom_csv == 'Tekanan_Uap_x10': 
                        rata_harian = rata_harian / 10 # Khusus Uap Air Rata-2 harian dibagi 10
                    
                    ws.write(start_row, 1, "BULAN", fmt_teks)
                    ws.write(start_row, 3, nama_bulan, fmt_teks)
                    ws.write(start_row, 4, str(tahun_val), fmt_teks)
                    ws.write(start_row + 1, 12, judul_param, fmt_judul)
                    
                    ws.write(start_row + 2, 0, "NO.", fmt_header)
                    for i in range(24): ws.write(start_row + 2, i + 1, str(i), fmt_header)
                    ws.write(start_row + 2, 25, "R A T A   2", fmt_header)
                    ws.write(start_row + 2, 26, "23 00", fmt_header)
                    ws.write(start_row + 2, 27, "05 00", fmt_header)
                    ws.write(start_row + 2, 28, "10 00", fmt_header)
                    
                    row_idx = start_row + 3
                    
                    # ----------------- LOOP TANGGAL 1-31 -----------------
                    for tgl in semua_tanggal:
                        ws.write(row_idx, 0, tgl, fmt_header)
                        
                        # 1. BAGIAN JAM 0-23 (Penerapan Sandi)
                        for h in range(24):
                            val = pivot.loc[tgl, h]
                            if pd.isna(val): ws.write(row_idx, h + 1, "", fmt_str)
                            else:
                                if 'QFF' in judul_param or 'QFE' in judul_param: ws.write(row_idx, h + 1, fmt_q_str(val), fmt_str)
                                elif 'SUHU UDARA' in judul_param: ws.write(row_idx, h + 1, fmt_t_str(val), fmt_str)
                                elif 'KELEMBABAN' in judul_param or 'UAP AIR' in judul_param: ws.write(row_idx, h + 1, fmt_int_str(val), fmt_str)
                                elif 'KECEPATAN ANGIN' in judul_param:
                                    if float(val) == 0: ws.write(row_idx, h + 1, "", fmt_str)
                                    else: ws.write(row_idx, h + 1, val, fmt_float)
                                else: ws.write(row_idx, h + 1, val, fmt_float)
                        
                        # 2. BAGIAN RATA-RATA HARIAN (Selalu pakai Koma)
                        mean_val = rata_harian.loc[tgl]
                        if pd.isna(mean_val): ws.write(row_idx, 25, "", fmt_str)
                        else:
                            if 'KECEPATAN ANGIN' in judul_param and float(mean_val) == 0: ws.write(row_idx, 25, "", fmt_str)
                            else: ws.write(row_idx, 25, mean_val, fmt_float)
                        
                        # 3. BAGIAN JAM SPESIFIK 23, 05, 10
                        for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                            val_s = pivot.loc[tgl, h_spec]
                            if pd.isna(val_s): ws.write(row_idx, c_idx, "", fmt_str)
                            else:
                                if 'QFF' in judul_param or 'QFE' in judul_param or 'SUHU UDARA' in judul_param: 
                                    ws.write(row_idx, c_idx, val_s, fmt_float) # Ada koma
                                elif 'KELEMBABAN' in judul_param or 'UAP AIR' in judul_param: 
                                    ws.write(row_idx, c_idx, fmt_int_str(val_s), fmt_str) # Tanpa Koma
                                elif 'KECEPATAN ANGIN' in judul_param:
                                    if float(val_s) == 0: ws.write(row_idx, c_idx, "", fmt_str)
                                    else: ws.write(row_idx, c_idx, val_s, fmt_float)
                                else: ws.write(row_idx, c_idx, val_s, fmt_float)
                        
                        row_idx += 1
                        
                    # ----------------- LOOP SUMMARY MAX MIN RATA-RATA -----------------
                    summary_funcs = [("MAX", pivot.max()), ("MIN", pivot.min()), ("RATA-RATA", pivot.mean())]
                    
                    for row_name, series_data in summary_funcs:
                        ws.write(row_idx, 0, row_name, fmt_summary)
                        
                        # Summary 0-23
                        for h in range(24):
                            val = series_data[h]
                            if pd.isna(val): ws.write(row_idx, h + 1, "", fmt_str)
                            else:
                                if 'QFF' in judul_param or 'QFE' in judul_param: ws.write(row_idx, h + 1, fmt_q_str(val), fmt_str)
                                elif 'SUHU UDARA' in judul_param: ws.write(row_idx, h + 1, fmt_t_str(val), fmt_str)
                                elif 'KELEMBABAN' in judul_param or 'UAP AIR' in judul_param: ws.write(row_idx, h + 1, fmt_int_str(val), fmt_str)
                                elif 'KECEPATAN ANGIN' in judul_param:
                                    if float(val) == 0: ws.write(row_idx, h + 1, "", fmt_str)
                                    else: ws.write(row_idx, h + 1, val, fmt_float)
                                else: ws.write(row_idx, h + 1, val, fmt_float)
                        
                        # Summary RATA 2
                        if row_name == "MAX": v_rata = rata_harian.max()
                        elif row_name == "MIN": v_rata = rata_harian.min()
                        else: v_rata = rata_harian.mean()
                        
                        if pd.isna(v_rata): ws.write(row_idx, 25, "", fmt_str)
                        else:
                            if 'KECEPATAN ANGIN' in judul_param and float(v_rata) == 0: ws.write(row_idx, 25, "", fmt_str)
                            else: ws.write(row_idx, 25, v_rata, fmt_float)
                            
                        # Summary Jam Spesifik 23, 05, 10
                        for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                            val_s = series_data.get(h_spec, np.nan)
                            if pd.isna(val_s): ws.write(row_idx, c_idx, "", fmt_str)
                            else:
                                if 'QFF' in judul_param or 'QFE' in judul_param or 'SUHU UDARA' in judul_param: 
                                    ws.write(row_idx, c_idx, val_s, fmt_float)
                                elif 'KELEMBABAN' in judul_param or 'UAP AIR' in judul_param: 
                                    ws.write(row_idx, c_idx, fmt_int_str(val_s), fmt_str)
                                elif 'KECEPATAN ANGIN' in judul_param:
                                    if float(val_s) == 0: ws.write(row_idx, c_idx, "", fmt_str)
                                    else: ws.write(row_idx, c_idx, val_s, fmt_float)
                                else: ws.write(row_idx, c_idx, val_s, fmt_float)
                                
                        row_idx += 1
                        
                    start_row = row_idx + 3 # Jarak antar tabel parameter

        st.success("✅ Seluruh Parameter berhasil disatukan dalam 1 Sheet memanjang ke bawah!")
        st.download_button(
            label=f"📥 Unduh Laporan Fklim Lengkap ({bulan_dipilih})",
            data=buffer.getvalue(),
            file_name=f"FKLIM_{bulan_dipilih}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
