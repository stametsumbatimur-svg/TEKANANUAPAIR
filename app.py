import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Generator Matriks Fklim BMKG", layout="wide")

st.title("📊 Auto-Generate Excel Matriks Fklim (1 Sheet)")
st.write("Format 100% Standar Fklim. Dilengkapi pewarnaan cerdas pada kolom RATA-RATA dan Kesimpulan Akhir (MAX/MIN Bulanan) untuk memudahkan pembacaan.")

# --- RUMUS TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2) # Dikali 10 sesuai standar Fklim

# --- FUNGSI FORMATTING KHUSUS SYNOP ---
def fmt_q_str(val):
    if pd.isna(val): return ""
    try: return f"{(int(round(float(val) * 10)) % 10000):04d}" # 1011.1 -> 0111
    except: return ""

def fmt_t_str(val):
    if pd.isna(val): return ""
    try: return str(int(round(float(val) * 10))) # 23.4 -> 234
    except: return ""

def fmt_int_str(val):
    if pd.isna(val): return ""
    try: return str(int(round(float(val)))) # Bulat tanpa koma
    except: return ""

# --- PARAMETER MAPPING (Sesuai Koreksi Poin 4: Arah angin, Tmax, Tmin dihapus) ---
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
        
        # Cleansing nilai error
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
        # LOGIKA PERMAINAN WARNA & FORMAT EXCEL
        # =====================================================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('MATRIKS FKLIM')
            
            # Format Dasar
            fmt_teks = wb.add_format({'bold': True, 'align': 'left'})
            fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
            
            # Format Data Kolom Biasa (0-23 & Jam Spesifik)
            fmt_str_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_float_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            
            # Format Kolom RATA 2 (Warna Biru Muda agar jelas arahnya ke bawah)
            fmt_str_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
            fmt_float_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0', 'bg_color': '#DCE6F1'})
            
            # Format Summary Bawah (Warna Kuning/Oranye) - 1 NILAI SAJA
            fmt_summary_judul = wb.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'})
            fmt_summary_kosong = wb.add_format({'border': 1, 'bg_color': '#FFF2CC'})
            fmt_summary_final = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4', 'num_format': '0.0'})
            fmt_summary_final_str = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4'})
            
            # Lebar Kolom
            ws.set_column('A:A', 6)
            ws.set_column('B:Y', 5)
            ws.set_column('Z:Z', 13)
            ws.set_column('AA:AC', 8)
            
            start_row = 0
            
            # Helper untuk menentukan format sel secara dinamis
            def get_cell(val, param, col_type):
                if pd.isna(val): 
                    if col_type == 'RATA2': return "", fmt_str_rata2
                    return "", fmt_str_biasa
                
                # KECEPATAN ANGIN (Aturan Poin 5: 0 Kosong, tanpa koma)
                if 'KECEPATAN ANGIN' in param:
                    v = int(round(float(val)))
                    res = "" if v == 0 else str(v)
                    if col_type == 'RATA2': return res, fmt_str_rata2
                    return res, fmt_str_biasa
                
                # QFF / QFE / SUHU
                if 'QFF' in param or 'QFE' in param or 'SUHU' in param:
                    if col_type == '0-23':
                        if 'QFF' in param or 'QFE' in param: return fmt_q_str(val), fmt_str_biasa # 0111
                        if 'SUHU' in param: return fmt_t_str(val), fmt_str_biasa # 234
                    elif col_type == 'RATA2':
                        return round(float(val), 1), fmt_float_rata2 # 1011.1
                    else: # JAM SPESIFIK (23, 05, 10)
                        return round(float(val), 1), fmt_float_biasa # 1011.1
                
                # RH / UAP AIR (Aturan Poin 6 & 7: Semua tanpa koma kecuali Rata2)
                if 'KELEMBABAN' in param or 'UAP AIR' in param:
                    if col_type == 'RATA2':
                        return round(float(val), 1), fmt_float_rata2
                    else: # 0-23 dan Jam Spesifik
                        return fmt_int_str(val), fmt_str_biasa
                
                return val, fmt_str_biasa

            # ----------------- LOOP GENERATE PARAMETER -----------------
            for kolom_csv, judul_param in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    
                    df_bulan_ini.loc[:, kolom_csv] = pd.to_numeric(df_bulan_ini[kolom_csv], errors='coerce')
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    for h in range(24):
                        if h not in pivot.columns: pivot[h] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot = pivot.reindex(semua_tanggal)
                    
                    rata_harian = pivot.mean(axis=1)
                    if kolom_csv == 'Tekanan_Uap_x10': rata_harian = rata_harian / 10 # Penyesuaian uap air
                    
                    # Tulis Judul & Header
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
                    
                    # TULIS DATA TANGGAL 1-31
                    for tgl in semua_tanggal:
                        ws.write(row_idx, 0, tgl, fmt_header)
                        
                        # Data 0-23
                        for h in range(24):
                            val, fmt = get_cell(pivot.loc[tgl, h], judul_param, '0-23')
                            ws.write(row_idx, h + 1, val, fmt)
                        
                        # Data RATA 2 (Warna Biru Muda)
                        val_rata, fmt_rata = get_cell(rata_harian.loc[tgl], judul_param, 'RATA2')
                        ws.write(row_idx, 25, val_rata, fmt_rata)
                        
                        # Data Jam Spesifik (23, 05, 10)
                        for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                            val_s, fmt_s = get_cell(pivot.loc[tgl, h_spec], judul_param, 'SPEC')
                            ws.write(row_idx, c_idx, val_s, fmt_s)
                        
                        row_idx += 1
                        
                    # -------------------------------------------------------------
                    # 1 NILAI SAJA UNTUK MAX, MIN, RATA-RATA (Menghilangkan Bingung)
                    # -------------------------------------------------------------
                    # Nilai Absolute Sebulan Penuh
                    abs_max = pivot.max().max()
                    abs_min = pivot.min().min()
                    total_rata = rata_harian.mean()
                    
                    summary_labels = [("MAXIMUM BULAN INI", abs_max), 
                                      ("MINIMUM BULAN INI", abs_min), 
                                      ("TOTAL RATA-RATA", total_rata)]
                    
                    for label, final_val in summary_labels:
                        # Gabung (Merge) kolom 0 sampai 24 untuk judul agar rapi dan tidak ada angka bertumpuk
                        ws.merge_range(row_idx, 0, row_idx, 24, label, fmt_summary_judul)
                        
                        # 1 Nilai di letakkan tepat di bawah kolom RATA 2 (Warna Oranye)
                        if pd.isna(final_val):
                            ws.write(row_idx, 25, "", fmt_summary_final_str)
                        else:
                            if 'KECEPATAN ANGIN' in judul_param:
                                v_ang = int(round(float(final_val)))
                                ws.write(row_idx, 25, "" if v_ang == 0 else str(v_ang), fmt_summary_final_str)
                            else:
                                ws.write(row_idx, 25, round(float(final_val), 1), fmt_summary_final)
                        
                        # Kosongkan kolom 23, 05, 10 di baris summary agar bersih
                        ws.write(row_idx, 26, "", fmt_summary_kosong)
                        ws.write(row_idx, 27, "", fmt_summary_kosong)
                        ws.write(row_idx, 28, "", fmt_summary_kosong)
                        
                        row_idx += 1
                        
                    start_row = row_idx + 3 # Spasi sebelum tabel parameter berikutnya

        st.success("✅ Seluruh Parameter berhasil disatukan dalam 1 Sheet memanjang ke bawah! Data siap cetak.")
        st.download_button(
            label=f"📥 Unduh Laporan Fklim Lengkap ({bulan_dipilih})",
            data=buffer.getvalue(),
            file_name=f"FKLIM_{bulan_dipilih}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
