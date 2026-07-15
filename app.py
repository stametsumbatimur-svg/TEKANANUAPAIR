import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Generator Matriks Fklim BMKG", layout="wide")

st.title("📊 Auto-Generate Excel Matriks Fklim (1 Sheet)")
st.write("Aplikasi ini akan mengubah data CSV Anda menjadi format Laporan Fklim yang ditumpuk ke bawah dalam 1 Sheet, lengkap dengan ekstraksi jam 23, 05, dan 10 UTC.")

# --- RUMUS TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    # Dikalikan 10 sesuai standar Fklim untuk Tekanan Uap
    return round(e_actual * 10, 2)

# --- DAFTAR PARAMETER YANG AKAN DISUSUN KE BAWAH ---
parameter_mapping = {
    'pressure_qff_mb_derived': 'QFF RATA-2 HARIAN',
    'pressure_qfe_mb_derived': 'QFE RATA-2 HARIAN',
    'temp_drybulb_c_tttttt': 'SUHU UDARA RATA-2 HARIAN',
    'temp_max_c_txtxtx': 'SUHU MAKSIMUM RATA-2 HARIAN',
    'temp_min_c_tntntn': 'SUHU MINIMUM RATA-2 HARIAN',
    'relative_humidity_pc': 'KELEMBABAN UDARA RATA-2 HARIAN',
    'wind_dir_deg_dd': 'ARAH ANGIN RATA-2 HARIAN',
    'wind_speed_ff': 'KECEPATAN ANGIN RATA-2 HARIAN',
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR RATA-2 HARIAN'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. BACA & BERSIHKAN DATA
        df_raw = pd.read_csv(uploaded_file)
        
        # Hapus nilai error sensor
        NA_VALUES = [9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A']
        df_raw.replace(NA_VALUES, np.nan, inplace=True)
        
        # 2. PARSING WAKTU
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
        # PROSES PEMBUATAN EXCEL (1 SHEET, DITUMPUK KE BAWAH)
        # =====================================================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('MATRIKS FKLIM')
            
            # Format Tampilan
            fmt_teks = wb.add_format({'bold': True, 'align': 'left'})
            fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_data = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            fmt_no = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            # Atur Lebar Kolom
            ws.set_column('A:A', 6)         # NO.
            ws.set_column('B:Y', 6.5)       # Jam 0-23
            ws.set_column('Z:Z', 13)        # R A T A   2
            ws.set_column('AA:AC', 8)       # 23 00, 05 00, 10 00
            
            start_row = 0
            
            # Looping setiap parameter untuk dicetak bertingkat ke bawah
            for kolom_csv, judul_param in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    
                    # Buat Pivot Table per parameter
                    # (Memaksa data menjadi numerik agar bisa dicari rata-ratanya)
                    df_bulan_ini.loc[:, kolom_csv] = pd.to_numeric(df_bulan_ini[kolom_csv], errors='coerce')
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    # Pastikan kolom jam 0-23 ada
                    for h in range(24):
                        if h not in pivot.columns: pivot[h] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot = pivot.reindex(semua_tanggal)
                    
                    # Tambahkan kolom Rata-rata dan Jam Spesifik
                    rata_harian = pivot.mean(axis=1)
                    
                    # --- MULAI MENULIS KE EXCEL ---
                    
                    # Baris 1: Keterangan Bulan & Tahun
                    ws.write(start_row, 1, "BULAN", fmt_teks)
                    ws.write(start_row, 3, nama_bulan, fmt_teks)
                    ws.write(start_row, 4, str(tahun_val), fmt_teks)
                    
                    # Baris 2: Judul Parameter (Di tengah)
                    ws.write(start_row + 1, 12, judul_param, fmt_judul)
                    
                    # Baris 3: Header Kolom
                    ws.write(start_row + 2, 0, "NO.", fmt_header)
                    for i in range(24):
                        ws.write(start_row + 2, i + 1, str(i), fmt_header)
                    ws.write(start_row + 2, 25, "R A T A   2", fmt_header)
                    ws.write(start_row + 2, 26, "23 00", fmt_header)
                    ws.write(start_row + 2, 27, "05 00", fmt_header)
                    ws.write(start_row + 2, 28, "10 00", fmt_header)
                    
                    # Baris 4 sd Selesai: Isi Data
                    row_idx = start_row + 3
                    for tgl in semua_tanggal:
                        ws.write(row_idx, 0, tgl, fmt_no) # Tulis Tanggal
                        
                        # Tulis Data 0-23
                        for h in range(24):
                            val = pivot.loc[tgl, h]
                            if pd.isna(val):
                                ws.write(row_idx, h + 1, "", fmt_data)
                            else:
                                ws.write(row_idx, h + 1, val, fmt_data)
                        
                        # Tulis RATA 2
                        mean_val = rata_harian.loc[tgl]
                        if pd.isna(mean_val): ws.write(row_idx, 25, "", fmt_data)
                        else: ws.write(row_idx, 25, mean_val, fmt_data)
                        
                        # Tulis Jam Spesifik (23, 05, 10)
                        val_23 = pivot.loc[tgl, 23]
                        val_05 = pivot.loc[tgl, 5]
                        val_10 = pivot.loc[tgl, 10]
                        
                        ws.write(row_idx, 26, val_23 if pd.notna(val_23) else "", fmt_data)
                        ws.write(row_idx, 27, val_05 if pd.notna(val_05) else "", fmt_data)
                        ws.write(row_idx, 28, val_10 if pd.notna(val_10) else "", fmt_data)
                        
                        row_idx += 1
                    
                    # Tambahkan jarak 3 baris kosong sebelum tabel parameter berikutnya
                    start_row = row_idx + 3

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
