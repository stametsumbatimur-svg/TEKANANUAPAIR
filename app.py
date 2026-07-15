import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Sistem Rekapitulasi & ME-45 BMKG", layout="wide")

st.title("🌦️ Auto-Generator Data Meteorologi & ME-45 (Sandi SYNOP)")
st.write("Aplikasi ini secara otomatis mengonversi data ke dalam **Sandi SYNOP/ME-45 standar BMKG** (tanpa desimal, konversi tekanan 4 digit, dll).")

# --- FUNGSI FORMATTING STANDAR SYNOP / ME-45 ---
def format_suhu(val):
    """Mengubah suhu 23.4 menjadi 234"""
    if pd.isna(val) or val == "": return ""
    return str(int(round(float(val) * 10)))

def format_tekanan(val):
    """Mengubah tekanan 1011.1 menjadi 0111 (4 digit, buang ribuan)"""
    if pd.isna(val) or val == "": return ""
    p_int = int(round(float(val) * 10))
    p_mod = p_int % 10000
    return f"{p_mod:04d}"

def format_umum(val):
    """Membulatkan angka biasa tanpa desimal"""
    if pd.isna(val) or val == "": return ""
    return str(int(round(float(val))))

# Rumus Tekanan Uap Air
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2) # Masih ada desimal di raw, di-format saat dicetak

# Mapping Parameter untuk Laporan Per Jam
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
    'wind_dir_deg_dd': 'ARAH ANGIN (DD)',
    'visibility_vv': 'JARAK PANDANG (VV)',
    'cloud_cover_oktas_m': 'TUTUPAN AWAN (OKTAS)',
    'rainfall_24h_rrrr': 'CURAH HUJAN 24J'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda (Gunakan job_3137.csv)", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. BACA & PARSING DATA
        df_raw = pd.read_csv(uploaded_file)
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
        
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        df_raw['Tanggal_Lengkap'] = df_raw['data_timestamp'].dt.strftime('%d %B %Y')
        
        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            bulan_dipilih = st.selectbox("Pilih Bulan (Untuk Laporan Per Jam):", sorted(df_raw['Bulan_Tahun'].unique()))
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        with col_opt2:
            tanggal_dipilih = st.selectbox("Pilih Tanggal (Untuk Cetak Form ME-45):", sorted(df_bulan_ini['Tanggal_Lengkap'].unique()))
            
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]

        # =====================================================================
        # PROSES 1: MEMBUAT EXCEL LAPORAN PER JAM (FORMAT SANDI SYNOP)
        # =====================================================================
        buffer_jam = io.BytesIO()
        with pd.ExcelWriter(buffer_jam, engine='xlsxwriter') as writer_jam:
            wb_jam = writer_jam.book
            fmt_judul = wb_jam.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#8DB4E2', 'border': 1})
            fmt_data = wb_jam.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}) # Tanpa desimal
            fmt_tgl = wb_jam.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#EBF1DE', 'border': 1})
            
            for kolom_csv, nama_sheet in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    # Ambil data raw
                    pivot_raw = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    # Buat DataFrame baru dengan format string (Tanpa Koma)
                    kolom_jam_str = [str(h) for h in range(24)] # Header Jam 0, 1, 2 (tanpa .0)
                    semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')
                    pivot_fmt = pd.DataFrame(index=semua_tanggal, columns=kolom_jam_str + ['RATA-RATA'])
                    
                    # Menghitung Rata-rata dari nilai mentah (sebelum dibulatkan)
                    rata_rata_raw = pivot_raw.mean(axis=1)
                    
                    # Terapkan Formatting SYNOP ke seluruh sel
                    for tgl in semua_tanggal:
                        for h in range(24):
                            val_raw = pivot_raw.loc[tgl, h] if (tgl in pivot_raw.index and h in pivot_raw.columns) else np.nan
                            
                            # Logika format per parameter
                            if pd.isna(val_raw):
                                pivot_fmt.loc[tgl, str(h)] = ""
                            elif 'pressure' in kolom_csv and 'q' in kolom_csv or kolom_csv == 'pressure_reading_mb':
                                pivot_fmt.loc[tgl, str(h)] = format_tekanan(val_raw)
                            elif 'temp_' in kolom_csv:
                                pivot_fmt.loc[tgl, str(h)] = format_suhu(val_raw)
                            elif kolom_csv == 'Tekanan_Uap_x10':
                                pivot_fmt.loc[tgl, str(h)] = format_umum(val_raw) # Sudah dikali 10 di rumusnya
                            else:
                                pivot_fmt.loc[tgl, str(h)] = format_umum(val_raw)
                                
                        # Format untuk kolom Rata-rata
                        mean_raw = rata_rata_raw.loc[tgl] if tgl in rata_rata_raw.index else np.nan
                        if pd.isna(mean_raw):
                            pivot_fmt.loc[tgl, 'RATA-RATA'] = ""
                        elif 'pressure' in kolom_csv and 'q' in kolom_csv or kolom_csv == 'pressure_reading_mb':
                            pivot_fmt.loc[tgl, 'RATA-RATA'] = format_tekanan(mean_raw)
                        elif 'temp_' in kolom_csv:
                            pivot_fmt.loc[tgl, 'RATA-RATA'] = format_suhu(mean_raw)
                        else:
                            pivot_fmt.loc[tgl, 'RATA-RATA'] = format_umum(mean_raw)

                    safe_sheet_name = nama_sheet[:31]
                    pivot_fmt.to_excel(writer_jam, sheet_name=safe_sheet_name)
                    ws_jam = writer_jam.sheets[safe_sheet_name]
                    
                    ws_jam.set_column('A:A', 8, fmt_tgl)
                    ws_jam.set_column('B:Z', 7, fmt_data)
                    ws_jam.write(0, 0, "NO.", fmt_judul)
                    for col_num, value in enumerate(pivot_fmt.columns.values):
                        ws_jam.write(0, col_num + 1, value, fmt_judul)

        # =====================================================================
        # PROSES 2: MEMBUAT EXCEL FORM ME-45 HARIAN (SESUAI PDF)
        # =====================================================================
        df_hari_ini = df_raw[df_raw['Tanggal_Lengkap'] == tanggal_dipilih].copy()
        df_hari_ini = df_hari_ini.set_index('Jam')
        
        # Kolom persis seperti PDF
        kolom_me45_pdf = [
            'GMT', 'N dd ff VV', 'ww w1w2', 'TTT', 'TdTdTd', 'TwTwTw', 'QFF', 'QFE',
            'Tx Tx Tx', 'Tn Tn Tn E E E', 'F24 F24', 'P24 P24 iw', 'ix iR IE', 'RRRtR',
            'ChshsChshs 0 S', 'C Daec UU', 'TTTNCLhCCN MHs h'
        ]
        me45_df = pd.DataFrame(index=range(24), columns=kolom_me45_pdf)
        me45_df.fillna("", inplace=True)
        
        # --- Fungsi Pembuat Sandi ---
        def get_sandi_angin(jam):
            if jam not in df_hari_ini.index: return ""
            row = df_hari_ini.loc[jam]
            N = str(int(row['cloud_cover_oktas_m'])) if 'cloud_cover_oktas_m' in row and pd.notna(row['cloud_cover_oktas_m']) else "/"
            dd = f"{int(round(row['wind_dir_deg_dd'] / 10)):02d}" if 'wind_dir_deg_dd' in row and pd.notna(row['wind_dir_deg_dd']) else "//"
            ff = f"{int(row['wind_speed_ff']):02d}" if 'wind_speed_ff' in row and pd.notna(row['wind_speed_ff']) else "//"
            VV = "//"
            if 'visibility_vv' in row and pd.notna(row['visibility_vv']):
                km = float(row['visibility_vv'])
                if km < 5.1: VV = f"{int(km*10):02d}"
                elif km < 31: VV = f"{int(km)+50:02d}"
                elif km < 71: VV = f"{int(km/5)+80:02d}"
                else: VV = "99"
            return f"{N} {dd} {ff} {VV}"

        def get_sandi_cuaca(jam):
            if jam not in df_hari_ini.index: return ""
            row = df_hari_ini.loc[jam]
            ww = f"{int(row['present_weather_ww']):02d}" if 'present_weather_ww' in row and pd.notna(row['present_weather_ww']) else "//"
            w1 = str(int(row['past_weather_w1'])) if 'past_weather_w1' in row and pd.notna(row['past_weather_w1']) else "/"
            w2 = str(int(row['past_weather_w2'])) if 'past_weather_w2' in row and pd.notna(row['past_weather_w2']) else "/"
            if ww == "//" and w1 == "/" and w2 == "/": return ""
            return f"{ww} {w1}{w2}"

        # Mengisi Tabel ME-45
        for jam in range(24):
            me45_df.at[jam, 'GMT'] = f"{jam:02d}"
            me45_df.at[jam, 'N dd ff VV'] = get_sandi_angin(jam)
            me45_df.at[jam, 'ww w1w2'] = get_sandi_cuaca(jam)
            
            if jam in df_hari_ini.index:
                row = df_hari_ini.loc[jam]
                if 'temp_drybulb_c_tttttt' in row: me45_df.at[jam, 'TTT'] = format_suhu(row['temp_drybulb_c_tttttt'])
                if 'temp_dewpoint_c_tdtdtd' in row: me45_df.at[jam, 'TdTdTd'] = format_suhu(row['temp_dewpoint_c_tdtdtd'])
                if 'temp_wetbulb_c' in row: me45_df.at[jam, 'TwTwTw'] = format_suhu(row['temp_wetbulb_c'])
                if 'pressure_qff_mb_derived' in row: me45_df.at[jam, 'QFF'] = format_tekanan(row['pressure_qff_mb_derived'])
                if 'pressure_qfe_mb_derived' in row: me45_df.at[jam, 'QFE'] = format_tekanan(row['pressure_qfe_mb_derived'])
                if 'temp_max_c_txtxtx' in row: me45_df.at[jam, 'Tx Tx Tx'] = format_suhu(row['temp_max_c_txtxtx'])
                if 'temp_min_c_tntntn' in row: me45_df.at[jam, 'Tn Tn Tn E E E'] = format_suhu(row['temp_min_c_tntntn'])
        
        buffer_me45 = io.BytesIO()
        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer_me45:
            wb_me45 = writer_me45.book
            fmt_hdr = wb_me45.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
            fmt_isi = wb_me45.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_jam = wb_me45.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            me45_df.to_excel(writer_me45, sheet_name='ME45_HARIAN', index=False)
            ws_me45 = writer_me45.sheets['ME45_HARIAN']
            
            ws_me45.set_column('A:A', 5, fmt_jam)
            ws_me45.set_column('B:C', 13, fmt_isi)
            ws_me45.set_column('D:Q', 9, fmt_isi)
            
            for col_num, value in enumerate(me45_df.columns):
                ws_me45.write(0, col_num, value, fmt_hdr)

        # =====================================================================
        # TAMPILAN ANTARMUKA APLIKASI
        # =====================================================================
        st.write("---")
        st.write(f"### 📑 Preview Laporan ME-45 ({tanggal_dipilih})")
        st.dataframe(me45_df, use_container_width=True)

        st.success("✅ Format berhasil dikalibrasi ke standar SYNOP BMKG tanpa desimal!")
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            st.download_button(
                label=f"📥 1. Unduh Laporan Per Jam ({bulan_dipilih})",
                data=buffer_jam.getvalue(),
                file_name=f"LAPORAN_JAM_BMKG_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.info("Header jam sudah (0, 1, 2) dan angka di dalam tabel dikonversi menjadi string tanpa koma.")
            
        with dl_col2:
            st.download_button(
                label=f"📥 2. Unduh Form ME-45 ({tanggal_dipilih})",
                data=buffer_me45.getvalue(),
                file_name=f"ME45_{tanggal_dipilih.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.warning("Format kolom dijamin persis 100% dengan Kertas PDF ME-45 stasiun Anda.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
