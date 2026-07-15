import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Laporan Meteorologi & ME-45", layout="wide")

st.title("🌦️ Auto-Rekapitulasi Data & Generator ME-45")
st.write("Aplikasi ini menghasilkan dua output sekaligus: **Laporan Bulanan Per Jam (Multi-Sheet)** dan **Form ME-45 Harian (Terisi Penuh dengan Sandi Cuaca)**.")

# --- RUMUS MENGHITUNG TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh):
        return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2)

# =====================================================================
# --- MAPPING PARAMETER UNTUK LAPORAN PER JAM ---
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
    'wind_dir_deg_dd': 'ARAH ANGIN (DD)',
    'visibility_vv': 'JARAK PANDANG (VV)',
    'cloud_cover_oktas_m': 'TUTUPAN AWAN (OKTAS)',
    'rainfall_24h_rrrr': 'CURAH HUJAN 24J'
}

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG (Gunakan job_3137.csv agar ME-45 penuh)", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. BACA & PARSING DATA
        df_raw = pd.read_csv(uploaded_file)
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        # Hitung Tekanan Uap
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), 
                axis=1
            )
        
        # Opsi Pilihan di Layar
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
        # PROSES 1: MEMBUAT EXCEL LAPORAN PER JAM (MULTI-SHEET)
        # =====================================================================
        buffer_jam = io.BytesIO()
        with pd.ExcelWriter(buffer_jam, engine='xlsxwriter') as writer_jam:
            wb_jam = writer_jam.book
            fmt_judul = wb_jam.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#8DB4E2', 'border': 1})
            fmt_data = wb_jam.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            fmt_tgl = wb_jam.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#EBF1DE', 'border': 1})
            
            for kolom_csv, nama_sheet in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    for hour in range(24):
                        if hour not in pivot.columns: pivot[hour] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot.columns = [f"{float(h)}" for h in range(24)]
                    
                    semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')
                    pivot = pivot.reindex(semua_tanggal)
                    
                    if kolom_csv == 'Tekanan_Uap_x10':
                        pivot['R A T A   2'] = pivot.iloc[:, 0:24].mean(axis=1) / 10
                    else:
                        pivot['R A T A   2'] = pivot.iloc[:, 0:24].mean(axis=1)
                        
                    safe_sheet_name = nama_sheet[:31]
                    pivot.to_excel(writer_jam, sheet_name=safe_sheet_name)
                    ws_jam = writer_jam.sheets[safe_sheet_name]
                    
                    ws_jam.set_column('A:A', 8, fmt_tgl)
                    ws_jam.set_column('B:Y', 7, fmt_data)
                    ws_jam.set_column('Z:Z', 12, fmt_data)
                    ws_jam.write(0, 0, "NO.", fmt_judul)
                    for col_num, value in enumerate(pivot.columns.values):
                        ws_jam.write(0, col_num + 1, value, fmt_judul)

        # =====================================================================
        # PROSES 2: MEMBUAT EXCEL FORM ME-45 HARIAN
        # =====================================================================
        df_hari_ini = df_raw[df_raw['Tanggal_Lengkap'] == tanggal_dipilih].copy()
        
        me45_df = pd.DataFrame({'GMT': range(24)})
        df_hari_ini = df_hari_ini.set_index('Jam')
        
        # --- Fungsi Pembuat Sandi N dd ff VV Otomatis ---
        def generate_sandi_angin(jam):
            if jam not in df_hari_ini.index: return ""
            row = df_hari_ini.loc[jam]
            
            # N (Awan Oktas)
            N = str(int(row['cloud_cover_oktas_m'])) if 'cloud_cover_oktas_m' in row and pd.notna(row['cloud_cover_oktas_m']) else "/"
            
            # dd (Arah Angin puluhan derajat)
            if 'wind_dir_deg_dd' in row and pd.notna(row['wind_dir_deg_dd']):
                dd = f"{int(round(row['wind_dir_deg_dd'] / 10)):02d}"
            else:
                dd = "//"
                
            # ff (Kecepatan knot)
            ff = f"{int(row['wind_speed_ff']):02d}" if 'wind_speed_ff' in row and pd.notna(row['wind_speed_ff']) else "//"
            
            # VV (Jarak Pandang konversi ke kode SYNOP)
            VV = "//"
            if 'visibility_vv' in row and pd.notna(row['visibility_vv']):
                km = float(row['visibility_vv'])
                if km < 5.1: VV = f"{int(km*10):02d}"
                elif km < 31: VV = f"{int(km)+50:02d}"
                elif km < 71: VV = f"{int(km/5)+80:02d}"
                else: VV = "99"
                
            return f"{N} {dd} {ff} {VV}"

        # --- Fungsi Pembuat Sandi ww w1 w2 ---
        def generate_sandi_cuaca(jam):
            if jam not in df_hari_ini.index: return ""
            row = df_hari_ini.loc[jam]
            
            ww = f"{int(row['present_weather_ww']):02d}" if 'present_weather_ww' in row and pd.notna(row['present_weather_ww']) else "//"
            w1 = str(int(row['past_weather_w1'])) if 'past_weather_w1' in row and pd.notna(row['past_weather_w1']) else "/"
            w2 = str(int(row['past_weather_w2'])) if 'past_weather_w2' in row and pd.notna(row['past_weather_w2']) else "/"
            
            if ww == "//" and w1 == "/" and w2 == "/": return ""
            return f"{ww} {w1}{w2}"

        # Memasukkan fungsi ke dalam dataframe
        me45_df['N dd ff VV'] = me45_df['GMT'].apply(generate_sandi_angin)
        me45_df['ww w1 w2'] = me45_df['GMT'].apply(generate_sandi_cuaca)
        
        # Parameter Angka ME-45
        me45_df['TTT'] = me45_df['GMT'].map(df_hari_ini['temp_drybulb_c_tttttt'] if 'temp_drybulb_c_tttttt' in df_hari_ini.columns else {})
        me45_df['TdTdTd'] = me45_df['GMT'].map(df_hari_ini['temp_dewpoint_c_tdtdtd'] if 'temp_dewpoint_c_tdtdtd' in df_hari_ini.columns else {})
        me45_df['TwTwTw'] = me45_df['GMT'].map(df_hari_ini['temp_wetbulb_c'] if 'temp_wetbulb_c' in df_hari_ini.columns else {})
        me45_df['QFF'] = me45_df['GMT'].map(df_hari_ini['pressure_qff_mb_derived'] if 'pressure_qff_mb_derived' in df_hari_ini.columns else {})
        me45_df['QFE'] = me45_df['GMT'].map(df_hari_ini['pressure_qfe_mb_derived'] if 'pressure_qfe_mb_derived' in df_hari_ini.columns else {})
        me45_df['TxTxTx'] = me45_df['GMT'].map(df_hari_ini['temp_max_c_txtxtx'] if 'temp_max_c_txtxtx' in df_hari_ini.columns else {})
        me45_df['TnTnTn'] = me45_df['GMT'].map(df_hari_ini['temp_min_c_tntntn'] if 'temp_min_c_tntntn' in df_hari_ini.columns else {})
        
        me45_df['GMT'] = me45_df['GMT'].apply(lambda x: f"{x:02d}")
        
        buffer_me45 = io.BytesIO()
        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer_me45:
            wb_me45 = writer_me45.book
            fmt_hdr = wb_me45.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9', 'text_wrap': True})
            fmt_isi = wb_me45.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
            fmt_sandi = wb_me45.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_jam = wb_me45.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            me45_df.to_excel(writer_me45, sheet_name='ME45_HARIAN', index=False)
            ws_me45 = writer_me45.sheets['ME45_HARIAN']
            
            ws_me45.set_column('A:A', 6, fmt_jam)
            ws_me45.set_column('B:C', 13, fmt_sandi)
            ws_me45.set_column('D:J', 9, fmt_isi)
            
            for col_num, value in enumerate(me45_df.columns):
                ws_me45.write(0, col_num, value, fmt_hdr)
            ws_me45.write(0, 0, "GMT", fmt_hdr)

        # =====================================================================
        # TAMPILAN ANTARMUKA DOWNLOAD BERDAMPINGAN
        # =====================================================================
        st.write("---")
        st.write(f"### 📑 Preview Laporan ME-45 ({tanggal_dipilih})")
        st.dataframe(me45_df.style.format(precision=1, na_rep=""), use_container_width=True)

        st.success("✅ Seluruh Data Berhasil Diproses! Silakan unduh format yang Anda butuhkan:")
        
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            st.download_button(
                label=f"📥 1. Unduh Laporan Per Jam ({bulan_dipilih})",
                data=buffer_jam.getvalue(),
                file_name=f"LAPORAN_JAM_BMKG_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.info("Berisi Laporan Excel Multi-Sheet format matriks 24 Jam selama 1 bulan penuh.")
            
        with dl_col2:
            st.download_button(
                label=f"📥 2. Unduh Form ME-45 ({tanggal_dipilih})",
                data=buffer_me45.getvalue(),
                file_name=f"ME45_{tanggal_dipilih.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.warning("Formulir Harian. Kini Sandi Awan, Cuaca, Angin, dan Jarak Pandang terisi Otomatis!")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
