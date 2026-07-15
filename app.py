import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Generator ME-45 Harian", layout="wide")

st.title("📄 Auto-Generate Form ME-45 BMKG (Harian)")
st.write("Aplikasi ini akan mengubah data CSV Anda menjadi format form **ME-45 Harian** yang persis dengan standar PDF BMKG (Jam menurun ke bawah, Parameter menyamping).")

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda (Misal: job_3071.csv)", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. Baca dan Siapkan Data
        df = pd.read_csv(uploaded_file)
        df['data_timestamp'] = pd.to_datetime(df['data_timestamp'])
        df['Tahun'] = df['data_timestamp'].dt.year
        df['Bulan'] = df['data_timestamp'].dt.month
        df['Tanggal'] = df['data_timestamp'].dt.day
        df['Jam_GMT'] = df['data_timestamp'].dt.hour # Asumsi data sudah dalam waktu yang sesuai
        
        # Buat dropdown untuk memilih Hari/Tanggal spesifik
        df['Tanggal_Lengkap'] = df['data_timestamp'].dt.strftime('%d %B %Y')
        daftar_tanggal = sorted(df['Tanggal_Lengkap'].unique())
        
        tanggal_pilih = st.selectbox("Pilih Tanggal Observasi untuk dicetak ke ME-45:", daftar_tanggal)
        df_hari_ini = df[df['Tanggal_Lengkap'] == tanggal_pilih].copy()
        
        # 2. Susun Struktur Tabel Persis ME-45
        # Kita buat DataFrame kosong 24 Jam (00 - 23)
        me45_df = pd.DataFrame({'GMT': range(24)})
        
        # Gabungkan data aktual ke kerangka 24 jam
        df_hari_ini = df_hari_ini.set_index('Jam_GMT')
        
        # Mapping parameter dari CSV ke format Header ME-45
        me45_df['N dd ff VV'] = "" # Kolom sandi yang butuh input visual
        me45_df['ww w1 w2'] = ""
        
        # TTT (Suhu Udara)
        me45_df['TTT'] = me45_df['GMT'].map(df_hari_ini['temp_drybulb_c_tttttt'] if 'temp_drybulb_c_tttttt' in df_hari_ini.columns else {})
        # TdTdTd (Titik Embun)
        me45_df['TdTdTd'] = me45_df['GMT'].map(df_hari_ini['temp_dewpoint_c_tdtdtd'] if 'temp_dewpoint_c_tdtdtd' in df_hari_ini.columns else {})
        # TwTwTw (Suhu Bola Basah)
        me45_df['TwTwTw'] = me45_df['GMT'].map(df_hari_ini['temp_wetbulb_c'] if 'temp_wetbulb_c' in df_hari_ini.columns else {})
        # QFF
        me45_df['QFF'] = me45_df['GMT'].map(df_hari_ini['pressure_qff_mb_derived'] if 'pressure_qff_mb_derived' in df_hari_ini.columns else {})
        # QFE
        me45_df['QFE'] = me45_df['GMT'].map(df_hari_ini['pressure_qfe_mb_derived'] if 'pressure_qfe_mb_derived' in df_hari_ini.columns else {})
        # Tx (Suhu Max)
        me45_df['TxTxTx'] = me45_df['GMT'].map(df_hari_ini['temp_max_c_txtxtx'] if 'temp_max_c_txtxtx' in df_hari_ini.columns else {})
        # Tn (Suhu Min)
        me45_df['TnTnTn'] = me45_df['GMT'].map(df_hari_ini['temp_min_c_tntntn'] if 'temp_min_c_tntntn' in df_hari_ini.columns else {})
        
        # Perbaiki format GMT menjadi 00, 01, 02 dst
        me45_df['GMT'] = me45_df['GMT'].apply(lambda x: f"{x:02d}")
        
        # --- TAMPILAN PREVIEW ---
        st.write("---")
        st.write(f"### 📑 Preview Laporan ME-45 ({tanggal_pilih})")
        st.dataframe(me45_df.style.format(precision=1, na_rep=""), use_container_width=True)
        
        # --- EXPORT KE EXCEL DENGAN DESAIN MIRIP PDF ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Buat format styling persis form cetak BMKG
            format_header = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'border': 1, 'bg_color': '#D9D9D9', 'text_wrap': True
            })
            format_isi = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'
            })
            format_jam = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
            })
            
            me45_df.to_excel(writer, sheet_name='ME45_HARIAN', index=False)
            ws = writer.sheets['ME45_HARIAN']
            
            # Set Lebar Kolom agar rapi saat diprint
            ws.set_column('A:A', 6, format_jam)     # GMT
            ws.set_column('B:C', 12, format_isi)    # Sandi Visual
            ws.set_column('D:J', 9, format_isi)     # Parameter Angka
            
            # Tulis Header Spesifik
            for col_num, value in enumerate(me45_df.columns):
                ws.write(0, col_num, value, format_header)
                
            # Tambahkan info Hari/Tanggal di atas tabel (seperti di PDF)
            ws.write(0, 0, "GMT", format_header)

        st.success(f"✅ Form ME-45 untuk {tanggal_pilih} berhasil dibuat!")
        
        st.download_button(
            label=f"📥 Unduh Excel ME-45 ({tanggal_pilih})",
            data=buffer.getvalue(),
            file_name=f"ME45_{tanggal_pilih.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
