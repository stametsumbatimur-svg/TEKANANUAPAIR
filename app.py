import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Laporan Matriks Data BMKG", layout="wide")

st.title("📊 Auto-Generate Laporan Excel Matriks 24 Jam")
st.write("Aplikasi ini akan mengubah data CSV Anda menjadi format Excel Matriks 24 Jam, lengkap dengan kolom **R A T A   2** dan ekstraksi jam spesifik (**23:00, 05:00, 10:00**).")

uploaded_file = st.file_uploader("Unggah file CSV Raw Data BMKG Anda", type=["csv"])

if uploaded_file is not None:
    try:
        # 1. BACA DATA
        df_raw = pd.read_csv(uploaded_file)
        
        # Sapu bersih data error standar (opsional untuk menjaga kebersihan data)
        NA_VALUES = [9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A']
        df_raw.replace(NA_VALUES, np.nan, inplace=True)
        
        # 2. PARSING WAKTU
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        
        # 3. KONTROL MENU DI LAYAR
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            bulan_dipilih = st.selectbox("1. Pilih Bulan Laporan:", sorted(df_raw['Bulan_Tahun'].unique()))
            
        with col2:
            # Biarkan user memilih parameter apa yang mau dicetak (Dinamis dari header CSV)
            kolom_abaikan = ['station_name', 'data_timestamp', 'created_at', 'updated_at', 'Tahun', 'Bulan_Angka', 'Tanggal', 'Jam', 'Bulan_Tahun']
            pilihan_param = [col for col in df_raw.columns if col not in kolom_abaikan]
            param_dipilih = st.selectbox("2. Pilih Parameter yang ingin direkap (misal: pressure_qff_mb_derived):", pilihan_param)
            
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]

        # 4. PROSES PEMBUATAN TABEL MATRIKS
        if st.button("🚀 Buat Laporan Excel"):
            
            # Pivot tabel jam 0-23
            # Pastikan datanya numerik agar bisa dihitung rata-ratanya
            df_bulan_ini[param_dipilih] = pd.to_numeric(df_bulan_ini[param_dipilih], errors='coerce')
            
            pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=param_dipilih, aggfunc='first')
            
            # Pastikan kolom 0-23 ada
            for h in range(24):
                if h not in pivot.columns:
                    pivot[h] = np.nan
            pivot = pivot[list(range(24))]
            
            # Ganti header jam menjadi string
            pivot.columns = [str(h) for h in range(24)]
            
            # Susun baris tanggal (1 sampai akhir bulan)
            semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')
            pivot = pivot.reindex(semua_tanggal)
            
            # Tambahkan Kolom Rata-rata
            pivot['R A T A   2'] = pivot.iloc[:, 0:24].mean(axis=1)
            
            # Tambahkan Kolom Jam Spesifik (23, 05, 10)
            pivot[' 23 00'] = pivot['23']
            pivot[' 05 00'] = pivot['5']
            # Jam 10 bisa kosong jika belum jamnya, amankan dengan get()
            pivot[' 10 00'] = pivot.get('10', np.nan)
            
            st.success("✅ Tabel berhasil dibuat! Preview sebagian data:")
            st.dataframe(pivot.style.format(precision=1), use_container_width=True)
            
            # 5. EXPORT KE EXCEL
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                wb = writer.book
                fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D9D9', 'border': 1})
                fmt_data = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'})
                fmt_tgl = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2', 'border': 1})
                
                pivot.to_excel(writer, sheet_name=param_dipilih[:31])
                ws = writer.sheets[param_dipilih[:31]]
                
                # Desain lebar kolom
                ws.set_column('A:A', 6, fmt_tgl)     # NO
                ws.set_column('B:Y', 6, fmt_data)    # Jam 0-23
                ws.set_column('Z:Z', 12, fmt_data)   # Rata-rata
                ws.set_column('AA:AC', 8, fmt_data)  # Jam spesifik
                
                # Tulis ulang Header agar rapi
                ws.write(0, 0, "NO.", fmt_judul)
                for col_num, value in enumerate(pivot.columns.values):
                    ws.write(0, col_num + 1, value, fmt_judul)
            
            st.download_button(
                label="📥 Unduh File Excel Anda",
                data=buffer.getvalue(),
                file_name=f"LAPORAN_{param_dipilih.upper()}_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
