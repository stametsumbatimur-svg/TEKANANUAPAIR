import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Analisis Tekanan Uap Air", layout="wide")
st.title("📊 Aplikasi Analisis Tekanan Uap Air & Koreksi Psikrometrik")
st.write("Aplikasi ini didesain khusus untuk membaca dan menganalisis format log data Tekanan Uap Air.")

# 1. Upload File
uploaded_file = st.file_uploader("Unggah file Excel (.xls / .xlsx) atau CSV Tekanan Uap Air Anda", type=["xls", "xlsx", "csv"])

if uploaded_file is not None:
    # Membaca data (Mengasumsikan sheet pertama adalah Tekanan Uap Air)
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=None)
        else:
            df = pd.read_excel(uploaded_file, header=None)
            
        # --- PROSES PEMBERSIHAN DATA ---
        # Mengambil data observasi jam 00-23 (Baris index 2 sampai 14)
        data_jam = df.iloc[2:15, 1:25].astype(float)
        hari = df.iloc[2:15, 0].astype(float).astype(int).values
        
        data_jam.index = hari
        data_jam.columns = [f"Jam {int(c):02d}.00" for c in df.iloc[0, 1:25].astype(float)]
        
        # --- METRIK UTAMA ---
        flat_values = data_jam.values.flatten()
        val_min = np.nanmin(flat_values)
        val_max = np.nanmax(flat_values)
        val_mean = np.nanmean(flat_values)
        
        st.subheader("📌 Ringkasan Data Observasi (Tanggal 1 - 13)")
        col1, col2, col3 = st.columns(3)
        col1.metric("Tekanan Maksimum (x0.1 hPa)", f"{val_max:.2f}")
        col2.metric("Tekanan Minimum (x0.1 hPa)", f"{val_min:.2f}")
        col3.metric("Rata-rata Keseluruhan", f"{val_mean:.2f}")
        
        st.markdown("---")
        
        # --- VISUALISASI ---
        st.subheader("📈 Visualisasi Tren & Fluktuasi")
        tab1, tab2 = st.tabs(["Tren Per Jam (Rata-rata)", "Heatmap Kepadatan"])
        
        with tab1:
            # Grafik rata-rata per jam
            hourly_mean = data_jam.mean(axis=0).reset_index()
            hourly_mean.columns = ['Jam', 'Rata-rata Tekanan Uap']
            fig_line = px.line(hourly_mean, x='Jam', y='Rata-rata Tekanan Uap', title="Tren Perubahan Tekanan Uap Air Rata-rata Berdasarkan Jam", markers=True)
            st.plotly_chart(fig_line, use_container_width=True)
            
        with tab2:
            # Heatmap Hari vs Jam
            fig_heat = px.imshow(data_jam.values,
                                 labels=dict(x="Waktu (Jam)", y="Tanggal", color="Tekanan Uap"),
                                 x=data_jam.columns,
                                 y=data_jam.index,
                                 color_continuous_scale='Viridis',
                                 title="Heatmap Tekanan Uap Air (Hari vs Jam)")
            st.plotly_chart(fig_heat, use_container_width=True)
            
        st.markdown("---")
        
        # --- TABEL REFERENSI & KOREKSI SUHU ---
        st.subheader("🔍 Kalkulator Interaktif Koreksi Suhu")
        st.write("Fitur ini mengekstrak tabel referensi psikrometrik yang ada di bagian bawah file Anda.")
        
        # Ekstraksi tabel Koreksi Suhu (Baris 39 ke bawah, kolom 13 ke kanan)
        tabel_suhu = df.iloc[39:61, 13:24].dropna(how='all').copy()
        if not tabel_suhu.empty:
            tabel_suhu.columns = ['Suhu'] + [f"Koreksi .{i}" for i in range(10)]
            tabel_suhu['Suhu'] = tabel_suhu['Suhu'].astype(float)
            
            # Input dari user
            input_suhu = st.number_input("Masukkan Nilai Suhu untuk mencari nilai Koreksi (Contoh: 25.4):", 
                                         min_value=float(tabel_suhu['Suhu'].min()), 
                                         max_value=float(tabel_suhu['Suhu'].max() + 0.9), 
                                         value=25.0, step=0.1)
            
            # Proses pencarian nilai di tabel
            suhu_bulat = int(np.floor(input_suhu))
            desimal = int(round((input_suhu - suhu_bulat) * 10))
            if desimal == 10: # Antisipasi pembulatan float
                suhu_bulat += 1
                desimal = 0
                
            row_match = tabel_suhu[tabel_suhu['Suhu'] == float(suhu_bulat)]
            
            if not row_match.empty:
                nilai_koreksi = row_match.iloc[0, desimal + 1]
                st.success(f"💡 Untuk Suhu **{input_suhu}°C**, Nilai Koreksi Suhu yang ditemukan adalah: **{nilai_koreksi}**")
            else:
                st.warning("Data suhu di luar jangkauan tabel referensi.")
        else:
            st.info("Tabel referensi koreksi suhu tidak ditemukan atau format berbeda.")
            
    except Exception as e:
        st.error(f"Gagal memproses file. Pastikan struktur file sesuai template aslinya. Error: {e}")

else:
    st.info("Silakan unggah file Excel Tekanan Uap Air terlebih dahulu untuk memulai analisis.")
