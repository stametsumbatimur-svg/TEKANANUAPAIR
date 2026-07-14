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
    try:
        # Membaca file berdasarkan ekstensinya
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=None)
        else:
            engine_choice = 'xlrd' if uploaded_file.name.endswith('.xls') else None
            df = pd.read_excel(uploaded_file, header=None, engine=engine_choice)
            
        # --- PROSES PEMBERSIHAN DATA ---
        # Mengambil data observasi jam 00-23 (Baris index 2 sampai 14)
        data_jam = df.iloc[2:15, 1:25].astype(float)
        
        # FIX: Mengubah ke float terlebih dahulu baru ke int agar aman dari error '1.0'
        hari = df.iloc[2:15, 0].astype(float).astype(int).values
        
        data_jam.index = hari
        data_jam.columns = [f"Jam {int(float(c)):02d}.00" for c in df.iloc[0, 1:25]]
        
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
            hourly_mean = data_jam.mean(axis=0).reset_index()
            hourly_mean.columns = ['Jam', 'Rata-rata Tekanan Uap']
            fig_line = px.line(hourly_mean, x='Jam', y='Rata-rata Tekanan Uap', title="Tren Perubahan Tekanan Uap Air Rata-rata Berdasarkan Jam", markers=True)
            st.plotly_chart(fig_line, use_container_width=True)
            
        with tab2:
            fig_heat = px.imshow(data_jam.values,
                                 labels=dict(x="Waktu (Jam)", y="Tanggal", color="Tekanan Uap"),
                                 x=data_jam.columns,
                                 y=data_jam.index,
                                 color_continuous_scale='Viridis',
                                 title="Heatmap Tekanan Uap Air (Hari vs Jam)")
            st.plotly_chart(fig_heat, use_container_width=True)
            
    except Exception as e:
        st.error(f"Gagal memproses file. Pastikan struktur file sesuai template aslinya. Error: {e}")

else:
    st.info("Silakan unggah file Excel Tekanan Uap Air terlebih dahulu untuk memulai analisis.")
