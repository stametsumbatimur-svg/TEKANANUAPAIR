import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Kalkulator RH & Tekanan Uap Otomatis", layout="centered")

st.title("⚡ Aplikasi Meteorologi: Hitung RH & Koreksi Suhu Otomatis")
st.write("Aplikasi ini otomatis menghitung Persen Kelembapan (RH) berdasarkan referensi log Excel Anda.")

# Data Tabel Koreksi Suhu (Baris 17°C - 38°C, Kolom Desimal .0 sampai .9)
tabel_suhu = {
    17: [18.8, 18.9, 19.1, 19.2, 19.4, 19.5, 19.7, 19.8, 20.0, 20.2],
    18: [20.4, 20.5, 20.6, 20.8, 20.9, 21.1, 21.2, 21.4, 21.5, 21.7],
    19: [21.8, 22.0, 22.0, 22.3, 22.5, 22.6, 22.8, 22.9, 23.1, 23.2],
    20: [23.4, 23.5, 23.7, 23.8, 24.0, 24.1, 24.3, 24.4, 24.6, 24.7],
    21: [24.9, 25.0, 25.2, 25.3, 25.5, 25.6, 25.8, 26.0, 26.1, 26.3],
    22: [26.4, 26.6, 26.8, 26.9, 27.1, 27.2, 27.4, 27.6, 27.7, 27.9],
    23: [28.1, 28.3, 28.4, 28.6, 28.8, 28.9, 29.1, 29.3, 29.5, 29.7],
    24: [29.8, 30.0, 30.2, 30.4, 30.6, 30.7, 30.9, 31.1, 31.3, 31.5],
    25: [31.7, 31.9, 32.1, 32.2, 32.4, 32.6, 32.8, 33.0, 33.2, 33.4],
    26: [33.6, 33.8, 34.0, 34.2, 34.4, 34.6, 34.8, 35.0, 35.2, 35.4],
    27: [35.8, 35.9, 36.1, 36.3, 36.5, 36.7, 36.9, 37.0, 37.4, 37.6],
    28: [37.8, 38.0, 38.2, 38.5, 38.6, 38.9, 39.1, 39.4, 39.6, 39.8],
    29: [40.1, 40.3, 40.5, 40.8, 41.0, 41.2, 41.5, 41.7, 41.9, 42.2],
    30: [42.4, 42.7, 42.9, 43.2, 43.4, 43.7, 43.9, 44.4, 44.4, 44.7],
    31: [44.9, 45.2, 45.4, 45.7, 46.0, 46.2, 46.5, 46.8, 47.0, 47.3],
    32: [47.6, 47.8, 48.1, 48.4, 48.6, 48.9, 49.2, 49.5, 49.7, 50.0],
    33: [50.3, 50.6, 50.9, 51.2, 51.4, 51.7, 52.0, 52.3, 52.6, 52.9],
    34: [53.5, 53.5, 53.8, 54.1, 55.4, 54.7, 55.0, 55.3, 55.6, 55.9],
    35: [56.2, 56.5, 56.9, 57.2, 57.6, 57.8, 58.1, 58.5, 58.8, 59.1],
    36: [59.4, 59.7, 60.1, 60.4, 60.7, 61.1, 61.4, 61.7, 62.1, 62.4],
    37: [62.8, 63.1, 63.5, 63.6, 64.1, 64.5, 64.8, 65.2, 65.6, 65.9],
    38: [66.3, 66.6, 67.0, 67.3, 67.7, 68.1, 68.4, 68.6, 69.2, 69.4]
}

# --- INPUT FORM UTAMA ---
st.subheader("📝 Input Data Observasi")

col1, col2 = st.columns(2)
with col1:
    suhu_input = st.number_input("1. Masukkan Suhu Udara (°C):", min_value=17.0, max_value=38.9, value=25.4, step=0.1)
with col2:
    uap_input = st.number_input("2. Masukkan Tekanan Uap Air (x0.1 hPa):", min_value=0.0, max_value=500.0, value=252.9, step=0.1)

st.markdown("---")

# --- PROSES PERHITUNGAN OTOMATIS SIMULTAN ---
suhu_bulat = int(np.floor(suhu_input))
desimal = int(round((suhu_input - suhu_bulat) * 10))
if desimal == 10:
    suhu_bulat += 1
    desimal = 0

if suhu_bulat in tabel_suhu:
    # 1. Dapatkan Nilai Tekanan Uap Jenuh (E_s) dari tabel koreksi suhu
    e_s = tabel_suhu[suhu_bulat][desimal]
    
    # 2. Konversi Tekanan Uap Input ke satuan hPa utuh (e)
    e_actual = uap_input / 10
    
    # 3. Hitung Persen RH otomatis
    rh_kalkulasi = (e_actual / e_s) * 100
    # Memastikan nilai RH tidak melebihi 100% secara logika matematis
    if rh_kalkulasi > 100.0:
        rh_kalkulasi = 100.0

    # --- TAMPILKAN HASIL REAL-TIME ---
    st.subheader("📊 Hasil Perhitungan Otomatis Sistem:")
    
    c1, c2, c3 = st.columns(3)
    c1.metric(label="Tekanan Uap Jenuh ($e_s$)", value=f"{e_s} hPa")
    c2.metric(label="Tekanan Uap Aktual ($e$)", value=f"{e_actual:.2f} hPa")
    
    # Beri warna hijau/highlight besar untuk Persen RH sebagai output utama
    st.info(f"### 🎯 Nilai Kelembapan Udara (Persen RH): **{rh_kalkulasi:.1f} %**")
    
    # Keterangan tambahan untuk validasi petugas
    if rh_kalkulasi < 30:
        st.warning("⚠️ Kondisi Udara Sangat Kering.")
    elif rh_kalkulasi > 95:
        st.write("💧 Kondisi Udara Sangat Lembap / Jenuh Udara Basah.")

else:
    st.error("Suhu yang Anda masukkan berada di luar jangkauan tabel matriks referensi.")
