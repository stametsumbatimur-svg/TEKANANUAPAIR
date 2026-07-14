import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Aplikasi Input & Kalkulator Meterologi", layout="centered")

st.title("📱 Aplikasi Input Data & Kalkulator Tekanan Uap Air")
st.write("Silakan pilih menu di bawah ini untuk melakukan input atau perhitungan otomatis.")

# Membuat Menu Navigasi Sederhana
menu = st.sidebar.selectbox("Pilih Menu:", ["1. Kalkulator Koreksi Suhu", "2. Input Log Data Harian"])

# --- DATA REFERENSI TABEL KOREKSI SUHU (DARI EXCEL ANDA) ---
# Menyusun data dari baris 17 sampai 38 dengan desimal .0 sampai .9
tabel_data = {
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

# ================= MENU 1: KALKULATOR KOREKSI SUHU =================
if menu == "1. Kalkulator Koreksi Suhu":
    st.subheader("🔍 Hitung Otomatis Nilai Koreksi Suhu")
    st.write("Masukkan nilai suhu udara hasil pengamatan untuk mendapatkan nilai koreksi jenuhnya secara otomatis.")
    
    # Input Angka dari User
    input_suhu = st.number_input("Masukkan Nilai Suhu (°C) - Contoh: 25.4", 
                                 min_value=17.0, max_value=38.9, 
                                 value=25.0, step=0.1)
    
    # Proses Perhitungan Otomatis
    suhu_bulat = int(np.floor(input_suhu))
    desimal = int(round((input_suhu - suhu_bulat) * 10))
    if desimal == 10:  # Antisipasi error pembulatan float
        suhu_bulat += 1
        desimal = 0
        
    if suhu_bulat in tabel_data:
        nilai_koreksi = tabel_data[suhu_bulat][desimal]
        
        # Tampilkan Hasil di Kotak Hijau Besar
        st.success(f"### 💡 Hasil Perhitungan Otomatis:")
        st.write(f"Suhu Udara: **{input_suhu} °C**")
        st.write(f"Nilai Tekanan Uap Jenuh / Koreksi: **{nilai_koreksi}**")
    else:
        st.warning("Maaf, data suhu di luar jangkauan tabel referensi (17°C - 38°C).")


# ================= MENU 2: INPUT LOG DATA HARIAN =================
elif menu == "2. Input Log Data Harian":
    st.subheader("📝 Form Pengisian Log Tekanan Uap Air Bulanan")
    st.write("Gunakan form ini untuk memasukkan data harian baru ke dalam sistem rekapitulasi.")
    
    # Inisialisasi tabel memori kosong (1-30 Hari, 24 Jam) jika belum ada
    if 'tabel_rekap' not in st.session_state:
        st.session_state.tabel_rekap = pd.DataFrame(
            np.nan, 
            index=[f"Tanggal {i}" for i in range(1, 32)], 
            columns=[f"Jam {j:02d}.00" for j in range(24)]
        )
    
    # Form Input Komponen
    with st.form("form_log"):
        col1, col2 = st.columns(2)
        with col1:
            tgl_input = st.number_input("Pilih Tanggal (1-31):", min_value=1, max_value=31, value=1)
            jam_input = st.selectbox("Pilih Jam Observasi:", [f"Jam {j:02d}.00" for j in range(24)])
        with col2:
            nilai_input = st.number_input("Masukkan Nilai Tekanan Uap (x0.1 hPa):", min_value=0.0, max_value=500.0, value=250.0, step=0.1)
            
        submit_btn = st.form_submit_button("Simpan Data Ke Tabel Rekap")
        
    if submit_btn:
        # Menyimpan data yang diinput ke tabel di memori aplikasi
        st.session_state.tabel_rekap.at[f"Tanggal {tgl_input}", jam_input] = nilai_input
        st.success(f"Data Tanggal {tgl_input} {jam_input} sebesar {nilai_input} Berhasil Disimpan!")
        
    # Menampilkan Hasil Rekap Sementara yang mirip dengan Excel Anda
    st.markdown("---")
    st.subheader("📊 Tabel Rekapitulasi Sementara Bulanan")
    
    # Menghitung Rata-rata otomatis per baris (Tanggal) seperti kolom RATA 2 di Excel Anda
    df_tampil = st.session_state.tabel_rekap.copy()
    df_tampil['RATA-RATA (hPa)'] = df_tampil.mean(axis=1) / 10 # dibagi 10 seperti logika excel anda
    
    st.dataframe(df_tampil.style.highlight_max(axis=0, color='#ffcccc').highlight_min(axis=0, color='#cceffc'))
    
    # Tombol download hasil inputan menjadi CSV baru jika diinginkan
    csv = df_tampil.to_csv().encode('utf-8')
    st.download_button("📥 Unduh Hasil Rekap (.csv)", data=csv, file_name="REKAP_TEKANAN_UAP_BARU.csv", mime="text/csv")
