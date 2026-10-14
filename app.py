import streamlit as st
import pandas as pd
import requests
import io
from calendar import monthrange

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SIASATI Sumut Extractor", page_icon="⛴️", layout="centered")

st.title("⛴️ SIASATI Extractor: ASDP Sumut")
st.markdown("Aplikasi penarik data detail Excel SIASATI Kemenhub khusus wilayah Sumatera Utara.")

# --- DAFTAR PELABUHAN SUMUT ---
pelabuhan_sumut = [
    "Ajibata", "Ambarita", "Simanindo", "Tigaras", 
    "Onan Runggu", "Sipinggan", "Muara",
    "Sibolga", "Nias", "Gunungsitoli", "Teluk Dalam", "Pulau Tello"
]

# --- ANTARMUKA PENGGUNA (UI) ---
with st.sidebar:
    st.header("⚙️ Pengaturan Parameter")
    st.markdown("Silakan ambil Cookie terbaru dari Inspect Element (F12) > Network.")
    cookie_input = st.text_area("Masukkan Cookie (XSRF-TOKEN & Session):", height=150)
    
    st.markdown("---")
    tahun_pilihan = st.selectbox("Pilih Tahun:", [2024, 2025, 2026])
    
    nama_bulan = {1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni", 
                  7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"}
    
    bulan_pilihan = st.selectbox("Pilih Bulan:", list(nama_bulan.keys()), format_func=lambda x: nama_bulan[x])
    
    tombol_ekstrak = st.button("🚀 Ekstrak Data", use_container_width=True)

# --- MESIN SCRAPING ---
if tombol_ekstrak:
    if not cookie_input.strip():
        st.error("⚠️ Mohon masukkan Cookie terlebih dahulu sebelum memulai!")
    else:
        with st.spinner(f"Sedang menembus server Kemenhub untuk data {nama_bulan[bulan_pilihan]} {tahun_pilihan}..."):
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Cookie': cookie_input.strip()
            }
            
            hari_terakhir = monthrange(tahun_pilihan, bulan_pilihan)[1]
            start_date = f"{tahun_pilihan}-{bulan_pilihan:02d}-01"
            end_date = f"{tahun_pilihan}-{bulan_pilihan:02d}-{hari_terakhir}"
            
            url_export = f"https://dashboard-siasati.kemenhub.go.id/export/matrikasdpperiode/asdp_sasati/{start_date}/{end_date}"
            
            try:
                response = requests.get(url_export, headers=headers)
                
                if response.status_code == 200 and 'spreadsheetml' in response.headers.get('Content-Type', ''):
                    df_mentah = pd.read_excel(io.BytesIO(response.content), header=2)
                    
                    if 'Unnamed: 0' in df_mentah.columns:
                        df_filter = df_mentah[df_mentah['Unnamed: 0'].isin(pelabuhan_sumut)].copy()
                        
                        if not df_filter.empty:
                            df_filter.rename(columns={
                                'Unnamed: 0': 'Pelabuhan',
                                'Unnamed: 1': 'Tanggal',
                                'Unnamed: 2': 'Periode',
                                'Unnamed: 3': 'Trip',
                                'Unnamed: 7': 'Penumpang_Dalam_Kendaraan',
                                'Unnamed: 8': 'Total_Penumpang'
                            }, inplace=True)
                            
                            df_filter.insert(0, 'Periode_Awal', start_date)
                            df_filter.insert(1, 'Periode_Akhir', end_date)
                            
                            st.success(f"✅ Berhasil! Ditemukan {len(df_filter)} baris aktivitas pelabuhan Sumut.")
                            st.dataframe(df_filter, use_container_width=True) 
                            
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                df_filter.to_excel(writer, index=False, sheet_name='ASDP_Sumut')
                            
                            st.download_button(
                                label="📥 Download Excel ASDP Sumut",
                                data=buffer.getvalue(),
                                file_name=f"Data_ASDP_Sumut_{nama_bulan[bulan_pilihan]}_{tahun_pilihan}.xlsx",
                                mime="application/vnd.ms-excel",
                                use_container_width=True
                            )
                        else:
                            st.warning("Data berhasil ditarik, tapi tidak ada aktivitas pelabuhan Sumut di bulan ini.")
                    else:
                        st.error("Struktur Excel dari Kemenhub berubah (Kolom nama pelabuhan tidak ditemukan).")
                else:
                    st.error(f"Gagal! Status Server: {response.status_code}. Kemungkinan besar Cookie Anda sudah kedaluwarsa/salah.")
                    
            except Exception as e:
                st.error(f"Terjadi error teknis: {e}")
