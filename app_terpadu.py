import streamlit as st
import pandas as pd
import requests
import io
from calendar import monthrange

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="SIASATI Sumut Terpadu", page_icon="📊", layout="wide")

st.title("📊 SIASATI Extractor: Bus, Kereta, & ASDP Sumut")
st.markdown("Aplikasi terpadu penarik data API SIASATI Kemenhub khusus wilayah Sumatera Utara.")

# --- KAMUS ID PRASARANA SUMATERA UTARA ---
# Semua ID dijadikan format teks (string) agar seragam saat di-filter
id_sumut = {
    'darat_bus': ['APS', 'B1010', 'B1015', 'B114', 'B117', 'B120', 'B123', 'B1301', 'B1304', 'B300', 'B301', 'MTR', 'PNB', 'SBG', 'TPG'],
    'kereta': ['ARB', 'BAP', 'BDT', 'BIJ', 'BJL', 'BTK', 'DMR', 'KBG', 'KIS', 'KNM', 'LAG', 'LBP', 'LDT', 'LMP', 'LTD', 'MBM', 'MBU', 'MDN', 'PBA', 'PHA', 'PRA', 'PUR', 'RAP', 'RPH', 'SBJ', 'SIR', 'TBI', 'TGD', 'TKE', 'TNB'],
    'asdp_sasati': ['3', '1713', '1751', '1752', '2768', '2981', '3308', '3326', '4064', '4073', '4097']
}

# --- ANTARMUKA PENGGUNA (UI) SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Pengaturan Parameter")
    st.markdown("Masukkan Cookie terbaru dari menu Inspect Element (F12) > Network.")
    cookie_input = st.text_area("Cookie Kemenhub:", height=150)
    
    st.markdown("---")
    moda_label = st.selectbox("Pilih Moda Transportasi:", ["Bus", "Kereta Api", "ASDP (Penyeberangan)"])
    
    # Mapping label UI ke parameter API Kemenhub
    moda_map = {
        "Bus": "darat_bus",
        "Kereta Api": "kereta",
        "ASDP (Penyeberangan)": "asdp_sasati"
    }
    moda_api = moda_map[moda_label]
    
    # Filter Tahun 2019 - 2030 (Default ke 2026 yang memiliki index ke-7)
    tahun_pilihan = st.selectbox("Pilih Tahun:", list(range(2019, 2031)), index=7)
    
    nama_bulan = {1:"Januari", 2:"Februari", 3:"Maret", 4:"April", 5:"Mei", 6:"Juni", 
                  7:"Juli", 8:"Agustus", 9:"September", 10:"Oktober", 11:"November", 12:"Desember"}
    bulan_pilihan = st.selectbox("Pilih Bulan:", list(nama_bulan.keys()), format_func=lambda x: nama_bulan[x])
    
    tombol_ekstrak = st.button("🚀 Ekstrak Data", use_container_width=True)

# --- MESIN PENARIK DATA (POST API) ---
if tombol_ekstrak:
    if not cookie_input.strip():
        st.error("⚠️ Mohon masukkan Cookie terlebih dahulu sebelum memulai!")
    else:
        with st.spinner(f"Menarik data {moda_label} untuk {nama_bulan[bulan_pilihan]} {tahun_pilihan}..."):
            
            # Setup Header API
            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Origin': 'https://dashboard-siasati.kemenhub.go.id',
                'Referer': f'https://dashboard-siasati.kemenhub.go.id/periode-tanggal?moda={moda_api}',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'X-Requested-With': 'XMLHttpRequest',
                'Cookie': cookie_input.strip()
            }

            url_api = "https://dashboard-siasati.kemenhub.go.id/api/matrix/moda"
            
            # Setup Tanggal
            hari_terakhir = monthrange(tahun_pilihan, bulan_pilihan)[1]
            start_date = f"{tahun_pilihan}-{bulan_pilihan:02d}-01"
            end_date = f"{tahun_pilihan}-{bulan_pilihan:02d}-{hari_terakhir}"
            
            # Payload Body API
            payload = {
                'moda': moda_api,
                'start_date': start_date,
                'end_date': end_date,
                'cache_check': 'true'
            }
            
            try:
                response = requests.post(url_api, headers=headers, data=payload)
                
                if response.status_code == 200:
                    json_data = response.json()
                    
                    # Konversi JSON ke DataFrame Pandas
                    if isinstance(json_data, list):
                        df_temp = pd.DataFrame(json_data)
                    elif 'data' in json_data:
                        df_temp = pd.DataFrame(json_data['data'])
                    else:
                        df_temp = pd.DataFrame([json_data])
                        
                    if not df_temp.empty:
                        # Menambahkan kolom tambahan agar sesuai dengan format Excel referensi BPS
                        df_temp['Periode_Awal'] = start_date
                        df_temp['Periode_Akhir'] = end_date
                        df_temp['Bulan'] = bulan_pilihan
                        df_temp['Tahun'] = tahun_pilihan
                        
                        # Memastikan kolom id_prasarana ada sebelum di-filter
                        if 'id_prasarana' in df_temp.columns:
                            # Ubah semua id_prasarana jadi teks agar cocok saat filter
                            df_temp['id_prasarana'] = df_temp['id_prasarana'].astype(str)
                            
                            # Melakukan filter wilayah Sumut
                            id_target = id_sumut[moda_api]
                            df_filter = df_temp[df_temp['id_prasarana'].isin(id_target)].copy()
                            
                            if not df_filter.empty:
                                st.success(f"✅ Berhasil! Ditemukan {len(df_filter)} data {moda_label} untuk wilayah Sumut.")
                                
                                # Menampilkan tabel di Web
                                st.dataframe(df_filter, use_container_width=True)
                                
                                # Menyiapkan Excel untuk di-download
                                buffer = io.BytesIO()
                                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                    df_filter.to_excel(writer, index=False, sheet_name=moda_label[:31])
                                
                                st.download_button(
                                    label=f"📥 Download Excel {moda_label} Sumut",
                                    data=buffer.getvalue(),
                                    file_name=f"Data_{moda_label}_Sumut_{nama_bulan[bulan_pilihan]}_{tahun_pilihan}.xlsx",
                                    mime="application/vnd.ms-excel",
                                    use_container_width=True
                                )
                            else:
                                st.warning(f"Data API berhasil ditarik, tapi tidak ditemukan transaksi {moda_label} di stasiun/pelabuhan Sumut pada bulan ini.")
                                with st.expander("Lihat Data Mentah (Semua Wilayah Se-Indonesia)"):
                                    st.dataframe(df_temp)
                        else:
                            st.error("Struktur API berubah: Kolom 'id_prasarana' tidak ditemukan dari Kemenhub.")
                            st.write("Preview Data Mentah:", df_temp.head())
                    else:
                        st.warning("Server merespons sukses (Status 200), tapi isi datanya kosong. Pastikan Cookie Anda valid dan belum kedaluwarsa.")
                else:
                    st.error(f"Gagal! Status Server Kemenhub: {response.status_code}. Kemungkinan besar Cookie Anda salah/kedaluwarsa.")
                    
            except Exception as e:
                st.error(f"Terjadi error teknis pada sistem: {e}")
