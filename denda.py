import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os
import glob
import openpyxl
import xlsxwriter

# Konfigurasi halaman
st.set_page_config(
    page_title="Aplikasi Simulasi Perhitungan Denda",
    page_icon="💸",
    layout="wide",
)

# CSS untuk styling
st.markdown("""
<style>
    .main-header {
        font-size: 30px;
        font-weight: bold;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 20px;
    }
    .subtitle {
        font-size: 20px;
        font-weight: bold;
        color: #0D47A1;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    .highlight {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #1E88E5;
        margin-bottom: 10px;
    }
    .stButton>button {
        background-color: #1E88E5;
        color: white;
        border-radius: 5px;
        border: none;
        padding: 10px 15px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #0D47A1;
    }
    .result-container {
        background-color: #e8f5e9;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #4CAF50;
        margin-top: 20px;
    }
    .filter-section {
        background-color: #f5f5f5;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    .calculation-info {
        background-color: #fff3e0;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #FF9800;
        margin-bottom: 10px;
    }
    .debug-info {
        background-color: #f3e5f5;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #9c27b0;
        margin-bottom: 10px;
        font-family: monospace;
        font-size: 12px;
    }
    .jenis-izin-box {
        background-color: #e8eaf6;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #3949ab;
        margin-bottom: 10px;
    }
    .warning-box {
        background-color: #ffecb3;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #ffa000;
        margin-bottom: 10px;
    }
    .file-selector {
        background-color: #e0f7fa;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #00acc1;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Judul aplikasi
st.markdown("<div class='main-header'>Aplikasi Simulasi Perhitungan Denda Pelanggaran Frekuensi Radio & Perangkat Telekomunikasi</div>", unsafe_allow_html=True)

# Konstanta untuk nilai maksimum poin berdasarkan jenis izin
MAKS_POIN_DEFAULT = {
    "IPFR": 600000,
    "ISR": 7000,
    "APT": 5000
}

# Fungsi untuk menemukan semua file Excel dalam folder Data
def find_excel_files(data_folder="Data"):
    # Pastikan folder Data ada
    if not os.path.exists(data_folder):
        try:
            os.makedirs(data_folder)
            st.info(f"Folder {data_folder} telah dibuat. Silakan tambahkan file Excel ke folder tersebut.")
            return []
        except Exception as e:
            st.error(f"Error saat membuat folder {data_folder}: {e}")
            return []
    
    # Cari semua file Excel dalam folder Data
    excel_files = glob.glob(os.path.join(data_folder, "*.xlsx")) + glob.glob(os.path.join(data_folder, "*.xls"))
    
    return excel_files

# Fungsi untuk membaca file Excel
@st.cache_data
def load_excel(file_path):
    try:
        # Baca data dari file Excel
        excel_data = pd.ExcelFile(file_path)
        
        # Buat dictionary untuk menyimpan semua sheet
        sheets = {}
        
        # Baca setiap sheet
        for sheet_name in excel_data.sheet_names:
            sheets[sheet_name] = pd.read_excel(excel_data, sheet_name=sheet_name)
            
        return sheets, True
    except Exception as e:
        st.error(f"Error saat membaca file Excel {os.path.basename(file_path)}: {e}")
        return None, False

# Fungsi untuk menemukan header berdasarkan nilai tertentu dalam dataframe
def find_header_row(df, header_values):
    for i, row in df.iterrows():
        # Periksa jika semua nilai yang dicari ada dalam baris saat ini
        if all(value in row.values for value in header_values):
            return i
    return None

# Fungsi untuk memproses data dari sheet FREK & ALAT
def process_frek_alat_data(df):
    # Cari baris header
    header_values = ["DINAS", "KATEGORI", "BAND"]
    header_row = find_header_row(df, header_values)
    
    if header_row is None:
        st.error("Tidak dapat menemukan baris header di sheet FREK & ALAT")
        return None
    
    # Reset header dengan baris yang ditemukan
    header = df.iloc[header_row]
    processed_df = df.iloc[header_row+1:].reset_index(drop=True)
    processed_df.columns = header.values
    
    # Filter kolom yang tidak diinginkan (NaN atau unnamed)
    valid_columns = [col for col in processed_df.columns if not (pd.isna(col) or 'Unnamed' in str(col))]
    processed_df = processed_df[valid_columns]
    
    # Konversi kolom numerik
    numeric_cols = ['ZONA', 'MAKS POIN', 'INDEKS PELANGGARAN PERTAMA', 'INDEKS PELANGGARAN BERULANG',
                    '%', 'TOTAL POIN', 'TARIF DENDA', 'DENDA', 
                    'JUMLAH FREKUENSI', 'JUMLAH PERANGKAT']
    
    for col in numeric_cols:
        if col in processed_df.columns:
            processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce')
    
    # Pastikan kolom JENIS IZIN ada
    if 'JENIS IZIN' not in processed_df.columns:
        st.warning("Kolom JENIS IZIN tidak ditemukan. Aplikasi akan mencoba menggunakan nilai default.")
    else:
        # Normalisasi nilai JENIS IZIN (uppercase dan strip whitespace)
        processed_df['JENIS IZIN'] = processed_df['JENIS IZIN'].astype(str).str.strip().str.upper()
    
    return processed_df

# Fungsi untuk memproses data dari sheet Referensi untuk mendapatkan faktor persentase
def process_referensi_data(sheets):
    try:
        # Default persentase data
        persentase_data = {"0-12": 1.0, "13-24": 0.5, ">25": 0.25}
        
        # Sheet Referensi (jika ada)
        if 'Referensi' in sheets:
            ref_df = sheets['Referensi']
            
            # Cari baris dengan nilai "0-12", "13-24", ">25"
            period_values = ["0-12", "13-24", ">25"]
            
            for i, row in ref_df.iterrows():
                row_str = [str(val).strip() if not pd.isna(val) else "" for val in row.values]
                row_str = " ".join(row_str).lower()
                
                # Cek jika baris ini mengandung referensi ke periode
                if any(period.lower() in row_str for period in period_values):
                    row_values = [str(val).strip() if not pd.isna(val) else "" for val in row.values]
                    
                    # Temukan indeks kolom untuk periode
                    period_indices = {}
                    for j, value in enumerate(row_values):
                        for period in period_values:
                            if period == value:
                                period_indices[period] = j
                    
                    # Jika periode ditemukan dan baris berikutnya tersedia
                    if period_indices and i+1 < ref_df.shape[0]:
                        next_row = ref_df.iloc[i+1]
                        
                        # Ambil nilai persentase dari baris berikutnya
                        for period, idx in period_indices.items():
                            try:
                                percentage = next_row.iloc[idx]
                                if isinstance(percentage, (int, float)):
                                    # Normalisasi persentase
                                    if percentage > 1:
                                        percentage = percentage / 100
                                    persentase_data[period] = percentage
                                elif isinstance(percentage, str) and percentage.replace('.', '', 1).replace(',', '', 1).isdigit():
                                    # Konversi string ke float
                                    percentage = float(percentage.replace(',', '.'))
                                    if percentage > 1:
                                        percentage = percentage / 100
                                    persentase_data[period] = percentage
                            except:
                                pass
        
        return persentase_data
    
    except Exception as e:
        st.error(f"Error saat memproses data referensi: {e}")
        return {"0-12": 1.0, "13-24": 0.5, ">25": 0.25}

# Fungsi untuk mendapatkan persentase berdasarkan JML BULAN
def get_percentage(persentase_data, jml_bulan):
    # Default persentase
    default_percentage = 1.0
    
    # Cek jika JML BULAN ada dalam data persentase
    if isinstance(persentase_data, dict) and jml_bulan in persentase_data:
        return persentase_data[jml_bulan]
    
    return default_percentage

# Fungsi untuk memfilter data berdasarkan kriteria
def filter_data(df, filters):
    # Pastikan df adalah DataFrame
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    
    filtered_df = df.copy()
    
    # Terapkan filter
    for column, value in filters.items():
        if column in filtered_df.columns and value and value != "Semua":
            if column == 'ZONA' and isinstance(value, str) and value != "Semua":
                try:
                    filtered_df = filtered_df[filtered_df[column] == int(value)]
                except:
                    pass
            else:
                filtered_df = filtered_df[filtered_df[column] == value]
    
    return filtered_df

# Fungsi untuk mendapatkan MAKS POIN berdasarkan JENIS IZIN
def get_maks_poin(jenis_izin):
    jenis_izin = str(jenis_izin).strip().upper()
    return MAKS_POIN_DEFAULT.get(jenis_izin, 0)

# Fungsi untuk menghitung TOTAL POIN, DENDA, dan TOTAL TAGIHAN DENDA
def calculate_denda(row, jumlah_frekuensi, jumlah_perangkat, persentase=1.0, jenis_pelanggaran="Pelanggaran Pertama"):
    """
    Menghitung TOTAL POIN, DENDA, dan TOTAL TAGIHAN DENDA berdasarkan rumus:
    TOTAL POIN = INDEKS PELANGGARAN * % * MAKS POIN
    DENDA = TOTAL POIN * TARIF DENDA
    TOTAL TAGIHAN DENDA = DENDA * JUMLAH FREKUENSI * JUMLAH PERANGKAT
    
    Menggunakan MAKS POIN berdasarkan JENIS IZIN jika tersedia.
    """
    try:
        # Ambil JENIS IZIN dan tetapkan MAKS POIN sesuai jenisnya
        jenis_izin = str(row.get('JENIS IZIN', '')).strip().upper()
        
        # Default nilai dari konstanta berdasarkan JENIS IZIN
        default_maks_poin = get_maks_poin(jenis_izin)
        
        # Gunakan MAKS POIN dari data jika tersedia, jika tidak gunakan default berdasarkan JENIS IZIN
        maks_poin_data = 0 if pd.isna(row.get('MAKS POIN')) else float(row.get('MAKS POIN', 0))
        maks_poin = default_maks_poin if maks_poin_data == 0 else maks_poin_data
        
        # Jika MAKS POIN masih 0, gunakan nilai 1 untuk menghindari division by zero
        if maks_poin == 0:
            maks_poin = 1
        
        # Ambil INDEKS PELANGGARAN berdasarkan jenis pelanggaran
        if jenis_pelanggaran == "Pelanggaran Pertama":
            indeks = 0 if pd.isna(row.get('INDEKS PELANGGARAN PERTAMA')) else float(row.get('INDEKS PELANGGARAN PERTAMA', 0))
        else:  # Pelanggaran Berulang
            indeks = 0 if pd.isna(row.get('INDEKS PELANGGARAN BERULANG')) else float(row.get('INDEKS PELANGGARAN BERULANG', 0))
        
        # Ambil persentase dari data jika tersedia, jika tidak gunakan parameter
        percentage_data = 0 if pd.isna(row.get('%')) else float(row.get('%', 0))
        if percentage_data > 1:  # Normalisasi jika di atas 1
            percentage_data = percentage_data / 100
        
        percentage = percentage_data if percentage_data > 0 else persentase
        
        # Ambil TARIF DENDA
        tarif_denda = 0 if pd.isna(row.get('TARIF DENDA')) else float(row.get('TARIF DENDA', 0))
        
        # Ambil nilai yang sudah ada jika tersedia
        existing_total_poin = None if pd.isna(row.get('TOTAL POIN')) else float(row.get('TOTAL POIN', 0))
        existing_denda = None if pd.isna(row.get('DENDA')) else float(row.get('DENDA', 0))
        
        # Hitung TOTAL POIN
        if existing_total_poin is not None and existing_total_poin > 0:
            total_poin = existing_total_poin
        else:
            total_poin = indeks * percentage * maks_poin
        
        # Hitung DENDA
        if existing_denda is not None and existing_denda > 0:
            denda = existing_denda
        else:
            denda = total_poin * tarif_denda
        
        # Hitung TOTAL TAGIHAN DENDA
        jumlah_frekuensi = 1 if jumlah_frekuensi <= 0 else jumlah_frekuensi
        jumlah_perangkat = 1 if jumlah_perangkat <= 0 else jumlah_perangkat
        
        total_tagihan_denda = denda * jumlah_frekuensi * jumlah_perangkat
        
        # Return semua nilai untuk debugging dan visualisasi
        return {
            'indeks': indeks,
            'persentase': percentage,
            'maks_poin': maks_poin,
            'total_poin': total_poin,
            'tarif_denda': tarif_denda,
            'denda': denda,
            'total_tagihan_denda': total_tagihan_denda
        }
    
    except Exception as e:
        st.error(f"Error saat menghitung denda: {e}")
        return {
            'indeks': 0,
            'persentase': 0,
            'maks_poin': 0,
            'total_poin': 0,
            'tarif_denda': 0,
            'denda': 0,
            'total_tagihan_denda': 0
        }

# Fungsi untuk mengonversi dataframe ke CSV (alternatif Excel untuk menghindari dependensi xlsxwriter)
def to_csv(df):
    output = BytesIO()
    df.to_csv(output, index=False)
    output.seek(0)
    return output.getvalue()

# Coba fungsi untuk mengonversi dataframe ke Excel
def to_excel(df):
    try:
        # Coba dengan openpyxl
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Hasil Perhitungan', index=False)
        output.seek(0)
        return output.getvalue(), True
    except ImportError:
        try:
            # Jika openpyxl tidak ada, coba dengan xlsxwriter
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Hasil Perhitungan', index=False)
            output.seek(0)
            return output.getvalue(), True
        except ImportError:
            # Jika kedua engine tidak ada, kembalikan None
            return None, False
        except Exception as e:
            st.error(f"Error saat membuat file Excel: {e}")
            return None, False
    except Exception as e:
        st.error(f"Error saat membuat file Excel: {e}")
        return None, False

# Temukan semua file Excel di folder Data
excel_files = find_excel_files()

# Main container
with st.container():
    # Informasi JENIS IZIN dan MAKS POIN
    st.markdown("""
    <div class='jenis-izin-box'>
        <strong>Informasi JENIS IZIN dan MAKS POIN:</strong>
        <ul>
            <li>IPFR (Izin Penggunaan Frekuensi Radio): 600.000 poin</li>
            <li>ISR (Izin Stasiun Radio): 7.000 poin</li>
            <li>APT (Alat Perangkat Telekomunikasi): 5.000 poin</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Debug expander
    debug_expander = st.expander("Debug Info (Developer Only)", expanded=False)
    
    # File selection section
    st.markdown("<div class='subtitle'>Pilih File Data</div>", unsafe_allow_html=True)
    
    if excel_files:
        # Tampilkan dropdown untuk memilih file
        file_options = [os.path.basename(file) for file in excel_files]
        selected_file = st.selectbox("Pilih file Excel:", file_options)
        
        # Dapatkan path lengkap file terpilih
        selected_file_path = excel_files[file_options.index(selected_file)]
        
        # Tampilkan info file terpilih
        st.markdown(f"""
        <div class='file-selector'>
            <p><strong>File terpilih:</strong> {selected_file}</p>
            <p><strong>Path:</strong> {selected_file_path}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Tombol untuk memuat data
        if st.button("Muat Data"):
            with st.spinner(f'Memproses file {selected_file}...'):
                # Baca data Excel
                sheets, success = load_excel(selected_file_path)
                
                if success:
                    st.success(f"File berhasil dimuat. Sheet yang tersedia: {', '.join(sheets.keys())}")
                    
                    # Proses data dari berbagai sheet
                    frek_alat_df = None
                    if 'FREK & ALAT' in sheets:
                        frek_alat_df = process_frek_alat_data(sheets['FREK & ALAT'])
                    
                    # Proses data persentase dari Referensi
                    persentase_data = process_referensi_data(sheets)
                    
                    # Simpan dataframes dalam session state agar bisa diakses di bagian lain aplikasi
                    st.session_state['frek_alat_df'] = frek_alat_df
                    st.session_state['persentase_data'] = persentase_data
                    st.session_state['selected_file'] = selected_file
                    
                    # Tampilkan debug info jika diperlukan
                    with debug_expander:
                        st.markdown("### Data Persentase:")
                        st.write(persentase_data)
                        
                        st.markdown("### MAKS POIN Default berdasarkan JENIS IZIN:")
                        st.write(MAKS_POIN_DEFAULT)
                        
                        if frek_alat_df is not None:
                            st.markdown("### Data FREK & ALAT (5 baris pertama):")
                            st.write(frek_alat_df.head())
                            
                            st.markdown("### Kolom yang Tersedia:")
                            st.write(frek_alat_df.columns.tolist())
    else:
        st.warning(f"""
        Tidak ada file Excel ditemukan di folder 'Data'. 
        Silakan tambahkan file Excel ke folder tersebut dan mulai ulang aplikasi.
        """)

# Section perhitungan denda
if 'frek_alat_df' in st.session_state and st.session_state['frek_alat_df'] is not None:
    frek_alat_df = st.session_state['frek_alat_df']
    persentase_data = st.session_state['persentase_data']
    selected_file = st.session_state.get('selected_file', 'Data')
    
    # Tampilkan informasi persentase
    st.markdown("<div class='calculation-info'>", unsafe_allow_html=True)
    st.markdown("**Informasi Persentase Berdasarkan JML BULAN:**", unsafe_allow_html=True)
    for period, percentage in persentase_data.items():
        st.markdown(f"- Periode {period}: {percentage*100:.0f}%", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Sidebar untuk filter
    st.sidebar.markdown("<div class='subtitle'>Filter Data</div>", unsafe_allow_html=True)
    
    # Filter untuk JENIS IZIN (prioritaskan sebelum filter lainnya jika tersedia)
    if 'JENIS IZIN' in frek_alat_df.columns:
        jenis_izin_values = frek_alat_df['JENIS IZIN'].dropna().unique()
        # Normalisasi nilai JENIS IZIN
        normalized_jenis_izin = [str(val).strip().upper() for val in jenis_izin_values]
        
        # Tambahkan pilihan IPFR, ISR, APT jika belum ada
        for key_izin in MAKS_POIN_DEFAULT.keys():
            if key_izin not in normalized_jenis_izin:
                normalized_jenis_izin.append(key_izin)
        
        unique_jenis_izin = ["Semua"] + sorted(normalized_jenis_izin)
        selected_jenis_izin = st.sidebar.selectbox("JENIS IZIN", unique_jenis_izin)
    else:
        # Jika kolom JENIS IZIN tidak ada, buat dropdown manual
        unique_jenis_izin = ["Semua"] + sorted(list(MAKS_POIN_DEFAULT.keys()))
        selected_jenis_izin = st.sidebar.selectbox("JENIS IZIN", unique_jenis_izin)
    
    # Filter untuk DINAS
    unique_dinas = ["Semua"] + sorted(list(frek_alat_df['DINAS'].dropna().unique()))
    selected_dinas = st.sidebar.selectbox("DINAS", unique_dinas)
    
    # Filter untuk KATEGORI
    filtered_kategori_df = frek_alat_df
    if selected_dinas != "Semua":
        filtered_kategori_df = frek_alat_df[frek_alat_df['DINAS'] == selected_dinas]
    
    unique_kategori = ["Semua"] + sorted(list(filtered_kategori_df['KATEGORI'].dropna().unique()))
    selected_kategori = st.sidebar.selectbox("KATEGORI", unique_kategori)
    
    # Filter untuk BAND
    filtered_band_df = filtered_kategori_df
    if selected_kategori != "Semua":
        filtered_band_df = filtered_kategori_df[filtered_kategori_df['KATEGORI'] == selected_kategori]
    
    unique_band = ["Semua"] + sorted(list(filtered_band_df['BAND'].dropna().unique()))
    selected_band = st.sidebar.selectbox("BAND", unique_band)
    
    # Filter untuk ZONA
    unique_zona = ["Semua"] + sorted(list(map(str, frek_alat_df['ZONA'].dropna().unique())))
    selected_zona = st.sidebar.selectbox("ZONA", unique_zona)
    
    # Filter untuk JML BULAN
    unique_jml_bulan = ["Semua"] + sorted(list(persentase_data.keys()))
    selected_jml_bulan = st.sidebar.selectbox("JML BULAN", unique_jml_bulan)
    
    # Pilihan untuk jenis pelanggaran
    st.markdown("<div class='subtitle'>Jenis Pelanggaran</div>", unsafe_allow_html=True)
    jenis_pelanggaran = st.radio(
        "Pilih Jenis Pelanggaran:",
        ["Pelanggaran Pertama", "Pelanggaran Berulang"],
        horizontal=True
    )
    
    # Tampilkan input untuk JUMLAH FREKUENSI dan JUMLAH PERANGKAT
    st.markdown("<div class='subtitle'>Input Jumlah Frekuensi & Perangkat</div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        jumlah_frekuensi = st.number_input("JUMLAH FREKUENSI", min_value=0, value=1, step=1)
    
    with col2:
        jumlah_perangkat = st.number_input("JUMLAH PERANGKAT", min_value=0, value=1, step=1)
    
    # Tombol untuk menghitung denda
    if st.button("Hitung Denda"):
        # Siapkan filter berdasarkan input pengguna
        filters = {
            'DINAS': selected_dinas,
            'KATEGORI': selected_kategori,
            'BAND': selected_band,
            'ZONA': selected_zona
        }
        
        # Tambahkan JENIS IZIN ke filter jika tersedia dan dipilih
        if 'JENIS IZIN' in frek_alat_df.columns and selected_jenis_izin != "Semua":
            filters['JENIS IZIN'] = selected_jenis_izin
        
        # Filter data FREK & ALAT
        filtered_df = filter_data(frek_alat_df, filters)
        
        # Jika tidak ada data yang sesuai filter tapi JENIS IZIN dipilih, buat data dummy
        if filtered_df.empty and selected_jenis_izin != "Semua":
            # Buat data dummy dengan JENIS IZIN yang dipilih
            dummy_data = {
                'JENIS IZIN': selected_jenis_izin,
                'MAKS POIN': get_maks_poin(selected_jenis_izin),
                'INDEKS PELANGGARAN PERTAMA': 1.0,
                'INDEKS PELANGGARAN BERULANG': 1.5,
                '%': 1.0,
                'TARIF DENDA': 0  # Default 0, bisa diubah sesuai kebutuhan
            }
            
            # Tambahkan filter lain yang dipilih
            for key, value in filters.items():
                if value != "Semua" and key != 'JENIS IZIN':
                    dummy_data[key] = value
            
            # Buat DataFrame dummy
            filtered_df = pd.DataFrame([dummy_data])
            
            st.warning(f"""
            Tidak ada data yang sesuai dengan filter yang dipilih.
            Menggunakan data default untuk JENIS IZIN '{selected_jenis_izin}' dengan MAKS POIN {get_maks_poin(selected_jenis_izin)}.
            """)
        
        if not filtered_df.empty:
            # Dapatkan persentase berdasarkan JML BULAN
            persentase = get_percentage(persentase_data, selected_jml_bulan) if selected_jml_bulan != "Semua" else 1.0
            
            # Ambil data pertama dari hasil filter
            selected_data = filtered_df.iloc[0]
            
            # Tambahkan JENIS IZIN ke selected_data jika belum ada
            if 'JENIS IZIN' not in selected_data and selected_jenis_izin != "Semua":
                selected_data['JENIS IZIN'] = selected_jenis_izin
            
            # Tambahkan MAKS POIN sesuai JENIS IZIN jika belum ada atau 0
            jenis_izin = str(selected_data.get('JENIS IZIN', '')).strip().upper()
            if 'MAKS POIN' not in selected_data or pd.isna(selected_data.get('MAKS POIN')) or selected_data.get('MAKS POIN') == 0:
                selected_data['MAKS POIN'] = get_maks_poin(jenis_izin)
            
            # Pastikan INDEKS PELANGGARAN ada
            if 'INDEKS PELANGGARAN PERTAMA' not in selected_data:
                selected_data['INDEKS PELANGGARAN PERTAMA'] = 1.0
            if 'INDEKS PELANGGARAN BERULANG' not in selected_data:
                selected_data['INDEKS PELANGGARAN BERULANG'] = 1.5
            
            # Hitung TOTAL POIN, DENDA, dan TOTAL TAGIHAN DENDA
            hasil_perhitungan = calculate_denda(
                selected_data, 
                jumlah_frekuensi, 
                jumlah_perangkat, 
                persentase,
                jenis_pelanggaran
            )
            
            # Debug info
            with debug_expander:
                st.markdown("### Data Terpilih:")
                st.write(selected_data)
                
                st.markdown("### Hasil Perhitungan:")
                st.write(hasil_perhitungan)
                
                st.markdown("### Parameter Input:")
                st.write(f"JENIS IZIN: {selected_data.get('JENIS IZIN', 'N/A')}")
                st.write(f"Jenis Pelanggaran: {jenis_pelanggaran}")
                st.write(f"Persentase JML BULAN: {persentase}")
                st.write(f"Jumlah Frekuensi: {jumlah_frekuensi}")
                st.write(f"Jumlah Perangkat: {jumlah_perangkat}")
            
            # Tampilkan hasil
            st.markdown("<div class='subtitle'>Hasil Perhitungan Denda</div>", unsafe_allow_html=True)
            
            # Tampilkan informasi JENIS IZIN dan MAKS POIN
            jenis_izin_info = f"""
            <div class='jenis-izin-box'>
                <p><strong>JENIS IZIN:</strong> {jenis_izin}</p>
                <p><strong>MAKS POIN Default:</strong> {get_maks_poin(jenis_izin)}</p>
                <p><strong>MAKS POIN yang digunakan:</strong> {hasil_perhitungan['maks_poin']}</p>
                <p><strong>Jenis Pelanggaran:</strong> {jenis_pelanggaran}</p>
            </div>
            """
            st.markdown(jenis_izin_info, unsafe_allow_html=True)
            
            # Tampilkan informasi perhitungan
            formula_text = f"""
            <p><strong>Formula Perhitungan:</strong></p>
            <ol>
                <li>TOTAL POIN = INDEKS PELANGGARAN ({hasil_perhitungan['indeks']}) * % ({hasil_perhitungan['persentase']*100:.0f}%) * MAKS POIN ({hasil_perhitungan['maks_poin']}) = {hasil_perhitungan['total_poin']:.2f}</li>
                <li>DENDA = TOTAL POIN ({hasil_perhitungan['total_poin']:.2f}) * TARIF DENDA ({hasil_perhitungan['tarif_denda']:.2f}) = {hasil_perhitungan['denda']:.2f}</li>
                <li>TOTAL TAGIHAN DENDA = DENDA ({hasil_perhitungan['denda']:.2f}) * JUMLAH FREKUENSI ({jumlah_frekuensi}) * JUMLAH PERANGKAT ({jumlah_perangkat}) = {hasil_perhitungan['total_tagihan_denda']:.2f}</li>
            </ol>
            """
            
            st.markdown(f"""
            <div class='calculation-info'>
                {formula_text}
                <p><strong>Filter yang Digunakan:</strong> {', '.join([f"{k}: {v}" for k, v in filters.items() if v != 'Semua'])}</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Tambahkan kolom hasil ke dataframe untuk visualisasi
            result_df = filtered_df.copy()
            result_df['JENIS PELANGGARAN'] = jenis_pelanggaran
            result_df['INDEKS YANG DIGUNAKAN'] = hasil_perhitungan['indeks']
            result_df['PERSENTASE'] = hasil_perhitungan['persentase']
            result_df['TOTAL POIN'] = hasil_perhitungan['total_poin']
            result_df['DENDA'] = hasil_perhitungan['denda']
            result_df['JUMLAH FREKUENSI'] = jumlah_frekuensi
            result_df['JUMLAH PERANGKAT'] = jumlah_perangkat
            result_df['TOTAL TAGIHAN DENDA'] = hasil_perhitungan['total_tagihan_denda']
            
            # Pastikan JENIS IZIN ada di result_df
            if 'JENIS IZIN' not in result_df.columns and selected_jenis_izin != "Semua":
                result_df['JENIS IZIN'] = selected_jenis_izin
            
            # Pastikan MAKS POIN sesuai dengan JENIS IZIN
            if 'MAKS POIN' not in result_df.columns or result_df['MAKS POIN'].iloc[0] == 0:
                result_df['MAKS POIN'] = hasil_perhitungan['maks_poin']
            
            # Tentukan kolom yang akan ditampilkan
            display_columns = [
                'JENIS IZIN', 'DINAS', 'KATEGORI', 'BAND', 'ZONA',
                'JENIS PELANGGARAN', 'INDEKS YANG DIGUNAKAN', 'PERSENTASE',
                'MAKS POIN', 'TOTAL POIN', 'TARIF DENDA', 'DENDA', 
                'JUMLAH FREKUENSI', 'JUMLAH PERANGKAT', 'TOTAL TAGIHAN DENDA'
            ]
            
            # Pastikan semua kolom yang dibutuhkan ada
            display_columns = [col for col in display_columns if col in result_df.columns]
            
            st.dataframe(result_df[display_columns])
            
            # Tampilkan total denda
            st.markdown(f"""
            <div class='result-container'>
                <h3>Total Tagihan Denda: Rp {hasil_perhitungan['total_tagihan_denda']:,.2f}</h3>
            </div>
            """, unsafe_allow_html=True)
            
            # Visualisasi data
            st.markdown("<div class='subtitle'>Visualisasi Data</div>", unsafe_allow_html=True)
            
            # Tentukan komponen untuk visualisasi - Diagram Alir Perhitungan
            flow_components = {
                'Indeks Pelanggaran': hasil_perhitungan['indeks'],
                '% Faktor': hasil_perhitungan['persentase'],
                'MAKS POIN': hasil_perhitungan['maks_poin'],
                'TOTAL POIN': hasil_perhitungan['total_poin'],
                'TARIF DENDA': hasil_perhitungan['tarif_denda'],
                'DENDA': hasil_perhitungan['denda'],
                'Jumlah Frekuensi': jumlah_frekuensi,
                'Jumlah Perangkat': jumlah_perangkat
            }
            
            # Buat dataframe untuk visualisasi aliran perhitungan
            flow_df = pd.DataFrame({
                'Komponen': list(flow_components.keys()),
                'Nilai': list(flow_components.values())
            })
            
            # Grafik batang untuk komponen perhitungan
            fig1 = px.bar(
                flow_df,
                x='Komponen',
                y='Nilai',
                title='Komponen Perhitungan Denda',
                color='Komponen'
            )
            
            # Atur urutan komponen sesuai alur perhitungan
            component_order = ['Indeks Pelanggaran', '% Faktor', 'MAKS POIN', 'TOTAL POIN', 'TARIF DENDA', 'DENDA', 'Jumlah Frekuensi', 'Jumlah Perangkat']
            fig1.update_xaxes(categoryorder='array', categoryarray=component_order)
            
            st.plotly_chart(fig1, use_container_width=True)
            
            # Grafik pie untuk proporsi komponen dalam hasil akhir
            prop_components = {
                'TOTAL POIN': hasil_perhitungan['total_poin'],
                'TARIF DENDA': hasil_perhitungan['tarif_denda'],
                'Jumlah Frekuensi': jumlah_frekuensi,
                'Jumlah Perangkat': jumlah_perangkat
            }
            
            # Buat dataframe untuk visualisasi proporsi
            prop_df = pd.DataFrame({
                'Komponen': list(prop_components.keys()),
                'Nilai': list(prop_components.values())
            })
            
            # Hitung total nilai untuk proporsi
            total_prop = sum(prop_components.values())
            prop_df['Proporsi'] = prop_df['Nilai'] / total_prop if total_prop > 0 else 0
            
            # Grafik pie untuk proporsi komponen
            fig2 = px.pie(
                prop_df,
                values='Proporsi',
                names='Komponen',
                title='Proporsi Komponen dalam Perhitungan'
            )
            st.plotly_chart(fig2, use_container_width=True)
            
            # Grafik Sankey untuk alur perhitungan
            st.markdown("### Alur Perhitungan Denda")
            
            # Membuat label untuk Sankey diagram
            labels = [
                'Indeks', 'Persentase', 'MAKS POIN', 
                'TOTAL POIN', 'TARIF DENDA', 'DENDA',
                'Jumlah Frekuensi', 'Jumlah Perangkat', 'TOTAL TAGIHAN DENDA'
            ]
            
            # Membuat source dan target untuk Sankey diagram
            source = [0, 1, 2, 3, 4, 5, 6, 7]  # dari mana aliran berasal
            target = [3, 3, 3, 5, 5, 8, 8, 8]  # ke mana aliran menuju
            
            # Nilai untuk aliran (dapat disesuaikan untuk visualisasi yang lebih baik)
            values = [
                hasil_perhitungan['indeks'],
                hasil_perhitungan['persentase'],
                hasil_perhitungan['maks_poin'],
                hasil_perhitungan['total_poin'],
                hasil_perhitungan['tarif_denda'],
                hasil_perhitungan['denda'],
                jumlah_frekuensi,
                jumlah_perangkat
            ]
            
            # Normalkan nilai untuk visualisasi yang lebih baik
            max_value = max(values) if max(values) > 0 else 1
            normalized_values = [v / max_value * 100 for v in values]
            
            # Buat diagram Sankey
            fig3 = go.Figure(data=[go.Sankey(
                node=dict(
                    pad=15,
                    thickness=20,
                    line=dict(color="black", width=0.5),
                    label=labels,
                    color="blue"
                ),
                link=dict(
                    source=source,
                    target=target,
                    value=normalized_values
                )
            )])
            
            fig3.update_layout(title_text="Alur Perhitungan Denda", font_size=10)
            st.plotly_chart(fig3, use_container_width=True)
            
            # Opsi untuk download hasil perhitungan
            st.markdown("<div class='subtitle'>Download Hasil</div>", unsafe_allow_html=True)
            
            # Download sebagai CSV untuk menghindari masalah dengan Excel engine
            csv_data = to_csv(result_df)
            st.download_button(
                label="Download Hasil Perhitungan (CSV)",
                data=csv_data,
                file_name=f"hasil_perhitungan_denda_{jenis_izin}.csv",
                mime="text/csv"
            )
            
            # Coba download Excel jika engine tersedia
            excel_data, excel_success = to_excel(result_df)
            if excel_success:
                st.download_button(
                    label="Download Hasil Perhitungan (Excel)",
                    data=excel_data,
                    file_name=f"hasil_perhitungan_denda_{jenis_izin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="excel_download"
                )
            else:
                st.info("""
                Untuk download Excel, silakan install paket 'openpyxl' atau 'xlsxwriter': 
                `pip install openpyxl` atau `pip install xlsxwriter`
                """)
        else:
            st.warning("Tidak ada data yang sesuai dengan filter yang dipilih. Pilih JENIS IZIN untuk melanjutkan perhitungan.")
else:
    st.info("Silakan pilih dan muat data terlebih dahulu untuk melanjutkan perhitungan denda.")

# Tampilkan informasi di bagian bawah
st.markdown("""
<div class='highlight'>
    <h4>Petunjuk Penggunaan:</h4>
    <ol>
        <li>Pilih file Excel dari folder 'Data' yang berisi data perhitungan denda</li>
        <li>Klik tombol "Muat Data" untuk memproses file</li>
        <li>Pilih JENIS IZIN (IPFR, ISR, atau APT) untuk menggunakan nilai MAKS POIN yang sesuai</li>
        <li>Pilih jenis pelanggaran (Pertama atau Berulang)</li>
        <li>Gunakan filter di sidebar untuk memilih data berdasarkan DINAS, KATEGORI, BAND, ZONA, dan JML BULAN</li>
        <li>Masukkan JUMLAH FREKUENSI dan JUMLAH PERANGKAT</li>
        <li>Klik tombol "Hitung Denda" untuk melihat hasil perhitungan</li>
        <li>Download hasil perhitungan dalam format CSV atau Excel jika diperlukan</li>
    </ol>
    <p><strong>Catatan Formula Perhitungan:</strong></p>
    <ol>
        <li>TOTAL POIN = INDEKS PELANGGARAN * % * MAKS POIN</li>
        <li>DENDA = TOTAL POIN * TARIF DENDA</li>
        <li>TOTAL TAGIHAN DENDA = DENDA * JUMLAH FREKUENSI * JUMLAH PERANGKAT</li>
    </ol>
    <p>Sistem akan menggunakan nilai MAKS POIN berdasarkan JENIS IZIN (IPFR=600.000, ISR=7.000, APT=5.000)</p>
</div>
""", unsafe_allow_html=True)

# Informasi folder data
st.markdown("""
<div class='info-box'>
    <p><strong>Informasi Folder Data:</strong></p>
    <p>Aplikasi ini secara otomatis membaca file Excel (.xlsx, .xls) dari folder 'Data' di direktori yang sama dengan aplikasi.</p>
    <p>Untuk menambahkan data baru, cukup letakkan file Excel Anda di folder tersebut.</p>
    <p>File harus berisi setidaknya sheet 'FREK & ALAT' dengan kolom yang sesuai.</p>
</div>
""", unsafe_allow_html=True)

# Footer
st.markdown("""
<div style='text-align: center; margin-top: 30px; padding: 10px; color: #604CC3;'>
    <p>© 2025 Aplikasi Simulasi Perhitungan Denda | Loka Monitor SFR Kendari</p>
</div>
""", unsafe_allow_html=True)