import streamlit as st
import pandas as pd
from io import BytesIO
import random
import os

st.set_page_config(page_title="Pemetaan Petir PPU", layout="wide")
st.title("📊 Aplikasi Pengolahan Data Petir PPU")

# Inisialisasi session_state
if "uploaded_files" not in st.session_state:
    st.session_state["uploaded_files"] = []
if "gabung_log" not in st.session_state:
    st.session_state["gabung_log"] = ""
if "reset_flag" not in st.session_state:
    st.session_state["reset_flag"] = False
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = "uploader_1"

# Pilihan bulan dengan nomor di depan
list_bulan = [
    f"{str(i+1).zfill(2)}{bulan}" for i, bulan in enumerate([
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ])
]

# Master daftar lokasi
list_lokasi = [
    "Api-api", "Argo Mulyo", "Babulu Darat", "Babulu Laut", "Bangun Mulya", "Binuang", "Bukit Raya",
    "Bukit Subur", "Buluminung", "Bumi Harapan", "Gersik", "Giri Mukti", "Giri Purwa", "Gunung Intan",
    "Gunung Makmur", "Gunung Mulia", "Gunung Seteleng", "Jenebora", "Kampung Baru", "Karang Jinawi",
    "Labangka", "Labangka Barat", "Lawe-lawe", "Maridan", "Mentawir", "Nenang", "Nipah-Nipah",
    "Pantai Lango", "Pejala", "Pemaluan", "Penajam", "Petung", "Rawa Mulia", "Riko", "Rintik",
    "Salo Loang", "Sebakung Jaya", "Semoi Dua", "Sepaku", "Sepan", "Sesulu", "Sesumpu", "Sidorejo",
    "Sotek", "Sri Raharja", "Suka Raja", "Suko Mulyo", "Sumber Sari", "Sungai Parit", "Tanjung Tengah",
    "Telemow", "Tengin Baru", "Waru", "Wono Sari"
]

df_master = pd.DataFrame({"Nama Lokasi": list_lokasi})

# Tab aplikasi
tab1, tab2, tab3 = st.tabs([
    "📁 Gabungkan File Excel",
    "📌 Rapikan Data CG+ / CG-",
    "ℹ️ Penjelasan & Panduan"
])

# ─── Tab 1: Gabungkan Excel ───
with tab1:
    st.header("📁 Gabungkan Banyak File Excel")

    uploaded_files = st.file_uploader(
        "Upload beberapa file Excel (.xlsx/.xls)",
        accept_multiple_files=True,
        type=["xlsx", "xls"],
        key=st.session_state["uploader_key"]
    )

    if uploaded_files:
        st.session_state["uploaded_files"] = uploaded_files
        st.session_state["reset_flag"] = False

    st.write(f"📦 {len(st.session_state['uploaded_files'])} file terupload.")

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        hapus_semua = st.button("🗑️ Hapus Semua File")
    with col2:
        reset_file = st.button("🔄 Reset Setelah Download")
    with col3:
        output_filename1 = st.text_input(
            "Nama file output (tanpa ekstensi .xlsx)",
            value="TotalGabungan"
        )

    if hapus_semua or reset_file or st.session_state["reset_flag"]:
        st.session_state["uploaded_files"] = []
        st.session_state["gabung_log"] = ""
        st.session_state["uploader_key"] = f"uploader_{random.randint(1000,9999)}"
        st.session_state["reset_flag"] = False
        st.success("✅ Semua file telah dihapus.")

    # Tambahkan variabel untuk preview data gabungan
    preview_combined = None
    error_gabungan = None

    if st.button("Gabungkan"):
        st.session_state["gabung_log"] = ""
        preview_combined = None
        error_gabungan = None

        if not st.session_state["uploaded_files"]:
            st.warning("⚠️ Harap upload minimal satu file Excel.")
        else:
            all_data = []
            for file in st.session_state["uploaded_files"]:
                filename = file.name.lower()
                ext = os.path.splitext(filename)[-1]
                try:
                    if ext == ".xls":
                        try:
                            xls = pd.ExcelFile(file, engine="xlrd")
                        except ImportError:
                            st.session_state["gabung_log"] += (
                                f"❌ {file.name} - Modul 'xlrd' belum terinstall. "
                                "Install dengan perintah: pip install xlrd\n"
                            )
                            continue
                    else:
                        xls = pd.ExcelFile(file, engine="openpyxl")
                except Exception as e:
                    st.session_state["gabung_log"] += f"❌ {file.name} - Gagal membaca file: {type(e).__name__}: {e}\n"
                    continue
                for sheet in xls.sheet_names:
                    try:
                        if ext == ".xls":
                            try:
                                df = pd.read_excel(xls, sheet, engine="xlrd")
                            except ImportError:
                                st.session_state["gabung_log"] += (
                                    f"❌ {file.name} - Modul 'xlrd' belum terinstall. "
                                    "Install dengan perintah: pip install xlrd\n"
                                )
                                continue
                        else:
                            df = pd.read_excel(xls, sheet, engine="openpyxl")
                    except Exception as e:
                        st.session_state["gabung_log"] += f"❌ {file.name} - Sheet '{sheet}' gagal dibaca: {type(e).__name__}: {e}\n"
                        continue
                    if df.empty:
                        st.session_state["gabung_log"] += f"⚠️ {file.name} - Sheet '{sheet}' kosong, dilewati.\n"
                        continue
                    all_data.append(df)
                    st.session_state["gabung_log"] += f"✅ {file.name} - Sheet '{sheet}' ({len(df)} baris)\n"

            if not all_data:
                st.session_state["gabung_log"] += "\n❌ Tidak ada data yang berhasil digabung."
                error_gabungan = "Tidak ada data yang berhasil digabung."
            else:
                combined = pd.concat(all_data, ignore_index=True)
                total = len(combined)
                st.session_state["gabung_log"] += f"\n📊 Total baris gabungan: {total}"

                # Preview 10 baris pertama
                preview_combined = combined.head(10)

                max_rows = 65000
                num_sheets = (total // max_rows) + 1
                output = BytesIO()

                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for i in range(num_sheets):
                        part = combined.iloc[i*max_rows:(i+1)*max_rows]
                        part.to_excel(writer, index=False, sheet_name=f"Data_{i+1}")

                output.seek(0)
                # Gunakan nama file dari input user
                filename = f"{output_filename1.strip() or 'TotalGabungan'}.xlsx"
                st.download_button(
                    "📥 Download File Gabungan",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=lambda: st.session_state.update({"reset_flag": True})
                )

    if st.session_state["gabung_log"]:
        st.text_area("Log Proses", st.session_state["gabung_log"], height=200)

    # Tampilkan preview data gabungan jika ada
    if 'preview_combined' in locals() and preview_combined is not None:
        st.markdown("#### Preview Data Gabungan (10 baris pertama):")
        st.dataframe(preview_combined)
    elif 'error_gabungan' in locals() and error_gabungan:
        st.error(error_gabungan)

# ─── Tab 2: Rapikan CG+/CG- ───
with tab2:
    st.header("📌 Rapikan Data CG+ dan CG-")
    file = st.file_uploader("Upload File Excel Data Petir", type=["xlsx", "xls"], key="cg_file")

    output_filename2 = st.text_input(
        "Nama file output (tanpa ekstensi .xlsx)",
        value="HasilRapi",
        key="output_filename2"
    )

    # Tambahkan variabel untuk preview hasil pivot
    preview_pivot = None

    if st.button("Proses Pivot"):
        if not file:
            st.warning("⚠️ Upload file terlebih dahulu.")
        else:
            with st.spinner("Memproses data..."):
                try:
                    filename = file.name.lower()
                    ext = os.path.splitext(filename)[-1]
                    if ext == ".xls":
                        try:
                            df = pd.read_excel(file, engine="xlrd")
                        except ImportError:
                            st.error("❌ Modul 'xlrd' belum terinstall. Install dengan perintah: pip install xlrd")
                            raise st.stop()
                    else:
                        df = pd.read_excel(file, engine="openpyxl")

                    # Ganti 'NAMOBJ' menjadi 'Kelurahan'
                    if not {'Kelurahan', 'Jenis', 'FREQUENCY'}.issubset(df.columns):
                        st.error("❌ Kolom wajib: Kelurahan, Jenis, FREQUENCY")
                    else:
                        pivot_df = df.pivot_table(
                            index='Kelurahan',
                            columns='Jenis',
                            values='FREQUENCY',
                            aggfunc='sum'
                        ).reset_index()

                        # Rename kolom menjadi CG+ dan CG-
                        pivot_df.columns.name = None
                        pivot_df = pivot_df.rename(columns={
                            'Positive Cloud to Ground': 'CG+',
                            'Negative Cloud to Ground': 'CG-',
                            'Kelurahan': 'Nama Lokasi'
                        })

                        # Ambil daftar kelurahan unik dari file yang di-upload
                        lokasi_unik = df['Kelurahan'].drop_duplicates().sort_values().reset_index(drop=True)
                        df_lokasi = pd.DataFrame({"Nama Lokasi": lokasi_unik})

                        # Merge dengan daftar kelurahan dari file, bukan dari master tetap
                        result = pd.merge(df_lokasi, pivot_df, on='Nama Lokasi', how='left').fillna(0)

                        # Pastikan tipe data angka
                        result['CG+'] = result['CG+'].astype(int)
                        result['CG-'] = result['CG-'].astype(int)

                        # Tambahkan nomor baris mulai dari 1
                        result.insert(0, "No", range(1, len(result)+1))

                        # Urutkan kolom supaya CG+ di depan
                        result = result[["No", "Nama Lokasi", "CG+", "CG-"]]

                        # Tampilkan preview hasil pivot (10 baris pertama)
                        preview_pivot = result.head(10)
                        st.markdown("#### Preview Data Hasil Pivot (10 baris pertama):")
                        st.dataframe(preview_pivot)

                        output2 = BytesIO()
                        with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                            result.to_excel(writer, index=False)
                        output2.seek(0)

                        # Gunakan nama file dari input user
                        filename2 = f"{output_filename2.strip() or 'HasilRapi'}.xlsx"
                        st.download_button(
                            "📥 Download Data Rapi",
                            data=output2,
                            file_name=filename2,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ Terjadi error saat membaca file: {type(e).__name__}: {e}")

# ─── Tab 3: Penjelasan & Panduan ───
with tab3:
    st.header("ℹ️ Penjelasan & Panduan Penggunaan")

    st.markdown("""
### 📁 Gabungkan File Excel
- **Fitur:** Menggabungkan beberapa file Excel (.xlsx/.xls) menjadi satu file.
- **Catatan:** File yang digabungkan akan disatukan semua sheet-nya. Jika jumlah baris sangat banyak, hasil akan dibagi ke beberapa sheet (maksimal 65.000 baris per sheet).
- **Langkah:**  
    1. Upload beberapa file Excel.  
    2. Klik **Gabungkan**.  
    3. Download hasil gabungan.

---

### 📌 Rapikan Data CG+ / CG-
- **Fitur:** Merapikan data petir berdasarkan kelurahan dan jenis petir (CG+ dan CG-).
- **Format Data:**  
    File Excel yang di-upload **wajib** memiliki kolom berikut:
    - `Kelurahan` : Nama lokasi/kelurahan.
    - `Jenis` : Jenis petir, misal "Positive Cloud to Ground" atau "Negative Cloud to Ground".
    - `FREQUENCY` : Jumlah kejadian.
- **Langkah:**  
    1. Upload file Excel dengan format kolom di atas.  
    2. Klik **Proses Pivot**.  
    3. Download hasil data rapi.

---

### ❓ Tips & Catatan
- Jika file `.xls` gagal dibaca, pastikan modul `xlrd` sudah terinstall.
- Nama kolom **harus persis** seperti di atas (case-sensitive).
- Jika ada error, cek log/error yang muncul di aplikasi.
- Untuk pertanyaan lebih lanjut, hubungi admin aplikasi.

""")
