"""
==============================================================================
APLIKASI SISTEM AKUMULASI DATA MULTI-FILE EXCEL
==============================================================================
Tujuan   : Membantu saya mengakumulasi (menjumlahkan secara element-wise)
           beberapa file Excel (.xlsx) yang memiliki struktur identik,
           sekaligus meminimalisir human error melalui validasi otomatis,
           penanganan missing values, dan penghapusan duplikasi.

Cara jalankan:
    pip install streamlit pandas openpyxl
    streamlit run app_akumulasi_data.py
==============================================================================
"""

import io
import functools
from datetime import datetime

import pandas as pd
import streamlit as st

# ==============================================================================
# KONFIGURASI HALAMAN
# ==============================================================================
st.set_page_config(
    page_title="Sistem Akumulasi Data Saya",
    page_icon="🔢",
    layout="wide",
)


# ==============================================================================
# FASE 1 — FUNGSI INPUT (DATA INGESTION)
# ==============================================================================
def load_uploaded_files(uploaded_files):
    """
    Membaca seluruh file Excel yang saya unggah ke dalam list of tuple
    berisi (nama_file, dataframe).
    Mengembalikan None apabila salah satu file gagal dibaca (corrupt/format salah).
    """
    dfs_info = []
    for f in uploaded_files:
        try:
            df = pd.read_excel(f)
            dfs_info.append((f.name, df))
        except Exception as e:
            st.error(f"Saya gagal membaca file **{f.name}**. Detail error: `{e}`")
            return None
    return dfs_info


def validate_structure(dfs_info):
    """
    Memvalidasi bahwa seluruh file memiliki struktur identik:
    - Nama & urutan kolom harus sama persis dengan file acuan (file pertama)
    - Jumlah baris harus sama persis dengan file acuan

    Validasi ini penting karena keseluruhan proses akumulasi saya bergantung
    pada asumsi bahwa setiap file punya dimensi yang identik (sesuai
    spesifikasi proyek saya untuk meminimalisir human error).
    """
    base_filename, base_df = dfs_info[0]
    base_columns = list(base_df.columns)
    base_row_count = base_df.shape[0]

    for filename, df in dfs_info[1:]:
        if list(df.columns) != base_columns:
            pesan = (
                f"Saya menemukan ketidaksesuaian nama/urutan kolom antara file "
                f"**{base_filename}** dan **{filename}**. Saya tidak dapat "
                f"melanjutkan proses sebelum struktur kolom benar-benar identik."
            )
            return False, pesan
        if df.shape[0] != base_row_count:
            pesan = (
                f"Saya menemukan perbedaan jumlah baris: file **{base_filename}** "
                f"memiliki {base_row_count} baris, sedangkan file **{filename}** "
                f"memiliki {df.shape[0]} baris. Saya tidak dapat melanjutkan proses "
                f"akumulasi sebelum jumlah baris disamakan."
            )
            return False, pesan

    return True, ""


# ==============================================================================
# FASE 2 — PEMBERSIHAN AWAL (PRE-CLEANSING)
# ==============================================================================
def pre_cleansing(dfs_info, numeric_cols):
    """
    Untuk setiap dataset yang saya unggah:
    1. Mengonversi kolom numerik ke tipe numerik yang valid
       (nilai yang tidak bisa dikonversi otomatis menjadi NaN -> errors='coerce').
    2. Mendeteksi sel kosong/NaN pada kolom numerik.
    3. Mengisi seluruh sel kosong tersebut dengan angka 0 agar perhitungan
       element-wise di fase berikutnya aman dari nilai null.

    Mengembalikan: list dataframe yang sudah bersih, total sel kosong yang
    saya isi, dan log detail per file.
    """
    cleaned_dfs = []
    total_missing = 0
    detail_logs = []

    for filename, df in dfs_info:
        df_copy = df.copy()

        # Pastikan kolom numerik benar-benar bertipe numerik
        for col in numeric_cols:
            df_copy[col] = pd.to_numeric(df_copy[col], errors="coerce")

        # Deteksi jumlah sel kosong (NaN) pada kolom numerik
        missing_count = int(df_copy[numeric_cols].isna().sum().sum())
        total_missing += missing_count
        if missing_count > 0:
            detail_logs.append(
                f"- File **{filename}**: {missing_count} sel kosong/tidak valid "
                f"saya isi dengan nilai 0."
            )

        # Isi nilai kosong dengan 0 (penanganan missing values)
        df_copy[numeric_cols] = df_copy[numeric_cols].fillna(0)

        cleaned_dfs.append((filename, df_copy))

    return cleaned_dfs, total_missing, detail_logs


def check_identity_consistency(dfs_info, identity_cols):
    """
    Pengecekan tambahan (langkah ekstra untuk meminimalisir human error):
    memastikan nilai pada kolom identitas/kategori (misal: Nama, Kode, Kategori)
    konsisten di seluruh file yang saya unggah. Jika berbeda, saya tetap
    melanjutkan proses menggunakan nilai dari file acuan, namun saya beri
    peringatan agar bisa diperiksa kembali secara manual.
    """
    if not identity_cols:
        return None

    base_filename, base_df = dfs_info[0]
    base_identity = base_df[identity_cols].reset_index(drop=True)

    file_bermasalah = []
    for filename, df in dfs_info[1:]:
        compare_identity = df[identity_cols].reset_index(drop=True)
        if not base_identity.equals(compare_identity):
            file_bermasalah.append(filename)

    if file_bermasalah:
        return (
            f"Saya mendeteksi bahwa nilai pada kolom identitas berbeda antara "
            f"file acuan (**{base_filename}**) dengan file: {', '.join(file_bermasalah)}. "
            f"Saya tetap memproses menggunakan nilai dari file acuan, namun saya "
            f"sarankan untuk memeriksa kembali konsistensi data tersebut."
        )
    return None


# ==============================================================================
# FASE 3 — KOMPUTASI (ELEMENT-WISE ADDITION)
# ==============================================================================
def compute_accumulation(cleaned_dfs, identity_cols, numeric_cols, original_columns):
    """
    Menjumlahkan nilai numerik dari seluruh dataset secara element-wise
    berdasarkan posisi baris & kolom (karena dimensi setiap dataset
    dipastikan identik). Kolom identitas/kategori diambil dari file acuan
    (file pertama) karena nilainya bersifat deskriptif, bukan numerik
    yang dijumlahkan.
    """
    base_filename, base_df = cleaned_dfs[0]

    # Ambil kolom identitas dari file acuan, reset index agar posisinya selaras
    if identity_cols:
        identity_df = base_df[identity_cols].reset_index(drop=True)
    else:
        identity_df = pd.DataFrame(index=range(base_df.shape[0]))

    # Kumpulkan kolom numerik dari setiap file, lalu jumlahkan secara element-wise
    numeric_frames = [
        df[numeric_cols].reset_index(drop=True) for _, df in cleaned_dfs
    ]
    summed_numeric = functools.reduce(lambda a, b: a.add(b, fill_value=0), numeric_frames)

    # Gabungkan kembali kolom identitas dengan hasil penjumlahan numerik
    final_df = pd.concat([identity_df, summed_numeric], axis=1)

    # Urutkan kembali kolom sesuai urutan asli pada file sumber
    final_df = final_df[original_columns]

    return final_df


# ==============================================================================
# FASE 4 — PEMBERSIHAN AKHIR (POST-CLEANSING UNTUK DUPLIKASI)
# ==============================================================================
def post_cleansing(final_df):
    """
    Mendeteksi baris yang terduplikasi secara identik pada dataframe final
    hasil akumulasi, lalu menghapusnya (menyisakan kemunculan pertama).
    """
    duplicate_mask = final_df.duplicated(keep="first")
    duplicate_count = int(duplicate_mask.sum())
    final_df_clean = final_df.drop_duplicates(keep="first").reset_index(drop=True)
    return final_df_clean, duplicate_count


# ==============================================================================
# FASE 5 — DASHBOARD OUTPUT & VISUALISASI
# ==============================================================================
def convert_df_to_excel_bytes(df):
    """Mengonversi dataframe final saya menjadi file Excel (.xlsx) dalam bentuk bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data Akumulasi")
    return buffer.getvalue()


def render_dashboard(final_df, numeric_cols):
    """Menampilkan tabel hasil akhir, metrik, grafik batang, dan tombol download."""

    st.subheader("📋 Dataframe Final Hasil Akumulasi Saya")
    st.write(
        "Berikut adalah dataframe final saya yang sudah bersih dari nilai kosong "
        "dan duplikasi, serta sudah terakumulasi secara element-wise:"
    )
    st.dataframe(final_df, use_container_width=True)
    st.caption(f"Dimensi dataframe final saya: {final_df.shape[0]} baris × {final_df.shape[1]} kolom.")

    st.divider()
    st.subheader("📊 Dashboard Analitik Ringkas")

    # --- Metrik total per kolom numerik ---
    st.write("Saya rangkum nilai total untuk setiap kolom numerik pada dataset akumulasi saya:")
    totals = final_df[numeric_cols].sum()

    kolom_per_baris = 4
    for i in range(0, len(numeric_cols), kolom_per_baris):
        baris_kolom = st.columns(kolom_per_baris)
        potongan = numeric_cols[i: i + kolom_per_baris]
        for j, nama_kolom in enumerate(potongan):
            baris_kolom[j].metric(label=nama_kolom, value=f"{totals[nama_kolom]:,.2f}")

    # --- Grafik batang tren akumulasi per kolom ---
    st.write("Saya visualisasikan total akumulasi setiap kolom numerik dalam grafik batang berikut:")
    chart_df = totals.reset_index()
    chart_df.columns = ["Kolom", "Total Akumulasi"]
    st.bar_chart(chart_df.set_index("Kolom"))

    # --- Opsional: tren pertumbuhan kumulatif per baris (cumsum) ---
    with st.expander("➕ Opsional: Tren Pertumbuhan Kumulatif per Baris"):
        st.write(
            "Saya juga menyediakan grafik tambahan untuk melihat bagaimana nilai "
            "saya bertumbuh secara kumulatif mengikuti urutan baris pada dataset:"
        )
        cumulative_df = final_df[numeric_cols].cumsum()
        st.line_chart(cumulative_df)

    st.divider()

    # --- Tombol download ---
    st.subheader("⬇️ Unduh Dataset Final Saya")
    excel_bytes = convert_df_to_excel_bytes(final_df)
    nama_file = f"hasil_akumulasi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        label="📥 Unduh sebagai Excel (.xlsx)",
        data=excel_bytes,
        file_name=nama_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ==============================================================================
# ALUR UTAMA APLIKASI (MAIN)
# ==============================================================================
def main():
    st.title("🔢 Sistem Akumulasi Data Multi-File Saya")
    st.markdown(
        "Saya membangun aplikasi ini untuk mengakumulasi beberapa dataset Excel "
        "yang memiliki struktur identik, sekaligus meminimalisir risiko human "
        "error dalam proses penjumlahan data secara manual."
    )

    with st.sidebar:
        st.header("ℹ️ Tentang Alur Kerja Saya")
        st.markdown(
            """
            Saya merancang aplikasi ini dengan alur kerja berikut:
            1. **Input** — Saya unggah beberapa file `.xlsx` sekaligus.
            2. **Pre-Cleansing** — Saya isi otomatis sel kosong dengan 0.
            3. **Komputasi** — Saya jumlahkan seluruh dataset secara element-wise.
            4. **Post-Cleansing** — Saya hapus baris duplikat pada hasil akhir.
            5. **Dashboard** — Saya tampilkan hasil, metrik, grafik, dan tombol unduh.
            """
        )

    # --------------------------------------------------------------------
    # 1. FITUR INPUT
    # --------------------------------------------------------------------
    st.header("1️⃣ Unggah Dataset Saya")
    uploaded_files = st.file_uploader(
        "Unggah beberapa file Excel (.xlsx) yang ingin saya akumulasikan",
        type=["xlsx"],
        accept_multiple_files=True,
        help=(
            "Saya hanya menerima file berekstensi .xlsx. Pastikan setiap file "
            "memiliki jumlah baris serta nama dan jumlah kolom yang identik."
        ),
    )

    if not uploaded_files:
        st.info("Saya menunggu Anda mengunggah minimal 2 file Excel untuk memulai proses akumulasi.")
        return

    if len(uploaded_files) < 2:
        st.warning("Saya membutuhkan setidaknya 2 file untuk melakukan proses akumulasi data.")
        return

    with st.spinner("Saya sedang membaca seluruh file yang diunggah..."):
        dfs_info = load_uploaded_files(uploaded_files)

    if dfs_info is None:
        st.stop()

    # Validasi struktur dataset (baris & kolom harus identik)
    valid, pesan_error = validate_structure(dfs_info)
    if not valid:
        st.error(pesan_error)
        st.stop()

    st.success(
        f"Saya berhasil memvalidasi bahwa {len(dfs_info)} file memiliki struktur "
        f"(jumlah baris & kolom) yang identik. Saya dapat melanjutkan proses."
    )

    with st.expander("👁️ Pratinjau Data Mentah (File Pertama)"):
        st.dataframe(dfs_info[0][1].head(), use_container_width=True)

    # --------------------------------------------------------------------
    # KONFIGURASI KOLOM — kolom identitas vs kolom numerik
    # --------------------------------------------------------------------
    st.header("2️⃣ Konfigurasi Kolom Saya")
    st.write(
        "Saya perlu menentukan kolom mana yang berfungsi sebagai identitas/kategori "
        "(tidak dijumlahkan, contoh: Nama, Kode Cabang, Kategori) dan kolom mana "
        "yang berisi nilai numerik untuk saya akumulasikan."
    )

    semua_kolom = list(dfs_info[0][1].columns)
    default_identitas = dfs_info[0][1].select_dtypes(exclude="number").columns.tolist()

    identity_cols = st.multiselect(
        "Pilih kolom identitas/kategori (tidak dijumlahkan):",
        options=semua_kolom,
        default=default_identitas,
    )
    numeric_cols = [c for c in semua_kolom if c not in identity_cols]

    if not numeric_cols:
        st.error(
            "Saya tidak menemukan kolom numerik untuk dijumlahkan. Mohon periksa "
            "kembali pilihan kolom identitas saya di atas."
        )
        st.stop()

    st.caption(f"Saya akan menjumlahkan {len(numeric_cols)} kolom numerik: {', '.join(numeric_cols)}")

    proses = st.button("🚀 Proses Akumulasi Data Saya", type="primary")

    # Simpan hasil di session_state agar tidak hilang saat widget lain berubah
    if proses:
        # ------------------------------------------------------------
        # 2. PRE-CLEANSING — penanganan missing values
        # ------------------------------------------------------------
        cleaned_dfs, total_missing, detail_logs = pre_cleansing(dfs_info, numeric_cols)

        if total_missing > 0:
            st.success(
                f"Saya telah mendeteksi total {total_missing} sel kosong/tidak valid "
                f"pada kolom numerik dan telah mengisinya dengan nilai 0 agar "
                f"kalkulasi saya tetap aman."
            )
            with st.expander("Lihat detail sel kosong yang saya isi"):
                st.markdown("\n".join(detail_logs))
        else:
            st.success("Saya tidak menemukan sel kosong pada kolom numerik di seluruh file saya.")

        # Pengecekan konsistensi kolom identitas antar file (langkah pencegahan human error)
        pesan_identitas = check_identity_consistency(cleaned_dfs, identity_cols)
        if pesan_identitas:
            st.warning(pesan_identitas)

        # ------------------------------------------------------------
        # 3. KOMPUTASI — element-wise addition
        # ------------------------------------------------------------
        final_df = compute_accumulation(cleaned_dfs, identity_cols, numeric_cols, semua_kolom)
        st.success(
            f"Saya telah menjumlahkan {len(cleaned_dfs)} dataset secara element-wise "
            f"berdasarkan posisi baris dan kolomnya."
        )

        # ------------------------------------------------------------
        # 4. POST-CLEANSING — hapus duplikasi pada hasil akhir
        # ------------------------------------------------------------
        final_df_clean, duplicate_count = post_cleansing(final_df)
        if duplicate_count > 0:
            st.success(
                f"Saya telah mendeteksi dan menghapus {duplicate_count} baris "
                f"terduplikasi secara identik pada dataframe final saya."
            )
        else:
            st.success("Saya tidak menemukan baris terduplikasi pada dataframe final saya.")

        # Simpan ke session_state agar dashboard tetap tampil pada rerun berikutnya
        st.session_state["final_df"] = final_df_clean
        st.session_state["numeric_cols"] = numeric_cols

    # --------------------------------------------------------------------
    # 5. DASHBOARD OUTPUT & VISUALISASI
    # --------------------------------------------------------------------
    if "final_df" in st.session_state:
        st.divider()
        st.header("3️⃣ Dashboard Hasil Akumulasi Saya")
        render_dashboard(st.session_state["final_df"], st.session_state["numeric_cols"])


if __name__ == "__main__":
    main()
