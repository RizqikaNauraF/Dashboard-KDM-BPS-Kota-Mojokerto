import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


st.set_page_config(
    page_title="Dashboard KDM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.write(f"📅 Data terakhir diperbarui pada: Senin, 1 September 2025, pukul 05.00")
st.title("📊 Dashboard KDM BPS Kota Mojokerto - Sensus Ekonomi 2026")

# --Load Data---
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = df.columns.str.strip().str.lower()
    return df

# --- Sidebar Upload ---
uploaded_file = st.sidebar.file_uploader("Upload file Excel KDM", type=["xlsx"])

if uploaded_file is not None:
    df = load_data(uploaded_file)
else:
    try:
        df = load_data("ProgressKDM.xlsx")
    except Exception:
        df = pd.DataFrame([])

# --- Jika Data Kosong ---
if df.empty:
    st.warning("⚠️ Tidak ada data tersedia. Upload file Excel terlebih dahulu.")
    st.stop()

# --- Sidebar Filter ---
filter_option = st.sidebar.radio(
    "Pilih Data:",
    ("Semua", "Pegawai", "Non Pegawai")
)

# --- Load Data Sesuai Filter ---
try:
    if filter_option == "Semua":
        df = pd.read_excel("ProgressKDM.xlsx", sheet_name="Semua")
    elif filter_option == "Pegawai":
        df = pd.read_excel("ProgressKDM.xlsx", sheet_name="Pegawai")
    elif filter_option == "Non Pegawai":
        df = pd.read_excel("ProgressKDM.xlsx", sheet_name="NonPegawai")
except Exception:
    df = pd.DataFrame([])

# Kolom lowercase
df.columns = [c.strip().lower() for c in df.columns]

# Pastikan kolom ada
required_cols = ["nama", "total", "terbaru", "perolehan minggu ini"]
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"❌ Kolom berikut tidak ada di file: {missing_cols}")
    st.stop()

# =========================
# Konversi & Pilih Tanggal (pakai selectbox, tanpa waktu)
# =========================
if "tanggal" in df.columns:
    df["tanggal"] = pd.to_datetime(df["tanggal"], dayfirst=True, errors="coerce")

    if df["tanggal"].notna().any():
        # Ambil tanggal unik (tanpa jam) dan urutkan
        available_dates = sorted(df["tanggal"].dt.date.unique())

        # DEFAULT TANGGAL
        default_date = datetime.strptime("01/09/2025", "%d/%m/%Y").date()
        if default_date not in available_dates:
            default_date = available_dates[0]

        # Selectbox pilih tanggal (tampilkan format dd-mm-YYYY)
        selected_date_str = st.selectbox(
            "📅 Pilih tanggal:",
            options=[d.strftime("%d-%m-%Y") for d in available_dates],
            index=available_dates.index(default_date)
        )

        # Konversi kembali ke tipe date
        selected_date = datetime.strptime(selected_date_str, "%d-%m-%Y").date()

        # Filter dataframe sesuai tanggal dipilih
        df = df[df["tanggal"].dt.date == selected_date]

# =========================
# Statistik Ringkas
# =========================
st.subheader("📌 Statistik Ringkas")

total_all = int(df["total"].sum())
total_terbaru = int(df["terbaru"].sum())
total_week = int(df["perolehan minggu ini"].sum())

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📍 Total Tagging Sampai Minggu Lalu", f"{df['total'].sum():,}")
with col2:
    st.metric("🆕 Total Tagging Terbaru", f"{total_terbaru:,}")
with col3:
    st.metric("📅 Total Minggu Ini", f"{total_week:,}")

col4, col5, col6 = st.columns(3)
with col4:
    st.metric("👥 Rata-rata per Pegawai", f"{df['terbaru'].mean():,.2f}")
with col5:
    st.metric("🏆 Max Tagging", f"{df['terbaru'].max():,}")
with col6:
    st.metric("📉 Min Tagging", f"{df['terbaru'].min():,}")


# =========================
# Search Nama
# =========================
search_name = st.text_input("🔍 Cari berdasarkan nama:").lower()
if search_name:
    df = df[df["nama"].str.lower().str.contains(search_name)]

# =========================
# Pilih Mode Ranking
# =========================
ranking_mode = st.radio("📈 Pilih mode ranking:", 
                        ["Total Sampai Dengan Minggu Lalu", "Total Terbaru", "Perolehan Minggu Ini"])

if ranking_mode == "Total Sampai Dengan Minggu Lalu":
    sort_col = "total"
elif ranking_mode == "Total Terbaru":
    sort_col = "terbaru"
else:
    sort_col = "perolehan minggu ini"

# =========================
# Leaderboard
# =========================
st.subheader("🏆 Leaderboard Pegawai")

# Pilihan berapa top
top_n = st.slider("Pilih jumlah Top-N:", 5, 30, 17)
ascending = st.checkbox("⬆️ Urutkan dari terkecil", value=False)

# Hitung agregasi
leaderboard = df.groupby(["nama", "satker"], as_index=False).agg({
    "total": "sum",
    "terbaru": "sum",
    "perolehan minggu ini": "sum"
})

# Urutkan sesuai pilihan user
leaderboard = leaderboard.sort_values(by=sort_col, ascending=ascending).head(top_n)

# Reset index dan kasih Rank
leaderboard = leaderboard.reset_index(drop=True)
leaderboard["Rank"] = leaderboard.index + 1

# Ubah nama kolom agar lebih rapi
leaderboard_display = leaderboard.rename(columns={
    "Rank": "Rank",
    "nama": "Nama",
    "satker": "Satker",
    "total": "Total Sampai Minggu Lalu",
    "terbaru": "Total Terbaru",
    "perolehan minggu ini": "Perolehan Minggu Ini"
})

# Tampilkan tabel dengan nama kolom baru
st.dataframe(
    leaderboard_display[[
        "Rank",
        "Nama",
        "Satker",
        "Total Sampai Minggu Lalu",
        "Total Terbaru",
        "Perolehan Minggu Ini"
    ]],
    use_container_width=True,
    hide_index=True
)


# =========================
# Grafik Ranking (Gradasi Hijau)
# =========================

# Urutkan sesuai pilihan ranking
if ascending:
    df_plot = leaderboard.sort_values(by=sort_col, ascending=True)
else:
    df_plot = leaderboard.sort_values(by=sort_col, ascending=False)

df_plot = df_plot[::-1]  # Biar ranking atas muncul di atas

fig = px.bar(
    df_plot,
    x=sort_col,
    y="nama",
    title=f"📊 Top {top_n} berdasarkan {ranking_mode}",
    text=sort_col,
    orientation="h",
    color=sort_col,  # kasih warna sesuai nilai ranking
    color_continuous_scale="Greens",  # gradasi hijau
)

fig.update_traces(
    textposition="outside",
    marker=dict(line=dict(width=0))  # biar clean
)

fig.update_layout(
    yaxis=dict(
        categoryorder="array",
        categoryarray=df_plot["nama"].tolist()
    ),
    coloraxis_showscale=False  # sembunyikan legend skala warna kalau mau
)

st.plotly_chart(fig, use_container_width=True)


# # =========================
# # Tren Mingguan
# # =========================
# if "tanggal" in df.columns:
#     st.subheader("📅 Tren Mingguan")
#     trend_metric = st.selectbox("Pilih metrik tren:", ["total", "terbaru", "perolehan minggu ini"])
#     trend = df.groupby("tanggal", as_index=False)[trend_metric].sum()
#     fig_trend = px.line(trend, x="tanggal", y=trend_metric, 
#                         title=f"Tren {trend_metric} dari waktu ke waktu",
#                         markers=True)
#     st.plotly_chart(fig_trend, use_container_width=True)


# # =========================
# # 📊 Perbandingan dengan File Baru
# # =========================
# st.subheader("📊 Perbandingan dengan Data Tanggal 15 Agustus 2025")

# try:
#     df_new = pd.read_excel("KDM_15-8.xlsx", engine="openpyxl")
#     df_new.columns = df_new.columns.str.strip().str.lower()

#     if "nama" in df_new.columns and "total" in df_new.columns:
#         df_compare = df.merge(
#             df_new[["nama", "total"]],
#             on="nama",
#             how="left",
#             suffixes=("", "_15")
#         )

#         # Hitung selisih (Total Terbaru - Total Tanggal 15)
#         df_compare["selisih"] = df_compare["terbaru"] - df_compare["total"]
        
#         # Rename kolom biar rapi
#         df_show = df_compare.rename(
#             columns={
#                 "total": "Total Tanggal 15",
#                 "terbaru": "Total Terbaru",
#                 "nama": "Nama",
#                 "satker": "Satker",
#                 "selisih": "Selisih"
#             }
#         )[
#             ["Nama", "Satker", "Total Terbaru", "Total Tanggal 15", "Selisih"]
#         ]

#         # Pilihan urutan
#         order = st.radio(
#             "Urutkan berdasarkan Selisih:",
#             ["Terbesar ke Terkecil", "Terkecil ke Terbesar"],
#             horizontal=True
#         )

#         ascending = True if order == "Terkecil ke Terbesar" else False

#         # Urutkan berdasarkan Selisih
#         df_show = df_show.sort_values(by="Selisih", ascending=ascending)

#         # Tambahkan Ranking berdasarkan urutan selisih
#         df_show.insert(0, "Ranking", range(1, len(df_show) + 1))

#         # Reset index supaya angka di samping hilang
#         df_show = df_show.reset_index(drop=True)

#         st.dataframe(df_show, use_container_width=True, hide_index=True)

#     else:
#         st.warning("⚠️ File KDM_15-8.xlsx tidak memiliki kolom 'nama' dan 'total'.")
# except FileNotFoundError:
#     st.info("📂 File 'KDM_15-8.xlsx' belum ditemukan di folder. Upload dulu untuk perbandingan.")

# =========================
# 📊 Perbandingan dengan File Baru
# =========================
st.subheader("📊 Perbandingan dengan Data Tanggal 15 Agustus 2025")

try:
    # Baca file Excel
    df_new = pd.read_excel("KDM_15-8.xlsx", engine="openpyxl")
    df_new.columns = df_new.columns.str.strip().str.lower()

    # Pastikan kolom yang dibutuhkan ada
    if "nama" in df_new.columns and "total" in df_new.columns:

        # Pastikan numeric
        df_new["total"] = pd.to_numeric(df_new["total"], errors="coerce")
        df["terbaru"] = pd.to_numeric(df["terbaru"], errors="coerce")

        # Ambil hanya kolom yang diperlukan lalu merge
        df_compare = df[["nama", "satker", "terbaru"]].merge(
            df_new[["nama", "total"]],
            on="nama",
            how="left"
        )

        # Hitung selisih
        df_compare["selisih"] = df_compare["terbaru"] - df_compare["total"]

        # Rename kolom biar rapi
        df_show = df_compare.rename(
            columns={
                "nama": "Nama",
                "satker": "Satker",
                "terbaru": "Total Terbaru",
                "total": "Total Tanggal 15",
                "selisih": "Selisih"
            }
        )[
            ["Nama", "Satker", "Total Terbaru", "Total Tanggal 15", "Selisih"]
        ]

        # Pilihan urutan
        order = st.radio(
            "Urutkan berdasarkan Selisih:",
            ["Terbesar ke Terkecil", "Terkecil ke Terbesar"],
            horizontal=True
        )
        ascending = True if order == "Terkecil ke Terbesar" else False

        # Urutkan & ranking
        df_show = df_show.sort_values(by="Selisih", ascending=ascending)
        df_show.insert(0, "Ranking", range(1, len(df_show) + 1))
        df_show = df_show.reset_index(drop=True)

        # Tampilkan tabel
        st.dataframe(df_show, use_container_width=True, hide_index=True)

    else:
        st.warning("⚠️ File KDM_15-8.xlsx tidak memiliki kolom 'nama' dan 'total'.")
except FileNotFoundError:
    st.info("📂 File 'KDM_15-8.xlsx' belum ditemukan di folder. Upload dulu untuk perbandingan.")



# =========================
# Export Data
# =========================
st.subheader("⬇️ Export Data")

# Export Excel
out_excel = io.BytesIO()
with pd.ExcelWriter(out_excel, engine="xlsxwriter") as writer:
    leaderboard.to_excel(writer, index=False, sheet_name="Leaderboard")
st.download_button("📥 Download Leaderboard Excel", 
                   data=out_excel, 
                   file_name="leaderboard.xlsx")

# Export PDF
def create_pdf(dataframe):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = [Paragraph("Leaderboard KDM", styles["Title"])]

    table_data = [list(dataframe.columns)] + dataframe.values.tolist()
    table = Table(table_data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
    ]))
    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

pdf_bytes = create_pdf(leaderboard)
st.download_button("📄 Download Leaderboard PDF", 
                   data=pdf_bytes, 
                   file_name="leaderboard.pdf", 
                   mime="application/pdf")

st.markdown("""
                <hr style="border: 0.5px solid #ccc;" />
                <center><small>&copy; 2025 BPS Kota Mojokerto</small></center>
                """, unsafe_allow_html=True)