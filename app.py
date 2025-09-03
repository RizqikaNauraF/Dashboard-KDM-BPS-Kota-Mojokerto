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
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stApp {
        background-color: #f5f7fa;
    }
</style>
""", unsafe_allow_html=True)

st.write(f"📅 Data terakhir diperbarui pada: Senin, 1 September 2025, pukul 05.00")
st.title("📊 Dashboard KDM BPS Kota Mojokerto - Sensus Ekonomi 2026")

# --Load Data---
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = df.columns.str.strip().str.lower()
    return df

# ===================== Sidebar ===================== #
st.sidebar.header("📂 Data")
default_path = "ProgressKDM.xlsx"

# Upload file opsional
uploaded_file = st.sidebar.file_uploader("Upload Excel (opsional)", type=["xlsx"])
source = uploaded_file if uploaded_file is not None else (default_path if os.path.exists(default_path) else None)

if source is None:
    st.warning(f"Letakkan file **{default_path}** di folder kerja, atau upload dari sidebar.")
    st.stop()

# Load data
def load_data(file_path):
    return pd.read_excel(file_path)

try:
    df = load_data(source)
except Exception as e:
    st.error(f"Gagal memuat data: {e}")
    st.stop()

# ---- Sidebar Filter ---- #
st.sidebar.markdown("---")
st.sidebar.subheader("📊 Pilih Data")
filter_option = st.sidebar.radio(
    "Pilih Data:",
    ("Semua", "Pegawai", "Non Pegawai")
)

# --- Load Data Sesuai Filter ---
try:
    if filter_option == "Semua":
        df = pd.read_excel(source, sheet_name="Semua")
    elif filter_option == "Pegawai":
        df = pd.read_excel(source, sheet_name="Pegawai")
    elif filter_option == "Non Pegawai":
        df = pd.read_excel(source, sheet_name="NonPegawai")
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
# st.subheader("📌 Statistik Ringkas")

# total_all = int(df["total"].sum())
# total_terbaru = int(df["terbaru"].sum())
# total_week = int(df["perolehan minggu ini"].sum())

# col1, col2, col3 = st.columns(3)
# with col1:
#     st.metric("📍 Total Tagging Sampai Minggu Lalu", f"{df['total'].sum():,}")
# with col2:
#     st.metric("🆕 Total Tagging Terbaru", f"{total_terbaru:,}")
# with col3:
#     st.metric("📅 Total Minggu Ini", f"{total_week:,}")

# col4, col5, col6 = st.columns(3)
# with col4:
#     st.metric("👥 Rata-rata per Pegawai", f"{df['terbaru'].mean():,.2f}")
# with col5:
#     st.metric("🏆 Max Tagging", f"{df['terbaru'].max():,}")
# with col6:
#     st.metric("📉 Min Tagging", f"{df['terbaru'].min():,}")


st.subheader("📌 Statistik Ringkas")

total_all = int(df["total"].sum())
total_terbaru = int(df["terbaru"].sum())
total_week = int(df["perolehan minggu ini"].sum())

# st.markdown("""
# <div style="display:flex; gap:20px; margin-bottom:20px;">
#     <div style="flex:1; background:#2ECC71; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>📍 Total Tagging Sampai Minggu Lalu</h4>
#         <p style="font-size:22px; font-weight:bold;">{:,}</p>
#     </div>
#     <div style="flex:1; background:#3498DB; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>🆕 Total Tagging Terbaru</h4>
#         <p style="font-size:22px; font-weight:bold;">{:,}</p>
#     </div>
#     <div style="flex:1; background:#E67E22; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>📅 Total Minggu Ini</h4>
#         <p style="font-size:22px; font-weight:bold;">{:,}</p>
#     </div>
# </div>
# """.format(df['total'].sum(), total_terbaru, total_week), unsafe_allow_html=True)

# st.markdown("""
# <div style="display:flex; gap:20px;">
#     <div style="flex:1; background:#9B59B6; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>👥 Rata-rata per Pegawai</h4>
#         <p style="font-size:22px; font-weight:bold;">{:.2f}</p>
#     </div>
#     <div style="flex:1; background:#1ABC9C; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>🏆 Max Tagging</h4>
#         <p style="font-size:22px; font-weight:bold;">{:,}</p>
#     </div>
#     <div style="flex:1; background:#E74C3C; padding:20px; border-radius:12px; color:white; text-align:center;">
#         <h4>📉 Min Tagging</h4>
#         <p style="font-size:22px; font-weight:bold;">{:,}</p>
#     </div>
# </div>
# """.format(df['terbaru'].mean(), df['terbaru'].max(), df['terbaru'].min()), unsafe_allow_html=True)

st.markdown(f"""
<div style="display:flex; flex-wrap:wrap; gap:20px; margin-bottom:20px;">
    <div style="flex:1 1 200px; background:#2ECC71; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>📍 Total Tagging Sampai Minggu Lalu</h4>
        <p style="font-size:22px; font-weight:bold;">{total_all:,}</p>
    </div>
    <div style="flex:1 1 200px; background:#3498DB; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>🆕 Total Tagging Terbaru</h4>
        <p style="font-size:22px; font-weight:bold;">{total_terbaru:,}</p>
    </div>
    <div style="flex:1 1 200px; background:#E67E22; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>📅 Total Minggu Ini</h4>
        <p style="font-size:22px; font-weight:bold;">{total_week:,}</p>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div style="display:flex; flex-wrap:wrap; gap:20px; margin-bottom:20px;">
    <div style="flex:1 1 200px; background:#9B59B6; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>👥 Rata-rata per Pegawai</h4>
        <p style="font-size:22px; font-weight:bold;">{df['terbaru'].mean():.2f}</p>
    </div>
    <div style="flex:1 1 200px; background:#1ABC9C; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>🏆 Max Tagging</h4>
        <p style="font-size:22px; font-weight:bold;">{df['terbaru'].max():,}</p>
    </div>
    <div style="flex:1 1 200px; background:#E74C3C; padding:20px; border-radius:12px; color:white; text-align:center;">
        <h4>📉 Min Tagging</h4>
        <p style="font-size:22px; font-weight:bold;">{df['terbaru'].min():,}</p>
    </div>
</div>
""", unsafe_allow_html=True)


st.markdown("---")

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

# # =========================
# # Leaderboard
# # =========================
# st.subheader("🏆 Leaderboard Pegawai")

# # Pilihan berapa top
# top_n = st.slider("Pilih jumlah Top-N:", 5, 30, 17)
# ascending = st.checkbox("⬆️ Urutkan dari terkecil", value=False)

# # Hitung agregasi
# leaderboard = df.groupby(["nama", "satker"], as_index=False).agg({
#     "total": "sum",
#     "terbaru": "sum",
#     "perolehan minggu ini": "sum"
# })

# # Urutkan sesuai pilihan user
# leaderboard = leaderboard.sort_values(by=sort_col, ascending=ascending).head(top_n)

# # Reset index dan kasih Rank
# leaderboard = leaderboard.reset_index(drop=True)
# leaderboard["Rank"] = leaderboard.index + 1

# # Ubah nama kolom agar lebih rapi
# leaderboard_display = leaderboard.rename(columns={
#     "Rank": "Rank",
#     "nama": "Nama",
#     "satker": "Satker",
#     "total": "Total Sampai Minggu Lalu",
#     "terbaru": "Total Terbaru",
#     "perolehan minggu ini": "Perolehan Minggu Ini"
# })

# # Tampilkan tabel dengan nama kolom baru
# st.dataframe(
#     leaderboard_display[[
#         "Rank",
#         "Nama",
#         "Satker",
#         "Total Sampai Minggu Lalu",
#         "Total Terbaru",
#         "Perolehan Minggu Ini"
#     ]],
#     use_container_width=True,
#     hide_index=True
# )

# =========================
# Leaderboard
# =========================
st.subheader("🏆 Leaderboard Pegawai")

# Pilihan berapa top
top_n = st.slider("Pilih jumlah Top-N yang tampil:", 5, 77, 17)
ascending = st.checkbox("⬆️ Urutkan dari terkecil", value=False)

# Hitung agregasi
leaderboard_full = df.groupby(["nama", "satker"], as_index=False).agg({
    "total": "sum",
    "terbaru": "sum",
    "perolehan minggu ini": "sum"
})

# Urutkan sesuai pilihan user
leaderboard = leaderboard_full.sort_values(by=sort_col, ascending=ascending).head(top_n)

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

# =========================
# Custom HTML Table
# =========================
html = """
<table style="width:100%; border-collapse:collapse; font-size:14px;">
  <tr style="background-color:#2E86C1; color:white; text-align:left;">
    <th style="padding:8px;">Rank</th>
    <th style="padding:8px;">Nama</th>
    <th style="padding:8px;">Satker</th>
    <th style="padding:8px;">Total Sampai Minggu Lalu</th>
    <th style="padding:8px;">Total Terbaru</th>
    <th style="padding:8px;">Perolehan Minggu Ini</th>
  </tr>
"""

for _, row in leaderboard_display.iterrows():
    warna = "green" if row["Perolehan Minggu Ini"] >= 0 else "red"
    html += f"""
  <tr style="background-color:#f9f9f9;">
    <td style="padding:8px;">{row['Rank']}</td>
    <td style="padding:8px;">{row['Nama']}</td>
    <td style="padding:8px;">{row['Satker']}</td>
    <td style="padding:8px;">{row['Total Sampai Minggu Lalu']:,}</td>
    <td style="padding:8px;">{row['Total Terbaru']:,}</td>
    <td style="padding:8px; color:{warna};">{row['Perolehan Minggu Ini']:+,}</td>
  </tr>
"""

html += "</table>"

# Render ke Streamlit
st.markdown(html, unsafe_allow_html=True)

# =========================
# Export Full Leaderboard
# =========================
st.subheader("⬇️ Export Leaderboard Full")

col1, col2 = st.columns(2)


# Data lengkap untuk export (full, tanpa top_n)
full_leaderboard = leaderboard_full.copy()  # full data, jangan sentuh top_n

# Terapkan urutan user
full_leaderboard_sorted = full_leaderboard.sort_values(
    by=sort_col,         # dari pilihan user
    ascending=ascending  # dari pilihan user
).reset_index(drop=True)

# Kasih Rank sesuai urutan user
full_leaderboard_sorted["Rank"] = full_leaderboard_sorted.index + 1

# Rename kolom untuk export
full_leaderboard_display = full_leaderboard_sorted.rename(columns={
    "Rank": "Rank",
    "nama": "Nama",
    "satker": "Satker",
    "total": "Total Sampai Minggu Lalu",
    "terbaru": "Total Terbaru",
    "perolehan minggu ini": "Perolehan Minggu Ini"
})

# Export Excel
with col1:
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine="xlsxwriter") as writer:
        full_leaderboard_display.to_excel(writer, index=False, sheet_name="Leaderboard")
    st.download_button(
        "📥 Download Leaderboard Excel",
        data=out_excel,
        file_name=f"Leaderboard_Full_{selected_date}.xlsx"
    )

# Export PDF
with col2:
    try:
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, LongTable
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm

        def create_pdf(df):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4,
                                    leftMargin=2*cm, rightMargin=2*cm,
                                    topMargin=2*cm, bottomMargin=2*cm)
            styles = getSampleStyleSheet()

            # Header rapi
            header = [Paragraph(str(col), styles["Heading4"]) for col in df.columns]

            # Isi tabel rapi
            table_data = [header]
            for row in df.values.tolist():
                table_data.append([Paragraph(str(cell), styles["Normal"]) for cell in row])

            # Hitung lebar kolom agar muat halaman
            page_width, page_height = A4
            usable_width = page_width - doc.leftMargin - doc.rightMargin
            col_width = usable_width / len(df.columns)

            # Buat tabel panjang agar lanjut ke halaman berikutnya
            table = LongTable(table_data, colWidths=[col_width]*len(df.columns))
            table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.grey),
                ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ]))

            story = [Paragraph(f"Leaderboard KDM - {selected_date.strftime('%d %B %Y')}", styles["Title"])]
            story.append(table)

            doc.build(story)
            buffer.seek(0)
            return buffer

        pdf_bytes = create_pdf(full_leaderboard_display)
        st.download_button(
            "📄 Download Leaderboard PDF",
            data=pdf_bytes,
            file_name=f"Leaderboard_Full_{selected_date}.pdf",
            mime="application/pdf"
        )
    except ImportError:
        st.info("Export PDF membutuhkan paket 'reportlab'. Install: pip install reportlab")

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
# # 📊 Perbandingan dengan File tanggal 15 Agustus 2025
# # =========================
# st.subheader("📊 Perbandingan dengan Data Tanggal 15 Agustus 2025")

# try:
#     # Baca file Excel
#     df_new = pd.read_excel("KDM_15-8.xlsx", engine="openpyxl")
#     df_new.columns = df_new.columns.str.strip().str.lower()

#     # Pastikan kolom yang dibutuhkan ada
#     if "nama" in df_new.columns and "total" in df_new.columns:

#         # Pastikan numeric
#         df_new["total"] = pd.to_numeric(df_new["total"], errors="coerce")
#         df["terbaru"] = pd.to_numeric(df["terbaru"], errors="coerce")

#         # Ambil hanya kolom yang diperlukan lalu merge
#         df_compare = df[["nama", "satker", "terbaru"]].merge(
#             df_new[["nama", "total"]],
#             on="nama",
#             how="left"
#         )

#         # Hitung selisih
#         df_compare["selisih"] = df_compare["terbaru"] - df_compare["total"]

#         # Rename kolom biar rapi
#         df_show = df_compare.rename(
#             columns={
#                 "nama": "Nama",
#                 "satker": "Satker",
#                 "terbaru": "Total Terbaru",
#                 "total": "Total Tanggal 15",
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

#         # Urutkan & ranking
#         df_show = df_show.sort_values(by="Selisih", ascending=ascending)
#         df_show.insert(0, "Ranking", range(1, len(df_show) + 1))
#         df_show = df_show.reset_index(drop=True)

#         # Tampilkan tabel
#         st.dataframe(df_show, use_container_width=True, hide_index=True)

#     else:
#         st.warning("⚠️ File KDM_15-8.xlsx tidak memiliki kolom 'nama' dan 'total'.")
# except FileNotFoundError:
#     st.info("📂 File 'KDM_15-8.xlsx' belum ditemukan di folder. Upload dulu untuk perbandingan.")

st.markdown("---")

# =========================
# 📊 Perbandingan dengan File tanggal 15 Agustus 2025
# =========================
st.subheader("📊 Perbandingan dengan Data Tanggal 15 Agustus 2025")

try:
    # Baca file Excel
    df_new = pd.read_excel("KDM_15-8.xlsx", engine="openpyxl")
    df_new.columns = df_new.columns.str.strip().str.lower()

    if "nama" in df_new.columns and "total" in df_new.columns:
        df_new["total"] = pd.to_numeric(df_new["total"], errors="coerce")
        df["terbaru"] = pd.to_numeric(df["terbaru"], errors="coerce")

        # Merge
        df_compare = df[["nama", "satker", "terbaru"]].merge(
            df_new[["nama", "total"]],
            on="nama",
            how="left"
        )

        # Hitung selisih
        df_compare["selisih"] = df_compare["terbaru"] - df_compare["total"]

        # Rename kolom
        df_show = df_compare.rename(
            columns={
                "nama": "Nama",
                "satker": "Satker",
                "terbaru": "Total Terbaru",
                "total": "Total Tanggal 15",
                "selisih": "Selisih"
            }
        )[["Nama", "Satker", "Total Terbaru", "Total Tanggal 15", "Selisih"]]

        # Pilihan urutan
        order = st.radio(
            "Urutkan berdasarkan Selisih:",
            ["Terbesar ke Terkecil", "Terkecil ke Terbesar"],
            horizontal=True
        )
        ascending = True if order == "Terkecil ke Terbesar" else False

        # Ranking & urutkan
        df_show = df_show.sort_values(by="Selisih", ascending=ascending)
        df_show.insert(0, "Ranking", range(1, len(df_show) + 1))
        df_show = df_show.reset_index(drop=True)

        # ---- Top-N slider ----
        top_n = st.slider("Pilih jumlah Top-N yang tampil:", 5, 77, 17, key="top_n_perbandingan")
        df_top = df_show.head(top_n)

        # === Tabel HTML custom Top-N ===
        table_html = """
<table style="width:100%; border-collapse:collapse; font-size:14px;">
  <tr style="background-color:#2E86C1; color:white;">
    <th style="padding:8px;">Ranking</th>
    <th style="padding:8px;">Nama</th>
    <th style="padding:8px;">Satker</th>
    <th style="padding:8px;">Total Terbaru</th>
    <th style="padding:8px;">Total Tanggal 15</th>
    <th style="padding:8px;">Selisih</th>
  </tr>
"""
        for _, row in df_top.iterrows():
            val_terbaru = f"{int(row['Total Terbaru']):,}" if pd.notna(row["Total Terbaru"]) else "0"
            val_tgl15 = f"{int(row['Total Tanggal 15']):,}" if pd.notna(row["Total Tanggal 15"]) else "0"
            if pd.notna(row["Selisih"]):
                val_selisih = f"{int(row['Selisih']):+,}"
                selisih_color = "green" if row["Selisih"] >= 0 else "red"
            else:
                val_selisih = "0"
                selisih_color = "black"

            table_html += f"""
  <tr style="background-color:#f9f9f9;">
    <td style="padding:8px; text-align:center;">{row['Ranking']}</td>
    <td style="padding:8px;">{row['Nama']}</td>
    <td style="padding:8px;">{row['Satker']}</td>
    <td style="padding:8px; text-align:right;">{val_terbaru}</td>
    <td style="padding:8px; text-align:right;">{val_tgl15}</td>
    <td style="padding:8px; color:{selisih_color}; text-align:right;">{val_selisih}</td>
  </tr>
"""
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)

        # ---- Export Data Full (sebelah-sebelahan) ----
        st.subheader("⬇️ Export Perbandingan Full")

        # Ganti NaN dengan 0 dan hapus .0
        df_export = df_show.fillna(0).copy()
        for col in ["Total Terbaru", "Total Tanggal 15", "Selisih"]:
            df_export[col] = df_export[col].astype(int)

        col1, col2 = st.columns(2)

        # Export Excel
        with col1:
            out_excel = io.BytesIO()
            with pd.ExcelWriter(out_excel, engine="xlsxwriter") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Perbandingan")
            st.download_button(
                "📥 Download Perbandingan Excel",
                data=out_excel,
                file_name=f"Perbandingan_Full_{selected_date}.xlsx"
            )

        # Export PDF
        with col2:
            try:
                from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, LongTable
                from reportlab.lib import colors
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.styles import getSampleStyleSheet
                from reportlab.lib.units import cm

                def create_pdf(df):
                    buffer = io.BytesIO()
                    doc = SimpleDocTemplate(buffer, pagesize=A4,
                                            leftMargin=2*cm, rightMargin=2*cm,
                                            topMargin=2*cm, bottomMargin=2*cm)
                    styles = getSampleStyleSheet()

                    # Header
                    header = [Paragraph(str(col), styles["Heading4"]) for col in df.columns]

                    # Isi tabel
                    table_data = [header]
                    for row in df.values.tolist():
                        table_data.append([Paragraph(str(cell), styles["Normal"]) for cell in row])

                    # Lebar kolom
                    page_width, page_height = A4
                    usable_width = page_width - doc.leftMargin - doc.rightMargin
                    col_width = usable_width / len(df.columns)

                    table = LongTable(table_data, colWidths=[col_width]*len(df.columns))
                    table.setStyle(TableStyle([
                        ("BACKGROUND", (0,0), (-1,0), colors.grey),
                        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
                        ("ALIGN", (0,0), (-1,-1), "CENTER"),
                        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
                    ]))

                    story = [Paragraph(f"Perbandingan data 15 Agustus 2025 dengan - {selected_date.strftime('%d %B %Y')}", styles["Title"])]
                    story.append(table)

                    doc.build(story)
                    buffer.seek(0)
                    return buffer

                pdf_bytes = create_pdf(df_export)
                st.download_button(
                    "📄 Download Perbandingan PDF",
                    data=pdf_bytes,
                    file_name=f"Perbandingan_Full_{selected_date}.pdf",
                    mime="application/pdf"
                )
            except ImportError:
                st.info("Export PDF membutuhkan paket 'reportlab'. Install: pip install reportlab")

    else:
        st.warning("⚠️ File KDM_15-8.xlsx tidak memiliki kolom 'nama' dan 'total'.")
except FileNotFoundError:
    st.info("📂 File 'KDM_15-8.xlsx' belum ditemukan di folder. Upload dulu untuk perbandingan.")

st.markdown("""
                <hr style="border: 0.5px solid #ccc;" />
                <center><small>&copy; 2025 BPS Kota Mojokerto</small></center>
                """, unsafe_allow_html=True)