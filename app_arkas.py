import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

import db_supabase as db

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(page_title="Pendamping ARKAS", layout="wide", page_icon="🧾")


def _gerbang_pin():
    """Gerbang PIN sederhana untuk aplikasi yang diakses lewat internet publik.
    Aktif hanya kalau APP_PASSWORD diisi di secrets — kalau tidak diisi,
    aplikasi tetap bisa dibuka langsung tanpa PIN (mis. untuk coba-coba lokal).
    Ini BUKAN login multi-pengguna, cuma satu PIN bersama untuk 1 sekolah."""
    pin_target = st.secrets.get("APP_PASSWORD", "")
    if not pin_target:
        return
    if st.session_state.get("_auth_ok"):
        return
    st.title("🔒 Pendamping ARKAS")
    st.caption("Aplikasi ini dilindungi PIN karena datanya tersimpan online.")
    pin = st.text_input("Masukkan PIN akses", type="password")
    if st.button("Masuk", use_container_width=True):
        if pin == pin_target:
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("PIN salah.")
    st.stop()


_gerbang_pin()

COLOR_PRIMARY = colors.HexColor("#0F6E4F")   # hijau keuangan
COLOR_ACCENT = colors.HexColor("#D9A400")    # emas
COLOR_DANGER = colors.HexColor("#C0392B")
COLOR_LIGHT = colors.HexColor("#EAF5EF")
DOCX_PRIMARY_RGB = RGBColor(0x0F, 0x6E, 0x4F)
DOCX_GREY_RGB = RGBColor(0x55, 0x55, 0x55)

# ============================================================
# TEMA VISUAL
# ============================================================
st.markdown("""
<style>
    .stApp { background: linear-gradient(180deg, #F5FAF7 0%, #FFFFFF 100%); }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F6E4F 0%, #0A4E38 100%);
    }
    section[data-testid="stSidebar"] * { color: #F2FAF6 !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] select { color: #16324F !important; }
    h1, h2, h3 { color: #0F6E4F; }
    div.stButton > button {
        background: linear-gradient(135deg, #0F6E4F 0%, #16965F 100%);
        color: white; border: none; border-radius: 8px; font-weight: 600;
        padding: 0.5rem 1.1rem;
    }
    div.stButton > button:hover { box-shadow: 0 4px 10px rgba(15,110,79,0.35); }
    .stTabs [data-baseweb="tab"] {
        background-color: #EAF5EF; border-radius: 8px 8px 0 0; padding: 8px 16px; font-weight: 600;
    }
    .stTabs [aria-selected="true"] { background-color: #0F6E4F !important; color: white !important; }
    div[data-testid="stMetric"] {
        background: #FFFFFF; border: 1px solid #DCEDE3; border-radius: 12px;
        padding: 10px 16px; box-shadow: 0 2px 8px rgba(15,110,79,0.06);
    }
</style>
""", unsafe_allow_html=True)

st.title("🧾 Pendamping ARKAS")
st.caption("Aplikasi bantu TIDAK RESMI untuk mempersiapkan & menganalisis data RKAS/BKU — "
           "bukan pengganti dan tidak terhubung ke Aplikasi ARKAS resmi Kemendikdasmen. "
           "Selalu verifikasi angka final di ARKAS & konsultasikan ke Dinas Pendidikan setempat.")

# ============================================================
# ATURAN UMUM BOSP 2026 (Permendikdasmen No. 8 Tahun 2026) - ringkasan acuan kepatuhan
# Catatan: ini ringkasan umum untuk BOS/BOP Reguler. Aturan bisa direvisi - selalu cek
# Juknis resmi terbaru & konfirmasi ke Dinas Pendidikan setempat sebelum menetapkan RKAS final.
# ============================================================
ATURAN_KOMPONEN = {
    "Pengembangan Perpustakaan / Buku": {"tipe": "min", "persen": 10,
        "keterangan": "Wajib paling sedikit 10% dari total pagu alokasi untuk penyediaan buku."},
    "Pembayaran Honor": {"tipe": "max_by_status", "persen_negeri": 20, "persen_swasta": 40,
        "keterangan": "Maksimal 20% (negeri) / 40% (swasta) dari 100% total pagu alokasi tahunan. "
                       "Hanya untuk guru/tendik non-ASN yang belum menerima tunjangan profesi."},
    "Pemeliharaan Sarana & Prasarana": {"tipe": "max", "persen": 20,
        "keterangan": "Maksimal 20% dari total pagu alokasi (pemeliharaan lahan/bangunan/ruang, "
                       "akses disabilitas, tanggap darurat bencana)."},
}
KOMPONEN_LAIN = ["Kegiatan Pembelajaran & Ekstrakurikuler", "Kegiatan Asesmen & Evaluasi",
                  "Pengembangan Mutu Guru & Tendik", "Langganan Daya & Jasa",
                  "Administrasi & Pelaporan", "Lainnya"]
SEMUA_KOMPONEN = list(ATURAN_KOMPONEN.keys()) + KOMPONEN_LAIN
SUMBER_DANA_BERATURAN = ["BOS Reguler", "BOP PAUD Reguler", "BOP Kesetaraan Reguler"]
SUMBER_DANA_OPSI = SUMBER_DANA_BERATURAN + ["BOS Kinerja", "BOP PAUD Kinerja", "Lainnya"]

# ============================================================
# STATE AWAL
# ============================================================
defaults = {
    "sekolah": "", "npsn": "", "status_sekolah": "Negeri", "kepsek_nama": "",
    "bendahara_nama": "", "tahun_anggaran": "2026", "wilayah_3t": False,
    "rincian_rkas": pd.DataFrame([
        {"Komponen": "Pengembangan Perpustakaan / Buku", "Uraian Kegiatan": "", "Jumlah (Rp)": 0, "Kena PPN 12%": False},
    ]),
    "checklist_tahapan": None,
    # --- Modul SPJ/LPJ BOS ---
    "spj_desa": "", "spj_kecamatan": "", "spj_kabupaten": "", "spj_provinsi": "",
    "spj_kepsek_nip": "", "spj_bendahara_nip": "",
    "spj_saldo_awal_kas": 0, "spj_saldo_awal_bank": 0,
    "spj_pagu_komponen": {k: 0 for k in SEMUA_KOMPONEN},
    "spj_bku": pd.DataFrame([
        {"Tanggal": date.today(), "Kode Rekening": "", "No. Bukti": "", "Uraian": "",
         "Komponen": SEMUA_KOMPONEN[0], "Metode": "Bank", "Penerimaan": 0, "Pengeluaran": 0,
         "PPN": 0, "PPh 21": 0, "PPh 22": 0, "PPh 23": 0, "Pajak Sudah Disetor": False},
    ]),
}
# Muat data tersimpan dari database (kalau sudah dikonfigurasi) - hanya sekali
# per sesi, dan HANYA mengisi key yang datanya memang ada, supaya key lain
# tetap kebagian nilai default di bawah.
if "_db_loaded" not in st.session_state:
    st.session_state["_db_loaded"] = True
    if db.is_configured():
        _dimuat = db.load_state()
        if _dimuat:
            for k, v in _dimuat.items():
                if v is not None:
                    st.session_state[k] = v

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# SIDEBAR - IDENTITAS SEKOLAH
# ============================================================
with st.sidebar:
    st.header("🏫 Identitas Sekolah")
    st.session_state.sekolah = st.text_input("Nama Sekolah", st.session_state.sekolah)
    st.session_state.npsn = st.text_input("NPSN (opsional)", st.session_state.npsn)
    st.session_state.status_sekolah = st.selectbox("Status Sekolah", ["Negeri", "Swasta"],
        index=["Negeri", "Swasta"].index(st.session_state.status_sekolah))
    st.session_state.tahun_anggaran = st.text_input("Tahun Anggaran", st.session_state.tahun_anggaran)
    st.divider()
    st.subheader("👤 Identitas Penanggung Jawab")
    st.session_state.kepsek_nama = st.text_input("Nama Kepala Sekolah", st.session_state.kepsek_nama)
    st.session_state.bendahara_nama = st.text_input("Nama Bendahara Sekolah", st.session_state.bendahara_nama)
    st.divider()
    st.session_state.wilayah_3t = st.checkbox(
        "Sekolah di wilayah 3T", value=st.session_state.wilayah_3t,
        help="Beberapa daerah 3T mendapat relaksasi batas komponen honor. Besaran relaksasi "
             "berbeda-beda tiap daerah — selalu konfirmasi ke Dinas Pendidikan setempat, "
             "aplikasi ini TIDAK mengubah perhitungan otomatis untuk kasus ini.",
    )
    st.divider()
    st.subheader("💾 Penyimpanan Data")
    if db.is_configured():
        st.success("Terhubung ke database — data bisa disimpan permanen.", icon="✅")
        colsimpan, colmuat = st.columns(2)
        with colsimpan:
            if st.button("💾 Simpan", use_container_width=True,
                         help="Simpan semua data (identitas, RKAS, BKU/SPJ, checklist) ke database."):
                if db.save_state():
                    st.toast("Data tersimpan ke database.", icon="✅")
                else:
                    st.toast("Gagal menyimpan data.", icon="⚠️")
        with colmuat:
            if st.button("🔄 Muat Ulang", use_container_width=True,
                         help="Batalkan perubahan di layar & ambil ulang versi tersimpan di database."):
                _dimuat = db.load_state()
                if _dimuat:
                    for k, v in _dimuat.items():
                        if v is not None:
                            st.session_state[k] = v
                    st.toast("Data dimuat ulang dari database.", icon="🔄")
                else:
                    st.toast("Belum ada data tersimpan di database.", icon="ℹ️")
                st.rerun()
        if st.session_state.get("_db_error"):
            st.warning(st.session_state["_db_error"], icon="⚠️")
        st.caption("Klik **Simpan** setelah selesai mengisi/mengubah data supaya tidak hilang "
                   "kalau aplikasi di-restart. Aplikasi tidak menyimpan otomatis di setiap ketikan.")
    else:
        st.info("Database belum terhubung — data cuma tersimpan selama sesi browser terbuka "
                "(hilang kalau aplikasi di-restart). Lihat README_DEPLOY.md untuk mengaktifkan "
                "penyimpanan permanen gratis dengan Supabase.", icon="ℹ️")
        if st.session_state.get("_db_connect_error"):
            st.warning(st.session_state["_db_connect_error"], icon="⚠️")
    st.divider()
    st.caption("📘 Ringkasan aturan mengacu Permendikdasmen No. 8 Tahun 2026 (Juknis BOSP 2026). "
               "Aturan dapat direvisi sewaktu-waktu — selalu cek Juknis resmi terbaru.")


def format_rupiah(angka) -> str:
    try:
        return f"Rp {angka:,.0f}".replace(",", ".")
    except Exception:
        return f"Rp {angka}"


def build_pdf_styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("Title", parent=base["Heading1"], fontSize=16, alignment=TA_CENTER,
                                 textColor=COLOR_PRIMARY, spaceAfter=4),
        "subtitle": ParagraphStyle("Subtitle", parent=base["Normal"], fontSize=10, alignment=TA_CENTER,
                                    textColor=colors.HexColor("#555555"), spaceAfter=10),
        "h2": ParagraphStyle("H2", parent=base["Heading2"], fontSize=12, textColor=COLOR_PRIMARY,
                              spaceBefore=10, spaceAfter=6),
        "meta_label": ParagraphStyle("MetaLabel", parent=base["Normal"], fontSize=9.5,
                                      textColor=colors.HexColor("#555555"), alignment=TA_RIGHT),
        "meta_value": ParagraphStyle("MetaValue", parent=base["Normal"], fontSize=10,
                                      fontName="Helvetica-Bold"),
        "cell": ParagraphStyle("Cell", parent=base["Normal"], fontSize=8.8, leading=11),
        "note": ParagraphStyle("Note", parent=base["Normal"], fontSize=8, leading=11,
                                textColor=colors.HexColor("#777777")),
    }


def _meta_pairs(meta: dict):
    rows, pair = [], []
    for k, v in meta.items():
        if not v:
            continue
        pair.append((k, v))
        if len(pair) == 2:
            rows.append(pair); pair = []
    if pair:
        rows.append(pair)
    return rows


def _pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(COLOR_ACCENT)
    canvas.setLineWidth(2)
    canvas.line(2 * cm, 1.6 * cm, A4[0] - 2 * cm, 1.6 * cm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#777777"))
    canvas.drawString(2 * cm, 1.2 * cm,
                       f"Dibuat dengan Pendamping ARKAS (tidak resmi) - {datetime.now().strftime('%d %B %Y')}")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Halaman {doc.page}")
    canvas.restoreState()


def meta_table_flowable(meta: dict, styles):
    rows = _meta_pairs(meta)
    if not rows:
        return None
    table_data = []
    for pair in rows:
        row = []
        for k, v in pair:
            row.append(Paragraph(f"<b>{k}</b>", styles["meta_label"]))
            row.append(Paragraph(f": {v}", styles["meta_value"]))
        if len(pair) == 1:
            row += ["", ""]
        table_data.append(row)
    t = Table(table_data, colWidths=[3.4 * cm, 5.4 * cm, 3.0 * cm, 5.0 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.7, COLOR_PRIMARY),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def df_ke_pdf_table(df: pd.DataFrame, styles, col_widths=None, header_bg=COLOR_PRIMARY):
    data = [[Paragraph(f"<b>{c}</b>", styles["cell"]) for c in df.columns]]
    for _, row in df.iterrows():
        data.append([Paragraph(str(v), styles["cell"]) for v in row])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, COLOR_LIGHT]),
    ]))
    return t


def to_excel_bytes(sheets: dict) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for nama_sheet, df in sheets.items():
            df.to_excel(writer, sheet_name=nama_sheet[:31], index=False)
    buffer.seek(0)
    return buffer.getvalue()


def download_row_pdf_excel(base_filename: str, pdf_bytes: bytes, excel_bytes: bytes, key_prefix: str):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Download PDF (ringkasan cetak)", data=pdf_bytes,
                            file_name=f"{base_filename}.pdf", mime="application/pdf",
                            key=f"{key_prefix}_pdf", use_container_width=True)
    with c2:
        st.download_button("⬇️ Download Excel (angka siap disalin ke ARKAS)", data=excel_bytes,
                            file_name=f"{base_filename}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"{key_prefix}_xlsx", use_container_width=True)


def cek_kepatuhan(rekap_per_komponen: dict, total_pagu: float, sumber_dana: str, status_sekolah: str):
    """Bandingkan rekap Rp per komponen terhadap ambang batas Juknis BOSP 2026.
    Mengembalikan list of dict: Komponen, Jumlah, Persen, Ketentuan, Status."""
    hasil = []
    berlaku = sumber_dana in SUMBER_DANA_BERATURAN
    for komponen, aturan in ATURAN_KOMPONEN.items():
        jumlah = rekap_per_komponen.get(komponen, 0)
        persen = (jumlah / total_pagu * 100) if total_pagu > 0 else 0
        if not berlaku:
            ket, status = "Ketentuan khusus (di luar cakupan ringkasan ini)", "ℹ️ Cek Juknis"
        elif aturan["tipe"] == "min":
            batas = aturan["persen"]
            ket = f"Minimal {batas}%"
            status = "✅ Terpenuhi" if persen >= batas else "⚠️ Belum Terpenuhi"
        elif aturan["tipe"] == "max":
            batas = aturan["persen"]
            ket = f"Maksimal {batas}%"
            status = "✅ Terpenuhi" if persen <= batas else "❌ Melebihi Batas"
        elif aturan["tipe"] == "max_by_status":
            batas = aturan["persen_negeri"] if status_sekolah == "Negeri" else aturan["persen_swasta"]
            ket = f"Maksimal {batas}% ({status_sekolah})"
            status = "✅ Terpenuhi" if persen <= batas else "❌ Melebihi Batas"
        else:
            ket, status = "-", "-"
        hasil.append({"Komponen": komponen, "Jumlah (Rp)": jumlah, "% dari Pagu": round(persen, 1),
                       "Ketentuan": ket, "Status": status})
    return hasil


NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
              "September", "Oktober", "November", "Desember"]
NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def tanggal_indonesia(d: date, dengan_hari: bool = False) -> str:
    """Format tanggal ke Bahasa Indonesia tanpa bergantung locale sistem (yang bisa saja
    default Inggris di server), supaya surat resmi tidak tertulis 'Saturday, August'."""
    teks = f"{d.day} {NAMA_BULAN[d.month - 1]} {d.year}"
    if dengan_hari:
        teks = f"{NAMA_HARI[d.weekday()]}, {teks}"
    return teks


def _siapkan_bku(df: pd.DataFrame) -> pd.DataFrame:
    """Bersihkan tipe data & hitung Saldo BKU berjalan (saldo awal Kas+Bank + kumulatif)."""
    d = df.copy()
    for kol in ["Penerimaan", "Pengeluaran", "PPN", "PPh 21", "PPh 22", "PPh 23"]:
        d[kol] = pd.to_numeric(d[kol], errors="coerce").fillna(0)
    d["Tanggal"] = pd.to_datetime(d["Tanggal"], errors="coerce")
    d = d.sort_values("Tanggal", kind="stable").reset_index(drop=True)
    saldo_awal = float(st.session_state.spj_saldo_awal_kas) + float(st.session_state.spj_saldo_awal_bank)
    d["Saldo"] = saldo_awal + (d["Penerimaan"] - d["Pengeluaran"]).cumsum()
    d["Bulan"] = d["Tanggal"].dt.month
    return d


def build_pembantu_kas(d: pd.DataFrame) -> pd.DataFrame:
    sub = d[d["Metode"] == "Tunai"].copy()
    sub["Saldo Kas"] = float(st.session_state.spj_saldo_awal_kas) + \
        (sub["Penerimaan"] - sub["Pengeluaran"]).cumsum()
    return sub[["Tanggal", "No. Bukti", "Uraian", "Penerimaan", "Pengeluaran", "Saldo Kas"]]


def build_pembantu_bank(d: pd.DataFrame) -> pd.DataFrame:
    sub = d[d["Metode"] == "Bank"].copy()
    sub["Saldo Bank"] = float(st.session_state.spj_saldo_awal_bank) + \
        (sub["Penerimaan"] - sub["Pengeluaran"]).cumsum()
    return sub[["Tanggal", "No. Bukti", "Uraian", "Penerimaan", "Pengeluaran", "Saldo Bank"]]


def build_pembantu_pajak(d: pd.DataFrame) -> pd.DataFrame:
    sub = d[(d["PPN"] > 0) | (d["PPh 21"] > 0) | (d["PPh 22"] > 0) | (d["PPh 23"] > 0)].copy()
    sub["Total Dipungut"] = sub[["PPN", "PPh 21", "PPh 22", "PPh 23"]].sum(axis=1)
    sub["Status"] = sub["Pajak Sudah Disetor"].map({True: "Sudah Disetor", False: "Belum Disetor"})
    return sub[["Tanggal", "No. Bukti", "Uraian", "PPN", "PPh 21", "PPh 22", "PPh 23",
                "Total Dipungut", "Status"]]


def build_pembantu_rob(d: pd.DataFrame) -> pd.DataFrame:
    rekap = d.groupby(["Kode Rekening", "Komponen"], as_index=False)["Pengeluaran"].sum()
    rekap["Pagu"] = rekap["Komponen"].map(st.session_state.spj_pagu_komponen).fillna(0)
    rekap["Sisa Anggaran"] = rekap["Pagu"] - rekap["Pengeluaran"]
    rekap = rekap.rename(columns={"Pengeluaran": "Realisasi"})
    return rekap[["Kode Rekening", "Komponen", "Pagu", "Realisasi", "Sisa Anggaran"]]


def build_register_kas(d: pd.DataFrame, bulan_ke: int) -> dict:
    sub = d[d["Bulan"] <= bulan_ke]
    kas = build_pembantu_kas(_siapkan_bku(d[d["Bulan"] <= bulan_ke]))
    bank = build_pembantu_bank(_siapkan_bku(d[d["Bulan"] <= bulan_ke]))
    saldo_kas = float(kas["Saldo Kas"].iloc[-1]) if len(kas) else float(st.session_state.spj_saldo_awal_kas)
    saldo_bank = float(bank["Saldo Bank"].iloc[-1]) if len(bank) else float(st.session_state.spj_saldo_awal_bank)
    return {"Tanggal Penutupan": f"Akhir {NAMA_BULAN[bulan_ke - 1]}",
            "Saldo Kas Tunai": saldo_kas, "Saldo Bank": saldo_bank,
            "Saldo Kas Umum (Total)": saldo_kas + saldo_bank}


def build_lra(d: pd.DataFrame) -> pd.DataFrame:
    rekap = d.groupby("Komponen", as_index=False)["Pengeluaran"].sum()
    rekap["Pagu"] = rekap["Komponen"].map(st.session_state.spj_pagu_komponen).fillna(0)
    rekap = rekap.rename(columns={"Pengeluaran": "Realisasi"})
    rekap["Sisa Anggaran"] = rekap["Pagu"] - rekap["Realisasi"]
    rekap["% Realisasi"] = (rekap["Realisasi"] / rekap["Pagu"].replace(0, pd.NA) * 100).fillna(0).round(1)
    total = pd.DataFrame([{"Komponen": "TOTAL", "Pagu": rekap["Pagu"].sum(),
                            "Realisasi": rekap["Realisasi"].sum(),
                            "Sisa Anggaran": rekap["Sisa Anggaran"].sum(),
                            "% Realisasi": round(rekap["Realisasi"].sum() / max(rekap["Pagu"].sum(), 1) * 100, 1)}])
    return pd.concat([rekap[["Komponen", "Pagu", "Realisasi", "Sisa Anggaran", "% Realisasi"]], total],
                      ignore_index=True)


def build_rprpd(d: pd.DataFrame) -> pd.DataFrame:
    baris = []
    saldo = float(st.session_state.spj_saldo_awal_kas) + float(st.session_state.spj_saldo_awal_bank)
    for b in range(1, 13):
        sub = d[d["Bulan"] == b]
        terima, keluar = float(sub["Penerimaan"].sum()), float(sub["Pengeluaran"].sum())
        saldo += terima - keluar
        baris.append({"Bulan": NAMA_BULAN[b - 1], "Penerimaan": terima,
                       "Pengeluaran": keluar, "Saldo Akhir Bulan": saldo})
    return pd.DataFrame(baris)


def teks_sptm(bulan_ke: int, total_realisasi: float) -> str:
    s = st.session_state
    rp = format_rupiah(total_realisasi)
    return f"""SURAT PERNYATAAN TANGGUNG JAWAB MUTLAK (SPTM)

Yang bertanda tangan di bawah ini:
Nama            : {s.kepsek_nama or '..............................'}
Jabatan         : Kepala Sekolah {s.sekolah or '..............................'}

Dengan ini menyatakan dengan sesungguhnya bahwa dana Bantuan Operasional Sekolah (BOS) yang
diterima {s.sekolah or '(Nama Sekolah)'} pada bulan {NAMA_BULAN[bulan_ke - 1]} Tahun Anggaran
{s.tahun_anggaran} telah digunakan sebesar {rp} sesuai peruntukan dan ketentuan yang berlaku
dalam Petunjuk Teknis BOSP {s.tahun_anggaran}, serta bertanggung jawab penuh atas kebenaran
penggunaan dana tersebut.

Apabila di kemudian hari ditemukan penyimpangan dalam penggunaan dana ini, kami bersedia
bertanggung jawab sesuai ketentuan peraturan perundang-undangan yang berlaku.

{s.spj_kabupaten or '.....'}, {tanggal_indonesia(date.today())}
Kepala Sekolah,


{s.kepsek_nama or '..............................'}
NIP. {s.spj_kepsek_nip or '..............................'}
"""


def teks_sptmh(bulan_ke: int, total_honor: float) -> str:
    s = st.session_state
    rp = format_rupiah(total_honor)
    return f"""SURAT PERNYATAAN TANGGUNG JAWAB MUTLAK HONORARIUM (SPTMH)

Yang bertanda tangan di bawah ini:
Nama            : {s.kepsek_nama or '..............................'}
Jabatan         : Kepala Sekolah {s.sekolah or '..............................'}

Menyatakan bahwa pembayaran honorarium dari dana BOS pada bulan {NAMA_BULAN[bulan_ke - 1]}
Tahun Anggaran {s.tahun_anggaran} sebesar {rp} telah dibayarkan kepada guru/tenaga
kependidikan non-ASN yang BELUM menerima tunjangan profesi, sesuai dengan ketentuan
maksimal alokasi honor pada Juknis BOSP {s.tahun_anggaran} (maksimal 20% untuk sekolah
negeri / 40% untuk sekolah swasta dari total pagu alokasi tahunan).

Kami bertanggung jawab penuh atas kebenaran data ini.

{s.spj_kabupaten or '.....'}, {tanggal_indonesia(date.today())}
Kepala Sekolah,


{s.kepsek_nama or '..............................'}
NIP. {s.spj_kepsek_nip or '..............................'}
"""


def teks_bap_kas(bulan_ke: int, saldo_kas: float, saldo_bank: float) -> str:
    s = st.session_state
    rp_kas, rp_bank = format_rupiah(saldo_kas), format_rupiah(saldo_bank)
    rp_total = format_rupiah(saldo_kas + saldo_bank)
    return f"""BERITA ACARA PEMERIKSAAN KAS (BAP KAS)

Pada hari ini, {tanggal_indonesia(date.today(), dengan_hari=True)}, bulan {NAMA_BULAN[bulan_ke - 1]}
Tahun {s.tahun_anggaran}, yang bertanda tangan di bawah ini:

1. Nama    : {s.kepsek_nama or '..............................'}
   Jabatan : Kepala Sekolah (Pemeriksa)

2. Nama    : {s.bendahara_nama or '..............................'}
   Jabatan : Bendahara BOS

Telah melakukan pemeriksaan Kas BOS {s.sekolah or '(Nama Sekolah)'} dengan hasil sebagai berikut:

  Saldo Kas Tunai   : {rp_kas}
  Saldo Bank        : {rp_bank}
  Saldo Kas Umum    : {rp_total}

Demikian Berita Acara Pemeriksaan Kas ini dibuat dengan sesungguhnya untuk dipergunakan
sebagaimana mestinya.

{s.spj_kabupaten or '.....'}, {tanggal_indonesia(date.today())}

Kepala Sekolah,                              Bendahara BOS,


{s.kepsek_nama or '..............................'}              {s.bendahara_nama or '..............................'}
NIP. {s.spj_kepsek_nip or '..........................'}          NIP. {s.spj_bendahara_nip or '..........................'}
"""


def build_validasi_rekening(d: pd.DataFrame) -> pd.DataFrame:
    kas = build_pembantu_kas(d)
    bank = build_pembantu_bank(d)
    pajak = build_pembantu_pajak(d)
    saldo_kas = float(kas["Saldo Kas"].iloc[-1]) if len(kas) else float(st.session_state.spj_saldo_awal_kas)
    saldo_bank = float(bank["Saldo Bank"].iloc[-1]) if len(bank) else float(st.session_state.spj_saldo_awal_bank)
    saldo_bku = float(d["Saldo"].iloc[-1]) if len(d) else \
        float(st.session_state.spj_saldo_awal_kas) + float(st.session_state.spj_saldo_awal_bank)
    total_dipungut = float(pajak["Total Dipungut"].sum()) if len(pajak) else 0.0
    total_disetor = float(pajak.loc[pajak["Status"] == "Sudah Disetor", "Total Dipungut"].sum()) if len(pajak) else 0.0
    baris = [
        {"Pemeriksaan": "Saldo BKU = Saldo Kas Tunai + Saldo Bank",
         "Nilai 1": saldo_bku, "Nilai 2": saldo_kas + saldo_bank,
         "Status": "✅ Cocok" if abs(saldo_bku - (saldo_kas + saldo_bank)) < 1 else "❌ SELISIH — cek input BKU"},
        {"Pemeriksaan": "Total Pajak Dipungut ≥ Total Pajak Disetor",
         "Nilai 1": total_dipungut, "Nilai 2": total_disetor,
         "Status": "✅ Wajar" if total_dipungut >= total_disetor else "❌ Disetor melebihi dipungut — cek input"},
    ]
    return pd.DataFrame(baris)


def pie_chart_alokasi(rekap: dict, judul: str):
    rekap_bersih = {k: v for k, v in rekap.items() if v > 0}
    if not rekap_bersih:
        return None
    fig, ax = plt.subplots(figsize=(5, 4))
    warna = plt.cm.Greens_r(range(30, 30 + 40 * len(rekap_bersih), 40)) if len(rekap_bersih) > 1 else ["#0F6E4F"]
    ax.pie(rekap_bersih.values(), labels=rekap_bersih.keys(), autopct="%1.1f%%",
           colors=plt.cm.tab20c.colors[:len(rekap_bersih)], textprops={"fontsize": 8})
    ax.set_title(judul, fontsize=10)
    fig.tight_layout()
    return fig


# ============================================================
# TABS UTAMA
# ============================================================
tab_sim, tab_analisis, tab_pajak, tab_spj, tab_checklist, tab_panduan = st.tabs(
    ["🧮 Simulator RKAS", "📊 Analisis Laporan (BKU/RKAS)", "🧾 Kalkulator PPN",
     "📋 SPJ/LPJ BOS (13 Modul)", "📅 Checklist Tahapan", "📖 Panduan Singkat BOSP 2026"]
)

# ============================================================
# TAB 1 - SIMULATOR RKAS
# ============================================================
with tab_sim:
    st.subheader("Simulator Penyusunan RKAS")
    st.caption("Susun rincian rencana kegiatan di sini, cek otomatis kepatuhannya terhadap batas "
               "komponen Juknis BOSP 2026, lalu unduh sebagai acuan saat menginput manual ke ARKAS.")

    c1, c2 = st.columns(2)
    with c1:
        sim_sumber_dana = st.selectbox("Sumber Dana", SUMBER_DANA_OPSI, key="sim_sumber_dana")
    with c2:
        sim_total_pagu = st.number_input("Total Pagu Alokasi (Rp)", min_value=0, step=1_000_000,
                                          value=0, key="sim_total_pagu")

    if sim_sumber_dana not in SUMBER_DANA_BERATURAN:
        st.info("ℹ️ Ringkasan batas komponen di aplikasi ini berlaku untuk BOS/BOP Reguler. "
                "Untuk sumber dana ini, komponen & batasannya berbeda — cek Juknis khusus terkait.")

    st.markdown("**📋 Rincian Rencana Kegiatan**")
    st.caption("Klik baris terakhir/tanda + di bawah tabel untuk menambah baris. Isi kolom Komponen, "
               "Uraian Kegiatan, dan Jumlah (Rp) untuk tiap rencana belanja.")
    edited_df = st.data_editor(
        st.session_state.rincian_rkas,
        num_rows="dynamic", use_container_width=True, key="editor_rkas",
        column_config={
            "Komponen": st.column_config.SelectboxColumn("Komponen", options=SEMUA_KOMPONEN, required=True),
            "Uraian Kegiatan": st.column_config.TextColumn("Uraian Kegiatan", width="large"),
            "Jumlah (Rp)": st.column_config.NumberColumn("Jumlah (Rp)", min_value=0, step=50_000, format="%d"),
            "Kena PPN 12%": st.column_config.CheckboxColumn("Kena PPN 12%"),
        },
    )
    st.session_state.rincian_rkas = edited_df

    if st.button("🧮 Hitung & Cek Kepatuhan", key="btn_hitung_sim", use_container_width=True):
        df = edited_df.dropna(subset=["Komponen"]).copy()
        df["Jumlah (Rp)"] = pd.to_numeric(df["Jumlah (Rp)"], errors="coerce").fillna(0)
        total_dianggarkan = df["Jumlah (Rp)"].sum()
        rekap_komponen = df.groupby("Komponen")["Jumlah (Rp)"].sum().to_dict()

        colm1, colm2, colm3 = st.columns(3)
        colm1.metric("Total Pagu", format_rupiah(sim_total_pagu))
        colm2.metric("Total Dianggarkan", format_rupiah(total_dianggarkan))
        sisa = sim_total_pagu - total_dianggarkan
        colm3.metric("Sisa Pagu Belum Dialokasikan", format_rupiah(sisa),
                     delta=None if sim_total_pagu == 0 else f"{sisa/sim_total_pagu*100:.1f}% dari pagu")
        if sim_total_pagu > 0 and total_dianggarkan > sim_total_pagu:
            st.error(f"❌ Total rencana kegiatan ({format_rupiah(total_dianggarkan)}) MELEBIHI total pagu "
                     f"({format_rupiah(sim_total_pagu)}). Kurangi/sesuaikan rincian kegiatan.")

        st.divider()
        st.markdown("**✅ Cek Kepatuhan Komponen (Ringkasan Juknis BOSP 2026)**")
        hasil_kepatuhan = cek_kepatuhan(rekap_komponen, sim_total_pagu, sim_sumber_dana,
                                         st.session_state.status_sekolah)
        df_kepatuhan = pd.DataFrame(hasil_kepatuhan)
        df_kepatuhan_tampil = df_kepatuhan.copy()
        df_kepatuhan_tampil["Jumlah (Rp)"] = df_kepatuhan_tampil["Jumlah (Rp)"].apply(format_rupiah)
        st.dataframe(df_kepatuhan_tampil, use_container_width=True, hide_index=True)
        if any("❌" in h["Status"] for h in hasil_kepatuhan):
            st.error("❌ Ada komponen yang melebihi batas maksimal Juknis. Sesuaikan rincian kegiatan "
                     "sebelum diinput ke ARKAS.")
        elif any("⚠️" in h["Status"] for h in hasil_kepatuhan):
            st.warning("⚠️ Ada komponen yang belum memenuhi batas minimal Juknis.")
        else:
            st.success("✅ Semua komponen yang tercakup ringkasan ini sudah sesuai batas Juknis.")

        st.divider()
        colp1, colp2 = st.columns(2)
        with colp1:
            fig = pie_chart_alokasi(rekap_komponen, "Proporsi Alokasi per Komponen")
            if fig:
                st.pyplot(fig, use_container_width=True)
        with colp2:
            df_ppn = df[df["Kena PPN 12%"] == True].copy()
            if not df_ppn.empty:
                df_ppn["DPP (Rp)"] = df_ppn["Jumlah (Rp)"] / 1.12
                df_ppn["PPN 12% (Rp)"] = df_ppn["Jumlah (Rp)"] - df_ppn["DPP (Rp)"]
                st.markdown("**🧾 Estimasi PPN dari Kegiatan Bertanda 'Kena PPN'**")
                st.metric("Total Estimasi PPN yang Harus Disetor", format_rupiah(df_ppn["PPN 12% (Rp)"].sum()))
                st.caption("Estimasi ini mengasumsikan Jumlah (Rp) yang diisi SUDAH termasuk PPN 12%. "
                           "Selalu cocokkan dengan bukti transaksi & aturan pajak yang berlaku saat pencatatan di BKU.")

        # Simpan hasil ke session_state agar bisa diunduh
        st.session_state["_sim_hasil"] = {
            "df_rincian": df, "df_kepatuhan": df_kepatuhan, "total_pagu": sim_total_pagu,
            "total_dianggarkan": total_dianggarkan, "sumber_dana": sim_sumber_dana,
        }

    if "_sim_hasil" in st.session_state:
        st.divider()
        st.markdown("**⬇️ Unduh Kertas Kerja Simulasi**")
        hsl = st.session_state["_sim_hasil"]
        meta = {
            "Satuan Pendidikan": st.session_state.sekolah, "NPSN": st.session_state.npsn,
            "Sumber Dana": hsl["sumber_dana"], "Tahun Anggaran": st.session_state.tahun_anggaran,
            "Status Sekolah": st.session_state.status_sekolah,
            "Kepala Sekolah": st.session_state.kepsek_nama, "Bendahara": st.session_state.bendahara_nama,
        }
        styles = build_pdf_styles()
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.8 * cm, bottomMargin=2.1 * cm,
                                 leftMargin=2 * cm, rightMargin=2 * cm, title="SIMULASI RKAS")
        story = [
            Paragraph("SIMULASI KERTAS KERJA RKAS", styles["title"]),
            Paragraph("Dokumen bantu internal - bukan dokumen resmi ARKAS", styles["subtitle"]),
        ]
        mt = meta_table_flowable(meta, styles)
        if mt:
            story += [mt, Spacer(1, 10)]
        story.append(Paragraph(f"Total Pagu: {format_rupiah(hsl['total_pagu'])} &bull; "
                                f"Total Dianggarkan: {format_rupiah(hsl['total_dianggarkan'])}", styles["h2"]))
        df_show = hsl["df_rincian"][["Komponen", "Uraian Kegiatan", "Jumlah (Rp)"]].copy()
        df_show["Jumlah (Rp)"] = df_show["Jumlah (Rp)"].apply(format_rupiah)
        story.append(df_ke_pdf_table(df_show, styles, col_widths=[4.5*cm, 8.5*cm, 4*cm]))
        story.append(Spacer(1, 12))
        story.append(Paragraph("Ringkasan Kepatuhan Komponen", styles["h2"]))
        df_kep_show = hsl["df_kepatuhan"].copy()
        df_kep_show["Jumlah (Rp)"] = df_kep_show["Jumlah (Rp)"].apply(format_rupiah)
        df_kep_show["% dari Pagu"] = df_kep_show["% dari Pagu"].astype(str) + "%"
        story.append(df_ke_pdf_table(df_kep_show, styles, col_widths=[4.3*cm, 3*cm, 2.3*cm, 4*cm, 3.4*cm]))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Catatan: dokumen ini adalah alat bantu perhitungan internal, bukan "
                                "pengganti ARKAS resmi. Nilai akhir tetap harus diinput & disahkan "
                                "melalui Aplikasi ARKAS Kemendikdasmen.", styles["note"]))
        doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
        buffer.seek(0)
        pdf_bytes = buffer.getvalue()

        excel_bytes = to_excel_bytes({
            "Rincian Kegiatan": hsl["df_rincian"][["Komponen", "Uraian Kegiatan", "Jumlah (Rp)", "Kena PPN 12%"]],
            "Ringkasan Kepatuhan": hsl["df_kepatuhan"],
        })
        download_row_pdf_excel("Simulasi_RKAS", pdf_bytes, excel_bytes, "dl_sim")

# ============================================================
# TAB 2 - ANALISIS LAPORAN (IMPOR BKU/RKAS DARI ARKAS)
# ============================================================
KATA_KUNCI_KOMPONEN = {
    "Pengembangan Perpustakaan / Buku": ["buku", "pustaka", "perpustakaan"],
    "Pembayaran Honor": ["honor", "honorarium"],
    "Pemeliharaan Sarana & Prasarana": ["pemeliharaan", "sarpras", "sarana", "prasarana", "perbaikan"],
    "Langganan Daya & Jasa": ["listrik", "internet", "telepon", "air", "daya", "jasa"],
    "Pengembangan Mutu Guru & Tendik": ["pelatihan", "workshop", "bimtek", "diklat", "mgmp", "kkg"],
    "Kegiatan Pembelajaran & Ekstrakurikuler": ["pembelajaran", "ekstrakurikuler", "lomba", "kegiatan siswa"],
    "Kegiatan Asesmen & Evaluasi": ["asesmen", "ujian", "penilaian", "evaluasi"],
    "Administrasi & Pelaporan": ["atk", "administrasi", "pelaporan", "fotokopi", "cetak"],
}


def klasifikasi_otomatis(uraian: str) -> str:
    u = str(uraian).lower()
    for komponen, kata_kunci in KATA_KUNCI_KOMPONEN.items():
        if any(kk in u for kk in kata_kunci):
            return komponen
    return "Lainnya"


with tab_analisis:
    st.subheader("Analisis Laporan dari Ekspor BKU / RKAS ARKAS")
    st.caption("Unggah hasil ekspor Excel/CSV dari ARKAS (Buku Kas Umum atau RKAS), lalu aplikasi ini "
               "membuat rekap otomatis, grafik realisasi, dan cek kepatuhan tanpa perlu hitung manual.")

    ac1, ac2 = st.columns(2)
    with ac1:
        an_sumber_dana = st.selectbox("Sumber Dana", SUMBER_DANA_OPSI, key="an_sumber_dana")
    with ac2:
        an_total_pagu = st.number_input("Total Pagu Alokasi (Rp) — untuk cek kepatuhan", min_value=0,
                                         step=1_000_000, value=0, key="an_total_pagu")

    file_bku = st.file_uploader("📎 Unggah berkas ekspor BKU/RKAS (XLSX/CSV)", type=["xlsx", "xls", "csv"],
                                 key="file_bku")

    if file_bku is not None:
        try:
            if file_bku.name.lower().endswith(".csv"):
                df_bku = pd.read_csv(file_bku)
            else:
                df_bku = pd.read_excel(file_bku)
            df_bku = df_bku.dropna(how="all")
        except Exception as e:
            st.error(f"❌ Gagal membaca berkas: {e}")
            df_bku = None

        if df_bku is not None and not df_bku.empty:
            st.success(f"✅ Berhasil membaca {len(df_bku)} baris data dari '{file_bku.name}'.")
            with st.expander("👀 Pratinjau data mentah"):
                st.dataframe(df_bku.head(50), use_container_width=True)

            st.markdown("**🔗 Pemetaan Kolom** — cocokkan kolom di berkas Anda dengan kebutuhan analisis")
            kolom_opsi = ["(tidak ada)"] + list(df_bku.columns)
            mc1, mc2, mc3 = st.columns(3)
            with mc1:
                kol_uraian = st.selectbox("Kolom Uraian/Kegiatan", kolom_opsi,
                                           index=kolom_opsi.index(df_bku.columns[0]) if len(df_bku.columns) else 0,
                                           key="kol_uraian")
            with mc2:
                kol_penerimaan = st.selectbox("Kolom Penerimaan (opsional)", kolom_opsi, key="kol_penerimaan")
            with mc3:
                kol_pengeluaran = st.selectbox("Kolom Pengeluaran (opsional)", kolom_opsi, key="kol_pengeluaran")
            kol_kategori = st.selectbox(
                "Kolom Kategori/Komponen (opsional — kosongkan agar diklasifikasi otomatis dari Uraian)",
                kolom_opsi, key="kol_kategori")

            if st.button("📊 Proses & Buat Rekap", key="btn_proses_bku", use_container_width=True):
                df_kerja = pd.DataFrame()
                df_kerja["Uraian"] = df_bku[kol_uraian] if kol_uraian != "(tidak ada)" else ""
                df_kerja["Penerimaan"] = pd.to_numeric(df_bku[kol_penerimaan], errors="coerce").fillna(0) \
                    if kol_penerimaan != "(tidak ada)" else 0
                df_kerja["Pengeluaran"] = pd.to_numeric(df_bku[kol_pengeluaran], errors="coerce").fillna(0) \
                    if kol_pengeluaran != "(tidak ada)" else 0
                if kol_kategori != "(tidak ada)":
                    df_kerja["Komponen"] = df_bku[kol_kategori].fillna("Lainnya")
                else:
                    df_kerja["Komponen"] = df_kerja["Uraian"].apply(klasifikasi_otomatis)

                total_terima = df_kerja["Penerimaan"].sum()
                total_keluar = df_kerja["Pengeluaran"].sum()
                colr1, colr2, colr3 = st.columns(3)
                colr1.metric("Total Penerimaan", format_rupiah(total_terima))
                colr2.metric("Total Pengeluaran", format_rupiah(total_keluar))
                colr3.metric("Saldo", format_rupiah(total_terima - total_keluar))

                st.divider()
                rekap_komponen = df_kerja.groupby("Komponen")["Pengeluaran"].sum()
                rekap_komponen = rekap_komponen[rekap_komponen > 0].to_dict()
                colv1, colv2 = st.columns([1, 1])
                with colv1:
                    st.markdown("**📋 Rekap Pengeluaran per Komponen**")
                    df_rekap_tampil = pd.DataFrame([
                        {"Komponen": k, "Jumlah (Rp)": format_rupiah(v),
                         "% dari Total Pengeluaran": f"{(v/total_keluar*100):.1f}%" if total_keluar else "0%"}
                        for k, v in sorted(rekap_komponen.items(), key=lambda x: -x[1])
                    ])
                    st.dataframe(df_rekap_tampil, use_container_width=True, hide_index=True)
                with colv2:
                    fig = pie_chart_alokasi(rekap_komponen, "Proporsi Realisasi per Komponen")
                    if fig:
                        st.pyplot(fig, use_container_width=True)

                st.divider()
                st.markdown("**✅ Cek Kepatuhan Realisasi terhadap Pagu**")
                if an_total_pagu > 0:
                    hasil_kepatuhan_an = cek_kepatuhan(rekap_komponen, an_total_pagu, an_sumber_dana,
                                                        st.session_state.status_sekolah)
                    df_kep_an = pd.DataFrame(hasil_kepatuhan_an)
                    df_kep_an_tampil = df_kep_an.copy()
                    df_kep_an_tampil["Jumlah (Rp)"] = df_kep_an_tampil["Jumlah (Rp)"].apply(format_rupiah)
                    st.dataframe(df_kep_an_tampil, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ Isi Total Pagu Alokasi di atas untuk mengaktifkan cek kepatuhan otomatis.")
                    df_kep_an = pd.DataFrame()

                # Simpan untuk unduhan
                st.session_state["_an_hasil"] = {
                    "df_kerja": df_kerja, "rekap_komponen": rekap_komponen,
                    "total_terima": total_terima, "total_keluar": total_keluar,
                    "df_kepatuhan": df_kep_an, "sumber_dana": an_sumber_dana,
                }

    if "_an_hasil" in st.session_state:
        st.divider()
        st.markdown("**⬇️ Unduh Laporan Analisis**")
        hsl = st.session_state["_an_hasil"]
        meta = {
            "Satuan Pendidikan": st.session_state.sekolah, "NPSN": st.session_state.npsn,
            "Sumber Dana": hsl["sumber_dana"], "Tahun Anggaran": st.session_state.tahun_anggaran,
            "Bendahara": st.session_state.bendahara_nama, "Kepala Sekolah": st.session_state.kepsek_nama,
        }
        styles = build_pdf_styles()
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.8 * cm, bottomMargin=2.1 * cm,
                                 leftMargin=2 * cm, rightMargin=2 * cm, title="ANALISIS LAPORAN BKU/RKAS")
        story = [
            Paragraph("ANALISIS LAPORAN REALISASI BKU/RKAS", styles["title"]),
            Paragraph("Dokumen bantu internal - bukan dokumen resmi ARKAS", styles["subtitle"]),
        ]
        mt = meta_table_flowable(meta, styles)
        if mt:
            story += [mt, Spacer(1, 10)]
        story.append(Paragraph(f"Total Penerimaan: {format_rupiah(hsl['total_terima'])} &bull; "
                                f"Total Pengeluaran: {format_rupiah(hsl['total_keluar'])} &bull; "
                                f"Saldo: {format_rupiah(hsl['total_terima'] - hsl['total_keluar'])}", styles["h2"]))
        df_rekap_pdf = pd.DataFrame([{"Komponen": k, "Jumlah (Rp)": format_rupiah(v)}
                                      for k, v in hsl["rekap_komponen"].items()])
        if not df_rekap_pdf.empty:
            story.append(df_ke_pdf_table(df_rekap_pdf, styles, col_widths=[9 * cm, 8 * cm]))
        if not hsl["df_kepatuhan"].empty:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Ringkasan Kepatuhan Komponen", styles["h2"]))
            df_kep_pdf = hsl["df_kepatuhan"].copy()
            df_kep_pdf["Jumlah (Rp)"] = df_kep_pdf["Jumlah (Rp)"].apply(format_rupiah)
            df_kep_pdf["% dari Pagu"] = df_kep_pdf["% dari Pagu"].astype(str) + "%"
            story.append(df_ke_pdf_table(df_kep_pdf, styles, col_widths=[4.3*cm, 3*cm, 2.3*cm, 4*cm, 3.4*cm]))
        story.append(Spacer(1, 10))
        story.append(Paragraph("Catatan: dokumen ini adalah alat bantu analisis internal berdasarkan "
                                "data yang diunggah pengguna, bukan laporan resmi ARKAS.", styles["note"]))
        doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
        buffer.seek(0)
        pdf_bytes_an = buffer.getvalue()

        excel_bytes_an = to_excel_bytes({
            "Data Transaksi": hsl["df_kerja"],
            "Rekap Komponen": pd.DataFrame([{"Komponen": k, "Jumlah (Rp)": v}
                                             for k, v in hsl["rekap_komponen"].items()]),
        })
        download_row_pdf_excel("Analisis_Laporan_BKU", pdf_bytes_an, excel_bytes_an, "dl_an")

# ============================================================
# TAB 3 - KALKULATOR PPN
# ============================================================
with tab_pajak:
    st.subheader("Kalkulator PPN 12%")
    st.caption("Bantu hitung Dasar Pengenaan Pajak (DPP) dan PPN untuk dicatat di BKU, sesuai "
               "penyesuaian tarif PPN 12% (PMK No. 131 Tahun 2024 & PMK No. 11 Tahun 2025).")

    mode_ppn = st.radio("Nilai yang saya masukkan adalah:",
                         ["Total belanja SUDAH termasuk PPN", "Harga barang SEBELUM PPN (DPP)"],
                         key="mode_ppn")
    nilai_input = st.number_input("Nilai (Rp)", min_value=0, step=10_000, value=0, key="nilai_ppn")

    if nilai_input > 0:
        if mode_ppn == "Total belanja SUDAH termasuk PPN":
            dpp = nilai_input / 1.12
            ppn = nilai_input - dpp
            total = nilai_input
        else:
            dpp = nilai_input
            ppn = dpp * 0.12
            total = dpp + ppn

        colp1, colp2, colp3 = st.columns(3)
        colp1.metric("DPP (Dasar Pengenaan Pajak)", format_rupiah(dpp))
        colp2.metric("PPN 12%", format_rupiah(ppn))
        colp3.metric("Total Dibayar/Dicatat", format_rupiah(total))
        st.caption("💡 Hasil ini estimasi untuk membantu pencatatan — selalu cocokkan dengan nilai pada "
                   "faktur pajak/nota resmi dari penyedia barang & jasa.")

    st.divider()
    st.markdown("**📋 Kalkulator Banyak Item Sekaligus**")
    df_ppn_massal = st.data_editor(
        pd.DataFrame([{"Uraian": "", "Nilai Termasuk PPN (Rp)": 0}]),
        num_rows="dynamic", use_container_width=True, key="editor_ppn_massal",
        column_config={
            "Nilai Termasuk PPN (Rp)": st.column_config.NumberColumn(min_value=0, step=10_000, format="%d"),
        },
    )
    if st.button("🧮 Hitung Semua", key="btn_hitung_ppn_massal"):
        dfp = df_ppn_massal.copy()
        dfp["Nilai Termasuk PPN (Rp)"] = pd.to_numeric(dfp["Nilai Termasuk PPN (Rp)"], errors="coerce").fillna(0)
        dfp["DPP (Rp)"] = dfp["Nilai Termasuk PPN (Rp)"] / 1.12
        dfp["PPN 12% (Rp)"] = dfp["Nilai Termasuk PPN (Rp)"] - dfp["DPP (Rp)"]
        dfp_tampil = dfp.copy()
        for kol in ["DPP (Rp)", "PPN 12% (Rp)", "Nilai Termasuk PPN (Rp)"]:
            dfp_tampil[kol] = dfp_tampil[kol].apply(format_rupiah)
        st.dataframe(dfp_tampil, use_container_width=True, hide_index=True)
        st.metric("Total PPN yang Harus Disetor", format_rupiah(dfp["PPN 12% (Rp)"].sum()))

# ============================================================
# TAB 4 - CHECKLIST TAHAPAN ARKAS/RKAS
# ============================================================
# ============================================================
# TAB SPJ/LPJ BOS - 13 MODUL (mengikuti struktur file LPJ BOS pada umumnya)
# ============================================================
with tab_spj:
    st.subheader("📋 SPJ/LPJ BOS — Otomatis dari BKU")
    st.info("⚠️ Ini alat bantu MENYUSUN draf SPJ/LPJ, bukan pengganti ARKAS. Cocokkan selalu "
            "angka akhir ke ARKAS resmi sebelum ditandatangani/diserahkan ke Dinas Pendidikan.")

    (sp_data, sp_bku, sp_kas, sp_bank, sp_pajak, sp_rob, sp_reg, sp_bap,
     sp_lra, sp_sptm, sp_sptmh, sp_rprpd, sp_validasi) = st.tabs([
        "1️⃣ Isi Data", "2️⃣ BKU", "3️⃣ P. Kas", "4️⃣ P. Bank", "5️⃣ P. Pajak", "6️⃣ P. R.O.B",
        "7️⃣ Register Kas", "8️⃣ BAP Kas", "9️⃣ LRA", "🔟 SPTM", "1️⃣1️⃣ SPTMH",
        "1️⃣2️⃣ R PRPD", "1️⃣3️⃣ Validasi",
    ])

    with sp_data:
        st.caption("Data ini otomatis dipakai di seluruh 13 modul di bawah. Identitas Sekolah, "
                   "Kepala Sekolah & Bendahara mengikuti isian di Sidebar kiri.")
        c1, c2 = st.columns(2)
        with c1:
            st.session_state.spj_desa = st.text_input("Desa/Kelurahan", st.session_state.spj_desa)
            st.session_state.spj_kecamatan = st.text_input("Kecamatan", st.session_state.spj_kecamatan)
            st.session_state.spj_kabupaten = st.text_input("Kabupaten/Kota", st.session_state.spj_kabupaten)
            st.session_state.spj_provinsi = st.text_input("Provinsi", st.session_state.spj_provinsi)
        with c2:
            st.session_state.spj_kepsek_nip = st.text_input("NIP Kepala Sekolah", st.session_state.spj_kepsek_nip)
            st.session_state.spj_bendahara_nip = st.text_input("NIP Bendahara", st.session_state.spj_bendahara_nip)
            st.session_state.spj_saldo_awal_kas = st.number_input(
                "Saldo Awal Kas Tunai (Rp)", value=int(st.session_state.spj_saldo_awal_kas), step=10000)
            st.session_state.spj_saldo_awal_bank = st.number_input(
                "Saldo Awal Bank (Rp)", value=int(st.session_state.spj_saldo_awal_bank), step=10000)

        st.divider()
        st.markdown("#### 💰 Pagu Anggaran per Komponen (Tahun Anggaran)")
        st.caption("Dipakai untuk menghitung Sisa Anggaran di Pembantu R.O.B & LRA.")
        pagu_rows = [{"Komponen": k, "Pagu (Rp)": v} for k, v in st.session_state.spj_pagu_komponen.items()]
        df_pagu = st.data_editor(pd.DataFrame(pagu_rows), key="editor_pagu", hide_index=True,
                                  use_container_width=True, disabled=["Komponen"])
        st.session_state.spj_pagu_komponen = dict(zip(df_pagu["Komponen"], df_pagu["Pagu (Rp)"]))
        st.metric("Total Pagu Anggaran", format_rupiah(sum(st.session_state.spj_pagu_komponen.values())))

    with sp_bku:
        st.markdown("#### 📖 Buku Kas Umum (BKU) — Input Semua Transaksi di Sini")
        st.caption("Metode: 'Tunai' masuk ke Pembantu Kas, 'Bank' masuk ke Pembantu Bank. "
                   "Isi kolom pajak (PPN/PPh) hanya kalau transaksi ini kena potongan pajak.")
        st.session_state.spj_bku = st.data_editor(
            st.session_state.spj_bku, num_rows="dynamic", key="editor_bku", use_container_width=True,
            column_config={
                "Tanggal": st.column_config.DateColumn("Tanggal", format="DD-MM-YYYY"),
                "Komponen": st.column_config.SelectboxColumn("Komponen", options=SEMUA_KOMPONEN),
                "Metode": st.column_config.SelectboxColumn("Metode", options=["Tunai", "Bank"]),
                "Penerimaan": st.column_config.NumberColumn("Penerimaan", format="%d"),
                "Pengeluaran": st.column_config.NumberColumn("Pengeluaran", format="%d"),
                "PPN": st.column_config.NumberColumn("PPN", format="%d"),
                "PPh 21": st.column_config.NumberColumn("PPh 21", format="%d"),
                "PPh 22": st.column_config.NumberColumn("PPh 22", format="%d"),
                "PPh 23": st.column_config.NumberColumn("PPh 23", format="%d"),
            },
        )
        d_bku = _siapkan_bku(st.session_state.spj_bku)
        tampil = d_bku.copy()
        tampil["Tanggal"] = tampil["Tanggal"].dt.strftime("%d-%m-%Y")
        st.dataframe(tampil[["Tanggal", "No. Bukti", "Uraian", "Komponen", "Metode",
                              "Penerimaan", "Pengeluaran", "Saldo"]], use_container_width=True, hide_index=True)
        cA, cB, cC = st.columns(3)
        cA.metric("Total Penerimaan", format_rupiah(d_bku["Penerimaan"].sum()))
        cB.metric("Total Pengeluaran", format_rupiah(d_bku["Pengeluaran"].sum()))
        cC.metric("Saldo Akhir BKU", format_rupiah(d_bku["Saldo"].iloc[-1] if len(d_bku) else
                                                     st.session_state.spj_saldo_awal_kas + st.session_state.spj_saldo_awal_bank))

    d_bku = _siapkan_bku(st.session_state.spj_bku)  # dipakai ulang di semua sub-tab di bawah

    with sp_kas:
        st.markdown("#### 💵 Buku Pembantu Kas (Tunai)")
        df_kas = build_pembantu_kas(d_bku)
        st.dataframe(df_kas, use_container_width=True, hide_index=True)
        st.metric("Saldo Kas Tunai Saat Ini", format_rupiah(
            df_kas["Saldo Kas"].iloc[-1] if len(df_kas) else st.session_state.spj_saldo_awal_kas))

    with sp_bank:
        st.markdown("#### 🏦 Buku Pembantu Bank")
        df_bank = build_pembantu_bank(d_bku)
        st.dataframe(df_bank, use_container_width=True, hide_index=True)
        st.metric("Saldo Bank Saat Ini", format_rupiah(
            df_bank["Saldo Bank"].iloc[-1] if len(df_bank) else st.session_state.spj_saldo_awal_bank))

    with sp_pajak:
        st.markdown("#### 🧾 Buku Pembantu Pajak")
        df_pajak = build_pembantu_pajak(d_bku)
        if len(df_pajak):
            st.dataframe(df_pajak, use_container_width=True, hide_index=True)
            belum = df_pajak.loc[df_pajak["Status"] == "Belum Disetor", "Total Dipungut"].sum()
            st.metric("Pajak Belum Disetor", format_rupiah(belum))
        else:
            st.info("Belum ada transaksi dengan potongan pajak di BKU.")

    with sp_rob:
        st.markdown("#### 📦 Buku Pembantu Rincian Objek Belanja (R.O.B)")
        df_rob = build_pembantu_rob(d_bku)
        st.dataframe(df_rob, use_container_width=True, hide_index=True)

    with sp_reg:
        st.markdown("#### 🔒 Register Penutupan Kas")
        bulan_reg = st.selectbox("Bulan Penutupan", list(range(1, 13)),
                                  format_func=lambda b: NAMA_BULAN[b - 1], key="bulan_reg")
        reg = build_register_kas(d_bku, bulan_reg)
        for k, v in reg.items():
            if k != "Tanggal Penutupan":
                st.metric(k, format_rupiah(v))
            else:
                st.caption(f"Penutupan: **{v}**")

    with sp_bap:
        st.markdown("#### 📝 Berita Acara Pemeriksaan Kas (BAP Kas)")
        bulan_bap = st.selectbox("Bulan", list(range(1, 13)), format_func=lambda b: NAMA_BULAN[b - 1], key="bulan_bap")
        reg_bap = build_register_kas(d_bku, bulan_bap)
        teks = teks_bap_kas(bulan_bap, reg_bap["Saldo Kas Tunai"], reg_bap["Saldo Bank"])
        st.text_area("Draf Surat (bisa disalin ke Word untuk ditandatangani)", teks, height=420)

    with sp_lra:
        st.markdown("#### 📈 Laporan Realisasi Anggaran (LRA)")
        df_lra = build_lra(d_bku)
        st.dataframe(df_lra.style.format({"Pagu": "Rp {:,.0f}", "Realisasi": "Rp {:,.0f}",
                                           "Sisa Anggaran": "Rp {:,.0f}", "% Realisasi": "{:.1f}%"}),
                     use_container_width=True, hide_index=True)

    with sp_sptm:
        st.markdown("#### ✍️ Surat Pernyataan Tanggung Jawab Mutlak (SPTM)")
        bulan_sptm = st.selectbox("Bulan", list(range(1, 13)), format_func=lambda b: NAMA_BULAN[b - 1], key="bulan_sptm")
        total_bln = float(d_bku.loc[d_bku["Bulan"] == bulan_sptm, "Pengeluaran"].sum())
        teks = teks_sptm(bulan_sptm, total_bln)
        st.text_area("Draf Surat", teks, height=380, key="ta_sptm")

    with sp_sptmh:
        st.markdown("#### ✍️ Surat Pernyataan Tanggung Jawab Mutlak Honorarium (SPTMH)")
        bulan_sptmh = st.selectbox("Bulan", list(range(1, 13)), format_func=lambda b: NAMA_BULAN[b - 1], key="bulan_sptmh")
        total_honor = float(d_bku.loc[(d_bku["Bulan"] == bulan_sptmh) &
                                       (d_bku["Komponen"] == "Pembayaran Honor"), "Pengeluaran"].sum())
        teks = teks_sptmh(bulan_sptmh, total_honor)
        st.text_area("Draf Surat", teks, height=380, key="ta_sptmh")

    with sp_rprpd:
        st.markdown("#### 📆 Rekapitulasi Penerimaan & Realisasi Pengeluaran Dana (R PRPD)")
        df_rprpd = build_rprpd(d_bku)
        st.dataframe(df_rprpd.style.format({"Penerimaan": "Rp {:,.0f}", "Pengeluaran": "Rp {:,.0f}",
                                             "Saldo Akhir Bulan": "Rp {:,.0f}"}),
                     use_container_width=True, hide_index=True)

    with sp_validasi:
        st.markdown("#### ✅ Validasi Data Rekening")
        df_val = build_validasi_rekening(d_bku)
        for _, row in df_val.iterrows():
            st.write(f"**{row['Pemeriksaan']}**")
            vc1, vc2, vc3 = st.columns(3)
            vc1.metric("Nilai 1", format_rupiah(row["Nilai 1"]))
            vc2.metric("Nilai 2", format_rupiah(row["Nilai 2"]))
            vc3.write(row["Status"])
            st.divider()

    st.divider()
    st.markdown("### ⬇️ Unduh Semua Laporan (1 File Excel, 12 Sheet)")
    if st.button("📦 Buat File Excel Lengkap", key="btn_export_spj", use_container_width=True):
        tampil_bku = d_bku.copy()
        tampil_bku["Tanggal"] = tampil_bku["Tanggal"].dt.strftime("%d-%m-%Y")
        sheets = {
            "BKU": tampil_bku[["Tanggal", "No. Bukti", "Uraian", "Komponen", "Metode",
                                "Penerimaan", "Pengeluaran", "Saldo"]],
            "Pembantu Kas": build_pembantu_kas(d_bku),
            "Pembantu Bank": build_pembantu_bank(d_bku),
            "Pembantu Pajak": build_pembantu_pajak(d_bku),
            "Pembantu R.O.B": build_pembantu_rob(d_bku),
            "LRA": build_lra(d_bku),
            "R PRPD": build_rprpd(d_bku),
            "Validasi Rekening": build_validasi_rekening(d_bku),
        }
        excel_bytes = to_excel_bytes(sheets)
        st.download_button("⬇️ Download SPJ_LPJ_BOS.xlsx", excel_bytes,
                            f"SPJ_LPJ_BOS_{st.session_state.sekolah or 'Sekolah'}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)

DAFTAR_TAHAPAN = [
    "Penyusunan Kertas Kerja RKAS T-1 (tahun berikutnya)",
    "Rapat RKAS bersama Kepala Sekolah, Guru, Tendik & Komite Sekolah",
    "Input/Salin Kertas Kerja RKAS di ARKAS 4",
    "Verifikasi & Validasi RKAS oleh Operator/Bendahara",
    "Pengajuan RKAS untuk Persetujuan/Pengesahan Dinas Pendidikan",
    "Sinkronisasi Data Perencanaan ke MARKAS-SIPD",
    "Aktivasi BKU Bulan Berjalan (dilakukan tiap awal bulan, berurutan)",
    "Pencatatan Realisasi Transaksi di BKU (rutin tiap bulan)",
    "Tutup BKU Akhir Bulan",
    "Penyusunan Laporan Semester/Tahunan Penggunaan Dana BOSP",
]

with tab_checklist:
    st.subheader("Checklist Tahapan RKAS/ARKAS")
    st.caption("Pantau progres tahapan administrasi RKAS & ARKAS sepanjang tahun anggaran. "
               "Data checklist ini tersimpan selama sesi browser terbuka.")

    if st.session_state.checklist_tahapan is None:
        st.session_state.checklist_tahapan = pd.DataFrame([
            {"Tahapan": t, "Selesai": False, "Target Tanggal": None, "Catatan": ""}
            for t in DAFTAR_TAHAPAN
        ])

    df_checklist = st.data_editor(
        st.session_state.checklist_tahapan,
        use_container_width=True, num_rows="dynamic", key="editor_checklist",
        column_config={
            "Tahapan": st.column_config.TextColumn("Tahapan", width="large"),
            "Selesai": st.column_config.CheckboxColumn("Selesai"),
            "Target Tanggal": st.column_config.DateColumn("Target Tanggal"),
            "Catatan": st.column_config.TextColumn("Catatan", width="medium"),
        },
    )
    st.session_state.checklist_tahapan = df_checklist

    total_tahap = len(df_checklist)
    selesai_tahap = int(df_checklist["Selesai"].sum()) if total_tahap else 0
    st.progress(selesai_tahap / total_tahap if total_tahap else 0,
                text=f"{selesai_tahap} dari {total_tahap} tahapan selesai")

# ============================================================
# TAB 5 - PANDUAN SINGKAT ATURAN BOSP 2026
# ============================================================
with tab_panduan:
    st.subheader("Panduan Singkat Ketentuan Dana BOSP 2026")
    st.caption("Ringkasan referensi cepat berdasarkan Permendikdasmen No. 8 Tahun 2026 (Juknis BOSP 2026). "
               "Ini ringkasan tidak resmi — selalu rujuk Juknis asli & Dinas Pendidikan setempat untuk "
               "kepastian aturan, terutama untuk kasus khusus/wilayah 3T.")

    st.markdown("""
| Komponen | Ketentuan Umum (BOS/BOP Reguler) |
|---|---|
| **Pengembangan Perpustakaan / Buku** | Wajib **minimal 10%** dari total pagu alokasi untuk penyediaan buku |
| **Pembayaran Honor** (non-ASN, belum tunjangan profesi) | Maksimal **20%** (sekolah negeri) / **40%** (sekolah swasta) dari 100% total pagu alokasi tahunan |
| **Pemeliharaan Sarana & Prasarana** | Maksimal **20%** dari total pagu alokasi |
| **PPN** | Tarif **12%** berlaku atas transaksi belanja kena pajak (PMK No. 131/2024 & PMK No. 11/2025) |
| **Dana BOS Kinerja** | Dialokasikan untuk sekolah berprestasi/berkinerja (10% teratas) dengan komponen & aturan tersendiri — beda dari BOS Reguler |
| **Wilayah 3T** | Berpotensi ada relaksasi batas komponen honor — besaran berbeda tiap daerah, konfirmasi ke Dinas Pendidikan setempat |
""")

    st.divider()
    st.markdown("""
**Alur Perencanaan T-1 (skema terbaru):**
1. Sekolah menyusun Kertas Kerja RKAS untuk tahun anggaran berikutnya (T-1) sebelum tahun berjalan berakhir.
2. Data RKAS di ARKAS 4 disinkronkan ke sistem **MARKAS** milik Dinas Pendidikan.
3. Dinas Pendidikan memantau & menyetujui perencanaan melalui MARKAS.
4. Data diteruskan ke **SIPD** (Sistem Informasi Pemerintahan Daerah) untuk penyusunan APBD terkait.

**Catatan Penting:**
- Persentase minimal/maksimal di atas dihitung dari **total pagu alokasi dalam satu tahun anggaran**, bukan dari realisasi bulanan.
- Aturan dapat direvisi oleh Kemendikdasmen dari waktu ke waktu — pastikan selalu memakai Juknis versi terbaru.
- Aplikasi ini adalah alat bantu tidak resmi. Keputusan akhir & keabsahan dokumen tetap mengacu pada Aplikasi ARKAS resmi dan persetujuan Dinas Pendidikan.
""")
