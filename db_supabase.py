"""
Lapisan penyimpanan permanen (Supabase) untuk Pendamping ARKAS.

Skema pakai: SATU SEKOLAH, TANPA LOGIN. Semua data disimpan dalam 1 baris
tabel `arkas_state` (id=1) berbentuk JSON — jadi data tetap ada walau
aplikasi di-restart / dibuka dari perangkat lain, asal APP_PASSWORD (kalau
diaktifkan) & koneksi Supabase-nya sama.

Kalau nanti perlu banyak sekolah dengan login masing-masing, tabel ini
tinggal diubah jadi banyak baris (1 baris per akun, id = user_id dari
Supabase Auth) — polanya sama, tinggal ganti ROW_ID jadi dinamis.

Cara aktifkan: isi SUPABASE_URL & SUPABASE_KEY di st.secrets (lihat
README_DEPLOY.md). Kalau belum diisi, aplikasi tetap jalan normal dengan
penyimpanan sesi browser saja (seperti sebelumnya) — fitur simpan/muat
database otomatis disembunyikan.
"""
import json
from datetime import date, datetime

import pandas as pd
import streamlit as st

try:
    from supabase import create_client
except ImportError:  # library belum terpasang - tetap jangan crash
    create_client = None

TABLE = "arkas_state"
ROW_ID = 1

# Key session_state yang ikut disimpan/dimuat dari database.
STATE_KEYS = [
    "sekolah", "npsn", "status_sekolah", "kepsek_nama", "bendahara_nama",
    "tahun_anggaran", "wilayah_3t",
    "rincian_rkas", "checklist_tahapan",
    "spj_desa", "spj_kecamatan", "spj_kabupaten", "spj_provinsi",
    "spj_kepsek_nip", "spj_bendahara_nip",
    "spj_saldo_awal_kas", "spj_saldo_awal_bank",
    "spj_pagu_komponen", "spj_bku",
]

# Key yang isinya DataFrame (perlu perlakuan khusus saat simpan/muat).
DATAFRAME_KEYS = {"rincian_rkas", "checklist_tahapan", "spj_bku"}

# Kolom bertipe tanggal di tiap DataFrame di atas, supaya saat dimuat lagi
# dari JSON, tipenya dikembalikan ke datetime.date (bukan string biasa).
DATE_COLUMNS = {
    "checklist_tahapan": ["Target Tanggal"],
    "spj_bku": ["Tanggal"],
}


@st.cache_resource(show_spinner=False)
def _get_client_cached(url: str, key: str):
    """Bikin koneksi Supabase (di-cache berdasar url+key, supaya kalau
    secrets diganti, cache lama otomatis tidak dipakai lagi)."""
    return create_client(url, key)


def _get_client():
    if create_client is None:
        st.session_state["_db_connect_error"] = (
            "Library 'supabase' belum ter-install. Pastikan nama file dependency "
            "di repo adalah requirements.txt dan berisi baris 'supabase>=2.7', "
            "lalu cek tab Manage app > Logs untuk pesan instalasi."
        )
        return None
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        st.session_state["_db_connect_error"] = (
            "SUPABASE_URL / SUPABASE_KEY tidak ditemukan di Secrets. Cek App settings > Secrets."
        )
        return None
    if not url or not key:
        st.session_state["_db_connect_error"] = "SUPABASE_URL atau SUPABASE_KEY kosong di Secrets."
        return None
    try:
        client = _get_client_cached(url, key)
        st.session_state.pop("_db_connect_error", None)
        return client
    except Exception as e:
        st.session_state["_db_connect_error"] = f"Gagal membuat koneksi Supabase: {e}"
        print(f"[db_supabase] Gagal membuat koneksi Supabase: {e!r}")
        return None


def is_configured() -> bool:
    """True kalau kredensial Supabase sudah diisi & library terpasang."""
    return _get_client() is not None


def _json_default(o):
    if isinstance(o, (date, datetime, pd.Timestamp)):
        return o.isoformat()
    try:
        if pd.isna(o):
            return None
    except (TypeError, ValueError):
        pass
    return str(o)


def _serialize_state(keys=None) -> dict:
    keys = keys or STATE_KEYS
    payload = {}
    for k in keys:
        if k not in st.session_state:
            continue
        v = st.session_state[k]
        if isinstance(v, pd.DataFrame):
            payload[k] = json.loads(v.to_json(orient="records", date_format="iso"))
        elif v is None:
            payload[k] = None
        else:
            payload[k] = json.loads(json.dumps(v, default=_json_default))
    return payload


def _deserialize_state(payload: dict) -> dict:
    out = {}
    for k, v in payload.items():
        if k in DATAFRAME_KEYS:
            if not v:
                out[k] = None
                continue
            df = pd.DataFrame(v)
            for col in DATE_COLUMNS.get(k, []):
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
            out[k] = df
        else:
            out[k] = v
    return out


def load_state():
    """Ambil data tersimpan dari Supabase. Balikin dict {key: value} atau
    None kalau belum ada data / Supabase belum dikonfigurasi / gagal konek."""
    client = _get_client()
    if client is None:
        return None
    try:
        res = client.table(TABLE).select("data").eq("id", ROW_ID).limit(1).execute()
    except Exception as e:
        st.session_state["_db_error"] = f"Gagal memuat data: {e}"
        return None
    rows = res.data or []
    if not rows or not rows[0].get("data"):
        return None
    try:
        return _deserialize_state(rows[0]["data"])
    except Exception as e:
        st.session_state["_db_error"] = f"Gagal membaca data tersimpan: {e}"
        return None


def save_state() -> bool:
    """Simpan seluruh state relevan ke Supabase (1 baris, ditimpa/upsert)."""
    client = _get_client()
    if client is None:
        return False
    payload = _serialize_state()
    try:
        client.table(TABLE).upsert({
            "id": ROW_ID,
            "data": payload,
            "updated_at": datetime.now().isoformat(),
        }).execute()
        st.session_state.pop("_db_error", None)
        return True
    except Exception as e:
        st.session_state["_db_error"] = f"Gagal menyimpan data: {e}"
        return False
