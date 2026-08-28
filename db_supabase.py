"""
Lapisan penyimpanan permanen (Supabase) untuk Pendamping ARKAS.
"""
import json
from datetime import date, datetime

import pandas as pd
import streamlit as st

try:
    from supabase import create_client
except ImportError:
    create_client = None

TABLE = "arkas_state"
ROW_ID = 1

STATE_KEYS = [
    "sekolah", "npsn", "status_sekolah", "kepsek_nama", "bendahara_nama",
    "tahun_anggaran", "wilayah_3t",
    "rincian_rkas", "checklist_tahapan",
    "spj_desa", "spj_kecamatan", "spj_kabupaten", "spj_provinsi",
    "spj_kepsek_nip", "spj_bendahara_nip",
    "spj_saldo_awal_kas", "spj_saldo_awal_bank",
    "spj_pagu_komponen", "spj_bku",
]

DATAFRAME_KEYS = {"rincian_rkas", "checklist_tahapan", "spj_bku"}

DATE_COLUMNS = {
    "checklist_tahapan": ["Target Tanggal"],
    "spj_bku": ["Tanggal"],
}


@st.cache_resource(show_spinner=False)
def _get_client():
    """Bikin koneksi Supabase sekali per proses server."""
    if create_client is None:
        return None
    try:
        url = str(st.secrets.get("SUPABASE_URL", "")).strip()
        key = str(st.secrets.get("SUPABASE_KEY", "")).strip()

        # Bersihkan URL otomatis jika ada akhiran /rest/v1/ atau /
        if url.endswith("/rest/v1/"):
            url = url[:-9]
        elif url.endswith("/rest/v1"):
            url = url[:-8]
        if url.endswith("/"):
            url = url[:-1]

        if not url or not key:
            return None

        return create_client(url, key)
    except Exception as e:
        st.session_state["_db_error"] = f"Kesalahan Konfigurasi Supabase: {e}"
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
    """Ambil data tersimpan dari Supabase."""
    client = _get_client()
    if client is None:
        return None
    try:
        res = client.table(TABLE).select("data").eq("id", ROW_ID).limit(1).execute()
        rows = res.data or []
        if not rows or not rows[0].get("data"):
            return None
        return _deserialize_state(rows[0]["data"])
    except Exception as e:
        st.session_state["_db_error"] = f"Gagal memuat data dari Supabase: {e}"
        return None


def save_state() -> bool:
    """Simpan seluruh state relevan ke Supabase."""
    client = _get_client()
    if client is None:
        return False
    payload = _serialize_state()
    try:
        client.table(TABLE).upsert({
            "id": ROW_ID,
            "data": payload,
        }).execute()
        st.session_state.pop("_db_error", None)
        return True
    except Exception as e:
        st.session_state["_db_error"] = f"Gagal menyimpan data ke Supabase: {e}"
        return False
