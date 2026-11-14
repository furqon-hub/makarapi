"""kustom.py - ubah isian daftar centang (JSON) menjadi aturan. Kosong/mati = tidak diperiksa."""
import copy
import json

from fastapi import HTTPException

BATAS = {
    "atas_cm": (0.5, 8.0), "bawah_cm": (0.5, 8.0), "kiri_cm": (0.5, 8.0), "kanan_cm": (0.5, 8.0),
    "font_ukuran_pt": (6.0, 24.0), "spasi_isi": (0.5, 3.0), "indentasi_cm": (0.0, 5.0),
    "hanging_cm": (0.3, 3.0),
}
KERTAS = {"A4": (21.0, 29.7), "Letter": (21.59, 27.94), "F4": (21.5, 33.0), "A5": (14.8, 21.0)}
PERATAAN = {"left", "right", "center", "justify"}
POSISI_NOMOR = {"bawah_tengah", "bawah_kanan", "atas_kanan", "atas_tengah"}
BAGIAN = {"kata_pengantar", "abstrak", "daftar_isi", "daftar_tabel", "daftar_gambar", "daftar_pustaka", "lampiran"}
POSISI_KET = {"atas", "bawah"}


def _angka(d, kunci):
    v = d.get(kunci)
    if v in (None, ""):
        return None
    try:
        v = float(v)
    except (TypeError, ValueError):
        raise HTTPException(422, f"Isian {kunci} harus berupa angka.")
    lo, hi = BATAS[kunci]
    if not (lo <= v <= hi):
        raise HTTPException(422, f"Isian {kunci} harus antara {lo} dan {hi}.")
    return v


def _bool(d, kunci):
    return True if d.get(kunci) is True else None


def aturan_dari_json(dasar, teks_json):
    try:
        j = json.loads(teks_json)
        if not isinstance(j, dict):
            raise ValueError
    except ValueError:
        raise HTTPException(422, "Isian kustom tidak valid.")
    a = copy.deepcopy(dasar)
    # semua aspek dimatikan dulu: daftar centang, yang tidak diisi tidak diperiksa
    a["ukuran_kertas"]["lebar_cm"] = a["ukuran_kertas"]["tinggi_cm"] = None
    for s in ("atas", "bawah", "kiri", "kanan"):
        a["margin"][f"{s}_cm"] = None
    a["font"]["nama"] = a["font"]["ukuran_isi_pt"] = None
    a["spasi"]["baris_isi"] = None
    a["indentasi"]["awal_paragraf_cm"] = None
    a["perataan"]["jenis"] = None

    h = j.get("halaman") or {}
    if h.get("kertas"):
        if h["kertas"] not in KERTAS:
            raise HTTPException(422, "Ukuran kertas tidak dikenal.")
        a["ukuran_kertas"]["lebar_cm"], a["ukuran_kertas"]["tinggi_cm"] = KERTAS[h["kertas"]]
    for s in ("atas", "bawah", "kiri", "kanan"):
        a["margin"][f"{s}_cm"] = _angka(h, f"{s}_cm")

    t = j.get("teks") or {}
    nama = (t.get("font_nama") or "").strip()
    a["font"]["nama"] = nama[:60] or None
    a["font"]["ukuran_isi_pt"] = _angka(t, "font_ukuran_pt")
    a["spasi"]["baris_isi"] = _angka(t, "spasi_isi")
    a["indentasi"]["awal_paragraf_cm"] = _angka(t, "indentasi_cm")
    if t.get("perataan"):
        if t["perataan"] not in PERATAAN:
            raise HTTPException(422, "Perataan tidak dikenal.")
        a["perataan"]["jenis"] = t["perataan"]

    jb = j.get("judul_bab") or {}
    js = j.get("judul_subbab") or {}
    nh = j.get("nomor_halaman") or {}
    ket = j.get("keterangan") or {}
    dp = j.get("daftar_pustaka") or {}
    posisi = nh.get("posisi") or None
    if posisi and posisi not in POSISI_NOMOR:
        raise HTTPException(422, "Posisi nomor halaman tidak dikenal.")
    for k in ("tabel_posisi", "gambar_posisi"):
        if ket.get(k) and ket[k] not in POSISI_KET:
            raise HTTPException(422, "Posisi keterangan harus atas atau bawah.")
    bw = [x for x in (j.get("bagian_wajib") or []) if x in BAGIAN]
    a["tambahan"] = {
        "judul_bab": {"kapital": _bool(jb, "kapital"), "tebal": _bool(jb, "tebal"), "rata_tengah": _bool(jb, "rata_tengah")},
        "judul_subbab": {"tebal": _bool(js, "tebal")},
        "nomor_halaman": {"posisi": posisi, "sampul_tanpa_nomor": _bool(nh, "sampul_tanpa_nomor")},
        "bagian_wajib": bw,
        "keterangan": {"tabel_posisi": ket.get("tabel_posisi") or None, "gambar_posisi": ket.get("gambar_posisi") or None},
        "daftar_pustaka": {"hanging_cm": _angka(dp, "hanging_cm"), "urut_abjad": _bool(dp, "urut_abjad")},
    }
    return a
