"""
api.py
======
Pembungkus REST API (FastAPI) untuk mesin MakaRapi, supaya bisa dipanggil
dari halaman web (app/web/index.html) yang juga dibungkus WebView menjadi APK.

Jalankan lokal:
    uvicorn api:app --reload --port 8000
Dokumentasi interaktif otomatis: http://localhost:8000/docs

Endpoint:
    GET  /        -> halaman web MakaRapi
    GET  /health  -> cek server hidup
    GET  /profil  -> daftar profil pedoman (bawaan FEBI, berkas profil/*.yaml per kampus, kustom)
    POST /check   -> unggah .docx, terima JSON skor + pelanggaran + kandidat istilah asing
    POST /fix     -> unggah .docx, terima file .docx hasil perbaikan otomatis

Batasan yang disengaja (lihat bagian Risiko di proposal):
  - Hanya .docx, maksimal 10 MB.
  - Rate limit sederhana per IP (dalam memori, 10 permintaan/menit).
  - File tidak disimpan: dihapus segera setelah diproses.
"""

import copy
import os
import re
import tempfile
import time
from collections import Counter, defaultdict, deque
from pathlib import Path
from typing import Optional

import yaml
from docx import Document
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

import auto_fix
import italic_detector
import kustom
import parser as parser_mod
import rule_engine

BASE_DIR = Path(__file__).resolve().parent.parent
RULES_PATH = str(BASE_DIR / "rules.yaml")
ISTILAH_PATH = str(BASE_DIR / "daftar_istilah_asing.yaml")
PROFIL_DIR = BASE_DIR / "profil"

MAX_BYTES = 10 * 1024 * 1024
RATE_LIMIT = 10          # permintaan
RATE_WINDOW = 60         # detik
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(title="MakaRapi API", version="0.5.0")
_hits = defaultdict(deque)

PROFIL = [
    {
        "id": "febi",
        "kampus": "FEBI UIN K.H. Abdurrahman Wahid Pekalongan",
        "nama": "FEBI - Makalah",
        "tervalidasi": True,
        "keterangan": "Bersumber dari Pedoman Penulisan Tugas Akhir FEBI 2026, diadaptasi untuk makalah: spasi 1,5.",
    },
    {
        "id": "febi_skripsi",
        "kampus": "FEBI UIN K.H. Abdurrahman Wahid Pekalongan",
        "nama": "FEBI - Skripsi / Sempro",
        "tervalidasi": True,
        "keterangan": "Nilai apa adanya dari Pedoman Penulisan Tugas Akhir FEBI 2026 (spasi 2,0). Untuk skripsi, proposal, dan sempro.",
    },
    {
        "id": "kustom",
        "kampus": "Lainnya",
        "nama": "Kustom (isi sendiri)",
        "tervalidasi": False,
        "keterangan": "Nilai diisi pengguna; hasil bergantung pada nilai yang dimasukkan.",
    },
]

_ID_PROFIL = re.compile(r"^[a-z0-9_]{2,40}$")
_ID_BAWAAN = {p["id"] for p in PROFIL}


def _gabung(dasar, timpa):
    """Timpa dict `dasar` dengan `timpa` secara rekursif. null (None) di `timpa`
    ikut menimpa: artinya aspek itu tidak diatur pedoman dan dilewati."""
    for k, v in timpa.items():
        if isinstance(v, dict) and isinstance(dasar.get(k), dict):
            _gabung(dasar[k], v)
        else:
            dasar[k] = v


def _profil_berkas():
    """Baca profil per kampus dari folder profil/*.yaml. Berkas yang rusak atau
    tidak sesuai bentuk dilewati (tidak menjatuhkan server)."""
    hasil = {}
    if not PROFIL_DIR.is_dir():
        return hasil
    for f in sorted(PROFIL_DIR.glob("*.yaml")):
        try:
            d = yaml.safe_load(f.read_text(encoding="utf-8"))
            pid = d["id"]
            if (not _ID_PROFIL.match(pid) or pid in _ID_BAWAAN or pid in hasil
                    or not isinstance(d.get("aturan"), dict)):
                continue
            hasil[pid] = d
        except Exception:
            continue
    return hasil


def _info_profil(d):
    """Bentuk ringkas profil berkas untuk GET /profil."""
    catatan = d.get("catatan") or ""
    return {
        "id": d["id"],
        "nama": d.get("nama", d["id"]),
        "kampus": d.get("kampus", "Lainnya"),
        "jenis": d.get("jenis"),
        "tahun": d.get("tahun"),
        "tervalidasi": bool(d.get("tervalidasi", False)),
        "keterangan": ("Dari pedoman publik. " + catatan).strip(),
        "sumber_judul": d.get("sumber_judul"),
        "sumber_url": d.get("sumber_url"),
        "lokasi_nilai": d.get("lokasi_nilai"),
    }


# Rentang nilai yang masuk akal untuk profil kustom: (min, max)
BATAS = {
    "atas_cm": (0.5, 8.0), "bawah_cm": (0.5, 8.0), "kiri_cm": (0.5, 8.0), "kanan_cm": (0.5, 8.0),
    "font_ukuran_pt": (6.0, 24.0), "spasi_isi": (0.5, 3.0), "indentasi_cm": (0.0, 5.0),
}
PERATAAN_VALID = {"left", "right", "center", "justify"}


def _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama, font_ukuran_pt,
                  spasi_isi, indentasi_cm, perataan, kustom_json=None):
    """Kembalikan dict aturan untuk profil terpilih. Profil kustom = profil FEBI
    yang nilainya ditimpa dengan isian pengguna (yang kosong memakai nilai FEBI)."""
    dasar = rule_engine.load_rules(RULES_PATH)
    if profil == "febi":
        return dasar
    if profil == "febi_skripsi":
        aturan = copy.deepcopy(dasar)
        aturan["spasi"]["baris_isi"] = dasar["spasi"]["baris_isi_skripsi_asli"]
        return aturan
    berkas = _profil_berkas()
    if profil in berkas:
        aturan = copy.deepcopy(dasar)
        _gabung(aturan, berkas[profil]["aturan"])
        return aturan
    if profil != "kustom":
        raise HTTPException(status_code=422, detail=f"Profil '{profil}' tidak dikenal.")

    if kustom_json:
        return kustom.aturan_dari_json(dasar, kustom_json)
    aturan = copy.deepcopy(dasar)
    isian = {
        "atas_cm": atas_cm, "bawah_cm": bawah_cm, "kiri_cm": kiri_cm, "kanan_cm": kanan_cm,
        "font_ukuran_pt": font_ukuran_pt, "spasi_isi": spasi_isi, "indentasi_cm": indentasi_cm,
    }
    for nama, nilai in isian.items():
        if nilai is None:
            continue
        lo, hi = BATAS[nama]
        if not (lo <= nilai <= hi):
            raise HTTPException(status_code=422, detail=f"Nilai {nama} harus antara {lo} dan {hi}.")
    for sisi in ("atas", "bawah", "kiri", "kanan"):
        if isian[f"{sisi}_cm"] is not None:
            aturan["margin"][f"{sisi}_cm"] = isian[f"{sisi}_cm"]
    if font_nama:
        aturan["font"]["nama"] = font_nama.strip()[:60]
    if font_ukuran_pt is not None:
        aturan["font"]["ukuran_isi_pt"] = font_ukuran_pt
    if spasi_isi is not None:
        aturan["spasi"]["baris_isi"] = spasi_isi
    if indentasi_cm is not None:
        aturan["indentasi"]["awal_paragraf_cm"] = indentasi_cm
    if perataan:
        if perataan not in PERATAAN_VALID:
            raise HTTPException(status_code=422, detail="Perataan harus left, right, center, atau justify.")
        aturan["perataan"]["jenis"] = perataan
    return aturan


@app.middleware("http")
async def rate_limiter(request: Request, call_next):
    if request.url.path in ("/check", "/fix"):
        ip = request.client.host if request.client else "unknown"
        now = time.time()
        q = _hits[ip]
        while q and now - q[0] > RATE_WINDOW:
            q.popleft()
        if len(q) >= RATE_LIMIT:
            return JSONResponse(
                status_code=429,
                content={"detail": f"Terlalu banyak permintaan (maks {RATE_LIMIT}/menit). Coba lagi sebentar."},
            )
        q.append(now)
    return await call_next(request)


# Ditambahkan SETELAH rate limiter supaya CORS menjadi lapisan terluar
# (jawaban 429 pun tetap bisa dibaca oleh browser).
# CORS: diperlukan agar aplikasi versi web (browser) boleh memanggil API ini.
# Aplikasi Android/iOS tidak membutuhkannya. API tidak memakai cookie/login,
# jadi "*" aman untuk MVP; persempit ke domain aplikasi saat rilis.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


async def _simpan_upload(file: UploadFile) -> str:
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Hanya file .docx yang didukung.")
    data = await file.read()
    if len(data) > MAX_BYTES:
        raise HTTPException(status_code=413, detail="Ukuran file melebihi 10 MB.")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.write(data)
    tmp.close()
    return tmp.name


WEB_DIR = Path(__file__).resolve().parent / "web"


@app.get("/", include_in_schema=False)
def beranda():
    """Halaman web MakaRapi (dipakai browser komputer dan cangkang WebView APK)."""
    return FileResponse(WEB_DIR / "index.html", media_type="text/html; charset=utf-8")


@app.get("/health")
def health():
    return {"status": "ok", "versi": app.version}


@app.get("/profil")
def daftar_profil():
    urut = [p for p in PROFIL if p["id"] != "kustom"]
    urut += [_info_profil(d) for d in sorted(_profil_berkas().values(), key=lambda x: (x.get("kampus", ""), x["id"]))]
    urut += [p for p in PROFIL if p["id"] == "kustom"]
    return urut


@app.post("/check")
async def check(
    file: UploadFile = File(...),
    profil: str = Form("febi"),
    atas_cm: Optional[float] = Form(None),
    bawah_cm: Optional[float] = Form(None),
    kiri_cm: Optional[float] = Form(None),
    kanan_cm: Optional[float] = Form(None),
    font_nama: Optional[str] = Form(None),
    font_ukuran_pt: Optional[float] = Form(None),
    spasi_isi: Optional[float] = Form(None),
    indentasi_cm: Optional[float] = Form(None),
    perataan: Optional[str] = Form(None),
    kustom_json: Optional[str] = Form(None),
):
    rules = _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama,
                          font_ukuran_pt, spasi_isi, indentasi_cm, perataan, kustom_json)
    path = await _simpan_upload(file)
    try:
        try:
            doc = Document(path)
        except Exception:
            raise HTTPException(status_code=422, detail="File .docx tidak dapat dibaca (rusak atau bukan dokumen Word).")

        hasil = rule_engine.run_all_checks(
            parser_mod.get_page_setup(doc), parser_mod.get_paragraphs_info(doc), rules, doc=doc
        )
        kandidat = italic_detector.find_italic_candidates(doc, ISTILAH_PATH)
        return {
            "nama_file": file.filename,
            "profil": profil,
            "skor_kepatuhan_persen": hasil["skor_kepatuhan_persen"],
            "jumlah_pelanggaran": hasil["jumlah_pelanggaran"],
            "total_pemeriksaan": hasil.get("total_checks"),
            "ringkasan_per_kategori": dict(Counter(v["kategori"] for v in hasil["violations"])),
            "pelanggaran": hasil["violations"],
            "kandidat_istilah_asing": [
                {"paragraf_index": k["paragraf_index"], "kata_kandidat": k["kata_kandidat"], "konteks": k["konteks"]}
                for k in kandidat
            ],
        }
    finally:
        os.unlink(path)


@app.post("/fix")
async def fix(
    file: UploadFile = File(...),
    profil: str = Form("febi"),
    atas_cm: Optional[float] = Form(None),
    bawah_cm: Optional[float] = Form(None),
    kiri_cm: Optional[float] = Form(None),
    kanan_cm: Optional[float] = Form(None),
    font_nama: Optional[str] = Form(None),
    font_ukuran_pt: Optional[float] = Form(None),
    spasi_isi: Optional[float] = Form(None),
    indentasi_cm: Optional[float] = Form(None),
    perataan: Optional[str] = Form(None),
    kustom_json: Optional[str] = Form(None),
):
    rules = _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama,
                          font_ukuran_pt, spasi_isi, indentasi_cm, perataan, kustom_json)
    path = await _simpan_upload(file)
    out_path = path.replace(".docx", "_fixed.docx")
    try:
        auto_fix.auto_fix_document(path, out_path, rules)
    except Exception:
        for p in (path, out_path):
            if os.path.exists(p):
                os.unlink(p)
        raise HTTPException(status_code=422, detail="Dokumen tidak dapat diperbaiki otomatis.")
    os.unlink(path)

    def _bersihkan():
        if os.path.exists(out_path):
            os.unlink(out_path)

    return FileResponse(
        out_path,
        media_type=DOCX_MIME,
        filename=f"fixed_{file.filename}",
        background=BackgroundTask(_bersihkan),
    )


@app.get("/.well-known/assetlinks.json", include_in_schema=False)
def assetlinks():
    """Verifikasi kepemilikan situs untuk APK (Digital Asset Links) agar bilah browser tidak tampil."""
    return FileResponse(WEB_DIR / "assetlinks.json", media_type="application/json")


_STATIS = {
    "manifest.webmanifest": "application/manifest+json",
    "sw.js": "application/javascript",
    "icon-192.png": "image/png",
    "icon-512.png": "image/png",
    "icon-maskable-512.png": "image/png",
}


@app.get("/{nama}", include_in_schema=False)
def berkas_statis(nama: str):
    """Berkas PWA (manifest, service worker, ikon) - daftar putih, bukan folder bebas."""
    if nama not in _STATIS:
        raise HTTPException(status_code=404, detail="Tidak ditemukan.")
    return FileResponse(WEB_DIR / nama, media_type=_STATIS[nama])
