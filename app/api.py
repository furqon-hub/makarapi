"""
api.py
======
Pembungkus REST API (FastAPI) untuk mesin MakaRapi, supaya bisa dipanggil
dari aplikasi mobile (React Native / Expo).

Jalankan lokal:
    uvicorn api:app --reload --port 8000
Dokumentasi interaktif otomatis: http://localhost:8000/docs

Endpoint:
    GET  /health  -> cek server hidup
    GET  /profil  -> daftar profil pedoman (febi, kustom)
    POST /check   -> unggah .docx, terima JSON skor + pelanggaran + kandidat istilah asing
    POST /fix     -> unggah .docx, terima file .docx hasil perbaikan otomatis

Batasan yang disengaja (lihat bagian Risiko di proposal):
  - Hanya .docx, maksimal 10 MB.
  - Rate limit sederhana per IP (dalam memori, 10 permintaan/menit).
  - File tidak disimpan: dihapus segera setelah diproses.
"""

import copy
import os
import tempfile
import time
from collections import Counter, defaultdict, deque
from pathlib import Path
from typing import Optional

from docx import Document
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

import auto_fix
import italic_detector
import parser as parser_mod
import rule_engine

BASE_DIR = Path(__file__).resolve().parent.parent
RULES_PATH = str(BASE_DIR / "rules.yaml")
ISTILAH_PATH = str(BASE_DIR / "daftar_istilah_asing.yaml")

MAX_BYTES = 10 * 1024 * 1024
RATE_LIMIT = 10          # permintaan
RATE_WINDOW = 60         # detik
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(title="MakaRapi API", version="0.2.0")
_hits = defaultdict(deque)

PROFIL = [
    {
        "id": "febi",
        "nama": "FEBI UIN Gusdur 2026",
        "tervalidasi": True,
        "keterangan": "Bersumber dari Pedoman Penulisan Tugas Akhir FEBI 2026 (diadaptasi untuk makalah: spasi 1,5).",
    },
    {
        "id": "kustom",
        "nama": "Kustom (isi sendiri)",
        "tervalidasi": False,
        "keterangan": "Nilai diisi pengguna; hasil bergantung pada nilai yang dimasukkan.",
    },
]

# Rentang nilai yang masuk akal untuk profil kustom: (min, max)
BATAS = {
    "atas_cm": (0.5, 8.0), "bawah_cm": (0.5, 8.0), "kiri_cm": (0.5, 8.0), "kanan_cm": (0.5, 8.0),
    "font_ukuran_pt": (6.0, 24.0), "spasi_isi": (0.5, 3.0), "indentasi_cm": (0.0, 5.0),
}
PERATAAN_VALID = {"left", "right", "center", "justify"}


def _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama, font_ukuran_pt,
                  spasi_isi, indentasi_cm, perataan):
    """Kembalikan dict aturan untuk profil terpilih. Profil kustom = profil FEBI
    yang nilainya ditimpa dengan isian pengguna (yang kosong memakai nilai FEBI)."""
    dasar = rule_engine.load_rules(RULES_PATH)
    if profil == "febi":
        return dasar
    if profil != "kustom":
        raise HTTPException(status_code=422, detail=f"Profil '{profil}' tidak dikenal.")

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


@app.get("/health")
def health():
    return {"status": "ok", "versi": app.version}


@app.get("/profil")
def daftar_profil():
    return PROFIL


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
):
    rules = _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama,
                          font_ukuran_pt, spasi_isi, indentasi_cm, perataan)
    path = await _simpan_upload(file)
    try:
        try:
            doc = Document(path)
        except Exception:
            raise HTTPException(status_code=422, detail="File .docx tidak dapat dibaca (rusak atau bukan dokumen Word).")

        hasil = rule_engine.run_all_checks(
            parser_mod.get_page_setup(doc), parser_mod.get_paragraphs_info(doc), rules
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
):
    rules = _aturan_untuk(profil, atas_cm, bawah_cm, kiri_cm, kanan_cm, font_nama,
                          font_ukuran_pt, spasi_isi, indentasi_cm, perataan)
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
