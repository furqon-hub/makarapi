"""
ext_checks.py
=============
Pemeriksaan TAMBAHAN di luar margin/font/spasi: judul bab & subbab, nomor halaman,
bagian wajib (daftar isi, daftar gambar/tabel, ...), keterangan gambar/tabel,
dan daftar pustaka. Semua dikendalikan blok `tambahan` pada aturan:
nilai null / false / kosong = "tidak diatur" -> tidak diperiksa, tidak diubah.

Tiap temuan punya `perbaikan`: "otomatis" (bisa dibereskan tombol Perbaiki) atau
"manual" (hanya ditandai; pengguna perlu mengerjakannya sendiri).
"""

import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm

import parser as P

TAMBAHAN_DEFAULT = {
    "judul_bab": {"kapital": None, "tebal": None, "rata_tengah": None},
    "judul_subbab": {"tebal": None},
    "nomor_halaman": {"posisi": None, "sampul_tanpa_nomor": None},
    "bagian_wajib": [],
    "keterangan": {"tabel_posisi": None, "gambar_posisi": None},
    "daftar_pustaka": {"hanging_cm": None, "urut_abjad": None},
}

BAGIAN = {
    "kata_pengantar": ("Kata Pengantar", re.compile(r"^KATA\s+PENGANTAR\s*$", re.I)),
    "abstrak": ("Abstrak", re.compile(r"^(ABSTRAK|ABSTRACT)\s*$", re.I)),
    "daftar_isi": ("Daftar Isi", re.compile(r"^DAFTAR\s+ISI\s*$", re.I)),
    "daftar_tabel": ("Daftar Tabel", re.compile(r"^DAFTAR\s+TABEL\s*$", re.I)),
    "daftar_gambar": ("Daftar Gambar", re.compile(r"^DAFTAR\s+GAMBAR\s*$", re.I)),
    "daftar_pustaka": ("Daftar Pustaka", P.DAFTAR_PUSTAKA_PATTERN),
    "lampiran": ("Lampiran", re.compile(r"^(DAFTAR\s+)?LAMPIRAN\b", re.I)),
}

POSISI_HALAMAN = {"bawah_tengah": ("footer", "center"), "bawah_kanan": ("footer", "right"),
                  "atas_kanan": ("header", "right"), "atas_tengah": ("header", "center")}
_RATA = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}

KAP_TABEL = re.compile(r"^Tabel\s+\d", re.I)
KAP_GAMBAR = re.compile(r"^(Gambar|Grafik|Bagan|Diagram)\s+\d", re.I)


def tambahan_aktif(rules):
    t = rules.get("tambahan") or {}
    return t


def _on(t, *kunci):
    x = t
    for k in kunci:
        if not isinstance(x, dict):
            return None
        x = x.get(k)
    return x


# ---------- utilitas dokumen ----------

def _urutan_badan(doc):
    """Daftar elemen badan berurutan: dict(jenis=p|tbl, idx=indeks doc.paragraphs, ...)."""
    hasil = []
    i = 0
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            p = doc.paragraphs[i]
            gambar = bool(el.xpath(".//w:drawing | .//w:pict"))
            hasil.append({"jenis": "p", "idx": i, "p": p, "gambar": gambar, "teks": p.text.strip()})
            i += 1
        elif el.tag == qn("w:tbl"):
            hasil.append({"jenis": "tbl"})
    return hasil


def _section_per_paragraf(doc):
    """Peta indeks paragraf -> indeks section (sectPr di pPr menutup section)."""
    peta, sec = [], 0
    for p in doc.paragraphs:
        peta.append(sec)
        if p._p.pPr is not None and p._p.pPr.find(qn("w:sectPr")) is not None:
            sec += 1
    return peta


def _punya_page(parent_el):
    xml = parent_el.xml
    return bool(re.search(r"<w:instrText[^>]*>\s*PAGE\b", xml) or re.search(r'w:instr="\s*PAGE\b', xml))


def _efektif_hf(section, jenis):
    """Header/footer efektif (menelusuri 'sama dengan sebelumnya')."""
    hf = getattr(section, jenis)
    return hf


def _halaman_info(doc, sec_idx, jenis):
    """(ada_PAGE, rata) untuk header/footer efektif section sec_idx."""
    sec = doc.sections[sec_idx]
    hf = getattr(sec, jenis)
    asal = hf
    k = sec_idx
    while asal.is_linked_to_previous and k > 0:
        k -= 1
        asal = getattr(doc.sections[k], jenis)
    for p in asal.paragraphs:
        if _punya_page(p._p):
            return True, p.alignment, asal
    return False, None, asal


def _judul_bab_tingkat(zona, p):
    return zona


# ---------- pemeriksaan ----------

def jalankan(doc, paragraphs_info, rules):
    """Kembalikan (violations, total_checks_tambahan)."""
    t = tambahan_aktif(rules)
    if not t:
        return [], 0
    zona = P.classify_zones_untuk_italic(doc)
    v, total = [], 0
    badan = _urutan_badan(doc)
    sec_map = _section_per_paragraf(doc)

    def tambah(kat, lokasi, ditemukan, seharusnya, perbaikan, severity="deterministik"):
        v.append({"kategori": kat, "lokasi": lokasi, "ditemukan": ditemukan,
                  "seharusnya": seharusnya, "severity": severity, "perbaikan": perbaikan})

    # -- judul bab & subbab --
    jb = t.get("judul_bab") or {}
    js = t.get("judul_subbab") or {}
    for i, p in enumerate(doc.paragraphs):
        teks = p.text.strip()
        if zona[i] == "judul_bab":
            lok = f"Judul bab: \"{teks[:50]}\""
            if jb.get("kapital"):
                total += 1
                if teks != teks.upper():
                    tambah("judul_bab_kapital", lok, "tidak huruf kapital", "HURUF KAPITAL", "otomatis")
            if jb.get("tebal"):
                total += 1
                if not P._semua_tebal(p):
                    tambah("judul_bab_tebal", lok, "tidak tebal", "cetak tebal", "otomatis")
            if jb.get("rata_tengah"):
                total += 1
                if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    tambah("judul_bab_rata", lok, "tidak rata tengah", "rata tengah", "otomatis")
        elif zona[i] == "subjudul" and _adalah_subbab(p, teks) and js.get("tebal"):
            total += 1
            if not P._semua_tebal(p):
                tambah("judul_subbab_tebal", f"Subbab: \"{teks[:50]}\"", "tidak tebal", "cetak tebal", "otomatis")

    # -- nomor halaman --
    nh = t.get("nomor_halaman") or {}
    posisi = nh.get("posisi")
    if posisi in POSISI_HALAMAN or nh.get("sampul_tanpa_nomor"):
        jenis, rata = POSISI_HALAMAN.get(posisi, ("footer", None))
        sec_bab = sorted({sec_map[i] for i, z in enumerate(zona) if z == "judul_bab"})
        if posisi in POSISI_HALAMAN:
            for s in sec_bab:
                total += 1
                ada, rata_ada, _ = _halaman_info(doc, s, jenis)
                lok = f"Section {s} (memuat bab)"
                if not ada:
                    tambah("nomor_halaman", lok, "tidak ada nomor halaman", f"nomor di {jenis}, {rata}", "otomatis")
                else:
                    total += 1
                    cocok = (rata_ada == _RATA[rata]) if rata_ada is not None else False
                    if not cocok:
                        tambah("nomor_halaman_posisi", lok, "posisi nomor tidak sesuai", f"{jenis}, {rata}", "otomatis")
        if nh.get("sampul_tanpa_nomor") and len(doc.sections) > 1:
            total += 1
            ada, _, _ = _halaman_info(doc, 0, jenis)
            if ada:
                tambah("nomor_halaman_sampul", "Section 0 (sampul)", "sampul bernomor", "sampul tanpa nomor", "manual")

    # -- bagian wajib --
    judul_pendek = [(e["teks"]) for e in badan if e["jenis"] == "p" and 0 < len(e["teks"]) <= 60]
    for kunci in (t.get("bagian_wajib") or []):
        if kunci not in BAGIAN:
            continue
        nama, pola = BAGIAN[kunci]
        total += 1
        if not any(pola.match(x) for x in judul_pendek):
            tambah("bagian_wajib", "Dokumen", f"\"{nama}\" tidak ditemukan", f"ada bagian {nama}", "manual")

    # -- keterangan tabel/gambar (hanya setelah BAB pertama) --
    ket = t.get("keterangan") or {}
    mulai = next((k for k, e in enumerate(badan) if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    if mulai is not None and (ket.get("tabel_posisi") or ket.get("gambar_posisi")):
        for k in range(mulai, len(badan)):
            e = badan[k]
            if e["jenis"] == "tbl":
                tipe, pola, aturan = "Tabel", KAP_TABEL, ket.get("tabel_posisi")
            elif e["jenis"] == "p" and e["gambar"]:
                tipe, pola, aturan = "Gambar", KAP_GAMBAR, ket.get("gambar_posisi")
            else:
                continue
            if not aturan:
                continue
            total += 1
            atas = _teks_tetangga(badan, k, -1)
            bawah = _teks_tetangga(badan, k, +1)
            ada_atas, ada_bawah = bool(pola.match(atas)), bool(pola.match(bawah))
            if not (ada_atas or ada_bawah):
                tambah("keterangan", f"{tipe} setelah \"{(atas or bawah)[:40]}\"", f"{tipe} tanpa keterangan",
                       f"keterangan {tipe.lower()} di {aturan}", "manual")
            elif (aturan == "atas" and not ada_atas) or (aturan == "bawah" and not ada_bawah):
                tambah("keterangan_posisi", f"{tipe} dengan keterangan di sisi lain", f"keterangan di {'bawah' if ada_bawah else 'atas'}",
                       f"di {aturan}", "manual")

    # -- daftar pustaka --
    dp = t.get("daftar_pustaka") or {}
    entri = [(i, p) for i, p in enumerate(doc.paragraphs) if zona[i] == "daftar_pustaka" and p.text.strip()]
    if dp.get("hanging_cm") is not None:
        h = float(dp["hanging_cm"])
        for i, p in entri:
            total += 1
            pf = p.paragraph_format
            kiri = pf.left_indent.cm if pf.left_indent is not None else 0.0
            baris1 = pf.first_line_indent.cm if pf.first_line_indent is not None else 0.0
            if abs(kiri - h) > 0.1 or abs(baris1 + h) > 0.1:
                tambah("daftar_pustaka_hanging", f"Pustaka: \"{p.text.strip()[:40]}\"",
                       f"kiri {kiri:.1f} cm, baris pertama {baris1:.1f} cm", f"menggantung {h} cm", "otomatis")
    if dp.get("urut_abjad") and len(entri) > 1:
        total += 1
        kunci = [re.sub(r"[^a-z0-9]", "", p.text.lower()) for _, p in entri]
        for a in range(1, len(kunci)):
            if kunci[a] < kunci[a - 1]:
                tambah("daftar_pustaka_urut", f"Pustaka: \"{entri[a][1].text.strip()[:40]}\"",
                       "tidak urut abjad", "urut A-Z", "manual", severity="perlu_verifikasi")
                break
    return v, total


def _adalah_subbab(p, teks):
    return bool(re.match(r"^\d+(\.\d+)+[.)]?\s+\S", teks)) or P._nama_style(p).startswith("heading 2") \
        or P._nama_style(p).startswith("heading 3")


def _teks_tetangga(badan, k, arah):
    j = k + arah
    while 0 <= j < len(badan):
        e = badan[j]
        if e["jenis"] == "tbl":
            return ""
        if e["teks"]:
            return e["teks"]
        if e["gambar"]:
            return ""
        j += arah
    return ""


# ---------- perbaikan aman ----------

def _tambah_page_field(paragraph):
    run = paragraph.add_run()
    for tipe, teks in (("begin", None), (None, " PAGE "), ("end", None)):
        if tipe:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipe)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = teks
        run._r.append(el)


def perbaiki(doc, rules):
    t = tambahan_aktif(rules)
    if not t:
        return
    zona = P.classify_zones_untuk_italic(doc)
    sec_map = _section_per_paragraf(doc)
    jb = t.get("judul_bab") or {}
    js = t.get("judul_subbab") or {}
    for i, p in enumerate(doc.paragraphs):
        teks = p.text.strip()
        if zona[i] == "judul_bab":
            if jb.get("kapital"):
                for r in p.runs:
                    r.text = r.text.upper()
            if jb.get("tebal"):
                for r in p.runs:
                    r.bold = True
            if jb.get("rata_tengah"):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif zona[i] == "subjudul" and js.get("tebal") and _adalah_subbab(p, teks):
            for r in p.runs:
                r.bold = True

    nh = t.get("nomor_halaman") or {}
    posisi = nh.get("posisi")
    if posisi in POSISI_HALAMAN:
        jenis, rata = POSISI_HALAMAN[posisi]
        for s in sorted({sec_map[i] for i, z in enumerate(zona) if z == "judul_bab"}):
            ada, _, asal = _halaman_info(doc, s, jenis)
            if not ada:
                hf = getattr(doc.sections[s], jenis)
                hf.is_linked_to_previous = False
                par = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
                _tambah_page_field(par)
                par.alignment = _RATA[rata]
            else:
                for par in asal.paragraphs:
                    if _punya_page(par._p):
                        par.alignment = _RATA[rata]

    dp = t.get("daftar_pustaka") or {}
    if dp.get("hanging_cm") is not None:
        h = float(dp["hanging_cm"])
        for i, p in enumerate(doc.paragraphs):
            if zona[i] == "daftar_pustaka" and p.text.strip():
                p.paragraph_format.left_indent = Cm(h)
                p.paragraph_format.first_line_indent = Cm(-h)
