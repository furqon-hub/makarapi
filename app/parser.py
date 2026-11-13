"""
parser.py
=========
Modul untuk membaca properti format dari file .docx.

KONSEP PENTING YANG DIPELAJARI DI SINI:
----------------------------------------
File .docx sebenarnya adalah berkas ZIP yang isinya XML. python-docx
membungkus XML itu jadi objek Python yang lebih mudah dipakai
(Document, Paragraph, Run, dst).

Masalah utama yang akan sering Anda temui:
Microsoft Word punya mekanisme PEWARISAN (inheritance) untuk properti
seperti font dan ukuran huruf:

    Run (karakter yang diketik)
      -> mewarisi dari Paragraph (kalau run tidak set properti sendiri)
        -> mewarisi dari Style paragraf (misal "Normal")
          -> mewarisi dari Default Style dokumen (docDefaults)

Jadi kalau Anda tulis:
    run.font.size

...dan hasilnya None, itu BUKAN berarti paragraf tidak punya ukuran
font. Artinya ukurannya diwarisi dari style, dan Anda harus menelusuri
rantai pewarisan itu untuk tahu nilai sebenarnya ("effective value").

Mengabaikan hal ini adalah penyebab #1 "false positive" (sistem
bilang salah format, padahal sebenarnya benar) pada sistem pemeriksa
format semacam ini -- ini juga yang disinggung di diskusi kita
sebelumnya soal mahasiswa yang mengetik "BAB I" lalu menebalkannya
manual, bukan lewat Heading Style.
"""

from docx import Document
from docx.shared import Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH


EMU_PER_CM = 360000  # 1 cm = 360000 EMU (satuan internal Word)


def emu_to_cm(value):
    """Konversi satuan internal Word (EMU) ke cm. None tetap None."""
    if value is None:
        return None
    return round(value / EMU_PER_CM, 3)


def get_page_setup(doc: Document):
    """
    Ambil ukuran kertas dan margin dari SETIAP section dokumen.

    Kenapa "setiap section", bukan cuma satu?
    Dokumen .docx bisa punya lebih dari satu "section" (misal:
    halaman sampul pakai margin berbeda dari BAB isi). Kita harus
    cek semua, bukan asumsikan cuma ada satu.
    """
    hasil = []
    for i, section in enumerate(doc.sections):
        hasil.append({
            "section_index": i,
            "lebar_cm": emu_to_cm(section.page_width),
            "tinggi_cm": emu_to_cm(section.page_height),
            "margin_atas_cm": emu_to_cm(section.top_margin),
            "margin_bawah_cm": emu_to_cm(section.bottom_margin),
            "margin_kiri_cm": emu_to_cm(section.left_margin),
            "margin_kanan_cm": emu_to_cm(section.right_margin),
        })
    return hasil


def _get_theme_fonts(document):
    """
    Lapisan pewarisan TERAKHIR yang sering terlewat: kalau dokumen
    tidak pernah disentuh font-nya sama sekali (diketik langsung,
    tanpa klik menu Font), docDefaults di styles.xml biasanya cuma
    berisi REFERENSI tema (mis. asciiTheme="minorHAnsi"), bukan nama
    font literal. Nama font sebenarnya baru ketemu di word/theme/
    theme1.xml, di bagian <a:minorFont> (dipakai teks isi) atau
    <a:majorFont> (dipakai heading).

    Tanpa langkah ini, SEMUA dokumen yang ditulis apa adanya oleh
    mahasiswa -- tanpa pernah membuka menu Font -- akan selalu
    dilaporkan "tidak terdeteksi", padahal fontnya sebenarnya ada dan
    konsisten. Ini ditemukan lewat pengujian nyata, bukan dugaan.

    Dikembalikan sebagai dict {"major": nama_atau_None, "minor": nama_atau_None}.
    """
    try:
        for rel in document.part.rels.values():
            if rel.reltype.endswith("/theme"):
                xml = rel.target_part.blob.decode("utf-8")
                import re
                major_match = re.search(r"<a:majorFont>.*?<a:latin[^/]*typeface=\"([^\"]+)\"", xml, re.DOTALL)
                minor_match = re.search(r"<a:minorFont>.*?<a:latin[^/]*typeface=\"([^\"]+)\"", xml, re.DOTALL)
                return {
                    "major": major_match.group(1) if major_match else None,
                    "minor": minor_match.group(1) if minor_match else None,
                }
    except Exception:
        pass  # kalau gagal baca tema, biarkan fallback ke None -> "perlu cek manual"
    return {"major": None, "minor": None}


def _get_doc_defaults(document):
    """
    Baca <w:docDefaults> di styles.xml: font TEMA mana yang dipakai
    teks isi (ascii/asciiTheme) dan ukuran default dalam half-point.
    """
    import re
    styles_xml = document.styles.element.xml
    m = re.search(r"<w:docDefaults>.*?</w:docDefaults>", styles_xml, re.DOTALL)
    if not m:
        return {"ascii_theme": None, "size_half_pt": None}
    block = m.group(0)
    ascii_theme = re.search(r'w:asciiTheme="([^"]+)"', block)
    size = re.search(r'<w:sz w:val="(\d+)"', block)
    return {
        "ascii_theme": ascii_theme.group(1) if ascii_theme else None,
        "size_half_pt": int(size.group(1)) if size else None,
    }


def _resolve_font_name(paragraph, document=None):
    """
    Cari nama font yang SEBENARNYA dipakai paragraf ini, menyusuri
    rantai pewarisan kalau run-nya tidak set font sendiri.

    Urutan pencarian:
      1. Font di run pertama yang punya teks (paling spesifik)
      2. Font di style paragraf
      3. Font tema dokumen (docDefaults -> theme1.xml) -- lapisan
         yang sering terlewat, lihat _get_theme_fonts()
    """
    for run in paragraph.runs:
        if run.text.strip() and run.font.name:
            return run.font.name

    style = paragraph.style
    while style is not None:
        if style.font.name:
            return style.font.name
        style = style.base_style  # naik satu level ke style induk

    if document is not None:
        defaults = _get_doc_defaults(document)
        if defaults["ascii_theme"]:
            theme_fonts = _get_theme_fonts(document)
            # "minorHAnsi"/"minorAscii" dst -> pakai font minor (isi teks)
            # "majorHAnsi"/"majorAscii" dst -> pakai font major (heading)
            key = "minor" if "minor" in defaults["ascii_theme"].lower() else "major"
            if theme_fonts[key]:
                return theme_fonts[key]

    return None  # tidak ketemu sama sekali -> perlu dicek manual


def _resolve_font_size_pt(paragraph, document=None):
    """Sama seperti _resolve_font_name, tapi untuk ukuran font."""
    for run in paragraph.runs:
        if run.text.strip() and run.font.size:
            return run.font.size.pt

    style = paragraph.style
    while style is not None:
        if style.font.size:
            return style.font.size.pt
        style = style.base_style

    if document is not None:
        defaults = _get_doc_defaults(document)
        if defaults["size_half_pt"]:
            return defaults["size_half_pt"] / 2  # half-point -> point

    return None


def _resolve_line_spacing(paragraph):
    """
    Ambil line spacing (spasi baris). Dikembalikan sebagai angka
    kelipatan (1.0, 1.5, 2.0, ...) kalau bentuknya "multiple",
    atau dalam satuan point kalau pengguna set nilai pasti (fixed).
    """
    pf = paragraph.paragraph_format
    if pf.line_spacing is None:
        # coba warisi dari style
        style = paragraph.style
        while style is not None:
            if style.paragraph_format.line_spacing is not None:
                return style.paragraph_format.line_spacing
            style = style.base_style
        return None
    return pf.line_spacing


import re

from docx.enum.text import WD_ALIGN_PARAGRAPH as _ALIGN

# "BAB I", "BAB II", dan juga "BAB 1" (angka Arab) -- ditemukan dari pengujian dokumen nyata
BAB_PATTERN = re.compile(r"^BAB\s+([IVXLCDM]+|\d+)\b", re.IGNORECASE)
DAFTAR_PUSTAKA_PATTERN = re.compile(
    r"^(DAFTAR\s+(PUSTAKA|RUJUKAN|REFERENSI)|REFERENSI|REFERENCES)\s*$", re.IGNORECASE
)
LAMPIRAN_PATTERN = re.compile(r"^LAMPIRAN\b", re.IGNORECASE)
KETERANGAN_PATTERN = re.compile(r"^(Gambar|Tabel|Grafik|Bagan|Diagram)\s+\d|^Sumber\s*:", re.IGNORECASE)
NOMOR_PATTERN = re.compile(r"^(\d+(\.\d+)*[.)]?|[A-Za-z][.)]|[IVXLC]+\.)\s+\S")
BULLET_PATTERN = re.compile(r"^[\u2022\u25CF\u25E6\-\u2013*]\s+\S")
TOC_PATTERN = re.compile(r"(\.{3,}|\t)\s*\d+\s*$")
AKHIR_KALIMAT = re.compile(r"[.!?][\"\u201d\u2019)\]]*$")


def _nama_style(p):
    try:
        return (p.style.name or "").lower()
    except Exception:
        return ""


def _punya_penomoran(p):
    ppr = p._p.pPr
    return ppr is not None and ppr.numPr is not None


def _semua_tebal(p):
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def _jenis_paragraf(p, teks, zona):
    """Tentukan jenis satu paragraf (bukan judul bab / daftar pustaka -- itu
    ditangani pemanggil). Mengembalikan salah satu zona:

      isi            kalimat isi di dalam bab (indentasi, rata, spasi diperiksa)
      kalimat_awal   kalimat isi di bagian awal, mis. Kata Pengantar
      butir_daftar   butir daftar bernomor/berpoin & lanjutannya (tidak diindentasi)
      subjudul       judul subbab / judul teks tebal / teks rata tengah
      keterangan     keterangan gambar/tabel, "Sumber: ..."
      awal           sampul, daftar isi, dan baris pendek lain di bagian awal
      daftar_pustaka / lampiran
    """
    if zona == "daftar_pustaka":
        return "daftar_pustaka"
    if zona == "lampiran":
        return "lampiran"

    nama = _nama_style(p)
    pf = p.paragraph_format
    tengah = pf.alignment == _ALIGN.CENTER
    di_awal = zona == "awal"

    if nama.startswith("toc") or TOC_PATTERN.search(teks):
        return "awal"
    if nama.startswith(("heading", "judul", "title", "subtitle")):
        return "awal" if di_awal else "subjudul"
    if KETERANGAN_PATTERN.match(teks):
        return "keterangan"
    if tengah:
        return "awal" if di_awal else "subjudul"
    if _punya_penomoran(p) or nama.startswith("list"):
        return "butir_daftar"
    if BULLET_PATTERN.match(teks):
        return "butir_daftar"

    pendek = len(teks) <= 120
    berakhir_kalimat = bool(AKHIR_KALIMAT.search(teks))
    if NOMOR_PATTERN.match(teks):
        if berakhir_kalimat and not _semua_tebal(p):
            return "butir_daftar"          # "1. Kalimat lengkap."  -> butir manual
        if pendek:
            return "subjudul"              # "1.1 Latar Belakang"   -> judul manual
    if pendek and not berakhir_kalimat and _semua_tebal(p):
        return "subjudul"

    if di_awal:
        # Bagian awal: hanya paragraf yang tampak seperti kalimat (panjang >=150 karakter, atau >=80 dan
        # berakhir tanda baca; tidak rata tengah) -- sisanya sampul/daftar isi/tanda tangan.
        if len(teks) >= 150 or (len(teks) >= 80 and berakhir_kalimat):
            return "kalimat_awal"
        return "awal"

    # Di dalam bab: lanjutan butir daftar = paragraf yang menjorok penuh dari kiri
    # tanpa indentasi baris pertama (bukan paragraf isi biasa).
    kiri = pf.left_indent.cm if pf.left_indent is not None else 0.0
    baris1 = pf.first_line_indent.cm if pf.first_line_indent is not None else 0.0
    if kiri >= 0.4 and abs(baris1) < 0.05:
        return "butir_daftar"
    return "isi"


def _classify_zones(doc: Document):
    """
    Klasifikasi SETIAP paragraf (indeks sama dengan doc.paragraphs).

    TEMUAN DARI PENGUJIAN KE DOKUMEN ASLI (penting untuk Metode/Pembahasan artikel):
    1) Satu aturan "isi" yang disamaratakan ke semua paragraf menghasilkan banyak
       false positive (sampul dituduh salah rata, judul bab dituduh salah
       indentasi, daftar pustaka dituduh salah indentasi padahal hanging indent).
    2) Perbaikan otomatis yang juga menyamaratakan aturan merusak dokumen:
       judul, sampul, butir daftar, dan daftar pustaka ikut menjorok 1 cm.
    Karena itu pemeriksa DAN perbaikan memakai klasifikasi yang sama; lihat
    ATURAN_PER_ZONA di bawah untuk aspek apa yang berlaku di tiap jenis.

    KETERBATASAN (bahan Batasan Penelitian): klasifikasi berbasis pola teks,
    style, penomoran, dan tampilan (tebal/rata tengah/panjang). Variasi yang tak
    terduga (mis. "CHAPTER 1", judul manual tanpa tebal dan tanpa nomor, paragraf isi
    yang sengaja dijorokkan penuh) bisa salah digolongkan. Hasil pemeriksaan
    ditampilkan sebagai saran, dan pengguna tetap perlu memeriksa hasil perbaikan.
    """
    hasil = []
    zona = "awal"
    for p in doc.paragraphs:
        teks = p.text.strip()
        if not teks:
            hasil.append(zona)
            continue
        judul_style = _nama_style(p).startswith(("heading", "judul", "title"))
        boleh_judul = len(teks) <= 80 or judul_style

        if boleh_judul and BAB_PATTERN.match(teks):
            zona = "isi"
            hasil.append("judul_bab")
            continue
        if boleh_judul and DAFTAR_PUSTAKA_PATTERN.match(teks):
            zona = "daftar_pustaka"
            hasil.append("judul_daftar_pustaka")
            continue
        if boleh_judul and LAMPIRAN_PATTERN.match(teks):
            zona = "lampiran"
            hasil.append("judul_lampiran")
            continue
        hasil.append(_jenis_paragraf(p, teks, zona))
    return hasil


# Aspek yang diperiksa DAN diperbaiki per jenis paragraf. Jenis yang tidak
# tercantum (judul, sampul, daftar isi, keterangan, lampiran) tidak diindentasi,
# tidak diratakan, dan tidak diubah spasi/ukurannya.
ATURAN_PER_ZONA = {
    "isi":            ("font", "ukuran", "spasi", "indentasi", "perataan"),
    "kalimat_awal":   ("font", "ukuran", "indentasi", "perataan"),
    "butir_daftar":   ("font", "ukuran", "spasi"),
    "daftar_pustaka": ("font", "ukuran"),
}


def classify_zones_untuk_italic(doc: Document):
    """Alias publik dari _classify_zones -- dipakai modul lain (italic_detector)
    yang butuh tahu zona paragraf tanpa mengimpor fungsi berawalan underscore."""
    return _classify_zones(doc)


def get_paragraphs_info(doc: Document):
    """
    Ambil info format per paragraf: teks (dipotong untuk identifikasi),
    font, ukuran font, spasi baris, indentasi awal, perataan, DAN
    zona dokumen (lihat _classify_zones) supaya rule_engine tahu
    aturan mana yang relevan diterapkan ke paragraf ini.

    Setiap baris hasil punya "index" supaya nanti laporan pelanggaran
    bisa menunjuk "paragraf ke-N", bukan cuma bilang "ada yang salah".
    """
    zona_semua = _classify_zones(doc)
    hasil = []
    for idx, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            continue  # lewati paragraf kosong, tidak relevan diperiksa

        pf = p.paragraph_format
        hasil.append({
            "index": idx,
            "teks_cuplikan": p.text.strip()[:60],
            "zona": zona_semua[idx],
            "font_name": _resolve_font_name(p, document=doc),
            "font_size_pt": _resolve_font_size_pt(p, document=doc),
            "line_spacing": _resolve_line_spacing(p),
            "first_line_indent_cm": emu_to_cm(pf.first_line_indent),
            "alignment": pf.alignment,  # objek enum, diterjemahkan di rule engine
            "is_italic_candidates": _find_non_italic_runs(p),
        })
    return hasil


def _find_non_italic_runs(paragraph):
    """
    Placeholder sederhana untuk tahap selanjutnya (deteksi istilah
    asing yang belum dimiringkan). Untuk sekarang hanya mencatat
    run mana saja yang TIDAK italic, supaya tahap berikut tinggal
    menyaring run mana yang berisi kata asing.
    """
    return [
        run.text for run in paragraph.runs
        if run.text.strip() and not run.font.italic
    ]


if __name__ == "__main__":
    # Cara cepat menguji modul ini sendiri dari terminal:
    #   python3 parser.py path/ke/dokumen.docx
    import sys
    import json

    if len(sys.argv) < 2:
        print("Pakai: python3 parser.py <file.docx>")
        sys.exit(1)

    doc = Document(sys.argv[1])
    print("=== Page setup ===")
    print(json.dumps(get_page_setup(doc), indent=2, ensure_ascii=False))
    print("=== Paragraf (10 pertama) ===")
    info = get_paragraphs_info(doc)
    for row in info[:10]:
        row.pop("is_italic_candidates")
        print(row)
