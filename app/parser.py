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

BAB_PATTERN = re.compile(r"^BAB\s+[IVXLCDM]+\b", re.IGNORECASE)
DAFTAR_PUSTAKA_PATTERN = re.compile(r"^DAFTAR\s+PUSTAKA\s*$", re.IGNORECASE)


def _classify_zones(doc: Document):
    """
    TEMUAN DARI PENGUJIAN KE DOKUMEN ASLI (penting dibaca):
    Rule engine awalnya menyamaratakan SATU aturan (margin, font,
    indentasi, perataan "isi") ke SEMUA paragraf tanpa kecuali.
    Begitu diuji ke makalah mahasiswa sungguhan, ini menghasilkan
    banyak sekali false positive: judul di halaman sampul dituduh
    "salah rata" padahal sampul memang konvensinya rata tengah,
    "BAB I" dituduh "salah indentasi" padahal judul bab memang tidak
    pakai indentasi, dan entri Daftar Pustaka dituduh "salah
    indentasi" padahal itu memang hanging indent yang benar.

    Fungsi ini membagi dokumen jadi beberapa ZONA berdasarkan pola
    teks yang umum dipakai di makalah/skripsi Indonesia, supaya
    rule_engine bisa menerapkan aturan yang BERBEDA per zona,
    bukan satu aturan untuk semua:

      - "awal"            : sampul, kata pengantar, daftar isi --
                             sebelum "BAB I" pertama kali muncul.
      - "judul_bab"        : baris "BAB I/II/III/..." itu sendiri.
      - "isi"              : paragraf isi di dalam bab.
      - "judul_daftar_pustaka": baris "DAFTAR PUSTAKA" itu sendiri.
      - "daftar_pustaka"   : entri-entri rujukan setelah itu.

    KETERBATASAN yang harus disadari (bahan Batasan Penelitian):
    Deteksi ini berbasis pola teks sederhana (regex "BAB <angka
    romawi>" dan "DAFTAR PUSTAKA"). Kalau mahasiswa menulis variasi
    lain (misalnya "Daftar Rujukan", atau "CHAPTER 1"), zona tidak
    akan terdeteksi benar. Ini perlu disebut sebagai batasan, bukan
    disembunyikan.
    """
    zona_per_paragraf = []
    zona_sekarang = "awal"

    for p in doc.paragraphs:
        teks = p.text.strip()

        if BAB_PATTERN.match(teks):
            zona_sekarang = "isi"
            zona_per_paragraf.append("judul_bab")
            continue

        if DAFTAR_PUSTAKA_PATTERN.match(teks):
            zona_sekarang = "daftar_pustaka"
            zona_per_paragraf.append("judul_daftar_pustaka")
            continue

        zona_per_paragraf.append(zona_sekarang)

    return zona_per_paragraf


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
