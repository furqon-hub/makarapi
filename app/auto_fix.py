"""
auto_fix.py
===========
Memperbaiki LANGSUNG properti yang AMAN diubah otomatis:
margin, ukuran kertas, jenis/ukuran font, spasi baris, indentasi,
perataan -- dengan paragraf isi dibedakan dari judul, sampul, daftar,
dan daftar pustaka (lihat fix_paragraph_formatting).

KEPUTUSAN DESAIN (penting untuk bab Metode/Pembahasan artikel):
Italic (istilah asing) dan typo TIDAK diperbaiki di sini -- itu
ditandai lewat Word comments di modul terpisah (annotator.py),
karena auto-fix untuk dua hal itu berisiko salah pada nama orang,
akronim, atau istilah yang sudah lazim. Keputusan ini disengaja,
bukan keterbatasan teknis yang belum dikerjakan.

Prinsip penting: JANGAN menimpa file asli pengguna. Selalu simpan
ke file baru. File asli adalah milik pengguna, bukan bahan kerja.
"""

import parser as parser_mod
import rule_engine
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING


def fix_page_setup(doc: Document, rules: dict):
    r = rules["ukuran_kertas"]
    m = rules["margin"]
    for section in doc.sections:
        # nilai null = tidak diatur pedoman -> dibiarkan apa adanya
        if r.get("lebar_cm") is not None:
            section.page_width = Cm(r["lebar_cm"])
        if r.get("tinggi_cm") is not None:
            section.page_height = Cm(r["tinggi_cm"])
        if m.get("atas_cm") is not None:
            section.top_margin = Cm(m["atas_cm"])
        if m.get("bawah_cm") is not None:
            section.bottom_margin = Cm(m["bawah_cm"])
        if m.get("kiri_cm") is not None:
            section.left_margin = Cm(m["kiri_cm"])
        if m.get("kanan_cm") is not None:
            section.right_margin = Cm(m["kanan_cm"])


def fix_paragraph_formatting(doc: Document, rules: dict):
    """
    Perbaiki format per paragraf HANYA sesuai jenisnya (lihat parser.ATURAN_PER_ZONA).

    Bug yang ditemukan dari pengujian dokumen nyata: versi awal menerapkan indentasi
    1 cm, rata kiri-kanan, dan spasi ke SEMUA paragraf, sehingga judul bab/subbab,
    sampul, daftar isi, butir daftar, dan daftar pustaka (hanging indent) ikut rusak.
    Sekarang indentasi awal paragraf, perataan rata kiri-kanan, dan spasi hanya
    dikenakan pada jenis paragraf yang memang memakainya (kalimat isi).
    Nama font diseragamkan di semua teks; ukuran font hanya di jenis yang diperiksa.
    """
    f_rules = rules["font"]
    s_rules = rules["spasi"]
    i_rules = rules["indentasi"]
    a_rules = rules["perataan"]

    alignment_target = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }.get(a_rules.get("jenis"))
    aktif = rule_engine.aspek_aktif(rules)  # hanya aspek yang diatur profil

    zona_semua = parser_mod.classify_zones_untuk_italic(doc)

    for p, zona in zip(doc.paragraphs, zona_semua):
        if not p.text.strip():
            continue  # jangan sentuh paragraf kosong (biasanya pemisah)

        aspek = set(parser_mod.ATURAN_PER_ZONA.get(zona, ())) & aktif
        pf = p.paragraph_format

        if "spasi" in aspek:
            if s_rules.get("mode", "multiple") == "exact":
                pf.line_spacing = Pt(s_rules["baris_isi"])
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            else:
                pf.line_spacing = s_rules["baris_isi"]
        if "indentasi" in aspek:
            pf.first_line_indent = Cm(i_rules["awal_paragraf_cm"])
        if "perataan" in aspek and alignment_target is not None:
            pf.alignment = alignment_target

        for run in p.runs:
            if run.text.strip():
                if "font" in aktif:
                    run.font.name = f_rules["nama"]
                if "ukuran" in aspek:
                    run.font.size = Pt(f_rules["ukuran_isi_pt"])


def auto_fix_document(input_path: str, output_path: str, rules: dict):
    """
    Fungsi utama: baca file input, perbaiki, simpan ke output_path
    (BUKAN menimpa input_path).
    """
    doc = Document(input_path)
    fix_page_setup(doc, rules)
    fix_paragraph_formatting(doc, rules)
    import ext_checks
    ext_checks.perbaiki(doc, rules)
    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys
    sys.path.insert(0, ".")
    import rule_engine

    if len(sys.argv) < 3:
        print("Pakai: python3 auto_fix.py <input.docx> <output.docx>")
        sys.exit(1)

    rules = rule_engine.load_rules("../rules.yaml")
    auto_fix_document(sys.argv[1], sys.argv[2], rules)
    print(f"Selesai. Hasil disimpan di: {sys.argv[2]}")
