"""
auto_fix.py
===========
Memperbaiki LANGSUNG properti yang AMAN diubah otomatis:
margin, ukuran kertas, jenis/ukuran font, spasi baris, indentasi,
perataan.

KEPUTUSAN DESAIN (penting untuk bab Metode/Pembahasan artikel):
Italic (istilah asing) dan typo TIDAK diperbaiki di sini -- itu
ditandai lewat Word comments di modul terpisah (annotator.py),
karena auto-fix untuk dua hal itu berisiko salah pada nama orang,
akronim, atau istilah yang sudah lazim. Keputusan ini disengaja,
bukan keterbatasan teknis yang belum dikerjakan.

Prinsip penting: JANGAN menimpa file asli pengguna. Selalu simpan
ke file baru. File asli adalah milik pengguna, bukan bahan kerja.
"""

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def fix_page_setup(doc: Document, rules: dict):
    r = rules["ukuran_kertas"]
    m = rules["margin"]
    for section in doc.sections:
        section.page_width = Cm(r["lebar_cm"])
        section.page_height = Cm(r["tinggi_cm"])
        section.top_margin = Cm(m["atas_cm"])
        section.bottom_margin = Cm(m["bawah_cm"])
        section.left_margin = Cm(m["kiri_cm"])
        section.right_margin = Cm(m["kanan_cm"])


def fix_paragraph_formatting(doc: Document, rules: dict):
    f_rules = rules["font"]
    s_rules = rules["spasi"]
    i_rules = rules["indentasi"]
    a_rules = rules["perataan"]

    alignment_target = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[a_rules["jenis"]]

    for p in doc.paragraphs:
        if not p.text.strip():
            continue  # jangan sentuh paragraf kosong (biasanya pemisah)

        pf = p.paragraph_format
        pf.line_spacing = s_rules["baris_isi"]
        pf.first_line_indent = Cm(i_rules["awal_paragraf_cm"])
        pf.alignment = alignment_target

        for run in p.runs:
            if run.text.strip():
                run.font.name = f_rules["nama"]
                run.font.size = Pt(f_rules["ukuran_isi_pt"])


def auto_fix_document(input_path: str, output_path: str, rules: dict):
    """
    Fungsi utama: baca file input, perbaiki, simpan ke output_path
    (BUKAN menimpa input_path).
    """
    doc = Document(input_path)
    fix_page_setup(doc, rules)
    fix_paragraph_formatting(doc, rules)
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
