"""
annotator.py
============
Menempelkan Word COMMENT (bukan mengubah teks) ke run yang mengandung
kandidat istilah asing. Pengguna membuka dokumen di Word/LibreOffice,
lihat komentarnya di panel kanan, lalu putuskan sendiri apakah mau
dimiringkan atau tidak.

Kenapa comment, bukan highlight warna atau tanda kurung?
Comment adalah mekanisme bawaan Word yang paling umum dipahami
pengguna, tidak mengubah tata letak dokumen, dan mudah dihapus massal
oleh pengguna (Review > Delete All Comments) kalau tidak diperlukan.
"""

from docx import Document
import italic_detector


def annotate_italic_candidates(doc: Document, candidates: list, author: str = "MakaRapi"):
    """
    candidates: hasil dari italic_detector.find_italic_candidates(doc)
    Catatan: HARUS dipanggil dengan `doc` yang SAMA dengan yang dipakai
    saat candidates dihasilkan -- karena tiap kandidat menyimpan
    referensi objek run langsung dari dokumen itu.
    """
    jumlah = 0
    for kandidat in candidates:
        kata_list = ", ".join(f"'{k}'" for k in kandidat["kata_kandidat"])
        teks_komentar = (
            f"Kemungkinan istilah asing yang belum dicetak miring: {kata_list}. "
            f"Periksa apakah kata ini sudah baku di KBBI; jika belum, "
            f"pertimbangkan menulisnya dengan format italic."
        )
        doc.add_comment(
            runs=kandidat["run"],
            text=teks_komentar,
            author=author,
            initials="MR",
        )
        jumlah += 1
    return jumlah


def process_document(input_path: str, output_path: str, daftar_path="../daftar_istilah_asing.yaml"):
    """Fungsi gabungan: baca, deteksi, tandai, simpan ke file baru."""
    doc = Document(input_path)
    candidates = italic_detector.find_italic_candidates(doc, daftar_path)
    jumlah = annotate_italic_candidates(doc, candidates)
    doc.save(output_path)
    return jumlah


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Pakai: python3 annotator.py <input.docx> <output.docx>")
        sys.exit(1)

    jumlah = process_document(sys.argv[1], sys.argv[2])
    print(f"Selesai. {jumlah} komentar ditambahkan. Hasil: {sys.argv[2]}")
