"""
italic_detector.py
===================
Mendeteksi KANDIDAT istilah asing yang belum dicetak miring.

RIWAYAT DESAIN (penting dibaca, ini perubahan arsitektur, bukan
sekadar perbaikan bug):

Versi PERTAMA memakai pendekatan "exclusion-list": tandai semua kata
yang cocok dengan kamus Bahasa Inggris (pustaka `english_words`),
lalu kecualikan yang ada di daftar putih kata serapan baku. Versi ini
GAGAL saat diuji ke makalah asli -- puluhan kata Indonesia biasa
("hakikat", "watak", "nalar", "corak", "runtut", "paragraf", dst)
kebetulan juga tercatat di kamus Inggris kuno (Webster 1934) sebagai
kata arkais atau istilah ilmiah, sehingga salah tertandai sebagai
istilah asing. Menambah daftar pengecualian satu per satu tidak akan
pernah cukup -- yang dibutuhkan sebenarnya seluruh kamus Bahasa
Indonesia untuk pengecualian yang benar, yang tidak tersedia sebagai
pustaka Python siap pakai.

Versi KEDUA (dipakai sekarang) membalik pendekatannya jadi
"inclusion-list": sistem HANYA menandai kata yang benar-benar ADA di
daftar istilah asing yang sudah dikenal (daftar_istilah_asing.yaml),
bukan menebak dari kamus umum. Konsekuensinya jujur dan harus ditulis
di Batasan Penelitian: presisi tinggi, tapi tidak akan mendeteksi
istilah asing BARU yang belum masuk daftar. Trade-off ini disengaja
-- untuk alat yang hasilnya cuma berupa SARAN (ditandai lewat
comment, bukan auto-fix), presisi tinggi lebih penting daripada
mencoba menangkap semua kemungkinan dan malah penuh salah tuduh.
"""

import re
import yaml

WORD_PATTERN = re.compile(r"[A-Za-z]+")


def load_daftar_istilah(path="../daftar_istilah_asing.yaml"):
    """Gabungkan semua kategori di file jadi satu set kata (lowercase)."""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    semua_istilah = set()
    for kategori, daftar in data.items():
        for kata in daftar:
            semua_istilah.add(kata.lower())
    return semua_istilah


def find_italic_candidates(doc, daftar_path="../daftar_istilah_asing.yaml"):
    """
    Telusuri semua paragraf dan semua run yang BELUM italic.
    Tandai kata yang PERSIS cocok (case-insensitive) dengan salah
    satu entri di daftar_istilah_asing.yaml.
    """
    daftar_istilah = load_daftar_istilah(daftar_path)
    hasil = []

    for p_idx, paragraph in enumerate(doc.paragraphs):
        for run in paragraph.runs:
            if run.font.italic:
                continue  # sudah dimiringkan, lewati
            if not run.text.strip():
                continue

            kata_dalam_run = WORD_PATTERN.findall(run.text)
            kandidat_di_run = [
                w for w in kata_dalam_run if w.lower() in daftar_istilah
            ]

            if kandidat_di_run:
                hasil.append({
                    "paragraf_index": p_idx,
                    "run": run,
                    "kata_kandidat": kandidat_di_run,
                    "konteks": run.text.strip()[:80],
                })

    return hasil


if __name__ == "__main__":
    import sys
    from docx import Document

    if len(sys.argv) < 2:
        print("Pakai: python3 italic_detector.py <file.docx>")
        sys.exit(1)

    doc = Document(sys.argv[1])
    kandidat = find_italic_candidates(doc)
    print(f"Ditemukan {len(kandidat)} run dengan kandidat istilah asing:\n")
    for k in kandidat:
        print(f"  Paragraf #{k['paragraf_index']}: {k['kata_kandidat']} -- \"{k['konteks']}\"")
