"""
main.py
=======
Antarmuka web (Streamlit) untuk MakaRapi.

Jalankan dengan:
    streamlit run main.py

Alur pengguna:
  1. Upload file .docx
  2. Sistem menampilkan laporan pelanggaran + skor kepatuhan
  3. Pengguna bisa unduh versi yang sudah diperbaiki otomatis
     (margin, font, spasi, indentasi, perataan)

CATATAN: deteksi italic/typo (flag via comment, bukan auto-fix)
belum dimasukkan di tahap ini -- menyusul sesuai peta jalan
(Bulan 2, Minggu 5-6) supaya tidak membangun semuanya sekaligus.
"""

import io
import tempfile
import os

import streamlit as st
from docx import Document

import rule_engine
import auto_fix
import parser as parser_mod
import italic_detector
import annotator


st.set_page_config(page_title="MakaRapi", page_icon="📄", layout="centered")

st.title("📄 MakaRapi")
st.caption("Pemeriksa & perbaikan otomatis format penulisan makalah sesuai Pedoman UIN Gusdur")

rules = rule_engine.load_rules("../rules.yaml")

uploaded = st.file_uploader("Unggah dokumen (.docx)", type=["docx"])

if uploaded is not None:
    # Simpan ke file sementara karena python-docx & LibreOffice butuh
    # path file, tidak bisa langsung baca dari objek upload Streamlit.
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
        tmp_in.write(uploaded.getvalue())
        input_path = tmp_in.name

    doc = Document(input_path)
    page_setup = parser_mod.get_page_setup(doc)
    paragraphs = parser_mod.get_paragraphs_info(doc)
    hasil = rule_engine.run_all_checks(page_setup, paragraphs, rules)

    col1, col2 = st.columns(2)
    col1.metric("Skor kepatuhan", f"{hasil['skor_kepatuhan_persen']}%")
    col2.metric("Jumlah pelanggaran", hasil["jumlah_pelanggaran"])

    if hasil["violations"]:
        st.subheader("Daftar pelanggaran")
        st.dataframe(hasil["violations"], use_container_width=True)
    else:
        st.success("Tidak ada pelanggaran format deterministik yang ditemukan.")

    st.subheader("Unduh hasil perbaikan otomatis")
    st.caption(
        "Hanya margin, ukuran kertas, jenis/ukuran font, spasi, indentasi, "
        "dan perataan yang diperbaiki otomatis. Italic dan typo tidak "
        "diubah langsung (akan ditandai lewat komentar pada tahap berikutnya)."
    )

    output_path = input_path.replace(".docx", "_fixed.docx")
    auto_fix.auto_fix_document(input_path, output_path, rules)

    with open(output_path, "rb") as f:
        st.download_button(
            label="⬇️ Unduh dokumen hasil perbaikan (margin/font/spasi)",
            data=f.read(),
            file_name=f"fixed_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.subheader("Kandidat istilah asing yang belum dimiringkan")
    st.caption(
        "Ditandai lewat komentar Word, bukan diubah langsung -- buka hasilnya "
        "di Word/LibreOffice dan lihat panel komentar di sisi kanan."
    )

    doc_untuk_anotasi = Document(output_path)  # pakai versi yang sudah di-auto-fix
    kandidat = italic_detector.find_italic_candidates(doc_untuk_anotasi, daftar_path="../daftar_istilah_asing.yaml")

    if kandidat:
        st.dataframe(
            [{"paragraf": k["paragraf_index"], "kata_kandidat": ", ".join(k["kata_kandidat"]),
              "konteks": k["konteks"]} for k in kandidat],
            use_container_width=True,
        )
        annotator.annotate_italic_candidates(doc_untuk_anotasi, kandidat)
        annotated_path = input_path.replace(".docx", "_annotated.docx")
        doc_untuk_anotasi.save(annotated_path)

        with open(annotated_path, "rb") as f:
            st.download_button(
                label="⬇️ Unduh dokumen dengan tanda komentar istilah asing",
                data=f.read(),
                file_name=f"annotated_{uploaded.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        os.unlink(annotated_path)
    else:
        st.success("Tidak ada kandidat istilah asing yang ditemukan.")

    os.unlink(input_path)
    os.unlink(output_path)
else:
    st.info("Silakan unggah dokumen .docx untuk mulai diperiksa.")
