"""Dokumen uji untuk italic_detector.py: campuran istilah asing murni,
kata serapan baku (harus TIDAK ditandai), dan kata Indonesia (harus
TIDAK ditandai)."""

from docx import Document

doc = Document()

doc.add_paragraph(
    "Penelitian ini menggunakan metode deployment otomatis untuk "
    "memperbarui sistem secara berkala."
)
# "deployment" -> SEHARUSNYA ketandai (istilah asing)
# "metode", "sistem" -> SEHARUSNYA TIDAK (kata serapan baku di whitelist)

doc.add_paragraph(
    "Framework yang digunakan mendukung pengujian otomatis melalui "
    "beberapa tahapan pipeline yang berjalan berurutan."
)
# "Framework", "pipeline" -> SEHARUSNYA ketandai
# "pengujian", "tahapan", "berjalan", "berurutan" -> kata Indonesia, SEHARUSNYA TIDAK

doc.add_paragraph(
    "Data dari aplikasi dikumpulkan melalui server pusat lalu "
    "dianalisis menggunakan program statistik."
)
# semua kata di sini ada di whitelist / pola imbuhan Indonesia
# -> SEHARUSNYA TIDAK ADA yang ketandai sama sekali

doc.save("sample_italic_test.docx")
print("Dibuat: sample_italic_test.docx")
