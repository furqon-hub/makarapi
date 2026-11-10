"""Membuat dokumen .docx contoh dengan format SENGAJA salah,
supaya kita bisa menguji apakah rule_engine berhasil mendeteksinya."""

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

section = doc.sections[0]
section.page_width = Cm(21.59)   # ini ukuran Letter, SALAH (seharusnya A4 21cm)
section.page_height = Cm(27.94)
section.top_margin = Cm(2.5)     # SALAH, seharusnya 4cm
section.bottom_margin = Cm(2.5)  # SALAH, seharusnya 3cm
section.left_margin = Cm(3.0)    # SALAH, seharusnya 4cm
section.right_margin = Cm(3.0)   # ini benar

p1 = doc.add_paragraph("Ini adalah paragraf contoh dengan font Arial yang salah dan ukuran 11pt.")
p1.paragraph_format.line_spacing = 1.0  # SALAH, seharusnya 2.0
p1.paragraph_format.first_line_indent = Cm(0.5)  # SALAH, seharusnya 1.27cm
p1.alignment = WD_ALIGN_PARAGRAPH.LEFT  # SALAH, seharusnya justify
for run in p1.runs:
    run.font.name = "Arial"
    run.font.size = Pt(11)

p2 = doc.add_paragraph("Paragraf kedua ini sudah benar formatnya, Times New Roman 12pt spasi 2.")
p2.paragraph_format.line_spacing = 2.0
p2.paragraph_format.first_line_indent = Cm(1.27)
p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for run in p2.runs:
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

doc.save("sample_salah.docx")
print("Dibuat: sample_salah.docx")
