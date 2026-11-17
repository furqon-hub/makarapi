"""
ext_checks.py
=============
Pemeriksaan TAMBAHAN di luar margin/font/spasi: judul bab & subbab (gaya Heading,
penomoran otomatis, font sama dengan isi), nomor halaman, halaman baru, bagian wajib
(Daftar Isi/Tabel/Gambar dibuatkan otomatis), keterangan tabel/gambar, daftar pustaka.

Semua dikendalikan blok `tambahan` pada aturan: null / false / kosong = "tidak diatur"
-> tidak diperiksa, tidak diubah. Tiap temuan punya `perbaikan`:
"otomatis" (diperbaiki tombol Perbaiki) atau "manual".
"""

import re

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import parser as P

BAGIAN = {
    "kata_pengantar": ("Kata Pengantar", re.compile(r"^KATA\s+PENGANTAR\s*$", re.I)),
    "abstrak": ("Abstrak", re.compile(r"^(ABSTRAK|ABSTRACT)\s*$", re.I)),
    "daftar_isi": ("Daftar Isi", re.compile(r"^DAFTAR\s+ISI\s*$", re.I)),
    "daftar_tabel": ("Daftar Tabel", re.compile(r"^DAFTAR\s+TABEL\s*$", re.I)),
    "daftar_gambar": ("Daftar Gambar", re.compile(r"^DAFTAR\s+GAMBAR\s*$", re.I)),
    "daftar_pustaka": ("Daftar Pustaka", P.DAFTAR_PUSTAKA_PATTERN),
    "lampiran": ("Lampiran", re.compile(r"^(DAFTAR\s+)?LAMPIRAN\b", re.I)),
}
DIBUAT_OTOMATIS = ("daftar_isi", "daftar_tabel", "daftar_gambar")

POSISI_HALAMAN = {"bawah_tengah": ("footer", "center"), "bawah_kanan": ("footer", "right"),
                  "atas_kanan": ("header", "right"), "atas_tengah": ("header", "center")}
_RATA = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}

KAP_TABEL = re.compile(r"^Tabel\s+\d", re.I)
KAP_GAMBAR = re.compile(r"^(Gambar|Grafik|Bagan|Diagram)\s+\d", re.I)
AWALAN_BAB = re.compile(r"^\s*BAB\s+([IVXLCDM]+|\d+)\s*[:.\-–]?\s*", re.I)
AWALAN_SUB = re.compile(r"^\s*(\d+(\.\d+)+[.)]?|\d+[.)]|[A-Za-z][.)])\s+")
NOMOR_SUB = re.compile(r"^\d+(\.\d+)+[.)]?\s+\S")
ROMAWI = [(10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]


def _romawi(n):
    out = ""
    for v, s in ROMAWI:
        while n >= v:
            out += s
            n -= v
    return out


# ---------- struktur dokumen ----------

def _urutan_badan(doc):
    hasil, i = [], 0
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            p = doc.paragraphs[i]
            hasil.append({"jenis": "p", "idx": i, "p": p, "el": el,
                          "gambar": bool(el.xpath(".//w:drawing | .//w:pict")), "teks": p.text.strip()})
            i += 1
        elif el.tag == qn("w:tbl"):
            hasil.append({"jenis": "tbl", "el": el})
    return hasil


def _section_per_paragraf(doc):
    peta, sec = [], 0
    for p in doc.paragraphs:
        peta.append(sec)
        if p._p.pPr is not None and p._p.pPr.find(qn("w:sectPr")) is not None:
            sec += 1
    return peta


def _punya_page(el):
    xml = el.xml
    return bool(re.search(r"<w:instrText[^>]*>\s*PAGE\b", xml) or re.search(r'w:instr="\s*PAGE\b', xml))


def _halaman_info(doc, sec_idx, jenis):
    asal = getattr(doc.sections[sec_idx], jenis)
    k = sec_idx
    while asal.is_linked_to_previous and k > 0:
        k -= 1
        asal = getattr(doc.sections[k], jenis)
    from docx.text.paragraph import Paragraph
    for el in asal._element.iter(qn("w:p")):   # termasuk paragraf di dalam content control (sdt)/tabel
        if _punya_page(el):
            return True, Paragraph(el, asal).alignment, asal
    return False, None, asal


def _level_subbab(p, teks):
    nama = P._nama_style(p)
    m = re.match(r"heading (\d)", nama)
    if m and int(m.group(1)) >= 2:
        return min(int(m.group(1)), 3)
    if nama.startswith("heading 1"):
        return None
    if p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return None
    if NOMOR_SUB.match(teks):
        return min(teks.split()[0].rstrip(".)").count(".") + 1, 3)
    if len(teks) <= 100 and not P.AKHIR_KALIMAT.search(teks):
        if re.match(r"^[a-z][.)]\s+\S", teks):
            return 3
        if P.NOMOR_PATTERN.match(teks) or P._semua_tebal(p):
            return 2
    return None


def _daftar_judul(doc, zona):
    """[(indeks, paragraf, level 1..3)] -- bab = 1, subbab = 2/3."""
    hasil = []
    for i, p in enumerate(doc.paragraphs):
        teks = p.text.strip()
        if not teks:
            continue
        if zona[i] == "judul_bab":
            hasil.append((i, p, 1))
        elif zona[i] == "subjudul":
            lv = _level_subbab(p, teks)
            if lv:
                hasil.append((i, p, lv))
    return hasil


def _adalah_subbab(p, teks):
    return _level_subbab(p, teks) is not None


def _awal_halaman_baru(badan, k):
    """Apakah elemen badan ke-k diawali halaman baru (page break / section break / pageBreakBefore)?"""
    p = badan[k]["p"]
    pf = p.paragraph_format
    if pf.page_break_before:
        return True
    st = p.style
    while st is not None:
        if st.paragraph_format.page_break_before:
            return True
        st = st.base_style
    j = k - 1
    while j >= 0:
        e = badan[j]
        if e["jenis"] == "tbl":
            return False
        el = e["el"]
        if el.xpath('.//w:br[@w:type="page"]'):
            return True
        ppr = el.find(qn("w:pPr"))
        if ppr is not None:
            sp = ppr.find(qn("w:sectPr"))
            if sp is not None:
                tp = sp.find(qn("w:type"))
                if tp is None or tp.get(qn("w:val")) != "continuous":
                    return True
        if e["teks"] or e["gambar"]:
            return False
        j -= 1
    return True  # awal dokumen


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


def _keterangan_dokumen(badan, zona):
    """(daftar keterangan tabel, gambar, jumlah tabel, jumlah gambar) setelah BAB pertama."""
    mulai = next((k for k, e in enumerate(badan) if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    kt, kg, nt, ng = [], [], 0, 0
    if mulai is None:
        return kt, kg, nt, ng
    for e in badan[mulai:]:
        if e["jenis"] == "tbl":
            nt += 1
        elif e["jenis"] == "p":
            if e["gambar"]:
                ng += 1
            if KAP_TABEL.match(e["teks"]):
                kt.append(e)
            elif KAP_GAMBAR.match(e["teks"]):
                kg.append(e)
    return kt, kg, nt, ng



# ---------- helper tambahan: bagian, daftar, butir, keterangan, tabel ----------

TOC_STYLE = ("toc", "table of figures")
PENANDA_BUTIR = [
    ("decimal", re.compile(r"^\s*\d+[.)]\s+")),
    ("lowerLetter", re.compile(r"^\s*[a-z][.)]\s+")),
    ("bullet", re.compile(r"^\s*[•●◦\-–*]\s+")),
]
NSID_BUTIR = {"decimal": "4D414B42", "lowerLetter": "4D414B43", "bullet": "4D414B44"}


def _jenis_penanda(teks):
    for jenis, pola in PENANDA_BUTIR:
        if pola.match(teks):
            return jenis
    return None


def _judul_bagian(badan, zona):
    """Elemen badan yang berupa judul bagian (Kata Pengantar, Daftar Isi, ...)."""
    hasil = []
    for e in badan:
        if e["jenis"] != "p" or not e["teks"] or len(e["teks"]) > 60:
            continue
        if zona[e["idx"]] == "judul_bab" or P.TOC_PATTERN.search(e["teks"]):
            continue
        if P._nama_style(e["p"]).startswith(TOC_STYLE):
            continue
        if any(pola.match(e["teks"]) for _, pola in BAGIAN.values()):
            hasil.append(e)
    return hasil


def _lebar_tersedia_cm(doc):
    s = doc.sections[-1]
    try:
        return (s.page_width - s.left_margin - s.right_margin).cm
    except Exception:
        return 15.5


def _ukuran_tabel(n_kolom, dasar):
    if not dasar:
        return None
    kurang = 0 if n_kolom <= 4 else min(n_kolom - 4, 4)
    return max(dasar - kurang, 8)


def _tabel_dalam_bab(doc, badan, zona):
    from docx.table import Table
    mulai = next((k for k, e in enumerate(badan) if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    if mulai is None:
        return []
    return [Table(e["el"], doc._body) for e in badan[mulai:] if e["jenis"] == "tbl"]


def _grid_twips(tbl):
    return [int(g.get(qn("w:w")) or 0) for g in tbl._tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol"))]


def _tabel_ukuran_run(tbl):
    ukuran = set()
    for row in tbl.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    if r.text.strip():
                        ukuran.add(r.font.size.pt if r.font.size else None)
    return ukuran


def _format_isi_tabel(doc, tbl, nama, dasar):
    n = len(_grid_twips(tbl)) or len(tbl.columns)
    uk = _ukuran_tabel(n, dasar)
    for row in tbl.rows:
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    _atur_font_run(r, nama, uk)
    # lebar: skalakan kolom bila melebihi area teks
    pakai = int(_lebar_tersedia_cm(doc) * 567)
    grid = _grid_twips(tbl)
    total = sum(grid)
    if total > pakai and total > 0:
        k = pakai / total
        for g in tbl._tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol")):
            g.set(qn("w:w"), str(int(int(g.get(qn("w:w")) or 0) * k)))
        for tc in tbl._tbl.iter(qn("w:tcW")):
            if tc.get(qn("w:type")) in (None, "dxa"):
                tc.set(qn("w:w"), str(int(int(tc.get(qn("w:w")) or 0) * k)))
        tw = tbl._tbl.tblPr.find(qn("w:tblW")) if tbl._tbl.tblPr is not None else None
        if tw is not None:
            tw.set(qn("w:type"), "dxa")
            tw.set(qn("w:w"), str(pakai))


def _format_bagian_judul(doc, p, nama, ukuran, spasi=None):
    p.style = _style(doc, "Heading 1")
    ppr = p._p.get_or_add_pPr()
    if ppr.find(qn("w:numPr")) is None:      # cegah Heading 1 yang bernomor memberi "BAB" pada bagian ini
        _pasang_numpr(p, 0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    if spasi:
        p.paragraph_format.line_spacing = spasi
    for r in p.runs:
        r.text = r.text.upper()
        r.bold = True
        _atur_font_run(r, nama, ukuran)


def _format_daftar(doc, rules):
    f = rules.get("font") or {}
    nama, ukuran = f.get("nama"), f.get("ukuran_isi_pt")
    spasi = (rules.get("spasi") or {}).get("baris_isi")
    for sn in ("toc 1", "toc 2", "toc 3", "table of figures"):
        st = _cari_style(doc, sn)
        if st is None and sn.startswith("toc"):
            st = _style(doc, sn)
        if st is None:
            continue
        _atur_font_style(st, nama, ukuran)
        if spasi:
            st.paragraph_format.line_spacing = spasi
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.space_after = Pt(0)
    for p in doc.paragraphs:
        if P._nama_style(p).startswith(TOC_STYLE):
            if spasi:
                p.paragraph_format.line_spacing = spasi
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                _atur_font_run(r, nama, ukuran)


def _numbering_butir(doc, jenis):
    """abstractNum untuk butir (decimal / lowerLetter / bullet); kembalikan abstractNumId."""
    num_el = doc.part.numbering_part.element
    nsid = NSID_BUTIR[jenis]
    ada = num_el.xpath('w:abstractNum[w:nsid/@w:val="%s"]/@w:abstractNumId' % nsid)
    if ada:
        return int(ada[0])
    from docx.oxml import parse_xml
    ids = [int(x) for x in num_el.xpath("w:abstractNum/@w:abstractNumId")] or [0]
    aid = max(ids) + 1
    teks = {"decimal": "%1.", "lowerLetter": "%1.", "bullet": "•"}[jenis]
    xml = (
        '<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="%d">'
        '<w:nsid w:val="%s"/><w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' % (aid, nsid, jenis, teks)
    )
    el = parse_xml(xml)
    pertama_num = num_el.find(qn("w:num"))
    if pertama_num is not None:
        pertama_num.addprevious(el)
    else:
        num_el.append(el)
    return aid


def _num_baru(doc, abstract_id):
    from docx.oxml import parse_xml
    num_el = doc.part.numbering_part.element
    nids = [int(x) for x in num_el.xpath("w:num/@w:numId")] or [0]
    nid = max(nids) + 1
    num_el.append(parse_xml(
        '<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="%d">'
        '<w:abstractNumId w:val="%d"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
        % (nid, abstract_id)))
    return nid


def _font_tanda_paragraf(p, nama, ukuran):
    """Font nomor/poin otomatis mengikuti tanda paragraf (pPr/rPr)."""
    ppr = p._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ujung = ppr.find(qn("w:sectPr"))
        if ujung is None:
            ujung = ppr.find(qn("w:pPrChange"))
        if ujung is not None:
            ujung.addprevious(rpr)
        else:
            ppr.append(rpr)
    if nama:
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            rf.set(qn("w:" + a), nama)
    if ukuran:
        for tag in ("w:sz", "w:szCs"):
            e = rpr.find(qn(tag))
            if e is None:
                e = OxmlElement(tag)
                rpr.append(e)
            e.set(qn("w:val"), str(int(round(ukuran * 2))))


def _pasang_numpr(p, nid, ilvl=0):
    ppr = p._p.get_or_add_pPr()
    for old in ppr.findall(qn("w:numPr")):
        ppr.remove(old)
    numpr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(ilvl))
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(nid))
    numpr.append(il); numpr.append(ni)
    sisip = None
    for tag in ("w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc"):
        sisip = ppr.find(qn(tag))
        if sisip is not None:
            break
    if sisip is not None:
        sisip.addprevious(numpr)
    else:
        ppr.append(numpr)


def _butir_manual(doc, zona):
    """[(paragraf, jenis_penanda)] butir yang penomorannya diketik manual."""
    hasil = []
    for i, p in enumerate(doc.paragraphs):
        if zona[i] != "butir_daftar" or P.numpr_efektif(p) is not None:
            continue
        j = _jenis_penanda(p.text.strip())
        if j:
            hasil.append((p, j))
    return hasil


def _bab_baris_baru_ok(p):
    teks = p.text
    if P.numpr_efektif(p) is not None:
        return teks.startswith("\n") or "\n" in teks.strip()[:1] or not teks.strip()
    m = re.match(r"^\s*BAB\s+([IVXLCDM]+|\d+)[ \t]*(\n|$)", teks, re.I)
    if m:
        return True
    return not AWALAN_BAB.match(teks)   # bukan pola "BAB X ..." -> tidak ada yang bisa dinilai


def _paksa_baris_baru(p):
    """'BAB I PENDAHULUAN' / nomor otomatis + judul -> nomor di baris atas, judul di bawahnya."""
    def run_br(awal_teks, contoh):
        r = OxmlElement("w:r")
        if contoh is not None and contoh.find(qn("w:rPr")) is not None:
            import copy
            r.append(copy.deepcopy(contoh.find(qn("w:rPr"))))
        if awal_teks:
            t = OxmlElement("w:t")
            t.text = awal_teks
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
        r.append(OxmlElement("w:br"))
        return r
    runs = p._p.findall(qn("w:r"))
    contoh = next((r for r in runs if r.find(qn("w:t")) is not None), runs[0] if runs else None)
    ppr = p._p.pPr
    idx = 1 if ppr is not None else 0
    if P.numpr_efektif(p) is not None:
        if not p.text.startswith("\n"):
            p._p.insert(idx, run_br("", contoh))
        return
    m = AWALAN_BAB.match(p.text)
    if not m:
        return
    awalan = p.text[:m.end()].strip()
    if not p.text[m.end():].strip():
        return
    _hapus_awalan(p, AWALAN_BAB)
    p._p.insert(idx, run_br(awalan, contoh))


# ---------- pemeriksaan ----------

def jalankan(doc, paragraphs_info, rules):
    """Kembalikan (violations, total_checks_tambahan)."""
    t = rules.get("tambahan") or {}
    if not t:
        return [], 0
    zona = P.classify_zones_untuk_italic(doc)
    v, total = [], 0
    badan = _urutan_badan(doc)
    sec_map = _section_per_paragraf(doc)
    f_rules = rules.get("font") or {}

    def tambah(kat, lokasi, ditemukan, seharusnya, perbaikan, severity="deterministik"):
        v.append({"kategori": kat, "lokasi": lokasi, "ditemukan": ditemukan,
                  "seharusnya": seharusnya, "severity": severity, "perbaikan": perbaikan})

    jb, js, jd = t.get("judul_bab") or {}, t.get("judul_subbab") or {}, t.get("judul") or {}
    judul = _daftar_judul(doc, zona)

    # -- judul bab & subbab --
    for i, p, lv in judul:
        teks = p.text.strip()
        lok = (f"Judul bab: \"{teks[:50]}\"" if lv == 1 else f"Subbab: \"{teks[:50]}\"")
        if lv == 1:
            if jb.get("kapital"):
                total += 1
                if teks != teks.upper():
                    tambah("judul_bab_kapital", lok, "tidak huruf kapital", "HURUF KAPITAL", "otomatis")
            if jb.get("tebal"):
                total += 1
                if not P._semua_tebal(p) and not _style_tebal(p):
                    tambah("judul_bab_tebal", lok, "tidak tebal", "cetak tebal", "otomatis")
            if jb.get("rata_tengah"):
                total += 1
                if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    tambah("judul_bab_rata", lok, "tidak rata tengah", "rata tengah", "otomatis")
        elif js.get("tebal"):
            total += 1
            if not P._semua_tebal(p) and not _style_tebal(p):
                tambah("judul_subbab_tebal", lok, "tidak tebal", "cetak tebal", "otomatis")
        if jd.get("heading"):
            total += 1
            if not P._nama_style(p).startswith(f"heading {lv}"):
                tambah("judul_heading", lok, f"gaya \"{p.style.name}\"", f"gaya Heading {lv}", "otomatis")
        if jd.get("penomoran"):
            total += 1
            if P.numpr_efektif(p) is None:
                manual = bool(AWALAN_BAB.match(teks)) if lv == 1 else bool(AWALAN_SUB.match(teks))
                tambah("judul_penomoran", lok, "nomor diketik manual" if manual else "tanpa nomor",
                       "penomoran otomatis (multilevel)", "otomatis")
        if jd.get("font_isi"):
            nama_t, ukuran_t = f_rules.get("nama"), f_rules.get("ukuran_isi_pt")
            if nama_t or ukuran_t:
                total += 1
                nm = P._resolve_font_name(p, document=doc) or "Times New Roman"
                uk = P._resolve_font_size_pt(p, document=doc)
                salah = []
                if nama_t and nm != nama_t:
                    salah.append(f"{nm}")
                if ukuran_t and uk is not None and abs(uk - ukuran_t) > 0.1:
                    salah.append(f"{uk:g} pt")
                spasi_t = (rules.get("spasi") or {}).get("baris_isi")
                if spasi_t:
                    ls = P._resolve_line_spacing_info(p)
                    if ls[0] is None or ls[1] != "multiple" or abs(ls[0] - spasi_t) > 0.05:
                        salah.append(f"spasi {ls[0] if ls[0] is not None else 'bawaan'}")
                if salah:
                    tambah("judul_font", lok, ", ".join(salah),
                           f"{nama_t or ''} {ukuran_t:g} pt".strip() if ukuran_t else nama_t, "otomatis")

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
                    if rata_ada != _RATA[rata]:
                        tambah("nomor_halaman_posisi", lok, "posisi nomor tidak sesuai", f"{jenis}, {rata}", "otomatis")
        if nh.get("sampul_tanpa_nomor") and len(doc.sections) > 1:
            total += 1
            ada, _, _ = _halaman_info(doc, 0, jenis)
            if ada:
                tambah("nomor_halaman_sampul", "Section 0 (sampul)", "sampul bernomor", "sampul tanpa nomor", "manual")

    # -- halaman baru --
    hb = t.get("halaman_baru") or {}
    if hb.get("bab") or hb.get("bagian"):
        for k, e in enumerate(badan):
            if e["jenis"] != "p" or not e["teks"]:
                continue
            adalah_bab = zona[e["idx"]] == "judul_bab"
            adalah_bagian = (not adalah_bab and len(e["teks"]) <= 60 and not P.TOC_PATTERN.search(e["teks"])
                         and not P._nama_style(e["p"]).startswith(TOC_STYLE)
                             and not P._nama_style(e["p"]).startswith(TOC_STYLE)
                             and any(pola.match(e["teks"]) for _, pola in BAGIAN.values()))
            if (adalah_bab and hb.get("bab")) or (adalah_bagian and hb.get("bagian")):
                total += 1
                if not _awal_halaman_baru(badan, k):
                    tambah("halaman_baru", f"\"{e['teks'][:40]}\"", "tidak diawali halaman baru",
                           "mulai di halaman baru", "otomatis")

    # -- bagian wajib --
    judul_pendek = [e["teks"] for e in badan if e["jenis"] == "p" and 0 < len(e["teks"]) <= 60]
    kt, kg, nt, ng = _keterangan_dokumen(badan, zona)
    ada_bab = any(z == "judul_bab" for z in zona)
    for kunci in (t.get("bagian_wajib") or []):
        if kunci not in BAGIAN:
            continue
        nama, pola = BAGIAN[kunci]
        if kunci == "daftar_tabel" and not (kt or nt):
            continue   # tidak ada tabel -> daftar tabel tidak diperlukan
        if kunci == "daftar_gambar" and not (kg or ng):
            continue
        total += 1
        if not any(pola.match(x) for x in judul_pendek):
            otom = kunci in DIBUAT_OTOMATIS and ada_bab
            tambah("bagian_wajib", "Dokumen", f"\"{nama}\" tidak ditemukan", f"ada bagian {nama}",
                   "otomatis" if otom else "manual")

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
            atas, bawah = _teks_tetangga(badan, k, -1), _teks_tetangga(badan, k, +1)
            ada_atas, ada_bawah = bool(pola.match(atas)), bool(pola.match(bawah))
            if not (ada_atas or ada_bawah):
                tambah("keterangan", f"{tipe} setelah \"{(atas or bawah)[:40]}\"", f"{tipe} tanpa keterangan",
                       f"keterangan {tipe.lower()} di {aturan}", "manual")
            elif (aturan == "atas" and not ada_atas) or (aturan == "bawah" and not ada_bawah):
                tambah("keterangan_posisi", f"{tipe} dengan keterangan di sisi lain",
                       f"keterangan di {'bawah' if ada_bawah else 'atas'}", f"di {aturan}", "otomatis")

    # -- judul BAB: nomor di atas, judul di bawah --
    if jb.get("baris_baru"):
        for i, p, lv in judul:
            if lv == 1:
                total += 1
                if not _bab_baris_baru_ok(p):
                    tambah("judul_bab_baris", f"Judul bab: \"{p.text.strip()[:50]}\"",
                           "judul sejajar dengan nomor BAB", "judul di baris bawah nomor BAB", "otomatis")

    # -- judul bagian (Kata Pengantar, Daftar Isi, ...) & isi daftar --
    bg = t.get("bagian") or {}
    if bg.get("judul_tengah"):
        for e in _judul_bagian(badan, zona):
            total += 1
            p = e["p"]
            salah = []
            if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                salah.append("tidak rata tengah")
            if not (P._semua_tebal(p) or _style_tebal(p)):
                salah.append("tidak tebal")
            if e["teks"] != e["teks"].upper():
                salah.append("bukan huruf kapital")
            if not P._nama_style(p).startswith("heading 1"):
                salah.append("bukan Heading 1 (tidak masuk Daftar Isi)")
            if salah:
                tambah("bagian_judul", f"Judul bagian: \"{e['teks'][:40]}\"", ", ".join(salah),
                       "rata tengah, tebal, KAPITAL, Heading 1", "otomatis")
    if bg.get("isi_daftar") and (f_rules.get("nama") or f_rules.get("ukuran_isi_pt")):
        spasi_t = (rules.get("spasi") or {}).get("baris_isi")
        for i, p in enumerate(doc.paragraphs):
            if not p.text.strip() or not P._nama_style(p).startswith(TOC_STYLE):
                continue
            total += 1
            salah = []
            nm = P._resolve_font_name(p, document=doc) or "Times New Roman"
            uk = P._resolve_font_size_pt(p, document=doc)
            if f_rules.get("nama") and nm != f_rules["nama"]:
                salah.append(nm)
            if f_rules.get("ukuran_isi_pt") and uk is not None and abs(uk - f_rules["ukuran_isi_pt"]) > 0.1:
                salah.append(f"{uk:g} pt")
            if spasi_t:
                ls = P._resolve_line_spacing_info(p)
                if ls[0] is None or ls[1] != "multiple" or abs(ls[0] - spasi_t) > 0.05:
                    salah.append(f"spasi {ls[0] if ls[0] is not None else 'bawaan'}")
            ali = p.paragraph_format.alignment
            st_ = p.style
            while ali is None and st_ is not None:
                ali = st_.paragraph_format.alignment
                st_ = st_.base_style
            if ali != WD_ALIGN_PARAGRAPH.JUSTIFY:
                salah.append("tidak rata kiri-kanan")
            if salah:
                tambah("daftar_format", f"Isi daftar: \"{p.text.strip()[:40]}\"", ", ".join(salah),
                       f"{f_rules.get('nama') or ''} {f_rules.get('ukuran_isi_pt') or ''} spasi {spasi_t or ''}".strip(),
                       "otomatis")

    # -- butir/poin: penomoran otomatis --
    if (t.get("butir") or {}).get("otomatis"):
        for p, jenis in _butir_manual(doc, zona):
            total += 1
            tambah("butir_penomoran", f"Butir: \"{p.text.strip()[:40]}\"", "nomor/poin diketik manual",
                   "penomoran otomatis", "otomatis")

    # -- format keterangan & isi tabel --
    kf = (t.get("keterangan") or {}).get("format")
    if kf:
        for e in badan:
            if e["jenis"] != "p" or not (KAP_TABEL.match(e["teks"]) or KAP_GAMBAR.match(e["teks"])):
                continue
            if zona[e["idx"]] != "keterangan":
                continue
            total += 1
            p = e["p"]
            salah = []
            if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                salah.append("tidak rata tengah")
            nm = P._resolve_font_name(p, document=doc) or "Times New Roman"
            uk = P._resolve_font_size_pt(p, document=doc)
            if f_rules.get("nama") and nm != f_rules["nama"]:
                salah.append(nm)
            if f_rules.get("ukuran_isi_pt") and uk is not None and abs(uk - f_rules["ukuran_isi_pt"]) > 0.1:
                salah.append(f"{uk:g} pt")
            if salah:
                tambah("keterangan_format", f"\"{e['teks'][:40]}\"", ", ".join(salah),
                       "rata tengah, font ikut isi", "otomatis")
    if (t.get("tabel") or {}).get("font") and f_rules.get("ukuran_isi_pt"):
        pakai = _lebar_tersedia_cm(doc) * 567
        for n, tbl in enumerate(_tabel_dalam_bab(doc, badan, zona), 1):
            total += 1
            grid = _grid_twips(tbl)
            target = _ukuran_tabel(len(grid) or len(tbl.columns), f_rules["ukuran_isi_pt"])
            ukuran = _tabel_ukuran_run(tbl)
            salah = []
            if ukuran and ukuran != {target}:
                salah.append("ukuran font " + "/".join("bawaan" if u is None else f"{u:g}" for u in sorted(ukuran, key=lambda x: x or 0)))
            if sum(grid) > pakai + 20:
                salah.append("tabel melebihi lebar halaman")
            if salah:
                tambah("tabel_isi", f"Tabel ke-{n}", ", ".join(salah),
                       f"font {target:g} pt, muat dalam halaman", "otomatis")

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
    if dp.get("justify"):
        for i, p in entri:
            total += 1
            if p.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                tambah("daftar_pustaka_rata", f"Pustaka: \"{p.text.strip()[:40]}\"", "tidak rata kiri-kanan",
                       "justify", "otomatis")
    if dp.get("urut_abjad") and len(entri) > 1:
        total += 1
        kunci = [re.sub(r"[^a-z0-9]", "", p.text.lower()) for _, p in entri]
        for a in range(1, len(kunci)):
            if kunci[a] < kunci[a - 1]:
                tambah("daftar_pustaka_urut", f"Pustaka: \"{entri[a][1].text.strip()[:40]}\"",
                       "tidak urut abjad", "urut A-Z", "manual", severity="perlu_verifikasi")
                break
    return v, total


def _style_tebal(p):
    st = p.style
    while st is not None:
        if st.font.bold is not None:
            return bool(st.font.bold)
        st = st.base_style
    return False


# ---------- perbaikan ----------

_ID_BAWAAN = {"toc 1": "TOC1", "toc 2": "TOC2", "toc 3": "TOC3", "table of figures": "TableofFigures"}


def _cari_style(doc, nama):
    for st in doc.styles:
        if (st.name or "").lower() == nama.lower() and st.type == WD_STYLE_TYPE.PARAGRAPH:
            return st
    return None


def _style(doc, nama, dasar="Normal"):
    st = _cari_style(doc, nama)
    if st is not None:
        return st
    bawaan = nama.lower() in _ID_BAWAAN
    st = doc.styles.add_style(nama.lower() if bawaan else nama, WD_STYLE_TYPE.PARAGRAPH)
    if bawaan:
        st.style_id = _ID_BAWAAN[nama.lower()]
        if nama.lower() in ("toc 2", "toc 3"):
            st.paragraph_format.left_indent = Cm(0.5 if nama.lower() == "toc 2" else 1.0)
    try:
        st.base_style = doc.styles[dasar]
    except KeyError:
        pass
    return st


def _atur_font_style(st, nama, ukuran, tebal=None):
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if nama:
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rf.get(qn("w:" + a)) is not None:
                del rf.attrib[qn("w:" + a)]
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            rf.set(qn("w:" + a), nama)
    if ukuran:
        st.font.size = Pt(ukuran)
    if tebal is not None:
        st.font.bold = tebal
    st.font.italic = False
    st.font.color.rgb = RGBColor(0, 0, 0)
    c = rpr.find(qn("w:color"))
    if c is not None:
        for a in ("themeColor", "themeShade", "themeTint"):
            if c.get(qn("w:" + a)) is not None:
                del c.attrib[qn("w:" + a)]


def _atur_font_run(r, nama, ukuran):
    if nama:
        r.font.name = nama
        rf = r._r.get_or_add_rPr().find(qn("w:rFonts"))
        for a in ("eastAsia", "cs"):
            rf.set(qn("w:" + a), nama)
        for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rf.get(qn("w:" + a)) is not None:
                del rf.attrib[qn("w:" + a)]
    if ukuran:
        r.font.size = Pt(ukuran)


def _numbering_judul(doc):
    """numId penomoran multilevel milik MakaRapi (dibuat sekali): BAB I / 1.1 / 1.1.1"""
    num_el = doc.part.numbering_part.element
    ada = num_el.xpath('w:abstractNum[w:nsid/@w:val="4D414B41"]/@w:abstractNumId')
    if ada:
        n = num_el.xpath('w:num[w:abstractNumId/@w:val="%s"]/@w:numId' % ada[0])
        if n:
            return int(n[0])
    ids = [int(x) for x in num_el.xpath("w:abstractNum/@w:abstractNumId")] or [0]
    nids = [int(x) for x in num_el.xpath("w:num/@w:numId")] or [0]
    aid, nid = max(ids) + 1, max(nids) + 1
    xml = (
        '<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="%d">'
        '<w:nsid w:val="4D414B41"/><w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:suff w:val="nothing"/>'
        '<w:lvlText w:val="BAB %%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="%%1.%%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="%%1.%%2.%%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="851" w:hanging="851"/></w:pPr></w:lvl>'
        '</w:abstractNum>' % aid
    )
    from docx.oxml import parse_xml
    abs_el = parse_xml(xml)
    pertama_num = num_el.find(qn("w:num"))
    if pertama_num is not None:
        pertama_num.addprevious(abs_el)
    else:
        num_el.append(abs_el)
    num = parse_xml('<w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="%d">'
                    '<w:abstractNumId w:val="%d"/></w:num>' % (nid, aid))
    num_el.append(num)
    return nid


def _hapus_awalan(p, pola):
    """Buang awalan (mis. 'BAB I ' / '1.1 ') dari awal teks paragraf, lintas run, termasuk <w:br/>."""
    teks = p.text
    m = pola.match(teks)
    if not m or m.end() == 0:
        return False
    sisa = m.end()
    for r in p.runs:
        for ch in list(r._r):
            if sisa <= 0:
                return True
            if ch.tag == qn("w:t"):
                tx = ch.text or ""
                if len(tx) <= sisa:
                    sisa -= len(tx)
                    r._r.remove(ch)
                else:
                    ch.text = tx[sisa:]
                    ch.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    sisa = 0
            elif ch.tag in (qn("w:br"), qn("w:tab")):
                sisa -= 1
                r._r.remove(ch)
    return True


def _gabung_judul_berikut(p):
    """'BAB I' kosong setelah awalan dibuang -> tarik judul di paragraf berikutnya ke sini."""
    if p.text.strip():
        return
    nxt = p._p.getnext()
    if nxt is None or nxt.tag != qn("w:p"):
        return
    from docx.text.paragraph import Paragraph
    np_ = Paragraph(nxt, p._parent)
    teks = np_.text.strip()
    if not teks or len(teks) > 100:
        return
    for r in list(nxt.findall(qn("w:r"))):
        p._p.append(r)
    nxt.getparent().remove(nxt)


def _terapkan_judul(doc, zona, rules, t):
    jb, js, jd = t.get("judul_bab") or {}, t.get("judul_subbab") or {}, t.get("judul") or {}
    f = rules.get("font") or {}
    nama, ukuran = f.get("nama"), f.get("ukuran_isi_pt")
    judul = _daftar_judul(doc, zona)
    bg = t.get("bagian") or {}
    spasi = (rules.get("spasi") or {}).get("baris_isi")
    pakai_heading = bool(jd.get("heading")) or ("daftar_isi" in (t.get("bagian_wajib") or [])) or bool(bg.get("judul_tengah"))
    nid = None
    if jd.get("penomoran") and any(P.numpr_efektif(p) is None for _, p, _ in judul):
        nid = _numbering_judul(doc)
    if jd.get("font_isi") or pakai_heading:
        for lv in (1, 2, 3):
            st = _style(doc, f"Heading {lv}")
            if jd.get("font_isi") or (bg.get("judul_tengah") and lv == 1):
                _atur_font_style(st, nama, ukuran, True if (lv == 1 and (jb.get("tebal") or bg.get("judul_tengah"))) or (lv > 1 and js.get("tebal")) else None)
            if jd.get("font_isi") and spasi:
                st.paragraph_format.line_spacing = spasi
            ppr = st.element.get_or_add_pPr()
            if ppr.find(qn("w:outlineLvl")) is None:
                ol = OxmlElement("w:outlineLvl")
                ol.set(qn("w:val"), str(lv - 1))
                ppr.append(ol)
    for i, p, lv in judul:
        teks = p.text.strip()
        if jd.get("penomoran") and P.numpr_efektif(p) is None:
            _hapus_awalan(p, AWALAN_BAB if lv == 1 else AWALAN_SUB)
            if lv == 1:
                _gabung_judul_berikut(p)
            ppr = p._p.get_or_add_pPr()
            for old in ppr.findall(qn("w:numPr")):
                ppr.remove(old)
            numpr = OxmlElement("w:numPr")
            il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(lv - 1))
            ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(nid))
            numpr.append(il); numpr.append(ni)
            # urutan skema: numPr setelah keepNext/keepLines/pageBreakBefore ... sebelum spacing/ind/jc
            sisip = None
            for tag in ("w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                        "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                        "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                        "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc"):
                sisip = ppr.find(qn(tag))
                if sisip is not None:
                    break
            if sisip is not None:
                sisip.addprevious(numpr)
            else:
                ppr.append(numpr)
            if lv > 1:   # indentasi diatur level penomoran, buang indentasi langsung
                pf = p.paragraph_format
                pf.left_indent = None
                pf.first_line_indent = None
        if pakai_heading:
            p.style = _style(doc, f"Heading {lv}")
        if jd.get("font_isi"):
            for r in p.runs:
                _atur_font_run(r, nama, ukuran)
            if spasi:
                p.paragraph_format.line_spacing = spasi
        if lv == 1:
            if jb.get("kapital"):
                for r in p.runs:
                    r.text = r.text.upper()
            if jb.get("tebal"):
                for r in p.runs:
                    r.bold = True
            if jb.get("rata_tengah"):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif js.get("tebal"):
            for r in p.runs:
                r.bold = True


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


def _fld(p, tipe):
    r = p.add_run()
    el = OxmlElement("w:fldChar")
    el.set(qn("w:fldCharType"), tipe)
    r._r.append(el)


def _bangun_daftar(doc, sebelum_el, judul_teks, instr, entri, nama, ukuran):
    """Sisipkan judul bagian + bidang TOC (berisi entri awal; Word memperbaruinya saat dibuka)."""
    h = doc.add_paragraph()
    r = h.add_run(judul_teks)
    r.bold = True
    _atur_font_run(r, nama, ukuran)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.page_break_before = True
    sebelum_el.addprevious(h._p)
    teks_awal = entri or ["(Klik kanan di sini, lalu pilih Perbarui Bidang)"]
    paras = []
    for k, tx in enumerate(teks_awal):
        lv = 1
        if isinstance(tx, tuple):
            lv, tx = tx
        p = doc.add_paragraph(style=_style(doc, f"toc {lv}"))
        if k == 0:
            _fld(p, "begin")
            ri = p.add_run()
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = instr
            ri._r.append(it)
            _fld(p, "separate")
        rr = p.add_run(tx)
        _atur_font_run(rr, nama, ukuran)
        sebelum_el.addprevious(p._p)
        paras.append(p)
    _fld(paras[-1], "end")


def _aktifkan_update_fields(doc):
    st = doc.settings.element
    if st.find(qn("w:updateFields")) is not None:
        return
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    for tag in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids",
                "m:mathPr", "w:themeFontLang", "w:clrSchemeMapping"):
        try:
            x = st.find(qn(tag))
        except Exception:
            x = None
        if x is not None:
            x.addprevious(uf)
            return
    st.append(uf)


def _buat_daftar(doc, rules, t):
    wajib = set(t.get("bagian_wajib") or [])
    if not (wajib & set(DIBUAT_OTOMATIS)):
        return
    zona = P.classify_zones_untuk_italic(doc)
    badan = _urutan_badan(doc)
    bab = next((e for e in badan if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    if bab is None:
        return
    judul_pendek = [e["teks"] for e in badan if e["jenis"] == "p" and 0 < len(e["teks"]) <= 60]
    kt, kg, nt, ng = _keterangan_dokumen(badan, zona)
    f = rules.get("font") or {}
    nama, ukuran = f.get("nama"), f.get("ukuran_isi_pt")
    ada = lambda k: any(BAGIAN[k][1].match(x) for x in judul_pendek)
    dibuat = False

    bagian_h1 = bool((t.get("bagian") or {}).get("judul_tengah"))
    akan_tabel = "daftar_tabel" in wajib and (kt or nt) and not ada("daftar_tabel")
    akan_gambar = "daftar_gambar" in wajib and (kg or ng) and not ada("daftar_gambar")
    if "daftar_isi" in wajib and not ada("daftar_isi"):
        entri, cn = [], [0, 0, 0]
        pos_bab = bab["idx"]
        ekstra = []
        if bagian_h1:
            for e in _judul_bagian(badan, zona):
                ekstra.append((e["idx"], e["teks"].upper()))
            ekstra.append((pos_bab - 0.5, "DAFTAR ISI"))
            if akan_tabel:
                ekstra.append((pos_bab - 0.4, "DAFTAR TABEL"))
            if akan_gambar:
                ekstra.append((pos_bab - 0.3, "DAFTAR GAMBAR"))
        pakai_nomor = bool((t.get("judul") or {}).get("penomoran"))
        semua = []
        for i, p, lv in _daftar_judul(doc, zona):
            tx = p.text.strip()
            label = ""
            if pakai_nomor and P.numpr_efektif(p) is not None:
                cn[lv - 1] += 1
                for j in range(lv, 3):
                    cn[j] = 0
                label = (f"BAB {_romawi(cn[0])} " if lv == 1 else ".".join(str(c) for c in cn[:lv]) + " ")
            semua.append((i, (lv, label + tx.replace("\n", " "))))
        semua += [(i, (1, tx)) for i, tx in ekstra]
        entri = [x for _, x in sorted(semua, key=lambda z: z[0])]
        _bangun_daftar(doc, bab["el"], "DAFTAR ISI", ' TOC \\o "1-3" \\h \\z \\u ', entri, nama, ukuran)
        dibuat = True
    for kunci, daftar, ada_n, nm_style, judul_t in (
            ("daftar_tabel", kt, nt, "Keterangan Tabel", "DAFTAR TABEL"),
            ("daftar_gambar", kg, ng, "Keterangan Gambar", "DAFTAR GAMBAR")):
        if kunci in wajib and (daftar or ada_n) and not ada(kunci):
            st = _style(doc, nm_style)
            for e in daftar:
                e["p"].style = st
            _bangun_daftar(doc, bab["el"], judul_t, f' TOC \\h \\z \\t "{nm_style},1" ',
                           [e["teks"] for e in daftar], nama, ukuran)
            dibuat = True
    if dibuat:
        _aktifkan_update_fields(doc)


def _atur_halaman_baru(doc, t):
    hb = t.get("halaman_baru") or {}
    if not (hb.get("bab") or hb.get("bagian")):
        return
    zona = P.classify_zones_untuk_italic(doc)
    badan = _urutan_badan(doc)
    for k, e in enumerate(badan):
        if e["jenis"] != "p" or not e["teks"]:
            continue
        adalah_bab = zona[e["idx"]] == "judul_bab"
        adalah_bagian = (not adalah_bab and len(e["teks"]) <= 60 and not P.TOC_PATTERN.search(e["teks"])
                         and not P._nama_style(e["p"]).startswith(TOC_STYLE)
                             and not P._nama_style(e["p"]).startswith(TOC_STYLE)
                         and any(pola.match(e["teks"]) for _, pola in BAGIAN.values()))
        if (adalah_bab and hb.get("bab")) or (adalah_bagian and hb.get("bagian")):
            if not _awal_halaman_baru(badan, k):
                e["p"].paragraph_format.page_break_before = True


def perbaiki(doc, rules):
    t = rules.get("tambahan") or {}
    if not t:
        return
    zona = P.classify_zones_untuk_italic(doc)
    sec_map = _section_per_paragraf(doc)
    _terapkan_judul(doc, zona, rules, t)

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
                from docx.text.paragraph import Paragraph
                for el in asal._element.iter(qn("w:p")):
                    if _punya_page(el):
                        Paragraph(el, asal).alignment = _RATA[rata]

    dp = t.get("daftar_pustaka") or {}
    if dp.get("justify"):
        zona = P.classify_zones_untuk_italic(doc)
        for i, p in enumerate(doc.paragraphs):
            if zona[i] == "daftar_pustaka" and p.text.strip():
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if dp.get("hanging_cm") is not None:
        h = float(dp["hanging_cm"])
        zona = P.classify_zones_untuk_italic(doc)
        for i, p in enumerate(doc.paragraphs):
            if zona[i] == "daftar_pustaka" and p.text.strip():
                p.paragraph_format.left_indent = Cm(h)
                p.paragraph_format.first_line_indent = Cm(-h)

    _buat_daftar(doc, rules, t)
    _perbaiki_lain(doc, rules, t)
    _atur_halaman_baru(doc, t)


def _pindah_keterangan(doc, t):
    """Keterangan gambar di bawah gambar, keterangan tabel di atas tabel (sesuai aturan)."""
    ket = t.get("keterangan") or {}
    if not (ket.get("tabel_posisi") or ket.get("gambar_posisi")):
        return
    zona = P.classify_zones_untuk_italic(doc)
    badan = _urutan_badan(doc)
    mulai = next((k for k, e in enumerate(badan) if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    if mulai is None:
        return
    for k in range(mulai, len(badan)):
        e = badan[k]
        if e["jenis"] == "tbl":
            pola, aturan = KAP_TABEL, ket.get("tabel_posisi")
        elif e["jenis"] == "p" and e["gambar"]:
            pola, aturan = KAP_GAMBAR, ket.get("gambar_posisi")
        else:
            continue
        if not aturan:
            continue
        atas, bawah = _teks_tetangga(badan, k, -1), _teks_tetangga(badan, k, +1)
        if pola.match(atas) and not pola.match(bawah) and aturan == "bawah":
            arah = -1
        elif pola.match(bawah) and not pola.match(atas) and aturan == "atas":
            arah = +1
        else:
            continue
        j = k + arah
        while 0 <= j < len(badan) and not (badan[j]["jenis"] == "p" and badan[j]["teks"]):
            j += arah
        cap = badan[j]["el"]
        if arah == -1:
            e["el"].addnext(cap)        # keterangan pindah ke bawah gambar
        else:
            e["el"].addprevious(cap)    # keterangan pindah ke atas tabel
        badan = _urutan_badan(doc)      # urutan berubah; hitung ulang
        for k2, e2 in enumerate(badan):
            if e2["el"] is e["el"]:
                k = k2
                break


def _perbaiki_lain(doc, rules, t):
    _pindah_keterangan(doc, t)
    f = rules.get("font") or {}
    nama, ukuran = f.get("nama"), f.get("ukuran_isi_pt")
    spasi = (rules.get("spasi") or {}).get("baris_isi")
    zona = P.classify_zones_untuk_italic(doc)
    badan = _urutan_badan(doc)

    # butir/poin -> penomoran otomatis (kelompokkan butir berurutan sejenis)
    if (t.get("butir") or {}).get("otomatis"):
        manual = {id(p._p): (p, j) for p, j in _butir_manual(doc, zona)}
        kelompok_jenis, nid, sebelumnya = None, None, None
        for e in badan:
            if e["jenis"] != "p" or id(e["p"]._p) not in manual:
                if e["jenis"] == "tbl" or e.get("teks") or e.get("gambar"):
                    sebelumnya = None      # konten lain memutus kelompok butir
                continue
            p, jenis = manual[id(e["p"]._p)]
            if sebelumnya != jenis:
                nid = _num_baru(doc, _numbering_butir(doc, jenis))
            sebelumnya = jenis
            pola = dict(PENANDA_BUTIR)[jenis]
            _hapus_awalan(p, pola)
            _pasang_numpr(p, nid)
            p.paragraph_format.left_indent = None
            p.paragraph_format.first_line_indent = None
            for r in p.runs:
                _atur_font_run(r, nama, ukuran)
            _font_tanda_paragraf(p, nama, ukuran)
            if spasi:
                p.paragraph_format.line_spacing = spasi
            rata = (rules.get("perataan") or {}).get("jenis")
            if rata == "justify":
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # judul BAB: nomor di atas, judul di bawah
    if (t.get("judul_bab") or {}).get("baris_baru"):
        for i, p, lv in _daftar_judul(doc, zona):
            if lv == 1 and not _bab_baris_baru_ok(p):
                _paksa_baris_baru(p)

    # judul bagian & isi daftar
    bg = t.get("bagian") or {}
    if bg.get("judul_tengah"):
        for e in _judul_bagian(badan, zona):
            _format_bagian_judul(doc, e["p"], nama, ukuran, spasi)
    if bg.get("isi_daftar"):
        _format_daftar(doc, rules)

    # keterangan & isi tabel
    if (t.get("keterangan") or {}).get("format"):
        for e in badan:
            if e["jenis"] != "p" or zona[e["idx"]] != "keterangan":
                continue
            p = e["p"]
            for r in p.runs:
                _atur_font_run(r, nama, ukuran)
            if KAP_TABEL.match(e["teks"]) or KAP_GAMBAR.match(e["teks"]):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.keep_with_next = bool(KAP_TABEL.match(e["teks"]))   # keterangan tabel di atas tabel
                if KAP_GAMBAR.match(e["teks"]):                                      # gambar menempel ke keterangannya
                    sebelum = p._p.getprevious()
                    while sebelum is not None and sebelum.tag == qn("w:p") and not "".join(sebelum.itertext()).strip() and not sebelum.xpath(".//w:drawing"):
                        sebelum = sebelum.getprevious()
                    if sebelum is not None and sebelum.tag == qn("w:p") and sebelum.xpath(".//w:drawing"):
                        from docx.text.paragraph import Paragraph
                        Paragraph(sebelum, p._parent).paragraph_format.keep_with_next = True
            if spasi:
                p.paragraph_format.line_spacing = spasi
    if (t.get("tabel") or {}).get("font") and ukuran:
        for tbl in _tabel_dalam_bab(doc, badan, zona):
            _format_isi_tabel(doc, tbl, nama, ukuran)
