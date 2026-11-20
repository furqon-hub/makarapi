"""
ekstrak.py - baca dokumen CONTOH yang sudah benar dan hasilkan isian daftar centang Kustom.

Hanya mencatat apa yang benar-benar terbaca. Tiap isian punya alasan singkat di `catatan`,
dan yang tidak terbaca dibiarkan kosong (tidak dicentang).
"""
import re
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH

import parser as P
import ext_checks as X
import kustom

_ALIGN_NAMA = {WD_ALIGN_PARAGRAPH.JUSTIFY: "justify", WD_ALIGN_PARAGRAPH.LEFT: "left",
               WD_ALIGN_PARAGRAPH.CENTER: "center", WD_ALIGN_PARAGRAPH.RIGHT: "right"}


def _umum(nilai, minimal=1):
    nilai = [n for n in nilai if n is not None]
    if not nilai:
        return None, 0
    n, c = Counter(nilai).most_common(1)[0]
    return (n, c) if c >= minimal else (None, 0)


def _rata_efektif(p):
    a = p.paragraph_format.alignment
    st = p.style
    while a is None and st is not None:
        a = st.paragraph_format.alignment
        st = st.base_style
    return _ALIGN_NAMA.get(a, "left") if a is not None else "left"


def _snap(cm):
    """Bulatkan ke 0.5 cm bila sangat dekat (2.54 -> 2.5 bukan, tapi 3.99 -> 4.0)."""
    for k in (1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0):
        if abs(cm - k) <= 0.06:
            return k
    return round(cm, 2)


def ekstrak(doc):
    zona = P.classify_zones_untuk_italic(doc)
    info = P.get_paragraphs_info(doc)
    badan = X._urutan_badan(doc)
    sec_map = X._section_per_paragraf(doc)
    cat, j = [], {}

    # --- halaman: section yang memuat bab pertama ---
    idx_bab = next((i for i, z in enumerate(zona) if z == "judul_bab"), None)
    ps = P.get_page_setup(doc)
    s = ps[sec_map[idx_bab]] if idx_bab is not None and sec_map[idx_bab] < len(ps) else ps[-1]
    h = {}
    for k_json, k_ps in (("atas_cm", "margin_atas_cm"), ("bawah_cm", "margin_bawah_cm"),
                         ("kiri_cm", "margin_kiri_cm"), ("kanan_cm", "margin_kanan_cm")):
        if s.get(k_ps) is not None:
            h[k_json] = _snap(s[k_ps])
    for nama, (w, t) in kustom.KERTAS.items():
        if abs(s["lebar_cm"] - w) <= 0.2 and abs(s["tinggi_cm"] - t) <= 0.2:
            h["kertas"] = nama
            break
    j["halaman"] = h
    cat.append(f"Halaman dibaca dari section yang memuat BAB pertama: margin {h.get('atas_cm')}/{h.get('bawah_cm')}/{h.get('kiri_cm')}/{h.get('kanan_cm')} cm.")

    # --- teks isi: paragraf kalimat/isi ---
    isi = [x for x in info if x["zona"] in ("kalimat_awal", "isi")]
    paras = {i: p for i, p in enumerate(doc.paragraphs)}
    t = {}
    nm, _ = _umum([x["font_name"] for x in isi])
    uk, _ = _umum([x["font_size_pt"] for x in isi])
    sp, _ = _umum([round(x["line_spacing"], 2) for x in isi if x["line_spacing"] and x["line_spacing_mode"] == "multiple"])
    ind, _ = _umum([round(x["first_line_indent_cm"], 1) for x in isi
                    if x["first_line_indent_cm"] and x["first_line_indent_cm"] > 0.2])
    rata, _ = _umum([_rata_efektif(paras[x["index"]]) for x in isi if x["index"] in paras])
    if nm: t["font_nama"] = nm
    if uk: t["font_ukuran_pt"] = uk
    if sp: t["spasi_isi"] = sp
    if ind: t["indentasi_cm"] = ind
    if rata: t["perataan"] = rata
    j["teks"] = t
    cat.append(f"Teks isi dari {len(isi)} paragraf: {nm or '?'} {uk or '?'} pt, spasi {sp or '?'}, indentasi {ind or '?'} cm, rata {rata or '?'}.")

    # --- judul ---
    babs = [p for i, p in enumerate(doc.paragraphs) if zona[i] == "judul_bab" and p.text.strip()]
    subs = [(p, lv) for i, p, lv in X._daftar_judul(doc, zona) if lv > 1]
    if babs:
        j["judul_bab"] = {
            "kapital": all(p.text.strip() == p.text.strip().upper() for p in babs) or None,
            "tebal": all(P._semua_tebal(p) or X._style_tebal(p) for p in babs) or None,
            "rata_tengah": all(p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER or _rata_efektif(p) == "center" for p in babs) or None,
            "baris_baru": all("\n" in p.text.strip() or p.text.strip().upper().startswith("BAB") and len(p.text.strip()) < 12
                              or (P.numpr_efektif(p) is not None and p.text.startswith("\n")) for p in babs) or None,
        }
        cat.append(f"Judul bab dari {len(babs)} bab.")
    if subs:
        j["judul_subbab"] = {"tebal": all(P._semua_tebal(p) or X._style_tebal(p) for p, _ in subs) or None}
    allh = [(p, 1) for p in babs] + subs
    if allh:
        j["judul"] = {
            "heading": all(P._nama_style(p).startswith(f"heading {lv}") for p, lv in allh) or None,
            "penomoran": all(P.numpr_efektif(p) is not None for p, _ in allh) or None,
        }
        if nm and uk:
            same = all((P._resolve_font_name(p, document=doc) or "Times New Roman") == nm and
                       (P._resolve_font_size_pt(p, document=doc) in (None, uk)) for p, _ in allh)
            j["judul"]["font_isi"] = same or None

    # --- nomor halaman ---
    if idx_bab is not None:
        for jenis_hf in ("footer", "header"):
            ada, rata_ada, _ = X._halaman_info(doc, sec_map[idx_bab], jenis_hf)
            if ada:
                kunci = {("footer", WD_ALIGN_PARAGRAPH.CENTER): "bawah_tengah", ("footer", WD_ALIGN_PARAGRAPH.RIGHT): "bawah_kanan",
                         ("header", WD_ALIGN_PARAGRAPH.RIGHT): "atas_kanan", ("header", WD_ALIGN_PARAGRAPH.CENTER): "atas_tengah"}.get((jenis_hf, rata_ada))
                if kunci:
                    j["nomor_halaman"] = {"posisi": kunci}
                    cat.append(f"Nomor halaman di {kunci.replace('_', ' ')}.")
                break
        if len(doc.sections) > 1 and not X._halaman_info(doc, 0, "footer")[0]:
            j.setdefault("nomor_halaman", {})["sampul_tanpa_nomor"] = True

    # --- halaman baru ---
    bagian = X._judul_bagian(badan, zona)
    hb = {}
    pos_bab = [k for k, e in enumerate(badan) if e["jenis"] == "p" and e["teks"] and zona[e["idx"]] == "judul_bab"]
    if pos_bab and all(X._awal_halaman_baru(badan, k) for k in pos_bab):
        hb["bab"] = True
    pos_bagian = [k for k, e in enumerate(badan) if any(e is b for b in bagian)]
    if pos_bagian and all(X._awal_halaman_baru(badan, k) for k in pos_bagian):
        hb["bagian"] = True
    if hb:
        j["halaman_baru"] = hb

    # --- judul & isi bagian awal ---
    if bagian:
        j["bagian"] = {"judul_tengah": all(e["p"].paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER and
                                           (P._semua_tebal(e["p"]) or X._style_tebal(e["p"])) and
                                           e["teks"] == e["teks"].upper() for e in bagian) or None}
    ada_toc = [p for p in doc.paragraphs if P._nama_style(p).startswith(X.TOC_STYLE) and p.text.strip()]
    if ada_toc and nm:
        j.setdefault("bagian", {})["isi_daftar"] = all(
            (P._resolve_font_name(p, document=doc) or "Times New Roman") == nm for p in ada_toc) or None

    # --- bagian wajib yang ada ---
    pendek = [e["teks"] for e in badan if e["jenis"] == "p" and 0 < len(e["teks"]) <= 60]
    j["bagian_wajib"] = [k for k, (_, pola) in X.BAGIAN.items() if any(pola.match(x) for x in pendek)]
    cat.append("Bagian yang ditemukan: " + (", ".join(X.BAGIAN[k][0] for k in j["bagian_wajib"]) or "tidak ada") + ".")

    # --- butir otomatis ---
    butir = [p for i, p in enumerate(doc.paragraphs) if zona[i] == "butir_daftar" and p.text.strip()]
    if butir and all(P.numpr_efektif(p) is not None for p in butir):
        j["butir"] = {"otomatis": True}

    # --- tabel & gambar ---
    kt, kg, nt, ng = X._keterangan_dokumen(badan, zona)
    ket = {}
    mulai = next((k for k, e in enumerate(badan) if e["jenis"] == "p" and zona[e["idx"]] == "judul_bab"), None)
    if mulai is not None:
        pos = {"Tabel": Counter(), "Gambar": Counter()}
        for k in range(mulai, len(badan)):
            e = badan[k]
            if e["jenis"] == "tbl":
                tipe, pola = "Tabel", X.KAP_TABEL
            elif e["jenis"] == "p" and e["gambar"]:
                tipe, pola = "Gambar", X.KAP_GAMBAR
            else:
                continue
            if pola.match(X._teks_tetangga(badan, k, -1)):
                pos[tipe]["atas"] += 1
            elif pola.match(X._teks_tetangga(badan, k, +1)):
                pos[tipe]["bawah"] += 1
        for tipe, kunci in (("Tabel", "tabel_posisi"), ("Gambar", "gambar_posisi")):
            if pos[tipe]:
                ket[kunci] = pos[tipe].most_common(1)[0][0]
        if ket:
            j["keterangan"] = ket
            cat.append(f"Keterangan: tabel di {ket.get('tabel_posisi', '?')}, gambar di {ket.get('gambar_posisi', '?')}.")

    # --- daftar pustaka ---
    ent = [p for i, p in enumerate(doc.paragraphs) if zona[i] == "daftar_pustaka" and p.text.strip()]
    if ent:
        gantung = [round(-p.paragraph_format.first_line_indent.cm, 1) for p in ent
                   if p.paragraph_format.first_line_indent is not None and p.paragraph_format.first_line_indent.cm < -0.2]
        dp = {}
        g, _ = _umum(gantung, minimal=max(1, len(ent) // 2))
        if g: dp["hanging_cm"] = g
        if all(_rata_efektif(p) == "justify" for p in ent): dp["justify"] = True
        keys = [re.sub(r"[^a-z0-9]", "", p.text.lower()) for p in ent]
        if keys == sorted(keys): dp["urut_abjad"] = True
        if dp: j["daftar_pustaka"] = dp

    # buang kunci kosong
    bersih = {}
    for k, v in j.items():
        if isinstance(v, dict):
            v = {a: b for a, b in v.items() if b not in (None, False, "")}
            if not v:
                continue
        elif isinstance(v, list) and not v:
            continue
        bersih[k] = v
    return bersih, cat
