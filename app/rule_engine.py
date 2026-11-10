"""
rule_engine.py
==============
Membandingkan hasil parsing (dari parser.py) dengan aturan di
rules.yaml, menghasilkan daftar pelanggaran.

KONSEP PENTING:
---------------
Rule engine ini SENGAJA dipisah dari parser. Parser hanya "membaca
fakta" dari dokumen, rule engine yang "menilai benar/salah". Kalau
nanti pedoman berubah, Anda cukup ubah rules.yaml -- kode di sini
dan di parser.py tidak perlu disentuh. Inilah yang jadi klaim
kebaruan di artikel: aturan tidak ditanam di kode (hard-coded).
"""

import yaml
from docx.enum.text import WD_ALIGN_PARAGRAPH

ALIGNMENT_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    None: None,
}


def load_rules(path="rules.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def check_page_setup(page_setup_list, rules):
    """Cek ukuran kertas & margin tiap section."""
    violations = []
    r = rules["ukuran_kertas"]
    m = rules["margin"]

    for sec in page_setup_list:
        loc = f"Section {sec['section_index']}"

        if sec["lebar_cm"] is not None and abs(sec["lebar_cm"] - r["lebar_cm"]) > 0.1:
            violations.append({
                "kategori": "ukuran_kertas",
                "lokasi": loc,
                "ditemukan": f"{sec['lebar_cm']} cm",
                "seharusnya": f"{r['lebar_cm']} cm",
                "severity": "deterministik",
            })

        for sisi, key in [("atas", "margin_atas_cm"), ("bawah", "margin_bawah_cm"),
                           ("kiri", "margin_kiri_cm"), ("kanan", "margin_kanan_cm")]:
            nilai = sec[key]
            target = m[f"{sisi}_cm"]
            if nilai is not None and abs(nilai - target) > m["toleransi_cm"]:
                violations.append({
                    "kategori": f"margin_{sisi}",
                    "lokasi": loc,
                    "ditemukan": f"{nilai} cm",
                    "seharusnya": f"{target} cm",
                    "severity": "deterministik",
                })
    return violations


def check_paragraphs(paragraphs_info, rules):
    """
    Cek font, ukuran font, spasi, indentasi, perataan per paragraf.

    PENTING -- hasil dari pengujian ke dokumen asli: tidak semua
    aturan "isi" berlaku untuk semua zona dokumen. Lihat penjelasan
    panjang di parser.py (_classify_zones). Di sini kita pakai zona
    itu untuk MEMBATASI aturan mana yang relevan:

      - perataan & indentasi "isi": HANYA diperiksa di zona "isi".
        Sampul (zona "awal") memang konvensinya rata tengah, judul
        bab ("judul_bab") memang konvensinya tanpa indentasi, dan
        Daftar Pustaka pakai hanging indent, bukan first-line indent.
      - jenis_font & ukuran_font: HANYA diperiksa di zona "isi" dan
        "daftar_pustaka". Sampul lazim pakai ukuran judul yang lebih
        besar -- itu konvensi, bukan kesalahan, jadi dikecualikan.
      - spasi_baris: HANYA diperiksa di zona "isi". Sampul, kata
        pengantar sebelum BAB I, dan Daftar Pustaka sering punya
        aturan spasi sendiri yang berbeda dari isi.
    """
    violations = []
    f_rules = rules["font"]
    s_rules = rules["spasi"]
    i_rules = rules["indentasi"]
    a_rules = rules["perataan"]

    for p in paragraphs_info:
        loc = f"Paragraf #{p['index']} (\"{p['teks_cuplikan']}...\")"
        zona = p.get("zona", "isi")

        cek_font = zona in ("isi", "daftar_pustaka")
        cek_perataan_indentasi_spasi = zona == "isi"

        if not cek_font:
            pass  # lewati cek font untuk zona "awal", "judul_bab", "judul_daftar_pustaka"
        elif p["font_name"] and p["font_name"] != f_rules["nama"]:
            violations.append({
                "kategori": "jenis_font",
                "lokasi": loc,
                "ditemukan": p["font_name"],
                "seharusnya": f_rules["nama"],
                "severity": "deterministik",
            })
        elif p["font_name"] is None:
            violations.append({
                "kategori": "jenis_font",
                "lokasi": loc,
                "ditemukan": "tidak terdeteksi (perlu cek manual)",
                "seharusnya": f_rules["nama"],
                "severity": "perlu_verifikasi",
            })

        if cek_font and p["font_size_pt"] is not None:
            if abs(p["font_size_pt"] - f_rules["ukuran_isi_pt"]) > 0.1:
                violations.append({
                    "kategori": "ukuran_font",
                    "lokasi": loc,
                    "ditemukan": f"{p['font_size_pt']} pt",
                    "seharusnya": f"{f_rules['ukuran_isi_pt']} pt",
                    "severity": "deterministik",
                })

        if cek_perataan_indentasi_spasi and p["line_spacing"] is not None:
            if abs(p["line_spacing"] - s_rules["baris_isi"]) > s_rules["toleransi"]:
                violations.append({
                    "kategori": "spasi_baris",
                    "lokasi": loc,
                    "ditemukan": f"{p['line_spacing']}",
                    "seharusnya": f"{s_rules['baris_isi']}",
                    "severity": "deterministik",
                })

        if cek_perataan_indentasi_spasi and p["first_line_indent_cm"] is not None:
            if abs(p["first_line_indent_cm"] - i_rules["awal_paragraf_cm"]) > i_rules["toleransi_cm"]:
                violations.append({
                    "kategori": "indentasi",
                    "lokasi": loc,
                    "ditemukan": f"{p['first_line_indent_cm']} cm",
                    "seharusnya": f"{i_rules['awal_paragraf_cm']} cm",
                    "severity": "deterministik",
                })

        if cek_perataan_indentasi_spasi:
            align_str = ALIGNMENT_MAP.get(p["alignment"], "tidak diketahui")
            if align_str is not None and align_str != a_rules["jenis"]:
                violations.append({
                    "kategori": "perataan",
                    "lokasi": loc,
                    "ditemukan": align_str,
                    "seharusnya": a_rules["jenis"],
                    "severity": "deterministik",
                })

    return violations


def compute_compliance_score(violations, total_checks):
    """
    Skor kepatuhan sederhana: persentase pemeriksaan yang LOLOS
    dari total pemeriksaan yang dilakukan.

    total_checks = jumlah total titik yang diperiksa (dihitung di
    luar fungsi ini, dari jumlah paragraf x jumlah aturan yang relevan).
    """
    if total_checks == 0:
        return 100.0
    lolos = total_checks - len(violations)
    return round(max(lolos, 0) / total_checks * 100, 1)


def run_all_checks(page_setup_list, paragraphs_info, rules):
    violations = []
    violations += check_page_setup(page_setup_list, rules)
    violations += check_paragraphs(paragraphs_info, rules)

    # Total titik periksa HARUS mengikuti zona yang benar-benar diperiksa
    # (lihat check_paragraphs) -- kalau tidak, skor akan salah hitung
    # karena menganggap paragraf sampul/judul bab juga diperiksa 5 aspek,
    # padahal sengaja dilewati.
    total_checks = 0
    for p in paragraphs_info:
        zona = p.get("zona", "isi")
        if zona == "isi":
            total_checks += 5  # jenis_font, ukuran_font, spasi, indentasi, perataan
        elif zona == "daftar_pustaka":
            total_checks += 2  # jenis_font, ukuran_font saja
        # zona "awal", "judul_bab", "judul_daftar_pustaka": tidak diperiksa sama sekali

    total_checks += len(page_setup_list) * 6

    score = compute_compliance_score(violations, total_checks)
    return {
        "violations": violations,
        "total_checks": total_checks,
        "jumlah_pelanggaran": len(violations),
        "skor_kepatuhan_persen": score,
    }


if __name__ == "__main__":
    import sys
    import json
    from docx import Document
    import parser as parser_mod

    if len(sys.argv) < 2:
        print("Pakai: python3 rule_engine.py <file.docx>")
        sys.exit(1)

    doc = Document(sys.argv[1])
    rules = load_rules("../rules.yaml")
    page_setup = parser_mod.get_page_setup(doc)
    paragraphs = parser_mod.get_paragraphs_info(doc)

    hasil = run_all_checks(page_setup, paragraphs, rules)
    print(json.dumps(hasil, indent=2, ensure_ascii=False))
