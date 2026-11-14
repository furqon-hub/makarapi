# Profil pedoman per kampus

Satu berkas `.yaml` = satu profil. API membacanya otomatis (tidak perlu mengubah kode).
Salin berkas yang sudah ada, ganti isinya, simpan di folder ini, lalu push dan reload server.

Aturan penting:
- `id`: huruf kecil, angka, garis bawah (2-40 karakter), unik, tidak boleh `febi`, `febi_skripsi`, `kustom`.
- `aturan`: hanya menimpa nilai dasar FEBI (`rules.yaml`). **Aspek yang TIDAK diatur pedoman
  harus ditulis `null`** supaya tidak diperiksa dan tidak diubah (jangan dibiarkan kosong,
  karena nilai FEBI akan terpakai).
- `spasi.mode`: `multiple` (kelipatan: 1.0, 1.5, 2.0) atau `exact` (tepat N pt, mis. 24).
- Setiap nilai sebaiknya dicatat sumbernya di `lokasi_nilai` (bab/butir/halaman).
- `tervalidasi: false` sampai ada orang yang mencocokkan nilainya dengan dokumen aslinya.
