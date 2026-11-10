export type Profil = {
  id: string;
  nama: string;
  tervalidasi: boolean;
  keterangan: string;
};

export type Pelanggaran = {
  kategori: string;
  lokasi: string;
  ditemukan: string;
  seharusnya: string;
  severity?: string;
};

export type Kandidat = {
  paragraf_index: number;
  kata_kandidat: string[];
  konteks: string;
};

export type HasilCek = {
  nama_file: string;
  profil: string;
  skor_kepatuhan_persen: number;
  jumlah_pelanggaran: number;
  total_pemeriksaan: number | null;
  ringkasan_per_kategori: Record<string, number>;
  pelanggaran: Pelanggaran[];
  kandidat_istilah_asing: Kandidat[];
};

export type FileTerpilih = {
  uri: string;
  name: string;
  mimeType?: string;
  size?: number;
  /** Hanya di web: objek File bawaan browser dari pemilih file. */
  webFile?: Blob;
};

// Isian profil kustom, disimpan sebagai teks (dari kolom input) dan dikirim apa adanya.
export type Kustom = {
  atas_cm?: string;
  bawah_cm?: string;
  kiri_cm?: string;
  kanan_cm?: string;
  font_nama?: string;
  font_ukuran_pt?: string;
  spasi_isi?: string;
  indentasi_cm?: string;
  perataan?: string;
};
