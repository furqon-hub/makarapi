import type { FileTerpilih, HasilCek, Kustom } from "./types";

// Penyimpanan sederhana di memori untuk meneruskan data antar layar.
// (Hasil terlalu besar untuk dikirim sebagai parameter rute.)
type Sesi = {
  hasil: HasilCek | null;
  file: FileTerpilih | null;
  profil: string;
  kustom: Kustom;
};

const sesi: Sesi = { hasil: null, file: null, profil: "febi", kustom: {} };

export const ambilSesi = (): Sesi => sesi;

export function simpanHasil(hasil: HasilCek, file: FileTerpilih, profil: string, kustom: Kustom) {
  sesi.hasil = hasil;
  sesi.file = file;
  sesi.profil = profil;
  sesi.kustom = kustom;
}
