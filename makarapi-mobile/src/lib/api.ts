import { Platform } from "react-native";
import { File as BerkasLokal } from "expo-file-system";

import { API_URL } from "./config";
import type { FileTerpilih, HasilCek, Kustom, Profil } from "./types";

export class GalatApi extends Error {
  status: number;
  constructor(status: number, pesan: string) {
    super(pesan);
    this.status = status;
  }
}

async function fetchDenganBatasWaktu(url: string, init: RequestInit, ms = 60000): Promise<Response> {
  const kontrol = new AbortController();
  const timer = setTimeout(() => kontrol.abort(), ms);
  try {
    return await fetch(url, { ...init, signal: kontrol.signal });
  } catch (e) {
    console.warn("[api] permintaan gagal:", url, e);
    const rinci = e instanceof Error ? e.message : "galat tidak dikenal";
    if (e instanceof Error && e.name === "AbortError") {
      throw new GalatApi(0, "Server tidak merespons (waktu habis). Coba lagi sebentar lagi.");
    }
    throw new GalatApi(0, `Permintaan gagal: ${rinci}. Periksa koneksi dan alamat server.`);
  } finally {
    clearTimeout(timer);
  }
}

async function pesanDariRespons(res: Response): Promise<string> {
  if (res.status === 429) return "Terlalu banyak permintaan. Tunggu sebentar lalu coba lagi.";
  try {
    const j = await res.json();
    if (typeof j?.detail === "string") return j.detail;
  } catch {
    // abaikan, pakai pesan umum
  }
  return `Terjadi kesalahan pada server (kode ${res.status}).`;
}

/** Membangunkan server hosting gratis yang "tidur"; gagal tidak perlu ditampilkan. */
export async function bangunkanServer(): Promise<void> {
  try {
    await fetchDenganBatasWaktu(`${API_URL}/health`, {}, 20000);
  } catch {
    // sengaja diabaikan
  }
}

export async function ambilProfil(): Promise<Profil[]> {
  const res = await fetchDenganBatasWaktu(`${API_URL}/profil`, {}, 20000);
  if (!res.ok) throw new GalatApi(res.status, await pesanDariRespons(res));
  return (await res.json()) as Profil[];
}

export function bentukFormData(file: FileTerpilih, profil: string, kustom: Kustom): FormData {
  const form = new FormData();
  // Expo SDK terbaru memakai fetch milik Expo yang TIDAK menerima objek {uri, name, type}
  // di FormData. File harus berupa Blob/File:
  //  - web: objek File bawaan browser dari pemilih file
  //  - Android/iOS: File dari expo-file-system (dibaca langsung dari uri)
  if (Platform.OS === "web") {
    if (!file.webFile) throw new GalatApi(0, "Berkas tidak terbaca di browser. Pilih ulang file.");
    form.append("file", file.webFile, file.name);
  } else {
    form.append("file", new BerkasLokal(file.uri) as unknown as Blob, file.name);
  }
  form.append("profil", profil);
  if (profil === "kustom") {
    (Object.keys(kustom) as (keyof Kustom)[]).forEach((k) => {
      const nilai = kustom[k]?.trim();
      if (nilai) form.append(k, nilai.replace(",", "."));
    });
  }
  return form;
}

export async function periksa(file: FileTerpilih, profil: string, kustom: Kustom): Promise<HasilCek> {
  const res = await fetchDenganBatasWaktu(`${API_URL}/check`, {
    method: "POST",
    body: bentukFormData(file, profil, kustom),
  });
  if (!res.ok) throw new GalatApi(res.status, await pesanDariRespons(res));
  return (await res.json()) as HasilCek;
}
