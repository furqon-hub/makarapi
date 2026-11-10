// Alamat API MakaRapi.
//
// - Emulator Android: http://10.0.2.2:8000 (alias ke laptop)
// - HP sungguhan (satu Wi-Fi dengan laptop): http://<IP-laptop>:8000, mis. http://192.168.1.10:8000
//   (cek IP laptop dengan `ipconfig` di Windows)
// - Setelah API di-hosting: isi dengan alamat hosting (https://...)
//
// Bisa diatur lewat variabel lingkungan EXPO_PUBLIC_API_URL tanpa mengubah kode.
export const API_URL: string = process.env.EXPO_PUBLIC_API_URL ?? "http://10.0.2.2:8000";

export const MAKS_UKURAN_BYTE = 10 * 1024 * 1024; // sama dengan batas di API
export const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
