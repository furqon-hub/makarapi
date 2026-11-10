import { useEffect, useState } from "react";
import {
  ActivityIndicator, Alert, Pressable, ScrollView, StyleSheet, Text, TextInput, View,
} from "react-native";
import * as DocumentPicker from "expo-document-picker";
import { useRouter } from "expo-router";

import { ambilProfil, bangunkanServer, GalatApi, periksa } from "../lib/api";
import { DOCX_MIME, MAKS_UKURAN_BYTE } from "../lib/config";
import { simpanHasil } from "../lib/store";
import type { FileTerpilih, Kustom, Profil } from "../lib/types";

// Profil cadangan bila daftar dari server belum bisa diambil.
const PROFIL_CADANGAN: Profil[] = [
  { id: "febi", nama: "FEBI UIN Gusdur 2026", tervalidasi: true, keterangan: "Bersumber dari pedoman resmi." },
  { id: "kustom", nama: "Kustom (isi sendiri)", tervalidasi: false, keterangan: "Hasil bergantung pada nilai yang Anda isi." },
];

const KOLOM_KUSTOM: { kunci: keyof Kustom; label: string; contoh: string }[] = [
  { kunci: "atas_cm", label: "Margin atas (cm)", contoh: "4" },
  { kunci: "bawah_cm", label: "Margin bawah (cm)", contoh: "3" },
  { kunci: "kiri_cm", label: "Margin kiri (cm)", contoh: "4" },
  { kunci: "kanan_cm", label: "Margin kanan (cm)", contoh: "3" },
  { kunci: "font_nama", label: "Jenis font", contoh: "Times New Roman" },
  { kunci: "font_ukuran_pt", label: "Ukuran font (pt)", contoh: "12" },
  { kunci: "spasi_isi", label: "Spasi baris", contoh: "1.5" },
  { kunci: "indentasi_cm", label: "Indentasi paragraf (cm)", contoh: "1" },
];

export default function Unggah() {
  const router = useRouter();
  const [daftarProfil, setDaftarProfil] = useState<Profil[]>(PROFIL_CADANGAN);
  const [profil, setProfil] = useState("febi");
  const [kustom, setKustom] = useState<Kustom>({});
  const [file, setFile] = useState<FileTerpilih | null>(null);
  const [memuat, setMemuat] = useState(false);
  const [galat, setGalat] = useState<string | null>(null);

  useEffect(() => {
    // Bangunkan server hosting yang "tidur", lalu ambil daftar profil.
    bangunkanServer().then(() =>
      ambilProfil().then(setDaftarProfil).catch(() => { }),
    );
  }, []);

  async function pilihFile() {
    setGalat(null);
    const hasil = await DocumentPicker.getDocumentAsync({
      type: DOCX_MIME,
      copyToCacheDirectory: true,
      multiple: false,
    });
    if (hasil.canceled) return;
    const aset = hasil.assets[0];
    if (!aset.name.toLowerCase().endsWith(".docx")) {
      setFile(null);
      setGalat("Hanya file .docx yang didukung.");
      return;
    }
    if (aset.size !== undefined && aset.size > MAKS_UKURAN_BYTE) {
      setFile(null);
      setGalat("Ukuran file melebihi 10 MB.");
      return;
    }
    setFile({
      uri: aset.uri,
      name: aset.name,
      mimeType: aset.mimeType ?? DOCX_MIME,
      size: aset.size,
      webFile: aset.file,
    });
  }

  async function jalankan() {
    if (!file) return;
    setGalat(null);
    setMemuat(true);
    try {
      const hasil = await periksa(file, profil, kustom);
      simpanHasil(hasil, file, profil, kustom);
      router.push("/hasil");
    } catch (e) {
      const pesan = e instanceof GalatApi ? e.message : "Terjadi kesalahan tak terduga.";
      setGalat(pesan);
      if (e instanceof GalatApi && e.status === 0) {
        Alert.alert("Tidak terhubung", pesan);
      }
    } finally {
      setMemuat(false);
    }
  }

  const profilAktif = daftarProfil.find((p) => p.id === profil);

  return (
    <ScrollView contentContainerStyle={s.isi} keyboardShouldPersistTaps="handled">
      <Text style={s.judul}>Periksa format makalah</Text>
      <Text style={s.sub}>Pilih pedoman, unggah file .docx, lalu lihat skor kepatuhannya.</Text>

      <Text style={s.label}>Pedoman</Text>
      <View style={s.baris}>
        {daftarProfil.map((p) => (
          <Pressable
            key={p.id}
            onPress={() => setProfil(p.id)}
            style={[s.chip, profil === p.id && s.chipAktif]}
          >
            <Text style={[s.chipTeks, profil === p.id && s.chipTeksAktif]}>{p.nama}</Text>
          </Pressable>
        ))}
      </View>
      {profilAktif && (
        <Text style={s.catatan}>
          {profilAktif.keterangan}
          {!profilAktif.tervalidasi ? " Profil ini belum tervalidasi terhadap pedoman resmi." : ""}
        </Text>
      )}

      {profil === "kustom" && (
        <View style={s.kartu}>
          <Text style={s.kartuJudul}>Nilai pedoman Anda</Text>
          <Text style={s.catatan}>Kolom yang dikosongkan memakai nilai FEBI.</Text>
          {KOLOM_KUSTOM.map((k) => (
            <View key={k.kunci} style={s.kolom}>
              <Text style={s.kolomLabel}>{k.label}</Text>
              <TextInput
                style={s.input}
                value={kustom[k.kunci] ?? ""}
                placeholder={k.contoh}
                keyboardType={k.kunci === "font_nama" ? "default" : "decimal-pad"}
                onChangeText={(t) => setKustom((lama) => ({ ...lama, [k.kunci]: t }))}
              />
            </View>
          ))}
          <Text style={s.kolomLabel}>Perataan teks</Text>
          <View style={s.baris}>
            {[
              { id: "justify", nama: "Rata kanan-kiri" },
              { id: "left", nama: "Rata kiri" },
            ].map((o) => (
              <Pressable
                key={o.id}
                onPress={() => setKustom((lama) => ({ ...lama, perataan: o.id }))}
                style={[s.chip, (kustom.perataan ?? "justify") === o.id && s.chipAktif]}
              >
                <Text style={[s.chipTeks, (kustom.perataan ?? "justify") === o.id && s.chipTeksAktif]}>{o.nama}</Text>
              </Pressable>
            ))}
          </View>
        </View>
      )}

      <Text style={s.label}>Dokumen</Text>
      <Pressable style={s.tombolSekunder} onPress={pilihFile} disabled={memuat}>
        <Text style={s.tombolSekunderTeks}>Pilih file .docx</Text>
      </Pressable>
      {file && (
        <Text style={s.catatan}>
          {file.name}
          {file.size !== undefined ? ` (${(file.size / 1024 / 1024).toFixed(2)} MB)` : ""}
        </Text>
      )}

      {galat && <Text style={s.galat}>{galat}</Text>}

      <Pressable
        style={[s.tombol, (!file || memuat) && s.tombolMati]}
        onPress={jalankan}
        disabled={!file || memuat}
      >
        {memuat ? <ActivityIndicator color="#FFFFFF" /> : <Text style={s.tombolTeks}>Periksa Sekarang</Text>}
      </Pressable>
      {memuat && <Text style={s.catatan}>Memproses... permintaan pertama bisa lebih lama.</Text>}

      <Text style={s.catatan}>File tidak disimpan di server setelah diproses.</Text>
    </ScrollView>
  );
}

const s = StyleSheet.create({
  isi: { padding: 16, gap: 8 },
  judul: { fontSize: 22, fontWeight: "700", color: "#1F3864" },
  sub: { fontSize: 14, color: "#444", marginBottom: 8 },
  label: { fontSize: 14, fontWeight: "600", marginTop: 12, color: "#1F3864" },
  baris: { flexDirection: "row", flexWrap: "wrap", gap: 8 },
  chip: { borderWidth: 1, borderColor: "#1F3864", borderRadius: 20, paddingVertical: 8, paddingHorizontal: 14 },
  chipAktif: { backgroundColor: "#1F3864" },
  chipTeks: { color: "#1F3864" },
  chipTeksAktif: { color: "#FFFFFF" },
  catatan: { fontSize: 12, color: "#666" },
  kartu: { borderWidth: 1, borderColor: "#DDD", borderRadius: 10, padding: 12, gap: 6, marginTop: 8 },
  kartuJudul: { fontWeight: "600" },
  kolom: { gap: 2 },
  kolomLabel: { fontSize: 12, color: "#444", marginTop: 4 },
  input: { borderWidth: 1, borderColor: "#CCC", borderRadius: 8, paddingHorizontal: 10, paddingVertical: 8 },
  tombolSekunder: { borderWidth: 1, borderColor: "#1F3864", borderRadius: 10, padding: 14, alignItems: "center" },
  tombolSekunderTeks: { color: "#1F3864", fontWeight: "600" },
  tombol: { backgroundColor: "#1F3864", borderRadius: 10, padding: 16, alignItems: "center", marginTop: 12 },
  tombolMati: { opacity: 0.4 },
  tombolTeks: { color: "#FFFFFF", fontWeight: "700", fontSize: 16 },
  galat: { color: "#B00020", marginTop: 8 },
});
