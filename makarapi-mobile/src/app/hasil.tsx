import { Pressable, ScrollView, StyleSheet, Text, View } from "react-native";
import { useRouter } from "expo-router";

import { ambilSesi } from "../lib/store";

const NAMA_KATEGORI: Record<string, string> = {
  ukuran_kertas: "Ukuran kertas",
  margin_atas: "Margin atas",
  margin_bawah: "Margin bawah",
  margin_kiri: "Margin kiri",
  margin_kanan: "Margin kanan",
  ukuran_font: "Ukuran font",
  jenis_font: "Jenis font",
  spasi_baris: "Spasi baris",
  indentasi: "Indentasi paragraf",
  perataan: "Perataan teks",
};

export default function Hasil() {
  const router = useRouter();
  const { hasil, file } = ambilSesi();

  if (!hasil) {
    return (
      <View style={s.tengah}>
        <Text>Belum ada hasil. Kembali dan unggah dokumen terlebih dahulu.</Text>
      </View>
    );
  }

  const kategori = Object.entries(hasil.ringkasan_per_kategori).sort((a, b) => b[1] - a[1]);

  return (
    <ScrollView contentContainerStyle={s.isi}>
      <Text style={s.nama}>{file?.name ?? hasil.nama_file}</Text>

      <View style={s.skorKartu}>
        <Text style={s.skor}>{hasil.skor_kepatuhan_persen}%</Text>
        <Text style={s.skorLabel}>skor kepatuhan</Text>
        <Text style={s.jumlah}>
          {hasil.jumlah_pelanggaran} pelanggaran
          {hasil.total_pemeriksaan ? ` dari ${hasil.total_pemeriksaan} pemeriksaan` : ""}
        </Text>
      </View>

      <Text style={s.judulBagian}>Per kategori</Text>
      {kategori.length === 0 ? (
        <Text style={s.catatan}>Tidak ada pelanggaran format yang ditemukan.</Text>
      ) : (
        kategori.map(([k, n]) => (
          <View key={k} style={s.barisKategori}>
            <Text>{NAMA_KATEGORI[k] ?? k}</Text>
            <Text style={s.angka}>{n}</Text>
          </View>
        ))
      )}

      <Pressable style={s.tombol} onPress={() => router.push("/istilah")}>
        <Text style={s.tombolTeks}>
          Lihat istilah asing ({hasil.kandidat_istilah_asing.length})
        </Text>
      </Pressable>

      <Pressable style={[s.tombol, s.tombolMati]} disabled>
        <Text style={s.tombolTeks}>Unduh hasil perbaikan (segera)</Text>
      </Pressable>
    </ScrollView>
  );
}

const s = StyleSheet.create({
  isi: { padding: 16, gap: 8 },
  tengah: { flex: 1, alignItems: "center", justifyContent: "center", padding: 24 },
  nama: { fontSize: 14, color: "#444" },
  skorKartu: { backgroundColor: "#F2F5FA", borderRadius: 12, padding: 20, alignItems: "center", marginVertical: 8 },
  skor: { fontSize: 48, fontWeight: "700", color: "#1F3864" },
  skorLabel: { fontSize: 14, color: "#444" },
  jumlah: { marginTop: 6, fontSize: 14 },
  judulBagian: { fontWeight: "600", marginTop: 12, color: "#1F3864" },
  catatan: { color: "#666" },
  barisKategori: {
    flexDirection: "row", justifyContent: "space-between",
    paddingVertical: 8, borderBottomWidth: 1, borderBottomColor: "#EEE",
  },
  angka: { fontWeight: "700" },
  tombol: { backgroundColor: "#1F3864", borderRadius: 10, padding: 14, alignItems: "center", marginTop: 12 },
  tombolMati: { opacity: 0.4 },
  tombolTeks: { color: "#FFFFFF", fontWeight: "600" },
});
