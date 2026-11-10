import { FlatList, StyleSheet, Text, View } from "react-native";

import { ambilSesi } from "../lib/store";

export default function Istilah() {
  const { hasil } = ambilSesi();
  const daftar = hasil?.kandidat_istilah_asing ?? [];

  return (
    <FlatList
      contentContainerStyle={s.isi}
      data={daftar}
      keyExtractor={(item, i) => `${item.paragraf_index}-${i}`}
      ListHeaderComponent={
        <Text style={s.catatan}>
          Saran saja: teks dokumen tidak diubah. Periksa apakah kata ini sudah baku di KBBI; jika belum, tulis dengan huruf miring.
        </Text>
      }
      ListEmptyComponent={<Text style={s.kosong}>Tidak ada kandidat istilah asing.</Text>}
      renderItem={({ item }) => (
        <View style={s.kartu}>
          <Text style={s.kata}>{item.kata_kandidat.join(", ")}</Text>
          <Text style={s.paragraf}>Paragraf {item.paragraf_index}</Text>
          <Text style={s.konteks}>{item.konteks}</Text>
        </View>
      )}
    />
  );
}

const s = StyleSheet.create({
  isi: { padding: 16, gap: 10 },
  catatan: { color: "#666", fontSize: 12, marginBottom: 8 },
  kosong: { textAlign: "center", marginTop: 24, color: "#444" },
  kartu: { borderWidth: 1, borderColor: "#DDD", borderRadius: 10, padding: 12, gap: 2 },
  kata: { fontWeight: "700", color: "#1F3864" },
  paragraf: { fontSize: 12, color: "#666" },
  konteks: { marginTop: 4 },
});
