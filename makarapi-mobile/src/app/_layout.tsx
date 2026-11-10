import { Stack } from "expo-router";
import { StatusBar } from "expo-status-bar";

export default function RootLayout() {
  return (
    <>
      <StatusBar style="light" />
      <Stack
        screenOptions={{
          headerStyle: { backgroundColor: "#1F3864" },
          headerTintColor: "#FFFFFF",
          headerTitleStyle: { fontWeight: "600" },
        }}
      >
        <Stack.Screen name="index" options={{ title: "MakaRapi" }} />
        <Stack.Screen name="hasil" options={{ title: "Hasil Pemeriksaan" }} />
        <Stack.Screen name="istilah" options={{ title: "Istilah Asing" }} />
      </Stack>
    </>
  );
}
