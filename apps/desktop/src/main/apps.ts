import { appForFileName } from "../shared/formats.js";

export type AppKind = "belge" | "tablo" | "sunum";
export type WindowKind = AppKind | "baslangic";

/** Uzantıya göre dosyayı açacak uygulama (biçim listesi: shared/formats.ts). */
export function appForPath(filePath: string): AppKind | null {
  return appForFileName(filePath);
}

export function isSupportedFile(filePath: string): boolean {
  return appForPath(filePath) !== null;
}

export const APP_TITLES: Record<WindowKind, string> = {
  belge: "OfficeHere Belge",
  tablo: "OfficeHere Tablo",
  sunum: "OfficeHere Sunum",
  baslangic: "OfficeHere",
};

/** Başlık çubuğu varsayılan renkleri ("Renkli" tema); sayfa açılınca güncellenir. */
export const APP_COLORS: Record<WindowKind, string> = {
  belge: "#1f5fbf",
  tablo: "#1d7a46",
  sunum: "#c4462b",
  baslangic: "#3d3d8f",
};

/** Geliştirme sırasında Vite sunucuları. */
export const DEV_PORTS: Record<WindowKind, number> = {
  belge: 5173,
  tablo: 5174,
  sunum: 5175,
  baslangic: 5176,
};
