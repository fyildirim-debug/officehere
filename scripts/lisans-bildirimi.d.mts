import type { Plugin } from "vite";

export interface LisansBildirimiSecenekleri {
  /** Yapılandırma dosyasına göre göreli paket dizini; üretim bağımlılıkları da bildirime eklenir. */
  uretimBagimliliklari?: string;
}

/** Vite eklentisi: derleme çıktısına THIRD_PARTY_NOTICES.txt ekler. */
export function lisansBildirimi(secenekler?: LisansBildirimiSecenekleri): Plugin;
