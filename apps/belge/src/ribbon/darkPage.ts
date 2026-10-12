// Görünüm › Koyu Sayfa: ekranda sayfa zemini koyu, otomatik (siyah) metin açık çizilir. Yalnız ekran
// görünümünü etkiler (yazdırma, Baskı Önizleme ve PDF beyaz kalır); seçim bu tarayıcıda hatırlanır.
// Görünüm <html data-bl-page="dark"> ile seçilir, stiller styles/dark-page.css dosyasındadır.
import { useSyncExternalStore } from "react";

export const DARK_PAGE_KEY = "officehere.belge.darkPage";

const listeners = new Set<() => void>();

function storage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

export function readDarkPage(): boolean {
  try {
    return storage()?.getItem(DARK_PAGE_KEY) === "1";
  } catch {
    return false;
  }
}

function apply(on: boolean) {
  if (typeof document === "undefined") return;
  if (on) document.documentElement.setAttribute("data-bl-page", "dark");
  else document.documentElement.removeAttribute("data-bl-page");
}

// Açılışta (şerit çizilmeden) kayıtlı seçim uygulanır.
let current = readDarkPage();
apply(current);

export function isDarkPage(): boolean {
  return current;
}

export function setDarkPage(on: boolean) {
  current = on;
  try {
    if (on) storage()?.setItem(DARK_PAGE_KEY, "1");
    else storage()?.removeItem(DARK_PAGE_KEY);
  } catch {
    // Depolama kapalıysa seçim yalnız bu oturumda geçerlidir.
  }
  apply(on);
  listeners.forEach((l) => l());
}

const subscribe = (l: () => void) => {
  listeners.add(l);
  return () => {
    listeners.delete(l);
  };
};

export function useDarkPage(): boolean {
  return useSyncExternalStore(subscribe, isDarkPage, () => false);
}
