import { useCallback, useEffect, useState } from "react";

export type UiTheme = "colorful" | "white" | "dark" | "black" | "system";

export const UI_THEMES: { id: UiTheme; label: string }[] = [
  { id: "colorful", label: "Renkli" },
  { id: "white", label: "Beyaz" },
  { id: "dark", label: "Koyu Gri" },
  { id: "black", label: "Siyah" },
  { id: "system", label: "Sistem ayarını kullan" },
];

const KEY = "officehere.uiTheme";
const EVENT = "officehere:ui-theme";

function read(): UiTheme {
  try {
    return (localStorage.getItem(KEY) as UiTheme) || "colorful";
  } catch {
    return "colorful";
  }
}

export function applyUiTheme(theme: UiTheme) {
  document.documentElement.setAttribute("data-oh-theme", theme);
}

/** Arayüz teması (Renkli / Beyaz / Koyu Gri / Siyah); tüm uygulamalarda ortaktır. */
export function useUiTheme(): [UiTheme, (t: UiTheme) => void] {
  const [theme, setTheme] = useState<UiTheme>(read);
  useEffect(() => {
    applyUiTheme(theme);
    try {
      if (localStorage.getItem(KEY) !== theme) localStorage.setItem(KEY, theme);
    } catch {
      // yok say
    }
  }, [theme]);
  useEffect(() => {
    // Başka pencerelerde (storage) ve aynı penceredeki diğer kanca örneklerinde (özel olay) eşitle.
    const onStorage = (e: StorageEvent) => e.key === KEY && e.newValue && setTheme(e.newValue as UiTheme);
    const onLocal = (e: Event) => setTheme((e as CustomEvent<UiTheme>).detail);
    window.addEventListener("storage", onStorage);
    window.addEventListener(EVENT, onLocal);
    return () => {
      window.removeEventListener("storage", onStorage);
      window.removeEventListener(EVENT, onLocal);
    };
  }, []);
  const set = useCallback((t: UiTheme) => {
    setTheme(t);
    window.dispatchEvent(new CustomEvent<UiTheme>(EVENT, { detail: t }));
  }, []);
  return [theme, set];
}
