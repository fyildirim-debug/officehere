import { useEffect, useState } from "react";

export type UiTheme = "colorful" | "white" | "dark" | "black" | "system";

export const UI_THEMES: { id: UiTheme; label: string }[] = [
  { id: "colorful", label: "Renkli" },
  { id: "white", label: "Beyaz" },
  { id: "dark", label: "Koyu Gri" },
  { id: "black", label: "Siyah" },
  { id: "system", label: "Sistem ayarını kullan" },
];

const KEY = "officehere.uiTheme";

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

/** Office teması (Renkli / Beyaz / Koyu Gri / Siyah); tüm uygulamalarda ortaktır. */
export function useUiTheme(): [UiTheme, (t: UiTheme) => void] {
  const [theme, setTheme] = useState<UiTheme>(read);
  useEffect(() => {
    applyUiTheme(theme);
    try {
      localStorage.setItem(KEY, theme);
    } catch {
      // yok say
    }
  }, [theme]);
  useEffect(() => {
    const onStorage = (e: StorageEvent) => e.key === KEY && e.newValue && setTheme(e.newValue as UiTheme);
    window.addEventListener("storage", onStorage);
    return () => window.removeEventListener("storage", onStorage);
  }, []);
  return [theme, setTheme];
}
