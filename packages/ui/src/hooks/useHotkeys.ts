import { useEffect, useRef } from "react";

export type HotkeyHandler = (e: KeyboardEvent) => void;

/**
 * Klavye kısayolları. Anahtar biçimi: "Mod+S", "Mod+Shift+Z", "F12", "Alt+Q".
 * "Mod" Windows/Linux'ta Ctrl, macOS'ta Cmd'dir. İşleyici bulunursa
 * varsayılan davranış engellenir.
 */
export function useHotkeys(map: Record<string, HotkeyHandler>, deps: unknown[] = []) {
  const ref = useRef(map);
  ref.current = map;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.isComposing) return;
      const combo = comboOf(e);
      const handler = ref.current[combo];
      if (handler) {
        e.preventDefault();
        handler(e);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);
}

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);

export function comboOf(e: KeyboardEvent): string {
  const parts: string[] = [];
  if (isMac ? e.metaKey : e.ctrlKey) parts.push("Mod");
  if (e.altKey) parts.push("Alt");
  if (e.shiftKey) parts.push("Shift");
  let key = e.key;
  if (key.length === 1) key = key.toUpperCase();
  if (e.code.startsWith("Key") && (e.ctrlKey || e.metaKey || e.altKey)) key = e.code.slice(3);
  if (e.code.startsWith("Digit") && (e.ctrlKey || e.metaKey || e.altKey)) key = e.code.slice(5);
  if (key === " ") key = "Space";
  parts.push(key);
  return parts.join("+");
}

/** Kısayolu kullanıcıya gösterilecek biçime çevirir ("Mod+S" → "Ctrl+S"). */
export function formatShortcut(combo: string): string {
  return combo.replace("Mod", isMac ? "⌘" : "Ctrl");
}
