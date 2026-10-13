/** Pencere genişliği verilen sınırın altında mı (telefon yerleşimi için)? */
import { useSyncExternalStore } from "react";

export function useNarrow(maxWidth: number): boolean {
  const q = `(max-width: ${maxWidth}px)`;
  return useSyncExternalStore(
    (l) => {
      const m = window.matchMedia(q);
      m.addEventListener("change", l);
      return () => m.removeEventListener("change", l);
    },
    () => window.matchMedia(q).matches,
    () => false,
  );
}
