/**
 * Otomatik kaydetme (masaüstü, .pptx): sunu değiştikçe boşta kalma süresinden
 * sonra aynı dosyaya yazılır. Kirli bayrak temizlenince Otomatik Kurtarma
 * taslağı da silinir (useAutoRecover).
 *
 * Dosya açılırken okunamayan bölümler olduysa (loadWarnings) ilk elle kayda kadar
 * otomatik kaydetme yapılmaz; yazılan dosya kendi okuyucumuzla yeniden açılarak
 * doğrulanır (açılamazsa önceki içerik geri yüklenir).
 */
import { useRef } from "react";
import { useAutoSave } from "@officehere/ui";
import { readPresentationFile } from "../io/formats";
import type { Actions } from "../state/actions";
import { documentPres } from "../state/masterActions";
import { presentationBytes, save } from "./fileOps";

/** Rozette gösterilen bekletme nedeni. */
export const LOAD_WARNING_HOLD = "Bu dosyada desteklenmeyen öğeler var; ilk kaydı siz onaylayın";

/** Yazılan .pptx kendi okuyucumuzla açılabiliyor mu ve yazılan slaytların hiçbiri kaybolmamış mı. */
export async function canReopenPptx(data: ArrayBuffer, expectedSlides = 0): Promise<boolean> {
  const { pres } = await readPresentationFile(data, "dogrulama.pptx");
  return Array.isArray(pres?.slides) && pres.slides.length >= expectedSlides;
}

export function useFileAutoSave(a: Actions, path: string | undefined): void {
  const warnings = (a.s as { loadWarnings?: string[] }).loadWarnings;
  const expected = useRef(0);
  useAutoSave({
    path,
    formats: ["pptx"],
    holdReason: warnings?.length ? LOAD_WARNING_HOLD : null,
    verify: (data) => canReopenPptx(data, expected.current),
    isDirty: () => a.s.dirty,
    // Yalnız sunu içeriği değişince (seçim, görünüm vb. değil).
    subscribe: (onChange) => {
      let pres = a.s.pres;
      return a.store.subscribe(() => {
        if (a.s.pres === pres) return;
        pres = a.s.pres;
        onChange();
      });
    },
    prepare: async () => {
      const { pres, masterBase } = a.s;
      expected.current = documentPres(a).slides.length;
      return { data: await presentationBytes(a), isCurrent: () => a.s.pres === pres && a.s.masterBase === masterBase };
    },
    markClean: () => a.store.set({ dirty: false }),
    save: () => void save(a),
  });
}
