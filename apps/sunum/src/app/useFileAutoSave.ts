/**
 * Otomatik kaydetme (masaüstü, .pptx): sunu değiştikçe boşta kalma süresinden
 * sonra aynı dosyaya yazılır. Kirli bayrak temizlenince Otomatik Kurtarma
 * taslağı da silinir (useAutoRecover).
 */
import { useAutoSave } from "@officehere/ui";
import type { Actions } from "../state/actions";
import { presentationBytes, save } from "./fileOps";

export function useFileAutoSave(a: Actions, path: string | undefined): void {
  useAutoSave({
    path,
    formats: ["pptx"],
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
      return { data: await presentationBytes(a), isCurrent: () => a.s.pres === pres && a.s.masterBase === masterBase };
    },
    markClean: () => a.store.set({ dirty: false }),
    save: () => void save(a),
  });
}
