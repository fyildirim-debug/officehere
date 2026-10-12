/**
 * Otomatik kaydetme (masaüstü, yalnız .xlsx): kitap değiştikçe boşta kalma
 * süresinden sonra aynı dosyaya yazılır. Kirli bayrak temizlenince Otomatik
 * Kurtarma taslağı da silinir (useAutoRecover). Başka kitap yüklenince
 * başlık çubuğundaki "Kaydedildi" bilgisi sıfırlanır.
 */
import { useEffect } from "react";
import { saveStatus, useAutoSave } from "@officehere/ui";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
import { writeXlsx } from "../io/xlsx";

export function useFileAutoSave(ctl: Controller, files: FileActions): void {
  useEffect(() => {
    let wb = ctl.wb;
    return ctl.subscribe(() => {
      if (ctl.wb === wb) return;
      wb = ctl.wb;
      saveStatus.reset();
    });
  }, [ctl]);

  useAutoSave({
    path: ctl.fileRef?.path,
    formats: ["xlsx"],
    isDirty: () => ctl.dirty,
    // Yalnız içerik değişiklikleri (seçim/kaydırma değil).
    subscribe: (onChange) => {
      let v = ctl.dataVersion;
      return ctl.subscribe(() => {
        if (ctl.dataVersion === v) return;
        v = ctl.dataVersion;
        onChange();
      });
    },
    prepare: async () => {
      const v = ctl.dataVersion;
      const wb = ctl.wb;
      return { data: await writeXlsx(wb), isCurrent: () => ctl.dataVersion === v && ctl.wb === wb };
    },
    markClean: () => {
      ctl.dirty = false;
      ctl.emit();
    },
    save: () => void files.save(),
  });
}
