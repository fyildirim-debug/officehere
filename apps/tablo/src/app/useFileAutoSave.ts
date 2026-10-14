/**
 * Otomatik kaydetme (masaüstü, yalnız .xlsx): kitap değiştikçe boşta kalma
 * süresinden sonra aynı dosyaya yazılır. Kirli bayrak temizlenince Otomatik
 * Kurtarma taslağı da silinir (useAutoRecover). Başka kitap yüklenince
 * başlık çubuğundaki "Kaydedildi" bilgisi sıfırlanır.
 *
 * Dosyadaki bazı öğeler okunamadıysa (loadWarnings) ilk elle kayda kadar
 * otomatik kaydetme yapılmaz; yazılan dosya kendi okuyucumuzla yeniden açılarak
 * doğrulanır (açılamazsa önceki içerik geri yüklenir).
 */
import { useEffect, useRef } from "react";
import { saveStatus, useAutoSave } from "@officehere/ui";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
import { openBook } from "../io/formats";
import { writeXlsx } from "../io/xlsx";

/** Rozette gösterilen bekletme nedeni. */
export const LOAD_WARNING_HOLD = "Bu dosyada desteklenmeyen öğeler var; ilk kaydı siz onaylayın";

/** Yazılan .xlsx kendi okuyucumuzla açılabiliyor mu ve yazılan sayfaların hiçbiri kaybolmamış mı. */
export async function canReopenXlsx(data: ArrayBuffer, expectedSheets = 1): Promise<boolean> {
  const opened = await openBook("dogrulama.xlsx", data);
  return opened.wb.sheets.length >= Math.max(1, expectedSheets);
}

export function useFileAutoSave(ctl: Controller, files: FileActions): void {
  const expected = useRef(1);
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
    holdReason: ctl.wb.loadWarnings.length ? LOAD_WARNING_HOLD : null,
    verify: (data) => canReopenXlsx(data, expected.current),
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
      expected.current = wb.sheets.length;
      return { data: await writeXlsx(wb), isCurrent: () => ctl.dataVersion === v && ctl.wb === wb };
    },
    markClean: () => {
      ctl.dirty = false;
      ctl.emit();
    },
    save: () => void files.save(),
  });
}
