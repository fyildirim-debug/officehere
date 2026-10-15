/**
 * Otomatik kaydetme (masaüstü, yalnız .xlsx): kitap değiştikçe boşta kalma
 * süresinden sonra aynı dosyaya yazılır. Kirli bayrak temizlenince Otomatik
 * Kurtarma taslağı da silinir (useAutoRecover). Başka kitap yüklenince
 * başlık çubuğundaki "Kaydedildi" bilgisi sıfırlanır.
 *
 * Dosyadaki bazı öğeler okunamadıysa (loadWarnings) ilk elle kayda kadar
 * otomatik kaydetme yapılmaz; yazılan dosya kendi okuyucumuzla yeniden açılarak
 * doğrulanır (açılamazsa önceki içerik geri yüklenir). Yazma ve doğrulama
 * Worker'dadır (ayrı istemciler); ana iş parçacığı yalnız kitabın anlık
 * görüntüsünü parça parça çıkarır, büyük kitapta arayüz donmaz.
 */
import { useEffect, useRef } from "react";
import { saveStatus, useAutoSave } from "@officehere/ui";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
import { verifyXlsxIo, writeBookIo } from "../io/client";

/** Rozette gösterilen bekletme nedeni. */
export const LOAD_WARNING_HOLD = "Bu dosyada desteklenmeyen öğeler var; ilk kaydı siz onaylayın";

/** Yazılan .xlsx kendi okuyucumuzla açılabiliyor mu ve yazılan sayfaların hiçbiri kaybolmamış mı. */
export async function canReopenXlsx(data: ArrayBuffer, expectedSheets = 1): Promise<boolean> {
  return verifyXlsxIo(data, expectedSheets);
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
      const data = await writeBookIo(wb, wb.active, "xlsx", undefined, undefined, { version: () => ctl.dataVersion });
      return { data: data as Uint8Array<ArrayBuffer>, isCurrent: () => ctl.dataVersion === v && ctl.wb === wb };
    },
    markClean: () => {
      ctl.dirty = false;
      ctl.emit();
    },
    save: () => void files.save(),
  });
}
