// Yazdırma ve PDF olarak dışarı aktarma: ekran düzeninin sayfa kopyalarından yazdırma kökü kurulur.
// Masaüstünde (getNative().exportPdf) doğrudan .pdf kaydedilir; yoksa tarayıcının yazdırma penceresi açılır.
import { useCallback, useEffect, useState } from "react";
import { flushSync } from "react-dom";
import { baseName, getNative } from "@officehere/core";
import { pageSizeId, type DocSettings } from "../model/document";
import type { PaginationController } from "../editor/pagination";
import { buildPrintRoot, removePrintRoot } from "../views/printRender";

interface PrintDeps {
  controller: PaginationController;
  settings: DocSettings;
  zoom: number;
  fileName: string;
  /** Yazdırmadan önce Dosya görünümü, üst bilgi düzenleme ve odak modu kapatılır. */
  leaveModes: () => void;
  notify: (text: string, kind?: "info" | "warning" | "error" | "success") => void;
}

const frames = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

export function usePrint({ controller, settings, zoom, fileName, leaveModes, notify }: PrintDeps) {
  const [printing, setPrinting] = useState(false);

  const preparePrint = useCallback(() => {
    const pagesEl = document.querySelector<HTMLElement>(".bl-pages");
    if (!pagesEl) return false;
    return !!buildPrintRoot(pagesEl, { pages: controller.info.pages, scale: zoom / 100, pageColor: settings.pageColor });
  }, [settings.pageColor, controller, zoom]);

  const finishPrint = useCallback(() => removePrintRoot(), []);

  /** Yazdır ya da (masaüstünde) doğrudan PDF olarak kaydet. */
  const print = useCallback(
    async (target: "print" | "pdf" = "print") => {
      const native = getNative();
      if (target === "pdf" && !native?.exportPdf) target = "print";
      flushSync(leaveModes);
      await frames();
      controller.measureNow?.();
      await frames();
      if (!preparePrint()) return;
      setPrinting(true);
      try {
        if (target === "pdf" && native?.exportPdf) {
          // Yönlendirme ve kağıt boyutu ilk sayfadan (belge ayarı / ilk bölüm).
          const first = controller.info.pages[0];
          const sizeId = pageSizeId({ ...settings.page, width: Math.round((first?.w ?? 794) * 15), height: Math.round((first?.h ?? 1123) * 15) });
          const landscape = !!first && first.w > first.h;
          const pageSize = sizeId !== "custom" ? (sizeId as "A4") : { width: (first?.w ?? 794) / 96, height: (first?.h ?? 1123) / 96 };
          const res = await native.exportPdf(`${baseName(fileName)}.pdf`, { landscape, pageSize });
          if (res) notify(`"${res.name}" kaydedildi.`, "success");
        } else window.print();
      } catch (err) {
        notify(`PDF oluşturulamadı: ${(err as Error).message}`, "error");
      } finally {
        finishPrint();
        setPrinting(false);
      }
    },
    [controller, leaveModes, preparePrint, finishPrint, settings.page, fileName, notify],
  );

  useEffect(() => {
    if (import.meta.env.DEV) Object.assign((window as unknown as { __belge: object }).__belge ?? {}, { preparePrint, finishPrint });
  }, [preparePrint, finishPrint]);

  // Tarayıcı menüsünden yazdırma.
  useEffect(() => {
    const before = () => {
      if (!document.getElementById("bl-print-root")) preparePrint();
    };
    window.addEventListener("beforeprint", before);
    window.addEventListener("afterprint", finishPrint);
    return () => {
      window.removeEventListener("beforeprint", before);
      window.removeEventListener("afterprint", finishPrint);
    };
  }, [preparePrint, finishPrint]);

  return { printing, print };
}
