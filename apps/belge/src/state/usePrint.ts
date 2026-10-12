// Yazdırma ve PDF olarak dışarı aktarma: ekran düzeninin sayfa kopyalarından yazdırma kökü kurulur.
// Masaüstünde (getNative().exportPdf) doğrudan .pdf kaydedilir; yoksa tarayıcının yazdırma penceresi açılır.
// Dosya › Yazdır ayarları (sayfa aralığı, kopya, harmanlama, yaprak başına sayfa, elle iki taraflı) uygulanır.
import { useCallback, useEffect, useRef, useState } from "react";
import { flushSync } from "react-dom";
import { baseName, getNative } from "@officehere/core";
import { pageSizeId, type DocSettings } from "../model/document";
import type { PaginationController } from "../editor/pagination";
import { buildPrintRoot, removePrintRoot } from "../views/printRender";
import { defaultPrintOptions, printSheets, selectedPages, type PrintOptions } from "../print/printOptions";

interface PrintDeps {
  controller: PaginationController;
  settings: DocSettings;
  zoom: number;
  fileName: string;
  /** İmlecin bulunduğu sayfa (Geçerli Sayfayı Yazdır). */
  currentPage: () => number;
  /** Yazdırmadan önce Dosya görünümü, üst bilgi düzenleme ve odak modu kapatılır. */
  leaveModes: () => void;
  notify: (text: string, kind?: "info" | "warning" | "error" | "success") => void;
}

const frames = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

export function usePrint({ controller, settings, zoom, fileName, currentPage, leaveModes, notify }: PrintDeps) {
  const [printing, setPrinting] = useState(false);
  const [options, setOptions] = useState<PrintOptions>(defaultPrintOptions);
  const [duplexWait, setDuplexWait] = useState(false);
  const pending = useRef<PrintOptions | null>(null);

  /** Yazdırma kökünü kurar; ayarlar verilirse yalnız seçilen sayfalar ve yapraklar basılır. */
  const preparePrint = useCallback(
    (opts?: PrintOptions, pass?: 1 | 2) => {
      const pagesEl = document.querySelector<HTMLElement>(".bl-shell-doc .bl-pages") ?? document.querySelector<HTMLElement>(".bl-pages");
      if (!pagesEl) return false;
      const g = { pages: controller.info.pages, scale: zoom / 100, pageColor: settings.pageColor };
      if (!opts) return !!buildPrintRoot(pagesEl, g);
      const sel = selectedPages(opts, controller.info.pageCount, currentPage());
      if (typeof sel === "string") {
        notify(sel, "warning");
        return false;
      }
      return !!buildPrintRoot(pagesEl, g, { sheets: printSheets(sel, opts, opts.duplex === "manual" ? pass : undefined), perSheet: opts.perSheet });
    },
    [settings.pageColor, controller, zoom, currentPage, notify],
  );

  const finishPrint = useCallback(() => removePrintRoot(), []);

  /** Yazdır ya da (masaüstünde) doğrudan PDF olarak kaydet. */
  const run = useCallback(
    async (target: "print" | "pdf", opts: PrintOptions | undefined, pass?: 1 | 2) => {
      const native = getNative();
      if (target === "pdf" && !native?.exportPdf) target = "print";
      flushSync(leaveModes);
      await frames();
      controller.measureNow?.();
      await frames();
      if (!preparePrint(opts, pass)) return false;
      setPrinting(true);
      try {
        if (target === "pdf" && native?.exportPdf) {
          // Yönlendirme ve kağıt boyutu ilk yapraktan.
          const first = document.querySelector<HTMLElement>("#bl-print-root > .bl-print-page");
          const w = parseFloat(first?.style.width ?? "794");
          const h = parseFloat(first?.style.height ?? "1123");
          const sizeId = pageSizeId({ ...settings.page, width: Math.round(w * 15), height: Math.round(h * 15) });
          const pageSize = sizeId !== "custom" ? (sizeId as "A4") : { width: w / 96, height: h / 96 };
          const res = await native.exportPdf(`${baseName(fileName)}.pdf`, { landscape: w > h, pageSize });
          if (res) notify(`"${res.name}" kaydedildi.`, "success");
        } else window.print();
      } catch (err) {
        notify(`PDF oluşturulamadı: ${(err as Error).message}`, "error");
      } finally {
        finishPrint();
        setPrinting(false);
      }
      return true;
    },
    [controller, leaveModes, preparePrint, finishPrint, settings.page, fileName, notify],
  );

  const print = useCallback(
    async (target: "print" | "pdf" = "print", opts?: PrintOptions) => {
      const manual = !!opts && opts.duplex === "manual" && target === "print";
      const ok = await run(target, opts, manual ? 1 : undefined);
      // Elle iki taraflı: ön yüzlerden sonra kağıt yeniden yerleştirilip arka yüzler basılır.
      if (ok && manual && opts) {
        pending.current = opts;
        setDuplexWait(true);
      }
    },
    [run],
  );

  const duplex = {
    open: duplexWait,
    /** Arka yüzleri yazdır. */
    next: () => {
      setDuplexWait(false);
      const o = pending.current;
      if (o) void run("print", o, 2);
    },
    cancel: () => setDuplexWait(false),
  };

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

  return { printing, print, options, setOptions, duplex };
}
