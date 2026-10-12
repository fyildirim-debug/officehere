// Dosya › Yazdır önizlemesi için sayfa kopyaları: Dosya görünümü açıkken belge ekran dışında yazdırma
// düzeninde (ölçek 1, sayfa aralığı yok) çizilir; düzen her değiştiğinde kopyalar yenilenir.
import { useEffect, useState } from "react";
import type { PageInfo, PaginationController } from "../editor/pagination";
import { renderPages } from "../views/printRender";

export function usePreviewPages(controller: PaginationController, active: boolean, pageColor: string | null) {
  const [preview, setPreview] = useState<{ pages: HTMLElement[]; infos: PageInfo[] } | null>(null);
  useEffect(() => {
    if (!active) {
      setPreview(null);
      return;
    }
    let t = 0;
    const build = () => {
      window.clearTimeout(t);
      t = window.setTimeout(() => {
        const el = document.querySelector<HTMLElement>(".bl-offscreen .bl-pages");
        const infos = controller.info.pages;
        if (!el || !infos.length) return;
        const pages = renderPages(el, { pages: infos, scale: 1, pageColor });
        if (pages) setPreview({ pages, infos });
      }, 150);
    };
    build();
    const off = controller.subscribe(build);
    return () => {
      off();
      window.clearTimeout(t);
    };
  }, [controller, active, pageColor]);
  return preview;
}
