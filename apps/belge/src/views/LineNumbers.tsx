// Satır numaraları katmanı: sayfa düzeni değişince (ve düzenlemeden kısa süre sonra) satırlar ölçülür,
// numaralar her satırın sütununun solunda, metinden ayarlanan uzaklıkta çizilir.
import { useEffect, useState } from "react";
import type { Editor } from "@tiptap/core";
import type { PageInfo, PaginationController } from "../editor/pagination";
import type { SecGeom } from "../editor/sections";
import { measureLines, numberLines } from "../editor/lineNumbers";

interface Props {
  editor: Editor;
  controller: PaginationController;
  pagesRef: React.RefObject<HTMLDivElement | null>;
  scale: number;
  pages: PageInfo[];
  geoms: SecGeom[];
  rectOf: (i: number) => { x: number; y: number };
}

interface Item {
  key: string;
  x: number;
  y: number;
  h: number;
  n: number;
}

/** "Otomatik" uzaklık (0,25 inç). */
const AUTO_DISTANCE = 24;

export function LineNumbers({ editor, controller, pagesRef, scale, pages, geoms, rectOf }: Props) {
  const [items, setItems] = useState<Item[]>([]);
  const enabled = geoms.some((g) => !!g.page.lineNumbers);
  useEffect(() => {
    if (!enabled) {
      setItems([]);
      return;
    }
    let timer = 0;
    const run = () => {
      timer = 0;
      const host = pagesRef.current;
      if (!host || editor.isDestroyed) return;
      const pageAt = (x: number, y: number) => {
        for (let i = 0; i < pages.length; i++) {
          const r = rectOf(i);
          if (y >= r.y && y <= r.y + pages[i].h && x >= r.x - 8 && x <= r.x + pages[i].w + 8) return i;
        }
        return Math.max(0, pages.length - 1);
      };
      const colAt = (page: number, x: number) => {
        const g = geoms[pages[page]?.section ?? 0];
        if (!g || g.cols < 2) return 0;
        const rel = x - rectOf(page).x - g.ml;
        return Math.max(0, Math.min(g.cols - 1, Math.floor(rel / (g.colW + g.colGap) + 0.02)));
      };
      const lines = measureLines(editor.view, host, scale, (si) => !!geoms[si]?.page.lineNumbers, pageAt, colAt);
      const numbered = numberLines(lines, (si) => geoms[si]?.page.lineNumbers);
      setItems(
        numbered.map(({ line, n }, k) => {
          const g = geoms[line.section] ?? geoms[0];
          const r = rectOf(line.page);
          const col = (line as typeof line & { col: number }).col % 100;
          const colLeft = r.x + g.ml + col * (g.colW + g.colGap);
          const dist = g.page.lineNumbers?.distance ? g.page.lineNumbers.distance / 15 : AUTO_DISTANCE;
          return { key: `${k}:${n}`, x: colLeft - dist, y: line.top, h: line.height, n };
        }),
      );
    };
    const schedule = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(run, 150);
    };
    schedule();
    const off = controller.subscribe(schedule);
    editor.on("update", schedule);
    return () => {
      window.clearTimeout(timer);
      off();
      editor.off("update", schedule);
    };
  }, [enabled, editor, controller, pagesRef, scale, pages, geoms, rectOf]);
  if (!enabled || !items.length) return null;
  return (
    <div className="bl-lnums bl-doc" aria-hidden="true">
      {items.map((it) => (
        <span key={it.key} className="bl-lnum" style={{ top: it.y, left: it.x - 64, height: it.h, lineHeight: `${it.h}px` }}>
          {it.n}
        </span>
      ))}
    </div>
  );
}
