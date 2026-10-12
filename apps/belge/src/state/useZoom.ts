// Yakınlaştırma hedefleri: yüzde, sayfa genişliği, metin genişliği, tam sayfa, birden çok sayfa.
import { useCallback, type RefObject } from "react";
import type { Editor } from "@tiptap/core";
import type { PaginationController } from "../editor/pagination";
import type { ZoomTarget } from "../dialogs/MiscDialogs";
import type { ViewOptions } from "./context";
import { TILE_GAP } from "../views/DocumentView";

export type ZoomArg = ZoomTarget | "100" | "onePage" | "multi" | "width";

export function useZoom(controller: PaginationController, editor: Editor, scrollRef: RefObject<HTMLDivElement | null>, zoom: number, setView: (p: Partial<ViewOptions>) => void) {
  return useCallback(
    (z: ZoomArg) => {
      const sc = scrollRef.current;
      // İmleçteki sayfanın geometrisi (bölümler farklı boyutta olabilir).
      const info = controller.info;
      const pg = info.pages[Math.min(info.pages.length - 1, controller.pageOf(editor.state.selection.from))] ?? info.pages[0];
      const w = info.bodyW || pg.w;
      const availW = (sc?.clientWidth ?? 1000) - 48;
      const availH = (sc?.clientHeight ?? 700) - 48;
      let pct = zoom;
      const target = typeof z === "string" ? (z === "100" ? { pct: 100 } : { fit: z === "onePage" ? "page" : z }) : z;
      if ("pct" in target) pct = target.pct;
      else if (target.fit === "width") pct = (availW / w) * 100;
      else if (target.fit === "text") pct = (availW / (pg.w - pg.ml - pg.mr)) * 100;
      else if (target.fit === "page") pct = Math.min(availW / w, availH / pg.h) * 100;
      // Birden Çok Sayfa: iki sayfa yan yana tam görünür.
      else if (target.fit === "multi") pct = Math.min(availW / (2 * w + TILE_GAP + 8), availH / pg.h) * 100;
      // Tek Sayfa / Tam sayfa: sığsa da sayfalar yan yana dizilmez.
      const onePage = "fit" in target && target.fit === "page";
      setView({ zoom: Math.max(10, Math.min(500, Math.floor(pct))), onePage });
    },
    [controller, editor, scrollRef, zoom, setView],
  );
}
