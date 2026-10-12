/**
 * Sanallaştırılmış canvas ızgara bileşeni: çizim döngüsü, kaydırma
 * çubukları, fare etkileşimleri (seçim, boyutlandırma, doldurma tutamacı,
 * taşıma, formül işaretleme) ve klavye.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react";
import type { Controller } from "../controller";
import * as E from "../actions/edit";
import { buildClipboard, pastePayload } from "../model/clipboard";
import { MAX_COLS, MAX_ROWS, contains, rect, rectCols, rectRows } from "../model/rect";
import { expandForMerges } from "../model/selection";
import type { CellPos, Rect } from "../model/types";
import { widthPxToChars, pxToPt, type AxisSizes } from "../model/axis";
import { CellEditor } from "./CellEditor";
import { handleGridKey } from "./keys";
import { GridOverlays } from "./GridOverlays";
import { sizeSelection } from "../actions/format";
import { handleOutlineClick } from "../actions/outline";
import { activeFilterRows } from "./gridUtil";
import { GridLayout } from "./layout";
import { measureText, renderGrid, type RenderState } from "./render";
import { SheetTabs } from "./SheetTabs";
import { autoFitCols, autoFitRows } from "./autofit";
import { cursorFor, fillTipText, hitTestGrid, openLink, selectedLines, useDarkUi, type Hit } from "./gridUtil";
import { pivotButtons } from "../pivot/pivotOps";
import { computePageBreaks } from "../io/print";
import { pageBreakCursor, pageBreakDrop, pageBreakHit, pageBreakMove, type PageBreakDrag } from "./pageBreakDrag";
import { filtersOf } from "../actions/data";

type Drag =
  | { kind: "select"; add: boolean }
  | { kind: "colsel"; start: number; add: boolean }
  | { kind: "rowsel"; start: number; add: boolean }
  | { kind: "colresize"; col: number; startX: number; startW: number; saved: AxisSizes }
  | { kind: "rowresize"; row: number; startY: number; startH: number; saved: AxisSizes }
  | { kind: "fill"; src: Rect; target: Rect | null }
  | { kind: "point"; anchor: CellPos; sheetId: number }
  | { kind: "move"; src: Rect; grab: CellPos; target: Rect | null; copy: boolean }
  | { kind: "pagebreak"; pb: PageBreakDrag };

export interface GridProps {
  ctl: Controller;
  onContextMenu: (e: { clientX: number; clientY: number; preventDefault: () => void }, kind: "cell" | "col" | "row" | "tab", index?: number) => void;
  onFilterButton: (row: number, col: number, x: number, y: number) => void;
}

const SCROLL_CAP = 15_000_000;

export function Grid({ ctl, onContextMenu, onFilterButton }: GridProps) {
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const sheet = ctl.sheet;
  const wrapRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const vbarRef = useRef<HTMLDivElement>(null);
  const hbarRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const scrolls = useRef(new Map<number, { x: number; y: number }>());
  const [size, setSize] = useState({ w: 800, h: 500 });
  const [, force] = useState(0);
  const drag = useRef<Drag | null>(null);
  const lastPointer = useRef<{ x: number; y: number } | null>(null);
  const autoScroll = useRef<number | null>(null);
  const syncing = useRef(false);
  const raf = useRef(0);
  const march = useRef(0);
  const [tip, setTip] = useState<{ x: number; y: number; text: string; title?: string } | null>(null);
  const [cursor, setCursor] = useState("cell");
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  const darkUi = useDarkUi();

  const scroll = () => {
    let s = scrolls.current.get(sheet.id);
    if (!s) scrolls.current.set(sheet.id, (s = { x: 0, y: 0 }));
    return s;
  };

  const measureDigits = useCallback((n: number) => measureText("8".repeat(n), `${11 * Math.max(0.65, Math.min(sheet.zoom / 100, 2.2))}px "Segoe UI", sans-serif`), [sheet.zoom]);
  const makeLayout = useCallback(() => new GridLayout(sheet, size.w, size.h, scroll().x, scroll().y, measureDigits), [sheet, size.w, size.h, measureDigits]); // eslint-disable-line react-hooks/exhaustive-deps
  const layoutRef = useRef<GridLayout | null>(null);
  layoutRef.current = makeLayout();
  const layout = layoutRef.current;

  /* --------------------------- Çizim --------------------------- */

  const draw = useCallback(() => {
    cancelAnimationFrame(raf.current);
    raf.current = requestAnimationFrame(() => {
      const c = canvasRef.current;
      if (!c) return;
      const ctx = c.getContext("2d");
      if (!ctx) return;
      const L = makeLayout();
      layoutRef.current = L;
      const e = ctl.edit;
      const refs =
        e && (e.sheetId === sheet.id || e.text.includes("!"))
          ? ctl
              .editRefs()
              .filter((r) => (r.sheetName ? r.sheetName.toLocaleUpperCase("tr-TR") === sheet.name.toLocaleUpperCase("tr-TR") : e.sheetId === sheet.id))
          : [];
      const filterButtons: { row: number; col: number; active: boolean }[] = [];
      for (const f of filtersOf(sheet)) for (let col = f.range.c1; col <= f.range.c2; col++) filterButtons.push({ row: f.range.r1, col, active: !!f.columns[col - f.range.c1] });
      for (const b of pivotButtons(sheet)) filterButtons.push({ row: b.row, col: b.col, active: b.active });
      const spillInfo = ctl.wb.engine.getSpill(sheet.id, ctl.sel.active.row, ctl.sel.active.col);
      const st: RenderState = {
        wb: ctl.wb,
        sheet,
        layout: L,
        cf: ctl.cf,
        dpr,
        selection: ctl.sel,
        showSelection: !(e && e.pointSpan && e.sheetId !== sheet.id),
        refs,
        clip: ctl.clip && ctl.clip.sheetId === sheet.id ? { rect: ctl.clip.rect, cut: ctl.clip.cut, offset: march.current } : null,
        fillPreview: drag.current?.kind === "fill" ? drag.current.target : null,
        movePreview: drag.current?.kind === "move" ? drag.current.target : null,
        editing: e && e.sheetId === sheet.id ? { row: e.row, col: e.col } : null,
        filterButtons,
        spill: spillInfo ? rect(spillInfo.anchorRow, spillInfo.anchorCol, spillInfo.anchorRow + spillInfo.rows - 1, spillInfo.anchorCol + spillInfo.cols - 1) : null,
        dark: darkUi,
        arrows: ctl.arrows,
        errorCheck: ctl.errorChecking,
        pageBreaks: sheet.viewMode === "pageBreak" ? computePageBreaks(ctl.wb, sheet) : null,
        pageBreakDrag: drag.current?.kind === "pagebreak" ? drag.current.pb : null,
        filtered: activeFilterRows(sheet),
      };
      renderGrid(ctx, st);
      syncScrollbars(L);
    });
  }, [ctl, sheet, makeLayout, dpr, darkUi]); // eslint-disable-line react-hooks/exhaustive-deps

  const syncScrollbars = (L: GridLayout) => {
    const ext = ctl.wb.dataExtent(sheet);
    const used = ctl.wb.usedRange(sheet);
    const s = scroll();
    const content = L.contentSize(Math.max(ext.rows, used ? used.r2 + 1 : 0, ctl.sel.active.row + 1), Math.max(ext.cols, used ? used.c2 + 1 : 0, ctl.sel.active.col + 1));
    const h = Math.min(SCROLL_CAP, Math.max(content.h + L.paneH * 0.5, s.y + L.paneH * 1.5));
    const w = Math.min(SCROLL_CAP, Math.max(content.w + L.paneW * 0.5, s.x + L.paneW * 1.5));
    const vb = vbarRef.current;
    const hb = hbarRef.current;
    syncing.current = true;
    if (vb) {
      (vb.firstElementChild as HTMLElement).style.height = `${h}px`;
      if (Math.abs(vb.scrollTop - s.y) > 1) vb.scrollTop = s.y;
    }
    if (hb) {
      (hb.firstElementChild as HTMLElement).style.width = `${w}px`;
      if (Math.abs(hb.scrollLeft - s.x) > 1) hb.scrollLeft = s.x;
    }
    requestAnimationFrame(() => (syncing.current = false));
  };

  const setScroll = useCallback(
    (x: number, y: number) => {
      const L = layoutRef.current ?? makeLayout();
      const max = L.maxScroll();
      const s = scroll();
      s.x = Math.max(0, Math.min(max.x, x));
      s.y = Math.max(0, Math.min(max.y, y));
      // Görünüm sol üst hücresi (sayfa değiştirince/kaydederken)
      draw();
      force((n) => n + 1);
    },
    [draw, makeLayout], // eslint-disable-line react-hooks/exhaustive-deps
  );

  useEffect(() => {
    draw();
  });

  // Boyut izleme
  useLayoutEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      const w = el.clientWidth;
      const h = el.clientHeight;
      setSize({ w, h });
      const c = canvasRef.current;
      if (c) {
        c.width = Math.round(w * dpr);
        c.height = Math.round(h * dpr);
        c.style.width = `${w}px`;
        c.style.height = `${h}px`;
      }
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [dpr]);

  // Kopyalama çerçevesi animasyonu
  useEffect(() => {
    if (!ctl.clip) return;
    const t = window.setInterval(() => {
      march.current = (march.current + 1) % 14;
      draw();
    }, 90);
    return () => window.clearInterval(t);
  }, [ctl.clip, draw]);

  // Denetleyici geri çağrıları
  useEffect(() => {
    ctl.scrollIntoView = (row, col) => {
      const L = makeLayout();
      const p = L.scrollFor(Math.min(row, MAX_ROWS - 1), Math.min(col, MAX_COLS - 1));
      const s = scroll();
      if (p.x !== s.x || p.y !== s.y) setScroll(p.x, p.y);
    };
    ctl.focusGrid = () => inputRef.current?.focus({ preventScroll: true });
    ctl.cellClientRect = (row, col) => {
      const L = layoutRef.current;
      const w = wrapRef.current;
      if (!L || !w) return null;
      const m = sheet.mergeAt(row, col);
      const rc = m ? L.cellRect(m.r1, m.c1, m.r2, m.c2) : L.cellRect(row, col);
      const b = w.getBoundingClientRect();
      return { x: b.left + rc.x, y: b.top + rc.y, w: rc.w, h: rc.h };
    };
  });

  // Sayfa değişince odak
  useEffect(() => {
    if (!ctl.edit) inputRef.current?.focus({ preventScroll: true });
  }, [sheet.id]); // eslint-disable-line react-hooks/exhaustive-deps

  /* --------------------------- İsabet testi --------------------------- */

  const hitTest = (x: number, y: number): Hit => hitTestGrid(layoutRef.current!, ctl, sheet, x, y);

  const cellAt = (x: number, y: number): CellPos => {
    const L = layoutRef.current!;
    return { row: L.rowAtY(Math.max(L.headerH, y)), col: L.colAtX(Math.max(L.headerW, x)) };
  };

  /* --------------------------- Fare --------------------------- */

  const local = (e: { clientX: number; clientY: number }) => {
    const r = wrapRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  /** React olayları portal (açılır menü) içinden de kabarcıklanır: yalnız ızgara DOM'undan gelenler işlenir. */
  const fromGrid = (e: { target: EventTarget | null }) => !!wrapRef.current && e.target instanceof Node && wrapRef.current.contains(e.target);

  const onPointerDown = (e: React.PointerEvent) => {
    if (!fromGrid(e)) return;
    if ((e.target as HTMLElement).closest(".tb-float, .tb-editor.is-visible, .tb-assist, .tb-errtag")) return;
    const { x, y } = local(e);
    // Anahat şeritleri (düzey ve +/- düğmeleri)
    if (e.button === 0 && handleOutlineClick(ctl, layoutRef.current!, x, y)) {
      e.preventDefault();
      return;
    }
    // Sayfa Sonu Önizleme: sayfa sonunu ya da yazdırma alanı kenarını sürükle
    const pb = e.button === 0 && !ctl.edit ? pageBreakHit(ctl, layoutRef.current!, x, y) : null;
    if (pb) {
      e.preventDefault();
      wrapRef.current?.setPointerCapture(e.pointerId);
      drag.current = { kind: "pagebreak", pb };
      return;
    }
    const hit = hitTest(x, y);
    setTip(null);
    if (e.button === 2) {
      if (hit.area === "cell" && !ctl.sel.ranges.some((r) => contains(r, hit.row, hit.col))) ctl.selectCell(hit.row, hit.col, false);
      if (hit.area === "colHeader" && !ctl.sel.ranges.some((r) => r.r1 === 0 && r.r2 === MAX_ROWS - 1 && hit.col >= r.c1 && hit.col <= r.c2)) ctl.selectColumns(hit.col, hit.col);
      if (hit.area === "rowHeader" && !ctl.sel.ranges.some((r) => r.c1 === 0 && r.c2 === MAX_COLS - 1 && hit.row >= r.r1 && hit.row <= r.r2)) ctl.selectRows(hit.row, hit.row);
      return;
    }
    if (e.button !== 0) return;
    e.preventDefault();
    // Formül yazarken işaretleme
    if (ctl.edit) {
      if (hit.area === "cell" && ctl.canPoint()) {
        const r = expandForMerges(sheet, rect(hit.row, hit.col));
        ctl.pointTo(r, sheet, { row: hit.row, col: hit.col }, { row: hit.row, col: hit.col });
        drag.current = { kind: "point", anchor: { row: hit.row, col: hit.col }, sheetId: sheet.id };
        wrapRef.current?.setPointerCapture(e.pointerId);
        return;
      }
      if (!ctl.commitEdit(null)) return;
    }
    inputRef.current?.focus({ preventScroll: true });
    wrapRef.current?.setPointerCapture(e.pointerId);
    lastPointer.current = { x, y };
    const add = e.ctrlKey || e.metaKey;
    switch (hit.area) {
      case "corner":
        ctl.selectAll();
        return;
      case "colHeader":
        if (hit.colBorder !== undefined) {
          drag.current = { kind: "colresize", col: hit.colBorder, startX: x, startW: sheet.cols.size(hit.colBorder), saved: sheet.cols.clone() };
          return;
        }
        if (e.shiftKey) ctl.selectColumns(ctl.sel.anchor.col, hit.col);
        else ctl.selectColumns(hit.col, hit.col, add);
        drag.current = { kind: "colsel", start: e.shiftKey ? ctl.sel.anchor.col : hit.col, add };
        return;
      case "rowHeader":
        if (hit.rowBorder !== undefined) {
          drag.current = { kind: "rowresize", row: hit.rowBorder, startY: y, startH: sheet.rows.size(hit.rowBorder), saved: sheet.rows.clone() };
          return;
        }
        if (e.shiftKey) ctl.selectRows(ctl.sel.anchor.row, hit.row);
        else ctl.selectRows(hit.row, hit.row, add);
        drag.current = { kind: "rowsel", start: e.shiftKey ? ctl.sel.anchor.row : hit.row, add };
        return;
      case "cell": {
        if (hit.filterButton) {
          const L = layoutRef.current!;
          const rc = L.cellRect(hit.row, hit.col);
          const wr = wrapRef.current!.getBoundingClientRect();
          onFilterButton(hit.row, hit.col, wr.left + rc.x, wr.top + rc.y + rc.h);
          return;
        }
        if (hit.fillHandle) {
          drag.current = { kind: "fill", src: { ...ctl.range }, target: null };
          return;
        }
        if (hit.selBorder) {
          drag.current = { kind: "move", src: { ...ctl.range }, grab: { row: hit.row, col: hit.col }, target: null, copy: add };
          return;
        }
        // Köprü: Ctrl olmadan tek tıkla açılır (Excel)
        const link = sheet.linkAt(hit.row, hit.col);
        if (link && !e.shiftKey && !add) {
          ctl.selectCell(hit.row, hit.col, false);
          openLink(ctl, link);
          return;
        }
        if (e.shiftKey) {
          ctl.corner = { row: hit.row, col: hit.col };
          const s = ctl.sel;
          ctl.setSelection({ ...s, ranges: [...s.ranges.slice(0, -1), expandForMerges(sheet, rect(s.anchor.row, s.anchor.col, hit.row, hit.col))] }, false);
        } else if (add) {
          const m = sheet.mergeAt(hit.row, hit.col);
          const r = m ?? rect(hit.row, hit.col);
          const a = { row: r.r1, col: r.c1 };
          ctl.corner = a;
          ctl.setSelection({ ranges: [...ctl.sel.ranges, r], active: a, anchor: a }, false);
        } else ctl.selectCell(hit.row, hit.col, false);
        drag.current = { kind: "select", add };
        return;
      }
    }
  };

  const applyDrag = (x: number, y: number) => {
    const d = drag.current;
    if (!d) return;
    const L = layoutRef.current!;
    const z = L.z;
    switch (d.kind) {
      case "select": {
        const p = cellAt(x, y);
        const s = ctl.sel;
        if (p.row === ctl.corner.row && p.col === ctl.corner.col) return;
        ctl.corner = p;
        ctl.setSelection({ ...s, ranges: [...s.ranges.slice(0, -1), expandForMerges(sheet, rect(s.anchor.row, s.anchor.col, p.row, p.col))] }, false);
        return;
      }
      case "colsel": {
        const col = L.colAtX(Math.max(L.headerW, x));
        const s = ctl.sel;
        const r = rect(0, Math.min(d.start, col), MAX_ROWS - 1, Math.max(d.start, col));
        ctl.corner = { row: MAX_ROWS - 1, col };
        ctl.setSelection({ ...s, ranges: [...s.ranges.slice(0, -1), r] }, false);
        return;
      }
      case "rowsel": {
        const row = L.rowAtY(Math.max(L.headerH, y));
        const s = ctl.sel;
        const r = rect(Math.min(d.start, row), 0, Math.max(d.start, row), MAX_COLS - 1);
        ctl.corner = { row, col: MAX_COLS - 1 };
        ctl.setSelection({ ...s, ranges: [...s.ranges.slice(0, -1), r] }, false);
        return;
      }
      case "colresize": {
        const w = Math.max(0, d.startW + (x - d.startX) / z);
        sheet.cols.set(d.col, w);
        setTip({ x: x + 12, y: 4, text: `Genişlik: ${widthPxToChars(w).toLocaleString("tr-TR")} (${Math.round(w)} piksel)` });
        draw();
        return;
      }
      case "rowresize": {
        const h = Math.max(0, d.startH + (y - d.startY) / z);
        sheet.rows.set(d.row, h);
        setTip({ x: L.headerW + 6, y: y + 8, text: `Yükseklik: ${pxToPt(h).toLocaleString("tr-TR")} (${Math.round(h)} piksel)` });
        draw();
        return;
      }
      case "fill": {
        const p = cellAt(x, y);
        const s = d.src;
        let target: Rect | null = null;
        const dRow = p.row > s.r2 ? p.row - s.r2 : p.row < s.r1 ? p.row - s.r1 : 0;
        const dCol = p.col > s.c2 ? p.col - s.c2 : p.col < s.c1 ? p.col - s.c1 : 0;
        if (Math.abs(dRow) >= Math.abs(dCol) && dRow !== 0) target = dRow > 0 ? rect(s.r1, s.c1, p.row, s.c2) : rect(p.row, s.c1, s.r2, s.c2);
        else if (dCol !== 0) target = dCol > 0 ? rect(s.r1, s.c1, s.r2, p.col) : rect(s.r1, p.col, s.r2, s.c2);
        d.target = target;
        if (target) {
          // İpucu: son değer
          setTip({ x: x + 14, y: y + 14, text: fillTipText(ctl, s, target) });
        } else setTip(null);
        draw();
        return;
      }
      case "move": {
        const p = cellAt(x, y);
        const dr = p.row - d.grab.row;
        const dc = p.col - d.grab.col;
        const r1 = Math.max(0, Math.min(MAX_ROWS - rectRows(d.src), d.src.r1 + dr));
        const c1 = Math.max(0, Math.min(MAX_COLS - rectCols(d.src), d.src.c1 + dc));
        d.target = rect(r1, c1, r1 + rectRows(d.src) - 1, c1 + rectCols(d.src) - 1);
        draw();
        return;
      }
      case "point": {
        const p = cellAt(x, y);
        const r = expandForMerges(sheet, rect(d.anchor.row, d.anchor.col, p.row, p.col));
        ctl.pointTo(r, sheet, d.anchor, p);
        return;
      }
      case "pagebreak":
        pageBreakMove(L, d.pb, x, y);
        draw();
        return;
    }
  };

  const onPointerMove = (e: React.PointerEvent) => {
    if (!drag.current && !fromGrid(e)) return;
    const { x, y } = local(e);
    if (drag.current) {
      lastPointer.current = { x, y };
      applyDrag(x, y);
      // Kenarlarda otomatik kaydırma
      const L = layoutRef.current!;
      const out = x < L.cellsX || x > L.width || y < L.cellsY || y > L.height;
      const k = drag.current.kind;
      if (out && (k === "select" || k === "fill" || k === "move" || k === "point" || k === "colsel" || k === "rowsel")) {
        if (autoScroll.current === null) {
          autoScroll.current = window.setInterval(() => {
            const lp = lastPointer.current;
            const LL = layoutRef.current;
            if (!lp || !LL || !drag.current) return;
            const s = scroll();
            let nx = s.x;
            let ny = s.y;
            if (lp.x > LL.width) nx += Math.min(200, (lp.x - LL.width) * 2 + 10);
            else if (lp.x < LL.cellsX && LL.fc === 0) nx -= Math.min(200, (LL.cellsX - lp.x) * 2 + 10);
            if (lp.y > LL.height) ny += Math.min(400, (lp.y - LL.height) * 2 + 10);
            else if (lp.y < LL.cellsY && LL.fr === 0) ny -= Math.min(400, (LL.cellsY - lp.y) * 2 + 10);
            setScroll(nx, ny);
            requestAnimationFrame(() => applyDrag(Math.min(Math.max(lp.x, LL.cellsX + 1), LL.width - 1), Math.min(Math.max(lp.y, LL.cellsY + 1), LL.height - 1)));
          }, 50);
        }
      } else if (autoScroll.current !== null) {
        window.clearInterval(autoScroll.current);
        autoScroll.current = null;
      }
      return;
    }
    // İmleç ve not ipucu
    const hit = hitTest(x, y);
    const cur = (sheet.viewMode === "pageBreak" && pageBreakCursor(ctl, layoutRef.current!, x, y)) || cursorFor(hit, sheet);
    if (cur !== cursor) setCursor(cur);
    if (hit.area === "cell") {
      const m = sheet.mergeAt(hit.row, hit.col);
      const note = sheet.noteAt(m ? m.r1 : hit.row, m ? m.c1 : hit.col);
      const link = sheet.linkAt(hit.row, hit.col);
      if (note) {
        const L = layoutRef.current!;
        const rc = m ? L.cellRect(m.r1, m.c1, m.r2, m.c2) : L.cellRect(hit.row, hit.col);
        setTip({ x: rc.x + rc.w + 8, y: rc.y, text: note.text, title: note.author ? `${note.author}:` : undefined });
      } else if (link) setTip({ x: x + 12, y: y + 18, text: `${link}\nTakip etmek için tıklayın.` });
      else if (tip) setTip(null);
    } else if (tip) setTip(null);
  };

  const finishDrag = (e: React.PointerEvent) => {
    if (autoScroll.current !== null) {
      window.clearInterval(autoScroll.current);
      autoScroll.current = null;
    }
    const d = drag.current;
    drag.current = null;
    try {
      wrapRef.current?.releasePointerCapture(e.pointerId);
    } catch {
      /* yok say */
    }
    if (!d) return;
    switch (d.kind) {
      case "select":
        // Biçim Boyacısı tek tıkla da uygulanır (sürükleme setSelection'da uygulanmış olabilir)
        if (ctl.painter) ctl.applyPainter();
        break;
      case "colresize": {
        const w = sheet.cols.size(d.col);
        sheet.cols.assign(d.saved);
        const items = selectedLines(ctl, "col", d.col);
        ctl.ops.setSizes(sheet, "col", items, w, "Sütun Genişliği");
        ctl.recordRepeat("Sütun Genişliği", () => sizeSelection(ctl, "col", w, "Sütun Genişliği"));
        setTip(null);
        break;
      }
      case "rowresize": {
        const h = sheet.rows.size(d.row);
        sheet.rows.assign(d.saved);
        const items = selectedLines(ctl, "row", d.row);
        ctl.ops.setSizes(sheet, "row", items, h, "Satır Yüksekliği");
        ctl.recordRepeat("Satır Yüksekliği", () => sizeSelection(ctl, "row", h, "Satır Yüksekliği"));
        setTip(null);
        break;
      }
      case "fill":
        setTip(null);
        if (d.target) {
          const shrink = d.target.r2 - d.target.r1 < d.src.r2 - d.src.r1 || d.target.c2 - d.target.c1 < d.src.c2 - d.src.c1;
          if (shrink) {
            // İçeri sürükleme: aradaki hücreleri temizle
            const clearR = d.target.r2 < d.src.r2 ? rect(d.target.r2 + 1, d.src.c1, d.src.r2, d.src.c2) : rect(d.src.r1, d.target.c2 + 1, d.src.r2, d.src.c2);
            ctl.ops.clear(sheet, [clearR], "contents");
          } else ctl.ops.fill(sheet, d.src, d.target, e.ctrlKey ? "copy" : "series");
          ctl.selectRange(d.target, ctl.sel.active);
        }
        break;
      case "pagebreak":
        pageBreakDrop(ctl, d.pb);
        draw();
        break;
      case "move":
        if (d.target && (d.target.r1 !== d.src.r1 || d.target.c1 !== d.src.c1)) {
          const { payload } = buildClipboard(ctl.wb, sheet, d.src, false);
          const dst = pastePayload(ctl.ops, sheet, d.target, payload, { what: "all" }, !d.copy);
          ctl.selectRange(dst);
        }
        break;
    }
    draw();
  };

  const onDoubleClick = (e: React.MouseEvent) => {
    if (!fromGrid(e) || (e.target as Element).closest(".tb-float, .tb-editor.is-visible, .tb-assist, .tb-errtag")) return;
    const { x, y } = local(e);
    const hit = hitTest(x, y);
    if (hit.area === "colHeader" && hit.colBorder !== undefined) {
      const items = selectedLines(ctl, "col", hit.colBorder);
      autoFitCols(ctl, items);
      return;
    }
    if (hit.area === "rowHeader" && hit.rowBorder !== undefined) {
      autoFitRows(ctl, selectedLines(ctl, "row", hit.rowBorder));
      return;
    }
    if (hit.area === "cell") {
      if (hit.fillHandle) {
        E.autoFillDown(ctl);
        return;
      }
      if (hit.filterButton) return;
      if (!ctl.edit) {
        ctl.selectCell(hit.row, hit.col, false);
        ctl.startEdit(undefined);
        requestAnimationFrame(() => {
          const ta = inputRef.current;
          if (ta) {
            ta.focus();
            ta.setSelectionRange(ta.value.length, ta.value.length);
          }
        });
      }
    }
  };

  const onWheel = (e: React.WheelEvent) => {
    if (!fromGrid(e)) return;
    if (e.ctrlKey) {
      // Ctrl+tekerlek: %15 adım (Excel gibi)
      const next = Math.max(10, Math.min(400, sheet.zoom + (e.deltaY < 0 ? 15 : -15)));
      sheet.zoom = next;
      ctl.emit();
      return;
    }
    const L = layoutRef.current!;
    let dx = e.deltaX;
    let dy = e.deltaY;
    if (e.deltaMode === 1) {
      dy *= 20 * L.z;
      dx *= 64 * L.z;
    } else if (e.deltaMode === 2) {
      dy *= L.paneH;
      dx *= L.paneW;
    }
    if (e.shiftKey && !dx) {
      dx = dy;
      dy = 0;
    }
    const s = scroll();
    setScroll(s.x + dx, s.y + dy);
  };

  /* --------------------------- Kaydırma çubukları --------------------------- */

  const onVScroll = () => {
    if (syncing.current) return;
    const vb = vbarRef.current!;
    const s = scroll();
    if (Math.abs(vb.scrollTop - s.y) > 1) setScroll(s.x, vb.scrollTop);
  };
  const onHScroll = () => {
    if (syncing.current) return;
    const hb = hbarRef.current!;
    const s = scroll();
    if (Math.abs(hb.scrollLeft - s.x) > 1) setScroll(hb.scrollLeft, s.y);
  };

  const gridKey = (ev: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const L = layoutRef.current!;
    // Sayfa Aşağı/Yukarı görünümü de bir ekran kaydırır (etkin hücre ekrandaki yerinde kalır)
    if ((ev.key === "PageDown" || ev.key === "PageUp") && !ev.ctrlKey && !ev.metaKey) {
      const d = ev.key === "PageDown" ? 1 : -1;
      const s = scroll();
      if (ev.altKey) setScroll(s.x + d * L.paneW, s.y);
      else setScroll(s.x, s.y + d * L.paneH);
    }
    const handled = handleGridKey(ctl, ev, {
      pageRows: () => Math.max(1, L.visibleRows()[1] - L.visibleRows()[0]),
      pageCols: () => Math.max(1, L.visibleCols()[1] - L.visibleCols()[0]),
    });
    if (handled) ev.preventDefault();
  };

  // Sayfanın görünüm konumunu kaydet
  if (layout) {
    const [r1] = layout.visibleRows();
    const [c1] = layout.visibleCols();
    sheet.scroll = { row: r1, col: c1 };
  }

  return (
    <div className="tb-grid">
      <div className="tb-grid-main">
        <div
          ref={wrapRef}
          className="tb-canvas-wrap"
          style={{ cursor }}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={finishDrag}
          onPointerCancel={finishDrag}
          onDoubleClick={onDoubleClick}
          onWheel={onWheel}
          onPointerLeave={() => setTip(null)}
          onContextMenu={(e) => {
            if (!fromGrid(e) || (e.target as HTMLElement).closest(".tb-float")) return;
            const { x, y } = local(e);
            const hit = hitTest(x, y);
            const kind = hit.area === "colHeader" ? "col" : hit.area === "rowHeader" ? "row" : "cell";
            onContextMenu(e, kind);
          }}
        >
          <canvas ref={canvasRef} className="tb-canvas" />
          <CellEditor ctl={ctl} layout={layout} inputRef={inputRef} onGridKey={gridKey} />
          <GridOverlays ctl={ctl} layout={layout} />
          {tip && (
            <div className="tb-tip" style={{ left: Math.min(tip.x, size.w - 220), top: tip.y }}>
              {tip.title && <strong>{tip.title}</strong>}
              {tip.text.split("\n").map((l, i) => (
                <div key={i}>{l}</div>
              ))}
            </div>
          )}
        </div>
        <div ref={vbarRef} className="tb-vscroll oh-scroll" onScroll={onVScroll}>
          <div style={{ width: 1, height: 1 }} />
        </div>
      </div>
      <div className="tb-bottombar">
        <SheetTabs ctl={ctl} onContextMenu={(e, i) => onContextMenu(e, "tab", i)} />
        <div ref={hbarRef} className="tb-hscroll oh-scroll" onScroll={onHScroll}>
          <div style={{ height: 1, width: 1 }} />
        </div>
      </div>
    </div>
  );
}

export { openLink };
