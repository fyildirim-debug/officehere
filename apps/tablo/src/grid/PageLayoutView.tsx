/**
 * Sayfa Düzeni görünümü: sayfalar kağıt boyutunda, yazdırılacakları yerde
 * (yan yana sütun sayfaları, alt alta satır sayfaları) gösterilir; cetveller,
 * sütun harfleri / satır numaraları, üst/alt bilgi alanları. Hücreler burada
 * seçilip doğrudan düzenlenir (tıkla, yaz, çift tıkla, ok tuşları, kısayollar).
 */
import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { PAGE_SIZES } from "@officehere/core";
import type { Controller } from "../controller";
import { computePageBreaks, expandHeaderFooter, fitScale, printRange, rangeToHtml } from "../io/print";
import type { Rect } from "../model/types";
import { handleEditKey, useFormulaAssist } from "./CellEditor";
import { handleGridKey } from "./keys";
import { HEAD_H, HEAD_W, HfEditor, Headings, RULER, Ruler } from "./PageLayoutParts";
import { setPage } from "../dialogs/PageSetupDialog";
import { titledLists } from "../io/printPages";

const MAX_PAGES = 60;

interface PlPage extends Rect {
  /** Yazdırma sırasındaki numara (önce aşağı, sonra yana). */
  num: number;
}

type MenuEvent = { clientX: number; clientY: number; preventDefault: () => void };

export function PageLayoutView({ ctl, onContextMenu }: { ctl: Controller; onContextMenu?: (e: MenuEvent, kind: "cell") => void }) {
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const drag = useRef<{ row: number; col: number } | null>(null);
  const [hfEdit, setHfEdit] = useState<{ which: "header" | "footer"; num: number } | null>(null);
  const sheet = ctl.sheet;
  const wb = ctl.wb;
  const page = sheet.page;
  const z = sheet.zoom / 100;
  const rootRef = useRef<HTMLDivElement>(null);
  const taRef = useRef<HTMLTextAreaElement>(null);
  const lastActive = useRef("");
  const assistApi = useFormulaAssist(ctl, false);
  const { pages, nCols, total, scale } = useMemo(() => {
    const r = printRange(wb, sheet);
    if (!r) return { pages: [] as PlPage[], nCols: 1, total: 0, scale: 1 };
    const br = computePageBreaks(wb, sheet, r);
    const rowStarts = [r.r1, ...br.rows.slice(0, -1)];
    const colStarts = [r.c1, ...br.cols.slice(0, -1)];
    const bands = Math.max(1, Math.floor(MAX_PAGES / colStarts.length));
    const out: PlPage[] = [];
    rowStarts.slice(0, bands).forEach((r1, ri) => colStarts.forEach((c1, ci) => out.push({ r1, c1, r2: br.rows[ri] - 1, c2: br.cols[ci] - 1, num: ci * rowStarts.length + ri + 1 })));
    return { pages: out, nCols: colStarts.length, total: rowStarts.length * colStarts.length, scale: fitScale(sheet, r) };
  }, [ctl.dataVersion, sheet]); // eslint-disable-line react-hooks/exhaustive-deps
  // Sayfaların HTML'i yalnız içerik değişince yeniden üretilir (düzenlerken her tuşta değil)
  const htmls = useMemo(
    () => pages.map((r) => {
      const t = titledLists(sheet, r);
      return rangeToHtml(wb, sheet, r, { gridlines: page.printGridlines || sheet.showGridlines, headings: false, cellAttrs: true, objects: !t.rowList && !t.colList, ...t });
    }),
    [pages, sheet.showGridlines, page.printGridlines], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const p = PAGE_SIZES[page.paper] ?? PAGE_SIZES.A4;
  const land = page.orientation === "landscape";
  const pw = land ? p.height : p.width;
  const ph = land ? p.width : p.height;
  const m = page.margins;
  const k = z * scale;
  const ruler = ctl.showRuler;
  const heads = sheet.showHeadings;
  const padL = (ruler ? RULER : 0) + (heads ? HEAD_W : 0);
  const padT = (ruler ? RULER : 0) + (heads ? HEAD_H : 0);
  const hf = (t: string | undefined, n: number) => (t ? expandHeaderFooter(t, { page: n, pages: total, sheet: sheet.name, file: ctl.fileName }) : "");
  const e = ctl.edit && ctl.edit.sheetId === sheet.id && !ctl.edit.inBar ? ctl.edit : null;

  // Klavye odağı ve etkin hücreyi görünür tutma bu görünümde
  useLayoutEffect(() => {
    ctl.focusGrid = () => taRef.current?.focus({ preventScroll: true });
    ctl.scrollIntoView = () => {
      lastActive.current = "";
    };
    // Hücrenin ekran dikdörtgeni (HTML hücrelerinden)
    ctl.cellClientRect = (row, col) => {
      const td = rootRef.current?.querySelector<HTMLElement>(`td[data-r="${row}"][data-c="${col}"]`);
      if (!td) return null;
      const b = td.getBoundingClientRect();
      return { x: b.left, y: b.top, w: b.width, h: b.height };
    };
    if (!document.activeElement || document.activeElement === document.body) taRef.current?.focus({ preventScroll: true });
  }, [ctl]);

  // Seçim vurgusu ve düzenleyicinin konumu (hücreler HTML olarak çizildiği için DOM üzerinden)
  useLayoutEffect(() => {
    const root = rootRef.current;
    const ta = taRef.current;
    if (!root || !ta) return;
    const { active, ranges } = ctl.sel;
    let target: HTMLElement | null = null;
    root.querySelectorAll<HTMLElement>("td[data-r]").forEach((td) => {
      const r = Number(td.dataset.r);
      const c = Number(td.dataset.c);
      const isActive = r === active.row && c === active.col;
      if (isActive && !target) target = td;
      td.classList.toggle("is-active", isActive);
      td.classList.toggle("is-sel", !isActive && ranges.some((s) => r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2));
    });
    const key = `${sheet.id}:${active.row}:${active.col}`;
    const td = target as HTMLElement | null;
    if (!td) {
      ta.style.cssText = "left:0;top:0";
      return;
    }
    const rr = root.getBoundingClientRect();
    const tr = td.getBoundingClientRect();
    ta.style.left = `${tr.left - rr.left + root.scrollLeft}px`;
    ta.style.top = `${tr.top - rr.top + root.scrollTop}px`;
    ta.style.width = `${tr.width}px`;
    ta.style.height = `${tr.height}px`;
    ta.style.fontSize = `${11 * k * (96 / 72)}px`;
    if (key !== lastActive.current) {
      lastActive.current = key;
      td.scrollIntoView({ block: "nearest", inline: "nearest" });
    }
  });

  const cellOf = (t: EventTarget) => {
    const td = (t as HTMLElement).closest?.("td[data-r]") as HTMLElement | null;
    return td ? { row: Number(td.dataset.r), col: Number(td.dataset.c) } : null;
  };
  const onPointerDown = (ev: React.PointerEvent) => {
    const c = cellOf(ev.target);
    if (!c || ev.button !== 0) return;
    ev.preventDefault();
    if (ctl.edit && !ctl.commitEdit(null)) return;
    if (ev.shiftKey) {
      const a = ctl.sel.anchor;
      ctl.selectRange({ r1: Math.min(a.row, c.row), c1: Math.min(a.col, c.col), r2: Math.max(a.row, c.row), c2: Math.max(a.col, c.col) }, a);
    } else ctl.selectCell(c.row, c.col, false);
    // Sürükleyerek aralık seçimi
    drag.current = ev.shiftKey ? ctl.sel.anchor : c;
    taRef.current?.focus({ preventScroll: true });
  };
  const onPointerMove = (ev: React.PointerEvent) => {
    const a = drag.current;
    if (!a || !(ev.buttons & 1)) return;
    const c = cellOf(document.elementFromPoint(ev.clientX, ev.clientY) ?? ev.target);
    if (!c) return;
    const r = { r1: Math.min(a.row, c.row), c1: Math.min(a.col, c.col), r2: Math.max(a.row, c.row), c2: Math.max(a.col, c.col) };
    const cur = ctl.range;
    if (cur.r1 !== r.r1 || cur.c1 !== r.c1 || cur.r2 !== r.r2 || cur.c2 !== r.c2) ctl.selectRange(r, a);
  };
  const onContext = (ev: React.MouseEvent) => {
    const c = cellOf(ev.target);
    if (!c || !onContextMenu) return;
    ev.preventDefault();
    if (ctl.edit) ctl.commitEdit(null);
    if (!ctl.sel.ranges.some((s) => c.row >= s.r1 && c.row <= s.r2 && c.col >= s.c1 && c.col <= s.c2)) ctl.selectCell(c.row, c.col, false);
    onContextMenu(ev, "cell");
  };
  const margins = (patch: Partial<typeof m>) => setPage(ctl, { margins: { ...m, ...patch } }, "Kenar Boşlukları");
  const onDoubleClick = (ev: React.MouseEvent) => {
    const c = cellOf(ev.target);
    if (!c || ctl.edit) return;
    ctl.selectCell(c.row, c.col, false);
    ctl.startEdit();
    requestAnimationFrame(() => {
      const ta = taRef.current;
      if (ta) ta.setSelectionRange(ta.value.length, ta.value.length);
    });
  };

  return (
    <div ref={rootRef} className="tb-pl oh-scroll" onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={() => (drag.current = null)} onDoubleClick={onDoubleClick} onContextMenu={onContext}>
      <div className="tb-pl-pages" style={{ gridTemplateColumns: `repeat(${Math.max(1, nCols)}, max-content)` }}>
        {pages.length === 0 && (
          <div className="tb-pl-page" style={{ width: pw * z, height: ph * z }}>
            <div className="tb-pl-empty">Bu sayfada yazdırılacak bir şey yok. Veri eklemek için Normal görünümde yazın.</div>
          </div>
        )}
        {pages.map((r, i) => (
          <div key={`${r.r1}:${r.c1}`} className="tb-pl-wrap" style={{ paddingLeft: padL, paddingTop: padT }}>
            {ruler && (
              <div style={{ position: "absolute", left: padL, top: 0 }}>
                <Ruler length={pw} z={z} start={m.left * 96} end={m.right * 96} onChange={(a, b) => margins({ left: a / 96, right: b / 96 })} />
              </div>
            )}
            {ruler && (
              <div style={{ position: "absolute", left: 0, top: padT }}>
                <Ruler length={ph} z={z} start={m.top * 96} end={m.bottom * 96} vertical onChange={(a, b) => margins({ top: a / 96, bottom: b / 96 })} />
              </div>
            )}
            {heads && (
              <div style={{ position: "absolute", left: padL + m.left * 96 * z, top: padT - HEAD_H }}>
                <Headings sheet={sheet} r={r} k={k} axis="col" active={ctl.sel.ranges} />
              </div>
            )}
            {heads && (
              <div style={{ position: "absolute", left: padL - HEAD_W, top: padT + m.top * 96 * z, width: HEAD_W }}>
                <Headings sheet={sheet} r={r} k={k} axis="row" active={ctl.sel.ranges} />
              </div>
            )}
            <div className="tb-pl-page" style={{ width: pw * z, height: ph * z }}>
              <div className="tb-pl-hf is-top" style={{ top: m.header * 96 * z, fontSize: 11 * z }} onPointerDown={(ev) => ev.stopPropagation()} onClick={() => setHfEdit({ which: "header", num: r.num })}>
                {hfEdit?.which === "header" && hfEdit.num === r.num ? (
                  <HfEditor value={page.header ?? ""} onDone={(v) => (setHfEdit(null), v !== null && v !== (page.header ?? "") && setPage(ctl, { header: v || undefined }, "Üst Bilgi"), ctl.focusGrid())} />
                ) : (
                  hf(page.header, r.num) || <span className="tb-pl-hf-empty">Üst bilgi eklemek için tıklayın</span>
                )}
              </div>
              <div className="tb-pl-body" style={{ left: m.left * 96 * z, top: m.top * 96 * z, width: (pw - (m.left + m.right) * 96) * z, height: (ph - (m.top + m.bottom) * 96) * z }}>
                <div style={{ zoom: k }} dangerouslySetInnerHTML={{ __html: htmls[i] }} />
              </div>
              <div className="tb-pl-hf is-bottom" style={{ bottom: m.footer * 96 * z, fontSize: 11 * z }} onPointerDown={(ev) => ev.stopPropagation()} onClick={() => setHfEdit({ which: "footer", num: r.num })}>
                {hfEdit?.which === "footer" && hfEdit.num === r.num ? (
                  <HfEditor value={page.footer ?? ""} onDone={(v) => (setHfEdit(null), v !== null && v !== (page.footer ?? "") && setPage(ctl, { footer: v || undefined }, "Alt Bilgi"), ctl.focusGrid())} />
                ) : (
                  hf(page.footer, r.num) || <span className="tb-pl-hf-empty">Alt bilgi eklemek için tıklayın</span>
                )}
              </div>
              <div className="tb-pl-num">
                Sayfa {r.num} / {total}
              </div>
            </div>
          </div>
        ))}
      </div>
      <textarea
        ref={taRef}
        className={`tb-pl-editor${e ? " is-editing" : ""}`}
        aria-label="Hücre düzenleyici"
        spellCheck={false}
        value={e ? e.text : ""}
        onChange={(ev) => {
          const ta = ev.currentTarget;
          if (!ctl.edit) {
            ctl.startEdit(ta.value);
            if (ctl.edit) ctl.updateEdit(ta.value, ta.value.length);
            return;
          }
          ctl.updateEdit(ta.value, ta.selectionStart);
        }}
        onKeyDown={(ev) => {
          if (ctl.edit && !ctl.edit.inBar) {
            if (handleEditKey(ctl, ev, assistApi)) ev.preventDefault();
            return;
          }
          if (handleGridKey(ctl, ev, { pageRows: () => 20, pageCols: () => 8 })) ev.preventDefault();
        }}
      />
    </div>
  );
}
