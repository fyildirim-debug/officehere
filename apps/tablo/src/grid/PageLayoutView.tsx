/** Sayfa Düzeni görünümü: sayfalar kağıt boyutunda, kenar boşlukları ve üst/alt bilgiyle gösterilir. */
import { useMemo, useSyncExternalStore } from "react";
import { PAGE_SIZES } from "@officehere/core";
import type { Controller } from "../controller";
import { computePageBreaks, expandHeaderFooter, fitScale, printRange, rangeToHtml } from "../io/print";

const MAX_PAGES = 60;

export function PageLayoutView({ ctl }: { ctl: Controller }) {
  const version = useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const sheet = ctl.sheet;
  const wb = ctl.wb;
  const page = sheet.page;
  const z = sheet.zoom / 100;
  const pages = useMemo(() => {
    const r = printRange(wb, sheet);
    if (!r) return [];
    const br = computePageBreaks(wb, sheet);
    const out: { r1: number; c1: number; r2: number; c2: number }[] = [];
    let r1 = r.r1;
    for (const rEnd of br.rows) {
      let c1 = r.c1;
      for (const cEnd of br.cols) {
        out.push({ r1, c1, r2: rEnd - 1, c2: cEnd - 1 });
        c1 = cEnd;
        if (out.length >= MAX_PAGES) return out;
      }
      r1 = rEnd;
    }
    return out;
  }, [version, sheet]); // eslint-disable-line react-hooks/exhaustive-deps
  const p = PAGE_SIZES[page.paper] ?? PAGE_SIZES.A4;
  const land = page.orientation === "landscape";
  const pw = land ? p.height : p.width;
  const ph = land ? p.width : p.height;
  const m = page.margins;
  const scale = pages.length ? fitScale(sheet, pages[0]) : 1;
  const hf = (t: string | undefined, i: number) => (t ? expandHeaderFooter(t, { page: i + 1, pages: pages.length, sheet: sheet.name, file: ctl.fileName }) : "");

  const onDbl = (e: React.MouseEvent) => {
    const td = (e.target as HTMLElement).closest("td[data-r]") as HTMLElement | null;
    if (!td) return;
    sheet.viewMode = "normal";
    ctl.selectCell(Number(td.dataset.r), Number(td.dataset.c));
    ctl.startEdit();
  };

  return (
    <div className="tb-pl oh-scroll" onDoubleClick={onDbl}>
      <div className="tb-pl-hint">Sayfa Düzeni görünümü. Bir hücreyi düzenlemek için çift tıklayın; Normal görünüme geçilir.</div>
      {pages.length === 0 && (
        <div className="tb-pl-page" style={{ width: pw * z, height: ph * z }}>
          <div className="tb-pl-empty">Bu sayfada yazdırılacak bir şey yok.</div>
        </div>
      )}
      {pages.map((r, i) => (
        <div key={i} className="tb-pl-page" style={{ width: pw * z, height: ph * z }}>
          <div className="tb-pl-hf is-top" style={{ top: m.header * 96 * z, fontSize: 11 * z }} onClick={() => ctl.openDialog("pageSetup", { tab: "hf" })}>
            {hf(page.header, i) || <span className="tb-pl-hf-empty">Üst bilgi eklemek için tıklayın</span>}
          </div>
          <div
            className="tb-pl-body"
            style={{ left: m.left * 96 * z, top: m.top * 96 * z, width: (pw - (m.left + m.right) * 96) * z, height: (ph - (m.top + m.bottom) * 96) * z }}
          >
            <div style={{ zoom: z * scale }} dangerouslySetInnerHTML={{ __html: rangeToHtml(wb, sheet, r, { gridlines: page.printGridlines || sheet.showGridlines, headings: page.printHeadings, cellAttrs: true, objects: true }) }} />
          </div>
          <div className="tb-pl-hf is-bottom" style={{ bottom: m.footer * 96 * z, fontSize: 11 * z }} onClick={() => ctl.openDialog("pageSetup", { tab: "hf" })}>
            {hf(page.footer, i) || <span className="tb-pl-hf-empty">Alt bilgi eklemek için tıklayın</span>}
          </div>
          <div className="tb-pl-num" style={{ fontSize: 10 }}>
            Sayfa {i + 1} / {pages.length}
          </div>
        </div>
      ))}
    </div>
  );
}
