// Baskı önizlemesi: yazdırılacak yapraklar (sayfa aralığı ve yaprak başına sayfa uygulanmış) gerçek sayfa
// kopyalarıyla çizilir; altta sayfa gezintisi ve yakınlaştırma.
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { ChevronLeft20Regular, ChevronRight20Regular, ZoomFit20Regular } from "@fluentui/react-icons";
import { assembleSheets } from "../views/printRender";
import type { PageInfo } from "../editor/pagination";
import { printSheets, selectedPages, type PrintOptions } from "./printOptions";

interface Props {
  preview: { pages: HTMLElement[]; infos: PageInfo[] } | null;
  options: PrintOptions;
  pageCount: number;
  currentPage: number;
}

export function PrintPreview({ preview, options, pageCount, currentPage }: Props) {
  const sel = selectedPages(options, pageCount, currentPage);
  const list = typeof sel === "string" ? [] : sel;
  const key = list.join(",");
  // Önizleme her yaprağı bir kez gösterir (kopyalar ve iki taraflı sıra hariç).
  const sheets = useMemo(() => printSheets(list, { ...options, copies: 1, duplex: "single" }), [key, options.perSheet]); // eslint-disable-line react-hooks/exhaustive-deps
  const [idx, setIdx] = useState(() => Math.max(0, sheets.findIndex((s) => s.includes(currentPage))));
  const [zoom, setZoom] = useState<number | null>(null);
  const [fit, setFit] = useState(0.5);
  const hostRef = useRef<HTMLDivElement>(null);
  const areaRef = useRef<HTMLDivElement>(null);
  const cur = Math.min(idx, Math.max(0, sheets.length - 1));

  useEffect(() => {
    if (idx >= sheets.length) setIdx(Math.max(0, sheets.length - 1));
  }, [sheets.length, idx]);

  useLayoutEffect(() => {
    const host = hostRef.current;
    const area = areaRef.current;
    if (!host || !area) return;
    host.innerHTML = "";
    const sheet = sheets[cur];
    if (!preview || !sheet || sheet.some((i) => !preview.pages[i])) return;
    const el = assembleSheets(preview.pages, preview.infos, [sheet], options.perSheet)[0];
    host.appendChild(el);
    const w = parseFloat(el.style.width) || 794;
    const h = parseFloat(el.style.height) || 1123;
    const f = Math.min((area.clientWidth - 40) / w, (area.clientHeight - 40) / h);
    setFit(Math.max(0.1, Math.min(2, f)));
    host.style.width = `${w}px`;
    host.style.height = `${h}px`;
  }, [preview, sheets, cur, options.perSheet]);

  const scale = zoom ?? fit;
  return (
    <div className="bl-pr-preview">
      <div className="bl-pr-area" ref={areaRef}>
        {sheets.length ? (
          <div className="bl-pr-sheet" ref={hostRef} style={{ zoom: scale }} data-sheet={cur + 1} />
        ) : (
          <div className="bl-pr-empty">Yazdırılacak sayfa yok.</div>
        )}
      </div>
      <div className="bl-pr-bar">
        <div className="bl-pr-nav">
          <button type="button" aria-label="Önceki Sayfa" disabled={cur <= 0} onClick={() => setIdx(cur - 1)}>
            <ChevronLeft20Regular />
          </button>
          <input
            type="number"
            aria-label="Geçerli sayfa"
            min={1}
            max={Math.max(1, sheets.length)}
            value={sheets.length ? cur + 1 : 0}
            onChange={(e) => setIdx(Math.max(0, Math.min(sheets.length - 1, Number(e.target.value) - 1)))}
          />
          <span className="bl-pr-count">/ {sheets.length}</span>
          <button type="button" aria-label="Sonraki Sayfa" disabled={cur >= sheets.length - 1} onClick={() => setIdx(cur + 1)}>
            <ChevronRight20Regular />
          </button>
        </div>
        <div className="bl-pr-zoom">
          <button type="button" aria-label="Uzaklaştır" onClick={() => setZoom(Math.max(0.1, +(scale - 0.1).toFixed(2)))}>
            −
          </button>
          <input type="range" aria-label="Yakınlaştırma" min={10} max={200} value={Math.round(scale * 100)} onChange={(e) => setZoom(Number(e.target.value) / 100)} />
          <button type="button" aria-label="Yakınlaştır" onClick={() => setZoom(Math.min(2, +(scale + 0.1).toFixed(2)))}>
            +
          </button>
          <span className="bl-pr-pct">%{Math.round(scale * 100)}</span>
          <button type="button" aria-label="Sayfaya Yakınlaştır" title="Sayfaya Yakınlaştır" onClick={() => setZoom(null)}>
            <ZoomFit20Regular />
          </button>
        </div>
      </div>
    </div>
  );
}
