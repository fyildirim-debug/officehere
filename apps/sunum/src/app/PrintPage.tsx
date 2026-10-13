/**
 * Dosya › Yazdır sayfası: yazdırılacak slaytlar (tümü, seçim, geçerli, özel aralık),
 * düzen (tam sayfa slaytlar, not sayfaları, ana hat, dinleyici notları), çerçeve,
 * gizli slaytlar, yönlendirme ve renk; sağda sayfa sayfa gezinilen baskı önizlemesi.
 * Düzen ayarları sunuyla saklanır (p:prnPr).
 */
import { useLayoutEffect, useMemo, useRef, useState } from "react";
import { ChevronLeft16Regular, ChevronRight16Regular, DocumentPdf24Regular, Print24Regular } from "@fluentui/react-icons";
import { BackstageHeading } from "@officehere/ui";
import { isDesktop } from "@officehere/core";
import { DEFAULT_PRINT_SETTINGS } from "../io/presProps";
import type { PrintSettings, PrintWhat } from "../model/types";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { PrintPageView } from "./PrintPages";
import { buildPrintPages, paperSize, parseSlideRange, printSlideIndices, type PrintJob, type PrintScope } from "./printLayout";

/** Düzen seçenekleri: değer "düzen" ya da "düzen:sıra". */
const LAYOUTS: { value: string; label: string; group: string }[] = [
  { value: "slides", label: "Tam Sayfa Slaytlar", group: "Yazdırma Düzeni" },
  { value: "notes", label: "Not Sayfaları", group: "Yazdırma Düzeni" },
  { value: "outline", label: "Ana Hat", group: "Yazdırma Düzeni" },
  { value: "handouts1", label: "1 Slayt", group: "Dinleyici Notları" },
  { value: "handouts2", label: "2 Slayt", group: "Dinleyici Notları" },
  { value: "handouts3", label: "3 Slayt (not satırlarıyla)", group: "Dinleyici Notları" },
  { value: "handouts4:horz", label: "4 Slayt Yatay", group: "Dinleyici Notları" },
  { value: "handouts4:vert", label: "4 Slayt Dikey", group: "Dinleyici Notları" },
  { value: "handouts6:horz", label: "6 Slayt Yatay", group: "Dinleyici Notları" },
  { value: "handouts6:vert", label: "6 Slayt Dikey", group: "Dinleyici Notları" },
  { value: "handouts9:horz", label: "9 Slayt Yatay", group: "Dinleyici Notları" },
  { value: "handouts9:vert", label: "9 Slayt Dikey", group: "Dinleyici Notları" },
];

const ORDERED = new Set(["handouts4", "handouts6", "handouts9"]);

function layoutValue(s: PrintSettings): string {
  return ORDERED.has(s.what) ? `${s.what}:${s.order ?? "horz"}` : s.what;
}

/** Önizleme kutusunun en büyük boyutu (px). */
const PREVIEW_MAX = { w: 620, h: 600 };

export function PrintPage({ state, actions: a, onPrint, onHeaderFooter }: { state: AppState; actions: Actions; onPrint: (job: PrintJob) => void; onHeaderFooter: () => void }) {
  const pres = state.pres;
  const settings = pres.printSettings ?? DEFAULT_PRINT_SETTINGS;
  const [scopeKind, setScopeKind] = useState<PrintScope["kind"]>("all");
  const [range, setRange] = useState("");
  const [page, setPage] = useState(0);
  const boxRef = useRef<HTMLDivElement>(null);
  const [avail, setAvail] = useState(PREVIEW_MAX);
  useLayoutEffect(() => {
    const el = boxRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setAvail({ w: Math.max(200, Math.min(PREVIEW_MAX.w, el.clientWidth - 24)), h: PREVIEW_MAX.h }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const set = (patch: Partial<PrintSettings>) => {
    a.store.commit((p) => ({ ...p, printSettings: { ...(p.printSettings ?? DEFAULT_PRINT_SETTINGS), ...patch } }), { merge: "print-settings" });
    setPage(0);
  };

  const selection = useMemo(() => state.slideSel.map((id) => pres.slides.findIndex((s) => s.id === id)).filter((i) => i >= 0), [state.slideSel, pres.slides]);
  const scope: PrintScope = scopeKind === "custom" ? { kind: "custom", text: range } : { kind: scopeKind } as PrintScope;
  const rangeInvalid = scopeKind === "custom" && range.trim() !== "" && parseSlideRange(range, pres.slides.length) === null;
  const indices = printSlideIndices(pres, settings, scope, state.slideIdx, selection);
  const pages = useMemo(() => buildPrintPages(pres, settings, indices), [pres, settings, indices.join(",")]); // eslint-disable-line react-hooks/exhaustive-deps
  const size = paperSize(pres, settings.what, settings.orientation);
  const k = Math.min(avail.w / size.w, avail.h / size.h);
  const cur = Math.min(page, Math.max(0, pages.length - 1));
  const hasHidden = pres.slides.some((s) => s.hidden);
  const job = (pdf: boolean): PrintJob => ({ settings, indices, pdf });
  const empty = !pages.length || rangeInvalid;

  return (
    <>
      <BackstageHeading>Yazdır</BackstageHeading>
      <div className="su-printpage">
        <div className="su-printpage-side">
          <div className="su-printpage-top">
            <button type="button" className="su-printpage-btn" disabled={empty} onClick={() => onPrint(job(false))} aria-label="Yazdır">
              <Print24Regular />
              <span>Yazdır</span>
            </button>
            <button type="button" className="su-printpage-btn" disabled={empty} onClick={() => onPrint(job(true))} aria-label="PDF Olarak Kaydet" title={isDesktop() ? "Seçili düzenle PDF dosyası oluşturun." : "Yazdırma penceresinde hedef olarak “PDF olarak kaydet” seçin."}>
              <DocumentPdf24Regular />
              <span>PDF Olarak Kaydet</span>
            </button>
          </div>
          <div className="su-printpage-label">Ayarlar</div>
          <select value={scopeKind} aria-label="Yazdırılacak slaytlar" onChange={(e) => (setScopeKind(e.target.value as PrintScope["kind"]), setPage(0))}>
            <option value="all">Tüm Slaytları Yazdır</option>
            <option value="selection">Seçimi Yazdır</option>
            <option value="current">Geçerli Slaytı Yazdır</option>
            <option value="custom">Özel Aralık</option>
          </select>
          <label className="su-printpage-range">
            <span>Slaytlar:</span>
            <input
              type="text"
              value={range}
              placeholder="ör. 1-3, 5"
              aria-label="Slayt aralığı"
              aria-invalid={rangeInvalid || undefined}
              onChange={(e) => {
                setRange(e.target.value);
                setScopeKind(e.target.value.trim() ? "custom" : "all");
                setPage(0);
              }}
            />
          </label>
          {rangeInvalid && <div className="su-printpage-error">Geçersiz slayt aralığı. Örnek: 1-3, 5</div>}
          <select
            value={layoutValue(settings)}
            aria-label="Yazdırma düzeni"
            onChange={(e) => {
              const [what, order] = e.target.value.split(":") as [PrintWhat, "horz" | "vert" | undefined];
              set({ what, order });
            }}
          >
            {["Yazdırma Düzeni", "Dinleyici Notları"].map((g) => (
              <optgroup key={g} label={g}>
                {LAYOUTS.filter((l) => l.group === g).map((l) => (
                  <option key={l.value} value={l.value}>
                    {l.group === "Dinleyici Notları" ? `Dinleyici Notları: ${l.label}` : l.label}
                  </option>
                ))}
              </optgroup>
            ))}
          </select>
          <label className="su-printpage-check">
            <input type="checkbox" checked={settings.frameSlides} onChange={(e) => set({ frameSlides: e.target.checked })} /> Slaytları Çerçevele
          </label>
          <label className="su-printpage-check" title={hasHidden ? undefined : "Sunuda gizli slayt yok."}>
            <input type="checkbox" checked={settings.hiddenSlides} disabled={!hasHidden} onChange={(e) => set({ hiddenSlides: e.target.checked })} /> Gizli Slaytları Yazdır
          </label>
          <select value={settings.orientation} aria-label="Yönlendirme" disabled={settings.what === "slides"} onChange={(e) => set({ orientation: e.target.value as PrintSettings["orientation"] })}>
            <option value="portrait">Dikey Yönlendirme</option>
            <option value="landscape">Yatay Yönlendirme</option>
          </select>
          <select value={settings.color} aria-label="Renk" onChange={(e) => set({ color: e.target.value as PrintSettings["color"] })}>
            <option value="clr">Renkli</option>
            <option value="gray">Gri Tonlama</option>
            <option value="bw">Saf Siyah ve Beyaz</option>
          </select>
          <button type="button" className="su-printpage-link" onClick={onHeaderFooter}>
            Üst Bilgi ve Alt Bilgiyi Düzenle
          </button>
        </div>
        <div className="su-printpage-preview" ref={boxRef}>
          {empty ? (
            <div className="print-prev-empty">{rangeInvalid ? "Slayt aralığı geçersiz." : "Yazdırılacak slayt yok."}</div>
          ) : (
            <div className="su-printpage-paper" style={{ width: size.w * k, height: size.h * k }} aria-label={`Sayfa ${cur + 1} önizlemesi`}>
              <div style={{ transform: `scale(${k})`, transformOrigin: "0 0", width: size.w, height: size.h }}>
                <PrintPageView pres={pres} page={pages[cur]} pageNo={cur + 1} size={size} settings={settings} />
              </div>
            </div>
          )}
          <div className="print-prev-nav">
            <button type="button" aria-label="Önceki Sayfa" disabled={cur <= 0} onClick={() => setPage(cur - 1)}>
              <ChevronLeft16Regular />
            </button>
            <span>
              {pages.length && !rangeInvalid ? cur + 1 : 0} / {rangeInvalid ? 0 : pages.length}
            </span>
            <button type="button" aria-label="Sonraki Sayfa" disabled={cur >= pages.length - 1} onClick={() => setPage(cur + 1)}>
              <ChevronRight16Regular />
            </button>
          </div>
        </div>
      </div>
    </>
  );
}
