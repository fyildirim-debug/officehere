/**
 * Yazdırma sayfalarının çizimi (baskı önizlemesi ve yazdırma/PDF ortak):
 * tam sayfa slaytlar, not sayfaları, ana hat ve dinleyici notları; not ve
 * dinleyici sayfalarında üst bilgi, tarih, alt bilgi ve sayfa numarası.
 */
import { useEffect } from "react";
import { createPortal } from "react-dom";
import { DEFAULT_NOTES_HF } from "../model/presentation";
import type { NotesHeaderFooter, Presentation, PrintSettings } from "../model/types";
import { ScaledSlide, SlideView } from "../render/SlideView";
import { NOTE_LINE_GAP, notesPageBoxes, PAGE_MARGIN, type PrintPageDesc } from "./printLayout";

interface PageProps {
  pres: Presentation;
  page: PrintPageDesc;
  /** 1 tabanlı sayfa numarası (not/dinleyici/ana hat sayfalarında basılır). */
  pageNo: number;
  size: { w: number; h: number };
  settings: PrintSettings;
}

const today = () => new Date().toLocaleDateString("tr-TR");

/** Not/dinleyici/ana hat sayfasının üst ve alt bilgi şeritleri. */
function HeaderFooterBands({ hf, pageNo }: { hf: NotesHeaderFooter; pageNo: number }) {
  const date = hf.dateTime ? (hf.dateAuto ? today() : hf.dateFixed) : "";
  return (
    <>
      <div className="pp-hf is-top">
        <span>{hf.header ? hf.headerText : ""}</span>
        <span>{date}</span>
      </div>
      <div className="pp-hf is-bottom">
        <span>{hf.footer ? hf.footerText : ""}</span>
        <span className="pp-pageno">{hf.pageNumber ? pageNo : ""}</span>
      </div>
    </>
  );
}

const boxStyle = (b: { x: number; y: number; w: number; h: number }) => ({ left: b.x, top: b.y, width: b.w, height: b.h });

/** Tek bir yazdırma sayfası (doğal boyutunda; önizlemede dışarıdan ölçeklenir). */
export function PrintPageView({ pres, page, pageNo, size, settings }: PageProps) {
  const hf = pres.notesHeaderFooter ?? DEFAULT_NOTES_HF;
  const ratio = pres.width / pres.height;
  const cls = `pp-page is-${page.kind}${settings.color !== "clr" ? ` is-${settings.color}` : ""}${settings.frameSlides ? " is-framed" : ""}`;
  if (page.kind === "slide") {
    const s = pres.slides[page.index];
    return (
      <div className={cls} style={{ width: size.w, height: size.h }} data-print-page={pageNo}>
        <div className="pp-content">
          <SlideView pres={pres} slide={s} slideNumber={page.index + 1} mode="print" />
        </div>
        {settings.frameSlides && <div className="pp-frame" />}
      </div>
    );
  }
  let body: React.ReactNode = null;
  if (page.kind === "notes") {
    const s = pres.slides[page.index];
    const b = notesPageBoxes(size, ratio);
    body = (
      <>
        <div className="pp-slide is-bordered" style={boxStyle(b.slide)}>
          <ScaledSlide pres={pres} slide={s} slideNumber={page.index + 1} width={b.slide.w} mode="print" />
        </div>
        <div className="pp-notes" style={boxStyle(b.notes)}>
          {s.notes.split("\n").map((l, i) => (
            <p key={i}>{l || " "}</p>
          ))}
        </div>
      </>
    );
  } else if (page.kind === "handout") {
    body = page.cells.map((c) => (
      <div key={c.index} className="pp-cell">
        <div className="pp-slide" style={boxStyle(c.slide)} data-slide-index={c.index}>
          <ScaledSlide pres={pres} slide={pres.slides[c.index]} slideNumber={c.index + 1} width={c.slide.w} mode="print" />
        </div>
        {c.lines && (
          <div className="pp-lines" style={boxStyle(c.lines)}>
            {Array.from({ length: Math.max(1, Math.floor(c.lines.h / NOTE_LINE_GAP)) }, (_, i) => (
              <div key={i} className="pp-line" style={{ height: NOTE_LINE_GAP }} />
            ))}
          </div>
        )}
      </div>
    ));
  } else if (page.kind === "outline") {
    body = (
      <div className="pp-outline" style={{ left: PAGE_MARGIN.x, top: PAGE_MARGIN.top, width: size.w - PAGE_MARGIN.x * 2 }}>
        {page.blocks.map((b) => (
          <div key={b.index} className="pp-outline-slide">
            <div className="pp-outline-title">
              <span className="pp-outline-num">{b.index + 1}</span>
              <span className="pp-outline-icon" aria-hidden="true" />
              <span>{b.title}</span>
            </div>
            {b.lines.map((l, i) => (
              <div key={i} className="pp-outline-line" style={{ paddingLeft: 52 + l.level * 24 }}>
                {l.text}
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  }
  // Not sayfasının numarası slayt numarasıdır; diğerlerinde basılan sayfanın sırası.
  return (
    <div className={cls} style={{ width: size.w, height: size.h }} data-print-page={pageNo}>
      <HeaderFooterBands hf={hf} pageNo={page.kind === "notes" ? page.index + 1 : pageNo} />
      {body}
    </div>
  );
}

/** Yazdırma kökü: yalnız yazdırmada görünür; @page boyutu kâğıt boyutudur. */
export function PrintRoot({
  pres,
  pages,
  size,
  settings,
  onDone,
  onReady,
}: {
  pres: Presentation;
  pages: PrintPageDesc[];
  size: { w: number; h: number };
  settings: PrintSettings;
  onDone: () => void;
  onReady?: () => Promise<void>;
}) {
  useEffect(() => {
    // Etki temizlenince (kapanma ya da geliştirmede StrictMode'un ikinci kurulumu) bekleyen
    // zincir iptal edilir: aksi halde ilk kurulumun zinciri de yazdırır (iki kez yazdırma).
    let cancelled = false;
    let finished = false;
    let timer: number | undefined;
    const style = document.createElement("style");
    style.textContent = `@page { size: ${(size.w / 96).toFixed(3)}in ${(size.h / 96).toFixed(3)}in; margin: 0; }`;
    document.head.appendChild(style);
    document.documentElement.classList.add("is-printing");
    const after = () => {
      if (cancelled || finished) return;
      finished = true;
      onDone();
    };
    window.addEventListener("afterprint", after);
    const imgs = Array.from(document.querySelectorAll<HTMLImageElement>(".print-root img"));
    void Promise.all(imgs.map((i) => (i.complete ? Promise.resolve() : new Promise((r) => ((i.onload = r), (i.onerror = r)))))).then(() =>
      requestAnimationFrame(() =>
        requestAnimationFrame(() => {
          if (cancelled) return;
          if (onReady) {
            // Masaüstü: yazdırma CSS'iyle doğrudan PDF (pencere açmadan).
            void onReady().finally(after);
            return;
          }
          window.print();
          // Bazı tarayıcılar afterprint göndermez.
          timer = window.setTimeout(after, 500);
        }),
      ),
    );
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
      window.removeEventListener("afterprint", after);
      style.remove();
      document.documentElement.classList.remove("is-printing");
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return createPortal(
    <div className="print-root">
      {pages.map((p, i) => (
        <div key={i} className="print-page" style={{ width: size.w, height: size.h }}>
          <PrintPageView pres={pres} page={p} pageNo={i + 1} size={size} settings={settings} />
        </div>
      ))}
    </div>,
    document.body,
  );
}
