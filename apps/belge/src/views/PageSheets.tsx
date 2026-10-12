// Sayfa zeminleri: beyaz sayfa, filigran, sayfa kenarlığı, üst/alt bilgi, kılavuz çizgileri, dipnotlar.
// Her sayfa kendi bölümünün boyut ve kenar boşluklarıyla çizilir.
import { memo, useMemo } from "react";
import { generateHTML, type JSONContent } from "@tiptap/core";
import { fontStack } from "@officehere/core";
import type { DocSettings, HfVariant } from "../model/document";
import { NUMPAGES_TOKEN, PAGE_TOKEN } from "../editor/nodes";
import { headerFooterExtensions } from "../editor/extensions";
import type { PageInfo } from "../editor/pagination";
import { FootnoteText, type FootnoteView } from "./FootnoteText";

export type { FootnoteView } from "./FootnoteText";

const hfExtensions = headerFooterExtensions();

function toHtml(json: JSONContent | null): string {
  if (!json) return "";
  try {
    return generateHTML(json, hfExtensions);
  } catch {
    return "";
  }
}

export function hasContent(json: JSONContent | null): boolean {
  if (!json?.content) return false;
  const walk = (n: JSONContent): boolean => (n.type === "text" && !!n.text?.trim()) || n.type === "pageField" || n.type === "image" || !!n.content?.some(walk);
  return json.content.some(walk);
}

/** Sayfanın üst/alt bilgi türü: bölümde "Farklı İlk Sayfa" ve belgede "Tek ve Çift Sayfalarda Farklı". */
export function hfVariantOf(page: PageInfo | undefined, index: number, settings: Pick<DocSettings, "oddEven">): HfVariant {
  if (page?.titlePage && page.first) return "first";
  if (settings.oddEven && (index + 1) % 2 === 0) return "even";
  return "default";
}

export function hfKey(which: "header" | "footer", v: HfVariant): "header" | "footer" | "headerFirst" | "footerFirst" | "headerEven" | "footerEven" {
  if (v === "first") return which === "header" ? "headerFirst" : "footerFirst";
  if (v === "even") return which === "header" ? "headerEven" : "footerEven";
  return which;
}

/** Düzenleme sırasında sayfadaki üst/alt bilgi alanı etiketi (Word'deki gibi). */
export function hfLabel(which: "header" | "footer", v: HfVariant, oddEven: boolean, section: number, sections: number): string {
  const base = which === "header" ? "Üst Bilgi" : "Alt Bilgi";
  const pre = v === "first" ? "İlk Sayfa " : v === "even" ? "Çift Sayfa " : oddEven ? "Tek Sayfa " : "";
  const name = pre ? `${pre}${base.replace("Bilgi", "Bilgisi")}` : base;
  return sections > 1 ? `${name} -Bölüm ${section + 1}-` : name;
}

interface Props {
  pages: PageInfo[];
  /** Yan yana dizimde sayfa konumları (null: düzen bilgisindeki konumlar). */
  rects: { x: number; y: number }[] | null;
  /** Sayfa başına dipnot kimlikleri ve kimlik → içerik eşlemesi. */
  footnotes: string[][];
  notes: Map<string, FootnoteView>;
  onNoteChange: (id: string, content: JSONContent) => void;
  settings: DocSettings;
  gridlines: boolean;
  editing: "header" | "footer" | null;
  /** Düzenlenen üst/alt bilginin bulunduğu sayfa. */
  editingPage: number;
  sectionCount: number;
}

export const PageSheets = memo(function PageSheets({ pages, rects, settings, gridlines, editing, editingPage, footnotes, notes, onNoteChange, sectionCount }: Props) {
  const html = useMemo(() => {
    const m = new Map<string, string>();
    for (const k of ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"] as const) {
      const json = settings[k];
      m.set(k, hasContent(json) ? toHtml(json) : "");
    }
    return m;
  }, [settings.header, settings.footer, settings.headerFirst, settings.footerFirst, settings.headerEven, settings.footerEven]); // eslint-disable-line react-hooks/exhaustive-deps
  const wm = settings.watermark;
  const pb = settings.pageBorder;
  const count = pages.length;
  const out = [];
  for (let i = 0; i < count; i++) {
    const g = pages[i];
    const v = hfVariantOf(g, i, settings);
    const headerHtml = html.get(hfKey("header", v)) ?? "";
    const footerHtml = html.get(hfKey("footer", v)) ?? "";
    const fill = (h: string) => h.split(PAGE_TOKEN).join(String(i + 1)).split(NUMPAGES_TOKEN).join(String(count));
    // Aynı türdeki üst/alt bilgi düzenlenirken diğer sayfalarda canlı kopya gösterilmez (düzenleyici tek sayfada).
    const hideHeader = editing === "header" && i === editingPage;
    const hideFooter = editing === "footer" && i === editingPage;
    out.push(
      <div key={i} className={`bl-page${g.blank ? " is-blank" : ""}`} data-page={i} data-section={g.section} style={{ top: rects?.[i]?.y ?? g.top, left: rects?.[i]?.x ?? g.left, width: g.w, height: g.h, background: settings.pageColor ?? undefined }}>
        {wm && (
          <div className="bl-watermark" aria-hidden="true">
            <span
              style={{
                color: wm.color,
                opacity: wm.semitransparent ? 0.35 : 0.8,
                fontFamily: fontStack(wm.font),
                fontSize: Math.min(g.w, g.h) / Math.max(4, wm.text.length * 0.62),
                transform: wm.layout === "diagonal" ? `rotate(${-Math.atan(g.h / g.w) * (180 / Math.PI)}deg)` : undefined,
              }}
            >
              {wm.text}
            </span>
          </div>
        )}
        {pb && <div className="bl-page-border" style={{ inset: 32, border: `${pb.style === "thick" ? Math.max(3, pb.width) : pb.width}pt ${pb.style === "thick" || pb.style === "single" ? "solid" : pb.style} ${pb.color}` }} />}
        {gridlines && <div className="bl-gridlines" style={{ top: g.mt, left: g.ml, right: g.mr, bottom: g.mb }} />}
        {!hideHeader && headerHtml && (
          <div className="bl-hf bl-doc bl-hf-header" style={{ top: g.headerY, left: g.ml, right: g.mr, maxHeight: Math.max(0, g.mt - g.headerY) }} dangerouslySetInnerHTML={{ __html: fill(headerHtml) }} />
        )}
        {!hideFooter && footerHtml && (
          <div className="bl-hf bl-doc bl-hf-footer" style={{ bottom: g.footerY, left: g.ml, right: g.mr, maxHeight: Math.max(0, g.mb - g.footerY) }} dangerouslySetInnerHTML={{ __html: fill(footerHtml) }} />
        )}
        {(footnotes[i]?.length ?? 0) > 0 && (
          <div className="bl-footnotes bl-doc" style={{ left: g.ml, right: g.mr, bottom: g.mb }}>
            <div className="bl-note-sep" />
            {footnotes[i].map((id) => {
              const n = notes.get(id);
              if (!n) return null;
              return (
                <div key={id} className="bl-fn-row">
                  <sup>{n.label}</sup>
                  <FootnoteText id={id} note={n} onChange={onNoteChange} />
                </div>
              );
            })}
          </div>
        )}
        {editing && (
          <>
            <div className="bl-hf-line" style={{ top: g.mt }}>
              <span>{hfLabel("header", v, settings.oddEven, g.section, sectionCount)}</span>
            </div>
            <div className="bl-hf-line is-footer" style={{ top: g.h - g.mb }}>
              <span>{hfLabel("footer", v, settings.oddEven, g.section, sectionCount)}</span>
            </div>
          </>
        )}
      </div>,
    );
  }
  return <>{out}</>;
});
