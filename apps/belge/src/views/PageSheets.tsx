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
import { HF_KEYS, hfKeyOf, type HfKey } from "../editor/sectionHf";

export type { FootnoteView } from "./FootnoteText";

const hfExtensions = headerFooterExtensions();

const htmlCache = new WeakMap<JSONContent, string>();

function toHtml(json: JSONContent | null): string {
  if (!json) return "";
  const hit = htmlCache.get(json);
  if (hit !== undefined) return hit;
  let out = "";
  try {
    out = hasContent(json) ? generateHTML(json, hfExtensions) : "";
  } catch {
    out = "";
  }
  htmlCache.set(json, out);
  return out;
}

/** Bölüm × tür üst/alt bilgi tablosu (HF_KEYS sırasıyla) ve "Öncekiyle Aynı" bilgisi. */
export interface HfTable {
  content: (JSONContent | null)[][];
  linked: boolean[][];
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

export function hfKey(which: "header" | "footer", v: HfVariant): HfKey {
  return hfKeyOf(which, v);
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
  /** Bölümlerin etkin üst/alt bilgileri. */
  hf: HfTable;
}

export const PageSheets = memo(function PageSheets({ pages, rects, settings, gridlines, editing, editingPage, footnotes, notes, onNoteChange, sectionCount, hf }: Props) {
  const html = useMemo(() => hf.content.map((row) => row.map((json) => toHtml(json))), [hf]);
  const htmlOf = (section: number, key: HfKey) => html[Math.min(section, html.length - 1)]?.[HF_KEYS.indexOf(key)] ?? "";
  const linkedOf = (section: number, key: HfKey) => hf.linked[Math.min(section, hf.linked.length - 1)]?.[HF_KEYS.indexOf(key)] ?? false;
  const wm = settings.watermark;
  const pb = settings.pageBorder;
  const count = pages.length;
  const out = [];
  for (let i = 0; i < count; i++) {
    const g = pages[i];
    const v = hfVariantOf(g, i, settings);
    const headerHtml = htmlOf(g.section, hfKey("header", v));
    const footerHtml = htmlOf(g.section, hfKey("footer", v));
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
              {linkedOf(g.section, hfKey("header", v)) && <span className="bl-hf-same">Öncekiyle Aynı</span>}
            </div>
            <div className="bl-hf-line is-footer" style={{ top: g.h - g.mb }}>
              <span>{hfLabel("footer", v, settings.oddEven, g.section, sectionCount)}</span>
              {linkedOf(g.section, hfKey("footer", v)) && <span className="bl-hf-same">Öncekiyle Aynı</span>}
            </div>
          </>
        )}
      </div>,
    );
  }
  return <>{out}</>;
});
