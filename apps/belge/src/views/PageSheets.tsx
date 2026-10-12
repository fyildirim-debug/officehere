// Sayfa zeminleri: beyaz sayfa, filigran, sayfa kenarlığı, üst/alt bilgi, kılavuz çizgileri.
import { memo, useLayoutEffect, useMemo, useRef } from "react";
import { generateHTML, type JSONContent } from "@tiptap/core";
import { fontStack } from "@officehere/core";
import type { DocSettings } from "../model/document";
import { NUMPAGES_TOKEN, PAGE_TOKEN } from "../editor/nodes";
import { headerFooterExtensions } from "../editor/extensions";

export interface PageGeom {
  w: number;
  h: number;
  mt: number;
  mb: number;
  ml: number;
  mr: number;
  headerY: number;
  footerY: number;
  stride: number;
}

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

export interface FootnoteView {
  label: string;
  text: string;
}

interface Props {
  count: number;
  /** Sayfa başına dipnot kimlikleri ve kimlik → metin eşlemesi. */
  footnotes: string[][];
  notes: Map<string, FootnoteView>;
  onNoteChange: (id: string, text: string) => void;
  g: PageGeom;
  settings: DocSettings;
  gridlines: boolean;
  editing: "header" | "footer" | null;
  /** Düzenlenen üst/alt bilginin bulunduğu sayfa. */
  editingPage: number;
}

export const PageSheets = memo(function PageSheets({ count, g, settings, gridlines, editing, editingPage, footnotes, notes, onNoteChange }: Props) {
  const headerHtml = useMemo(() => (hasContent(settings.header) ? toHtml(settings.header) : ""), [settings.header]);
  const footerHtml = useMemo(() => (hasContent(settings.footer) ? toHtml(settings.footer) : ""), [settings.footer]);
  const wm = settings.watermark;
  const pb = settings.pageBorder;
  const pages = [];
  for (let i = 0; i < count; i++) {
    const fill = (html: string) => html.split(PAGE_TOKEN).join(String(i + 1)).split(NUMPAGES_TOKEN).join(String(count));
    pages.push(
      <div key={i} className="bl-page" data-page={i} style={{ top: i * g.stride, width: g.w, height: g.h, background: settings.pageColor ?? undefined }}>
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
        {!(editing === "header" && i === editingPage) && headerHtml && (
          <div className="bl-hf bl-doc bl-hf-header" style={{ top: g.headerY, left: g.ml, right: g.mr, maxHeight: g.mt - g.headerY }} dangerouslySetInnerHTML={{ __html: fill(headerHtml) }} />
        )}
        {!(editing === "footer" && i === editingPage) && footerHtml && (
          <div className="bl-hf bl-doc bl-hf-footer" style={{ bottom: g.footerY, left: g.ml, right: g.mr, maxHeight: g.mb - g.footerY }} dangerouslySetInnerHTML={{ __html: fill(footerHtml) }} />
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
                  <FootnoteText id={id} text={n.text} onChange={onNoteChange} />
                </div>
              );
            })}
          </div>
        )}
        {editing && (
          <>
            <div className="bl-hf-line" style={{ top: g.mt }}>
              <span>Üst Bilgi</span>
            </div>
            <div className="bl-hf-line is-footer" style={{ top: g.h - g.mb }}>
              <span>Alt Bilgi</span>
            </div>
          </>
        )}
      </div>,
    );
  }
  return <>{pages}</>;
});

/** Sayfa altındaki düzenlenebilir dipnot metni. */
function FootnoteText({ id, text, onChange }: { id: string; text: string; onChange: (id: string, text: string) => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    if (document.activeElement !== el && el.value !== text) el.value = text;
    el.style.height = "auto";
    el.style.height = `${el.scrollHeight}px`;
  }, [text]);
  return (
    <textarea
      ref={ref}
      className="bl-fn-text"
      rows={1}
      defaultValue={text}
      placeholder="Dipnot metni"
      data-note-id={id}
      onChange={(e) => {
        e.currentTarget.style.height = "auto";
        e.currentTarget.style.height = `${e.currentTarget.scrollHeight}px`;
        onChange(id, e.currentTarget.value);
      }}
    />
  );
}
