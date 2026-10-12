// İzlenen değişikliklerin görünümü: sol kenarda değişiklik çubukları ve sağ kenarda düzeltme balonları
// ("Silindi: …", "Biçimlendirilmiş: Kalın").
import { useLayoutEffect, useState } from "react";
import type { Editor } from "@tiptap/core";
import type { Mark, Node as PMNode } from "@tiptap/pm/model";
import type { PageInfo } from "../editor/pagination";
import { listChanges } from "../editor/trackChanges";
import type { PageRect } from "./DocumentView";

export interface RevisionBalloon {
  id: string;
  from: number;
  to: number;
  kind: "deletion" | "formatChange";
  author: string;
  date: string;
  /** "Silindi: metin" / "Biçimlendirilmiş: Kalın". */
  label: string;
  text: string;
}

const FMT_LABEL: Record<string, string> = {
  bold: "Kalın",
  italic: "İtalik",
  underline: "Altı çizili",
  strike: "Üstü çizili",
  superscript: "Üst simge",
  subscript: "Alt simge",
  highlight: "Vurgulu",
};

/** Biçim değişikliğinin açıklaması (eski ve yeni işaretlerin farkı). */
export function describeFormat(node: PMNode, oldJson: string): string {
  let old: { type: string; attrs?: Record<string, unknown> }[] = [];
  try {
    old = JSON.parse(oldJson) as typeof old;
  } catch {
    old = [];
  }
  const now = node.marks.filter((m: Mark) => !["insertion", "deletion", "formatChange", "comment", "link"].includes(m.type.name));
  const parts: string[] = [];
  for (const m of now) {
    const prev = old.find((o) => o.type === m.type.name);
    if (FMT_LABEL[m.type.name] && !prev) parts.push(FMT_LABEL[m.type.name]);
    if (m.type.name === "textStyle") {
      const a = m.attrs as Record<string, unknown>;
      const p = (prev?.attrs ?? {}) as Record<string, unknown>;
      if (a.fontFamily && a.fontFamily !== p.fontFamily) parts.push(`Yazı tipi: ${String(a.fontFamily)}`);
      if (a.fontSize && a.fontSize !== p.fontSize) parts.push(`${String(a.fontSize).replace("pt", "")} nk`);
      if (a.color && a.color !== p.color) parts.push("Yazı tipi rengi");
    }
  }
  for (const o of old) if (FMT_LABEL[o.type] && !now.some((m) => m.type.name === o.type)) parts.push(`${FMT_LABEL[o.type]} Değil`);
  return parts.join(", ") || "Biçim";
}

/** Balonlarda gösterilecek düzeltmeler (kip: tüm düzeltmeler ya da yalnız biçimlendirme). */
export function revisionBalloons(editor: Editor, mode: "revisions" | "formatting" | "inline"): RevisionBalloon[] {
  if (mode === "inline") return [];
  const out: RevisionBalloon[] = [];
  for (const c of listChanges(editor.state.doc)) {
    if (c.kind === "insertion" || (c.kind === "deletion" && mode !== "revisions")) continue;
    let label = "Silindi: ";
    if (c.kind === "formatChange") {
      const node = editor.state.doc.nodeAt(c.from);
      const fm = node?.marks.find((m) => m.type.name === "formatChange");
      label = `Biçimlendirilmiş: ${node && fm ? describeFormat(node, String(fm.attrs.old ?? "[]")) : "Biçim"}`;
    }
    out.push({ id: `rev-${c.kind}-${c.from}`, from: c.from, to: c.to, kind: c.kind, author: c.author, date: c.date, label, text: c.kind === "deletion" ? c.text : "" });
  }
  return out;
}

/**
 * Görünür bir konumun ekran koordinatları: gizli (balondaki silinmiş) metnin içindeki konumlar için
 * yakındaki görünür konum kullanılır.
 */
export function visibleCoords(view: Editor["view"], pos: number): { top: number; bottom: number; left: number } | null {
  const size = view.state.doc.content.size;
  for (const [p, side] of [
    [pos, 1],
    [pos, -1],
  ] as const) {
    try {
      const c = view.coordsAtPos(Math.max(0, Math.min(size, p)), side);
      if (c.bottom - c.top > 0.5 && (c.top !== 0 || c.left !== 0)) return c;
    } catch {
      // geçersiz konum
    }
  }
  // Gizli aralığın dışına doğru ilerle.
  for (let d = 1; d < 400; d++) {
    for (const p of [pos + d, pos - d]) {
      if (p < 0 || p > size) continue;
      try {
        const c = view.coordsAtPos(p, 1);
        if (c.bottom - c.top > 0.5 && (c.top !== 0 || c.left !== 0)) return c;
      } catch {
        // geçersiz konum
      }
    }
  }
  return null;
}

interface BarProps {
  editor: Editor;
  container: HTMLElement | null;
  pages: PageInfo[];
  rectOf: (i: number) => PageRect;
  pageAt: (pos: number) => number;
  scale: number;
  version: number;
}

/** Değişiklik içeren satırların yanında, sol kenar boşluğunda dikey çubuk (Word'deki gibi). */
export function ChangeBars({ editor, container, pages, rectOf, pageAt, scale, version }: BarProps) {
  const [bars, setBars] = useState<{ x: number; y: number; h: number }[]>([]);
  useLayoutEffect(() => {
    if (!container) return;
    const base = container.getBoundingClientRect();
    const raw: { page: number; top: number; bottom: number }[] = [];
    for (const c of listChanges(editor.state.doc)) {
      try {
        const a = visibleCoords(editor.view, c.from);
        const b = visibleCoords(editor.view, Math.max(c.from, c.to)) ?? a;
        if (!a || !b) continue;
        const page = pageAt(c.from);
        const top = (Math.min(a.top, b.top) - base.top) / scale;
        const bottom = (Math.max(a.bottom, b.bottom) - base.top) / scale;
        raw.push({ page, top, bottom });
      } catch {
        // konum çözümlenemedi
      }
    }
    raw.sort((x, y) => x.top - y.top);
    const merged: typeof raw = [];
    for (const r of raw) {
      const last = merged[merged.length - 1];
      if (last && last.page === r.page && r.top <= last.bottom + 2) last.bottom = Math.max(last.bottom, r.bottom);
      else merged.push({ ...r });
    }
    const out = merged.map((r) => {
      const pg = pages[r.page] ?? pages[0];
      const rect = rectOf(r.page);
      return { x: rect.x + (pg ? pg.ml : 96) - 18, y: r.top, h: Math.max(4, r.bottom - r.top) };
    });
    setBars((prev) => (JSON.stringify(prev) === JSON.stringify(out) ? prev : out));
  }, [editor, container, pages, rectOf, pageAt, scale, version]);
  return (
    <>
      {bars.map((b, i) => (
        <div key={i} className="bl-changebar" style={{ left: b.x, top: b.y, height: b.h }} aria-hidden="true" />
      ))}
    </>
  );
}
