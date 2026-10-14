/**
 * Anahattan Slaytlar: düz metin (.txt/.md) ya da belge (.docx) ana hattından
 * slayt üretir. Girintisiz satır (ya da Başlık 1) yeni slaydın başlığıdır;
 * sekmeyle/boşlukla girintili satırlar (ya da Başlık 2–6) madde düzeyleridir.
 */
import { bodyShape, linesToBody, titleShape } from "./outline";
import { layoutByType, newSlide } from "./presentation";
import type { Presentation, Slide } from "./types";
import { loadIo } from "../io/formats";

export interface OutlineEntry {
  title: string;
  /** "\t" ile düzeyi belirtilmiş madde satırları. */
  lines: string[];
}

/** Düz metin ana hattı. Baştaki 4 boşluk ya da bir sekme bir düzeydir; "#"/"-" işaretleri de anlaşılır. */
export function parseTextOutline(text: string): OutlineEntry[] {
  const out: OutlineEntry[] = [];
  for (const raw of text.replace(/\r\n?/g, "\n").split("\n")) {
    if (!raw.trim()) continue;
    const lead = /^[\t ]*/.exec(raw)![0];
    let level = 0;
    for (const ch of lead) level += ch === "\t" ? 1 : 0.25;
    level = Math.floor(level);
    let body = raw.slice(lead.length);
    const md = /^(#{1,6})\s+/.exec(body);
    if (md) {
      level = md[1].length - 1;
      body = body.slice(md[0].length);
    } else if (/^[-*•]\s+/.test(body) && level === 0) {
      level = 1;
      body = body.replace(/^[-*•]\s+/, "");
    } else body = body.replace(/^[-*•]\s+/, "");
    if (level === 0 || !out.length) out.push({ title: body, lines: [] });
    else out[out.length - 1].lines.push("\t".repeat(Math.min(4, level - 1)) + body);
  }
  return out;
}

/** .docx: Başlık 1 → slayt başlığı, Başlık 2..6 → madde düzeyi; diğer paragraflar yok sayılır. */
export async function parseDocxOutline(data: ArrayBuffer | Uint8Array): Promise<OutlineEntry[]> {
  const { JSZip } = await loadIo();
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) return [];
  const stylesXml = (await zip.file("word/styles.xml")?.async("string")) ?? "";
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const styles = new DOMParser().parseFromString(stylesXml || "<w:styles xmlns:w='x'/>", "application/xml");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  // Stil kimliği → ana hat düzeyi (w:outlineLvl ya da "Heading N" adı).
  const lvlOf = new Map<string, number>();
  for (const st of Array.from(styles.getElementsByTagNameNS(W, "style"))) {
    const id = st.getAttributeNS(W, "styleId") ?? "";
    const ol = st.getElementsByTagNameNS(W, "outlineLvl")[0]?.getAttributeNS(W, "val");
    const name = st.getElementsByTagNameNS(W, "name")[0]?.getAttributeNS(W, "val") ?? "";
    const m = /(heading|başlık|balk)\s*(\d)/i.exec(name) ?? /(heading|başlık|balk)(\d)/i.exec(id);
    if (ol !== null && ol !== undefined && ol !== "") lvlOf.set(id, Number(ol));
    else if (m) lvlOf.set(id, Number(m[2]) - 1);
  }
  const lines: string[] = [];
  for (const p of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    const sid = p.getElementsByTagNameNS(W, "pStyle")[0]?.getAttributeNS(W, "val") ?? "";
    let lvl = lvlOf.get(sid);
    if (lvl === undefined) {
      const m = /(heading|başlık|balk)(\d)/i.exec(sid);
      if (m) lvl = Number(m[2]) - 1;
    }
    if (lvl === undefined || lvl > 5) continue;
    const text = Array.from(p.getElementsByTagNameNS(W, "t"))
      .map((t) => t.textContent ?? "")
      .join("")
      .trim();
    if (text) lines.push("\t".repeat(lvl) + text);
  }
  return parseTextOutline(lines.join("\n"));
}

/** Ana hattan "Başlık ve İçerik" slaytları. */
export function slidesFromOutline(pres: Presentation, entries: OutlineEntry[]): Slide[] {
  const obj = layoutByType(pres, "obj");
  return entries.map((e) => {
    const s = newSlide(pres, obj.id);
    return {
      ...s,
      shapes: s.shapes.map((sh) => {
        if (sh.kind !== "sp") return sh;
        if (sh === titleShape(s)) return { ...sh, text: linesToBody(e.title, sh.text) };
        if (sh === bodyShape(s) && e.lines.length) return { ...sh, text: linesToBody(e.lines.join("\n"), sh.text) };
        return sh;
      }),
    };
  });
}
