// Dizin: XE alanları (girdi işaretleri) ve OfficeHereIndex içerik denetimi içinde INDEX alanı.
import type { JSONContent } from "@tiptap/core";
import type { IndexItem } from "../../editor/indexing";
import { indexLine, xeCode } from "../../editor/indexing";
import { attr, cleanText, descendants, esc, kid } from "./xml";

const fldBegin = (instr: string) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>`;
const fldEnd = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** XE: sonucu olmayan alan (Word'de yalnız alan kodları gösterilince görünür). */
export function xeXml(node: JSONContent): string {
  return fldBegin(xeCode(String(node.attrs?.main ?? ""), (node.attrs?.sub as string | null) ?? undefined)) + fldEnd;
}

/** INDEX alanı: her girdi bir paragraf (Index 1 / Index 2 stilleri). */
export function indexXml(node: JSONContent): string {
  const entries = (node.attrs?.entries as IndexItem[]) ?? [];
  const cols = Math.max(1, Number(node.attrs?.columns) || 2);
  const begin = `${fldBegin(`INDEX \\c "${cols}" \\z "1055"`)}<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const run = (t: string) => `<w:r><w:t xml:space="preserve">${esc(cleanText(t))}</w:t></w:r>`;
  let x = `<w:sdt><w:sdtPr><w:tag w:val="OfficeHereIndex"/></w:sdtPr><w:sdtContent>`;
  if (!entries.length) x += `<w:p><w:pPr><w:pStyle w:val="Index1"/></w:pPr>${begin}${run("Dizin girdisi bulunamadı.")}${fldEnd}</w:p>`;
  entries.forEach((e, i) => {
    const style = e.sub ? "Index2" : "Index1";
    const ind = e.sub ? '<w:ind w:left="480" w:hanging="240"/>' : '<w:ind w:left="240" w:hanging="240"/>';
    x += `<w:p><w:pPr><w:pStyle w:val="${style}"/>${ind}</w:pPr>${i === 0 ? begin : ""}${run(indexLine(e))}${i === entries.length - 1 ? fldEnd : ""}</w:p>`;
  });
  return `${x}</w:sdtContent></w:sdt>`;
}

/** OfficeHereIndex içerik denetimi → dizin bloğu (girdiler sonuç paragraflarından). */
export function indexNode(content: Element, instr: string): JSONContent {
  const columns = Number(/\\c\s+"?(\d+)"?/.exec(instr)?.[1] ?? 2) || 2;
  const entries: IndexItem[] = [];
  let lastMain = "";
  for (const p of descendants(content, "p")) {
    const text = descendants(p, "t")
      .map((t) => t.textContent ?? "")
      .join("")
      .trim();
    if (!text || text === "Dizin girdisi bulunamadı.") continue;
    const style = attr(kid(kid(p, "pPr"), "pStyle"), "val") ?? "";
    const m = /^(.*?)(?:,\s*([\d,\s]+))?$/.exec(text);
    const name = (m?.[1] ?? text).trim();
    const pages = (m?.[2] ?? "")
      .split(/[,\s]+/)
      .map((x) => parseInt(x, 10))
      .filter((n) => Number.isFinite(n));
    if (/Index2|dizin\s*2/i.test(style) && lastMain) entries.push({ main: lastMain, sub: name, pages });
    else {
      lastMain = name;
      entries.push({ main: name, pages });
    }
  }
  return { type: "indexBlock", attrs: { columns, entries } };
}

/** XE alan kodu → girdi düğümü. */
export function xeNode(instr: string): JSONContent | null {
  const m = /^\s*XE\s+"([^"]*)"/i.exec(instr);
  if (!m) return null;
  const [main, ...rest] = m[1].split(":");
  return main.trim() ? { type: "indexEntry", attrs: { main: main.trim(), sub: rest.join(":").trim() || null } } : null;
}
