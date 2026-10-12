// Adres mektup birleştirme: alıcı listesi (CSV / XLSX) okuma, birleştirme (her kayıt için belge kopyası)
// ve birleştirilen belgenin yeni pencerede açılması.
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { getNative } from "@officehere/core";
import type { BelgeDocument } from "../model/document";

export interface RecipientList {
  name: string;
  fields: string[];
  records: Record<string, string>[];
}

/** CSV (virgül, noktalı virgül ya da sekme ayraçlı; tırnaklı alanlar) → satırlar. */
export function parseCsv(text: string): string[][] {
  const src = text.replace(/^﻿/, "");
  const firstLine = src.split(/\r?\n/, 1)[0] ?? "";
  const counts = [",", ";", "\t"].map((d) => [d, firstLine.split(d).length] as const);
  const delim = counts.sort((a, b) => b[1] - a[1])[0][0];
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (quoted) {
      if (c === '"' && src[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (c === '"') quoted = false;
      else cell += c;
    } else if (c === '"' && cell === "") quoted = true;
    else if (c === delim) {
      row.push(cell);
      cell = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && src[i + 1] === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += c;
  }
  if (cell !== "" || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((r) => r.some((x) => x.trim() !== ""));
}

const colIndex = (ref: string) => {
  const letters = /^[A-Z]+/.exec(ref)?.[0] ?? "A";
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

/** XLSX'in ilk çalışma sayfası → satırlar (paylaşılan dizeler, satır içi dizeler, sayılar). */
export async function parseXlsx(data: ArrayBuffer | Uint8Array): Promise<string[][]> {
  const zip = await JSZip.loadAsync(data);
  const parser = new DOMParser();
  const xml = async (p: string) => {
    const f = zip.file(p);
    return f ? parser.parseFromString(await f.async("string"), "application/xml") : null;
  };
  const all = (d: Document | Element, name: string) => Array.from(d.getElementsByTagName("*")).filter((e) => e.localName === name);
  const shared = (await xml("xl/sharedStrings.xml")) as Document | null;
  const strings = shared
    ? all(shared, "si").map((si) =>
        all(si, "t")
          .map((t) => t.textContent ?? "")
          .join(""),
      )
    : [];
  // İlk sayfanın yolu: workbook ilişkilerinden (yoksa sheet1).
  let sheetPath = "xl/worksheets/sheet1.xml";
  const wb = await xml("xl/workbook.xml");
  const rels = await xml("xl/_rels/workbook.xml.rels");
  const first = wb ? all(wb, "sheet")[0] : undefined;
  const rid = first?.getAttribute("r:id") ?? first?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  const target =
    rels && rid
      ? all(rels, "Relationship")
          .find((r) => r.getAttribute("Id") === rid)
          ?.getAttribute("Target")
      : null;
  if (target) sheetPath = target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\//, "")}`;
  const sheet = await xml(sheetPath);
  if (!sheet) return [];
  const rows: string[][] = [];
  for (const r of all(sheet, "row")) {
    const out: string[] = [];
    for (const c of all(r, "c")) {
      const i = colIndex(c.getAttribute("r") ?? "A");
      const t = c.getAttribute("t");
      const v = all(c, "v")[0]?.textContent ?? "";
      let val = v;
      if (t === "s") val = strings[Number(v)] ?? "";
      else if (t === "inlineStr")
        val = all(c, "t")
          .map((x) => x.textContent ?? "")
          .join("");
      else if (t === "b") val = v === "1" ? "DOĞRU" : "YANLIŞ";
      else if (v && /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(v)) val = Number(v).toLocaleString("tr-TR", { maximumFractionDigits: 10, useGrouping: false });
      while (out.length < i) out.push("");
      out[i] = val;
    }
    rows.push(out);
  }
  return rows.filter((r) => r.some((x) => x.trim() !== ""));
}

/** Satırlar → alıcı listesi (ilk satır alan adları). */
export function rowsToList(name: string, rows: string[][]): RecipientList {
  const head = (rows[0] ?? []).map((h, i) => h.trim() || `Alan${i + 1}`);
  const seen = new Set<string>();
  const fields = head.map((h) => {
    let n = h.replace(/\s+/g, "_");
    for (let k = 2; seen.has(n); k++) n = `${h}_${k}`;
    seen.add(n);
    return n;
  });
  const records = rows.slice(1).map((r) => Object.fromEntries(fields.map((f, i) => [f, (r[i] ?? "").trim()])));
  return { name, fields, records };
}

export async function readRecipients(name: string, data: ArrayBuffer): Promise<RecipientList> {
  if (/\.xlsx$/i.test(name)) return rowsToList(name, await parseXlsx(data));
  return rowsToList(name, parseCsv(new TextDecoder("utf-8").decode(data)));
}

/** İçerikteki birleştirme alanlarını kaydın değerleriyle değiştirir (alanın biçimi korunur). */
export function fillRecord(node: JSONContent, rec: Record<string, string>): JSONContent {
  if (!node.content) return node;
  return {
    ...node,
    content: node.content.flatMap((c): JSONContent[] => {
      if (c.type !== "mergeField") return [fillRecord(c, rec)];
      const text = rec[(c.attrs?.name as string) ?? ""] ?? "";
      return text ? [{ type: "text", text, ...(c.marks ? { marks: c.marks } : {}) }] : [];
    }),
  };
}

/** Birleştirme: her kayıt için belgenin bölümleri; kayıtlar arası bölüm sonu (sonraki sayfa). */
export function mergeDocument(content: JSONContent, records: Record<string, string>[]): JSONContent {
  const sections: JSONContent[] = [];
  for (const rec of records) {
    const filled = fillRecord(content, rec);
    (filled.content ?? []).forEach((sec, i) => {
      sections.push(i === 0 && sections.length ? { ...sec, attrs: { ...(sec.attrs ?? {}), kind: "nextPage" } } : sec);
    });
  }
  return { type: "doc", content: sections.length ? sections : [{ type: "section", content: [{ type: "paragraph" }] }] };
}

const HANDOFF = "officehere.belge.handoff";

export interface Handoff {
  doc: BelgeDocument;
  name: string;
  then?: "print";
  /** Kurtarılan taslak: yeni pencere kendi taslağına yazınca bu anahtar silinir. */
  restoreKey?: string;
}

/** Belgeyi yeni pencerede açar (birleştirme sonucu "Mektuplar1" ya da kurtarılan taslak). */
export function openInNewWindow(doc: BelgeDocument, name: string, then?: "print", restoreKey?: string) {
  try {
    localStorage.setItem(HANDOFF, JSON.stringify({ doc, name, then, restoreKey, at: Date.now() }));
  } catch {
    return false;
  }
  const native = getNative();
  if (native) void native.newWindow("belge");
  else window.open(`${location.pathname}?handoff=1`, "_blank");
  return true;
}

let taken: Handoff | null | undefined;

/** Açılışta yeni pencereye aktarılan belge (varsa; bir kez alınır, sonraki çağrılar aynı sonucu döndürür). */
export function takeHandoff(): Handoff | null {
  if (taken !== undefined) return taken;
  taken = null;
  try {
    if (!getNative() && !new URLSearchParams(location.search).has("handoff")) return null;
    const raw = localStorage.getItem(HANDOFF);
    if (!raw) return null;
    const v = JSON.parse(raw) as Handoff & { at: number };
    if (Date.now() - v.at > 60_000) return null;
    localStorage.removeItem(HANDOFF);
    taken = v;
    return v;
  } catch {
    return null;
  }
}

/** Bu pencere başka bir pencereden aktarılan belgeyle mi açıldı. */
export function isHandoffWindow(): boolean {
  return !!takeHandoff();
}
