/**
 * Otomatik filtre ölçütlerinin XLSX'e yazılması ve okunması (ECMA-376
 * §18.3.2 autoFilter/filterColumn): değer listesi (<filters>), özel ölçütler
 * (<customFilters>), ilk 10 (<top10>). Sayfa otomatik filtresi sayfa XML'inde,
 * tablo filtreleri tablo parçasında durur. exceljs ölçütleri yazmaz/okumaz.
 */
import JSZip from "jszip";
import { formatRange } from "@officehere/formula";
import type { AutoFilter, FilterCriteria } from "../model/types";
import type { WorkbookModel } from "../model/workbook";
import { esc, relsPathOf, resolveTarget, sheetPaths, unescapeXml } from "./ooxmlZip";

const OPS: Record<string, string> = { eq: "equal", ne: "notEqual", gt: "greaterThan", ge: "greaterThanOrEqual", lt: "lessThan", le: "lessThanOrEqual" };

function customFilter(op: string, v: string): string {
  const n = Number(v.replace(/\./g, "").replace(",", "."));
  const val = v.trim() !== "" && Number.isFinite(n) ? String(n) : v;
  switch (op) {
    case "begins":
      return `<customFilter val="${esc(val)}*"/>`;
    case "ends":
      return `<customFilter val="*${esc(val)}"/>`;
    case "contains":
      return `<customFilter val="*${esc(val)}*"/>`;
    case "notContains":
      return `<customFilter operator="notEqual" val="*${esc(val)}*"/>`;
    default:
      return `<customFilter${OPS[op] && op !== "eq" ? ` operator="${OPS[op]}"` : ""} val="${esc(val)}"/>`;
  }
}

export function filterColumnXml(colId: number, c: FilterCriteria): string | null {
  switch (c.kind) {
    case "values":
      return `<filterColumn colId="${colId}"><filters${c.blanks ? ' blank="1"' : ""}>${c.values.map((v) => `<filter val="${esc(v)}"/>`).join("")}</filters></filterColumn>`;
    case "custom": {
      const parts = [customFilter(c.op1, c.v1)];
      if (c.op2) parts.push(customFilter(c.op2, c.v2 ?? ""));
      return `<filterColumn colId="${colId}"><customFilters${c.op2 && c.join !== "or" ? ' and="1"' : ""}>${parts.join("")}</customFilters></filterColumn>`;
    }
    case "top10":
      return `<filterColumn colId="${colId}"><top10${c.bottom ? ' top="0"' : ""}${c.percent ? ' percent="1"' : ""} val="${c.count}"/></filterColumn>`;
  }
  return null;
}

export function autoFilterXml(f: AutoFilter): string {
  const cols = Object.entries(f.columns)
    .map(([k, c]) => filterColumnXml(Number(k), c))
    .filter(Boolean)
    .join("");
  return cols ? `<autoFilter ref="${formatRange(f.range)}">${cols}</autoFilter>` : `<autoFilter ref="${formatRange(f.range)}"/>`;
}

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unescapeXml(m[1]) : undefined;
}

const OP_BACK: Record<string, string> = { equal: "eq", notEqual: "ne", greaterThan: "gt", greaterThanOrEqual: "ge", lessThan: "lt", lessThanOrEqual: "le" };

function parseCustom(tag: string): { op: string; v: string } {
  const op = OP_BACK[attr(tag, "operator") ?? "equal"] ?? "eq";
  const val = attr(tag, "val") ?? "";
  const m = /^(\*?)(.*?)(\*?)$/.exec(val)!;
  const tr = (s: string) => (/^-?\d+(\.\d+)?$/.test(s) ? s.replace(".", ",") : s);
  if (m[1] && m[3] && m[2]) return { op: op === "ne" ? "notContains" : "contains", v: m[2] };
  if (m[3] && m[2] && op === "eq") return { op: "begins", v: m[2] };
  if (m[1] && m[2] && op === "eq") return { op: "ends", v: m[2] };
  return { op, v: tr(val) };
}

/** <autoFilter> içindeki filtre sütunları → model ölçütleri. */
export function parseFilterColumns(autoFilterXml: string): Record<number, FilterCriteria> {
  const out: Record<number, FilterCriteria> = {};
  for (const m of autoFilterXml.matchAll(/<(?:\w+:)?filterColumn\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?filterColumn>/g)) {
    const col = Number(attr(m[1], "colId"));
    if (!Number.isFinite(col)) continue;
    const body = m[2];
    const filters = /<(?:\w+:)?filters\b([^>]*)\/?>([\s\S]*?)(?:<\/(?:\w+:)?filters>|$)/.exec(body);
    if (filters && /<(?:\w+:)?filters\b/.test(body)) {
      const values = [...body.matchAll(/<(?:\w+:)?filter\b([^>]*)\/?>/g)].map((f) => attr(f[1], "val") ?? "").filter((v) => v !== "");
      out[col] = { kind: "values", values, blanks: attr(filters[1], "blank") === "1" || attr(filters[1], "blank") === "true" };
      continue;
    }
    const cf = /<(?:\w+:)?customFilters\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?customFilters>/.exec(body);
    if (cf) {
      const tags = [...cf[2].matchAll(/<(?:\w+:)?customFilter\b([^>]*)\/?>/g)].map((t) => parseCustom(t[1]));
      if (tags.length) {
        const and = attr(cf[1], "and") === "1" || attr(cf[1], "and") === "true";
        out[col] = { kind: "custom", op1: tags[0].op, v1: tags[0].v, ...(tags[1] ? { join: and ? "and" : "or", op2: tags[1].op, v2: tags[1].v } : {}) } as FilterCriteria;
      }
      continue;
    }
    const top = /<(?:\w+:)?top10\b([^>]*)\/?>/.exec(body);
    if (top) out[col] = { kind: "top10", count: Number(attr(top[1], "val") ?? 10), bottom: attr(top[1], "top") === "0" || attr(top[1], "top") === "false", percent: attr(top[1], "percent") === "1" || attr(top[1], "percent") === "true" };
  }
  return out;
}

/** Tablo parçasındaki autoFilter'ı ölçütlerle değiştirir. */
export function patchTableAutoFilter(tableXml: string, f: AutoFilter | undefined): string {
  if (!f || !Object.keys(f.columns).length) return tableXml;
  // Tablo autoFilter'ında colId tablo sütun sırasıdır (filtre aralığına göre)
  return tableXml.replace(/<autoFilter\b[^>]*\/>|<autoFilter\b[\s\S]*?<\/autoFilter>/, autoFilterXml(f));
}

/** Dosyadan sayfa ve tablo filtre ölçütlerini okur. */
export async function readFiltersFromData(data: ArrayBuffer, wb: WorkbookModel): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const p = paths.get(sheet.name);
    const xml = p ? await zip.file(p)?.async("string") : undefined;
    if (!p || !xml) continue;
    const af = /<(?:\w+:)?autoFilter\b[^>]*>[\s\S]*?<\/(?:\w+:)?autoFilter>/.exec(xml)?.[0];
    if (af && sheet.autoFilter) sheet.autoFilter.columns = parseFilterColumns(af);
    if (!sheet.tables.length) continue;
    const rels = await zip.file(relsPathOf(p))?.async("string");
    for (const m of (rels ?? "").matchAll(/<Relationship\b[^>]*>/g)) {
      if (!/Type="[^"]*\/table"/.test(m[0])) continue;
      const target = /Target="([^"]+)"/.exec(m[0])?.[1];
      const tx = target ? await zip.file(resolveTarget(p, target))?.async("string") : undefined;
      if (!tx) continue;
      const name = attr(/<(?:\w+:)?table\b[^>]*>/.exec(tx)?.[0] ?? "", "displayName") ?? "";
      const t = sheet.tables.find((x) => x.name.replace(/[^\p{L}\p{N}_]/gu, "_") === name || x.name === name);
      const taf = /<(?:\w+:)?autoFilter\b[^>]*>[\s\S]*?<\/(?:\w+:)?autoFilter>/.exec(tx)?.[0];
      if (t?.filter && taf) t.filter.columns = parseFilterColumns(taf);
    }
  }
}
