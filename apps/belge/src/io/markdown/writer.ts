// Markdown yazıcı: belge → CommonMark + GFM (tablolar, üstü çizili, görev listeleri, dipnotlar).
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument } from "../../model/document";
import { fieldNowText } from "../../editor/nodes";
import { atomText, isMonoFont, marksToProps, plainBlocks, sectionsOf, textOf, type Mark } from "../common";
import { CODE_SHADE } from "./reader";

type MdMark = "link" | "strong" | "em" | "del" | "u" | "sup" | "sub" | "mark" | "code";

/** Açılış sırası: bağlantı en dışta, kod en içte. */
const ORDER: MdMark[] = ["link", "strong", "em", "del", "u", "mark", "sup", "sub", "code"];

function mdMarks(marks: Mark[] | undefined): Map<MdMark, string> {
  const p = marksToProps(marks);
  const out = new Map<MdMark, string>();
  if (p.link) out.set("link", p.link);
  if (p.bold || p.charStyle === "Strong") out.set("strong", "");
  if (p.italic || p.charStyle === "Emphasis") out.set("em", "");
  if (p.strike) out.set("del", "");
  if (p.underline) out.set("u", "");
  if (p.highlight) out.set("mark", "");
  if (p.sup) out.set("sup", "");
  if (p.sub) out.set("sub", "");
  if (isMonoFont(p.font)) out.set("code", "");
  return out;
}

const OPEN: Record<Exclude<MdMark, "link" | "code">, [string, string]> = {
  strong: ["**", "**"],
  em: ["*", "*"],
  del: ["~~", "~~"],
  u: ["<u>", "</u>"],
  mark: ["<mark>", "</mark>"],
  sup: ["<sup>", "</sup>"],
  sub: ["<sub>", "</sub>"],
};

/** Düz metin kaçışı (satır başı yapıları dahil). */
function escapeText(s: string, lineStart: boolean): string {
  let t = s.replace(/[\\`*_[\]<>~|]/g, "\\$&").replace(/&(?=[a-zA-Z0-9#]+;)/g, "\\&");
  if (lineStart) {
    t = t.replace(/^(\s*)([#>+=-])/, "$1\\$2").replace(/^(\s*\d+)([.)])/, "$1\\$2");
  }
  return t;
}

function linkDest(href: string): string {
  return /[\s()<>]/.test(href) ? `<${href.replace(/[<>]/g, (c) => encodeURIComponent(c))}>` : href;
}

function codeSpan(s: string): string {
  const longest = Math.max(0, ...(s.match(/`+/g) ?? []).map((x) => x.length));
  const fence = "`".repeat(longest + 1);
  const pad = s.startsWith("`") || s.endsWith("`") || (s.startsWith(" ") && s.endsWith(" ") && s.trim()) ? " " : "";
  return `${fence}${pad}${s}${pad}${fence}`;
}

class MdWriter {
  notes: string[] = [];
  noteNum = 0;

  /** Satır içi içerik; "indent" sonraki satırların girintisidir (liste ve alıntı içinde). */
  inline(nodes: JSONContent[], indent = "", table = false): string {
    let out = "";
    const open: { m: MdMark; v: string }[] = [];
    let pendingWs = "";
    let codeBuf = "";
    const lineStart = () => out === "" || out.endsWith("\n" + indent);
    const flushCode = () => {
      if (codeBuf) out += codeSpan(codeBuf);
      codeBuf = "";
    };
    const closeTo = (k: number) => {
      while (open.length > k) {
        const top = open.pop()!;
        if (top.m === "code") flushCode();
        else if (top.m === "link") out += `](${linkDest(top.v)})`;
        else out += OPEN[top.m][1];
      }
    };
    const setMarks = (want: Map<MdMark, string>, leadingWs: string) => {
      // Ortak önek korunur, kalanı kapatılıp yeniden açılır.
      let keep = 0;
      while (keep < open.length && want.has(open[keep].m) && want.get(open[keep].m) === open[keep].v) keep++;
      // Kod parçası en içte kalmalı: yeni işaret açılacaksa önce kapanır.
      const missing = ORDER.some((m) => want.has(m) && !open.slice(0, keep).some((o) => o.m === m));
      const ci = open.findIndex((o) => o.m === "code");
      if (missing && ci >= 0 && ci < keep) keep = ci;
      if (keep < open.length) closeTo(keep);
      out += pendingWs + leadingWs;
      pendingWs = "";
      for (const m of ORDER) {
        if (!want.has(m) || open.some((o) => o.m === m)) continue;
        if (m === "link") out += "[";
        else if (m !== "code") out += OPEN[m][0];
        open.push({ m, v: want.get(m)! });
      }
    };
    for (const n of nodes) {
      // İzlenen değişiklikte silinmiş metin yazılmaz.
      if (n.marks?.some((m) => m.type === "deletion")) continue;
      if (n.type === "text") {
        let s = n.text ?? "";
        if (table) s = s.replace(/\n/g, " ");
        const want = mdMarks(n.marks);
        if (want.has("code")) {
          setMarks(want, "");
          codeBuf += s;
          continue;
        }
        const lead = /^\s*/.exec(s)![0];
        const trail = s.length > lead.length ? /\s*$/.exec(s)![0] : "";
        const core = s.slice(lead.length, s.length - trail.length);
        if (!core) {
          pendingWs += s;
          continue;
        }
        setMarks(want, lead);
        out += escapeText(core, lineStart()).replace(/\t/g, "    ");
        pendingWs = trail;
        continue;
      }
      if (n.type === "hardBreak") {
        closeTo(0);
        out += pendingWs.replace(/\s+$/, "");
        pendingWs = "";
        out += table ? "<br>" : `\\\n${indent}`;
        continue;
      }
      if (n.type === "image") {
        const a = n.attrs ?? {};
        const src = String(a.src ?? "");
        if (!src) continue;
        setMarks(mdMarks(n.marks?.filter((m) => m.type === "link")), "");
        out += `![${escapeText(String(a.alt ?? ""), false)}](${linkDest(src)})`;
        continue;
      }
      if (n.type === "note") {
        setMarks(new Map(), "");
        const num = ++this.noteNum;
        out += `[^${num}]`;
        const a = n.attrs ?? {};
        const blocks = (a.content as JSONContent[] | null) ?? String(a.text ?? "").split("\n").map((t) => ({ type: "paragraph", content: t ? [{ type: "text", text: t }] : [] }));
        const body = blocks.map((b) => this.block(b, "    ")).filter(Boolean).join("\n\n    ");
        this.notes.push(`[^${num}]: ${body}`);
        continue;
      }
      // Markdown'da alan yoktur: tarih/saat o anki değeriyle yazılır, sayfa numarası atlanır.
      if (n.type === "pageField" && n.attrs?.field !== "DATE" && n.attrs?.field !== "TIME") continue;
      const t = n.type === "pageField" ? fieldNowText(n.attrs?.field, new Date(), n.attrs?.format) : atomText(n);
      if (t) {
        setMarks(mdMarks(n.marks), "");
        out += escapeText(t, lineStart());
      }
    }
    closeTo(0);
    out += pendingWs.replace(/\s+$/, "");
    return out;
  }

  isCode(n: JSONContent): boolean {
    if (n.type !== "paragraph" || n.attrs?.shading !== CODE_SHADE) return false;
    return (n.content ?? []).every((c) => c.type === "text" && isMonoFont(marksToProps(c.marks).font));
  }

  /** Blok dizisi (kod paragrafları birleştirilir). */
  blocks(nodes: JSONContent[] | undefined, indent = ""): string {
    const parts: string[] = [];
    const list = nodes ?? [];
    for (let i = 0; i < list.length; i++) {
      const n = list[i];
      if (this.isCode(n)) {
        const lines: string[] = [];
        while (i < list.length && this.isCode(list[i])) lines.push((list[i++].content ?? []).map((c) => c.text ?? "").join(""));
        i--;
        const body = lines.join("\n");
        const longest = Math.max(2, ...(body.match(/`+/g) ?? []).map((x) => x.length));
        const fence = "`".repeat(longest + 1);
        parts.push([fence, ...lines, fence].join(`\n${indent}`));
        continue;
      }
      const b = this.block(n, indent);
      if (b) parts.push(b);
    }
    return parts.join(`\n\n${indent}`);
  }

  block(n: JSONContent, indent = ""): string {
    switch (n.type) {
      case "heading": {
        const t = this.inline(n.content ?? [], indent).replace(/\\\n\s*/g, " ");
        return t ? `${"#".repeat(Math.min(6, Number(n.attrs?.level) || 1))} ${t}` : "";
      }
      case "paragraph": {
        const t = this.inline(n.content ?? [], indent);
        if (!t.trim()) return "";
        const sid = n.attrs?.styleId;
        if (sid === "Quote" || sid === "IntenseQuote") return `> ${t.split(`\n${indent}`).join(`\n${indent}> `)}`;
        return t;
      }
      case "bulletList":
      case "orderedList":
        return this.list(n, indent);
      case "table":
        return this.table(n, indent);
      case "horizontalRule":
        return "---";
      case "textBox":
        return this.blocks(n.content, indent);
      case "bibliography":
      case "indexBlock":
        return this.blocks(plainBlocks(n), indent);
      default:
        return n.content ? this.blocks(n.content, indent) : "";
    }
  }

  list(n: JSONContent, indent: string): string {
    const ordered = n.type === "orderedList";
    let num = Number(n.attrs?.start) || 1;
    const delim = n.attrs?.listStyle === "decimal-paren" ? ")" : ".";
    const lines: string[] = [];
    for (const li of n.content ?? []) {
      const marker = ordered ? `${num++}${delim} ` : "- ";
      const inner = indent + " ".repeat(marker.length);
      const kids = li.content ?? [];
      let first = "";
      let rest = kids;
      if (kids[0]?.type === "paragraph") {
        const p = kids[0];
        const txt = textOf(p);
        const task = /^[☐☑☒]\s/.exec(txt);
        let content = p.content ?? [];
        if (task && content[0]?.type === "text") content = [{ ...content[0], text: (content[0].text ?? "").replace(/^[☐☑☒]\s/, "") }, ...content.slice(1)];
        first = (task ? (task[0][0] === "☐" ? "[ ] " : "[x] ") : "") + this.inline(content, inner);
        rest = kids.slice(1);
      }
      let item = `${marker}${first}`;
      for (const k of rest) {
        const b = this.block(k, inner);
        if (!b) continue;
        item += k.type === "bulletList" || k.type === "orderedList" ? `\n${inner}${b}` : `\n\n${inner}${b}`;
      }
      lines.push(item);
    }
    return lines.join(`\n${indent}`);
  }

  table(n: JSONContent, indent: string): string {
    // Hücre ızgarası: birleştirilmiş hücrelerin kapladığı yerler boş kalır.
    const grid: string[][] = [];
    const aligns: (string | null)[] = [];
    (n.content ?? []).forEach((row, r) => {
      grid[r] ??= [];
      let c = 0;
      for (const cell of row.content ?? []) {
        while (grid[r][c] !== undefined) c++;
        const text = (cell.content ?? [])
          .map((b) => (b.type === "paragraph" || b.type === "heading" ? this.inline(b.content ?? [], "", true) : textOf(b).replace(/\|/g, "\\|")))
          .filter(Boolean)
          .join("<br>")
          .replace(/\n/g, " ");
        const cs = Math.max(1, Number(cell.attrs?.colspan) || 1);
        const rs = Math.max(1, Number(cell.attrs?.rowspan) || 1);
        if (r === 0) aligns[c] = (cell.content?.[0]?.attrs?.textAlign as string | null) ?? null;
        for (let dr = 0; dr < rs; dr++) for (let dc = 0; dc < cs; dc++) (grid[r + dr] ??= [])[c + dc] = dr === 0 && dc === 0 ? text : "";
        c += cs;
      }
    });
    const cols = Math.max(1, ...grid.map((r) => r.length));
    const row = (r: string[]) => `| ${Array.from({ length: cols }, (_, k) => r[k] ?? "").join(" | ")} |`;
    const sep = `|${Array.from({ length: cols }, (_, k) => (aligns[k] === "center" ? ":---:" : aligns[k] === "right" ? "---:" : aligns[k] === "left" ? ":---" : "---")).join("|")}|`;
    return [row(grid[0] ?? []), sep, ...grid.slice(1).map(row)].join(`\n${indent}`);
  }
}

/** Belge → Markdown metni. */
export function writeMarkdown(doc: BelgeDocument): string {
  const w = new MdWriter();
  const parts: string[] = [];
  const title = doc.settings.properties.title.trim();
  if (title) parts.push(`---\ntitle: "${title.replace(/"/g, "'")}"\n---`);
  for (const sec of sectionsOf(doc.content)) {
    const b = w.blocks(sec.content);
    if (b) parts.push(b);
  }
  if (w.notes.length) parts.push(w.notes.join("\n\n"));
  return parts.join("\n\n") + "\n";
}
