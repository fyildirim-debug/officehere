// Düz metin (.txt) ve HTML içe/dışa aktarma.
import type { JSONContent } from "@tiptap/core";
import { fontStack } from "@officehere/core";
import { pageMetrics, type DocSettings } from "../model/document";
import { resolveStyles, styleSheetCss } from "../model/styles";
import documentCss from "../styles/document.css?raw";

/** Baytları metne çevirir: UTF-8 (BOM'lu/BOM'suz) ya da Türkçe Windows-1254. */
export function decodeText(data: ArrayBuffer | Uint8Array): string {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes).replace(/^﻿/, "");
  } catch {
    try {
      return new TextDecoder("windows-1254").decode(bytes);
    } catch {
      return new TextDecoder("latin1").decode(bytes);
    }
  }
}

export function textToDoc(text: string): JSONContent {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  return {
    type: "doc",
    content: lines.map((l) => {
      const parts = l.split("\f");
      const out: JSONContent = { type: "paragraph", ...(parts[0] ? { content: [{ type: "text", text: parts[0] }] } : {}) };
      return out;
    }),
  };
}

interface NoteAcc {
  f: number;
  e: number;
  list: string[];
}

const ROMAN = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"];

function inlineText(n: JSONContent, notes?: NoteAcc): string {
  if (n.type === "text") return n.text ?? "";
  if (n.type === "hardBreak") return "\n";
  if (n.type === "pageField") return "1";
  if (n.type === "note" && notes) {
    const end = n.attrs?.kind === "endnote";
    const label = end ? (ROMAN[++notes.e] ?? String(notes.e)) : String(++notes.f);
    notes.list.push(`[${label}] ${String(n.attrs?.text ?? "")}`);
    return `[${label}]`;
  }
  return (n.content ?? []).map((c) => inlineText(c, notes)).join("");
}

/** Belge → düz metin (Windows satır sonları; madde işaretleri ve numaralar korunur). */
export function docToText(doc: JSONContent, eol = "\r\n"): string {
  const lines: string[] = [];
  const notes: NoteAcc = { f: 0, e: 0, list: [] };
  const walk = (nodes: JSONContent[] | undefined, indent: string, listInfo?: { ordered: boolean; n: number }) => {
    for (const n of nodes ?? []) {
      switch (n.type) {
        case "paragraph":
        case "heading":
          lines.push(indent + inlineText(n, notes).replace(/\n/g, eol));
          break;
        case "bulletList":
        case "orderedList": {
          let i = (n.attrs?.start as number) ?? 1;
          for (const li of n.content ?? []) {
            const marker = n.type === "orderedList" ? `${i++}. ` : "• ";
            (li.content ?? []).forEach((c, k) => {
              if (k === 0 && (c.type === "paragraph" || c.type === "heading")) lines.push(indent + marker + inlineText(c, notes));
              else walk([c], indent + "    ", listInfo);
            });
          }
          break;
        }
        case "table":
          for (const row of n.content ?? []) lines.push(indent + (row.content ?? []).map((cell) => (cell.content ?? []).map((c) => inlineText(c, notes)).join(" ")).join("\t"));
          break;
        case "pageBreak":
          lines.push("\f");
          break;
        case "horizontalRule":
          lines.push("");
          break;
        case "toc":
          for (const e of (n.attrs?.entries as { text: string; page: number; level: number }[]) ?? []) lines.push(`${"  ".repeat(e.level - 1)}${e.text}\t${e.page}`);
          break;
        default:
          walk(n.content, indent, listInfo);
      }
    }
  };
  walk(doc.content, "");
  if (notes.list.length) lines.push("", "____________", ...notes.list);
  return lines.join(eol) + eol;
}

/** Bağımsız HTML belgesi (stiller ve sayfa ayarlarıyla). */
export function htmlDocument(bodyHtml: string, settings: DocSettings, title: string): string {
  const styles = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing);
  const m = pageMetrics(settings.page);
  const css = `
@page { size: ${(settings.page.width / 1440) * 2.54}cm ${(settings.page.height / 1440) * 2.54}cm; margin: ${(settings.page.margin.top / 1440) * 2.54}cm ${(settings.page.margin.right / 1440) * 2.54}cm ${(settings.page.margin.bottom / 1440) * 2.54}cm ${(settings.page.margin.left / 1440) * 2.54}cm; }
body { margin: 0; background: ${settings.pageColor ?? "#fff"}; color: #000; }
.bl-doc { max-width: ${Math.round(m.contentW)}px; margin: 24px auto; padding: 0 16px; font-family: ${fontStack(styles.Normal.props.font ?? "Aptos")}; }
.bl-doc [data-page-break] { break-after: page; height: 0; overflow: hidden; }
.bl-doc p, .bl-doc h1, .bl-doc h2, .bl-doc h3, .bl-doc h4, .bl-doc h5, .bl-doc h6 { white-space: pre-wrap; }
${documentCss}
${styleSheetCss(styles)}
`;
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  return `<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8">
<meta name="generator" content="OfficeHere Belge">
<title>${esc(title)}</title>
<style>${css}</style>
</head>
<body>
<div class="bl-doc">
${bodyHtml}
</div>
</body>
</html>
`;
}

/** HTML belgesinden gövde içeriğini çıkarır (stil ve betik etiketleri atılır). */
export function extractHtmlBody(html: string): string {
  const doc = new DOMParser().parseFromString(html, "text/html");
  doc.querySelectorAll("script, style, meta, link, title").forEach((e) => e.remove());
  const root = doc.querySelector(".bl-doc") ?? doc.body;
  return root.innerHTML;
}
