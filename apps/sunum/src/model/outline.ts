/** Ana hat görünümü için başlık/gövde metni ile satır listesi arasında dönüşüm. */
import type { Paragraph, Slide, SpShape, TextBody } from "./types";

const TITLE_TYPES = new Set(["title", "ctrTitle"]);
const BODY_TYPES = new Set(["body", "obj", "subTitle"]);

export function titleShape(s: Slide): SpShape | undefined {
  return s.shapes.find((x): x is SpShape => x.kind === "sp" && !!x.ph && TITLE_TYPES.has(x.ph.type));
}

export function bodyShape(s: Slide): SpShape | undefined {
  return s.shapes.find((x): x is SpShape => x.kind === "sp" && !!x.ph && BODY_TYPES.has(x.ph.type));
}

export const paraText = (p: Paragraph) => p.runs.map((r) => r.text).join("");

export function bodyToLines(t: TextBody | undefined): string {
  if (!t) return "";
  const lines = t.paragraphs.map((p) => "\t".repeat(p.level) + paraText(p));
  return lines.length === 1 && !lines[0] ? "" : lines.join("\n");
}

/** Satırları paragraflara çevirir; aynı sıradaki paragrafın biçimi korunur. */
export function linesToBody(text: string, old: TextBody | undefined): TextBody {
  const base: TextBody = old ?? { paragraphs: [] };
  const lines = text.split("\n");
  const paragraphs: Paragraph[] = lines.map((line, k) => {
    const m = /^\t*/.exec(line)![0].length;
    const body = line.slice(m);
    const prev = base.paragraphs[k];
    const runProps = prev?.runs.find((r) => !r.props.field)?.props ?? prev?.endProps ?? {};
    return { level: Math.min(8, m), props: prev?.props ?? {}, runs: body ? [{ text: body, props: { ...runProps } }] : [], endProps: prev?.endProps };
  });
  return { ...base, paragraphs };
}

