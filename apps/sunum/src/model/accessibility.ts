/**
 * Erişilebilirlik Denetleyicisi: alternatif metin eksikliği, slayt başlığı,
 * okuma sırası, metin kontrastı (WCAG 2 oranı) ve tablo üst bilgisi.
 */
import { hexToRgb, resolveHex } from "./color";
import { layoutOf, levelResolver, effectiveRun } from "./textStyle";
import { isLineGeom } from "../render/geometry";
import type { Fill, Presentation, Shape, Slide } from "./types";

export type IssueLevel = "error" | "warning" | "tip";
export type IssueKind = "altText" | "noTitle" | "dupTitle" | "readingOrder" | "contrast" | "tableHeader";

export interface A11yIssue {
  level: IssueLevel;
  kind: IssueKind;
  slideIndex: number;
  shapeId?: string;
  /** Listede görünen nesne adı. */
  label: string;
}

export const ISSUE_TITLES: Record<IssueKind, string> = {
  altText: "Alternatif Metin Eksik",
  noTitle: "Slayt Başlığı Eksik",
  dupTitle: "Yinelenen Slayt Başlığı",
  readingOrder: "Okuma Sırasını Denetleyin",
  contrast: "Okunması Zor Metin Kontrastı",
  tableHeader: "Tablo Üst Bilgisi Eksik",
};

/** Sorunun nedeni ve nasıl düzeltileceği (kendi ifademizle). */
export const ISSUE_HELP: Record<IssueKind, string> = {
  altText: "Ekran okuyucular resim, grafik ve şekillerin ne anlattığını alternatif metinden okur. Nesneyi kısa ve açıklayıcı bir cümleyle tanımlayın.",
  noTitle: "Başlıklar slaytlar arasında gezinmeyi ve slaydın konusunu anlamayı kolaylaştırır. Başlık yer tutucusuna bir başlık yazın.",
  dupTitle: "Aynı başlığı taşıyan slaytlar, slayt listesinde birbirinden ayırt edilemez. Her slayda benzersiz bir başlık verin.",
  readingOrder: "Ekran okuyucular nesneleri Seçim Bölmesi'ndeki sırayla okur. Başlığın önce, diğer içeriğin yukarıdan aşağıya okunduğundan emin olun.",
  contrast: "Metin rengi ile arka plan arasındaki fark az olduğunda metin zor okunur. Daha koyu ya da daha açık bir renk seçin.",
  tableHeader: "Üst bilgi satırı, tablo sütunlarının ne içerdiğini açıklar. Tablo Tasarımı sekmesinde Üst Bilgi Satırı'nı açın.",
};

const shapeText = (s: Shape) => (s.kind === "sp" && s.text ? s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ").trim() : "");
const isTitle = (s: Shape) => s.kind === "sp" && (s.ph?.type === "title" || s.ph?.type === "ctrTitle");

/** WCAG göreli parlaklık ve kontrast oranı. */
export function contrastRatio(a: string, b: string): number {
  const lum = (hex: string) => {
    const [r, g, bb] = hexToRgb(hex).map((v) => {
      const c = v / 255;
      return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
    });
    return 0.2126 * r + 0.7152 * g + 0.0722 * bb;
  };
  const la = lum(a);
  const lb = lum(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

function solidHex(f: Fill | undefined, pres: Presentation, theme = pres.theme): string | null {
  if (!f) return null;
  if (f.type === "solid") return resolveHex(f.color, theme);
  if (f.type === "gradient" && f.stops.length) return resolveHex(f.stops[Math.floor(f.stops.length / 2)].color, theme);
  return null;
}

function needsAlt(s: Shape): boolean {
  if (s.descr?.trim()) return false;
  if (s.kind === "pic" || s.kind === "chart" || s.kind === "group") return true;
  if (s.kind === "sp") return !s.ph && !s.ink && !isLineGeom(s.geom) && !shapeText(s) && !s.txBox;
  return false;
}

export function checkAccessibility(pres: Presentation): A11yIssue[] {
  const out: A11yIssue[] = [];
  const titles = new Map<string, number[]>();
  pres.slides.forEach((slide, si) => {
    if (slide.hidden) return;
    const theme = layoutOf(pres, slide)?.theme ?? pres.theme;
    const visible = slide.shapes.filter((s) => !s.hidden);
    for (const s of visible) if (needsAlt(s)) out.push({ level: "error", kind: "altText", slideIndex: si, shapeId: s.id, label: s.name });
    const title = visible.find(isTitle);
    const tt = title ? shapeText(title) : "";
    if (!tt) out.push({ level: "error", kind: "noTitle", slideIndex: si, label: `Slayt ${si + 1}` });
    else titles.set(tt.toLocaleLowerCase("tr-TR"), [...(titles.get(tt.toLocaleLowerCase("tr-TR")) ?? []), si]);
    for (const s of visible)
      if (s.kind === "table" && !s.style.firstRow) out.push({ level: "error", kind: "tableHeader", slideIndex: si, shapeId: s.id, label: s.name });
    readingOrder(slide, visible, si, out);
    contrast(pres, slide, visible, si, theme, out);
  });
  for (const list of titles.values()) if (list.length > 1) for (const si of list) out.push({ level: "warning", kind: "dupTitle", slideIndex: si, label: `Slayt ${si + 1}` });
  const rank: Record<IssueLevel, number> = { error: 0, warning: 1, tip: 2 };
  return out.sort((a, b) => rank[a.level] - rank[b.level] || a.slideIndex - b.slideIndex);
}

/** Başlık okuma sırasında ilk değilse ya da içerik görsel sıradan (yukarıdan aşağı) çok farklı okunuyorsa. */
function readingOrder(slide: Slide, shapes: Shape[], si: number, out: A11yIssue[]) {
  const content = shapes.filter((s) => (s.kind === "sp" && shapeText(s)) || s.kind === "pic" || s.kind === "chart" || s.kind === "table");
  if (content.length < 2) return;
  const ti = content.findIndex(isTitle);
  const visual = content.slice().sort((a, b) => (Math.abs(a.y - b.y) > 20 ? a.y - b.y : a.x - b.x));
  const inversions = content.reduce((n, s, i) => n + (visual.indexOf(s) < i - 1 ? 1 : 0), 0);
  if (ti > 0 || inversions > 0) out.push({ level: "warning", kind: "readingOrder", slideIndex: si, label: `Slayt ${si + 1}` });
  void slide;
}

function contrast(pres: Presentation, slide: Slide, shapes: Shape[], si: number, theme: Presentation["theme"], out: A11yIssue[]) {
  const layout = layoutOf(pres, slide);
  const slideBg = solidHex(slide.background ?? layout?.background ?? pres.master.background, pres, theme) ?? "#FFFFFF";
  for (const s of shapes) {
    if (s.kind !== "sp" || !s.text || !shapeText(s)) continue;
    const bg = solidHex(s.fill, pres, theme) ?? slideBg;
    const lvl = levelResolver(s, { master: pres.master, layout });
    let worst = Infinity;
    for (const p of s.text.paragraphs)
      for (const r of p.runs) {
        if (!r.text.trim()) continue;
        const eff = effectiveRun(r.props, p, lvl);
        const ratio = contrastRatio(resolveHex(eff.color, theme), bg);
        const large = (eff.size ?? 18) >= 18 || ((eff.size ?? 18) >= 14 && !!eff.bold);
        worst = Math.min(worst, ratio / (large ? 3 : 4.5));
      }
    if (worst < 1) out.push({ level: "warning", kind: "contrast", slideIndex: si, shapeId: s.id, label: s.name });
  }
}
