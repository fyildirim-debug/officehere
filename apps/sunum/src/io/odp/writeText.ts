/** ODP metin yazıcısı: paragraflar, listeler (madde/numara düzeyleri), aralıklar ve alanlar. */
import { effectivePara, effectiveRun } from "../../model/textStyle";
import type { BulletDef, Color, LevelStyle, Paragraph, ParaProps, RunProps, TextBody } from "../../model/types";
import { esc } from "../xml";
import { cm, odfAttrs } from "./common";
import { paraProps, textProps, type AutoStyles, type Shared } from "./writeStyles";

export interface TextWriteCtx {
  sh: Shared;
  auto: AutoStyles;
  lvl: (n: number) => LevelStyle;
  /** "#slide:<kimlik>" → "#<sayfa adı>". */
  link: (l: string) => string | undefined;
  /** Tablo stilinden gelen hücre metni biçimi (çalıştırmada yoksa). */
  cell?: { color?: Color; bold?: boolean };
  /**
   * Okuyucunun yeniden açılışta bu metne stillerden (ana sayfa sunu stili, şekil
   * grafik stili ya da varsayılanlar) devralacağı çalıştırma biçimi, düzey başına.
   * Verilirse çalıştırmalara yalnız açıkça ayarlanmış ya da devralınandan farklı
   * özellikler yazılır; devralınan değerler çalıştırmalara sabitlenmez (tema/asıl
   * slayt değişince metin de değişsin).
   */
  inherited?: (level: number) => RunProps;
}

/** ODF metin özelliklerine yazılabilen çalıştırma biçimi anahtarları. */
const RUN_KEYS = ["size", "bold", "italic", "underline", "strike", "color", "font", "caps", "baseline", "highlight", "shadow", "spacing"] as const;

/** Bir özelliğin ODF'deki yazılı biçimi (renk/yazı tipi çözülmüş olarak karşılaştırmak için). */
const written = (k: (typeof RUN_KEYS)[number], v: unknown, sh: Shared) => textProps({ [k]: v } as RunProps, sh);

/**
 * Yazılacak çalıştırma biçimi: açıkça ayarlanmış özellikler + devralınan değerden
 * farklı olan etkin özellikler. `inh` yoksa (eski davranış) etkin biçimin tamamı.
 */
/** Açık/kapalı biçimlerin "kapalı" değeri: devralınan açık değeri iptal etmek için yazılır. */
const NEUTRAL: Partial<RunProps> = { bold: false, italic: false, underline: false, strike: false, shadow: false, caps: "none" };

export function ownRun(explicit: RunProps, eff: RunProps, inh: RunProps | undefined, sh: Shared): RunProps {
  if (!inh) return eff;
  const out: RunProps = {};
  for (const k of RUN_KEYS) {
    // Etkin değeri olmayan (kapalı) biçim, devralınan açık değerin altında kalmasın.
    const v = eff[k] ?? (inh[k] !== undefined ? NEUTRAL[k] : undefined);
    if (v === undefined) continue;
    if (explicit[k] !== undefined || written(k, v, sh) !== written(k, inh[k], sh)) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

/** Paragraf stilindeki metin özellikleri (ODF'de tüm çalıştırmalara geçer): yalnız her çalıştırmada aynı olanlar. */
function commonRun(end: RunProps, runs: RunProps[], sh: Shared): RunProps {
  if (!runs.length) return end;
  const out: RunProps = {};
  for (const k of RUN_KEYS) {
    const v = end[k];
    if (v === undefined) continue;
    const w = written(k, v, sh);
    if (runs.every((r) => r[k] !== undefined && written(k, r[k], sh) === w)) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

/** Metni ODF öğelerine çevirir (satır sonu, sekme, ardışık boşluk). */
function textXml(t: string, paraStart: boolean): string {
  let out = "";
  let i = 0;
  while (i < t.length) {
    const ch = t[i];
    if (ch === "\v" || ch === "\n" || ch === "\r") {
      out += "<text:line-break/>";
      i++;
    } else if (ch === "\t") {
      out += "<text:tab/>";
      i++;
    } else if (ch === " ") {
      let n = 0;
      while (t[i + n] === " ") n++;
      const lead = paraStart && i === 0;
      if (lead) out += n > 1 ? `<text:s text:c="${n}"/>` : "<text:s/>";
      else out += n > 1 ? ` <text:s text:c="${n - 1}"/>` : " ";
      i += n;
    } else {
      let j = i;
      while (j < t.length && !"\v\n\r\t ".includes(t[j])) j++;
      out += esc(t.slice(i, j));
      i = j;
    }
  }
  return out;
}

function runXml(text: string, props: RunProps, own: RunProps, ctx: TextWriteCtx, paraStart: boolean): string {
  const st = ctx.auto.style("text", "T", textProps(own, ctx.sh));
  let inner: string;
  if (props.field === "slidenum") inner = `<text:page-number text:select-page="current">${esc(text === "‹#›" ? "1" : text)}</text:page-number>`;
  else if (props.field === "datetime") inner = `<text:date>${esc(text)}</text:date>`;
  else inner = textXml(text, paraStart);
  const span = `<text:span text:style-name="${st}">${inner}</text:span>`;
  const href = props.link ? (props.link.startsWith("#") ? ctx.link(props.link) : props.link) : undefined;
  return href ? `<text:a${odfAttrs({ "xlink:type": "simple", "xlink:href": href })}>${span}</text:a>` : span;
}

function paraXml(p: Paragraph, eff: ParaProps, inList: boolean, ctx: TextWriteCtx): string {
  const withCell = (r: RunProps): RunProps => (ctx.cell ? { ...r, color: r.color ?? ctx.cell.color, bold: r.bold ?? ctx.cell.bold } : r);
  const inh = ctx.inherited?.(p.level);
  // Tablo stilinden gelen hücre biçimi açık değil etkin biçim sayılır (devralınandan farklıysa yazılır).
  const own = (r: RunProps) => ownRun(r, effectiveRun(withCell(r), p, ctx.lvl), inh, ctx.sh);
  const runs = p.runs.map((r) => own(r.props));
  const end = own(p.endProps ?? p.runs[p.runs.length - 1]?.props ?? {});
  const common = inh ? commonRun(end, runs, ctx.sh) : end;
  const ps = ctx.auto.style("paragraph", "P", paraProps(eff, inList) + textProps({ ...common, link: undefined, field: undefined }, ctx.sh));
  let pos = 0;
  const body = p.runs
    .map((r, i) => {
      const x = runXml(r.text, r.props, runs[i], ctx, pos === 0);
      pos += r.text.length;
      return x;
    })
    .join("");
  return `<text:p text:style-name="${ps}">${body}</text:p>`;
}

const NUM: Record<string, [string, string]> = {
  arabicPeriod: ["1", "."],
  arabicParenR: ["1", ")"],
  romanUcPeriod: ["I", "."],
  romanLcPeriod: ["i", "."],
  alphaUcPeriod: ["A", "."],
  alphaLcPeriod: ["a", "."],
  alphaLcParenR: ["a", ")"],
  alphaUcParenR: ["A", ")"],
};

interface LevelDef {
  bullet: BulletDef;
  marL: number;
  indent: number;
}

/** Liste stili: düzey başına madde/numara ve girinti. */
export function listStyleXml(levels: Map<number, LevelDef>, sh: Shared): string {
  let out = "";
  for (let k = 0; k < 9; k++) {
    const d = levels.get(k) ?? { bullet: { type: "char", char: "•" } as BulletDef, marL: 24 + k * 48, indent: -24 };
    const label = Math.max(0, -d.indent);
    const props = `<style:list-level-properties${odfAttrs({ "text:space-before": cm(Math.max(0, d.marL - label)), "text:min-label-width": cm(label) })}/>`;
    const b = d.bullet;
    const color = b.type !== "none" && b.color ? { "fo:color": sh.color(b.color) } : { "style:use-window-font-color": "true" };
    const size = b.type !== "none" && b.size ? `${Math.round(b.size)}%` : "100%";
    if (b.type === "auto") {
      const [fmt, suf] = NUM[b.scheme] ?? NUM.arabicPeriod;
      out += `<text:list-level-style-number${odfAttrs({ "text:level": k + 1, "style:num-format": fmt, "style:num-suffix": suf, "text:start-value": b.startAt && b.startAt > 1 ? b.startAt : undefined })}>${props}<style:text-properties${odfAttrs({ ...color, "fo:font-size": size })}/></text:list-level-style-number>`;
    } else if (b.type === "char") {
      out += `<text:list-level-style-bullet${odfAttrs({ "text:level": k + 1, "text:bullet-char": b.char || "•" })}>${props}<style:text-properties${odfAttrs({ "fo:font-family": b.font ?? undefined, ...color, "fo:font-size": size })}/></text:list-level-style-bullet>`;
    } else {
      out += `<text:list-level-style-number${odfAttrs({ "text:level": k + 1, "style:num-format": "" })}>${props}</text:list-level-style-number>`;
    }
  }
  return out;
}

const bulleted = (e: ParaProps) => !!e.bullet && e.bullet.type !== "none";

/** Metin gövdesini ODF paragraflarına yazar. */
export function textBodyXml(body: TextBody | undefined, ctx: TextWriteCtx): string {
  const paras = body?.paragraphs ?? [];
  let out = "";
  let i = 0;
  while (i < paras.length) {
    const eff = effectivePara(paras[i], ctx.lvl);
    if (!bulleted(eff)) {
      out += paraXml(paras[i], eff, false, ctx);
      i++;
      continue;
    }
    // Ardışık maddeli paragraflar tek liste.
    const group: { p: Paragraph; eff: ParaProps }[] = [];
    while (i < paras.length) {
      const e = effectivePara(paras[i], ctx.lvl);
      if (!bulleted(e)) break;
      group.push({ p: paras[i], eff: e });
      i++;
    }
    const levels = new Map<number, LevelDef>();
    for (const g of group) {
      if (levels.has(g.p.level)) continue;
      levels.set(g.p.level, { bullet: g.eff.bullet!, marL: g.eff.marL ?? 24 + g.p.level * 48, indent: g.eff.indent ?? -24 });
    }
    const ls = ctx.auto.list(listStyleXml(levels, ctx.sh));
    out += `<text:list text:style-name="${ls}">`;
    for (const g of group) {
      const depth = Math.max(0, Math.min(8, g.p.level));
      out += "<text:list-item>" + "<text:list><text:list-item>".repeat(depth) + paraXml(g.p, g.eff, true, ctx) + "</text:list-item></text:list>".repeat(depth) + "</text:list-item>";
    }
    out += "</text:list>";
  }
  return out;
}
