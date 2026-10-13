/**
 * Sunu özellikleri (ppt/presProps.xml, ECMA-376 §19.2.1.27): slayt gösterisi
 * ayarları (p:showPr) ve yazdırma ayarları (p:prnPr). Lazer işaretçisi rengi
 * PowerPoint 2010 uzantısıyla (p14:laserClr) saklanır.
 */
import type { PrintSettings, PrintWhat, Presentation } from "../model/types";
import { DEFAULT_SHOW_SETTINGS, type ShowSettings } from "../show/settings";
import { attr, boolAttr, descendants, kid, kids, NS, numAttr, parseXml } from "./xml";

const NS_DECL = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;
const LASER_EXT_URI = "{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}";

export const DEFAULT_PRINT_SETTINGS: PrintSettings = { what: "slides", color: "clr", hiddenSlides: false, frameSlides: false, orientation: "portrait" };

const PRINT_WHAT: PrintWhat[] = ["slides", "notes", "outline", "handouts1", "handouts2", "handouts3", "handouts4", "handouts6", "handouts9"];

const hex = (c: string) => c.replace(/^#/, "").toUpperCase();

function prnPrXml(p: PrintSettings): string {
  const a: string[] = [];
  if (p.what !== "slides") a.push(`prnWhat="${p.what}"`);
  if (p.color !== "clr") a.push(`clrMode="${p.color}"`);
  if (p.hiddenSlides) a.push('hiddenSlides="1"');
  if (p.frameSlides) a.push('frameSlides="1"');
  return `<p:prnPr${a.length ? ` ${a.join(" ")}` : ""}/>`;
}

function showPrXml(s: ShowSettings, pres: Presentation): string {
  const a = [`showNarration="${s.noNarration ? 0 : 1}"`];
  if (s.loop) a.unshift('loop="1"');
  if (s.noAnimations) a.push('showAnimation="0"');
  if (s.advance === "manual") a.push('useTimings="0"');
  const type = s.type === "kiosk" ? "<p:kiosk/>" : s.type === "browsed" ? '<p:browse showScrollbar="0"/>' : "<p:present/>";
  let range = "<p:sldAll/>";
  if (s.range.kind === "range") range = `<p:sldRg st="${s.range.from}" end="${s.range.to}"/>`;
  else if (s.range.kind === "custom") {
    const custom = s.range.id;
    const i = (pres.customShows ?? []).findIndex((c) => c.id === custom);
    if (i >= 0) range = `<p:custShow id="${i}"/>`;
  }
  return (
    `<p:showPr ${a.join(" ")}>${type}${range}<p:penClr><a:srgbClr val="${hex(s.penColor)}"/></p:penClr>` +
    `<p:extLst><p:ext uri="${LASER_EXT_URI}"><p14:laserClr xmlns:p14="${NS.p14}"><a:srgbClr val="${hex(s.laserColor)}"/></p14:laserClr></p:ext></p:extLst></p:showPr>`
  );
}

/** presProps.xml; ayar yoksa boş öğe (önceki çıktıyla aynı). Şema sırası: prnPr, showPr. */
export function presPropsXml(pres: Presentation): string {
  const body = (pres.printSettings ? prnPrXml(pres.printSettings) : "") + (pres.showSettings ? showPrXml(pres.showSettings, pres) : "");
  return body ? `<p:presentationPr ${NS_DECL}>${body}</p:presentationPr>` : `<p:presentationPr ${NS_DECL}/>`;
}

const PRESET_COLORS: Record<string, string> = { red: "#FF0000", black: "#000000", white: "#FFFFFF", blue: "#0000FF", green: "#008000", yellow: "#FFFF00", lime: "#00FF00", cyan: "#00FFFF", magenta: "#FF00FF" };

function colorOf(el: Element | undefined): string | undefined {
  const c = kids(el)[0];
  if (!c) return undefined;
  const v = attr(c, "val");
  if (c.localName === "srgbClr" && v && /^[0-9a-f]{6}$/i.test(v)) return `#${v.toUpperCase()}`;
  if (c.localName === "prstClr" && v) return PRESET_COLORS[v];
  if (c.localName === "sysClr") {
    const l = attr(c, "lastClr");
    return l ? `#${l.toUpperCase()}` : undefined;
  }
  return undefined;
}

/**
 * presProps.xml okuması. customShowIds: dosyadaki custShow kimliği → model kimliği.
 * showPr yoksa gösteri ayarları tanımsız kalır (varsayılanlar kullanılır).
 */
export function readPresProps(xml: string | undefined, customShowIds: Map<string, string>): { showSettings?: ShowSettings; printSettings?: PrintSettings } {
  if (!xml) return {};
  let root: Element;
  try {
    root = parseXml(xml).documentElement;
  } catch {
    return {};
  }
  const out: { showSettings?: ShowSettings; printSettings?: PrintSettings } = {};
  const prn = kid(root, "prnPr");
  if (prn) {
    const what = attr(prn, "prnWhat") as PrintWhat | undefined;
    const clr = attr(prn, "clrMode");
    out.printSettings = {
      ...DEFAULT_PRINT_SETTINGS,
      what: what && PRINT_WHAT.includes(what) ? what : "slides",
      color: clr === "gray" || clr === "bw" ? clr : "clr",
      hiddenSlides: !!boolAttr(prn, "hiddenSlides"),
      frameSlides: !!boolAttr(prn, "frameSlides"),
    };
  }
  const sp = kid(root, "showPr");
  if (sp) {
    const s: ShowSettings = { ...DEFAULT_SHOW_SETTINGS, range: { kind: "all" } };
    // Şema varsayılanları: loop=0, showNarration=0, showAnimation=1, useTimings=1.
    s.loop = !!boolAttr(sp, "loop");
    s.noNarration = !boolAttr(sp, "showNarration");
    s.noAnimations = boolAttr(sp, "showAnimation") === false;
    s.advance = boolAttr(sp, "useTimings") === false ? "manual" : "timings";
    s.type = kid(sp, "kiosk") ? "kiosk" : kid(sp, "browse") ? "browsed" : "speaker";
    const rg = kid(sp, "sldRg");
    const cs = kid(sp, "custShow");
    if (rg) {
      const st = numAttr(rg, "st") ?? 1;
      s.range = { kind: "range", from: st, to: Math.max(st, numAttr(rg, "end") ?? st) };
    } else if (cs) {
      const id = customShowIds.get(attr(cs, "id") ?? "");
      if (id) s.range = { kind: "custom", id };
    }
    s.penColor = colorOf(kid(sp, "penClr")) ?? s.penColor;
    s.laserColor = colorOf(descendants(sp, "laserClr")[0]) ?? s.laserColor;
    out.showSettings = s;
  }
  return out;
}

/** p:notesSz yönlendirmesi (cx > cy: yatay). */
export function notesOrientation(cx: number | undefined, cy: number | undefined): PrintSettings["orientation"] {
  return cx !== undefined && cy !== undefined && cx > cy ? "landscape" : "portrait";
}

