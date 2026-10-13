/**
 * Yabancı biçim okuyucularının (ODP, PPT) ortak modeli kurma adımları: asıl
 * slayt görünümü, ikincil ana sayfalar için düzen kopyaları ve yer tutuculara
 * göre slayt düzeni seçimi.
 */
import { uid } from "../model/factory";
import type { Fill, Layout, LayoutType, Presentation, Shape, SpShape } from "../model/types";

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ImportMaster {
  key: string;
  display: string;
  bg: Fill;
  deco: Shape[];
  title?: Box;
  body?: Box;
}

/** Asıl slaydı içe aktarılan ana sayfanın arka planı, süsleri ve yer tutucu konumlarıyla günceller. */
export function applyPrimaryMaster(pres: Presentation, m: ImportMaster) {
  pres.master.background = m.bg;
  pres.master.shapes = [
    ...m.deco,
    ...pres.master.shapes.map((s) => {
      if (s.kind !== "sp" || !s.ph) return s;
      const src = s.ph.type === "title" ? m.title : s.ph.type === "body" ? m.body : undefined;
      return src ? { ...s, x: src.x, y: src.y, w: src.w, h: src.h } : s;
    }),
  ];
  pres.masterEdited = true;
}

/** Ana sayfa + düzen türü → modeldeki düzen (ikincil ana sayfalar için kopya). */
export function layoutResolver(pres: Presentation, masters: Map<string, ImportMaster>, primary: ImportMaster | undefined): (key: string | undefined, type: LayoutType) => Layout {
  const cache = new Map<string, Layout>();
  const baseLayout = (type: LayoutType) => pres.layouts.find((l) => l.type === type) ?? pres.layouts[0];
  return (key, type) => {
    const base = baseLayout(type);
    const m = key !== undefined ? masters.get(key) : undefined;
    if (!m || m === primary) return base;
    const ck = `${m.key}|${type}`;
    let l = cache.get(ck);
    if (!l) {
      l = { ...structuredClone(base), id: uid("L"), name: `${base.name} (${m.display})`, background: m.bg, showMasterShapes: false, shapes: [...m.deco.map((d) => ({ ...structuredClone(d), id: uid() })), ...structuredClone(base.shapes)] };
      cache.set(ck, l);
      pres.layouts.push(l);
    }
    return l;
  };
}

/** Slayttaki yer tutuculara göre düzen türü; yer tutucu türleri/dizinleri düzene uyarlanır. */
export function chooseLayout(shapes: Shape[]): LayoutType {
  const sps = shapes.filter((s): s is SpShape => s.kind === "sp" && !!s.ph);
  const title = sps.find((s) => s.ph!.type === "title" || s.ph!.type === "ctrTitle");
  const sub = sps.find((s) => s.ph!.type === "subTitle");
  const bodies = sps.filter((s) => s.ph!.type === "body" || s.ph!.type === "obj");
  let type: LayoutType = "blank";
  if (sub) type = "title";
  else if (title && bodies.length >= 2) type = "twoObj";
  else if (bodies.length) type = "obj";
  else if (title) type = "titleOnly";
  if (sub) {
    sub.ph = { type: "subTitle", idx: 1 };
    if (title) title.ph = { type: "ctrTitle" };
  } else if (title) title.ph = { type: "title" };
  bodies.forEach((b, k) => (b.ph = { type: type === "obj" || type === "twoObj" ? "obj" : "body", idx: k + 1 }));
  return type;
}
