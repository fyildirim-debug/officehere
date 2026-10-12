/**
 * Asıl slayt ve düzenleri tema tanımından üretir. Yer tutucu konumları
 * 1280×720 tabanında verilir ve slayt boyutuna göre ölçeklenir.
 */
import { makeSp, NO_FILL, NO_LINE, uid } from "../model/factory";
import type { Align, Color, Fill, Layout, LayoutType, LevelStyle, Master, Placeholder, Shape, SpShape, TextBody, TextStyles } from "../model/types";

export interface ThemeStyle {
  bg: Fill;
  /** Başlık slaytı ve bölüm üst bilgisi arka planı. */
  titleBg?: Fill;
  deco?: (w: number, h: number) => Shape[];
  titleDeco?: (w: number, h: number) => Shape[];
  titleColor?: Color;
  bodyColor?: Color;
  titleAlign?: Align;
  titleCaps?: boolean;
  titleSize?: number;
  titleBold?: boolean;
  /** Başlık slaytındaki başlık rengi (titleBg koyuysa). */
  titleSlideColor?: Color;
  titleSlideSubColor?: Color;
  /** Yer tutucu alanı kenar boşlukları (1280×720 tabanında px). */
  area?: { l?: number; r?: number; t?: number };
  /** Bölüm üst bilgisi başlık slaytı görünümü yerine asıl slayt görünümünü kullansın. */
  secHeadUsesMaster?: boolean;
}

const BULLET: LevelStyle["para"]["bullet"] = { type: "char", char: "•", font: "Arial" };

export function buildTextStyles(st: ThemeStyle): TextStyles {
  const body: LevelStyle[] = [];
  const sizes = [28, 24, 20, 18, 18, 18, 18, 18, 18];
  for (let i = 0; i < 9; i++) {
    body.push({
      para: { align: "l", marL: 24 + i * 48, indent: -24, lineSpacing: 90, spaceBefore: i === 0 ? 10 : 5, bullet: BULLET },
      run: { font: "+mn", size: sizes[i], color: st.bodyColor ?? { scheme: "dk1" } },
    });
  }
  const title: LevelStyle[] = [
    {
      para: { align: st.titleAlign ?? "l", lineSpacing: 90, bullet: { type: "none" } },
      run: { font: "+mj", size: st.titleSize ?? 44, color: st.titleColor ?? { scheme: "dk1" }, caps: st.titleCaps ? "all" : undefined, bold: st.titleBold },
    },
  ];
  const other: LevelStyle[] = [];
  for (let i = 0; i < 9; i++) other.push({ para: { align: "l", marL: i * 48, indent: 0, lineSpacing: 100 }, run: { font: "+mn", size: 18, color: st.bodyColor ?? { scheme: "dk1" } } });
  return { title, body, other };
}

function ph(type: Placeholder["type"], idx: number | undefined, name: string, x: number, y: number, w: number, h: number, text: Partial<TextBody> = {}): SpShape {
  return makeSp("rect", x, y, w, h, {
    name,
    ph: idx === undefined ? { type } : { type, idx },
    fill: NO_FILL,
    line: NO_LINE,
    text: { paragraphs: [{ runs: [], level: 0, props: {} }], ...text },
  });
}

const noBullet = (size: number, extra: Partial<LevelStyle["run"]> = {}, align: Align = "l"): Partial<LevelStyle> => ({
  para: { bullet: { type: "none" }, marL: 0, indent: 0, align },
  run: { size, ...extra },
});

/** Alt bilgi yer tutucuları (tarih, alt bilgi, slayt numarası). */
function footers(sx: number, sy: number, area: { l: number; r: number }, color?: Color): SpShape[] {
  const y = 668 * sy;
  const hh = 36 * sy;
  const c: Color = color ?? { scheme: "dk1", lumMod: 0.75, lumOff: 0.25 };
  const L = area.l * sx;
  const R = (1280 - area.r) * sx;
  const W = R - L;
  return [
    ph("dt", 10, "Tarih Yer Tutucusu", L, y, W * 0.26, hh, { anchor: "ctr", levels: [{ para: { align: "l" }, run: { size: 12, color: c } }] }),
    ph("ftr", 11, "Alt Bilgi Yer Tutucusu", L + W * 0.3, y, W * 0.4, hh, { anchor: "ctr", levels: [{ para: { align: "ctr" }, run: { size: 12, color: c } }] }),
    ph("sldNum", 12, "Slayt Numarası Yer Tutucusu", R - W * 0.26, y, W * 0.26, hh, { anchor: "ctr", levels: [{ para: { align: "r" }, run: { size: 12, color: c } }] }),
  ];
}

export const LAYOUT_NAMES: Record<string, string> = {
  title: "Başlık Slaytı",
  obj: "Başlık ve İçerik",
  secHead: "Bölüm Üst Bilgisi",
  twoObj: "İki İçerik",
  twoTxTwoObj: "Karşılaştırma",
  titleOnly: "Yalnızca Başlık",
  blank: "Boş",
  objTx: "Resim Yazılı İçerik",
  picTx: "Resim Yazılı Resim",
};

export const LAYOUT_ORDER: LayoutType[] = ["title", "obj", "secHead", "twoObj", "twoTxTwoObj", "titleOnly", "blank", "objTx", "picTx"];

export function buildMasterAndLayouts(w: number, h: number, st: ThemeStyle): { master: Master; layouts: Layout[] } {
  const sx = w / 1280;
  const sy = h / 720;
  const area = { l: st.area?.l ?? 88, r: st.area?.r ?? 88, t: st.area?.t ?? 0 };
  const L = area.l * sx;
  const W = (1280 - area.l - area.r) * sx;
  const T = area.t * sy;
  const titleY = 38 * sy + T;
  const titleH = 140 * sy;
  const bodyY = 192 * sy + T * 0.5;
  const bodyH = 457 * sy - T * 0.5;
  const gap = 24 * sx;
  const half = (W - gap) / 2;

  const masterShapes: Shape[] = [
    ...(st.deco ? st.deco(w, h) : []),
    ph("title", undefined, "Başlık Yer Tutucusu", L, titleY, W, titleH, { anchor: "ctr", autofit: "shrink", prompt: "Asıl başlık stilini düzenlemek için tıklayın" }),
    ph("body", 1, "Metin Yer Tutucusu", L, bodyY, W, bodyH, { anchor: "t", autofit: "shrink", prompt: "Asıl metin stillerini düzenlemek için tıklayın" }),
    ...footers(sx, sy, area, st.bodyColor ? { ...st.bodyColor, alpha: 0.65 } : undefined),
  ];
  const master: Master = { background: st.bg, shapes: masterShapes, textStyles: buildTextStyles(st) };

  const titleCol = st.titleSlideColor;
  const subCol = st.titleSlideSubColor;
  const mk = (type: LayoutType, shapes: Shape[], extra: Partial<Layout> = {}): Layout => ({
    id: uid("L"),
    type,
    name: LAYOUT_NAMES[type],
    shapes: [...shapes, ...footers(sx, sy, area, (type === "title" || (type === "secHead" && !st.secHeadUsesMaster) ? subCol : undefined) ?? (st.bodyColor ? { ...st.bodyColor, alpha: 0.65 } : undefined))],
    showMasterShapes: true,
    ...extra,
  });

  const titleDeco = st.titleDeco ? st.titleDeco(w, h) : undefined;
  const titleExtra: Partial<Layout> = titleDeco ? { showMasterShapes: false, background: st.titleBg } : st.titleBg ? { background: st.titleBg } : {};

  const layouts: Layout[] = [
    mk(
      "title",
      [
        ...(titleDeco ?? []),
        ph("ctrTitle", undefined, "Başlık 1", 160 * sx, 118 * sy, 960 * sx, 250 * sy, {
          anchor: "b",
          autofit: "shrink",
          levels: [{ para: { align: st.titleAlign === "l" ? "l" : "ctr" }, run: { size: 60, color: titleCol } }],
        }),
        ph("subTitle", 1, "Alt Başlık 2", 160 * sx, 378 * sy, 960 * sx, 174 * sy, {
          anchor: "t",
          autofit: "shrink",
          levels: [0, 1, 2, 3, 4].map((i) => noBullet([24, 20, 18, 16, 16][i], subCol ? { color: subCol } : {}, st.titleAlign === "l" ? "l" : "ctr")),
        }),
      ],
      titleExtra,
    ),
    mk("obj", [ph("title", undefined, "Başlık 1", L, titleY, W, titleH), ph("obj", 1, "İçerik Yer Tutucusu 2", L, bodyY, W, bodyH)]),
    mk(
      "secHead",
      [
        ...(st.secHeadUsesMaster ? [] : (titleDeco ?? [])),
        ph("title", undefined, "Başlık 1", L, 180 * sy, W, 300 * sy, { anchor: "b", levels: [{ run: { size: 60, color: st.secHeadUsesMaster ? undefined : titleCol } }] }),
        ph("body", 1, "Metin Yer Tutucusu 2", L, 482 * sy, W, 158 * sy, { levels: [noBullet(24, { color: (st.secHeadUsesMaster ? undefined : subCol) ?? { scheme: "dk1", lumMod: 0.75, lumOff: 0.25 } })] }),
      ],
      st.secHeadUsesMaster ? {} : titleExtra,
    ),
    mk("twoObj", [
      ph("title", undefined, "Başlık 1", L, titleY, W, titleH),
      ph("obj", 1, "İçerik Yer Tutucusu 2", L, bodyY, half, bodyH, { levels: [{ run: { size: 28 } }] }),
      ph("obj", 2, "İçerik Yer Tutucusu 3", L + half + gap, bodyY, half, bodyH, { levels: [{ run: { size: 28 } }] }),
    ]),
    mk("twoTxTwoObj", [
      ph("title", undefined, "Başlık 1", L, titleY, W, titleH),
      ph("body", 1, "Metin Yer Tutucusu 2", L, 176 * sy + T * 0.5, half, 86 * sy, { anchor: "b", levels: [noBullet(24, { bold: true })] }),
      ph("obj", 2, "İçerik Yer Tutucusu 3", L, 263 * sy + T * 0.5, half, 387 * sy - T * 0.5),
      ph("body", 3, "Metin Yer Tutucusu 4", L + half + gap, 176 * sy + T * 0.5, half, 86 * sy, { anchor: "b", levels: [noBullet(24, { bold: true })] }),
      ph("obj", 4, "İçerik Yer Tutucusu 5", L + half + gap, 263 * sy + T * 0.5, half, 387 * sy - T * 0.5),
    ]),
    mk("titleOnly", [ph("title", undefined, "Başlık 1", L, titleY, W, titleH)]),
    mk("blank", []),
    mk("objTx", [
      ph("title", undefined, "Başlık 1", L, 48 * sy + T, W * 0.374, 168 * sy, { anchor: "b", levels: [{ run: { size: 32 } }] }),
      ph("obj", 1, "İçerik Yer Tutucusu 2", L + W * 0.415, 104 * sy + T * 0.5, W * 0.585, 512 * sy - T * 0.5, { levels: [{ run: { size: 32 } }, { run: { size: 28 } }, { run: { size: 24 } }] }),
      ph("body", 2, "Metin Yer Tutucusu 3", L, 216 * sy + T, W * 0.374, 400 * sy - T, { levels: [noBullet(16)] }),
    ]),
    mk("picTx", [
      ph("title", undefined, "Başlık 1", L, 48 * sy + T, W * 0.374, 168 * sy, { anchor: "b", levels: [{ run: { size: 32 } }] }),
      ph("pic", 1, "Resim Yer Tutucusu 2", L + W * 0.415, 104 * sy + T * 0.5, W * 0.585, 512 * sy - T * 0.5, { levels: [noBullet(32)] }),
      ph("body", 2, "Metin Yer Tutucusu 3", L, 216 * sy + T, W * 0.374, 400 * sy - T, { levels: [noBullet(16)] }),
    ]),
  ];
  return { master, layouts };
}
