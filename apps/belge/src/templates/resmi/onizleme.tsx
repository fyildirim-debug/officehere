// Resmî yazı şablonlarının küçük sayfa önizlemeleri (SVG; şablon galerisi ve Dosya › Yeni kutucukları).
import type { ReactNode } from "react";

const GRI = "#c8c8c8";
const KOYU = "#7a7a7a";
const KIRMIZI = "#c00000";

type Cubuk = [x: number, y: number, w: number, renk?: string, h?: number];

/** A4 oranlı sayfa: 210 × 297 birim (mm); yazı alanı 15…195. */
function sayfa(cubuklar: Cubuk[], ek: ReactNode = null, yatay = false): ReactNode {
  const w = yatay ? 220 : 210;
  const h = yatay ? 110 : 297;
  return (
    <svg className="bl-resmi-onizleme" viewBox={`0 0 ${w} ${h}`} width="100%" height="100%" preserveAspectRatio="xMidYMid meet" aria-hidden="true" focusable="false">
      <rect x="0" y="0" width={w} height={h} fill="#ffffff" />
      {cubuklar.map(([x, y, cw, renk = GRI, ch = 3.2], i) => (
        <rect key={i} x={x} y={y} width={cw} height={ch} fill={renk} />
      ))}
      {ek}
    </svg>
  );
}

const ortada = (y: number, w: number, renk = KOYU, cx = 105): Cubuk => [cx - w / 2, y, w, renk];
const baslik = (y = 15): Cubuk[] => [ortada(y, 14), ortada(y + 6, 62), ortada(y + 12, 48)];
const sayiKonu = (y: number): Cubuk[] => [
  [15, y, 70, KOYU],
  [167, y, 28, KOYU],
  [15, y + 6, 80, KOYU],
];
const metin = (y: number, n: number): Cubuk[] => Array.from({ length: n }, (_, i): Cubuk => (i % 4 === 0 ? [24, y + i * 6, 171] : i % 4 === 3 ? [15, y + i * 6, 120] : [15, y + i * 6, 180]));
const imza = (y: number, x = 140): Cubuk[] => [
  [x + 8, y, 34, KOYU],
  [x + 12, y + 6, 26],
];
const altBilgi = (): Cubuk[] => [
  [15, 268, 180, KOYU, 0.6],
  [15, 271, 96, GRI, 2],
  [15, 275, 110, GRI, 2],
  [15, 279, 84, GRI, 2],
  [150, 271, 40, GRI, 2],
  [158, 275, 24, GRI, 2],
  [150, 279, 40, GRI, 2],
  [100, 286, 10, GRI, 2],
];

export const ONIZLEMELER: Record<string, ReactNode> = {
  "resmi-ust-yazi": sayfa([...baslik(), ...sayiKonu(45), ortada(68, 60), ortada(74, 46), [15, 92, 150], ...metin(104, 7), ...imza(160), [15, 178, 70], ...altBilgi()]),
  "resmi-dagitimli": sayfa([
    ...baslik(),
    ...sayiKonu(45),
    ortada(68, 54),
    ...metin(86, 6),
    ...imza(134),
    [15, 152, 60],
    [15, 162, 22, KOYU],
    [15, 168, 18],
    [100, 168, 16],
    [15, 174, 54],
    [100, 174, 52],
    [15, 180, 54],
    ...altBilgi(),
  ]),
  "resmi-olur": sayfa([
    ...baslik(),
    ...sayiKonu(45),
    ortada(68, 50),
    ...metin(86, 5),
    ...imza(124),
    [22, 146, 52],
    [26, 164, 34, KOYU],
    [30, 170, 26],
    ortada(186, 16, KOYU),
    ortada(194, 34, KOYU),
    ortada(200, 26),
    [15, 214, 64],
    ...altBilgi(),
  ]),
  "resmi-tekit": sayfa([[167, 15, 28, KIRMIZI], ...baslik(21), ...sayiKonu(51), ortada(74, 64), ortada(80, 52), [15, 98, 150], ...metin(110, 6), ...imza(160), ...altBilgi()]),
  "resmi-ek-dagitim": sayfa([
    ortada(18, 30, KOYU),
    [15, 34, 120],
    [15, 40, 132],
    [15, 46, 110],
    [15, 52, 126],
    [15, 58, 104],
    ortada(84, 40, KOYU),
    [15, 100, 22, KOYU],
    [100, 100, 16, KOYU],
    [15, 106, 60],
    [100, 106, 62],
    [15, 112, 56],
    [100, 112, 50],
    [15, 118, 64],
    ...altBilgi(),
  ]),
  "resmi-dilekce": sayfa([[167, 15, 28, KOYU], ortada(40, 80, KOYU), ...metin(58, 6), [150, 106, 34, KOYU], [158, 112, 18], [15, 130, 120], [15, 136, 70], [15, 142, 66], [15, 156, 60]]),
  "resmi-tutanak": sayfa(
    [...baslik(), ortada(45, 30, KOYU), ...metin(56, 8), [167, 110, 28, KOYU]],
    <g fill="none" stroke={KOYU} strokeWidth="0.6">
      <rect x="15" y="122" width="180" height="40" />
      <path d="M15 132h180M15 142h180M15 152h180M75 122v40M135 122v40" />
    </g>,
  ),
  "resmi-zarf": sayfa(
    [
      [12, 10, 10, KOYU],
      [12, 15, 50, KOYU],
      [12, 20, 42],
      [12, 27, 48],
      [12, 32, 30],
      [186, 10, 22, KIRMIZI],
      ortada(62, 70, KOYU, 110),
      ortada(68, 52, GRI, 110),
      ortada(74, 76, GRI, 110),
      ortada(80, 50, GRI, 110),
    ],
    null,
    true,
  ),
};

/** Dosya › Yeni'deki "Resmî yazışma" kutucuğu: üst yazı önizlemesi. */
export const RESMI_KUME_ONIZLEME = ONIZLEMELER["resmi-ust-yazi"];
