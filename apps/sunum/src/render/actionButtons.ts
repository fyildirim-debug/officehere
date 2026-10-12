/**
 * Eylem düğmeleri (ECMA-376 actionButton* hazır şekilleri): dikdörtgen gövde
 * ve ortada koyulaştırılmış dolgulu simge. Yollar bu proje için yazılmıştır.
 */
import type { GeomPath, GeomResult } from "./geometry";

const f = (n: number) => (Math.round(n * 100) / 100).toString();
const P = (x: number, y: number) => `${f(x)} ${f(y)}`;
const poly = (pts: [number, number][]) => pts.map((p, i) => `${i ? "L" : "M"}${P(p[0], p[1])}`).join(" ") + " Z";

export function actionButtonGeometry(geom: string, w: number, h: number): GeomResult {
  const body: GeomPath = { d: poly([[0, 0], [w, 0], [w, h], [0, h]]), fill: "norm", stroke: true };
  const ss = Math.min(w, h);
  const cx = w / 2;
  const cy = h / 2;
  const d = ss * 0.375; // simgenin yarı genişliği
  const icon: GeomPath[] = [];
  const dark = (dd: string, mode: GeomPath["fill"] = "darken"): void => void icon.push({ d: dd, fill: mode, stroke: true });
  switch (geom) {
    case "actionButtonBackPrevious":
      dark(poly([[cx - d, cy], [cx + d, cy - d], [cx + d, cy + d]]));
      break;
    case "actionButtonForwardNext":
      dark(poly([[cx + d, cy], [cx - d, cy - d], [cx - d, cy + d]]));
      break;
    case "actionButtonBeginning": {
      const b = d * 0.25;
      dark(poly([[cx - d + b * 1.5, cy], [cx + d, cy - d], [cx + d, cy + d]]));
      dark(poly([[cx - d, cy - d], [cx - d + b, cy - d], [cx - d + b, cy + d], [cx - d, cy + d]]));
      break;
    }
    case "actionButtonEnd": {
      const b = d * 0.25;
      dark(poly([[cx + d - b * 1.5, cy], [cx - d, cy - d], [cx - d, cy + d]]));
      dark(poly([[cx + d - b, cy - d], [cx + d, cy - d], [cx + d, cy + d], [cx + d - b, cy + d]]));
      break;
    }
    case "actionButtonHome":
      dark(poly([[cx, cy - d], [cx + d, cy], [cx + d * 0.7, cy], [cx + d * 0.7, cy + d], [cx - d * 0.7, cy + d], [cx - d * 0.7, cy], [cx - d, cy]]));
      dark(poly([[cx - d * 0.2, cy + d * 0.3], [cx + d * 0.2, cy + d * 0.3], [cx + d * 0.2, cy + d], [cx - d * 0.2, cy + d]]), "darkenLess");
      break;
    case "actionButtonInformation":
      dark(`M${P(cx - d, cy)} A${f(d)} ${f(d)} 0 1 1 ${P(cx + d, cy)} A${f(d)} ${f(d)} 0 1 1 ${P(cx - d, cy)} Z`);
      dark(`M${P(cx - d * 0.14, cy - d * 0.55)} A${f(d * 0.14)} ${f(d * 0.14)} 0 1 1 ${P(cx + d * 0.14, cy - d * 0.55)} A${f(d * 0.14)} ${f(d * 0.14)} 0 1 1 ${P(cx - d * 0.14, cy - d * 0.55)} Z`, "lighten");
      dark(poly([[cx - d * 0.2, cy - d * 0.2], [cx + d * 0.12, cy - d * 0.2], [cx + d * 0.12, cy + d * 0.5], [cx + d * 0.25, cy + d * 0.5], [cx + d * 0.25, cy + d * 0.65], [cx - d * 0.25, cy + d * 0.65], [cx - d * 0.25, cy + d * 0.5], [cx - d * 0.12, cy + d * 0.5], [cx - d * 0.12, cy - d * 0.05], [cx - d * 0.2, cy - d * 0.05]]), "lighten");
      break;
    case "actionButtonReturn": {
      // U dönüşü ok: aşağı inip sola dönen kalın yol ve ok başı.
      const t = d * 0.3;
      dark(
        `M${P(cx + d, cy - d * 0.6)} L${P(cx + d, cy + d * 0.3)} A${f(d * 0.7)} ${f(d * 0.7)} 0 0 1 ${P(cx + d * 0.3, cy + d)} L${P(cx - d * 0.3, cy + d)} L${P(cx - d * 0.3, cy + d - t)} L${P(cx + d * 0.3, cy + d - t)} A${f(d * 0.4)} ${f(d * 0.4)} 0 0 0 ${P(cx + d - t, cy + d * 0.3)} L${P(cx + d - t, cy - d * 0.6)} Z`,
      );
      dark(poly([[cx - d, cy + d - t / 2], [cx - d * 0.3, cy + d * 0.45], [cx - d * 0.3, cy + d * 1.05 + t / 2 - t]]));
      break;
    }
    case "actionButtonDocument": {
      const c = d * 0.35;
      dark(poly([[cx - d * 0.7, cy - d], [cx + d * 0.7 - c, cy - d], [cx + d * 0.7, cy - d + c], [cx + d * 0.7, cy + d], [cx - d * 0.7, cy + d]]));
      dark(poly([[cx + d * 0.7 - c, cy - d], [cx + d * 0.7 - c, cy - d + c], [cx + d * 0.7, cy - d + c]]), "darkenLess");
      break;
    }
    case "actionButtonHelp":
      dark(
        `M${P(cx - d * 0.55, cy - d * 0.35)} A${f(d * 0.55)} ${f(d * 0.55)} 0 1 1 ${P(cx + d * 0.2, cy + d * 0.15)} L${P(cx + d * 0.15, cy + d * 0.45)} L${P(cx - d * 0.15, cy + d * 0.45)} L${P(cx - d * 0.15, cy + d * 0.05)} A${f(d * 0.32)} ${f(d * 0.32)} 0 1 0 ${P(cx - d * 0.25, cy - d * 0.35)} Z`,
      );
      dark(`M${P(cx - d * 0.16, cy + d * 0.78)} A${f(d * 0.16)} ${f(d * 0.16)} 0 1 1 ${P(cx + d * 0.16, cy + d * 0.78)} A${f(d * 0.16)} ${f(d * 0.16)} 0 1 1 ${P(cx - d * 0.16, cy + d * 0.78)} Z`);
      break;
    default:
      break;
  }
  return { paths: [body, ...icon], text: { x: 0, y: 0, w, h } };
}

/** Düğme türüne göre PowerPoint'in önerdiği varsayılan eylem (köprü). */
export function defaultActionFor(geom: string): string | undefined {
  switch (geom) {
    case "actionButtonBackPrevious":
      return "#prev";
    case "actionButtonForwardNext":
      return "#next";
    case "actionButtonBeginning":
    case "actionButtonHome":
      return "#first";
    case "actionButtonEnd":
      return "#last";
    case "actionButtonReturn":
      return "#prev";
    default:
      return undefined;
  }
}
