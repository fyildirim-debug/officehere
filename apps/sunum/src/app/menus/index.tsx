/** Sağ tık menüleri: nesne türüne göre seçilen menü ve slayt/küçük resim/bölüm menüleri. */
import type { ReactNode } from "react";
import { chartMenu } from "./chart";
import type { MenuDeps } from "./common";
import { groupMenu, multiMenu } from "./group";
import { objectMenuKind } from "./kind";
import { lineMenu } from "./line";
import { mediaMenu, pictureMenu } from "./picture";
import { shapeMenu } from "./shape";
import { tableMenu } from "./table";
import { textBoxMenu } from "./text";

export type { MenuDeps } from "./common";
export { objectMenuKind, type ObjectMenuKind } from "./kind";
export { masterMenu, sectionMenu, slideAreaMenu, thumbMenu } from "./slide";

/** Seçili nesne(ler)e özgü menü; seçim yoksa null. */
export function objectMenu(d: MenuDeps): ReactNode | null {
  const shapes = d.a.selectedShapes();
  const s = shapes[0];
  switch (objectMenuKind(shapes)) {
    case "multi":
      return multiMenu(d);
    case "picture":
      return s.kind === "pic" ? pictureMenu(d, s) : null;
    case "media":
      return s.kind === "pic" ? mediaMenu(d, s) : null;
    case "chart":
      return chartMenu(d);
    case "table":
      return s.kind === "table" ? tableMenu(d, s) : null;
    case "group":
      return groupMenu(d);
    case "text":
      return s.kind === "sp" ? textBoxMenu(d, s) : null;
    case "line":
      return lineMenu(d);
    case "shape":
      return s.kind === "sp" ? shapeMenu(d, s) : null;
    default:
      return null;
  }
}
