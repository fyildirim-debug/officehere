/** Sağ tıklanan (seçili) nesnelerin hangi bağlam menüsünü alacağı. */
import type { Shape } from "../../model/types";
import { isLineGeom } from "../../render/geometry";

export type ObjectMenuKind = "none" | "multi" | "picture" | "media" | "chart" | "table" | "group" | "text" | "line" | "shape";

export function objectMenuKind(shapes: Shape[]): ObjectMenuKind {
  if (!shapes.length) return "none";
  if (shapes.length > 1) return "multi";
  const s = shapes[0];
  switch (s.kind) {
    case "pic":
      return s.media ? "media" : "picture";
    case "chart":
      return "chart";
    case "table":
      return "table";
    case "group":
      return "group";
    case "sp":
      // Mürekkep çizimi, çizgi ve bağlayıcıda metin ya da nokta düzenleme yoktur.
      if (s.ink || s.connector || isLineGeom(s.geom)) return "line";
      if (s.txBox || s.ph) return "text";
      return "shape";
  }
}
