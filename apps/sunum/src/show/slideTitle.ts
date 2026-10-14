/** Slayt başlığı (başlık yer tutucusunun düz metni); gösteri dışındaki modüller de kullanır. */
import type { Slide } from "../model/types";

export function slideTitle(s: Slide): string {
  const t = s.shapes.find((x) => x.kind === "sp" && (x.ph?.type === "title" || x.ph?.type === "ctrTitle"));
  if (!t || t.kind !== "sp" || !t.text) return "";
  return t.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ").trim();
}
