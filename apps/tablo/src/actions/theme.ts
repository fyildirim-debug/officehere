/** Sayfa Düzeni > Temalar: tema, renk ve yazı tipi şeması değişikliği (tek geri alma adımı). */
import type { Controller } from "../controller";
import { remapStyle, themeColorMap, type WorkbookTheme } from "../model/theme";
import { activateTheme } from "../model/themeRuntime";

/** Temayı değiştirir; palet renkleriyle boyanmış hücreler yeni temanın karşılığını alır. */
export function applyTheme(ctl: Controller, patch: Partial<WorkbookTheme>, label = "Tema"): void {
  const wb = ctl.wb;
  const old = wb.theme;
  const next: WorkbookTheme = { ...old, ...patch, colors: [...(patch.colors ?? old.colors)] as WorkbookTheme["colors"] };
  const map = themeColorMap(old.colors, next.colors);
  ctl.ops.snapshotOp(label, () => {
    wb.theme = next;
    if (map.size) {
      const cache = new Map<number, number>();
      const remapId = (id: number): number => {
        const hit = cache.get(id);
        if (hit !== undefined) return hit;
        const s = remapStyle(wb.styles.get(id), map);
        const nid = s ? wb.styles.intern(s) : id;
        cache.set(id, nid);
        return nid;
      };
      for (const sh of wb.sheets) {
        for (const m of [sh.styles, sh.colStyles, sh.rowStyles]) for (const [k, id] of m) m.set(k, remapId(id));
        for (const r of sh.condFormats) {
          if (r.style) r.style = remapStyle(r.style, map) ?? r.style;
          if (r.color) r.color = map.get(r.color.toUpperCase()) ?? r.color;
        }
        for (const s of sh.shapes) {
          s.fill = map.get(s.fill.toUpperCase()) ?? s.fill;
          s.stroke = map.get(s.stroke.toUpperCase()) ?? s.stroke;
        }
        if (sh.tabColor) sh.tabColor = map.get(sh.tabColor.toUpperCase()) ?? sh.tabColor;
      }
    }
    activateTheme(next);
  });
}
