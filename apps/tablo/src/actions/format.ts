/** Biçimlendirme eylemleri (Giriş sekmesi ve kısayollar). */
import type { Controller } from "../controller";
import { isDateFormat } from "@officehere/formula";
import { ACCOUNTING_FORMAT, COMMA_FORMAT, DEFAULT_FONT, DEFAULT_SIZE, type CellStyle } from "../model/styles";
import type { Rect } from "../model/types";
import { FONT_SIZES } from "@officehere/core";

export function activeStyle(ctl: Controller): CellStyle {
  const { row, col } = ctl.sel.active;
  return ctl.wb.styleAt(ctl.sheet, row, col);
}

function ranges(ctl: Controller): Rect[] {
  return ctl.sel.ranges;
}

export function applyStyle(ctl: Controller, patch: Partial<CellStyle> | ((s: CellStyle) => Partial<CellStyle>), label?: string): void {
  if (ctl.edit) ctl.commitEdit(null);
  if (ctl.sheet.protected) {
    ctl.alert("Korumalı bir sayfada biçimlendirme yapılamaz. Önce Gözden Geçir sekmesinden sayfa korumasını kaldırın.");
    return;
  }
  ctl.ops.applyStyle(ctl.sheet, ranges(ctl), patch, label);
}

export function toggle(ctl: Controller, key: "bold" | "italic" | "strike"): void {
  const on = !activeStyle(ctl)[key];
  const labels = { bold: "Kalın", italic: "İtalik", strike: "Üstü Çizili" };
  applyStyle(ctl, { [key]: on }, labels[key]);
}

export function toggleUnderline(ctl: Controller, kind: "single" | "double" = "single"): void {
  const cur = activeStyle(ctl).underline;
  applyStyle(ctl, { underline: cur === kind ? undefined : kind }, "Altı Çizili");
}

export function setFont(ctl: Controller, font: string): void {
  applyStyle(ctl, { font: font === DEFAULT_FONT ? undefined : font }, "Yazı Tipi");
}

export function setFontSize(ctl: Controller, size: number): void {
  if (!(size >= 1 && size <= 409)) return;
  applyStyle(ctl, { size: size === DEFAULT_SIZE ? undefined : size }, "Yazı Tipi Boyutu");
}

export function growFont(ctl: Controller, dir: 1 | -1): void {
  applyStyle(
    ctl,
    (s) => {
      const cur = s.size ?? DEFAULT_SIZE;
      const next = dir > 0 ? FONT_SIZES.find((x) => x > cur) ?? cur + 2 : [...FONT_SIZES].reverse().find((x) => x < cur) ?? Math.max(1, cur - 1);
      return { size: next };
    },
    dir > 0 ? "Yazı Tipi Boyutunu Büyüt" : "Yazı Tipi Boyutunu Küçült",
  );
}

export function setColor(ctl: Controller, color: string | null): void {
  applyStyle(ctl, { color: color ?? undefined }, "Yazı Tipi Rengi");
}

export function setFill(ctl: Controller, color: string | null): void {
  applyStyle(ctl, { fill: color ?? undefined, pattern: undefined }, "Dolgu Rengi");
}

export function setHAlign(ctl: Controller, a: CellStyle["hAlign"]): void {
  const cur = activeStyle(ctl).hAlign;
  applyStyle(ctl, { hAlign: cur === a ? undefined : a }, "Hizalama");
}

export function setVAlign(ctl: Controller, a: CellStyle["vAlign"]): void {
  applyStyle(ctl, { vAlign: a === "bottom" ? undefined : a }, "Hizalama");
}

export function toggleWrap(ctl: Controller): void {
  applyStyle(ctl, { wrap: !activeStyle(ctl).wrap }, "Metni Kaydır");
}

export function indent(ctl: Controller, dir: 1 | -1): void {
  applyStyle(
    ctl,
    (s) => {
      const n = Math.max(0, Math.min(250, (s.indent ?? 0) + dir));
      return { indent: n, hAlign: n > 0 && (!s.hAlign || s.hAlign === "general") ? "left" : s.hAlign };
    },
    dir > 0 ? "Girintiyi Artır" : "Girintiyi Azalt",
  );
}

export function setRotation(ctl: Controller, deg: number): void {
  const cur = activeStyle(ctl).rotation ?? 0;
  applyStyle(ctl, { rotation: cur === deg ? undefined : deg }, "Yönlendirme");
}

export function setNumFmt(ctl: Controller, code: string): void {
  applyStyle(ctl, { numFmt: code === "General" ? undefined : code }, "Sayı Biçimi");
}

/** Ondalık artır/azalt: biçim kodundaki ondalık basamak sayısını değiştirir. */
export function changeDecimals(ctl: Controller, dir: 1 | -1): void {
  const { row, col } = ctl.sel.active;
  const v = ctl.wb.getValue(ctl.sheet, row, col);
  applyStyle(
    ctl,
    (s) => {
      let code = s.numFmt && s.numFmt !== "General" ? s.numFmt : null;
      if (code && isDateFormat(code)) return {};
      if (!code) {
        // Genel: mevcut değerin ondalık basamaklarından başla
        const dec = typeof v === "number" ? Math.min(10, (String(v).split(".")[1] ?? "").length) : 0;
        code = dec ? `0.${"0".repeat(dec)}` : "0";
      }
      return { numFmt: adjustDecimals(code, dir) };
    },
    dir > 0 ? "Ondalık Artır" : "Ondalık Azalt",
  );
}

export function adjustDecimals(code: string, dir: 1 | -1): string {
  return code
    .split(";")
    .map((sec) => {
      // Tırnak içini koru
      const parts = sec.split(/("[^"]*")/);
      let done = false;
      for (let i = parts.length - 1; i >= 0 && !done; i--) {
        if (parts[i].startsWith('"')) continue;
        const p = parts[i];
        const m = /([0#?])(\.([0#?]+))?(?![\s\S]*[0#?])/.exec(p);
        if (!m) continue;
        const idx = m.index;
        let rep: string;
        if (dir > 0) rep = m[2] ? `${m[1]}.${m[3]}0` : `${m[1]}.0`;
        else if (m[3] && m[3].length > 1) rep = `${m[1]}.${m[3].slice(0, -1)}`;
        else rep = m[1];
        parts[i] = p.slice(0, idx) + rep + p.slice(idx + m[0].length);
        done = true;
      }
      return parts.join("");
    })
    .join(";");
}

export function accounting(ctl: Controller): void {
  setNumFmt(ctl, ACCOUNTING_FORMAT);
}

export function percent(ctl: Controller): void {
  setNumFmt(ctl, "0%");
}

export function comma(ctl: Controller): void {
  setNumFmt(ctl, COMMA_FORMAT);
}

export function borders(ctl: Controller, kind: string, color = "#000000", style: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" = "thin"): void {
  if (ctl.edit) ctl.commitEdit(null);
  for (const r of ctl.sel.ranges) ctl.ops.applyBorders(ctl.sheet, r, kind, { style, color });
}

export function mergeCells(ctl: Controller, mode: "center" | "across" | "merge" | "unmerge"): void {
  if (ctl.edit) ctl.commitEdit(null);
  const sheet = ctl.sheet;
  if (mode === "unmerge") {
    ctl.ops.unmerge(sheet, ctl.sel.ranges);
    return;
  }
  const already = ctl.sel.ranges.every((r) => sheet.merges.some((m) => m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2));
  if (mode === "center" && already) {
    ctl.ops.unmerge(sheet, ctl.sel.ranges);
    return;
  }
  const doMerge = () => {
    ctl.ops.merge(sheet, ctl.sel.ranges, mode);
    const r = ctl.range;
    ctl.selectRange(r, { row: r.r1, col: r.c1 });
  };
  if (ctl.ops.mergeLosesData(sheet, ctl.sel.ranges)) {
    ctl.openDialog("confirm", {
      title: "OfficeHere Tablo",
      message: "Hücreleri birleştirmek yalnızca sol üst değeri korur ve diğer değerleri atar.",
      onOk: doMerge,
    });
  } else doMerge();
}

export function clearFormats(ctl: Controller): void {
  ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "formats");
}

export function applyCellStyle(ctl: Controller, style: CellStyle): void {
  applyStyle(
    ctl,
    (s) => {
      const keep: Partial<CellStyle> = { numFmt: style.numFmt ?? s.numFmt, hAlign: s.hAlign, vAlign: s.vAlign, wrap: s.wrap, indent: s.indent };
      const reset: Partial<CellStyle> = { font: undefined, size: undefined, bold: undefined, italic: undefined, underline: undefined, color: undefined, fill: undefined, bTop: undefined, bBottom: undefined, bLeft: undefined, bRight: undefined };
      return { ...reset, ...keep, ...style };
    },
    "Hücre Stili",
  );
}
