/** Etkin temayı modül düzeyindeki varsayılanlara (gövde yazı tipi, tablo/grafik renkleri) uygular. */
import { setChartThemePalette } from "../charts/chartData";
import { setDefaultFont } from "./styles";
import { setTableThemeAccents } from "./tablestyles";
import type { WorkbookTheme } from "./theme";

let lastKey = "";

export function activateTheme(t: WorkbookTheme): void {
  const key = `${t.minorFont}|${t.colors.join(",")}`;
  if (key === lastKey) return;
  lastKey = key;
  setDefaultFont(t.minorFont);
  const accents = t.colors.slice(4);
  setTableThemeAccents(accents);
  setChartThemePalette(accents);
}
