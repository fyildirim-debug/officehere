/**
 * Yazı tipi listesi. Microsoft yazı tipleri dağıtılmaz; kullanıcının
 * sisteminde varsa kullanılır, yoksa ölçü uyumlu açık yazı tiplerine düşülür
 * (Carlito ↔ Calibri, Caladea ↔ Cambria, Liberation ↔ Arial/Times/Courier).
 */

export const FONT_FALLBACKS: Record<string, string> = {
  Aptos: '"Aptos", "Segoe UI", "Carlito", "Calibri", "Liberation Sans", Arial, sans-serif',
  "Aptos Display": '"Aptos Display", "Aptos", "Segoe UI", "Carlito", "Liberation Sans", Arial, sans-serif',
  Calibri: '"Calibri", "Carlito", "Liberation Sans", Arial, sans-serif',
  "Calibri Light": '"Calibri Light", "Calibri", "Carlito", "Liberation Sans", Arial, sans-serif',
  Cambria: '"Cambria", "Caladea", "Liberation Serif", "Times New Roman", serif',
  Arial: 'Arial, "Liberation Sans", "Arimo", Helvetica, sans-serif',
  "Times New Roman": '"Times New Roman", "Liberation Serif", "Tinos", Times, serif',
  "Courier New": '"Courier New", "Liberation Mono", "Cousine", monospace',
  Georgia: "Georgia, serif",
  Verdana: "Verdana, sans-serif",
  Tahoma: "Tahoma, Verdana, sans-serif",
  "Trebuchet MS": '"Trebuchet MS", sans-serif',
  "Segoe UI": '"Segoe UI", system-ui, sans-serif',
  "Comic Sans MS": '"Comic Sans MS", "Comic Neue", cursive',
  Impact: "Impact, sans-serif",
  Consolas: 'Consolas, "Cascadia Mono", "DejaVu Sans Mono", monospace',
  Garamond: 'Garamond, "EB Garamond", serif',
  "Book Antiqua": '"Book Antiqua", Palatino, serif',
};

export const FONT_LIST = Object.keys(FONT_FALLBACKS).sort((a, b) => a.localeCompare(b, "tr"));

export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** CSS font-family değeri (yedek zinciriyle). */
export function fontStack(name: string): string {
  return FONT_FALLBACKS[name] ?? `"${name}", sans-serif`;
}
