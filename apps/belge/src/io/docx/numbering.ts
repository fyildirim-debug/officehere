// Liste stilleri ↔ OOXML numaralandırma (numFmt + lvlText) eşlemesi.

export interface ListSig {
  type: "bullet" | "ordered";
  /** listStyle (null = varsayılan: disc / decimal). */
  style: string | null;
  multilevel: string | null;
}

export interface LevelDef {
  numFmt: string;
  lvlText: string;
  font?: string;
}

const BULLET_CHARS: Record<string, string> = {
  disc: "•",
  circle: "o",
  square: "▪",
  diamond: "❖",
  arrow: "➢",
  check: "✓",
  dash: "–",
};

const DEPTH_BULLETS = ["disc", "circle", "square"];

const ORDERED: Record<string, { fmt: string; suffix: string }> = {
  decimal: { fmt: "decimal", suffix: "." },
  "decimal-paren": { fmt: "decimal", suffix: ")" },
  "upper-roman": { fmt: "upperRoman", suffix: "." },
  "upper-alpha": { fmt: "upperLetter", suffix: "." },
  "lower-alpha-paren": { fmt: "lowerLetter", suffix: ")" },
  "lower-alpha": { fmt: "lowerLetter", suffix: "." },
  "lower-roman": { fmt: "lowerRoman", suffix: "." },
};

const MIXED = ["decimal", "lower-alpha", "lower-roman"];

export function sigKey(s: ListSig): string {
  return `${s.type}|${s.style ?? ""}|${s.multilevel ?? ""}`;
}

/** Düzey (0-8) tanımı. */
export function levelDef(s: ListSig, lvl: number): LevelDef {
  if (s.type === "bullet") {
    const style = s.style ?? DEPTH_BULLETS[lvl % 3];
    const ch = BULLET_CHARS[style] ?? "•";
    return { numFmt: "bullet", lvlText: ch, font: ch === "o" ? "Courier New" : undefined };
  }
  if (s.multilevel === "outline") {
    const text = Array.from({ length: lvl + 1 }, (_, i) => `%${i + 1}`).join(".") + ".";
    return { numFmt: "decimal", lvlText: text };
  }
  if (s.multilevel === "mixed") {
    const o = ORDERED[MIXED[lvl % 3]];
    return { numFmt: o.fmt, lvlText: `%${lvl + 1}${o.suffix}` };
  }
  // Varsayılan numaralandırma düzeyleri: 1. → a. → i. (yalnız ilk düzey değil, iç içe listeler de).
  const o = ORDERED[s.style ?? MIXED[lvl % 3]] ?? ORDERED.decimal;
  return { numFmt: o.fmt, lvlText: `%${lvl + 1}${o.suffix}` };
}

/** Okunan düzey tanımından liste imzası. */
export function sigFromLevel(numFmt: string, lvlText: string, lvl: number, allLevels?: { numFmt: string; lvlText: string }[]): ListSig {
  if (numFmt === "bullet" || numFmt === "none") {
    const ch = lvlText.trim();
    const hit = Object.entries(BULLET_CHARS).find(([, c]) => c === ch);
    let style: string | null = hit ? hit[0] : null;
    // Simge yazı tipinin özel kod noktaları (Symbol/Wingdings).
    if (!hit) {
      const code = ch.codePointAt(0) ?? 0;
      if (code === 0xf0b7 || ch === "●" || ch === "·") style = "disc";
      else if (code === 0xf0a7 || ch === "■") style = "square";
      else if (code === 0xf0d8 || code === 0xf0e8) style = "arrow";
      else if (code === 0xf0fc) style = "check";
      else if (code === 0xf076) style = "diamond";
      else if (ch === "-" || ch === "–") style = "dash";
      else if (ch === "○" || ch === "◦") style = "circle";
    }
    if (style && DEPTH_BULLETS[lvl % 3] === style) style = null;
    if (style === "disc") style = null;
    return { type: "bullet", style, multilevel: null };
  }
  // Çok düzeyli anahat: "%1.%2."
  if (/%\d+\.%\d+/.test(lvlText) || (allLevels && /%1\.%2/.test(allLevels[1]?.lvlText ?? ""))) return { type: "ordered", style: null, multilevel: "outline" };
  if (allLevels && allLevels[0]?.numFmt === "decimal" && allLevels[1]?.numFmt === "lowerLetter" && allLevels[2]?.numFmt === "lowerRoman") return { type: "ordered", style: null, multilevel: "mixed" };
  const suffix = lvlText.replace(/%\d+/g, "").trim();
  const paren = suffix === ")";
  let style: string | null = null;
  switch (numFmt) {
    case "upperRoman":
      style = "upper-roman";
      break;
    case "lowerRoman":
      style = "lower-roman";
      break;
    case "upperLetter":
      style = "upper-alpha";
      break;
    case "lowerLetter":
      style = paren ? "lower-alpha-paren" : "lower-alpha";
      break;
    default:
      style = paren ? "decimal-paren" : null;
  }
  return { type: "ordered", style, multilevel: null };
}
