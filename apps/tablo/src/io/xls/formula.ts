/**
 * BIFF8 ayrıştırılmış formül belirteçleri (ptg, MS-XLS §2.5.198) → A1 formülü.
 * Ters Lehçe gösterim yığınla çözülür; öncelik gerektirdiğinde parantez eklenir.
 * Çevrilemeyen belirteçte (dış başvuru, 3B aralık, makro işlevi) undefined döner.
 */
import { colToName, EN_SYNTAX, FUNCTIONS, parse, quoteSheetName, tokenize } from "@officehere/formula";
import { FTAB } from "./ftab";
import { f64, u16, u32, u8, readString } from "./records";

export interface XlsFormulaContext {
  sheetNames: string[];
  /** EXTERNSHEET: SUPBOOK dizini ve sayfa aralığı. */
  xti: { sup: number; first: number; last: number }[];
  supbooks: { kind: "self" | "addin" | "external"; externNames: string[] }[];
  /** NAME kayıtları (1 tabanlı ptgName dizini için sırayla). */
  names: { name: string; builtin: boolean }[];
}

const ERRORS: Record<number, string> = { 0x00: "#NULL!", 0x07: "#DIV/0!", 0x0f: "#VALUE!", 0x17: "#REF!", 0x1d: "#NAME?", 0x24: "#NUM!", 0x2a: "#N/A", 0x2b: "#GETTING_DATA" };

export function errorText(code: number): string {
  return ERRORS[code] ?? "#N/A";
}

const BIN: Record<number, [string, number]> = {
  0x03: ["+", 3], 0x04: ["-", 3], 0x05: ["*", 4], 0x06: ["/", 4], 0x07: ["^", 5], 0x08: ["&", 2],
  0x09: ["<", 1], 0x0a: ["<=", 1], 0x0b: ["=", 1], 0x0c: [">=", 1], 0x0d: [">", 1], 0x0e: ["<>", 1],
};

interface Item {
  t: string;
  p: number;
  /** ptgName'in gösterdiği ad (kullanıcı işlevi adı için). */
  name?: string;
}

const ATOM = 100;
const MAX_ROW8 = 0xffff;
const MAX_COL8 = 0xff;

class Unsupported extends Error {}

function cellText(row: number, col: number, rowRel: boolean, colRel: boolean): string {
  return `${colRel ? "" : "$"}${colToName(col)}${rowRel ? "" : "$"}${row + 1}`;
}

function strLit(s: string): string {
  return `"${s.replace(/"/g, '""')}"`;
}

/**
 * Belirteçleri formüle çevirir. base: hücrenin konumu (paylaşılan formüllerdeki
 * göreli ptgRefN/ptgAreaN başvuruları için).
 */
export function decodeFormula(rgce: Uint8Array, extra: Uint8Array, ctx: XlsFormulaContext, base: { row: number; col: number }): string | undefined {
  try {
    const text = decode(rgce, extra, ctx, base);
    if (text === undefined) return undefined;
    const f = "=" + text;
    parse(text);
    for (const t of tokenize(text, EN_SYNTAX)) if (t.type === "func" && !FUNCTIONS.has(String(t.value).toUpperCase())) return undefined;
    return f;
  } catch {
    return undefined;
  }
}

function decode(d: Uint8Array, extra: Uint8Array, ctx: XlsFormulaContext, base: { row: number; col: number }): string | undefined {
  const st: Item[] = [];
  let p = 0;
  let ep = 0;
  const pop = (): Item => {
    const it = st.pop();
    if (!it) throw new Unsupported("yığın boş");
    return it;
  };
  const sheetPrefix = (ixti: number): string | null => {
    const x = ctx.xti[ixti];
    if (!x) return null;
    const sb = ctx.supbooks[x.sup];
    if (!sb || sb.kind !== "self") throw new Unsupported("dış başvuru");
    if (x.first === 0xfffe || x.first === 0xffff) return null;
    if (x.first !== x.last) throw new Unsupported("3B başvuru");
    const name = ctx.sheetNames[x.first];
    if (name === undefined) return null;
    return quoteSheetName(name) + "!";
  };
  const ref = (rw: number, cf: number, rel: boolean): string => {
    const colRel = !!(cf & 0x4000);
    const rowRel = !!(cf & 0x8000);
    let row = rw;
    let col = cf & 0x3fff;
    if (rel) {
      if (rowRel) row = (base.row + rw) & 0xffff;
      if (colRel) col = (base.col + (cf & 0xff)) & 0xff;
    }
    return cellText(row, col, rowRel, colRel);
  };
  const area = (r1: number, r2: number, c1f: number, c2f: number, rel: boolean): string => {
    const c1 = c1f & 0x3fff;
    const c2 = c2f & 0x3fff;
    // Tüm sütun (1..65536) ve tüm satır (A..IV) başvuruları
    if (!rel && r1 === 0 && r2 === MAX_ROW8) return `${c1f & 0x4000 ? "" : "$"}${colToName(c1)}:${c2f & 0x4000 ? "" : "$"}${colToName(c2)}`;
    if (!rel && c1 === 0 && c2 === MAX_COL8) return `${c1f & 0x8000 ? "" : "$"}${r1 + 1}:${c2f & 0x8000 ? "" : "$"}${r2 + 1}`;
    return `${ref(r1, c1f, rel)}:${ref(r2, c2f, rel)}`;
  };
  let guard = 0;
  while (p < d.length) {
    if (guard++ > 10000) throw new Unsupported("çok uzun");
    const ptg = d[p++];
    if (ptg < 0x20) {
      const bin = BIN[ptg];
      if (bin) {
        const b = pop();
        const a = pop();
        const [op, prec] = bin;
        const left = a.p < prec ? `(${a.t})` : a.t;
        const right = b.p <= prec ? `(${b.t})` : b.t;
        st.push({ t: `${left}${op}${right}`, p: prec });
        continue;
      }
      switch (ptg) {
        case 0x01:
        case 0x02:
          throw new Unsupported("paylaşılan/tablo formülü");
        case 0x0f: {
          const b = pop();
          const a = pop();
          st.push({ t: `${a.t} ${b.t}`, p: 8 });
          break;
        }
        case 0x10: {
          const b = pop();
          const a = pop();
          st.push({ t: `${a.t},${b.t}`, p: 0 });
          break;
        }
        case 0x11: {
          const b = pop();
          const a = pop();
          st.push({ t: `${a.t}:${b.t}`, p: 9 });
          break;
        }
        case 0x12: {
          const a = pop();
          st.push({ t: `+${a.p < 7 ? `(${a.t})` : a.t}`, p: 7 });
          break;
        }
        case 0x13: {
          const a = pop();
          st.push({ t: `-${a.p < 7 ? `(${a.t})` : a.t}`, p: 7 });
          break;
        }
        case 0x14: {
          const a = pop();
          st.push({ t: `${a.p < 6 ? `(${a.t})` : a.t}%`, p: 6 });
          break;
        }
        case 0x15: {
          const a = pop();
          st.push({ t: `(${a.t})`, p: ATOM });
          break;
        }
        case 0x16:
          st.push({ t: "", p: ATOM });
          break;
        case 0x17: {
          const s = readString(d, p, 1);
          p = s.end;
          st.push({ t: strLit(s.text), p: ATOM });
          break;
        }
        case 0x19: {
          const grbit = d[p++];
          const data = u16(d, p);
          p += 2;
          if (grbit & 0x04) p += (data + 1) * 2;
          else if (grbit & 0x10) {
            const a = pop();
            st.push({ t: `SUM(${a.t})`, p: ATOM });
          }
          break;
        }
        case 0x1c:
          st.push({ t: errorText(d[p++]), p: ATOM });
          break;
        case 0x1d:
          st.push({ t: d[p++] ? "TRUE" : "FALSE", p: ATOM });
          break;
        case 0x1e:
          st.push({ t: String(u16(d, p)), p: ATOM });
          p += 2;
          break;
        case 0x1f: {
          const v = f64(d, p);
          p += 8;
          st.push({ t: Number.isFinite(v) ? String(v) : "#NUM!", p: ATOM });
          break;
        }
        default:
          throw new Unsupported(`belirteç ${ptg}`);
      }
      continue;
    }
    // Sınıflı belirteçler: 0x20 başvuru, 0x40 değer, 0x60 dizi sınıfı
    const id = (ptg & 0x1f) | 0x20;
    switch (id) {
      case 0x20: {
        p += 7;
        const cols = u8(extra, ep) + 1;
        const rows = u16(extra, ep + 1) + 1;
        ep += 3;
        const out: string[] = [];
        for (let r = 0; r < rows; r++) {
          const vals: string[] = [];
          for (let c = 0; c < cols; c++) {
            const t = u8(extra, ep++);
            if (t === 0x01) {
              vals.push(String(f64(extra, ep)));
              ep += 8;
            } else if (t === 0x02) {
              const s = readString(extra, ep, 2);
              ep = s.end;
              vals.push(strLit(s.text));
            } else if (t === 0x04) {
              vals.push(extra[ep] ? "TRUE" : "FALSE");
              ep += 8;
            } else if (t === 0x10) {
              vals.push(errorText(extra[ep]));
              ep += 8;
            } else {
              vals.push('""');
              ep += 8;
            }
          }
          out.push(vals.join(","));
        }
        st.push({ t: `{${out.join(";")}}`, p: ATOM });
        break;
      }
      case 0x21:
      case 0x22: {
        let argc: number;
        let tab: number;
        if (id === 0x21) {
          tab = u16(d, p);
          p += 2;
          const def = FTAB.get(tab);
          if (!def || def.argc < 0) throw new Unsupported("işlev");
          argc = def.argc;
        } else {
          argc = d[p] & 0x7f;
          tab = u16(d, p + 1);
          p += 3;
          if (tab & 0x8000) throw new Unsupported("komut işlevi");
        }
        const args: Item[] = [];
        for (let i = 0; i < argc; i++) args.unshift(pop());
        let name: string;
        if (tab === 0xff) {
          // Kullanıcı/eklenti işlevi: ilk bağımsız değişken işlev adıdır
          const fn = args.shift();
          name = (fn?.name ?? "").replace(/^_xlfn\.|^_xlws\./i, "").toUpperCase();
          if (!name) throw new Unsupported("adsız işlev");
        } else {
          const def = FTAB.get(tab);
          if (!def) throw new Unsupported("bilinmeyen işlev");
          name = def.name;
        }
        st.push({ t: `${name}(${args.map((a) => a.t).join(",")})`, p: ATOM });
        break;
      }
      case 0x23: {
        const idx = u32(d, p);
        p += 4;
        const n = ctx.names[idx - 1];
        if (!n || n.builtin) throw new Unsupported("ad");
        st.push({ t: n.name, p: ATOM, name: n.name });
        break;
      }
      case 0x24:
        st.push({ t: ref(u16(d, p), u16(d, p + 2), false), p: ATOM });
        p += 4;
        break;
      case 0x25:
        st.push({ t: area(u16(d, p), u16(d, p + 2), u16(d, p + 4), u16(d, p + 6), false), p: ATOM });
        p += 8;
        break;
      case 0x26:
        // Alt ifade izler; ek veri: Ref8U listesi
        p += 6;
        ep += 2 + u16(extra, ep) * 8;
        break;
      case 0x27:
      case 0x28:
        p += 6;
        break;
      case 0x29:
        p += 2;
        break;
      case 0x2a:
        p += 4;
        st.push({ t: "#REF!", p: ATOM });
        break;
      case 0x2b:
        p += 8;
        st.push({ t: "#REF!", p: ATOM });
        break;
      case 0x2c:
        st.push({ t: ref(u16(d, p), u16(d, p + 2), true), p: ATOM });
        p += 4;
        break;
      case 0x2d:
        st.push({ t: area(u16(d, p), u16(d, p + 2), u16(d, p + 4), u16(d, p + 6), true), p: ATOM });
        p += 8;
        break;
      case 0x39: {
        const ixti = u16(d, p);
        const idx = u32(d, p + 2);
        p += 6;
        const x = ctx.xti[ixti];
        const sb = x ? ctx.supbooks[x.sup] : undefined;
        const n = sb?.externNames[idx - 1];
        if (!sb || sb.kind === "external" || !n) throw new Unsupported("dış ad");
        st.push({ t: n, p: ATOM, name: n });
        break;
      }
      case 0x3a: {
        const pre = sheetPrefix(u16(d, p));
        st.push({ t: pre === null ? "#REF!" : pre + ref(u16(d, p + 2), u16(d, p + 4), false), p: ATOM });
        p += 6;
        break;
      }
      case 0x3b: {
        const pre = sheetPrefix(u16(d, p));
        st.push({ t: pre === null ? "#REF!" : pre + area(u16(d, p + 2), u16(d, p + 4), u16(d, p + 6), u16(d, p + 8), false), p: ATOM });
        p += 10;
        break;
      }
      case 0x3c:
        p += 6;
        st.push({ t: "#REF!", p: ATOM });
        break;
      case 0x3d:
        p += 10;
        st.push({ t: "#REF!", p: ATOM });
        break;
      default:
        throw new Unsupported(`belirteç ${ptg}`);
    }
  }
  if (st.length !== 1) return undefined;
  return st[0].t;
}

/** Paylaşılan formülün bağlantı hücresi: ptgExp (rgce yalnız 0x01 + satır + sütun). */
export function sharedRef(rgce: Uint8Array): { row: number; col: number } | null {
  if (rgce.length >= 5 && rgce[0] === 0x01) return { row: u16(rgce, 1), col: u16(rgce, 3) };
  return null;
}
