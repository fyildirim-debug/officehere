/**
 * LibreOffice çapraz denetimi: aynı veri ve formüller LibreOffice Calc'ta
 * hesaplatıldı (test/fixtures/lo-crosscheck.json); motorumuzun sonuçları
 * bunlarla karşılaştırılır. Bilinçli farklar (Excel davranışı esas alınır)
 * EXCEL_DIFF listesindedir.
 */
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { FormulaEngine, isError, type Scalar } from "../src";

interface Case {
  f: string;
  lo: { t: "n" | "s" | "e"; v: number | string; code?: number };
}
const fx = JSON.parse(readFileSync(new URL("./fixtures/lo-crosscheck.json", import.meta.url), "utf8")) as { data: (Scalar | null)[][]; cases: Case[] };

/** LibreOffice'in Excel'den ayrıldığı durumlar (motor Excel'e uyar). */
const DYN = "Dinamik dizi: Excel 365 sonucu taşırır (alttaki hücreler dolu olduğundan #TAŞMA!); LibreOffice örtük kesişim uygular";
const LOCALE = "Türkçe yerel ayar (ondalık virgül, binlik nokta, ₺, Türkçe büyük/küçük harf); LibreOffice İngilizce yerel ayarla çalıştı";
const DIRECT = "Excel doğrudan bağımsız değişken olarak verilen sayı metnini ve mantıksal değeri sayar";
const BINARY = "Sıralanmamış veride ikili arama sonucu (Excel algoritması); LibreOffice doğrusal arar";
const SECONDS = "Excel saat bileşenlerini en yakın saniyeye yuvarlayarak hesaplar";
const EXCEL_DIFF = new Map<string, string>([
  ["=FREQUENCY(A1:A10,{0,5,10})", DYN],
  ["=INDEX(A1:E10,0,1)", DYN],
  ["=A1:A3+1", DYN],
  ["=(A1:A2)*2", DYN],
  ["=TYPE(A1:A2)", "Excel çok hücreli başvuru için 64 (dizi) verir"],
  ["=TRUE>1", "Excel karşılaştırmasında mantıksal değer her sayıdan büyüktür"],
  ['="x"&1.5', LOCALE],
  ['=UPPER("istanbul")', LOCALE],
  ['=LOWER("IŞIK")', LOCALE],
  ['=TEXT(1234.567,"#,##0.00")', LOCALE],
  ["=FIXED(1234.567,1)", LOCALE],
  ["=DOLLAR(1234.5,1)", LOCALE],
  ['=ROUND("3.456",1)', LOCALE],
  ['=SUM("1",TRUE,1)', DIRECT],
  ['=MAX("5",3)', DIRECT],
  ["=DATE(1900,3,1)-DATE(1900,2,28)", "Excel 1900 artık yıl hatası (var sayılan 29.02.1900)"],
  ["=VLOOKUP(5,A1:E10,2,TRUE)", BINARY],
  ["=MATCH(9,E10:E1,-1)", BINARY],
  ["=MINUTE(0.0006944)", SECONDS],
  ["=HOUR(0.99999999)", SECONDS],
]);

/** Henüz desteklenmeyen işlevler (sonuç #NAME?). */
export const UNSUPPORTED = new Set<string>();

function engineWith(): { e: FormulaEngine; s: number } {
  const e = new FormulaEngine();
  const s = e.addSheet("Veri");
  fx.data.forEach((row, i) => row.forEach((v, j) => v !== null && e.setValue(s, i, j, v)));
  return { e, s };
}

function same(ours: Scalar, lo: Case["lo"]): boolean {
  if (lo.t === "e") return isError(ours) && ours.code === normalizeLoError(String(lo.v), lo.code);
  if (lo.t === "n") {
    const n = typeof ours === "boolean" ? (ours ? 1 : 0) : ours;
    if (typeof n !== "number") return false;
    const want = lo.v as number;
    return Math.abs(n - want) <= 1e-9 * Math.max(1, Math.abs(want));
  }
  return (ours === null ? "" : String(ours)) === lo.v;
}

function normalizeLoError(s: string, code?: number): string {
  if (s.startsWith("#")) return s;
  // LibreOffice "Err:5xx" kodları
  if (code === 504) return "#VALUE!";
  if (code === 502) return "#NUM!";
  return s;
}

describe("LibreOffice çapraz denetimi", () => {
  const { e, s } = engineWith();
  const results = fx.cases.map((c, i) => {
    e.setFormula(s, i, 6, c.f);
    return c;
  });
  e.recalc();
  const mismatches: string[] = [];
  results.forEach((c, i) => {
    const ours = e.getValue(s, i, 6);
    if (EXCEL_DIFF.has(c.f)) return;
    if (isError(ours) && ours.code === "#NAME?" && c.lo.v !== "#NAME?") {
      UNSUPPORTED.add(c.f);
      return;
    }
    if (!same(ours, c.lo)) mismatches.push(`${c.f}  motor=${isError(ours) ? ours.code : JSON.stringify(ours)}  LO=${JSON.stringify(c.lo.v)}`);
  });
  it(`${fx.cases.length} formül LibreOffice sonuçlarıyla eşleşir (bilinçli farklar: ${EXCEL_DIFF.size})`, () => {
    if (mismatches.length) console.log("UYUŞMAZLIKLAR:\n" + mismatches.join("\n"));
    expect(mismatches).toEqual([]);
    expect([...UNSUPPORTED]).toEqual([]);
  });
});
