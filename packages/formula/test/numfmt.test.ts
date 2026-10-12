import { describe, expect, it } from "vitest";
import { formatNumber, formatValue, formatGeneral, formatFromLocal, formatToLocal, isDateFormat, numberToText } from "../src/numfmt";
import { dateToSerial } from "../src/dates";

const f = (v: number, code: string) => formatNumber(v, code).text;

describe("sayı biçimleri", () => {
  it("General", () => {
    expect(formatGeneral(0)).toBe("0");
    expect(formatGeneral(1234)).toBe("1234");
    expect(formatGeneral(1.5)).toBe("1,5");
    expect(formatGeneral(1 / 3)).toBe("0,333333333");
    expect(formatGeneral(2 / 3)).toBe("0,666666667");
    expect(formatGeneral(-1.25)).toBe("-1,25");
    expect(formatGeneral(123456789012)).toBe("1,23457E+11");
    expect(formatGeneral(0.1 + 0.2)).toBe("0,3");
    expect(formatGeneral(1e-10)).toBe("1E-10");
  });
  it("sabit ondalık ve binlik", () => {
    expect(f(1234.567, "0.00")).toBe("1234,57");
    expect(f(1234.567, "#,##0.00")).toBe("1.234,57");
    expect(f(1234567.891, "#,##0")).toBe("1.234.568");
    expect(f(0.5, "0")).toBe("1");
    expect(f(0, "#,##0.00")).toBe("0,00");
    expect(f(-1234.5, "#,##0.00")).toBe("-1.234,50");
    expect(f(5, "000")).toBe("005");
    expect(f(0.123, "#.##")).toBe(",12");
    expect(f(1.5, "0.0#")).toBe("1,5");
    expect(f(1.555, "0.0#")).toBe("1,56");
    expect(f(1000000, "#,##0,")).toBe("1.000");
    expect(f(1500000, "0.0,,")).toBe("1,5");
  });
  it("yüzde ve bilimsel", () => {
    expect(f(0.25, "0%")).toBe("25%");
    expect(f(0.1234, "0.00%")).toBe("12,34%");
    expect(f(12345, "0.00E+00")).toBe("1,23E+04");
    expect(f(0.00012, "0.00E+00")).toBe("1,20E-04");
  });
  it("bölümler ve renkler", () => {
    expect(f(-5, "0;(0)")).toBe("(5)");
    expect(f(0, "0;-0;\"sıfır\"")).toBe("sıfır");
    const r = formatNumber(-3, "#,##0;[Red]-#,##0");
    expect(r.text).toBe("-3");
    expect(r.color).toBe("#FF0000");
    expect(formatNumber(-3, "0;[Kırmızı]0").color).toBe("#FF0000");
    expect(f(150, "[>100]\"büyük\";\"küçük\"")).toBe("büyük");
  });
  it("para birimi", () => {
    expect(f(1234.5, '#,##0.00 "₺"')).toBe("1.234,50 ₺");
    expect(f(1234.5, "[$₺-41F]#,##0.00")).toBe("₺1.234,50");
    expect(f(-1234.5, '_-* #,##0.00 "₺"_-;-* #,##0.00 "₺"_-;_-* "-"?? "₺"_-;_-@_-')).toBe("-1.234,50 ₺ ");
  });
  it("kesirler", () => {
    expect(f(1.5, "# ?/?")).toBe("1 1/2");
    expect(f(0.75, "?/?")).toBe("3/4");
    expect(f(2.333, "# ??/??")).toBe("2  1/3 ");
  });
  it("tarih ve saat", () => {
    const s = dateToSerial(2024, 3, 5);
    expect(f(s, "dd.mm.yyyy")).toBe("05.03.2024");
    expect(f(s, "d mmmm yyyy dddd")).toBe("5 Mart 2024 Salı");
    expect(f(s, "mmm yy")).toBe("Mar 24");
    expect(f(s + 0.5 + 1 / 1440 * 5, "hh:mm")).toBe("12:05");
    expect(f(s + 0.75, "h:mm AM/PM")).toBe("6:00 ÖS");
    expect(f(1.5, "[h]:mm")).toBe("36:00");
    expect(f(0.000011574, "hh:mm:ss")).toBe("00:00:01");
    expect(f(60, "dd.mm.yyyy")).toBe("29.02.1900");
    expect(f(61, "dd.mm.yyyy")).toBe("01.03.1900");
    expect(f(1, "dd.mm.yyyy")).toBe("01.01.1900");
    expect(isDateFormat("dd.mm.yyyy")).toBe(true);
    expect(isDateFormat("#,##0")).toBe(false);
    expect(f(s, "yyyy-mm-dd hh:mm")).toBe("2024-03-05 00:00");
  });
  it("metin bölümü", () => {
    expect(formatValue("abc", "@").text).toBe("abc");
    expect(formatValue("x", '0;0;0;"<"@">"').text).toBe("<x>");
    expect(formatValue(true, "General").text).toBe("DOĞRU");
  });
  it("Türkçe kod dönüşümü", () => {
    expect(formatFromLocal("#.##0,00")).toBe("#,##0.00");
    expect(formatFromLocal("gg.aa.yyyy")).toBe("dd.mm.yyyy");
    expect(formatFromLocal("[Kırmızı]0")).toBe("[Red]0");
    expect(formatFromLocal("Genel")).toBe("General");
    expect(formatFromLocal("ss:dd:nn")).toBe("hh:mm:ss");
    expect(formatToLocal("#,##0.00")).toBe("#.##0,00");
    expect(formatToLocal("dd.mm.yyyy")).toBe("gg.aa.yyyy");
    expect(formatToLocal("hh:mm")).toBe("ss:dd");
    expect(formatToLocal("[Red]0.0")).toBe("[Kırmızı]0,0");
    expect(formatToLocal("General")).toBe("Genel");
  });
  it("numberToText", () => {
    expect(numberToText(1 / 3)).toBe("0,333333333333333");
    expect(numberToText(1.5, ".")).toBe("1.5");
    expect(numberToText(1e21)).toBe("1E+21");
    expect(numberToText(0.1 + 0.2)).toBe("0,3");
  });
});
