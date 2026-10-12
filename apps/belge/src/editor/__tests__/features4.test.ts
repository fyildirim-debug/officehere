// @vitest-environment jsdom
// 4. tur yeni özellikler (saf mantık): satır numaralandırma, karşılaştırma, dizin, grafik ölçeği,
// kırpma boyutu, koruma (parola karması ve izin süzgeci).
import { afterEach, describe, expect, it } from "vitest";
import { Editor, type JSONContent } from "@tiptap/core";
import { normalizeDoc } from "../../model/document";
import { schemaExtensions } from "../extensions";
import { numberLines, type LineBox } from "../lineNumbers";
import { compareDocuments, similarity } from "../compare";
import { collectIndex, indexUpdate } from "../indexing";
import { niceScale, chartSvg, defaultChart } from "../chart";
import { croppedSize, cropCss } from "../image";
import { allowedUnder, checkPassword, hashPassword, makeProtection } from "../protection";

let editor: Editor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
});
const make = (content: object) => (editor = new Editor({ extensions: schemaExtensions(), content: normalizeDoc(content as JSONContent) }));
const p = (text: string, extra: JSONContent[] = []): JSONContent => ({ type: "paragraph", content: [...(text ? [{ type: "text", text }] : []), ...extra] });
const doc = (...c: JSONContent[]): JSONContent => ({ type: "doc", content: [{ type: "section", content: c }] });

describe("satır numaraları", () => {
  const L = (page: number, section = 0, suppress = false): LineBox => ({ top: 0, height: 10, left: 0, page, section, suppress });
  it("sürekli, sayfada yeniden başlat, aralık ve gizlenen paragraf", () => {
    const lines = [L(0), L(0), L(0, 0, true), L(1), L(1)];
    expect(numberLines(lines, () => ({ restart: "continuous" })).map((x) => x.n)).toEqual([1, 2, 3, 4]);
    expect(numberLines(lines, () => ({ restart: "newPage" })).map((x) => x.n)).toEqual([1, 2, 1, 2]);
    expect(numberLines(lines, () => ({ restart: "continuous", countBy: 2 })).map((x) => x.n)).toEqual([2, 4]);
    expect(numberLines(lines, () => ({ restart: "continuous", start: 10 })).map((x) => x.n)).toEqual([10, 11, 12, 13]);
  });
  it("bölümde yeniden başlat; numarasız bölüm sayılmaz", () => {
    const lines = [L(0, 0), L(0, 0), L(1, 1), L(1, 2), L(2, 2)];
    const cfg = (s: number) => (s === 1 ? null : { restart: "newSection" as const });
    expect(numberLines(lines, cfg).map((x) => [x.line.section, x.n])).toEqual([
      [0, 1],
      [0, 2],
      [2, 1],
      [2, 2],
    ]);
  });
});

describe("Karşılaştır", () => {
  const marksOf = (d: JSONContent, type: string): string[] => {
    const out: string[] = [];
    const walk = (n: JSONContent) => {
      if (n.type === "text" && n.marks?.some((m) => m.type === type)) out.push(n.text ?? "");
      n.content?.forEach(walk);
    };
    walk(d);
    return out;
  };
  it("sözcük düzeyinde ekleme/silme, eklenen ve silinen paragraflar", () => {
    const a = doc(p("Birinci paragraf aynı kalır."), p("Kırmızı elma ağaçta durur."), p("Silinecek paragraf."));
    const b = doc(p("Birinci paragraf aynı kalır."), p("Yeşil elma dalda durur."), p("Yeni eklenen paragraf burada."));
    const r = compareDocuments(a, b, "Ayşe", "2026-01-01T00:00:00Z");
    expect(marksOf(r, "deletion").join("|")).toContain("Kırmızı");
    expect(marksOf(r, "insertion").join("|")).toContain("Yeşil");
    expect(marksOf(r, "deletion").join(" ")).toContain("Silinecek");
    expect(marksOf(r, "insertion").join(" ")).toContain("Yeni");
    const first = r.content![0].content![0];
    expect(first.content![0].marks).toBeUndefined();
    const ins = JSON.stringify(r).match(/"author":"Ayşe"/g) ?? [];
    expect(ins.length).toBeGreaterThan(2);
    // Sonuç belgesi geçerli şemaya uyar.
    expect(() => make(r)).not.toThrow();
  });
  it("benzerlik ölçüsü", () => {
    expect(similarity("a b c", "a b c")).toBe(1);
    expect(similarity("a b c d", "a b x d")).toBeCloseTo(0.75);
    expect(similarity("", "x")).toBe(0);
  });
});

describe("Dizin", () => {
  it("girdiler Türkçe sırayla, alt girdiler ve sayfa numaraları", () => {
    const xe = (main: string, sub?: string): JSONContent => ({ type: "indexEntry", attrs: { main, sub: sub ?? null } });
    const ed = make(doc(p("Şehir ", [xe("Şanlıurfa")]), p("Başkent ", [xe("Ankara"), xe("Şehirler", "İzmir")]), p("Yine ", [xe("Ankara")]), { type: "indexBlock", attrs: { columns: 2, entries: [] } }));
    let yine = 0;
    ed.state.doc.descendants((n, pos) => (n.isTextblock && n.textContent.startsWith("Yine") ? ((yine = pos), false) : true));
    const pageOf = (pos: number) => (pos >= yine ? 1 : 0);
    const items = collectIndex(ed.state.doc, pageOf);
    expect(items.map((e) => `${e.main}${e.sub ? `>${e.sub}` : ""}:${e.pages.join(",")}`)).toEqual(["Ankara:1,2", "Şanlıurfa:1", "Şehirler:", "Şehirler>İzmir:1"]);
    const tr = indexUpdate(ed.state, pageOf)!;
    ed.view.dispatch(tr);
    let entries: unknown = null;
    ed.state.doc.descendants((n) => (n.type.name === "indexBlock" ? ((entries = n.attrs.entries), false) : true));
    expect(entries).toEqual(items);
    expect(indexUpdate(ed.state, pageOf)).toBeNull();
  });
});

describe("Grafik, kırpma", () => {
  it("eksen ölçeği yuvarlak adımlarla", () => {
    expect(niceScale(0, 4.5)).toEqual({ lo: 0, hi: 5, step: 1 });
    expect(niceScale(-3, 7).lo).toBeLessThanOrEqual(-3);
  });
  it("SVG grafik türleri", () => {
    for (const t of ["column", "bar", "line", "pie"] as const) {
      const svg = chartSvg(defaultChart(t, ["#112233", "#445566", "#778899"]));
      expect(svg).toMatch(/^<svg/);
      expect(svg).toContain("Grafik Başlığı");
    }
    expect(chartSvg(defaultChart("pie", ["#112233"]))).toContain("<path");
    expect(chartSvg(defaultChart("line", ["#112233"]))).toContain("<polyline");
  });
  it("kırpma ölçeği korur", () => {
    expect(croppedSize(200, 100, null, { l: 0.25, t: 0, r: 0.25, b: 0.5 })).toEqual({ width: 100, height: 50 });
    expect(croppedSize(100, 50, { l: 0.25, t: 0, r: 0.25, b: 0.5 }, null)).toEqual({ width: 200, height: 100 });
    expect(cropCss({ l: 0.1, t: 0, r: 0, b: 0.2 })).toBe("inset(0% 0% 20% 10%)");
    expect(cropCss(null)).toBeNull();
  });
});

describe("Düzenlemeyi Kısıtla", () => {
  it("parola karması (SHA-512, tuz, yineleme) ve doğrulama", async () => {
    const salt = btoa("0123456789abcdef");
    const h1 = await hashPassword("gizli", salt, 10);
    expect(h1).toBe(await hashPassword("gizli", salt, 10));
    expect(h1).not.toBe(await hashPassword("Gizli", salt, 10));
    const prot = await makeProtection("comments", "parola");
    expect(prot).toMatchObject({ mode: "comments", enforced: true, algorithm: "SHA-512", spinCount: 100000 });
    expect(await checkPassword(prot, "parola")).toBe(true);
    expect(await checkPassword(prot, "yanlış")).toBe(false);
    expect(await checkPassword({ mode: "readOnly", enforced: true }, "")).toBe(true);
  }, 120_000);
  it("izin süzgeci: salt okunur her değişikliği, açıklama kipi yalnız açıklama işaretini geçirir", () => {
    const ed = make(doc(p("korunan metin")));
    const typing = ed.state.tr.insertText("x", 2);
    const comment = ed.state.tr.addMark(2, 6, ed.state.schema.marks.comment.create({ id: "c1" }));
    const bold = ed.state.tr.addMark(2, 6, ed.state.schema.marks.bold.create());
    const ro = { mode: "readOnly" as const, enforced: true };
    const cm = { mode: "comments" as const, enforced: true };
    expect(allowedUnder(ro, typing)).toBe(false);
    expect(allowedUnder(ro, comment)).toBe(false);
    expect(allowedUnder(cm, comment)).toBe(true);
    expect(allowedUnder(cm, bold)).toBe(false);
    expect(allowedUnder({ ...cm, enforced: false }, typing)).toBe(true);
    expect(allowedUnder(cm, ed.state.tr.setMeta("x", 1))).toBe(true);
  });
});
