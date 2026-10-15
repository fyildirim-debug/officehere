// @vitest-environment jsdom
// Büyük belgenin parça parça yüklenmesi (ilk ekran önce): sonuç, tek seferde yüklemeyle aynı belgedir; geçmişe
// girmez, kirli saymaz, değişiklik izlemede kayıt üretmez, salt okunur korumada da tamamlanır; kaydetme ya da
// klavye girdisi kalanı hemen ekler.
import { Editor, type JSONContent } from "@tiptap/core";
import { undoDepth } from "@tiptap/pm/history";
import { editorExtensions } from "../../editor/extensions";
import { PaginationController } from "../../editor/pagination";
import { protectionSource } from "../../editor/protection";
import { setTracking } from "../../editor/trackChanges";
import { emptyDoc } from "../../model/document";
import { resolveStyles } from "../../model/styles";
import { PROGRESSIVE_MIN, loadIntoEditor, loadProgressive } from "../useFileOps";

// jsdom'da yerleşim yok: sayfalama eklentisinin gözlemcisi boş.
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

const hooks = { styles: () => resolveStyles("default", "OfficeHere"), openDialog: () => {}, toggleShowMarks: () => {}, newComment: () => {}, toggleTrack: () => {}, insertNote: () => {} };
const editors: Editor[] = [];
function make(): Editor {
  const e = new Editor({ extensions: editorExtensions(new PaginationController(), hooks), content: emptyDoc() });
  editors.push(e);
  return e;
}
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy();
  protectionSource.get = () => null;
});

const para = (t: string): JSONContent => ({ type: "paragraph", content: [{ type: "text", text: t }] });
const cell = (t: string): JSONContent => ({ type: "tableCell", content: [para(t)] });
const table = (n: number): JSONContent => ({ type: "table", content: [{ type: "tableRow", content: [cell(`a${n}`), cell(`b${n}`)] }] });

/** Üç bölümlü büyük belge: başlıklar, listeler, bölüm sonunda tablo (yükleyici araya paragraf sokmamalı), alanlar. */
function bigDoc(): JSONContent {
  const sec = (k: number, n: number): JSONContent => {
    const content: JSONContent[] = [];
    for (let i = 0; i < n; i++) {
      if (i % 40 === 0) content.push({ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: `Bölüm ${k} başlık ${i}` }] });
      if (i % 25 === 7) content.push(table(i));
      if (i % 30 === 3) content.push({ type: "bulletList", content: [{ type: "listItem", content: [para(`madde ${i}`)] }] });
      content.push(para(`${k}.${i} ${"Uzun bir paragraf metni ığüşöç İĞÜŞÖÇ. ".repeat(3)}`));
    }
    content.push(table(9999));
    return { type: "section", content };
  };
  return { type: "doc", content: [sec(1, 400), sec(2, 900), sec(3, 300)] };
}

async function loaded(e: Editor, json: JSONContent): Promise<{ job: ReturnType<typeof loadProgressive> }> {
  const job = loadProgressive(e, json);
  expect(job).not.toBeNull();
  expect(await job!.done).toBe(true);
  return { job };
}

describe("parça parça yükleme", () => {
  it("önceki belgede değişiklik izleme ya da salt okunur koruma varken yeni belge eksiksiz ve izlenmeden yüklenir", () => {
    const e = make();
    setTracking(e.view, true, "Deneme");
    loadIntoEditor(e, { type: "doc", content: [para("yeni belge")] });
    expect(JSON.stringify(e.getJSON())).not.toContain("insertion");
    protectionSource.get = () => ({ enforced: true, mode: "readOnly" }) as ReturnType<typeof protectionSource.get>;
    loadIntoEditor(e, { type: "doc", content: [para("üçüncü belge")] });
    expect(e.state.doc.textContent).toBe("üçüncü belge");
  });

  it("küçük belgede devreye girmez", () => {
    expect(loadProgressive(make(), { type: "doc", content: [para("kısa")] })).toBeNull();
  });

  it("tek seferde yüklemeyle aynı belgeyi kurar; geçmiş boş, güncelleme olayı yok", async () => {
    const json = bigDoc();
    expect(JSON.stringify(json).length).toBeGreaterThan(PROGRESSIVE_MIN);
    const a = make();
    loadIntoEditor(a, json);
    const b = make();
    let updates = 0;
    b.on("update", () => updates++);
    const first = loadProgressive(b, json)!;
    // İlk parça hemen: belgenin başı görünür, sonu henüz yok.
    expect(b.state.doc.content.size).toBeLessThan(a.state.doc.content.size);
    expect(b.state.doc.textContent.startsWith(a.state.doc.textContent.slice(0, 200))).toBe(true);
    expect(await first.done).toBe(true);
    expect(b.getJSON()).toEqual(a.getJSON());
    expect(undoDepth(b.state)).toBe(0);
    expect(updates).toBe(0);
    expect(b.state.selection.from).toBe(a.state.selection.from);
  });

  it("değişiklik izleme açıkken ve salt okunur korumada da aynı belge", async () => {
    const json = bigDoc();
    const a = make();
    loadIntoEditor(a, json);
    const b = make();
    setTracking(b.view, true, "Deneme");
    protectionSource.get = () => ({ enforced: true, mode: "readOnly" }) as ReturnType<typeof protectionSource.get>;
    await loaded(b, json);
    protectionSource.get = () => null;
    expect(JSON.stringify(b.getJSON())).not.toContain("insertion");
    expect(b.getJSON()).toEqual(a.getJSON());
  });

  it("flush kalanı hemen ekler; klavye girdisi de", async () => {
    const json = bigDoc();
    const a = make();
    loadIntoEditor(a, json);
    const b = make();
    const job = loadProgressive(b, json)!;
    job.flush();
    expect(await job.done).toBe(true);
    expect(b.getJSON()).toEqual(a.getJSON());
    const c = make();
    const job2 = loadProgressive(c, json)!;
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "a" }));
    expect(await job2.done).toBe(true);
    expect(c.getJSON()).toEqual(a.getJSON());
  });

  it("durdurulunca eklenenler kalır, sonuç false", async () => {
    const b = make();
    const job = loadProgressive(b, bigDoc())!;
    const size = b.state.doc.content.size;
    job.stop();
    expect(await job.done).toBe(false);
    await new Promise((r) => setTimeout(r, 150));
    expect(b.state.doc.content.size).toBe(size);
  });
});
