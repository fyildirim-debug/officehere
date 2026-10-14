// Otomatik kurtarma taslakları: pencereye özgü anahtar, sahipsiz taslakların bulunması, devralma,
// eski yerel depo (localStorage) taslaklarının bir kez taşınması ve kota hatasının bildirilmesi.
// IndexedDB yerine aynı arayüzdeki bellek deposu kullanılır. Kaynak deposu: açılan .docx'in özgün paketi
// taslakla birlikte saklanır ve kurtarmada yeniden kaydedilir.
import { afterAll, beforeEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { getSource, openSource, releaseSource, sha256Hex } from "@officehere/core/opc";

class MemoryStorage {
  private m = new Map<string, string>();
  get length() {
    return this.m.size;
  }
  key(i: number) {
    return [...this.m.keys()][i] ?? null;
  }
  getItem(k: string) {
    return this.m.has(k) ? this.m.get(k)! : null;
  }
  setItem(k: string, v: string) {
    this.m.set(k, String(v));
  }
  removeItem(k: string) {
    this.m.delete(k);
  }
  clear() {
    this.m.clear();
  }
}

const store = new MemoryStorage();
(globalThis as unknown as { localStorage: MemoryStorage }).localStorage = store;

const D = await import("../drafts");
const content = (t: string) => ({ type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [{ type: "text", text: t }] }] }] });
const draft = (t: string, at = 1000) => ({ content: content(t), settings: {} as never, fileName: "Belge1", savedAt: at });
const stored = (key: string, t: string, at = 1000) => ({ ...draft(t, at), key: `${D.DRAFT_PREFIX}${key}`, owner: key });

/** Başka bir pencereyi taklit eder: ping'lere kendi kimliğiyle yanıt verir. */
function liveWindow(id: string) {
  const ch = new BroadcastChannel("officehere.belge.windows");
  ch.onmessage = (e: MessageEvent<{ t: string; nonce: string }>) => {
    if (e.data?.t === "ping") ch.postMessage({ t: "pong", from: id, nonce: e.data.nonce });
  };
  return () => ch.close();
}

let db: ReturnType<typeof D.memoryDraftBackend>;
beforeEach(async () => {
  store.clear();
  db = D.memoryDraftBackend();
  await D.setDraftBackendForTest(db);
});
afterAll(() => store.clear());

describe("pencereye özgü taslaklar", () => {
  it("taslak bu pencerenin anahtarına yazılır; başka pencerenin taslağını ezmez; yerel depoya yazılmaz", async () => {
    await db.put(stored("diger", "DİĞER"));
    expect(await D.writeDraft(draft("BENİM"))).toBe("ok");
    expect(db.records.get(D.draftKey())?.content).toEqual(content("BENİM"));
    expect(db.records.get(`${D.DRAFT_PREFIX}diger`)?.content).toEqual(content("DİĞER"));
    expect(store.length).toBe(0);
    expect((await D.readOwnDraft())?.content).toEqual(content("BENİM"));
    await D.clearOwnDraft();
    expect(db.records.has(D.draftKey())).toBe(false);
    expect(db.records.has(`${D.DRAFT_PREFIX}diger`)).toBe(true);
  });

  it("dosya yolu taslakla birlikte saklanır", async () => {
    await D.writeDraft({ ...draft("YOLLU"), path: "/ev/rapor.docx" });
    expect((await D.readOwnDraft())?.path).toBe("/ev/rapor.docx");
  });

  it("yazma ve silme sırası korunur (kaydedince silinen taslak geri gelmez)", async () => {
    const w = D.writeDraft(draft("SON"));
    const c = D.clearOwnDraft();
    await Promise.all([w, c]);
    expect(db.records.has(D.draftKey())).toBe(false);
  });

  it("yalnız sahipsiz taslaklar önerilir (canlı pencereninki ve bu pencereninki önerilmez)", async () => {
    await D.writeDraft(draft("BENİM", 3000));
    await db.put(stored("canli", "CANLI", 2000));
    await db.put(stored("kapali", "KAPALI", 1000));
    const stop = liveWindow("canli");
    const orphans = await D.orphanDrafts(150);
    stop();
    expect(orphans.map((d) => d.owner)).toEqual(["kapali"]);
  });

  it("nabzı taze pencere yanıt vermese de canlı sayılır", async () => {
    await db.put(stored("mesgul", "MEŞGUL"));
    store.setItem("officehere.belge.alive.mesgul", String(Date.now()));
    expect(await D.orphanDrafts(50)).toEqual([]);
    store.setItem("officehere.belge.alive.mesgul", String(Date.now() - 60_000));
    expect((await D.orphanDrafts(50)).map((d) => d.owner)).toEqual(["mesgul"]);
  });

  it("devralma: önce kendi anahtarına yazar, sonra eskisini siler", async () => {
    await db.put(stored("kapali", "KURTAR"));
    const [d] = await D.listDrafts();
    await D.adoptDraft(d);
    expect(db.records.has(`${D.DRAFT_PREFIX}kapali`)).toBe(false);
    expect((await D.readOwnDraft())?.content).toEqual(content("KURTAR"));
    expect((await D.readOwnDraft())?.savedAt).toBe(1000);
  });

  it("kullanıcı isteğiyle kapanan pencere artık taslak yazmaz", async () => {
    await D.writeDraft(draft("İLK"));
    await D.forgetOwnDraft();
    expect(await D.writeDraft(draft("SONRA"))).toBe("closing");
    expect(db.records.has(D.draftKey())).toBe(false);
  });

  it("önizleme metni", () => {
    expect(D.draftPreview({ content: content("Merhaba dünya") as never })).toBe("Merhaba dünya");
    expect(D.draftPreview({ content: content("x".repeat(80)) as never }, 10)).toBe(`${"x".repeat(10)}…`);
  });
});

describe("IndexedDB'ye taşıma ve kota", () => {
  it("eski yerel depo taslakları bir kez taşınır ve yerel depodan silinir", async () => {
    store.setItem(`${D.DRAFT_PREFIX}eskipencere`, JSON.stringify({ ...draft("ESKİ PENCERE", 500), owner: "eskipencere" }));
    store.setItem(D.LEGACY_DRAFT_KEY, JSON.stringify(draft("ÇOK ESKİ", 400)));
    store.setItem("officehere.baska.anahtar", "dokunma");
    const fresh = D.memoryDraftBackend();
    expect(await D.setDraftBackendForTest(fresh)).toBe(2);
    expect(store.getItem(`${D.DRAFT_PREFIX}eskipencere`)).toBeNull();
    expect(store.getItem(D.LEGACY_DRAFT_KEY)).toBeNull();
    expect(store.getItem("officehere.baska.anahtar")).toBe("dokunma");
    const all = await D.listDrafts();
    expect(all.map((d) => d.fileName && D.draftPreview(d))).toEqual(["ESKİ PENCERE", "ÇOK ESKİ"]);
    // Sahipsiz sayılırlar (eski pencerenin nabzı yok, eski tek anahtarın sahibi yok).
    expect((await D.orphanDrafts(50)).length).toBe(2);
    // İkinci taşımada yapılacak iş kalmaz.
    expect(await D.setDraftBackendForTest(fresh)).toBe(0);
  });

  it("taşıma başarısızsa yerel kopya silinmez (taslak kaybolmaz)", async () => {
    store.setItem(`${D.DRAFT_PREFIX}buyuk`, JSON.stringify({ ...draft("x".repeat(5000)), owner: "buyuk" }));
    await D.setDraftBackendForTest(D.memoryDraftBackend({ quotaBytes: 100 }));
    expect(store.getItem(`${D.DRAFT_PREFIX}buyuk`)).not.toBeNull();
  });

  it("kota dolarsa yazma 'quota' döner (sessizce yutulmaz), önceki taslak korunur", async () => {
    const small = D.memoryDraftBackend({ quotaBytes: 2000 });
    await D.setDraftBackendForTest(small);
    expect(await D.writeDraft(draft("küçük"))).toBe("ok");
    expect(await D.writeDraft(draft("R".repeat(10_000)))).toBe("quota");
    expect((await D.readOwnDraft())?.content).toEqual(content("küçük"));
  });

  it("resimli büyük belge (yerel deponun 5 MB sınırını aşan) taslağı yazılır", async () => {
    const image = `data:image/png;base64,${"A".repeat(6 * 1024 * 1024)}`;
    const big = { ...draft("RESİMLİ"), content: { type: "doc", content: [{ type: "paragraph", content: [{ type: "image", attrs: { src: image } }] }] } };
    expect(await D.writeDraft(big as never)).toBe("ok");
    expect(JSON.stringify((await D.readOwnDraft())?.content).length).toBeGreaterThan(6 * 1024 * 1024);
  });
});

describe("kaynak deposu (geçiş katmanı: kurtarmada özgün paket)", () => {
  /** Kaynak kaydına girmiş küçük bir paket (içeriği önemsiz, ZIP olması yeter). */
  async function registered(text = "özgün", extra?: Uint8Array) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
    zip.file("veri.txt", text);
    if (extra) zip.file("buyuk.bin", extra);
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "STORE" });
    const id = await sha256Hex(bytes);
    await openSource(bytes, { id });
    return { id, bytes };
  }
  const withSource = (t: string, id: string) => ({ ...draft(t), settings: { kaynak: id } as never });

  it("paket, onu anan ilk taslakla bir kez yazılır; taslak silinince paket de silinir", async () => {
    const { id, bytes } = await registered();
    expect(await D.writeDraft(withSource("BİR", id))).toBe("ok");
    expect(await D.writeDraft(withSource("İKİ", id))).toBe("ok");
    expect(db.sources.size).toBe(1);
    expect([...db.sources.get(id)!.bytes]).toEqual([...bytes]);
    expect(db.records.get(D.draftKey())?.kaynak).toBe(id);
    await D.clearOwnDraft();
    expect(db.sources.size).toBe(0);
  });

  it("başka bir taslak andıkça paket silinmez", async () => {
    const { id } = await registered("paylaşılan");
    await D.writeDraft(withSource("BENİM", id));
    await db.put({ ...stored("diger", "DİĞER"), kaynak: id });
    await D.clearOwnDraft();
    expect(db.sources.has(id)).toBe(true);
  });

  it("kurtarma: bellekte olmayan paket depodan aynı kimlikle açılır; bulunamazsa 'yok'", async () => {
    const { id } = await registered("kurtarma");
    await D.writeDraft(withSource("KURTAR", id));
    releaseSource(id);
    expect(getSource(id)).toBeUndefined();
    expect(await D.ensureSource(id)).toBe("depo");
    expect(getSource(id)).toBeTruthy();
    expect(await D.ensureSource(id)).toBe("bellek");
    expect(await D.ensureSource("0".repeat(64))).toBe("yok");
    expect(await D.ensureSource(null)).toBeNull();
  });

  it("içerik kimliğiyle eşleşmeyen (bozulmuş) paket kullanılmaz", async () => {
    const { id } = await registered("bozulacak");
    await D.writeDraft(withSource("BOZUK", id));
    releaseSource(id);
    db.sources.get(id)!.bytes[0] ^= 0xff;
    expect(await D.ensureSource(id)).toBe("yok");
  });

  it("paket kotaya sığmazsa taslak paketsiz yazılır (taslak kaybolmaz)", async () => {
    const { id } = await registered("büyük", crypto.getRandomValues(new Uint8Array(20_000)));
    const small = D.memoryDraftBackend({ quotaBytes: 5_000 });
    await D.setDraftBackendForTest(small);
    expect(await D.writeDraft(withSource("KÜÇÜK TASLAK", id))).toBe("ok");
    expect(small.sources.size).toBe(0);
    expect(small.records.get(D.draftKey())?.kaynak).toBe(id);
  });
});
