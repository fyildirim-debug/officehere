// @vitest-environment jsdom
// Worker'daki okuyucu (yerleşik XML ayrıştırıcısı + xmldom DOM) ile süreç içi okuyucu (DOMParser) derlemdeki her
// belgede aynı modeli, uyarıları ve hatayı üretir. Ek dizinler: OH_ESITLIK_DIZIN="dizin1,dizin2".
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

const KOK = resolve(__dirname, "../../../../..");
const DIZINLER = [
  "apps/belge/src/io/__fixtures__",
  "apps/belge/src/io/__fixtures__/bicimler",
  "test-derlem/poi",
  "test-derlem/lo",
  "test-derlem/py",
  ...(process.env.OH_ESITLIK_DIZIN ?? "").split(",").filter(Boolean),
];
// HTML ve Markdown ana iş parçacığında açılır (DOM gerekir).
const UZANTI = /\.(docx|docm|dotx|dotm|doc|odt|ott|fodt|rtf|txt)$/i;

const dosyalar = DIZINLER.map((d) => resolve(KOK, d))
  .filter((d) => existsSync(d))
  .flatMap((d) => readdirSync(d).map((f) => join(d, f)))
  .filter((f) => UZANTI.test(f) && statSync(f).isFile());

type Sonuc = { doc?: unknown; format?: string; notes?: string[]; warnings?: string[]; details?: string[]; hata?: string; main?: boolean };

/** Opak öğe XML'inde sayısal karakter başvurusu yazımı (&#9; / &#x9;) seri hale getiriciye göre değişir. */
function olagan(v: unknown): unknown {
  return JSON.parse(JSON.stringify(v), (_k, x: unknown) => (typeof x === "string" ? x.replace(/&#x([0-9A-Fa-f]+);/g, (_m, h: string) => `&#${parseInt(h, 16)};`) : x));
}

/** İlk farkın yolu ve değerleri (hata iletisi için). */
function ilkFark(a: unknown, b: unknown, yol = ""): string {
  if (Object.is(a, b)) return "";
  if (a && b && typeof a === "object" && typeof b === "object") {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
      const d = ilkFark((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k], `${yol}.${k}`);
      if (d) return d;
    }
    return "";
  }
  return `${yol}: ${JSON.stringify(a)?.slice(0, 300)} ≠ ${JSON.stringify(b)?.slice(0, 300)}`;
}

async function oku(kip: "dom" | "yerlesik"): Promise<Map<string, Sonuc>> {
  vi.resetModules();
  const g = globalThis as { XMLSerializer?: unknown };
  const domSerializer = window.XMLSerializer;
  if (kip === "yerlesik") (await import("../xmlDom")).installXmlDom({ force: true });
  else g.XMLSerializer = domSerializer;
  const { readDocument, DomRequired } = await import("../index");
  const out = new Map<string, Sonuc>();
  // Tarihsiz öğelere okuma anı yazılır: iki kipte aynı an.
  vi.useFakeTimers({ toFake: ["Date"], now: new Date("2026-01-02T03:04:05Z") });
  try {
    for (const f of dosyalar) {
      const b = readFileSync(f);
      const ab = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
      try {
        const r = await readDocument(f, ab, { noDom: true });
        out.set(f, { doc: olagan(r.doc), format: r.format, notes: r.notes, warnings: r.warnings.messages(), details: r.warnings.details.filter((d) => !/XML|satır|line|parse/i.test(d)) });
      } catch (err) {
        if (err instanceof DomRequired) out.set(f, { main: true });
        else out.set(f, { hata: `${(err as Error).name}: ${(err as Error).message}` });
      }
    }
  } finally {
    vi.useRealTimers();
    (await import("../docx/xml")).setBuiltinXml(null);
    g.XMLSerializer = domSerializer;
  }
  return out;
}

describe("Worker okuyucusu süreç içi okuyucuyla eşit", () => {
  it(`${dosyalar.length} dosyada model, uyarı ve hata aynı`, async () => {
    expect(dosyalar.length).toBeGreaterThan(20);
    const dom = await oku("dom");
    const yerlesik = await oku("yerlesik");
    const farkli: string[] = [];
    for (const f of dosyalar) {
      const a = dom.get(f);
      const b = yerlesik.get(f);
      try {
        expect(b).toEqual(a);
      } catch {
        farkli.push(`${f.slice(KOK.length + 1)}: ${ilkFark(a, b)}`);
      }
    }
    expect(farkli).toEqual([]);
  }, 300_000);
});
