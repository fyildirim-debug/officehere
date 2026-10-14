import { describe, expect, it } from "vitest";
import { crc32, inflateEntry, isZip, OpcError, rawData, readZip, ZipWriter } from "../../src/opc/zip";

const enc = new TextEncoder();
const dec = new TextDecoder();

async function arsiv(parts: Record<string, string | Uint8Array>, compress = true): Promise<Uint8Array> {
  const z = new ZipWriter();
  for (const [n, d] of Object.entries(parts)) z.add(n, d, { compress });
  return z.finish();
}

describe("ZIP okuyucu ve yazıcı", () => {
  it("CRC-32 bilinen değer", () => {
    expect(crc32(enc.encode("123456789"))).toBe(0xcbf43926);
    expect(crc32(new Uint8Array())).toBe(0);
  });

  it("yazılan arşiv okunur: adlar sırayla, içerik ve CRC doğru (sıkıştırmalı ve sıkıştırmasız)", async () => {
    const uzun = "merhaba dünya ".repeat(500);
    for (const compress of [true, false]) {
      const b = await arsiv({ "[Content_Types].xml": "<Types/>", "word/document.xml": uzun, "ğüşİ/çöp.bin": new Uint8Array([1, 2, 3]) }, compress);
      expect(isZip(b)).toBe(true);
      const z = readZip(b);
      expect(z.entries.map((e) => e.name)).toEqual(["[Content_Types].xml", "word/document.xml", "ğüşİ/çöp.bin"]);
      expect(z.skipped).toBe(0);
      const doc = z.entries[1];
      expect(doc.method).toBe(compress ? 8 : 0);
      const data = await inflateEntry(b, doc, 1 << 20);
      expect(dec.decode(data)).toBe(uzun);
      expect(crc32(data)).toBe(doc.crc32);
      expect(doc.size).toBe(enc.encode(uzun).length);
      expect(Array.from(await inflateEntry(b, z.entries[2], 100))).toEqual([1, 2, 3]);
    }
  });

  it("ham kopya yeniden sıkıştırmaz: sıkıştırılmış baytlar aynen taşınır", async () => {
    const a = await arsiv({ "media/resim.bin": new Uint8Array(4096).map((_, i) => i % 7) });
    const za = readZip(a);
    const w = new ZipWriter();
    w.add("ilk.xml", "<a/>");
    w.addRaw("kopya/resim.bin", a, za.entries[0]);
    const b = await w.finish();
    const zb = readZip(b);
    const e = zb.entries[1];
    expect(e.name).toBe("kopya/resim.bin");
    expect(Array.from(rawData(b, e))).toEqual(Array.from(rawData(a, za.entries[0])));
    expect(e.crc32).toBe(za.entries[0].crc32);
    expect(Array.from(await inflateEntry(b, e, 1 << 20))).toEqual(Array.from(await inflateEntry(a, za.entries[0], 1 << 20)));
  });

  it("aynı ad (büyük/küçük harf duyarsız) yazıcıda yerini alır, okuyucuda ilk girdi kalır", async () => {
    const w = new ZipWriter();
    w.add("a.xml", "1");
    w.add("b.xml", "2");
    w.add("A.XML", "3");
    const z = readZip(await w.finish());
    expect(z.entries.map((e) => e.name)).toEqual(["A.XML", "b.xml"]);
    // Elle yinelenen girdi: merkez dizine aynı adı iki kez yazan arşiv
    const b = await arsiv({ "x.xml": "1", "y.xml": "2" });
    const bozuk = b.slice();
    // merkez dizindeki (son geçen) "y.xml" adını "X.xml" yap
    const cd = findLast(bozuk, enc.encode("y.xml"));
    bozuk.set(enc.encode("X.xml"), cd);
    const zb = readZip(bozuk);
    expect(zb.entries.map((e) => e.name)).toEqual(["x.xml"]);
    expect(zb.skipped).toBe(1);
  });

  it("ZIP olmayan ve kesik girdi OpcError verir, çökmez", async () => {
    expect(() => readZip(new Uint8Array(10))).toThrow(OpcError);
    expect(() => readZip(enc.encode("PK\u0003\u0004 bu bir zip değil ama yeterince uzun bir metin"))).toThrow(OpcError);
    try {
      readZip(new Uint8Array(100));
    } catch (e) {
      expect((e as OpcError).code).toBe("not-zip");
    }
    const b = await arsiv({ "a.xml": "x".repeat(1000) });
    // Yerel başlığı bozuk girdi atlanır
    const bozuk = b.slice();
    bozuk[0] = 0;
    const z = readZip(bozuk);
    expect(z.entries).toHaveLength(0);
    expect(z.skipped).toBe(1);
    expect(isZip(bozuk)).toBe(false);
  });

  it("bozuk deflate verisi 'corrupt' hatası verir", async () => {
    const b = await arsiv({ "a.xml": "abc ".repeat(400) });
    const z = readZip(b);
    const bozuk = b.slice();
    const e = z.entries[0];
    bozuk.fill(0xff, e.dataOffset, e.dataOffset + e.compressedSize);
    await expect(inflateEntry(bozuk, e, 1 << 20)).rejects.toMatchObject({ code: "corrupt" });
  });

  it("veri tanımlayıcılı (bit 3) girdi merkez dizinden okunur, yazarken bit kalkar", async () => {
    const b = await arsiv({ "a.xml": "veri ".repeat(300) });
    const dv = new DataView(b.buffer);
    // yerel başlıkta bit 3 ve sıfır boyut/CRC (akışla yazılmış arşiv gibi)
    dv.setUint16(6, 0x0008, true);
    dv.setUint32(14, 0, true);
    dv.setUint32(18, 0, true);
    dv.setUint32(22, 0, true);
    const z = readZip(b);
    expect(dec.decode(await inflateEntry(b, z.entries[0], 1 << 20))).toBe("veri ".repeat(300));
    const w = new ZipWriter();
    w.addRaw("a.xml", b, { ...z.entries[0], flags: 0x0008 });
    const c = await w.finish();
    expect(new DataView(c.buffer).getUint16(6, true) & 0x0008).toBe(0);
    expect(new DataView(c.buffer).getUint32(14, true)).toBe(z.entries[0].crc32);
  });

  it("sıkıştırma bombası: sınır aşılınca kesilir (bildirilen boyut yalan olsa da)", async () => {
    const sifir = new Uint8Array(8 * 1024 * 1024);
    const b = await arsiv({ "bomba.xml": sifir });
    const z = readZip(b);
    const e = z.entries[0];
    expect(e.compressedSize).toBeLessThan(64 * 1024);
    await expect(inflateEntry(b, e, 1024 * 1024)).rejects.toMatchObject({ code: "too-large" });
    // merkez dizin boyutu küçük gösterse de gerçek çıktı sınırla kesilir
    await expect(inflateEntry(b, { ...e, size: 10 }, 1024 * 1024)).rejects.toMatchObject({ code: "too-large" });
    expect((await inflateEntry(b, e, 16 * 1024 * 1024)).length).toBe(sifir.length);
  });

  it("şifreli girdi ve bilinmeyen yöntem 'unsupported'", async () => {
    const b = await arsiv({ "a.xml": "x".repeat(200) });
    const e = readZip(b).entries[0];
    await expect(inflateEntry(b, { ...e, flags: 1 }, 1000)).rejects.toMatchObject({ code: "unsupported" });
    await expect(inflateEntry(b, { ...e, method: 12 }, 1000)).rejects.toMatchObject({ code: "unsupported" });
  });

  it("ZIP64 ek alanındaki boyutlar okunur", async () => {
    const name = enc.encode("z.xml");
    const data = enc.encode("<zip64/>");
    const out = new Uint8Array(30 + name.length + data.length + 46 + name.length + 20 + 22);
    const dv = new DataView(out.buffer);
    let p = 0;
    dv.setUint32(p, 0x04034b50, true);
    dv.setUint16(p + 4, 45, true);
    dv.setUint32(p + 14, crc32(data), true);
    dv.setUint32(p + 18, 0xffffffff, true);
    dv.setUint32(p + 22, 0xffffffff, true);
    dv.setUint16(p + 26, name.length, true);
    out.set(name, p + 30);
    out.set(data, p + 30 + name.length);
    p += 30 + name.length + data.length;
    const cd = p;
    dv.setUint32(p, 0x02014b50, true);
    dv.setUint32(p + 16, crc32(data), true);
    dv.setUint32(p + 20, 0xffffffff, true);
    dv.setUint32(p + 24, 0xffffffff, true);
    dv.setUint16(p + 28, name.length, true);
    dv.setUint16(p + 30, 20, true);
    out.set(name, p + 46);
    const ex = p + 46 + name.length;
    dv.setUint16(ex, 0x0001, true);
    dv.setUint16(ex + 2, 16, true);
    dv.setUint32(ex + 4, data.length, true); // açık boyut
    dv.setUint32(ex + 12, data.length, true); // sıkışık boyut
    p = ex + 20;
    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 8, 1, true);
    dv.setUint16(p + 10, 1, true);
    dv.setUint32(p + 12, p - cd, true);
    dv.setUint32(p + 16, cd, true);
    const z = readZip(out);
    expect(z.entries).toHaveLength(1);
    expect(z.entries[0]).toMatchObject({ size: data.length, compressedSize: data.length });
    expect(dec.decode(await inflateEntry(out, z.entries[0], 100))).toBe("<zip64/>");
  });

  it("çok girdili arşiv: merkez dizin sayısı yalan olsa da döngü sınırlıdır", async () => {
    const parts: Record<string, string> = {};
    for (let i = 0; i < 300; i++) parts[`p/${i}.xml`] = `<a i="${i}"/>`;
    const b = await arsiv(parts);
    const dv = new DataView(b.buffer);
    const eocd = b.length - 22;
    dv.setUint16(eocd + 8, 0xfffe, true);
    dv.setUint16(eocd + 10, 0xfffe, true);
    const z = readZip(b);
    expect(z.entries).toHaveLength(300);
  });
});

function findLast(hay: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = hay.length - needle.length; i >= 0; i--) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}
