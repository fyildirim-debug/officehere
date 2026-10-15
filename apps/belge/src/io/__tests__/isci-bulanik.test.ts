// @vitest-environment jsdom
// Bozuk XML'de yerleşik ayrıştırıcı (Worker) tarayıcınınki kadar katıdır: rastgele bozulan parçalarda ikisi de
// aynı sonucu verir (ikisi de reddeder ya da ikisi de aynı ağacı kurar).
import { installXmlDom } from "../xmlDom";
import { parseXml, setBuiltinXml } from "../docx/xml";

/** Belirlenimci sözde rastgele (mulberry32). */
function rastgele(tohum: number) {
  let a = tohum >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const ORNEK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body><!-- açıklama --><w:p w14:paraId="1A2B3C4D"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:i w:val="0"/></w:rPr><w:t xml:space="preserve">Başlık &amp; &lt;özel&gt; &#x130;&#305; "tırnak" 'tek'</w:t></w:r></w:p>
<w:p><w:r><w:t><![CDATA[ham <metin> & ]]></w:t></w:r><w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r><?yonerge veri?></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p a="satır
sonu	sekme&#10;"/></mc:Choice><mc:Fallback><w:p xmlns="urn:varsayilan"><x xmlns="" y="1"/></w:p></mc:Fallback></mc:AlternateContent><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

const PARCALAR = ["<", ">", "&", "/", '"', "'", "=", ":", " ", "x:", "&amp", "]]>", "<!--", "-->", "\u0001", "w:", "xmlns:q=\"\"", "<w:p>", "</w:p>", "&#0;", "&nbsp;", "<?xml ?>", "￾"];

function boz(s: string, r: () => number): string {
  const i = Math.floor(r() * s.length);
  const n = Math.floor(r() * 3);
  if (n === 0) return s.slice(0, i) + s.slice(i + 1 + Math.floor(r() * 4));
  if (n === 1) return s.slice(0, i) + PARCALAR[Math.floor(r() * PARCALAR.length)] + s.slice(i);
  return s.slice(0, i);
}

/** Ağacın karşılaştırılabilir özeti. */
function ozet(n: Node): unknown {
  if (n.nodeType === 1) {
    const e = n as Element;
    const attrs = Array.from(e.attributes).map((a) => [a.namespaceURI, a.prefix, a.localName, a.value]);
    const kids: unknown[] = [];
    for (let c = e.firstChild; c; c = c.nextSibling) kids.push(ozet(c));
    return [e.namespaceURI, e.prefix, e.localName, attrs, kids];
  }
  if (n.nodeType === 3 || n.nodeType === 4) return [n.nodeType, n.nodeValue];
  if (n.nodeType === 8) return ["#", n.nodeValue];
  if (n.nodeType === 7) return ["?", (n as ProcessingInstruction).target, n.nodeValue];
  return [n.nodeType];
}

/** Örnekte değişen bölge (hata iletisi için). */
function bolge(s: string): string {
  let a = 0;
  while (a < s.length && a < ORNEK.length && s[a] === ORNEK[a]) a++;
  let b = 0;
  while (b < s.length - a && b < ORNEK.length - a && s[s.length - 1 - b] === ORNEK[ORNEK.length - 1 - b]) b++;
  return `${JSON.stringify(ORNEK.slice(Math.max(0, a - 30), ORNEK.length - b + 10))} → ${JSON.stringify(s.slice(Math.max(0, a - 30), s.length - b + 10))}`;
}

/** İki ağaç özetinin ilk farkı. */
function ilkFark(a: unknown, b: unknown, yol = ""): string {
  if (JSON.stringify(a) === JSON.stringify(b)) return "";
  if (Array.isArray(a) && Array.isArray(b)) {
    for (let i = 0; i < Math.max(a.length, b.length); i++) {
      const d = ilkFark(a[i], b[i], `${yol}/${i}`);
      if (d) return d;
    }
  }
  return `${yol}: ${JSON.stringify(a)?.slice(0, 150)} ≠ ${JSON.stringify(b)?.slice(0, 150)}`;
}

function sonuc(xml: string, yerlesik: boolean): unknown {
  if (yerlesik) installXmlDom({ force: true });
  else setBuiltinXml(null);
  try {
    return ozet(parseXml(xml).documentElement);
  } catch {
    return "HATA";
  }
}

describe("yerleşik XML ayrıştırıcısı", () => {
  afterAll(() => setBuiltinXml(null));

  it("sağlam örnekte DOMParser ile aynı ağacı kurar", () => {
    expect(sonuc(ORNEK, true)).toEqual(sonuc(ORNEK, false));
    expect(sonuc(ORNEK, true)).not.toBe("HATA");
  });

  // Uzun koşu: OH_BULANIK_ADET=40000 OH_BULANIK_TOHUM=7
  const ADET = Number(process.env.OH_BULANIK_ADET ?? 2000);
  it(`${ADET} rastgele bozulmada DOMParser ile aynı sonucu verir`, () => {
    const r = rastgele(Number(process.env.OH_BULANIK_TOHUM ?? 41));
    const farkli: string[] = [];
    let hata = 0;
    for (let k = 0; k < ADET; k++) {
      let s = ORNEK;
      for (let m = 1 + Math.floor(r() * 2); m > 0; m--) s = boz(s, r);
      // jsdom'un ayrıştırıcısı (saxes) ad alanı adresinin baş/son boşluğunu kırpar; tarayıcı kırpmaz.
      if (/xmlns(?::[^=\s]+)?\s*=\s*(["'])(?:\s[^"']*|[^"']*\s)\1/.test(s)) continue;
      const a = sonuc(s, false);
      const b = sonuc(s, true);
      if (a === "HATA") hata++;
      if (JSON.stringify(a) !== JSON.stringify(b) && farkli.length < 12) farkli.push(`${bolge(s)} ⇒ ${a === "HATA" ? "yalnız DOMParser reddetti" : b === "HATA" ? "yalnız yerleşik reddetti" : ilkFark(a, b)}`);
    }
    expect(farkli).toEqual([]);
    expect(hata).toBeGreaterThan(ADET / 10);
  }, 600_000);
});
