// @vitest-environment jsdom
// Resmî yazı şablonlarının ölçüleri (Yönetmelik m.6-27): kenar, punto, girinti, ":" hizası, sayı ve tarih aynı
// satır, satır boşlukları, imza bloğunun konumu; .docx çıktısındaki karşılıkları ve saydamlık yokluğu.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { twToCm } from "../../model/document";
import { writeDocx } from "../../io";
import { resmiYaziDenetle } from "../../denetim/resmiYazi";
import { TEMPLATES } from "../templates";
import { RESMI_SABLONLAR, resmiBelge } from "./sablonlar";
import { GIRINTI_PT, IKI_NOKTA_PT, ILGI_METIN_PT, IMZA_BLOK_CM, KIRMIZI, YAN_METIN_PT, YAZI_ALANI_PT } from "./olcu";

const FORM = { tarih: "15.10.2026" };
const bloklar = (id: string, form = {}) => resmiBelge(id, { ...FORM, ...form }).content.content ?? [];
const metni = (n: JSONContent): string => (n.content ?? []).map((c) => (c.type === "text" ? (c.text ?? "") : metni(c))).join("");
const bosMu = (n: JSONContent) => n.type === "paragraph" && !metni(n);
const ptCm = (pt: number) => (pt / 72) * 2.54;
const bul = (b: JSONContent[], re: RegExp) => b.findIndex((n) => re.test(metni(n)));
/** i ile j arasındaki boş paragraf sayısı (ikisi hariç). */
const arasi = (b: JSONContent[], i: number, j: number) => b.slice(i + 1, j).filter(bosMu).length;
const USTYAZI = RESMI_SABLONLAR.filter((s) => !["resmi-zarf", "resmi-dilekce", "resmi-tutanak"].includes(s.id)).map((s) => s.id);

describe("Resmî yazı şablonları: sayfa ve yazı tipi", () => {
  it("küme 8 şablondan oluşur ve hepsi Dosya › Yeni şablonlarında kayıtlı", () => {
    expect(RESMI_SABLONLAR).toHaveLength(8);
    for (const s of RESMI_SABLONLAR) expect(TEMPLATES.find((t) => t.id === s.id)?.kume).toBe("resmi");
    expect(TEMPLATES.find((t) => t.id === "resmi")?.kume).toBeUndefined();
  });

  it("A4; üst, sol, sağ kenar 1,5 cm (m.6, m.8); logo varsa üst 0,5 cm", () => {
    for (const s of RESMI_SABLONLAR.filter((x) => x.id !== "resmi-zarf")) {
      const { page } = resmiBelge(s.id, FORM).settings;
      expect([page.width, page.height], s.id).toEqual([11906, 16838]);
      for (const k of ["top", "left", "right"] as const) expect(twToCm(page.margin[k]), `${s.id} ${k}`).toBeCloseTo(1.5, 2);
    }
    const logolu = resmiBelge("resmi-ust-yazi", { ...FORM, logo: true });
    expect(twToCm(logolu.settings.page.margin.top)).toBeCloseTo(0.5, 2);
    // Sol üst logo yer tutucusu: belgenin başında dolgusuz metin kutusu; başlık üst kenardan 1,5 cm'de (0,5 + 1).
    const [kutu, tc] = logolu.content.content ?? [];
    expect(kutu.type).toBe("textBox");
    expect(kutu.attrs).toMatchObject({ x: 0, y: 0, fill: null });
    expect(metni(tc)).toBe("T.C.");
    expect(ptCm(tc.attrs?.spaceBefore as number)).toBeCloseTo(1, 2);
  });

  it("Times New Roman 12 ya da Arial 11; paragraf aralığı 0, tek satır (m.7, m.16/4)", () => {
    for (const s of RESMI_SABLONLAR) {
      const n = resmiBelge(s.id, FORM).settings.customStyles?.Normal;
      expect(n?.props, s.id).toMatchObject({ font: "Times New Roman", size: 12, spaceBefore: 0, spaceAfter: 0, line: 1 });
    }
    expect(resmiBelge("resmi-ust-yazi", { yaziTipi: "arial" }).settings.customStyles?.Normal.props).toMatchObject({ font: "Arial", size: 11 });
  });
});

describe("Resmî yazı şablonları: bölümlerin yeri ve hizası", () => {
  it("başlık ortalı: T.C., idare BÜYÜK HARFLE, birim (m.10); bağlı idarede 4 satır", () => {
    const b = bloklar("resmi-ust-yazi", { idare: "Örnek Bakanlığı", birim: "Strateji Dairesi Başkanlığı" });
    expect(b.slice(0, 3).map(metni)).toEqual(["T.C.", "ÖRNEK BAKANLIĞI", "Strateji Dairesi Başkanlığı"]);
    for (const n of b.slice(0, 3)) expect(n.attrs?.textAlign).toBe("center");
    const d = bloklar("resmi-ust-yazi", { baslikTuru: "bagli", idare: "ÖRNEK BAKANLIĞI", araIdare: "Örnek Kurumu Başkanlığı", birim: "Personel Dairesi Başkanlığı" });
    expect(d.slice(0, 4).map(metni)).toEqual(["T.C.", "ÖRNEK BAKANLIĞI", "Örnek Kurumu Başkanlığı", "Personel Dairesi Başkanlığı"]);
  });

  it("sayı ve tarih aynı satırda; tarih yazı alanının en sağında (m.11, m.12)", () => {
    for (const id of USTYAZI) {
      const b = bloklar(id);
      const sayi = b[bul(b, /^Sayı\t/)];
      expect(metni(sayi), id).toMatch(/^Sayı\t:\tE-12345678-903\.07\.02-\S+\t15\.10\.2026$/);
      const sag = (sayi.attrs?.tabs as { pos: number; align: string }[]).find((t) => t.align === "right");
      expect(ptCm(sag!.pos)).toBeCloseTo(18, 1);
      expect(ptCm(YAZI_ALANI_PT)).toBeCloseTo(21 - 3, 2);
    }
  });

  it("Sayı, Konu ve İlgi'nin ':' işareti aynı hizada; ikinci satırlar yan başlığın altı boş başlar (m.13, m.15)", () => {
    const b = bloklar("resmi-ust-yazi", { ilgiler: "Birinci yazı.\nİkinci yazı." });
    const satirlar = [bul(b, /^Sayı\t/), bul(b, /^Konu\t/), bul(b, /^İlgi\t/), bul(b, /^\t\tb\)/)].map((i) => b[i]);
    for (const n of satirlar) {
      const tabs = n.attrs?.tabs as { pos: number; align: string }[];
      expect(tabs[0]).toMatchObject({ pos: IKI_NOKTA_PT, align: "left" });
      expect(tabs[1]).toMatchObject({ pos: YAN_METIN_PT, align: "left" });
    }
    expect(metni(satirlar[1])).toMatch(/^Konu\t:\t/);
    expect(satirlar[1].attrs).toMatchObject({ indentLeft: YAN_METIN_PT, indentFirst: -YAN_METIN_PT });
    expect(metni(satirlar[2])).toBe("İlgi\t:\ta)\tBirinci yazı.");
    expect(satirlar[2].attrs).toMatchObject({ indentLeft: ILGI_METIN_PT, indentFirst: -ILGI_METIN_PT });
    expect(metni(satirlar[3])).toBe("\t\tb)\tİkinci yazı.");
    // Tek ilgide harf yok (m.15/5).
    const t = bloklar("resmi-ust-yazi", { ilgiler: "Tek yazı." });
    expect(metni(t[bul(t, /^İlgi\t/)])).toBe("İlgi\t:\tTek yazı.");
  });

  it("satır boşlukları: başlık→sayı 2, konu→muhatap 2, muhatap→ilgi 2, ilgi→metin 1, ilgi yoksa muhatap→metin 2, imza 2-4", () => {
    const b = bloklar("resmi-ust-yazi");
    const birim = bul(b, /^Destek Hizmetleri/);
    const sayi = bul(b, /^Sayı\t/);
    const konu = bul(b, /^Konu\t/);
    const muh = bul(b, /^ÖRNEK VALİLİĞİNE$/);
    const muhBirim = bul(b, /^\(İl Millî Eğitim Müdürlüğü\)$/);
    const ilgi = bul(b, /^İlgi\t/);
    const metin = bul(b, /^İlgi yazı ile/);
    const kapanis = bul(b, /arz ederim\.$/);
    const imza = bul(b, /^Ad SOYAD$/);
    expect(arasi(b, birim, sayi)).toBe(2);
    expect(konu).toBe(sayi + 1);
    expect(arasi(b, konu, muh)).toBe(2);
    expect(muhBirim).toBe(muh + 1);
    expect(arasi(b, muhBirim, ilgi)).toBe(2);
    expect(arasi(b, ilgi, metin)).toBe(1);
    expect(metin + 2).toBe(kapanis);
    expect(arasi(b, kapanis, imza)).toBeGreaterThanOrEqual(2);
    expect(arasi(b, kapanis, imza)).toBeLessThanOrEqual(4);
    const ilgisiz = bloklar("resmi-ust-yazi", { ilgiler: "", muhatapBirim: "" });
    expect(arasi(ilgisiz, bul(ilgisiz, /^ÖRNEK VALİLİĞİNE$/), bul(ilgisiz, /^İlgi yazı ile/))).toBe(2);
  });

  it("muhatap ortalı, BÜYÜK HARFLE; birim parantez içinde alt satırda; gerçek kişide 'Sayın' (m.14)", () => {
    const b = bloklar("resmi-ust-yazi", { muhatap: "örnek valiliğine", muhatapBirim: "İl Millî Eğitim Müdürlüğü" });
    const i = bul(b, /^ÖRNEK VALİLİĞİNE$/);
    expect(b[i].attrs?.textAlign).toBe("center");
    expect(metni(b[i + 1])).toBe("(İl Millî Eğitim Müdürlüğü)");
    const k = bloklar("resmi-ust-yazi", { muhatap: "Sayın Ayşe YILMAZ", muhatapBirim: "" });
    expect(bul(k, /^Sayın Ayşe YILMAZ$/)).toBeGreaterThan(0);
  });

  it("metin paragrafları 1,25 cm girintili ve iki yana yaslı (m.16/4)", () => {
    expect(ptCm(GIRINTI_PT)).toBeCloseTo(1.25, 2);
    for (const id of [...USTYAZI, "resmi-dilekce", "resmi-tutanak"]) {
      const b = bloklar(id);
      const metinler = b.filter((n) => n.type === "paragraph" && metni(n).length > 60 && !/^((Sayı|Konu|İlgi)\t|\t|(Adres|Telefon|e-Posta): )/.test(metni(n)));
      expect(metinler.length, id).toBeGreaterThan(0);
      for (const n of metinler) expect(n.attrs, `${id}: ${metni(n).slice(0, 30)}`).toMatchObject({ indentFirst: GIRINTI_PT, textAlign: "justify" });
    }
  });

  it("imza bloğu yazı alanının en sağında ortalı (m.17/1)", () => {
    for (const id of [...USTYAZI, "resmi-dilekce"]) {
      const b = bloklar(id);
      const i = b.findIndex((n, k) => metni(n) === "Ad SOYAD" && (metni(b[k + 1] ?? {}) === "Unvan" || metni(b[k + 1] ?? {}) === "İmza") && b[k].attrs?.indentLeft);
      expect(i, id).toBeGreaterThan(0);
      for (const n of [b[i], b[i + 1]]) {
        expect(n.attrs?.textAlign).toBe("center");
        expect(n.attrs?.indentRight ?? 0).toBe(0);
        // Blok: sol girintiden yazı alanının sağ kenarına; genişlik 7 cm, ortası sağ yarıda.
        const sol = ptCm(n.attrs?.indentLeft as number);
        expect(18 - sol).toBeCloseTo(IMZA_BLOK_CM, 1);
        expect((sol + 18) / 2).toBeGreaterThan(9);
      }
    }
  });

  it("olur: ara makam solda 'Uygun görüşle arz ederim.', ortada 'OLUR' ve altında ad soyad, unvan (m.20)", () => {
    const b = bloklar("resmi-olur");
    const ara = bul(b, /^Uygun görüşle arz ederim\.$/);
    expect(b[ara].attrs).toMatchObject({ textAlign: "center" });
    expect(ptCm(b[ara].attrs?.indentRight as number)).toBeCloseTo(18 - IMZA_BLOK_CM, 1);
    const olur = bul(b, /^OLUR$/);
    expect(olur).toBeGreaterThan(ara);
    expect(b[olur].attrs).toMatchObject({ textAlign: "center" });
    expect(b[olur].attrs?.indentLeft ?? null).toBeNull();
    expect([metni(b[olur + 1]), metni(b[olur + 2])]).toEqual(["Ad SOYAD", "Unvan"]);
    // m.18/1: olurda ek, oluru veren makamın imzasından sonra.
    expect(bul(b, /^Ek: /)).toBeGreaterThan(olur);
    // m.20/3: el ile imzada OLUR'dan sonra tarih ve imza yeri.
    const z = bloklar("resmi-olur", { ortam: "Z" });
    expect(metni(z[bul(z, /^OLUR$/) + 1])).toBe("…/…/20…");
  });

  it("ek ve dağıtım: tek ek aynı satırda, çok ek numaralı; Gereği ve Bilgi aynı satırda, Bilgi ortaya doğru (m.18, m.19)", () => {
    const tek = bloklar("resmi-ust-yazi", { ekler: "Liste (2 sayfa)" });
    expect(bul(tek, /^Ek: Liste \(2 sayfa\)$/)).toBeGreaterThan(0);
    const cok = bloklar("resmi-ust-yazi", { ekler: "A\nB" });
    const ek = bul(cok, /^Ek:$/);
    expect([metni(cok[ek + 1]), metni(cok[ek + 2])]).toEqual(["1- A", "2- B"]);
    const d = bloklar("resmi-dagitimli");
    const g = bul(d, /^Gereği:\tBilgi:$/);
    expect(metni(d[g - 1])).toBe("Dağıtım:");
    expect(ptCm((d[g].attrs?.tabs as { pos: number }[])[0].pos)).toBeCloseTo(9, 1);
    expect(metni(d[g + 1])).toBe("Örnek Kaymakamlığına\tÖrnek Bakanlığına");
    expect(bul(d, /^Dağıtım:$/)).toBeGreaterThan(bul(d, /^Ek: /));
  });

  it("ek ve dağıtım listesi ayrı sayfalarda; üst yazıda 'Ek: Ek Listesi' ve 'DAĞITIM LİSTESİ' (m.18/4, m.19/3)", () => {
    const b = bloklar("resmi-ek-dagitim");
    expect(b.filter((n) => n.type === "pageBreak")).toHaveLength(2);
    expect(bul(b, /^Ek: Ek Listesi$/)).toBeGreaterThan(0);
    const dl = bul(b, /^DAĞITIM LİSTESİ$/);
    expect(metni(b[dl - 1])).toBe("Dağıtım:");
    expect(b.findIndex((n) => n.type === "pageBreak")).toBeGreaterThan(dl);
    expect(b.filter((n) => /^(EK LİSTESİ|DAĞITIM LİSTESİ)$/.test(metni(n)) && n.attrs?.textAlign === "center")).toHaveLength(2);
  });

  it("tekit: sağ üstte kırmızı GÜNLÜDÜR (m.26/3); tutanakta tarih metnin bitiminde ve imza tablosu (m.12/4)", () => {
    const t = bloklar("resmi-tekit");
    expect(metni(t[0])).toBe("GÜNLÜDÜR");
    expect(t[0].attrs?.textAlign).toBe("right");
    expect(t[0].content?.[0].marks?.[0].attrs?.color).toBe(KIRMIZI);
    const r = bloklar("resmi-tutanak");
    const tarih = bul(r, /^15\.10\.2026$/);
    expect(r[tarih].attrs?.textAlign).toBe("right");
    expect(bul(r, /^İşbu tutanak/)).toBeLessThan(tarih);
    const tablo = r.find((n) => n.type === "table")!;
    expect(r.indexOf(tablo)).toBeGreaterThan(tarih);
    expect(tablo.content?.[0].content?.map(metni)).toEqual(["Ad SOYAD", "Unvan", "İmza"]);
    expect(tablo.content).toHaveLength(4);
  });

  it("zarf: DL 22 × 11 cm, C4 32,4 × 22,9 cm; sol üstte başlık, sayı, tarih; sağ üstte kırmızı ivedilik (m.31/2)", () => {
    const dl = resmiBelge("resmi-zarf", FORM);
    expect([twToCm(dl.settings.page.width), twToCm(dl.settings.page.height)].map((x) => +x.toFixed(1))).toEqual([22, 11]);
    const c4 = resmiBelge("resmi-zarf", { ...FORM, zarfBoyu: "C4" });
    expect([twToCm(c4.settings.page.width), twToCm(c4.settings.page.height)].map((x) => +x.toFixed(1))).toEqual([32.4, 22.9]);
    const b = dl.content.content ?? [];
    expect(metni(b[0])).toBe("T.C.\tACELE");
    expect(b[0].content?.[1].marks?.[0].attrs?.color).toBe(KIRMIZI);
    expect(bul(b, /^Sayı: E-/)).toBeLessThan(bul(b, /^ÖRNEK BAKANLIĞINA$/));
    expect(b[bul(b, /^ÖRNEK BAKANLIĞINA$/)].attrs?.textAlign).toBe("center");
  });

  it("alt bilgi: doğrulama yeri, çizgi, 8 punto iletişim bilgileri ve ortada sayfa numarası 'n/N' (m.23, m.24, m.27)", () => {
    const f = resmiBelge("resmi-ust-yazi", FORM).settings.footer!;
    const [d1, d2, tablo, no] = f.content ?? [];
    expect(bosMu(d1) && bosMu(d2)).toBe(true);
    expect(d2.attrs?.borders?.bottom).toMatchObject({ style: "single" });
    const yazilar = JSON.stringify(tablo);
    expect(yazilar).toContain("Adres: ");
    expect(yazilar).toContain("Bilgi için: ");
    // Tüm alt bilgi paragrafları 8 puntoluk, aralıksız alt bilgi stilinde.
    const stil = resmiBelge("resmi-ust-yazi", FORM).settings.customStyles?.ResmiAltBilgi;
    expect(stil?.props).toMatchObject({ size: 8, spaceBefore: 0, spaceAfter: 0, line: 1 });
    const paragraflar = JSON.stringify(f).match(/"type":"paragraph","attrs":\{[^}]*\}/g) ?? [];
    expect(paragraflar.length).toBe(10);
    for (const x of paragraflar) expect(x).toContain('"styleId":"ResmiAltBilgi"');
    expect(no.attrs?.textAlign).toBe("center");
    expect(no.content?.map((c) => c.attrs?.field ?? c.text)).toEqual(["PAGE", "/", "NUMPAGES"]);
  });
});

describe("Resmî yazı şablonları: kendi denetimleri ve .docx çıktısı", () => {
  it("her şablon Yönetmeliğe uygunluk denetiminden hatasız geçer (zarf yalnız A4 uyarısı)", () => {
    for (const s of RESMI_SABLONLAR) {
      const sonuc = resmiYaziDenetle(resmiBelge(s.id, FORM));
      const sorun = sonuc.flatMap((k) => k.bulgular.filter((b) => b.durum === "hata" || b.durum === "uyari").map((b) => `${k.kural}: ${b.ileti}`));
      if (s.id === "resmi-zarf") expect(sorun).toEqual([expect.stringMatching(/^kenar: Kâğıt A4 değil/)]);
      else expect(sorun, s.id).toEqual([]);
    }
  });

  it(".docx: kenar 850 twip, girinti 709 twip, iki yana yaslı, sekmeler, TNR 12, alt bilgi 8 punto; saydamlık yok", async () => {
    const zip = await JSZip.loadAsync(await writeDocx(resmiBelge("resmi-ust-yazi", { ...FORM, logo: true })));
    const belge = await zip.file("word/document.xml")!.async("string");
    expect(belge).toMatch(/<w:pgMar w:top="283" w:right="850" w:bottom="\d+" w:left="850"/);
    expect(belge).toContain('<w:ind w:firstLine="709"/>');
    expect(belge).toContain('<w:jc w:val="both"/>');
    expect(belge).toContain('<w:tab w:val="left" w:pos="709"/>');
    expect(belge).toContain('<w:tab w:val="right" w:pos="10205"/>');
    const stiller = await zip.file("word/styles.xml")!.async("string");
    // Normal stilin yazı tipi, boyutu ve aralığı belge varsayılanlarına yazılır.
    const varsayilan = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(stiller)?.[0] ?? "";
    expect(varsayilan).toContain('w:ascii="Times New Roman"');
    expect(varsayilan).toContain('<w:sz w:val="24"/>');
    expect(varsayilan).toContain('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>');
    const alt = Object.keys(zip.files).filter((n) => /^word\/footer\d*\.xml$/.test(n));
    expect(alt.length).toBeGreaterThan(0);
    expect(await zip.file(alt[0])!.async("string")).toContain('<w:pStyle w:val="ResmiAltBilgi"/>');
    expect(/<w:style [^>]*w:styleId="ResmiAltBilgi">[\s\S]*?<\/w:style>/.exec(stiller)?.[0]).toContain('<w:sz w:val="16"/>');
    // PDF/A-1b için: hiçbir parçada saydamlık (alfa, yarı saydam dolgu) yok.
    for (const ad of Object.keys(zip.files).filter((n) => n.endsWith(".xml"))) {
      const x = await zip.file(ad)!.async("string");
      expect(x, ad).not.toMatch(/<a:alpha|<a:alphaMod|w14:alpha|opacity=|o:opacity/);
    }
  });
});
