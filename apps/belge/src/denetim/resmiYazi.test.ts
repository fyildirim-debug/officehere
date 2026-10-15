// Yönetmeliğe uygunluk denetimi: her kural için olumlu (uygun belge) ve olumsuz (kuralı çiğneyen belge) sınamalar.
import { describe, expect, it } from "vitest";
import type { JSONContent } from "@tiptap/core";
import { cmToTw, defaultSettings, type BelgeDocument } from "../model/document";
import { resmiBelge } from "../templates/resmi/sablonlar";
import { GIRINTI_PT } from "../templates/resmi/olcu";
import { paragraflar, resmiYaziDenetle, sayiDuzelt, tarihDuzelt, type DenetimYonu, type KuralKimligi } from "./resmiYazi";

const FORM = { tarih: "15.10.2026" };
const ust = (form = {}) => resmiBelge("resmi-ust-yazi", { ...FORM, ...form });
const kural = (doc: BelgeDocument, k: KuralKimligi, yon: DenetimYonu = "otomatik") => resmiYaziDenetle(doc, { yon }).find((r) => r.kural === k)!;
const p = (text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, content: text ? [{ type: "text", text }] : undefined });
/** Belgedeki ilk eşleşen paragrafın metnini değiştirir. */
function degistir(doc: BelgeDocument, re: RegExp, yeni: string): BelgeDocument {
  const kopya = structuredClone(doc);
  const n = (kopya.content.content ?? []).find((x) => re.test(x.content?.map((c) => c.text ?? "").join("") ?? ""));
  if (!n) throw new Error(`bulunamadı: ${re}`);
  n.content = [{ type: "text", text: yeni }];
  return kopya;
}

describe("paragraf çıkarma", () => {
  it("paragraf ve başlıkları belge sırasıyla, tablo ve metin kutusu içindekilerle birlikte verir", () => {
    const list = paragraflar({
      type: "doc",
      content: [
        { type: "textBox", content: [p("Kutu")] },
        p("Bir", { textAlign: "center" }),
        { type: "table", content: [{ type: "tableRow", content: [{ type: "tableCell", content: [p("Hücre")] }] }] },
        { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Başlık", marks: [{ type: "textStyle", attrs: { fontFamily: '"Arial", sans-serif', fontSize: "16px" } }] }] },
      ],
    });
    expect(list.map((x) => [x.indeks, x.metin, x.tabloda, x.kutuda])).toEqual([
      [0, "Kutu", false, true],
      [1, "Bir", false, false],
      [2, "Hücre", true, false],
      [3, "Başlık", false, false],
    ]);
    expect(list[3].kosular[0]).toEqual({ metin: "Başlık", yaziTipi: "Arial", punto: 12 });
  });
});

describe("Kural 1: kâğıt ve kenar boşlukları (m.6, m.8)", () => {
  it("olumlu: A4 ve 1,5 cm; logolu belgede üst 0,5 cm", () => {
    expect(kural(ust(), "kenar").durum).toBe("uygun");
    expect(kural(ust({ logo: true }), "kenar")).toMatchObject({ durum: "uygun", ozet: expect.stringContaining("0,5 cm (logo)") });
  });
  it("olumsuz: varsayılan 2,5 cm kenarlar; logosuz 0,5 cm üst kenar; A4 olmayan kâğıt", () => {
    const d = { settings: defaultSettings(), content: { type: "doc", content: [p("Metin")] } };
    const k = kural(d, "kenar");
    expect(k.durum).toBe("hata");
    expect(k.bulgular.map((b) => b.ileti)).toEqual(expect.arrayContaining([expect.stringMatching(/^Sol kenar boşluğu 2,5 cm; 1,5 cm olmalı/), expect.stringMatching(/^Üst kenar boşluğu 2,5 cm/)]));
    const logosuz = ust();
    logosuz.settings.page.margin.top = cmToTw(0.5);
    expect(kural(logosuz, "kenar").bulgular[0].ileti).toMatch(/yalnız logo kullanıldığında/);
    const zarf = resmiBelge("resmi-zarf", FORM);
    expect(kural(zarf, "kenar")).toMatchObject({ durum: "uyari", bulgular: [{ ileti: expect.stringMatching(/^Kâğıt A4 değil \(22 × 11 cm\)/) }] });
  });
});

describe("Kural 2: yazı tipi ve punto (m.7)", () => {
  it("olumlu: Times New Roman 12, Arial 11; alt bilgide 8 punto", () => {
    expect(kural(ust(), "yaziTipi")).toMatchObject({ durum: "uygun", ozet: "Times New Roman 12 punto." });
    expect(kural(ust({ yaziTipi: "arial" }), "yaziTipi")).toMatchObject({ durum: "uygun", ozet: "Arial 11 punto." });
  });
  it("olumsuz: tema yazı tipi, yanlış punto, başka yazı tipi, 9 puntodan küçük, alt bilgide 8'den küçük, iki yazı tipi", () => {
    const tema = { settings: defaultSettings(), content: { type: "doc", content: [p("Metin")] } };
    expect(kural(tema, "yaziTipi").bulgular[0]).toMatchObject({ durum: "hata", ileti: expect.stringMatching(/^Belgenin yazı tipi Aptos/) });
    const on1 = ust();
    on1.settings.customStyles!.Normal.props.size = 11;
    expect(kural(on1, "yaziTipi").bulgular[0].ileti).toMatch(/Times New Roman 11 punto; Times New Roman için 12 punto esastır/);
    const d = ust();
    d.content.content!.push(
      { type: "paragraph", content: [{ type: "text", text: "Süslü", marks: [{ type: "textStyle", attrs: { fontFamily: "Georgia" } }] }] },
      { type: "paragraph", content: [{ type: "text", text: "Küçük", marks: [{ type: "textStyle", attrs: { fontSize: "8pt" } }] }] },
      { type: "paragraph", content: [{ type: "text", text: "Büyük", marks: [{ type: "textStyle", attrs: { fontSize: "16pt" } }] }] },
      { type: "paragraph", content: [{ type: "text", text: "Arial", marks: [{ type: "textStyle", attrs: { fontFamily: "Arial" } }] }] },
    );
    d.settings.footer = { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "Adres", marks: [{ type: "textStyle", attrs: { fontSize: "7pt" } }] }] }] };
    const k = kural(d, "yaziTipi");
    expect(k.durum).toBe("hata");
    const ileti = k.bulgular.map((b) => `${b.durum}: ${b.ileti}`);
    expect(ileti).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/^hata: “Georgia” yazı tipi/),
        "hata: Yazı alanında 9 puntodan küçük metin var (m.7/1).",
        expect.stringMatching(/^uyari: 1 metin parçası farklı puntoda/),
        "uyari: Belgede hem Times New Roman hem Arial var; tek yazı tipi kullanın.",
      ]),
    );
    const alt = ust();
    alt.settings.footer = d.settings.footer;
    expect(kural(alt, "yaziTipi").bulgular).toEqual([expect.objectContaining({ durum: "hata", bolge: "altBilgi", ileti: "İletişim bilgilerinde 8 puntodan küçük metin var (m.7/1)." })]);
    // Başlık stilleri tema yazı tipini taşır.
    const baslikli = ust();
    baslikli.content.content!.push({ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Başlık" }] });
    expect(kural(baslikli, "yaziTipi").bulgular[0].ileti).toMatch(/“Aptos Display” yazı tipi kullanılmış/);
  });
});

describe("Kural 3: konu satırı (m.13)", () => {
  it("olumlu: baş harfler büyük, bağlaçlar küçük, sonda noktalama yok, sayı satırının altında", () => {
    expect(kural(ust({ konu: "Eğitim ve Öğretim Programı ile İlgili Görüş" }), "konu").durum).toBe("uygun");
    expect(kural(ust({ konu: "2027 Yılı e-Devlet Hizmetleri" }), "konu").durum).toBe("uygun");
  });
  it("olumsuz: sonda nokta, küçük harfle başlayan kelimeler, boş konu, sayıdan ayrı konu", () => {
    expect(kural(ust({ konu: "Eğitim Programı." }), "konu").bulgular[0]).toMatchObject({ durum: "hata", ileti: expect.stringMatching(/sonunda noktalama işareti \(“\.”\)/) });
    expect(kural(ust({ konu: "Eğitim Programı Hk." }), "konu").durum).toBe("hata");
    expect(kural(ust({ konu: "eğitim programı ve takvimi" }), "konu").bulgular[0].ileti).toBe("Konudaki kelimelerin baş harfleri büyük yazılır: “eğitim”, “programı”, “takvimi” (m.13/1).");
    expect(kural(ust({ konu: "Ve Program" }), "konu").durum).toBe("uygun");
    expect(kural(ust({ konu: "ve Program" }), "konu").durum).toBe("hata");
    expect(kural(ust({ konu: "" }), "konu").bulgular[0].ileti).toBe("Konu boş.");
    const ayri = ust();
    const i = ayri.content.content!.findIndex((n) => n.content?.[0]?.text?.startsWith("Konu\t"));
    ayri.content.content!.splice(i, 0, p(""), p("Ara satır"));
    expect(kural(ayri, "konu").bulgular[0]).toMatchObject({ durum: "uyari", ileti: "Konu, “Sayı:” satırının bir alt satırına yazılır (m.13/1)." });
  });
  it("konu satırı yoksa bilgi verir", () => {
    expect(kural(resmiBelge("resmi-dilekce", FORM), "konu").durum).toBe("bilgi");
  });
});

describe("Kural 4: sayıların yazımı (m.16/7)", () => {
  it("olumlu: 1.452; 25.126; 10.545,72; yıl, kanun numarası, posta kodu, telefon, saat, dosya kodu", () => {
    for (const s of ["1.452", "25.126", "326.197", "10.545,72", "1,5", "999"]) expect(sayiDuzelt(s), s).toBeNull();
    expect(sayiDuzelt("2027", " yılında")).toBeNull();
    expect(sayiDuzelt("3071", " sayılı Kanun")).toBeNull();
    expect(sayiDuzelt("06100")).toBeNull();
    expect(sayiDuzelt("12345678901")).toBeNull();
    expect(sayiDuzelt("14.30")).toBeNull();
    expect(sayiDuzelt("903.07.02")).toBeNull();
    const d = ust({
      metin: "2027 yılında 5442 sayılı Kanun kapsamında 1.452 kişiye 10.545,72 TL ödenmiştir. Saat 14.30'da E-12345678-903.07.02-4752 sayılı yazı ile 06100 posta kodlu adrese gönderilmiştir.",
    });
    expect(kural(d, "sayiBicimi").durum).toBe("uygun");
  });
  it("olumsuz: gruplanmamış, İngilizce ayırıcılı ve ondalık noktalı sayılar düzeltme önerisiyle", () => {
    expect(sayiDuzelt("12500")).toBe("12.500");
    expect(sayiDuzelt("1452", " adet")).toBe("1.452");
    expect(sayiDuzelt("1452", " TL")).toBe("1.452");
    expect(sayiDuzelt("10545,72")).toBe("10.545,72");
    expect(sayiDuzelt("10,545.72")).toBe("10.545,72");
    expect(sayiDuzelt("10545.72")).toBe("10.545,72");
    expect(sayiDuzelt("1250000")).toBe("1.250.000");
    const d = ust({ metin: "Toplam 1452 adet malzeme için 10545,72 TL ve 1,250,000.50 TL ödenmiştir." });
    const k = kural(d, "sayiBicimi");
    expect(k.durum).toBe("hata");
    expect(k.bulgular.map((b) => b.ileti)).toEqual([
      "“1452” → “1.452” biçiminde yazılır (m.16/7).",
      "“10545,72” → “10.545,72” biçiminde yazılır (m.16/7).",
      "“1,250,000.50” → “1.250.000,50” biçiminde yazılır (m.16/7).",
    ]);
    expect(k.bulgular[0].paragraf).toBeGreaterThan(0);
  });
});

describe("Kural 5: tarih biçimi (m.12)", () => {
  it("olumlu: GG.AA.YYYY ve '10 Ekim 2019'; tarih sayı satırının en sağında", () => {
    expect(tarihDuzelt("01.09.2019")).toBeNull();
    const d = ust({ metin: "01.09.2019 tarihli ve 10 Ekim 2019 tarihli yazılar incelenmiştir." });
    expect(kural(d, "tarihBicimi").durum).toBe("uygun");
  });
  it("olumsuz: tek haneli gün/ay, eğik çizgi, tire, iki haneli yıl, küçük harfli ay", () => {
    expect(tarihDuzelt("1.9.2019")).toEqual({ oneri: "01.09.2019" });
    expect(tarihDuzelt("01/09/2019")).toEqual({ oneri: "01.09.2019" });
    expect(tarihDuzelt("2019-09-01")).toEqual({ oneri: "01.09.2019" });
    expect(tarihDuzelt("01.09.19")).toEqual({ oneri: null });
    const d = ust({ metin: "1.9.2019 ve 01/09/2019 ile 2019-09-01 ve 01.09.19 ile 10 ekim 2019 tarihli yazılar." });
    const k = kural(d, "tarihBicimi");
    expect(k.durum).toBe("hata");
    expect(k.bulgular.map((b) => b.ileti)).toEqual([
      "“1.9.2019” tarihi gün ve ay iki, yıl dört haneli, aralarında nokta ile yazılır: “01.09.2019” (m.12/1).",
      "“01/09/2019” tarihi gün ve ay iki, yıl dört haneli, aralarında nokta ile yazılır: “01.09.2019” (m.12/1).",
      "“2019-09-01” tarihi gün ve ay iki, yıl dört haneli, aralarında nokta ile yazılır: “01.09.2019” (m.12/1).",
      "“01.09.19” tarihi gün ve ay iki, yıl dört haneli, aralarında nokta ile yazılır (m.12/1).",
      "Ay adı büyük harfle başlar: “10 Ekim 2019” (m.12/1).",
    ]);
  });
  it("olumsuz: tarih sayıdan ayrı satırda ya da en sağda değil; tarih yoksa bilgi", () => {
    const ayri = degistir(ust(), /^Sayı\t/, "Sayı\t:\tE-12345678-903.07.02-1");
    ayri.content.content!.splice(3, 0, p("15.10.2026", { textAlign: "right" }));
    expect(kural(ayri, "tarihBicimi").bulgular[0]).toMatchObject({ durum: "hata", ileti: "Tarih, sayı ile aynı satırda ve yazı alanının en sağında yazılır (m.12/1)." });
    const yanyana = degistir(ust(), /^Sayı\t/, "Sayı\t:\tE-12345678-903.07.02-1 15.10.2026");
    expect(kural(yanyana, "tarihBicimi").bulgular[0].durum).toBe("uyari");
    const yok = degistir(ust(), /^Sayı\t/, "Sayı\t:\tE-12345678-903.07.02-1");
    expect(kural(yok, "tarihBicimi").durum).toBe("bilgi");
  });
});

describe("Kural 6: arz / rica uyumu (m.16/12)", () => {
  const kapanis = (k: string, form = {}) => ust({ kapanis: k, ...form });
  it("olumlu: üst makama arz, alt makama rica, karma dağıtımda arz ve rica, gerçek kişiye Saygılarımla", () => {
    expect(kural(kapanis("Bilgilerinizi ve gereğini arz ederim."), "kapanis", "ust").durum).toBe("uygun");
    expect(kural(kapanis("Gereğini rica ederim."), "kapanis", "alt").durum).toBe("uygun");
    expect(kural(kapanis("Bilgilerinizi ve gereğini arz ve rica ederim."), "kapanis", "karma").durum).toBe("uygun");
    expect(kural(kapanis("Bilgilerinizi ve gereğini arz/rica ederim."), "kapanis", "karma").durum).toBe("uygun");
    expect(kural(kapanis("Saygılarımla.", { muhatap: "Sayın Ayşe YILMAZ", muhatapBirim: "" }), "kapanis").durum).toBe("uygun");
    // Otomatik: MAKAMINA → üst makam.
    expect(kural(resmiBelge("resmi-olur", FORM), "kapanis").durum).toBe("uygun");
  });
  it("olumsuz: düzeyle uyuşmayan kapanış, noktasız kapanış, kapanışsız yazı", () => {
    expect(kural(kapanis("Gereğini rica ederim."), "kapanis", "ust").bulgular[0]).toMatchObject({
      durum: "hata",
      ileti: expect.stringMatching(/^Muhatap üst ya da aynı düzey makam: kapanış “… arz ederim.” olmalı/),
    });
    expect(kural(kapanis("Gereğini arz ederim."), "kapanis", "alt").durum).toBe("hata");
    expect(kural(kapanis("Gereğini arz ederim."), "kapanis", "karma").durum).toBe("hata");
    expect(kural(kapanis("Gereğini arz ederim.", { muhatap: "Sayın Ayşe YILMAZ" }), "kapanis").durum).toBe("uyari");
    expect(kural(kapanis("Olurlarınıza rica ederim.", { muhatap: "VALİLİK MAKAMINA" }), "kapanis").durum).toBe("hata");
    expect(kural(kapanis("Gereğini arz ederim"), "kapanis", "ust").bulgular[0]).toMatchObject({ durum: "uyari", ileti: "Kapanış ibaresi nokta ile biter (“… arz ederim.”)." });
    expect(kural(kapanis(""), "kapanis").bulgular[0]).toMatchObject({ durum: "uyari", ileti: expect.stringMatching(/kapanış ibaresiyle bitmiyor/) });
  });
  it("olur yazısında ara makamın 'Uygun görüşle arz ederim.' satırı kapanış sayılmaz", () => {
    const olur = resmiBelge("resmi-olur", { ...FORM, kapanis: "Olurlarınıza rica ederim." });
    expect(kural(olur, "kapanis", "ust").bulgular[0].ileti).toMatch(/belgede “Olurlarınıza rica ederim\.”/);
  });
  it("düzey belirsizse ya da dağıtımlıysa bilgi verir", () => {
    expect(kural(kapanis("Gereğini arz ederim."), "kapanis").durum).toBe("bilgi");
    expect(kural(resmiBelge("resmi-dagitimli", FORM), "kapanis").bulgular[0].ileti).toMatch(/^Dağıtımlı yazı/);
  });
});

describe("Kural 7: paragraf düzeni (m.16/4)", () => {
  it("olumlu: 1,25 cm girinti, iki yana yaslı, aralıksız", () => {
    expect(kural(ust(), "paragraf").durum).toBe("uygun");
  });
  it("olumsuz: girintisiz, sola yaslı ve aralıklı paragraflar", () => {
    const uzun = "Bu paragraf yönetmeliğe uygun olmayan biçimde girintisiz ve sola yaslı yazılmış uzun bir metindir.";
    const d = ust();
    d.content.content!.push(p(uzun), p(uzun, { indentFirst: GIRINTI_PT, textAlign: "justify", spaceAfter: 8 }));
    const k = kural(d, "paragraf");
    expect(k.bulgular.map((b) => b.ileti)).toEqual([
      "1 paragraf: ilk satır 1,25 cm içeriden başlamalı (m.16/4).",
      "1 paragraf: metin iki yana yaslanmalı (m.16/4).",
      "1 paragraf: paragraflar arasında satır boşluğu bırakılmaz (m.16/4).",
    ]);
    // Normal stilde paragraf sonrası boşluk (varsayılan 8 nk) de sayılır.
    const varsayilan = { settings: defaultSettings(), content: { type: "doc", content: [p(uzun, { indentFirst: GIRINTI_PT, textAlign: "justify" })] } };
    expect(kural(varsayilan, "paragraf").bulgular[0].ileti).toMatch(/paragraflar arasında satır boşluğu/);
  });
});
