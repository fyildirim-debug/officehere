// Dosya > Yeni şablonları (OfficeHere'in kendi tasarımları, Türkçe içerik).
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, type BelgeDocument } from "../model/document";
import { bekleyenAlanlar } from "./resmi/alanlar";
import { ONIZLEMELER, RESMI_KUME_ONIZLEME } from "./resmi/onizleme";
import { RESMI_SABLONLAR, resmiBelge } from "./resmi/sablonlar";

const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const b = (text: string) => t(text, [{ type: "bold" }]);
const col = (text: string, color: string, extra: Record<string, unknown> = {}) => t(text, [{ type: "textStyle", attrs: { color, ...extra } }]);
const p = (content?: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({
  type: "paragraph",
  attrs,
  ...(content ? { content: typeof content === "string" ? [t(content)] : content } : {}),
});
const h = (level: number, text: string): JSONContent => ({ type: "heading", attrs: { level }, content: [t(text)] });
const ul = (...items: (string | JSONContent[])[]): JSONContent => ({ type: "bulletList", content: items.map((i) => ({ type: "listItem", content: [p(i)] })) });
const ol = (...items: string[]): JSONContent => ({ type: "orderedList", attrs: { start: 1 }, content: items.map((i) => ({ type: "listItem", content: [p(i)] })) });
const cell = (content: string | JSONContent[], attrs: Record<string, unknown> = {}): JSONContent => ({ type: "tableCell", attrs, content: [p(content)] });
const row = (...cells: JSONContent[]): JSONContent => ({ type: "tableRow", content: cells });

const today = () => new Date().toLocaleDateString("tr-TR", { day: "numeric", month: "long", year: "numeric" });

export interface TemplateDef {
  id: string;
  label: string;
  build: () => BelgeDocument;
  preview: React.ReactNode;
  /** Küme şablonu: Dosya › Yeni'de ayrı kutucuk olarak değil, kümenin galerisinde gösterilir. */
  kume?: "resmi";
}

/** Dosya › Yeni'deki "Resmî yazışma" kutucuğu: tıklanınca şablon galerisi ve alan formu açılır. */
export const RESMI_KUME_ID = "resmi";

function letter(): BelgeDocument {
  const s = defaultSettings();
  s.properties.title = "Mektup";
  return {
    settings: s,
    content: {
      type: "doc",
      content: [
        p([col("Adınız Soyadınız", "#0F4761", { fontSize: "20pt", fontFamily: "Aptos Display" })], { spaceAfter: 0 }),
        p([col("Sokak Adresi • İlçe / İl • Telefon • e-posta@ornek.com", "#595959")], { borders: { bottom: { style: "single", size: 0.75, color: "#0F4761", space: 4 } }, spaceAfter: 24 }),
        p(today(), { textAlign: "right" }),
        p([b("Alıcı Adı Soyadı")], { spaceAfter: 0 }),
        p("Unvan", { spaceAfter: 0 }),
        p("Kurum Adı", { spaceAfter: 0 }),
        p("Adres", { spaceAfter: 18 }),
        p("Sayın Alıcı Adı,"),
        p("Bu mektubu, geçen hafta yaptığımız görüşmenin ardından konuştuğumuz konulardaki düşüncelerimi paylaşmak amacıyla yazıyorum. Mektubunuzun gövdesine buradan başlayabilirsiniz; yalnızca bu metni seçip yazmaya başlamanız yeterli.", { textAlign: "justify" }),
        p("Belgenin görünümünü Tasarım sekmesindeki temalarla değiştirebilir, paragraf aralıklarını Düzen sekmesinden ayarlayabilirsiniz. Her şeyin istediğiniz gibi göründüğünden emin olduktan sonra Dosya > Yazdır ile çıktı alabilir ya da PDF olarak kaydedebilirsiniz.", { textAlign: "justify" }),
        p("Değerli zamanınızı ayırdığınız için teşekkür eder, en kısa sürede görüşmeyi dilerim.", { textAlign: "justify", spaceAfter: 24 }),
        p("Saygılarımla,", { spaceAfter: 36 }),
        p([b("Adınız Soyadınız")], { spaceAfter: 0 }),
        p([col("Unvanınız", "#595959")]),
      ],
    },
  };
}

function resume(): BelgeDocument {
  const s = defaultSettings();
  s.properties.title = "Özgeçmiş";
  s.page.margin = { ...s.page.margin, top: 1134, bottom: 1134, left: 1134, right: 1134 };
  const section = (text: string): JSONContent => ({
    type: "heading",
    attrs: { level: 2, borders: { bottom: { style: "single", size: 0.75, color: "#156082", space: 2 } }, spaceBefore: 14 },
    content: [t(text.toLocaleUpperCase("tr-TR"), [{ type: "textStyle", attrs: { letterSpacing: 1, fontSize: "13pt" } }])],
  });
  return {
    settings: s,
    content: {
      type: "doc",
      content: [
        p([col("AYŞE YILMAZ", "#0E2841", { fontSize: "28pt", fontFamily: "Aptos Display", letterSpacing: 2 })], { spaceAfter: 0 }),
        p([col("Kıdemli Yazılım Mühendisi", "#156082", { fontSize: "14pt" })], { spaceAfter: 4 }),
        p([col("İstanbul • +90 5xx xxx xx xx • ayse.yilmaz@ornek.com • linkedin.com/in/ayseyilmaz", "#595959", { fontSize: "10pt" })]),
        section("Profil"),
        p("Ölçeklenebilir web uygulamaları geliştirmede 8 yıllık deneyime sahip, ekip liderliği ve mentorluk yapmış bir yazılım mühendisi. Kullanıcı odaklı ürünler tasarlamayı ve teknik borcu azaltan sürdürülebilir mimariler kurmayı seviyorum.", { textAlign: "justify" }),
        section("Deneyim"),
        p([b("Kıdemli Yazılım Mühendisi"), t(" — Örnek Teknoloji A.Ş.")], { spaceAfter: 0 }),
        p([col("Ocak 2021 – Günümüz, İstanbul", "#595959", { fontSize: "10pt" })], { spaceAfter: 4 }),
        ul("Ödeme altyapısını mikro hizmetlere taşıyan 6 kişilik ekibe liderlik ettim.", "Sayfa yükleme sürelerini %40 kısaltan önbellekleme katmanını tasarladım.", "Yeni katılan geliştiriciler için eğitim programı oluşturdum."),
        p([b("Yazılım Mühendisi"), t(" — Deneme Yazılım Ltd.")], { spaceAfter: 0 }),
        p([col("Haziran 2017 – Aralık 2020, Ankara", "#595959", { fontSize: "10pt" })], { spaceAfter: 4 }),
        ul("Kamu kurumları için belge yönetim sistemi geliştirdim.", "Otomatik test kapsamını %35'ten %85'e çıkardım."),
        section("Eğitim"),
        p([b("Bilgisayar Mühendisliği, Lisans"), t(" — Örnek Üniversitesi")], { spaceAfter: 0 }),
        p([col("2013 – 2017 • Genel not ortalaması 3,45 / 4,00", "#595959", { fontSize: "10pt" })]),
        section("Yetenekler"),
        {
          type: "table",
          attrs: { tableStyle: "none", headerRow: false, bandedRows: false },
          content: [
            row(cell([b("Diller: "), t("TypeScript, Go, Python, SQL")]), cell([b("Araçlar: "), t("Docker, Kubernetes, Git")])),
            row(cell([b("Yabancı dil: "), t("İngilizce (C1), Almanca (A2)")]), cell([b("Diğer: "), t("Çevik yöntemler, sistem tasarımı")])),
          ],
        },
        p(),
      ],
    },
  };
}

function report(): BelgeDocument {
  const s = defaultSettings();
  s.properties.title = "Proje Raporu";
  s.footer = { type: "doc", content: [p([t("Proje Raporu  •  Sayfa "), { type: "pageField", attrs: { field: "PAGE" } }, t(" / "), { type: "pageField", attrs: { field: "NUMPAGES" } }], { textAlign: "right" })] };
  return {
    settings: s,
    content: {
      type: "doc",
      content: [
        p([t("2026 Yılı Dijital Dönüşüm Projesi")], { styleId: "Title" }),
        p([t("Dönem Sonu Değerlendirme Raporu")], { styleId: "Subtitle" }),
        p([col(`Hazırlayan: Proje Yönetim Ofisi  •  ${today()}`, "#595959")], { spaceAfter: 18 }),
        h(1, "Yönetici Özeti"),
        p("Bu rapor, projenin ilk altı aylık döneminde elde edilen sonuçları, karşılaşılan riskleri ve önümüzdeki dönem için önerilen adımları özetlemektedir. Hedeflenen kilometre taşlarının büyük bölümü zamanında tamamlanmış, bütçe planlanan sınırlar içinde kalmıştır.", { textAlign: "justify" }),
        h(1, "Hedefler ve Sonuçlar"),
        p("Dönem başında belirlenen hedefler ve gerçekleşme durumları aşağıdaki tabloda özetlenmiştir."),
        {
          type: "table",
          attrs: { tableStyle: "grid-blue", headerRow: true, bandedRows: true },
          content: [
            row(cell([b("Hedef")]), cell([b("Plan")]), cell([b("Gerçekleşen")]), cell([b("Durum")])),
            row(cell("Altyapı geçişi"), cell("%100"), cell("%100"), cell("Tamamlandı")),
            row(cell("Personel eğitimi"), cell("250 kişi"), cell("212 kişi"), cell("Devam ediyor")),
            row(cell("Süreç otomasyonu"), cell("12 süreç"), cell("9 süreç"), cell("Devam ediyor")),
          ],
        },
        p(),
        h(2, "Öne Çıkan Başarılar"),
        ul("Veri merkezi taşıması kesintisiz tamamlandı.", "Müşteri başvuru süresi ortalama 3 günden 1 güne indi.", "Kâğıt kullanımı geçen yılın aynı dönemine göre %60 azaldı."),
        h(1, "Riskler"),
        p("Önümüzdeki dönemde dikkat edilmesi gereken başlıca riskler şunlardır:"),
        ol("Tedarikçi teslim sürelerindeki gecikmeler", "Eğitim takviminin yoğun iş dönemleriyle çakışması", "Bilgi güvenliği denetimlerinin ek iş yükü"),
        h(1, "Sonuç ve Öneriler"),
        p("Projenin genel gidişatı olumludur. Eğitim programının çevrimiçi modüllerle desteklenmesi ve otomasyon kapsamının önceliklendirilmesi önerilmektedir.", { textAlign: "justify" }),
      ],
    },
  };
}

function meeting(): BelgeDocument {
  const s = defaultSettings();
  s.properties.title = "Toplantı Notları";
  return {
    settings: s,
    content: {
      type: "doc",
      content: [
        p([col("TOPLANTI NOTLARI", "#156082", { fontSize: "11pt", letterSpacing: 2 })], { spaceAfter: 0 }),
        p([t("Haftalık Ürün Değerlendirme Toplantısı")], { styleId: "Title" }),
        {
          type: "table",
          attrs: { tableStyle: "light-blue", headerRow: false, bandedRows: false },
          content: [
            row(cell([b("Tarih")]), cell(today())),
            row(cell([b("Saat / Yer")]), cell("10:00 – 11:00, Toplantı Odası 3")),
            row(cell([b("Katılımcılar")]), cell("Ahmet K., Elif D., Mert S., Zeynep A.")),
            row(cell([b("Not alan")]), cell("Zeynep A.")),
          ],
        },
        p(),
        h(2, "Gündem"),
        ol("Geçen haftanın aksiyonlarının gözden geçirilmesi", "Yeni sürüm takvimi", "Kullanıcı geri bildirimleri", "Diğer konular"),
        h(2, "Görüşülenler"),
        p([b("Sürüm takvimi: "), t("2.4 sürümünün test süreci uzadığı için yayın tarihi bir hafta ileri alındı.")]),
        p([b("Geri bildirimler: "), t("Mobil uygulamada arama hızıyla ilgili şikâyetler arttı; inceleme başlatılacak.")]),
        h(2, "Kararlar"),
        ul("2.4 sürümü 24 Ekim'de yayınlanacak.", "Arama performansı için ayrı bir çalışma grubu kurulacak."),
        h(2, "Aksiyonlar"),
        {
          type: "table",
          attrs: { tableStyle: "grid-blue", headerRow: true, bandedRows: true },
          content: [
            row(cell([b("Görev")]), cell([b("Sorumlu")]), cell([b("Son tarih")])),
            row(cell("Test planını güncelle"), cell("Elif D."), cell("18.10.2026")),
            row(cell("Arama performans analizi"), cell("Mert S."), cell("21.10.2026")),
            row(cell("Sürüm notlarını hazırla"), cell("Ahmet K."), cell("23.10.2026")),
          ],
        },
        p(),
        p([col("Bir sonraki toplantı: Gelecek hafta aynı saatte.", "#595959")], { styleId: null }),
      ],
    },
  };
}

function petition(): BelgeDocument {
  const s = defaultSettings();
  s.properties.title = "Dilekçe";
  return {
    settings: s,
    content: {
      type: "doc",
      content: [
        p("", { spaceAfter: 24 }),
        p([b("…………………… MÜDÜRLÜĞÜNE")], { textAlign: "center", spaceAfter: 0 }),
        p([b("…………………… / ……………")], { textAlign: "center", spaceAfter: 36 }),
        p("Kurumunuzun ………………………… biriminde ………………… olarak görev yapmaktayım. ………………………… nedeniyle …………………… tarihleri arasında yıllık izin kullanmak istiyorum.", { indentFirst: 35.4, textAlign: "justify", lineHeight: 1.5 }),
        p("Gereğinin yapılmasını saygılarımla arz ederim.", { indentFirst: 35.4, spaceAfter: 36 }),
        p(today(), { textAlign: "right", spaceAfter: 0 }),
        p("Ad Soyad", { textAlign: "right", spaceAfter: 0 }),
        p("İmza", { textAlign: "right", spaceAfter: 48 }),
        p([b("Adres: ")], { spaceAfter: 0 }),
        p("…………………………………………………………", { spaceAfter: 0 }),
        p([b("Telefon: "), t("…………………………")], { spaceAfter: 0 }),
        p([b("T.C. Kimlik No: "), t("…………………………")]),
      ],
    },
  };
}

const lines = (n: number, cls = "") => Array.from({ length: n }, (_, i) => <i key={i} className={cls} />);

export const TEMPLATES: TemplateDef[] = [
  { id: "blank", label: "Boş belge", build: () => ({ settings: defaultSettings(), content: { type: "doc", content: [{ type: "paragraph" }] } }), preview: <span className="bl-tpl" /> },
  { id: RESMI_KUME_ID, label: "Resmî yazışma", build: () => resmiBelge("resmi-ust-yazi", bekleyenAlanlar("resmi-ust-yazi")), preview: RESMI_KUME_ONIZLEME },
  {
    id: "letter",
    label: "Mektup",
    build: letter,
    preview: (
      <span className="bl-tpl is-letter">
        <b />
        <u />
        {lines(9)}
      </span>
    ),
  },
  {
    id: "resume",
    label: "Özgeçmiş",
    build: resume,
    preview: (
      <span className="bl-tpl is-resume">
        <b />
        <em />
        {lines(3)}
        <u />
        {lines(4)}
        <u />
        {lines(3)}
      </span>
    ),
  },
  {
    id: "report",
    label: "Rapor",
    build: report,
    preview: (
      <span className="bl-tpl is-report">
        <b />
        <em />
        {lines(3)}
        <s />
        {lines(4)}
      </span>
    ),
  },
  {
    id: "meeting",
    label: "Toplantı Notları",
    build: meeting,
    preview: (
      <span className="bl-tpl is-meeting">
        <em />
        <b />
        <s />
        {lines(3)}
        <s />
      </span>
    ),
  },
  {
    id: "petition",
    label: "Dilekçe",
    build: petition,
    preview: (
      <span className="bl-tpl is-petition">
        <b />
        {lines(5)}
        <em />
      </span>
    ),
  },
  // Resmî Yazışma (2020) kümesi (templates/resmi): form değerleri bekleyenAlanlar ile bir kez alınır.
  ...RESMI_SABLONLAR.map((r): TemplateDef => ({ id: r.id, label: r.etiket, kume: "resmi", build: () => resmiBelge(r.id, bekleyenAlanlar(r.id)), preview: ONIZLEMELER[r.id] })),
];
