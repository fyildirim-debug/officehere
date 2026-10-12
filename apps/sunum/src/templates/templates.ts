/** Dosya > Yeni şablonları: boş sunu ve temalardan üretilmiş hazır sunular (Türkçe içerik). */
import { makeSp, makeTable, textBody, uid } from "../model/factory";
import { blankPresentation, newPresentation, newSlide } from "../model/presentation";
import type { Animation, LayoutType, Presentation, Shape, Slide, SpShape, TableShape } from "../model/types";

export interface TemplateDef {
  id: string;
  label: string;
  build: () => Presentation;
}

function setPh(slide: Slide, match: (s: SpShape) => boolean, text: string): string | undefined {
  let id: string | undefined;
  slide.shapes = slide.shapes.map((s) => {
    if (id || s.kind !== "sp" || !s.ph || !match(s)) return s;
    id = s.id;
    return { ...s, text: { ...(s.text ?? { paragraphs: [] }), ...textBody(text) } };
  });
  return id;
}

const isTitle = (s: SpShape) => s.ph?.type === "title" || s.ph?.type === "ctrTitle";
const isSub = (s: SpShape) => s.ph?.type === "subTitle" || (s.ph?.type === "body" && s.ph.idx === 1);
const isIdx = (i: number) => (s: SpShape) => s.ph?.idx === i && !isTitle(s);

function add(p: Presentation, layout: LayoutType, title: string, ...bodies: string[]): Slide {
  const s = newSlide(p, layout);
  setPh(s, isTitle, title);
  bodies.forEach((b, i) => setPh(s, isIdx(i + 1), b));
  p.slides.push(s);
  return s;
}

/** İçerik yer tutucusundaki paragrafları numaralı listeye çevirir. */
function numbered(s: Slide): Slide {
  s.shapes = s.shapes.map((x) =>
    x.kind === "sp" && x.ph?.type === "obj" && x.text
      ? { ...x, text: { ...x.text, paragraphs: x.text.paragraphs.map((pp) => ({ ...pp, props: { ...pp.props, bullet: { type: "auto", scheme: "arabicPeriod" }, marL: 54, indent: -54 } })) } }
      : x,
  );
  return s;
}

function titleSlide(p: Presentation, title: string, sub: string) {
  const s = p.slides[0];
  setPh(s, isTitle, title);
  setPh(s, isSub, sub);
}

function anim(shapeId: string | undefined, effect: string, trigger: Animation["trigger"] = "click", option?: string, cls: Animation["cls"] = "entr"): Animation[] {
  return shapeId ? [{ id: uid("a"), shapeId, cls, effect, option, trigger, duration: 0.5, delay: 0 }] : [];
}

function tableOf(rows: string[][], x: number, y: number, w: number, styleId = "medium2-accent1"): TableShape {
  const t = makeTable(rows.length, rows[0].length, x, y, w, 44);
  t.style.styleId = styleId;
  t.rows.forEach((r, i) => r.cells.forEach((c, j) => (c.text = textBody(rows[i][j] ?? ""))));
  return t;
}

function chevrons(labels: string[], y: number, slideW: number): Shape[] {
  const n = labels.length;
  const gap = 8;
  const w = (slideW - 176 - gap * (n - 1)) / n;
  return labels.map((l, i) =>
    makeSp(i === 0 ? "homePlate" : "chevron", 88 + i * (w + gap), y, w, 90, {
      name: `Aşama ${i + 1}`,
      fill: { type: "solid", color: { scheme: (["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const)[i % 6] } },
      line: { fill: { type: "none" }, width: 1, dash: "solid" },
      text: { ...textBody(l), anchor: "ctr", levels: [{ run: { color: { scheme: "lt1" }, size: 18, bold: true }, para: { align: "ctr" } }] },
    }),
  );
}

function projectDeck(): Presentation {
  const p = newPresentation("kiyi", 0);
  p.props.title = "Proje Sunumu";
  titleSlide(p, "Müşteri Portalı Yenileme Projesi", "Proje başlangıç sunumu · Proje Yönetim Ofisi");
  p.slides[0].transition = { type: "fade", duration: 1, advanceOnClick: true };
  const g = add(p, "obj", "Gündem", "Projenin amacı ve kapsamı\nHedefler ve başarı ölçütleri\nZaman çizelgesi\nEkip ve roller\nRiskler ve önlemler\nSorular");
  g.transition = { type: "push", option: "u", duration: 0.75, advanceOnClick: true };
  const s3 = add(p, "twoObj", "Hedefler", "İş hedefleri\n\tMüşteri memnuniyetini %20 artırmak\n\tÇağrı merkezi yükünü azaltmak\n\tSelf servis işlemleri yaygınlaştırmak", "Teknik hedefler\n\tMobil uyumlu arayüz\n\tSayfa yükleme süresi < 2 sn\n\tErişilebilirlik (WCAG 2.1 AA)");
  s3.transition = { type: "push", option: "u", duration: 0.75, advanceOnClick: true };
  const s4 = add(p, "titleOnly", "Zaman Çizelgesi");
  const ch = chevrons(["Analiz", "Tasarım", "Geliştirme", "Test", "Yayın"], 230, p.width);
  s4.shapes.push(...ch);
  s4.shapes.push(tableOf([["Aşama", "Başlangıç", "Bitiş", "Sorumlu"], ["Analiz", "06.01.2027", "31.01.2027", "İş Analizi"], ["Tasarım", "01.02.2027", "28.02.2027", "UX Ekibi"], ["Geliştirme", "01.03.2027", "30.05.2027", "Yazılım"], ["Test ve Yayın", "01.06.2027", "30.06.2027", "Kalite Güvence"]], 88, 360, p.width - 176));
  s4.animations = ch.flatMap((c, i) => anim(c.id, "wipe", i === 0 ? "click" : "after", "l"));
  s4.transition = { type: "push", option: "u", duration: 0.75, advanceOnClick: true };
  const s5 = add(p, "obj", "Ekip ve Roller", "Proje Sponsoru – Genel Müdür Yardımcısı\nProje Yöneticisi – PYO\nÜrün Sahibi – Dijital Kanallar\nGeliştirme Ekibi – 6 kişi\nKalite Güvence – 2 kişi");
  s5.notes = "Ekip üyelerini tek tek tanıtın; sorumluluk matrisi ekte paylaşılacak.";
  const s6 = add(p, "obj", "Riskler ve Önlemler", "Kapsam kayması → değişiklik kontrol kurulu\nEntegrasyon gecikmeleri → erken prototip\nKaynak eksikliği → dış kaynak planı");
  s6.transition = { type: "push", option: "u", duration: 0.75, advanceOnClick: true };
  const end = add(p, "secHead", "Teşekkürler", "Sorularınız için: pyo@ornek.com.tr");
  end.transition = { type: "fade", duration: 1, advanceOnClick: true };
  return p;
}

function trainingDeck(): Presentation {
  const p = newPresentation("kagit", 0);
  p.props.title = "Eğitim";
  titleSlide(p, "Etkili Sunum Teknikleri", "Kurum İçi Eğitim Programı · Modül 1");
  add(p, "obj", "Eğitimin Amaçları", "Hedef kitleyi tanımak\nMesajı yapılandırmak\nGörsel tasarım ilkelerini uygulamak\nBeden dili ve ses kullanımı\nSorularla etkileşim kurmak");
  numbered(add(p, "obj", "Modüller", "Hazırlık ve planlama\nİçerik ve hikâye kurgusu\nSlayt tasarımı\nSunum anı\nDeğerlendirme"));
  add(p, "twoTxTwoObj", "Yapılması ve Kaçınılması Gerekenler", "Yapın", "Her slaytta tek bir ana fikir\nOkunaklı yazı tipi boyutu (en az 24 nk)\nGörsellerle destekleyin", "Kaçının", "Uzun paragraflar\nÇok sayıda yazı tipi\nOkunamayan renk kontrastları");
  add(p, "secHead", "Uygulama", "Her katılımcı 3 dakikalık bir sunum hazırlayacak");
  const q = numbered(add(p, "objTx", "Kısa Sınav", "Bir slaytta kaç ana fikir olmalı?\nÖnerilen en küçük yazı boyutu nedir?\nSunumun ilk 30 saniyesinde ne yapılmalı?", "Cevaplarınızı not alın; oturum sonunda birlikte değerlendireceğiz."));
  q.notes = "Cevaplar: 1) Bir 2) 24 nk 3) Dikkat çekici bir açılış";
  add(p, "obj", "Özet", "Hazırlık, sunumun yarısıdır\nAz metin, güçlü görsel\nProva yapın ve süreyi ölçün");
  return p;
}

function productDeck(): Presentation {
  const p = newPresentation("geceyarisi", 0);
  p.props.title = "Ürün Tanıtımı";
  titleSlide(p, "Nova X1 Akıllı Saat", "Yeni nesil sağlık ve fitness asistanı");
  p.slides[0].transition = { type: "zoom", option: "in", duration: 1, advanceOnClick: true };
  const f = add(p, "titleOnly", "Öne Çıkan Özellikler");
  const items: [string, string][] = [
    ["heart", "7/24 nabız ve oksijen takibi"],
    ["lightningBolt", "14 gün pil ömrü"],
    ["sun", "Güneş ışığında okunabilir ekran"],
    ["star5", "100+ spor modu"],
  ];
  const shapes: Shape[] = [];
  items.forEach(([geom, label], i) => {
    const x = 120 + (i % 2) * 540;
    const y = 220 + Math.floor(i / 2) * 200;
    const icon = makeSp(geom, x, y, 90, 90, { name: `Simge ${i + 1}`, fill: { type: "solid", color: { scheme: "accent3" } }, line: { fill: { type: "none" }, width: 1, dash: "solid" } });
    const t = makeSp("rect", x + 110, y + 15, 380, 60, { name: `Özellik ${i + 1}`, txBox: true, text: { ...textBody(label), anchor: "ctr", levels: [{ run: { size: 24 } }] } });
    shapes.push(icon, t);
  });
  f.shapes.push(...shapes);
  f.animations = shapes.filter((_, i) => i % 2 === 0).flatMap((s, i) => anim(s.id, "zoom", i === 0 ? "click" : "after"));
  const spec = add(p, "titleOnly", "Teknik Özellikler");
  spec.shapes.push(tableOf([["Özellik", "Nova X1", "Nova X1 Pro"], ["Ekran", "1,4″ AMOLED", "1,6″ AMOLED"], ["Pil ömrü", "10 gün", "14 gün"], ["Su geçirmezlik", "5 ATM", "10 ATM"], ["GPS", "Dahili", "Çift frekans"], ["Fiyat", "4.999 ₺", "6.999 ₺"]], 140, 200, p.width - 280, "dark1-accent1"));
  add(p, "twoObj", "Neden Nova X1?", "Kullanıcılar için\n\tKişisel koçluk\n\tUyku analizi\n\tKolay eşleştirme", "Kurumlar için\n\tToplu yönetim paneli\n\tVeri gizliliği\n\tAPI entegrasyonu");
  const cta = add(p, "title", "Ön Siparişler Açıldı!", "nova-x1.ornek.com.tr · 1 Ekim 2027'de raflarda");
  cta.transition = { type: "shape", option: "circle", duration: 1.25, advanceOnClick: true };
  return p;
}

function meetingDeck(): Presentation {
  const p = newPresentation("cizgi", 2);
  p.props.title = "Toplantı";
  titleSlide(p, "Haftalık Ekip Toplantısı", "Pazartesi 10:00 · Toplantı Odası 3");
  add(p, "obj", "Gündem", "Geçen haftanın değerlendirmesi\nProje durumları\nKararlar\nAksiyon maddeleri\nDiğer konular");
  const st = add(p, "titleOnly", "Proje Durumları");
  st.shapes.push(tableOf([["Proje", "Durum", "İlerleme", "Not"], ["Web Sitesi", "Yolunda", "%80", "Yayın 15 Ekim"], ["Mobil Uygulama", "Riskli", "%45", "Test ortamı gecikti"], ["Veri Ambarı", "Yolunda", "%60", "—"]], 104, 200, p.width - 208, "light1-accent1"));
  add(p, "obj", "Kararlar", "Mobil test ortamı için ek bütçe onaylandı\nHaftalık demo cuma günleri yapılacak\nDokümantasyon şablonu güncellenecek");
  const ac = add(p, "titleOnly", "Aksiyon Maddeleri");
  ac.shapes.push(tableOf([["Görev", "Sorumlu", "Termin"], ["Test ortamı kurulumu", "Ayşe Y.", "18.10.2027"], ["Demo planı", "Mehmet K.", "20.10.2027"], ["Şablon güncellemesi", "Zeynep A.", "22.10.2027"]], 104, 200, p.width - 208, "medium2-accent1"));
  add(p, "secHead", "Sonraki Toplantı", "Pazartesi 10:00 · Aynı yer");
  return p;
}

function companyDeck(): Presentation {
  const p = newPresentation("geometri", 0);
  p.props.title = "Şirket Tanıtımı";
  titleSlide(p, "Anadolu Teknoloji A.Ş.", "Şirket Tanıtım Sunumu");
  add(p, "obj", "Hakkımızda", "2012'de İzmir'de kuruldu\n250+ çalışan, 4 ofis\nYazılım, danışmanlık ve eğitim hizmetleri\nISO 9001 ve ISO 27001 sertifikalı");
  const v = add(p, "twoObj", "Vizyon ve Misyon", "Vizyon\n\tBölgenin en güvenilir teknoloji ortağı olmak", "Misyon\n\tMüşterilerimizin dijital dönüşümünü uçtan uca desteklemek");
  v.transition = { type: "wipe", option: "l", duration: 1, advanceOnClick: true };
  const n = add(p, "titleOnly", "Rakamlarla Biz");
  const stats: [string, string][] = [
    ["12", "yıllık deneyim"],
    ["480+", "tamamlanan proje"],
    ["%96", "müşteri memnuniyeti"],
  ];
  stats.forEach(([num, label], i) => {
    const x = 120 + i * 360;
    n.shapes.push(
      makeSp("ellipse", x, 230, 200, 200, {
        name: `Sayı ${i + 1}`,
        fill: { type: "solid", color: { scheme: (["accent1", "accent2", "accent3"] as const)[i] } },
        line: { fill: { type: "none" }, width: 1, dash: "solid" },
        text: { ...textBody(num), anchor: "ctr", levels: [{ run: { size: 40, bold: true, color: { scheme: "lt1" } }, para: { align: "ctr" } }] },
      }),
      makeSp("rect", x - 20, 450, 240, 50, { name: `Etiket ${i + 1}`, txBox: true, text: { ...textBody(label), levels: [{ run: { size: 20 }, para: { align: "ctr" } }] } }),
    );
  });
  add(p, "title", "Birlikte Çalışalım", "info@anadoluteknoloji.com.tr · +90 232 000 00 00");
  return p;
}

export const TEMPLATES: TemplateDef[] = [
  { id: "blank", label: "Boş sunu", build: () => blankPresentation() },
  { id: "project", label: "Proje Sunumu", build: projectDeck },
  { id: "training", label: "Eğitim", build: trainingDeck },
  { id: "product", label: "Ürün Tanıtımı", build: productDeck },
  { id: "meeting", label: "Toplantı", build: meetingDeck },
  { id: "company", label: "Şirket Tanıtımı", build: companyDeck },
];
