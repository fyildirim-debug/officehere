# Hukuki çerçeve

OfficeHere, Microsoft Word, Excel ve PowerPoint ile **işlevsel ve dosya düzeyinde
uyumlu**, sıfırdan yazılmış bağımsız bir ofis paketidir. Bu belge projenin hangi
sınırlar içinde geliştirildiğini ve katkıcıların uyması gereken kuralları anlatır.

> Bu belge hukuki danışmanlık değildir. Projeyi ticari olarak dağıtmadan önce
> bir avukata danışmanız önerilir.

## Temiz oda (clean-room) ilkesi

| Serbest | Yasak |
| --- | --- |
| Ürünleri kullanıcı gibi kullanıp davranışı gözlemlemek (kara kutu) | Microsoft yazılımlarının kaynak kodunu, ikililerini incelemek; tersine mühendislik |
| Açık standartlar ve yayımlanmış belirtimler (ECMA-376, ISO/IEC 29500, [MS-OI29500] vb.) | Sızdırılmış ya da gizlilik sözleşmesiyle edinilmiş belgeler |
| Kendi çizdiğimiz arayüz, simge, şablon ve örnek içerik | Microsoft simgelerini, logolarını, şablon görsellerini, ekran görüntülerini kopyalamak |
| Genel arayüz kalıpları (sekmeli komut şeridi, durum çubuğu, sayfa görünümü) | Microsoft yazı tiplerini (Calibri, Aptos, Segoe UI…) paketle dağıtmak |

## Dosya biçimleri

- `.docx`, `.xlsx`, `.pptx` desteği **Office Open XML** standardına (ECMA-376 /
  ISO/IEC 29500) dayanır. Bu belirtimler ve Microsoft'un uzantı belirtimleri
  ([MS-DOCX], [MS-XLSX], [MS-PPTX], [MS-OI29500]) Microsoft'un *Open
  Specification Promise* kapsamındadır: Microsoft, belirtimi uygulayan
  yazılımlara karşı bu belirtimlerden doğan patent iddialarında bulunmayacağını
  taahhüt eder. Uzantı öğeleri (ör. `x14:dataBar`, `p14:sectionLst`) yalnız bu
  yayımlanmış belirtimlere göre yazılır.
- Formül sözdizimi ve fonksiyon adları birlikte çalışabilirlik için gereklidir ve
  standardın parçasıdır. Türkçe fonksiyon adları (TOPLA, EĞER, DÜŞEYARA…) kısa,
  betimleyici sözcüklerdir; başka açık kaynak tablolama yazılımlarında da kullanılır.

## Arayüz

- Tüm bileşenler (`packages/ui`) bu projede sıfırdan yazılmıştır; Microsoft'un
  arayüz kodu ya da görsel varlıkları kullanılmaz.
- Sekmeli komut şeridi (ribbon) bugün pek çok açık kaynak pakette bulunan genel
  bir kalıptır (ör. LibreOffice NotebookBar, ONLYOFFICE). Microsoft geçmişte bu
  arayüz için bir lisans programı yürütmüş ve 2004–2006 döneminde patent
  başvuruları yapmıştır; patent koruması başvurudan itibaren 20 yıl sürdüğünden
  bu başvuruların büyük bölümünün süresi dolmuş olmalıdır. Ticari dağıtımdan
  önce güncel durumu bir avukatla doğrulayın.
- Simgeler: **Fluent UI System Icons** — Microsoft'un MIT lisansıyla herkesin
  kullanımına açtığı açık kaynak simge setidir (`@fluentui/react-icons`); lisans
  metni üçüncü taraf bildirimlerinde yer alır. Office uygulamalarının kendi
  simgeleri (Word/Excel/PowerPoint logoları, şerit simgeleri) kullanılmaz.
  Uygulama ve dosya türü simgeleri bu projeye özgüdür.
- Arayüz metinleri bu projede yazılmıştır. Uzun yardım ve hata metinleri
  Microsoft ürünlerinden kopyalanmaz; komut adları Türkçe ofis yazılımlarında
  yerleşik, kısa ve betimleyici terimlerdir (Kes, Kopyala, Yapıştır…). Microsoft'a
  özgü özellik adları yerine genel adlar kullanılır: *Özet Tablo* (PivotTable),
  *Sanatsal Metin* (WordArt), *Akıllı Doldurma* (Flash Fill), *Mini Grafik*
  (Sparklines), *Arayüz teması*.

## Markalar

- Ürün adları **OfficeHere Belge**, **OfficeHere Tablo**, **OfficeHere Sunum**'dur.
  Arayüzde "Word", "Excel", "PowerPoint" adları kullanılmaz; yalnız Hakkında
  sayfasındaki marka notunda geçer.
- "OfficeHere" adı genel "office" sözcüğünü içerir (LibreOffice, OpenOffice,
  ONLYOFFICE, WPS Office gibi). Ad ticari olarak kullanılacaksa önce TÜRKPATENT,
  EUIPO ve USPTO'da benzer marka araştırması yapılması önerilir.
- Uyumluluğu anlatırken bu adlar yalnız betimleyici (nominatif) olarak ve marka
  notuyla birlikte geçer: *Microsoft, Word, Excel ve PowerPoint, Microsoft
  Corporation'ın ticari markalarıdır. OfficeHere, Microsoft Corporation ile
  bağlantılı değildir ve onun tarafından desteklenmez.*

## Yazı tipleri

Pakette hiçbir yazı tipi dosyası yoktur; Microsoft yazı tipleri dağıtılmaz ve
harici yazı tipi sunucularından (ör. Google Fonts) yükleme yapılmaz. Yazı tipi
adları (Calibri, Aptos…) yalnız belgedeki biçim bilgisi olarak okunup yazılır.
İstenen yazı tipi kullanıcının sisteminde yoksa, kuruluysa ölçü uyumlu açık yazı
tiplerine düşülür: Carlito (Calibri), Caladea (Cambria), Liberation
Sans/Serif/Mono (Arial, Times New Roman, Courier New).

## Bağımlılıkların lisansları

Projeye yalnız izin verici (permissive) lisanslı bağımlılıklar eklenir: MIT, ISC,
BSD-2/3-Clause, 0BSD, Apache-2.0, Zlib, BlueOak-1.0.0, Unlicense, CC0-1.0,
Python-2.0. GPL/LGPL/AGPL gibi karşılıklılık (copyleft) şartı taşıyan bağımlılık
eklenmez. Bu liste derlemede denetlenir (`IZINLI_LISANSLAR`,
`scripts/lisans-bildirimi.mjs`): listede olmayan ya da lisansı belirsiz bir
bileşen pakete girerse derleme durur. Başlıca bağımlılıklar:

| Paket | Kullanım | Lisans |
| --- | --- | --- |
| React | Arayüz | MIT |
| Electron | Masaüstü kabuğu | MIT (Chromium: BSD-3-Clause ve diğerleri) |
| Tiptap, ProseMirror | Belge editörü | MIT |
| ExcelJS | XLSX okuma/yazma | MIT |
| JSZip | OOXML paketleri | MIT (MIT/GPLv3 çift lisanstan MIT seçilir) |
| pako (JSZip içinde) | ZIP sıkıştırma | MIT ve Zlib |
| mammoth | DOCX içe aktarma yedeği | BSD-2-Clause |
| @fluentui/react-icons | Simgeler | MIT |
| electron-updater | Masaüstü güncelleyici | MIT (bağımlılıkları MIT/ISC/BlueOak/Python-2.0) |

Denetim notları (2026-10):

- `pnpm licenses list --prod` ile 260+ üretim bağımlılığı tarandı; GPL-only,
  LGPL, AGPL ya da ticari kısıtlı lisans yoktur.
- `buffers@0.1.1` (ExcelJS → unzipper → binary) pakette lisans bildirmez. Yalnız
  ExcelJS'in Node.js akış okuyucusunda kullanılır; tarayıcı paketine
  (`exceljs.min.js`) girmez, dolayısıyla dağıtılmaz. Derleme eklentisi bunu
  pakete giren modüllerden doğrular.
- Kullanılmayan `docx` bağımlılığı kaldırıldı.

## Üçüncü taraf bildirimleri

MIT, BSD, ISC ve Apache lisansları, dağıtılan her kopyada telif notunun ve lisans
metninin bulunmasını şart koşar. Bu şart otomatik karşılanır:

- `scripts/lisans-bildirimi.mjs` Vite eklentisi, her derlemede pakete gerçekten
  giren modülleri tarar ve her uygulamanın çıktısına `THIRD_PARTY_NOTICES.txt`
  yazar (bileşen adı, sürüm, lisans, kaynak ve tam lisans metni).
- Önceden derlenmiş paketlerin içine gömülü bileşenler de listelenir:
  `exceljs.min.js` için kaynak haritasından (crypto-browserify, core-js…),
  `jszip.min.js` için incelenerek belirlenen listeden (pako, lie…).
- Gömülü bileşenlerin ve lisans dosyası olmayan paketlerin metinleri (her ortamda aynı sonuç için) `scripts/lisanslar/`
  altında tutulur (`node scripts/lisans-onbellek.mjs <paket>` ile npm'den
  indirilir; npm paketinde lisans dosyası yoksa metin kaynak deposundan ya da
  README'deki lisans bölümünden birebir alınır).
- Lisans metni bulunamayan bir bileşen pakete girerse **derleme durur**; böylece
  bildirimsiz bileşen dağıtılamaz.
- Masaüstü ana sürecine paketlenen bağımlılıklar (electron-updater ve ağacı)
  başlangıç ekranının bildirimine eklenir. Lisans dosyası hiçbir yerde olmayan
  `lazy-val` için package.json'daki MIT bildirimi ve yazar adıyla standart MIT
  metni kullanılır.
- Bildirim, her uygulamanın Hakkında sayfasından ve başlangıç ekranının
  Seçenekler sayfasından açılır. Masaüstü paketinde ayrıca `LICENSE.txt`
  (OfficeHere), `LICENSE.electron.txt` ve `LICENSES.chromium.html` bulunur;
  web sürümünde `LICENSE.txt` site kökündedir.

## Kalan riskler

Aşağıdakiler yazılımla çözülemez; ticari dağıtımdan önce bir avukatla
değerlendirilmelidir:

- **Patentler:** Komut şeridi arayüzüne ilişkin Microsoft patent başvuruları
  2004–2006 dönemindedir; 20 yıllık koruma süresinin büyük ölçüde dolduğu
  varsayılmıştır. Dosya biçimleri Open Specification Promise kapsamındadır.
  Tek tek özellikler (ör. bir hesaplama ya da düzen davranışı) için patent
  araştırması yapılmamıştır.
- **Görünüm (trade dress):** Arayüz, Office'e benzer bir düzen kullanır ancak
  logo, simge ve adlar farklıdır ve her yerde bağımsızlık notu bulunur.
- **Marka:** "OfficeHere" adı için tescil araştırması yapılmamıştır.

## Katkıcılar için

Katkı göndererek, katkının bu belgeye uygun biçimde sizin tarafınızdan
üretildiğini ve MIT lisansıyla dağıtılabileceğini kabul etmiş olursunuz.
