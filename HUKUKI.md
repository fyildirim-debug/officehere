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
  ISO/IEC 29500) dayanır. Bu belirtimler Microsoft'un *Open Specification Promise*
  kapsamındadır: Microsoft, belirtimi uygulayan yazılımlara karşı bu
  belirtimlerden doğan patent iddialarında bulunmayacağını taahhüt eder.
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
  kullanımına açtığı açık kaynak simge setidir (`@fluentui/react-icons`).
  Uygulama ve dosya türü simgeleri bu projeye özgüdür.

## Markalar

- Ürün adları **OfficeHere Belge**, **OfficeHere Tablo**, **OfficeHere Sunum**'dur.
  Arayüzde "Word", "Excel", "PowerPoint" adları kullanılmaz.
- Uyumluluğu anlatırken bu adlar yalnız betimleyici (nominatif) olarak ve marka
  notuyla birlikte geçer: *Microsoft, Word, Excel ve PowerPoint, Microsoft
  Corporation'ın ticari markalarıdır. OfficeHere, Microsoft Corporation ile
  bağlantılı değildir ve onun tarafından desteklenmez.*

## Yazı tipleri

Microsoft yazı tipleri dağıtılmaz. Belge bir Microsoft yazı tipi istiyorsa ve bu
yazı tipi kullanıcının sisteminde yoksa, ölçü uyumlu açık yazı tiplerine düşülür:
Carlito (Calibri), Caladea (Cambria), Liberation Sans/Serif/Mono (Arial, Times New
Roman, Courier New).

## Bağımlılıkların lisansları

Projeye yalnız MIT, BSD, Apache-2.0 ve ISC lisanslı bağımlılıklar eklenir;
GPL/LGPL/AGPL bağımlılık eklenmez. Başlıca bağımlılıklar:

| Paket | Kullanım | Lisans |
| --- | --- | --- |
| React | Arayüz | MIT |
| Electron | Masaüstü kabuğu | MIT |
| Tiptap, ProseMirror | Belge editörü | MIT |
| ExcelJS | XLSX okuma/yazma | MIT |
| PptxGenJS | PPTX yazma | MIT |
| JSZip | OOXML paketleri | MIT (MIT/GPLv3 çift lisanstan MIT seçilir) |
| docx | DOCX yazma | MIT |
| mammoth | DOCX içe aktarma yedeği | BSD-2-Clause |
| @fluentui/react-icons | Simgeler | MIT |

Tam liste için `pnpm licenses list` komutunu kullanın.

## Katkıcılar için

Katkı göndererek, katkının bu belgeye uygun biçimde sizin tarafınızdan
üretildiğini ve MIT lisansıyla dağıtılabileceğini kabul etmiş olursunuz.
