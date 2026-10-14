# POI derlemi: gerçek Office dosyaları

Depo derlemi (`test-derlem/lo`, `test-derlem/py`) LibreOffice ve python ile üretildiği için Office'e özgü
parçaları (makro, SmartArt, OLE, özel XML, imza, x14/w14/p14 uzantıları…) içermez. Bu dizin, geçiş
katmanının ("tanımadığını kaybetme") ölçümü için Office'in kendi ürettiği 41 dosyayı içerir.

- **Kaynak:** [Apache POI](https://github.com/apache/poi) `trunk/test-data/` (14 Ekim 2026). Dosyalar
  değiştirilmeden, POI'deki adlarıyla alındı.
- **Lisans:** Apache License 2.0 ([LICENSE](LICENSE)); atıf: [NOTICE](NOTICE). Yalnız test girdisidir,
  uygulama paketlerine girmez.
- **Kullanım:** `pnpm derlem` her dosyayı açar, yerel biçimde kaydeder ve kaybolan parça / ilişki / uzantı
  öğesi sayılarını `scripts/derlem/temel.json` → `poi` temel çizgisiyle karşılaştırır. Uygulama başına
  parça kaybı toplamı artarsa CI durur; azalınca `pnpm derlem:temel` ile temel güncellenir.
  `pnpm derlem:opc` kaydedilen paketlerin tutarlılığını denetler.

## Seçim ölçütleri

Dosyalar tek tek seçildi. Dışarıda bırakılanlar:

- Hata raporlarından gelenler (`BugNNNNN`, `bugNNNNN`, numarayla başlayan adlar, `clusterfuzz-*`,
  `crash-*`, `tika-*`): kökeni ve içeriği doğrulanamaz.
- Kötü niyetli ya da bozuk girdiler (`ExternalEntityInText.docx` XXE, `poc-xmlbomb*`, `xxe_in_schema`,
  açılamayan `1_NoIden.xlsx`): bunlar bulanık testin işi (`pnpm bulanik`).
- 2 MB'tan büyük dosyalar.
- Üçüncü tarafların gerçek içeriğini taşıyanlar (haber metni, kurum sunumları, okul ders sunumları:
  ör. `drawing.docx`, `customGeo.pptx`, `keyframes.pptx`, `prProps.pptx`) ve kökeni belirsiz ses/video/
  fotoğraf içerenler (`EmbeddedAudio.pptx`, `chart-picture-bg.pptx`, `WithDrawing.xlsx`).

## Dosyalar

| Dosya | POI dizini | Neden |
|---|---|---|
| SimpleMacro.docm | document | VBA projesi (`vbaProject.bin`), makrolu içerik türü |
| EmbeddedDocument.docx | document | Gömülü OLE çalışma sayfası (.xls) + EMF önizleme |
| WordWithAttachments.docx | document | OLE ekleri (.doc/.xls/.xlsx), customXml, yorum, dipnot, `customizations.xml`, özel özellikler |
| checkboxes.docx | document | Form onay kutuları, metin kutusu, üç üst/alt bilgi çifti, customXml, ekli şablon ilişkisi |
| FieldCodes.docx | document | Alan kodları, w14, customXml |
| Numbering.docx | document | Numaralandırma, w15, `mc:AlternateContent` |
| TestTableColumns.docx | document | Tablo sütunları, w14/w15 |
| comment.docx | document | Yorum (LibreOffice üretimi; karşılaştırma için) |
| delins.docx | document | İzlenen değişiklikler (ekleme/silme), customXml |
| footnotes.docx | document | Dipnotlar, customXml |
| Headers.docx | document | Üst/alt bilgi türleri |
| WithGIF.docx | document | GIF resim, a14 |
| hello-world-signed.docx | xmldsign | XML-DSig imzası (`_xmlsignatures/`) |
| SimpleMacro.xlsm | spreadsheet | VBA projesi, x14 |
| ExcelWithAttachments.xlsm | spreadsheet | VBA + OLE ekleri (.doc/.pptx), yorumlar, yazıcı ayarları |
| WithEmbeded.xlsx | spreadsheet | Gömülü OLE nesneleri |
| NewStyleConditionalFormattings.xlsx | spreadsheet | x14 koşullu biçim (simge kümeleri, veri çubukları) |
| CustomXMLMappings.xlsx | spreadsheet | XML eşlemeleri (`xmlMaps.xml`) |
| DataValidationEvaluations.xlsx | spreadsheet | 17 veri doğrulama kuralı (liste, aralık, formül) |
| ExcelPivotTableSample.xlsx | spreadsheet | Özet tablo + önbellek, `connections.xml` |
| ExcelTables.xlsx | spreadsheet | Tablo nesneleri |
| StructuredReferences.xlsx | spreadsheet | Yapılandırılmış başvurular, dinamik dizi meta verisi |
| WithThreeCharts.xlsx | spreadsheet | Üç grafik |
| SimpleScatterChart.xlsx | spreadsheet | Grafik sayfası, grafik stili ve renkleri (`cs:`), c14 |
| WithTextBox.xlsx | spreadsheet | Çizim metin kutusu (`xdr:sp`) |
| comments.xlsx | spreadsheet | Yorumlar (VML) |
| right-to-left.xlsx | spreadsheet | Zincirli yorumlar (`threadedComments`, `persons`), İbranice/Arapça metin |
| hello-world-signed.xlsx | xmldsign | XML-DSig imzası |
| SimpleMacro.pptm | slideshow | VBA projesi |
| PPTWithAttachments.pptm | slideshow | VBA + OLE ekleri, notlar, el notu asıl sayfası |
| SmartArt.pptx | slideshow | SmartArt (`dgm`, `dsp`) |
| smartart-simple.pptx | slideshow | SmartArt |
| smartart-rotated-text.pptx | slideshow | SmartArt, döndürülmüş metin |
| bar-chart.pptx | slideshow | Çubuk grafik, gömülü çalışma kitabı, c14, `mc:AlternateContent` |
| pie-chart.pptx | slideshow | Pasta grafik |
| SampleShow.pptx | slideshow | p14/p15 uzantıları |
| backgrounds.pptx | slideshow | Slayt arka planları |
| layouts.pptx | slideshow | Asıl slayt düzenleri |
| shapes.pptx | slideshow | Şekiller, a14 |
| table_test.pptx | slideshow | Tablo |
| hello-world-signed.pptx | xmldsign | XML-DSig imzası |

## Bilinen eksikler

- **Dilimleyici (slicer) ve zaman çizelgesi:** POI test verisinde adından ya da POI testlerinden
  anlaşılan bir örnek bulunamadı; incelenen adayların hiçbirinde yok. Sentetik zenginleştiriciyle (teknik
  tasarım K§1.2.9) eklenecek.
- **Yapı taşları (`word/glossary/`):** incelenen adaylardan yalnız hata raporu kökenlilerde (Bug54849,
  Bug54771a) vardı; alınmadı.
- **x14 veri doğrulama (başka sayfaya liste):** incelenen adaylarda bulunamadı. x14 uzantıları koşullu
  biçim, özet tablo ve stil listelerinden ölçülüyor.
