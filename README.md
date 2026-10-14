<div align="center">

<img src="apps/desktop/build/icons/128x128.png" width="96" alt="OfficeHere" />

# OfficeHere

**Açık kaynak, Türkçe, tanıdık ofis paketi.**
Belge, Tablo ve Sunum — `.docx`, `.xlsx`, `.pptx` başta olmak üzere OpenDocument (`.odt` `.ods` `.odp`),
eski 97-2003 (`.doc` `.xls` `.ppt`), RTF, Markdown ve CSV dosyalarını açar, düzenler ve kaydeder.

[İndir](https://github.com/fyildirim-debug/officehere/releases) ·
[Hukuki çerçeve](HUKUKI.md) ·
[Mimari](MIMARI.md) ·
[Katkı](KATKI.md)

</div>

![OfficeHere Belge](docs/ekran/belge.png)

<table>
<tr>
<td><img src="docs/ekran/tablo.png" alt="OfficeHere Tablo" /></td>
<td><img src="docs/ekran/sunum.png" alt="OfficeHere Sunum" /></td>
</tr>
<tr>
<td align="center">Tablo — koşullu biçimlendirme, grafikler, Türkçe formüller</td>
<td align="center">Sunum — temalar, slayt bölmesi, notlar</td>
</tr>
</table>

## Uygulamalar

| | Uygulama | Karşılığı | Dosyalar |
| --- | --- | --- | --- |
| <img src="apps/desktop/build/filetypes/belge.png" width="28" /> | **OfficeHere Belge** | Kelime işlemci | `.docx` `.odt` `.rtf` `.md` `.html` `.txt` · açar: `.doc` `.dotx` `.dotm` `.ott` → PDF |
| <img src="apps/desktop/build/filetypes/tablo.png" width="28" /> | **OfficeHere Tablo** | Hesap tablosu | `.xlsx` `.ods` `.csv` `.tsv` · açar: `.xls` `.xlsm` `.xltx` `.xltm` `.ots` `.html` → PDF, HTML |
| <img src="apps/desktop/build/filetypes/sunum.png" width="28" /> | **OfficeHere Sunum** | Sunu | `.pptx` `.odp` `.ppsx` `.potx` · açar: `.ppt` `.pptm` `.potm` `.ppsm` `.otp` → PDF, PNG |

Üçü de aynı tanıdık düzeni paylaşır: renkli başlık çubuğu, Hızlı Erişim Araç
Çubuğu, sekmeli komut şeridi, Dosya görünümü, durum çubuğu ve yakınlaştırma.
Arayüz tamamen Türkçedir; Renkli, Beyaz, Koyu Gri ve Siyah temaları vardır.

### Belge

- Gerçek sayfa görünümü: A4/A3/A5/Letter/Legal, dikey/yatay, kenar boşlukları, sütunlar, bölümler
- Sayfalama (dul/yetim satır denetimi, birlikte tut), cetvel, üst/alt bilgi ve sayfa numarası
- Stil galerisi ve stil kümeleri, temalar, filigran, sayfa rengi ve kenarlığı
- Listeler ve çok düzeyli listeler, tablolar (birleştir/böl, tablo stilleri), resimler, şekiller, kapak sayfası, simgeler
- İçindekiler, dipnot ve son not, açıklamalar, değişiklikleri izleme (kabul/ret, değişiklik çubukları)
- Resim yazısı, çapraz başvuru, kaynakça (APA/IEEE), dizin, denklem, grafik, satır numaraları
- Metni kaydır, resim kırpma, Stiller bölmesi, sekme durakları, Otomatik Düzelt, mini araç çubuğu
- Adres mektup birleştirme (CSV/XLSX), belge karşılaştırma, düzenleme koruması
- Bul ve Değiştir (Türkçe İ/i, I/ı duyarlı), Anahat/Taslak/Okuma/Web görünümleri, ayrıntılı yazdırma ayarları
- Büyük İlk Harf, biçimli ve otomatik güncellenen tarih/saat, Simge (son kullanılanlar), Alt+X kod dönüşümü
- Sesli Oku (okunan sözcük vurgulanır, hız ve ses seçimi), Sözcük Sayısı (seçim istatistikleri)
- Numaralandırmaya yeniden başla / devam et / değer ayarla
- Bilinen klavye kısayolları, sağ tık menüleri, otomatik kurtarma ve çökme koruması

### Tablo

- 1.048.576 satır × 16.384 sütunluk akıcı ızgara, bölmeleri dondurma, sayfa sekmeleri
- 370'ten fazla işlev; **Türkçe işlev adları** (`=TOPLA(A1:A10)`, `=EĞER(…;…;…)`, `=DÜŞEYARA`) ve İngilizce adlar birlikte
- Formül çubuğu, renkli başvuru vurguları, F4, otomatik tamamlama, doldurma tutamacıyla seriler
- Hücreleri Biçimlendir, sayı biçimleri (₺, %, tarih), koşullu biçimlendirme, tablo ve hücre stilleri
- Tablo olarak biçimlendirme ve yapısal başvurular (`=TOPLA(Tablo1[Tutar])`), Özet Tablo, Dilimleyici
- Sıralama, filtre, gelişmiş filtre, veri doğrulama, Anahat/Alt Toplam, Hedef Ara, Senaryo Yöneticisi, Veri Tablosu
- Formül Denetimi (etkileyen/etkilenen okları, Formülü Değerlendir), Akıllı Doldurma, simge kümeleri
- Grafikler ve mini grafikler, resimler, şekiller, notlar, sayfa ve kitap koruma, hücre içi onay kutusu
- Dinamik diziler (FİLTRE, SIRALA, BENZERSİZ, ÇAPRAZARA, LET, LAMBDA), REGEX işlevleri, veritabanı ve mühendislik işlevleri, DOT/EĞİLİM
- Otomatik Doldurma Seçenekleri, Yinelenenleri Kaldır, Metni Sütunlara Dönüştür, Birleştir, Bağlantıyı Yapıştır
- Sayfa Düzeni ve Sayfa Sonu Önizleme, yazdırma ayarları ve PDF; 100 bin satırda akıcı filtre/sıralama
- Python/openpyxl, LibreOffice ve diğer araçların ürettiği XLSX dosyalarıyla uyum

### Sunum

- Normal, Ana Hat, Slayt Sıralayıcısı, Not Sayfası, Okuma ve Asıl Slayt görünümleri
- Şekiller, metin kutuları, resimler (kırpma), tablolar, grafikler, video ve ses
- Akıllı kılavuzlar, döndürme, hizalama ve dağıtma, gruplama, bölümler, köprüler
- 10 tema × 4 çeşitleme, arka plan biçimlendirme, slayt boyutu
- Gerçekten oynayan geçişler ve animasyonlar, Animasyon Bölmesi
- Tam ekran slayt gösterisi, iki pencereli sunucu görünümü, kalem/vurgulayıcı/lazer işaretçi
- Dönüştür geçişi, hareket yolları, eylem düğmeleri, özet ve slayt yakınlaştırma, zamanlama provası
- Şekilleri birleştirme, noktaları düzenleme, resimleri sıkıştırma, erişilebilirlik denetleyicisi, damlalık
- Yazdırma düzenleri: tam sayfa, not sayfaları, ana hat ve dinleyici notları (1–9 slayt); videoya aktarma (WebM)

## Kurulum

[Sürümler](https://github.com/fyildirim-debug/officehere/releases) sayfasından
işletim sisteminize uygun paketi indirin:

| Sistem | Paket |
| --- | --- |
| Windows 10/11 | `OfficeHere-<sürüm>-win-x64.exe` (64 bit kurulum), `-win-arm64.exe` (ARM), `-win.exe` (iki mimari tek kurulumda) ya da `-tasinabilir.exe` (kurulumsuz) |
| Linux | `.AppImage` ya da `.deb` |
| macOS (Apple Silicon) | `-mac-arm64.dmg` (imzasız: ilk açılışta sağ tık → Aç) |

Kurulum `.docx`, `.doc`, `.odt`, `.rtf`, `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv`,
`.pptx`, `.ppt`, `.ppsx` ve `.odp` dosyalarını OfficeHere ile ilişkilendirir.

**Diğer biçimler:** Eski `.doc`/`.xls`/`.ppt` dosyaları yalnız açılır; kaydedince
`.docx`/`.xlsx`/`.pptx` olur (özgün dosya değişmez). `.odt`, `.ods`, `.odp`, `.rtf`,
`.md` ve `.csv` dosyaları aynı biçimde kaydedilebilir; ilk kayıtta bazı özelliklerin
kaybolabileceği sorulur.

**Güncelleştirmeler:** Uygulama GitHub Sürümleri'ni açılışta ve altı saatte bir
denetler (Seçenekler'den kapatılabilir). Windows kurulumu ve AppImage yeni sürümü
arka planda indirir, yeniden başlatınca kurar; taşınabilir `.exe`, `.deb` ve
macOS paketi yeni sürümü bildirir ve indirme sayfasını açar. Otomatik
güncelleştirme 0.1.2 ile gelir; daha eski sürümlerden bir kez elle yükseltin.

**Otomatik kaydetme:** Masaüstünde diskte kayıtlı `.docx`, `.xlsx` ve `.pptx`
belgeleri değiştikçe kendiliğinden kaydedilir; ilk otomatik kayıttan önce özgün
dosyanın yedeği (`<dosya>.bak`) alınır. Seçenekler'den kapatılabilir. Paketler henüz kod imzalı değildir; Windows SmartScreen ilk
açılışta uyarı gösterebilir ("Ek bilgi" → "Yine de çalıştır").

Tarayıcıda kurulumsuz kullanmak için web sürümünü derleyip herhangi bir statik
sunucuda yayınlayabilirsiniz (`pnpm web` ya da `docker build -t officehere .`).

### İndirilen paketi doğrulama

Her sürümde `SHA256SUMS.txt` (tüm paketlerin SHA-256 özetleri) ve onun
[Sigstore](https://www.sigstore.dev/) imzası `SHA256SUMS.txt.sigstore.json` bulunur. İmza anahtarsızdır:
sürümü üreten GitHub Actions iş akışının kimliğiyle atılır ve herkese açık şeffaflık kaydına (Rekor) yazılır.
[cosign](https://docs.sigstore.dev/cosign/system_config/installation/) ile:

```bash
# 1) Sağlama listesinin bu depodaki sürüm iş akışınca imzalandığını doğrulayın
#    (cosign 3.x; cosign 2.x kullanıyorsanız --new-bundle-format ekleyin)
cosign verify-blob SHA256SUMS.txt \
  --bundle SHA256SUMS.txt.sigstore.json \
  --certificate-identity-regexp '^https://github\.com/fyildirim-debug/officehere/\.github/workflows/release\.yml@' \
  --certificate-oidc-issuer https://token.actions.githubusercontent.com

# 2) İndirdiğiniz paketin özetini listeyle karşılaştırın
sha256sum --ignore-missing -c SHA256SUMS.txt          # Linux
shasum -a 256 --ignore-missing -c SHA256SUMS.txt      # macOS
```

Windows'ta PowerShell ile: `Get-FileHash .\OfficeHere-<sürüm>-win-x64.exe -Algorithm SHA256` çıktısını
`SHA256SUMS.txt` içindeki satırla karşılaştırın.

## Geliştirme

Gereksinimler: Node.js 22+, pnpm 10.

```bash
pnpm install
pnpm dev:belge        # http://localhost:5173
pnpm dev:tablo        # http://localhost:5174
pnpm dev:sunum        # http://localhost:5175
pnpm dev:desktop      # Electron kabuğu + başlangıç ekranı

pnpm typecheck        # tüm paketler
pnpm test             # birim testleri (Vitest; performans testleri hariç)
pnpm test:perf        # performans testleri (ayrı ve sıralı; makine hızına göre normalize)
pnpm denetim          # yerel ayarsız toUpperCase/toLowerCase ve waitForTimeout sayaçları
pnpm web && pnpm e2e  # uçtan uca testler (Playwright)
pnpm dist:win         # Windows kurulum paketi (ayrıca dist:linux, dist:mac)

pnpm derlem           # derlem kaydet-aç turu → .derlem/rapor/derlem.md
pnpm bulanik          # 756 bozuk dosyayla bulanık test → .derlem/rapor/bulanik.md
pnpm derlem:uret      # büyük derlem dosyaları (LibreOffice + python; isteğe bağlı)
pnpm derlem:lo        # kaydedilen dosyaları LibreOffice ile PDF'e çevirip denetler
pnpm derlem:temel     # iyileşmeden sonra temel çizgiyi günceller (scripts/derlem/temel.json)
```

Ayrıntılar: [MIMARI.md](MIMARI.md), [KATKI.md](KATKI.md).

## Uyumluluk ve testler

- Dosya biçimleri açık **Office Open XML (ECMA-376 / ISO/IEC 29500)** ve **OpenDocument (ISO/IEC 26300)** standartlarına, eski ikili biçimler Microsoft'un yayımladığı açık belirtimlere ([MS-CFB], [MS-DOC], [MS-XLS], [MS-PPT]) göre uygulanır. Belge (`.docx` `.odt` `.rtf` `.doc` `.md`), Sunum (`.pptx` `.odp` `.ppt`) ve Tablo'nun `.ods`/`.xls`/`.csv` okuyucu/yazıcıları projenin kendi kodudur. Üç kütüphane kullanılır: **XLSX** okuma/yazmada [ExcelJS](https://github.com/exceljs/exceljs) (MIT; grafik, filtre, ad, uzantı gibi ExcelJS'in desteklemediği parçalar kendi kodumuzla okunup yazılır), ayrıştırılamayan **DOCX** dosyalarında yedek dönüşüm için [mammoth](https://github.com/mwilliamson/mammoth.js) (BSD-2-Clause) ve ZIP paketleri için [JSZip](https://github.com/Stuk/jszip) (MIT). Tam liste: [HUKUKI.md](HUKUKI.md#bağımlılıkların-lisansları).
- Her uygulamanın gidiş-dönüş testleri (yaz → oku → karşılaştır) vardır; LibreOffice, python-docx / openpyxl / python-pptx ile üretilmiş örnek dosyalar okuma testlerindedir.
- **Sağlamlık** (CI, her gönderimde): LibreOffice ve python ile üretilmiş gerçek dünya derlemi ve Apache POI test verisinden seçilmiş 41 gerçek Office dosyası ([test-derlem/poi](test-derlem/poi/README.md), Apache-2.0, [NOTICE](test-derlem/poi/NOTICE)) kaydet-aç turundan geçirilir (`pnpm derlem`: metin, biçim, hücre, slayt, not ve OOXML/ODF parça, ilişki ve uzantı öğesi kayıpları sayılır), kaydedilen paketlerin tutarlılığı denetlenir (`pnpm derlem:opc`), dosyalar LibreOffice'e açtırılır (`pnpm derlem:lo`) ve 756 bozuk dosyayla bulanık test yapılır (`pnpm bulanik`: takılma, yığın taşması ve anlaşılmaz hata iletileri). Sonuçlar kaydedilmiş temel çizgiyle karşılaştırılır; yeni takılma, yığın taşması, POI derleminde artan parça kaybı ya da yeni paket hatası derlemeyi durdurur.
- Formül motorunun sonuçları LibreOffice ile yüzlerce formülde çapraz doğrulanır.
- Uçtan uca testler her uygulamayı gerçek bir tarayıcıda kullanıcı gibi kullanır; CI her gönderimde hepsini çalıştırır. Windows ve macOS'ta duman testi `main` dalına her gönderimde koşar.

## Hukuki

OfficeHere bağımsız bir projedir; Microsoft Corporation ile bağlantılı değildir
ve onun tarafından desteklenmez. Microsoft, Word, Excel ve PowerPoint, Microsoft
Corporation'ın ticari markalarıdır. Office uygulamalarının kodu, logoları,
şablonları ya da yazı tipleri kullanılmaz; arayüz simgeleri Microsoft'un MIT
lisansıyla herkese açtığı Fluent UI System Icons setidir. Dosya biçimi desteği
açık ECMA-376 / ISO/IEC 29500 standardına dayanır. Ayrıntılar: [HUKUKI.md](HUKUKI.md).

## Lisans

[MIT](LICENSE) © 2026 [Furkan YILDIRIM](https://furkanyildirim.com)

Kullanılan açık kaynak bileşenlerin lisans metinleri her derlemede otomatik
üretilen `THIRD_PARTY_NOTICES.txt` dosyasındadır (uygulamalarda Dosya → Hakkında).
