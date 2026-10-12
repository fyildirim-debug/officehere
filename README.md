<div align="center">

<img src="apps/desktop/build/icons/128x128.png" width="96" alt="OfficeHere" />

# OfficeHere

**Açık kaynak, Türkçe, tanıdık ofis paketi.**
Belge, Tablo ve Sunum — `.docx`, `.xlsx`, `.pptx` dosyalarını açar, düzenler ve kaydeder.

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
| <img src="apps/desktop/build/filetypes/belge.png" width="28" /> | **OfficeHere Belge** | Kelime işlemci | `.docx` `.html` `.txt` → PDF |
| <img src="apps/desktop/build/filetypes/tablo.png" width="28" /> | **OfficeHere Tablo** | Hesap tablosu | `.xlsx` `.csv` → PDF |
| <img src="apps/desktop/build/filetypes/sunum.png" width="28" /> | **OfficeHere Sunum** | Sunu | `.pptx` → PDF, PNG |

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
- Bilinen klavye kısayolları, sağ tık menüleri, otomatik kurtarma ve çökme koruması

### Tablo

- 1.048.576 satır × 16.384 sütunluk akıcı ızgara, bölmeleri dondurma, sayfa sekmeleri
- 370'ten fazla işlev; **Türkçe işlev adları** (`=TOPLA(A1:A10)`, `=EĞER(…;…;…)`, `=DÜŞEYARA`) ve İngilizce adlar birlikte
- Formül çubuğu, renkli başvuru vurguları, F4, otomatik tamamlama, doldurma tutamacıyla seriler
- Hücreleri Biçimlendir, sayı biçimleri (₺, %, tarih), koşullu biçimlendirme, tablo ve hücre stilleri
- Tablo olarak biçimlendirme ve yapısal başvurular (`=TOPLA(Tablo1[Tutar])`), Özet Tablo, Dilimleyici
- Sıralama, filtre, gelişmiş filtre, veri doğrulama, Anahat/Alt Toplam, Hedef Ara, Senaryo Yöneticisi, Veri Tablosu
- Formül Denetimi (etkileyen/etkilenen okları, Formülü Değerlendir), Akıllı Doldurma, simge kümeleri
- Grafikler ve mini grafikler, resimler, şekiller, notlar, sayfa ve kitap koruma
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
- Şekilleri birleştirme, noktaları düzenleme, resimleri sıkıştırma, erişilebilirlik denetleyicisi

## Kurulum

[Sürümler](https://github.com/fyildirim-debug/officehere/releases) sayfasından
işletim sisteminize uygun paketi indirin:

| Sistem | Paket |
| --- | --- |
| Windows 10/11 | `OfficeHere-<sürüm>-win-x64.exe` (64 bit kurulum), `-win-arm64.exe` (ARM), `-win.exe` (iki mimari tek kurulumda) ya da `-tasinabilir.exe` (kurulumsuz) |
| Linux | `.AppImage` ya da `.deb` |
| macOS (Apple Silicon) | `-mac-arm64.dmg` (imzasız: ilk açılışta sağ tık → Aç) |

Kurulum `.docx`, `.xlsx`, `.csv` ve `.pptx` dosyalarını OfficeHere ile
ilişkilendirir. Paketler henüz kod imzalı değildir; Windows SmartScreen ilk
açılışta uyarı gösterebilir ("Ek bilgi" → "Yine de çalıştır").

Tarayıcıda kurulumsuz kullanmak için web sürümünü derleyip herhangi bir statik
sunucuda yayınlayabilirsiniz (`pnpm web` ya da `docker build -t officehere .`).

## Geliştirme

Gereksinimler: Node.js 22+, pnpm 10.

```bash
pnpm install
pnpm dev:belge        # http://localhost:5173
pnpm dev:tablo        # http://localhost:5174
pnpm dev:sunum        # http://localhost:5175
pnpm dev:desktop      # Electron kabuğu + başlangıç ekranı

pnpm typecheck        # tüm paketler
pnpm test             # birim testleri (Vitest)
pnpm web && pnpm e2e  # uçtan uca testler (Playwright)
pnpm dist:win         # Windows kurulum paketi (ayrıca dist:linux, dist:mac)
```

Ayrıntılar: [MIMARI.md](MIMARI.md), [KATKI.md](KATKI.md).

## Uyumluluk ve testler

- Dosya biçimleri açık **Office Open XML (ECMA-376 / ISO/IEC 29500)** standardına göre kendi okuyucu/yazıcılarımızla uygulanır.
- Her uygulamanın gidiş-dönüş testleri (yaz → oku → karşılaştır) vardır; üretilen dosyalar LibreOffice ve python-docx / openpyxl / python-pptx ile dış doğrulamadan geçirilir.
- Formül motorunun sonuçları LibreOffice ile yüzlerce formülde çapraz doğrulanır.
- Uçtan uca testler her uygulamayı gerçek bir tarayıcıda kullanıcı gibi kullanır; CI her gönderimde hepsini çalıştırır.

## Hukuki

OfficeHere bağımsız bir projedir; Microsoft Corporation ile bağlantılı değildir
ve onun tarafından desteklenmez. Microsoft, Word, Excel ve PowerPoint, Microsoft
Corporation'ın ticari markalarıdır. Hiçbir Microsoft kodu, simgesi, şablonu ya
da yazı tipi kullanılmaz. Ayrıntılar: [HUKUKI.md](HUKUKI.md).

## Lisans

[MIT](LICENSE) © 2026 Furkan YILDIRIM
