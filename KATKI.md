# Katkı rehberi

## Kurulum

Gereksinimler: Node.js 22+, pnpm 10.

```bash
pnpm install
pnpm dev:belge      # http://localhost:5173
pnpm dev:tablo      # http://localhost:5174
pnpm dev:sunum      # http://localhost:5175
pnpm dev:desktop    # Electron kabuğu + üç uygulama
```

## Denetimler

Göndermeden önce şunlar geçmelidir:

```bash
pnpm denetim     # yeni toUpperCase()/toLowerCase() ya da waitForTimeout eklenmedi mi
pnpm typecheck
pnpm test
pnpm build
```

Dosya biçimi kodunu değiştirdiyseniz `pnpm derlem` ve `pnpm bulanik` ile
derlem turunu ve bozuk dosya testini çalıştırın (CI'da "Sağlamlık" işi). Kayıp
azaldıysa temel çizgiyi `pnpm derlem:temel` ile güncelleyin.

## Kurallar

- **Hukuk:** [HUKUKI.md](HUKUKI.md) içindeki temiz oda kurallarına uyun. Microsoft
  kodu, simgesi, şablonu ya da yazı tipi kopyalanmaz; yalnız MIT/BSD/Apache-2.0/ISC
  gibi izin verici lisanslı bağımlılık eklenir. Lisans bildirimi derlemede
  otomatik üretilir; derleme "Lisans metni bulunamadı" diye durursa
  `node scripts/lisans-onbellek.mjs <paket>` çalıştırıp çıkan dosyayı ekleyin.
- **Dil:** Arayüz metinleri ve kod yorumları Türkçedir. Arayüz terimleri Türkçe
  ofis yazılımlarında yerleşik karşılıklarıyla yazılır (Giriş, Ekle, Düzen,
  Görünüm, Yapıştır, Biçim Boyacısı…).
- **Yerel ayar:** Büyük/küçük harf dönüşümünde `toLocaleUpperCase("tr-TR")`,
  sayılarda `1.234,56`, tarihte `gg.aa.yyyy`.
- **Ortak kod:** Birden çok uygulamanın ihtiyaç duyduğu arayüz parçası
  `packages/ui`'ye, çerçeveden bağımsız yardımcılar `packages/core`'a gider.
- **Testler:** Dosya biçimi okuyucu/yazıcıları ve formül fonksiyonları için
  gidiş-dönüş ve birim testleri yazın.
- **Commit mesajları:** Türkçe, ne değiştiğini söyleyen kısa bir başlık.
