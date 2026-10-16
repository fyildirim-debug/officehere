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

## Beta süreci

- Ön sürümler `vX.Y.Z-beta.N` etiketiyle yayımlanır (`release.yml`, girdi `surum`). Tireli etiket
  GitHub'da "Pre-release" olur; güncelleme dosyaları `beta*.yml` adını alır ve kararlı kanaldaki
  kullanıcılar (`latest*.yml`) betayı hiçbir yoldan görmez.
- Kullanıcı Seçenekler → **Güncelleme kanalı: Beta** ile katılır; mobilde beta APK GitHub sürümünden
  elle kurulur.
- Kararlı sürüm, son betadan en az **7 gün** sonra ve beta geri bildirimindeki engelleyici sorunlar
  kapanınca çıkar. Geri bildirim "Beta geri bildirimi" formuyla ya da e-postayla toplanır.
- Sürüm notları `docs/surumler/<etiket>.md`; yayımlamadan önce paketlerin `package.json` sürümleri
  etiketle aynı olmalıdır (iş akışı uyuşmazlıkta durur).

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
