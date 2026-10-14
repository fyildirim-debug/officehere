# Görsel regresyon temel görüntüleri

`linux/<ad>.png`: `e2e/gorsel/*.spec.ts` testlerinin karşılaştırdığı temel görüntüler (Playwright `toHaveScreenshot`;
yapılandırma `playwright.gorsel.config.ts`).

| Önek | İçerik |
|---|---|
| `belge-<dosya>-s<n>` | Belge: derlem dosyasının n. sayfası (`.bl-page`) |
| `tablo-<dosya>-<n>` | Tablo: derlem dosyasının n. çekilen sayfasının ızgarası (`.tb-canvas-wrap`) |
| `sunum-<dosya>-s<n>` | Sunum: derlem dosyasının n. slaydı (`.stage`) |
| `kabuk-<uygulama>-acik/koyu` | Boş belgede tam pencere, açık (Renkli) ve koyu (Koyu Gri) tema |
| `serit-<uygulama>-giris/ekle/dar` | Şerit: Giriş, Ekle, 900 px dar pencere (daralmış gruplar) |
| `ozdenetim-*` | Karşılaştırıcının öz denetimi (1 px kaydırma temelden ayrışmalı) |

Karşılaştırma: `maxDiffPixelRatio 0.002` ve `maxDiffPixels 40` (hangisi küçükse), `threshold 0.1`, animasyonlar kapalı,
metin imleci gizli, yeniden deneme yok. Piksel sınırı, seyrek sayfada 1 px kaymanın da yakalanması içindir; çizim
aynı imajda bayt bayt belirlenimcidir. Belirlenimcilik için Chromium bayrakları ve resimlerin önden çözülmemesi
`playwright.gorsel.config.ts` ve `e2e/gorsel/ortak.ts` içinde açıklanmıştır.

## Ortam: yalnız sabit Docker imajı

Yazı tipi çizimi makineden makineye değiştiği için temeller **yalnız** CI ile aynı ortamda üretilir ve
karşılaştırılır: `mcr.microsoft.com/playwright:v1.56.1-noble` + sürümü sabit `fonts-crosextra-carlito`,
`fonts-crosextra-caladea`, `fonts-liberation`, `fonts-dejavu-core` (sürümler `e2e/gorsel/calistir.mjs` ve
`.github/workflows/ci.yml` içinde, ikisi aynı tutulur). Bu dizine başka bir makinede üretilmiş görüntü koymayın.

## Komutlar

```sh
pnpm gorsel                        # karşılaştır (Docker'da); fark görüntüleri: test-results/gorsel/
pnpm gorsel:guncelle               # eksik ya da eşiği aşan temelleri yeniden yaz (Docker'da)
pnpm gorsel:guncelle -- --hepsi    # dizini silip hepsini yeniden üret (kullanılmayanlar da gider)
pnpm gorsel -- -g Sunum            # diğer argümanlar Playwright'a geçer
```

Docker yoksa betik ne yapılacağını anlatıp çıkar. Docker'sız yalnız **bu makinede** belirlenimcilik denemesi
(temeller geçici dizine yazılır, depodakilerle karşılaştırılmaz; önce uygulamalar derlenmiş olmalı):

```sh
pnpm -r --filter ./apps/belge --filter ./apps/tablo --filter ./apps/sunum run build
OH_GORSEL_YEREL=1 pnpm gorsel:guncelle && OH_GORSEL_YEREL=1 pnpm gorsel
OH_GORSEL_YEREL=1 OH_GORSEL_KAYDIR=1 pnpm gorsel   # kasıtlı 1 px kaydırma: içeriği olan her görüntü kırmızı
```

## CI (`ci.yml`, `gorsel` işi)

İş, aynı imajda (`container:`) çalışır ve kipini kendisi seçer:

| Durum | Kip | Sonuç |
|---|---|---|
| Depoda temel görüntü yok (`linux/*.png` boş) | üret | İş **yeşil**; görüntüler `gorsel-temel` artefaktında |
| Elle tetikleme (Actions › CI › Run workflow), `gorsel_temel` işaretli | hepsini üret | İş yeşil; `gorsel-temel` artefaktı. Diğer CI işleri atlanır |
| PR'da `gorsel-degisti` etiketi | değişenleri üret | İş yeşil; `gorsel-temel` artefaktı (tam dizin) |
| Diğer | karşılaştır | Fark varsa iş kırmızı; fark görüntüleri `gorsel-fark` artefaktında |

Temel güncelleme yolu: `pnpm gorsel:guncelle` (Docker) ya da `gorsel-temel` artefaktını indirip içindeki PNG'leri
bu dizinin `linux/` alt dizinine koymak; ardından commit ve PR'a `gorsel-degisti` etiketi. Gözden geçiren PNG
farklarına bakar. Chromium/Playwright yükseltmesi temellerin topluca güncellenmesini gerektirir (imaj sürümü
`@playwright/test` sürümüyle aynı tutulur).

Boyut: ~80 görüntü, ~4-5 MB. Belirgin büyürse git LFS'e geçilir.
