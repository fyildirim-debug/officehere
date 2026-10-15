# Dağıtım kanalları

OfficeHere masaüstü paketlerinin hangi kanallardan, nasıl ve hangi maliyetle dağıtıldığı. Kod
tarafı hazırdır; hesap, belirteç, anahtar ve alan adı gerektiren adımlar Furkan'ındır ve
[sonda sıralıdır](#furkanın-yapacakları). Hiçbir gizli değer depoya girmez: hepsi onaylı GitHub
ortamlarında (Settings → Environments) tutulur.

| Kanal | Ne | Nerede üretilir | İmza / SmartScreen | Maliyet | Durum |
| --- | --- | --- | --- | --- | --- |
| GitHub Sürümleri | NSIS `.exe` (x64, arm64, ikisi bir arada), taşınabilir `.exe`, `.AppImage`, `.deb`, `.dmg` | `release.yml` → `paket`, `yayin` | İmzasız; `SHA256SUMS.txt` cosign ile imzalı | 0 | Yayında |
| Microsoft Store | `.appx` (x64, arm64) | `release.yml` → `magaza` (artefakt `magaza-appx`) | Mağaza yeniden imzalar: **uyarı yok** | 0 (bireysel hesap) | Hesap bekliyor |
| winget | NSIS `.exe` (GitHub Sürümleri'nden) | `.github/winget/*.yaml` (ilk gönderim), `release.yml` → `winget` (`komac update`) | İmzasız; SmartScreen yine uyarabilir | 0 | İlk PR ve belirteç bekliyor |
| apt deposu (Pardus, Debian, Ubuntu) | İmzalı `stable` deposu, `officehere` paketi | `scripts/apt/*`, `release.yml` → `apt` | GPG (deb822 `Signed-By`) | 0 (R2 ücretsiz katman) | Anahtar, kova ve alan adı bekliyor |
| Pardus Yazılım Merkezi | Aynı `.deb` / apt deposu | `pardus.yml` (Pardus 23/25 kurulum + açılış sınaması) | — | 0 | apt deposundan sonra başvuru |

Sonraya bırakılanlar (3. aşama): Flathub (manifest elle yazılır, YZ beyanı gerekir), Snap, Certum
OSS kod imzası (€49/yıl), Apple Developer + Homebrew ($99/yıl).

## Microsoft Store

- **Paket:** `OH_STORE=1 pnpm --filter @officehere/desktop dist -- --win appx --x64 --arm64 --publish never`
  (yalnız Windows'ta; `makeappx.exe` Windows SDK aracıdır). Çıktı:
  `OfficeHere-<sürüm>-magaza-x64.appx` ve `-arm64.appx`. Paket imzasız üretilir; mağaza yeniden imzalar.
- **Mağaza sürümü farkı:** `OH_STORE=1` pakete `"ohDagitim": "store"` işaretini koyar
  (`scripts/dist.mjs` → `-c.extraMetadata`). `updates.ts` bu işareti (ya da AppX kabında
  `process.windowsStore`'u) görünce güncelleştirme denetimini tümden kapatır, electron-updater'ı
  yüklemez; Seçenekler'de "Bu sürüm uygulama mağazasından güncellenir; OfficeHere kendisi
  güncelleştirme denetlemez." yazar. Mağaza sürümü bu yüzden **hiç ağ isteği yapmaz** (yazım
  sözlüğü onayı dışında).
- **Kimlik:** `electron-builder.yml` → `appx` bölümündeki `identityName` ve yayıncı yer tutucudur.
  Partner Center'da ad ayrıldıktan sonra **Ürün yönetimi → Ürün kimliği** sayfasındaki
  `Package/Identity/Name` ve `Package/Identity/Publisher` değerleri depo **değişkeni** olarak girilir
  (gizli değildir): `STORE_KIMLIK`, `STORE_YAYINCI`. Değişken yoksa paket yer tutucularla üretilir
  (yerel deneme için yeterli, mağazaya yüklenemez).
- **Görseller:** `apps/desktop/build/appx/` (`scripts/icons.mjs` üretir). Yoksa electron-builder örnek
  görseller koyar ve mağaza sertifikasyonu "varsayılan görsel" diye reddeder.
- **Sürüm numarası:** mağaza `X.Y.Z.0` ister; `0.2.0-beta.1` de `0.2.0.0` olur. Aynı numara iki kez
  gönderilemeyeceği için mağazaya **yalnız kararlı sürümler** yüklenir.
- **Gönderim (elle):** sürüm iş akışı bittikten sonra Actions → çalıştırma → `magaza-appx` artefaktını
  indir → Partner Center → yeni gönderim → Paketler'e iki `.appx`'i yükle. Mağaza listesi (Türkçe):
  açıklamada Microsoft ürün adı geçmez, yalnız uzantılar (`.docx`, `.xlsx`, `.pptx`). Gizlilik ilkesi
  bağlantısı: README "Ağ erişimi" bölümü; gizlilik metnine "güncelleme denetimi Store sürümünde yok"
  cümlesi girer. `runFullTrust` gerekçesi: "Masaüstü uygulaması; kullanıcının seçtiği dosyaları açar ve
  kaydeder."
- **Deneme derlemesi (yayımsız):** Actions → Sürüm → Run workflow, `surum`: `deneme` (v ile
  başlamayan her değer). Paketler, mağaza paketi ve Pardus sınaması üretilir; sürüm yayımlanmaz.
- **Elle doğrulanacak (Windows):** dosya ilişkilendirmeleri, atlama listesi görevleri ve "Klasörde
  göster" mağaza kabında çalışıyor mu (MSIX kabının dosya sistemi sanallaştırması).

## winget

- **Kimlik:** `FurkanYildirim.OfficeHere`. Manifestler `.github/winget/` altındadır (v0.1.6; şema
  1.10.0; x64 ve arm64 NSIS, kullanıcı `/currentuser` ve makine `/allusers` kapsamı, sessiz kurulum
  `/S`). Özetler sürümün cosign imzalı `SHA256SUMS.txt`'iyle aynıdır.
- **İlk gönderim (bir kez, elle):** winget-pkgs'te paket yokken `komac update` çalışmaz.
  - Yol 1: `microsoft/winget-pkgs` çatalında `manifests/f/FurkanYildirim/OfficeHere/0.1.6/` dizinine
    dört dosyayı koyup PR açın; Windows'ta önce `winget validate --manifest .github/winget` ve
    `winget install --manifest .github/winget` ile deneyin.
  - Yol 2 (daha yeni bir kararlı sürüm çıktıysa): `komac new FurkanYildirim.OfficeHere --version <v>
    --urls <x64.exe> <arm64.exe>` (alanları `.github/winget`'tekilerle doldurun) `--submit`.
- **Sonraki sürümler (otomatik):** `release.yml` → `winget` işi yalnız kararlı sürümde, yayından sonra
  `komac update` ile PR açar (komac 2.17.1, özet sabitli). `WINGET_TOKEN` yoksa iş başarıyla atlanır;
  sürüm akışı düşmez.
- **Belirteç:** klasik GitHub belirteci, yalnız `public_repo` kapsamı (komac ince taneli belirteçle PR
  açamıyor). `winget` ortamının gizlisi; ortamda "Required reviewers: Furkan" açık olursa her sürümde
  onay ister.
- winget, SmartScreen'i aşmaz; yalnız kurulumu kolaylaştırır. Moderatör onayı ve Defender taraması
  vardır.

## apt deposu (Pardus, Debian, Ubuntu)

- **Betikler:** `scripts/apt/depo-uret.sh` (reprepro; `Release`, `Release.gpg`, `InRelease`; açık
  anahtar ve deb822 `.sources` üretir, yayından önce imzayı ve özetleri `gpgv` ile doğrular),
  `scripts/apt/yukle.sh` (Cloudflare R2, S3 API; önce `pool/`, sonra `dists/`), `scripts/apt/sina.sh`
  (geçici test anahtarıyla uçtan uca: yerel HTTP, `Signed-By`, `apt-get install officehere`, yabancı
  anahtar ve bozuk imza reddi), `scripts/apt/duman.sh` (Xvfb'de açılış).
- **Depo:** tek dağıtım `stable`, bileşen `main`, `amd64` (ve boş `arm64`). Her sürümde baştan üretilir
  (durumsuz): dizinde yalnız son sürüm vardır; eski `.deb`'ler kovada kalır. `NotAutomatic: yes` +
  `ButAutomaticUpgrades: yes`: depo önceliği 100'dür, dağıtımın paketlerinin önüne geçmez, kurulu
  `officehere` güncellenir (Debian "UseThirdParty" önerisi).
- **Kullanıcı kurulumu** (alan adı belli olunca README'ye girer):

  ```sh
  sudo install -d -m 0755 /etc/apt/keyrings
  curl -fsSL https://apt.ALAN-ADI/officehere.pgp | sudo tee /etc/apt/keyrings/officehere.pgp >/dev/null
  curl -fsSL https://apt.ALAN-ADI/officehere.sources | sudo tee /etc/apt/sources.list.d/officehere.sources >/dev/null
  sudo apt update && sudo apt install officehere
  ```

  `officehere.sources`:

  ```
  Types: deb
  URIs: https://apt.ALAN-ADI
  Suites: stable
  Components: main
  Architectures: amd64
  Signed-By: /etc/apt/keyrings/officehere.pgp
  ```

- **Anahtar düzeni:** birincil anahtar yalnız sertifika (`[C]`) ve çevrimdışı; CI yalnız imzalama alt
  anahtarını alır. Alt anahtar sızarsa birincil anahtarla iptal edilip yenisi eklenir; kullanıcılar
  aynı `officehere.pgp`'yi yeniden indirir.

  ```sh
  gpg --quick-gen-key "OfficeHere apt <posta@furkanyildirim.com>" ed25519 cert 5y
  gpg --quick-add-key <BİRİNCİL-PARMAK-İZİ> ed25519 sign 2y
  gpg --armor --export-secret-keys <BİRİNCİL-PARMAK-İZİ> > cevrimdisi-yedek.asc   # çevrimdışı sakla
  gpg --armor --export-secret-subkeys <BİRİNCİL-PARMAK-İZİ> > ci-alt-anahtar.asc    # APT_GPG_ANAHTARI
  ```

  `ci-alt-anahtar.asc` parolasız olmalıdır (CI'da girilemez); gizliye koyduktan sonra silin.
- **CI (`release.yml` → `apt`):** yalnız kararlı sürümde, yayından sonra. `apt` ortamı:
  gizliler `APT_GPG_ANAHTARI`, `R2_ERISIM_ANAHTARI`, `R2_GIZLI_ANAHTAR`; değişkenler `APT_ANAHTAR`
  (birincil parmak izi), `APT_ADRES` (`https://apt.ALAN-ADI`), `R2_HESAP`, `R2_KOVA`. Biri eksikse iş
  atlanır.
- **R2:** kova (ör. `officehere-apt`) → Settings → Custom Domains → `apt.ALAN-ADI` (herkese açık
  okuma). API belirteci: yalnız bu kova, "Object Read & Write". Ücretsiz katman 10 GB-ay ve çıkış
  trafiği ücretsiz; `.deb` ~100 MB olduğundan ~90 sürüm sığar, eskiler arada elle silinir (ya da kova
  yaşam döngüsü kuralı).

## Pardus

- **Sınama (`pardus.yml`):** `pardus/yirmiuc:latest` (Pardus 23.5, Debian 12) ve
  `pardus/yirmibes:latest` (Pardus 25.2, Debian 13) kapsayıcılarında `.deb` test anahtarıyla imzalı
  yerel depodan `apt-get install officehere` ile kurulur; Xvfb ve grafik kitaplıkları kurulmadan önce
  `ldd` ile bildirilmemiş bağımlılık aranır; sonra başlangıç ekranının yüklendiği ve yerel günlüğe
  açılış yazıldığı denetlenir. Masaüstü dosyalarına dokunan PR'larda, `main`'de, her sürümde
  (`release.yml` → `pardus`, yayını engellemez) ve elle (`surum` ile yayımlanmış bir sürüm) çalışır.
- **Bulunan ve düzeltilen hata:** yayımlanan 0.1.6 `.deb`'i en küçük Pardus 23/25 kurulumunda
  açılmıyordu (`libgbm.so.1` yok); `deb.depends`'e `libgbm1` eklendi. Masaüstü Pardus'ta Mesa kurulu
  olduğu için görünmüyordu.
- **Kapsayıcı notu:** GitHub kapsayıcısı kullanıcı ad alanı ve CAP_SYS_ADMIN vermediği için CI'da
  Chromium korumalı alanı kapatılır (`OH_DUMAN_KORUMASIZ=1`). Yerel chroot denemesinde Pardus 23 ve
  25'te SUID korumalı alanla açıldı.
- **Başvuru (apt deposu yayına girdikten sonra, K12):** [Uygulama Öner](https://apps.pardus.org.tr/suggestapp)
  formu (Uygulama Adı: OfficeHere; Web Sitesi: depo adresi; Neden: Türkçe arayüz, çevrimdışı, MIT,
  hesap/telemetri yok, `.docx/.xlsx/.pptx` + ODF, resmî yazışma şablonları, imzalı apt deposu) +
  dev@pardus.org.tr'ye kısa paketleme teklifi ("`.deb`'i ve imzalı depoyu biz sağlıyoruz; Pardus 23/25
  CI'da sınanıyor") + forum.pardus.org.tr duyurusu.

## Atlama listesi, Dock ve masaüstü eylemleri

"Yeni Belge", "Yeni Tablo", "Yeni Sunu" kısayolları `apps/desktop/src/main/tasks.ts`'tedir:
Windows'ta atlama listesi görevi (görev çubuğu ve Başlat menüsünde sağ tık; simgeler
`resources/filetypes/*.ico`), macOS'ta Dock menüsü, Linux'ta `.deb` masaüstü girdisinin
eylemleri (`electron-builder.yml` → `deb.desktop`). Görev uygulamayı `--yeni=belge|tablo|sunum` ile
başlatır; uygulama açıksa istek tek örnek kilidiyle çalışan örneğe iletilir ve o uygulamanın boş
penceresi açılır. Mağaza kabında görevlerin çalıştığı Windows'ta elle doğrulanmalıdır.

## Maliyetler

| Kalem | Tutar | Not |
| --- | --- | --- |
| Microsoft Store bireysel hesap | 0 | Kimlik + özçekim doğrulaması; günler sürebilir |
| winget | 0 | Klasik belirteç |
| Cloudflare R2 + alt alan adı | 0 | 10 GB-ay ücretsiz; çıkış ücretsiz; mevcut alan adı |
| GitHub Actions | 0 | Herkese açık depo |
| Certum OSS kod imzası | €49/yıl | Ertelendi (3. aşama) |
| Apple Developer | $99/yıl | Ertelendi |

## Furkan'ın yapacakları

1. **Microsoft Store:** Partner Center bireysel hesabı (0 ₺; kimlik + özçekim; kayıt formunda Türkiye
   pazarını denetleyin). "OfficeHere" adını ayırın. Ürün kimliği sayfasındaki `Name` ve `Publisher`'ı
   depo değişkenleri `STORE_KIMLIK`, `STORE_YAYINCI` olarak girin. İlk kararlı sürümden sonra
   `magaza-appx` artefaktıyla gönderim yapın (liste metni, gizlilik bağlantısı, yaş derecelendirmesi).
2. **winget:** klasik belirteç (`public_repo`) → `winget` ortamı (Required reviewers: siz) → gizli
   `WINGET_TOKEN`. İlk gönderimi `.github/winget` manifestleriyle elle yapın (yukarıda).
3. **apt deposu:** GPG anahtarı (birincil çevrimdışı + imzalama alt anahtarı; birincilin çevrimdışı
   yedeği), alan adı seçimi (`apt.<alan>`), R2 kovası + özel alan adı + kovaya özel API belirteci;
   `apt` ortamına gizliler ve değişkenler.
4. **Pardus başvurusu:** apt deposu yayına girdikten sonra form + e-posta + forum (yukarıdaki metin).
5. **Windows'ta elle doğrulama:** NSIS kurulumunda atlama listesi (Yeni Belge/Tablo/Sunu), mağaza
   paketinde dosya ilişkilendirmeleri ve görevler.
