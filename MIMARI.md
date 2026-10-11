# Mimari

```
officehere/
├── packages/
│   ├── core/       @officehere/core     Çerçeveden bağımsız ortak kod
│   ├── ui/         @officehere/ui       Ofis arayüz kiti (React)
│   └── formula/    @officehere/formula  Formül motoru (saf TypeScript)
├── apps/
│   ├── belge/      @officehere/belge    Kelime işlemci (Vite + React + Tiptap/ProseMirror)
│   ├── tablo/      @officehere/tablo    Hesap tablosu (Vite + React + canvas ızgara)
│   ├── sunum/      @officehere/sunum    Sunu (Vite + React + HTML/SVG slayt çizici)
│   └── desktop/    @officehere/desktop  Electron kabuğu + başlangıç ekranı
└── scripts/        Web sürümü derleme
```

## Katmanlar

**core** — Dosya açma/kaydetme tek bir arayüzden yapılır (`openFile`, `saveFile`,
`saveFileAs`, `readPath`). Masaüstünde Electron köprüsü (`window.officehereNative`),
tarayıcıda File System Access API ya da indirme kullanılır. Ayrıca OOXML birimleri
(EMU, twip, pt), tema renkleri, yazı tipi yedek zincirleri ve geri al yığını
buradadır.

**ui** — Başlık çubuğu, şerit, gruplar, düğmeler, açılan kutular, renk paleti,
galeri, Dosya görünümü (`OfficeBackstage`), durum çubuğu, iletişim kutuları ve
temalar. Uygulama vurgusu `<html data-app>`, tema `<html data-oh-theme>` ile
seçilir. `useDocumentLifecycle` belgeyi masaüstü kabuğuna bağlar (başlık, kirli
durum, kapatmadan önce sorma, `?file=` ile açma).

**formula** — Excel formül dilinin ayrıştırıcısı, değerlendiricisi, bağımlılık
grafiği ve fonksiyon kitaplığı. İçeride kanonik İngilizce ad ve `,` ayırıcı
kullanılır; Türkçe yerel ayar katmanı `TOPLA(A1;B1)` biçimini çevirir.

**apps/** — Her uygulama bağımsız bir Vite uygulamasıdır; tarayıcıda tek başına
çalışır, masaüstünde Electron penceresine yüklenir. Dosya biçimi kodu her
uygulamanın `src/io/` dizinindedir.

## Masaüstü kabuğu

```
Electron ana süreç (apps/desktop/src/main)
  ├─ tek örnek kilidi, dosya ilişkilendirme (docx/xlsx/csv/pptx)
  ├─ pencere başına bir belge; uzantıya göre uygulama seçimi
  ├─ IPC: dosya diyalogları, okuma/yazma (geçici dosya + taşıma), PDF dışa aktarma,
  │       son kullanılanlar, yeni pencere, başlık/kirli durum, kapatma onayı
  └─ preload (sandbox) → window.officehereNative
```

Paketlenmiş uygulamada web derlemeleri `resources/apps/<uygulama>/` altındadır.
Geliştirmede `pnpm dev:desktop` Vite sunucularını (5173–5176) açar ve Electron'u
`OH_DEV=1` ile başlatır.

## Dağıtım

- **Masaüstü:** `pnpm dist:win` (NSIS kurulum + taşınabilir), `pnpm dist:linux`
  (AppImage + deb), `pnpm dist:mac` (dmg). `v*` etiketi gönderildiğinde GitHub
  Actions üç platformun paketlerini üretip sürüme ekler.
- **Web:** `pnpm web` → `dist-web/` (başlangıç ekranı kökte, uygulamalar alt
  dizinlerde); herhangi bir statik sunucuyla ya da `Dockerfile` ile yayınlanır.
