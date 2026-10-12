# @officehere/ui

Belge, Tablo ve Sunum uygulamalarının ortak ofis arayüz kiti: başlık çubuğu,
şerit (ribbon), Dosya görünümü (backstage), durum çubuğu, menüler, renk
paleti, galeri ve iletişim kutuları. Uygulamalar `@officehere/ui/styles.css`
dosyasını bir kez içe aktarır.

## İskelet

```tsx
<AppShell
  app="belge"                       // "belge" | "tablo" | "sunum" — vurgu rengini seçer
  titleBar={<TitleBar app="belge" title="Belge1" dirty quickAccess={<QatButton … />} onSearch={…} />}
  ribbon={<Ribbon tabs={[…]} activeTab={tab} onTabChange={setTab} onFileClick={openBackstage} onCollapsedChange={…} collapsed={…} />}
  belowRibbon={<MessageBar … /> /* ya da formül çubuğu */}
  statusBar={<StatusBar left={…} right={<><ViewSwitcher … /><ZoomSlider … /></>} />}
  backstage={backstageOpen ? <Backstage …>…</Backstage> : undefined}
>
  {/* çalışma alanı: flex sütun, kalan yüksekliği doldurur */}
</AppShell>
```

## Şerit

| Bileşen | Kullanım |
| --- | --- |
| `Ribbon` | `tabs: { id, label, content, contextual?, hidden? }[]`. Sekmeye çift tık şeridi daraltır. |
| `RibbonGroup` | `label`, `onLauncher` (sağ alttaki iletişim kutusu oku), `icon` (pencere daralınca gösterilen simge). |
| `RibbonStack` / `RibbonRow` | Küçük denetimleri 3 satıra kadar alt alta / yan yana dizer. |
| `RibbonButton` | `size: "large" \| "medium" \| "small"`, `icon`, `label`, `checked`, `disabled`, `tip`, `shortcut`, `description`, `onClick`. |
| `RibbonSplitButton` | Ana eylem + menü. `colorBar` simge altına renk şeridi çizer. |
| `RibbonMenuButton` | Tıklanınca menü açar. `menu` ReactNode ya da `(close) => ReactNode`. |
| `RibbonCombo` | Yazılabilir açılan kutu (yazı tipi, boyut). `onCommit(value)`. |
| `RibbonSpinner` | Sayı kutusu (aralık, girinti). |
| `RibbonCheckbox` | Görünüm sekmesindeki onay kutuları. |
| `RibbonGallery` | Stil/tema galerisi; oklarla kaydırma, açılır tam liste. |
| `ColorSplitButton` | Yazı tipi rengi / vurgulama / dolgu. `variant="highlight"` vurgu paleti. |
| `TableGridPicker` | "Tablo Ekle" ızgarası (menü içinde). |

Büyük düğme simgeleri 32 px, küçükler 16 px gösterilir; Fluent simgelerinin
herhangi bir boyutu verilebilir, CSS ölçekler. Düğmeler fareyle basılınca
odağı editörden çalmaz.

## Menüler ve katmanlar

- `Menu`, `MenuItem` (`icon`, `label`, `shortcut`, `checked`, `submenu`, `keepOpen`), `MenuSeparator`, `MenuHeader`, `MenuSection`.
- `useContextMenu()` → `{ open(event, content), element }`; `element` bileşende bir kez çizilir.
- `Popover` herhangi bir öğeye ya da noktaya bağlı açılır katmandır. `onClose(reason)` kapanma
  nedenini (`"escape"`, `"outside"`, `"blur"`, `"arrow"`, `"tab"`) verir; `focusOnOpen` açılış odağını seçer.
- İpucu için öğeye `tipProps(başlık, kısayol, açıklama)` yaymak yeterlidir.

## Klavye

- Şerit sekmeleri: Tab etkin sekmeye gelir; ←/→ sekme değiştirir (uçlarda döner), Home/End; ↓ ya da Tab
  gövdedeki ilk denetime iner. Gövdede denetimler normal Tab sırasındadır.
- Menü düğmeleri, bölünmüş düğme oku ve daraltılmış grup: Enter/Space/↓ açar ve ilk öğeye odaklanır;
  kapanınca odak düğmeye döner. Fareyle açılan menüde odak menü kabındadır, öğe vurgulanmaz.
- Menü içeriği (`Menu` ya da düz `.oh-menu` kabı fark etmez, `Popover` sağlar): ↑/↓ (devre dışı ve
  ayraçlar atlanır, uçlarda döner; renk paleti ve galeri ızgarasında konuma göre), Home/End, harfle atlama
  (Türkçe), Enter/Space çalıştırır, → alt menüyü açar, ← ya da Esc alt menüyü kapatır, Esc en dış menüyü
  kapatıp odağı geri verir, Tab menüyü kapatır. Yazı alanı içeren katmanlarda Tab içeride döner.
- `RibbonCombo`: Alt+↓ ya da ↓ listeyi açar, ↑/↓ etkin seçeneği taşır (`aria-activedescendant`), Enter seçer,
  Esc kapatıp önceki değeri korur.
- `RibbonGallery`: tek Tab durağı; ←/→/↑/↓, Home/End, Alt+↓ galeriyi genişletir.
- Saf yardımcılar (`menuNav.ts`: sonraki etkin öğe, harfle eşleşme, ızgara yönü) birim testlidir.

## Dosya görünümü

`Backstage` (`items`, `activeId`, `onSelect`, `onBack`) içinde `BackstageHeading`,
`BackstageSubheading`, `TemplateRow` + `TemplateTile`, `RecentFileList`,
`BackstageAction`, `PropertyList` kullanılır.

## İletişim kutuları

`Dialog` (`open`, `title`, `onClose`, `footer`, `onSubmit`), `DialogButton`,
`Field`, `FieldSet`, `DialogTabs`, `SaveChangesDialog`, `MessageBar`.

## Kancalar

- `useHotkeys({ "Mod+S": save, "Mod+Shift+Z": redo, F12: saveAs })` — `Mod` Ctrl/Cmd.
- `useUiTheme()` → Renkli / Beyaz / Koyu Gri / Siyah / Sistem; tüm uygulamalarda ortak.

## Hazır Dosya görünümü ve belge yaşam döngüsü

- `OfficeBackstage` — üç uygulamada ortak Dosya menüsü (Giriş, Yeni, Aç, Bilgi,
  Kaydet, Farklı Kaydet, Yazdır, Dışarı Aktar, Kapat, Hakkında, Seçenekler).
  Uygulama yalnız şablonları, dışa aktarma biçimlerini ve eylemleri verir.
- `useDocumentLifecycle({ title, dirty, onOpenPath, onSave })` → `{ dialog, confirmDiscard }`.
  Masaüstü kabuğuna başlık/kirli durumunu bildirir, pencere kapanırken
  "Değişiklikler kaydedilsin mi?" sorar, `?file=<yol>` ve dosya ilişkilendirme
  isteklerini `onOpenPath`'e iletir; tarayıcıda sayfadan ayrılma uyarısı verir.
  `confirmDiscard()` "Yeni" / "Aç" / "Kapat" öncesinde çağrılır.

## Komut araması

`CommandSearch` başlık çubuğunun arama yuvasına (`.oh-titlebar-search`; `TitleBar`'a
`onSearch` verilmez) yerleşir. `tabs` şerit sekmeleridir: içerikleri görünmeyen bir
kapta bir kez çizilip ipuçlu (`data-tip`) öğelerden dizin çıkarılır; seçilen komutun
sekmesine geçilip düğmesine tıklanır. `extra` şeritte olmayan komutları verir
(`fileCommands(...)` Yeni, Aç, Kaydet, Farklı Kaydet, PDF, Yazdır, Seçenekler…), `find`
listenin sonundaki "bul" satırıdır; `findOnEnter` ile seçim yokken Enter onu çalıştırır.
Son kullanılan 5 komut `localStorage`'da uygulama başına (`oh-cmd-recent:<app>`) tutulur
ve kutu boşken listelenir. Alt+Q kutuya odaklanır.
