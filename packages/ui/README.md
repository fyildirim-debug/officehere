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
- `Popover` herhangi bir öğeye ya da noktaya bağlı açılır katmandır.
- İpucu için öğeye `tipProps(başlık, kısayol, açıklama)` yaymak yeterlidir.

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
