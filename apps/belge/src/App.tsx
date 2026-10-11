// Geçici iskelet: arayüz kitinin örnek kullanımı. Belge editörü bunun yerine gelecek.
import { useState } from "react";
import {
  ArrowRedo20Regular,
  ArrowUndo20Regular,
  ClipboardPaste32Regular,
  Copy20Regular,
  Cut20Regular,
  PaintBrush20Regular,
  Save20Regular,
  TextBold20Regular,
  TextItalic20Regular,
  TextUnderline20Regular,
  TextStrikethrough20Regular,
  TextColor20Regular,
  Highlight20Regular,
  TextAlignLeft20Regular,
  TextAlignCenter20Regular,
  TextAlignRight20Regular,
  TextAlignJustify20Regular,
  TextBulletListLtr20Regular,
  TextNumberListLtr20Regular,
  Search20Regular,
  ArrowSwap20Regular,
  DocumentOnePage20Regular,
  ReadingModeMobile20Regular,
  Globe20Regular,
  Table32Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import {
  AppShell,
  Backstage,
  BackstageHeading,
  ColorSplitButton,
  Menu,
  MenuItem,
  QatButton,
  Ribbon,
  RibbonButton,
  RibbonCombo,
  RibbonGallery,
  RibbonGroup,
  RibbonMenuButton,
  RibbonRow,
  RibbonSplitButton,
  RibbonStack,
  StatusBar,
  StatusItem,
  TableGridPicker,
  TemplateRow,
  TemplateTile,
  TitleBar,
  ViewSwitcher,
  ZoomSlider,
} from "@officehere/ui";

export function App() {
  const [tab, setTab] = useState("home");
  const [zoom, setZoom] = useState(100);
  const [bs, setBs] = useState(false);
  const [font, setFont] = useState("Aptos");
  const [size, setSize] = useState("11");
  const [color, setColor] = useState<string | null>("#C00000");

  const home = (
    <>
      <RibbonGroup label="Pano" onLauncher={() => {}}>
        <RibbonSplitButton size="large" icon={<ClipboardPaste32Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => {}} menu={<Menu><MenuItem label="Kaynak Biçimlendirmesini Koru" /><MenuItem label="Yalnızca Metni Koru" /></Menu>} />
        <RibbonStack>
          <RibbonButton icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" />
          <RibbonButton icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" />
          <RibbonButton icon={<PaintBrush20Regular />} label="Biçim Boyacısı" />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Yazı Tipi" onLauncher={() => {}}>
        <RibbonStack>
          <RibbonRow>
            <RibbonCombo value={font} width={130} options={FONT_LIST.map((f) => ({ value: f, style: { fontFamily: fontStack(f) } }))} onCommit={setFont} tip="Yazı Tipi" shortcut="Ctrl+Shift+F" />
            <RibbonCombo value={size} width={48} options={FONT_SIZES.map(String)} onCommit={setSize} tip="Yazı Tipi Boyutu" />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked />
            <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" />
            <RibbonButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" />
            <RibbonButton size="small" icon={<TextStrikethrough20Regular />} tip="Üstü Çizili" />
            <ColorSplitButton icon={<Highlight20Regular />} tip="Metin Vurgu Rengi" color="#FFFF00" variant="highlight" noneLabel="Renk Yok" onApply={() => {}} />
            <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={color} noneLabel="Otomatik" onApply={setColor} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Paragraf" onLauncher={() => {}}>
        <RibbonStack>
          <RibbonRow>
            <RibbonMenuButton size="small" icon={<TextBulletListLtr20Regular />} tip="Madde İşaretleri" menu={<Menu><MenuItem label="●" /></Menu>} />
            <RibbonMenuButton size="small" icon={<TextNumberListLtr20Regular />} tip="Numaralandırma" menu={<Menu><MenuItem label="1." /></Menu>} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" checked />
            <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" />
            <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" />
            <RibbonButton size="small" icon={<TextAlignJustify20Regular />} tip="İki Yana Yasla" />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Stiller" onLauncher={() => {}}>
        <RibbonGallery
          columns={5}
          items={["Normal", "Aralık Yok", "Başlık 1", "Başlık 2", "Konu Başlığı", "Alt Başlık", "Vurgu"].map((s, i) => ({
            id: s,
            label: s,
            preview: <span style={{ fontSize: i === 4 ? 16 : 12, color: i >= 2 && i <= 3 ? "#0F4761" : undefined }}>AaBbÇç<br /><small>{s}</small></span>,
          }))}
          selectedId="Normal"
          onSelect={() => {}}
        />
      </RibbonGroup>
      <RibbonGroup label="Düzenleme">
        <RibbonStack>
          <RibbonButton icon={<Search20Regular />} label="Bul" />
          <RibbonButton icon={<ArrowSwap20Regular />} label="Değiştir" />
        </RibbonStack>
      </RibbonGroup>
    </>
  );

  const insert = (
    <RibbonGroup label="Tablolar">
      <RibbonMenuButton size="large" icon={<Table32Regular />} label="Tablo" menu={<div className="oh-menu"><TableGridPicker onPick={() => {}} /></div>} />
    </RibbonGroup>
  );

  return (
    <AppShell
      app="belge"
      titleBar={
        <TitleBar
          app="belge"
          title="Belge1"
          dirty
          onSearch={() => {}}
          quickAccess={
            <>
              <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={() => {}} />
              <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" onClick={() => {}} />
              <QatButton icon={<ArrowRedo20Regular />} tip="Yinele" shortcut="Ctrl+Y" onClick={() => {}} disabled />
            </>
          }
        />
      }
      backstage={
        bs ? (
          <Backstage items={[{ id: "home", label: "Giriş" }, { id: "new", label: "Yeni" }, { id: "open", label: "Aç" }, { id: "save", label: "Kaydet", separator: true, onClick: () => {} }, { id: "options", label: "Seçenekler", bottom: true }]} activeId="home" onSelect={() => {}} onBack={() => setBs(false)}>
            <BackstageHeading>İyi akşamlar</BackstageHeading>
            <TemplateRow>
              <TemplateTile label="Boş belge" preview={<div />} onClick={() => setBs(false)} />
            </TemplateRow>
          </Backstage>
        ) : undefined
      }
      ribbon={
        <Ribbon
          tabs={[
            { id: "home", label: "Giriş", content: home },
            { id: "insert", label: "Ekle", content: insert },
            { id: "design", label: "Tasarım", content: <div /> },
            { id: "layout", label: "Düzen", content: <div /> },
            { id: "references", label: "Başvurular", content: <div /> },
            { id: "review", label: "Gözden Geçir", content: <div /> },
            { id: "view", label: "Görünüm", content: <div /> },
            { id: "help", label: "Yardım", content: <div /> },
          ]}
          activeTab={tab}
          onTabChange={setTab}
          onFileClick={() => setBs(true)}
          onCollapsedChange={() => {}}
        />
      }
      statusBar={
        <StatusBar
          left={
            <>
              <StatusItem>Sayfa 1 / 1</StatusItem>
              <StatusItem>0 sözcük</StatusItem>
              <StatusItem>
                <Globe20Regular style={{ width: 14 }} /> Türkçe (Türkiye)
              </StatusItem>
            </>
          }
          right={
            <>
              <ViewSwitcher
                active="print"
                onChange={() => {}}
                views={[
                  { id: "read", icon: <ReadingModeMobile20Regular />, tip: "Okuma Modu" },
                  { id: "print", icon: <DocumentOnePage20Regular />, tip: "Yazdırma Düzeni" },
                  { id: "web", icon: <Globe20Regular />, tip: "Web Düzeni" },
                ]}
              />
              <ZoomSlider value={zoom} onChange={setZoom} />
            </>
          }
        />
      }
    >
      <div className="oh-scroll" style={{ flex: 1, overflow: "auto", display: "flex", justifyContent: "center", padding: 24 }}>
        <div style={{ width: 794, minHeight: 1123, background: "#fff", boxShadow: "0 0 0 1px #d4d4d4, 0 2px 6px rgba(0,0,0,.12)", padding: 96, color: "#000", fontFamily: fontStack(font), fontSize: `${size}pt` }}>
          <p>Merhaba dünya.</p>
        </div>
      </div>
    </AppShell>
  );
}
