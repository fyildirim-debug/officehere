// Görünüm, Yardım ve Üst Bilgi ve Alt Bilgi sekmeleri.
import {
  DocumentOnePage32Regular,
  Globe32Regular,
  ZoomIn32Regular,
  DocumentOnePage20Regular,
  DocumentMultiple20Regular,
  ArrowAutofitWidth20Regular,
  WindowNew32Regular,
  ScanText32Regular,
  Keyboard24Regular,
  Info32Regular,
  PersonFeedback32Regular,
  DocumentHeader32Regular,
  DocumentFooter24Regular,
  DocumentPageNumber24Regular,
  Calendar20Regular,
  Image20Regular,
  ArrowUp20Regular,
  ArrowDown20Regular,
  Dismiss32Regular,
  ReadingModeMobile24Regular,
} from "@fluentui/react-icons";
import { getNative } from "@officehere/core";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { FOOTER_PRESETS, HEADER_PRESETS, pageNumberDoc } from "../templates/parts";

export function ViewTab({ zoomTo }: { zoomTo: (kind: "100" | "onePage" | "multi" | "width") => void }) {
  const ctx = useBelge();
  const v = ctx.view;
  return (
    <>
      <RibbonGroup label="Görünümler" icon={<DocumentOnePage32Regular />}>
        <RibbonButton size="large" icon={<ReadingModeMobile24Regular />} label="Okuma Modu" checked={v.mode === "read"} onClick={() => ctx.setView({ mode: "read" })} />
        <RibbonButton size="large" icon={<DocumentOnePage32Regular />} label="Yazdırma Düzeni" checked={v.mode === "print"} onClick={() => ctx.setView({ mode: "print" })} />
        <RibbonButton size="large" icon={<Globe32Regular />} label="Web Düzeni" checked={v.mode === "web"} onClick={() => ctx.setView({ mode: "web" })} />
      </RibbonGroup>
      <RibbonGroup label="Tam Ekran" icon={<ScanText32Regular />}>
        <RibbonButton size="large" icon={<ScanText32Regular />} label="Odak" description="Dikkat dağıtıcı öğeleri kaldırıp belgenize odaklanın. Çıkmak için Esc." checked={v.focusMode} onClick={() => ctx.setView({ focusMode: !v.focusMode })} />
      </RibbonGroup>
      <RibbonGroup label="Göster" icon={<DocumentOnePage32Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Cetvel" checked={v.ruler} onChange={(x) => ctx.setView({ ruler: x })} />
          <RibbonCheckbox label="Kılavuz Çizgileri" checked={v.gridlines} onChange={(x) => ctx.setView({ gridlines: x })} />
          <RibbonCheckbox label="Gezinti Bölmesi" checked={v.navPane} onChange={(x) => ctx.setView({ navPane: x })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Yakınlaştır" icon={<ZoomIn32Regular />}>
        <RibbonButton size="large" icon={<ZoomIn32Regular />} label="Yakınlaştır" onClick={() => ctx.openDialog("zoom")} />
        <RibbonButton size="large" icon={<span className="bl-pct">%100</span>} label="%100" onClick={() => zoomTo("100")} />
        <RibbonStack>
          <RibbonButton icon={<DocumentOnePage20Regular />} label="Tek Sayfa" onClick={() => zoomTo("onePage")} />
          <RibbonButton icon={<DocumentMultiple20Regular />} label="Birden Çok Sayfa" onClick={() => zoomTo("multi")} />
          <RibbonButton icon={<ArrowAutofitWidth20Regular />} label="Sayfa Genişliği" onClick={() => zoomTo("width")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Pencere" icon={<WindowNew32Regular />}>
        <RibbonButton
          size="large"
          icon={<WindowNew32Regular />}
          label="Yeni Pencere"
          onClick={() => {
            const n = getNative();
            if (n) void n.newWindow("belge");
            else window.open(window.location.href.split("?")[0], "_blank");
          }}
        />
      </RibbonGroup>
    </>
  );
}

export function HelpTab({ onFeedback }: { onFeedback: () => void }) {
  const ctx = useBelge();
  return (
    <RibbonGroup label="Yardım" icon={<Keyboard24Regular />}>
      <RibbonButton size="large" icon={<Keyboard24Regular />} label="Klavye Kısayolları" onClick={() => ctx.openDialog("shortcuts")} />
      <RibbonButton size="large" icon={<Info32Regular />} label="Hakkında" onClick={() => ctx.openBackstage("about")} />
      <RibbonButton size="large" icon={<PersonFeedback32Regular />} label="Geri Bildirim" onClick={onFeedback} />
    </RibbonGroup>
  );
}

const CM_PER_TW = 2.54 / 1440;

export function HeaderFooterTab() {
  const ctx = useBelge();
  const { updateSettings, settings } = ctx;
  const m = settings.page.margin;
  // Hazır içerik uygulanınca açık düzenleyici yeni içerikle yeniden kurulur.
  const setPart = (which: "header" | "footer", json: typeof settings.header) => {
    updateSettings((s) => ({ ...s, [which]: json }));
    ctx.editHeaderFooter(which);
  };
  return (
    <>
      <RibbonGroup label="Üst Bilgi ve Alt Bilgi" icon={<DocumentHeader32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DocumentHeader32Regular />}
          label="Üst Bilgi"
          menu={
            <Menu width={220}>
              {HEADER_PRESETS.map((h) => (
                <MenuItem key={h.id} label={h.label} onClick={() => setPart("header", h.content)} />
              ))}
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentFooter24Regular />}
          label="Alt Bilgi"
          menu={
            <Menu width={220}>
              {FOOTER_PRESETS.map((h) => (
                <MenuItem key={h.id} label={h.label} onClick={() => setPart("footer", h.content)} />
              ))}
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentPageNumber24Regular />}
          label="Sayfa Numarası"
          menu={
            <Menu>
              <MenuItem label="Geçerli Konum" onClick={() => ctx.editor.chain().focus().insertPageField("PAGE").run()} />
              <MenuItem label="Sayfa X / Y (Geçerli Konum)" onClick={() => ctx.editor.chain().focus().insertContent([{ type: "text", text: "Sayfa " }, { type: "pageField", attrs: { field: "PAGE" } }, { type: "text", text: " / " }, { type: "pageField", attrs: { field: "NUMPAGES" } }]).run()} />
              <MenuItem label="Sayfanın Altı (Ortada)" onClick={() => setPart("footer", pageNumberDoc("center"))} />
              <MenuItem label="Sayfanın Üstü (Sağda)" onClick={() => setPart("header", pageNumberDoc("right"))} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Ekle" icon={<Calendar20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<Calendar20Regular />} label="Tarih ve Saat" onClick={() => ctx.openDialog("dateTime")} />
          <RibbonButton icon={<Image20Regular />} label="Resimler" onClick={ctx.insertImageFromFile} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Gezinti" icon={<ArrowUp20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<ArrowUp20Regular />} label="Üst Bilgiye Git" checked={ctx.headerFooterEditing === "header"} onClick={() => ctx.editHeaderFooter("header")} />
          <RibbonButton icon={<ArrowDown20Regular />} label="Alt Bilgiye Git" checked={ctx.headerFooterEditing === "footer"} onClick={() => ctx.editHeaderFooter("footer")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Konum" icon={<DocumentHeader32Regular />}>
        <RibbonStack>
          <RibbonSpinner label="Üst Bilgi Üstten:" value={m.header * CM_PER_TW} step={0.1} min={0} max={10} unit="cm" decimals={2} width={76} onChange={(v) => updateSettings((s) => ({ ...s, page: { ...s.page, margin: { ...s.page.margin, header: Math.round(v / CM_PER_TW) } } }))} />
          <RibbonSpinner label="Alt Bilgi Alttan:" value={m.footer * CM_PER_TW} step={0.1} min={0} max={10} unit="cm" decimals={2} width={76} onChange={(v) => updateSettings((s) => ({ ...s, page: { ...s.page, margin: { ...s.page.margin, footer: Math.round(v / CM_PER_TW) } } }))} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Kapat" icon={<Dismiss32Regular />}>
        <RibbonButton size="large" className="bl-wide" icon={<Dismiss32Regular />} label="Üst Bilgi ve Alt Bilgiyi Kapat" onClick={() => ctx.editHeaderFooter(null)} />
      </RibbonGroup>
    </>
  );
}

