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
  ArrowPrevious20Regular,
  ArrowNext20Regular,
  Link20Regular,
  Dismiss32Regular,
  ReadingModeMobile24Regular, TextBulletListTree24Regular, Drafts24Regular, DarkTheme24Regular } from "@fluentui/react-icons";
import { getNative } from "@officehere/core";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack, ReportBugButton } from "@officehere/ui";
import { useBelge } from "../state/context";
import { FOOTER_PRESETS, HEADER_PRESETS, pageNumberDoc } from "../templates/parts";
import { hfKey, hfVariantOf } from "../views/PageSheets";
import { hfLinked, sectionIds, setHfLink, withHfContent } from "../editor/sectionHf";
import { applyPageSetup, sectionPageAt } from "../state/pageSetup";
import type { PageSetup } from "../model/document";
import { setDarkPage, useDarkPage } from "./darkPage";

export function ViewTab({ zoomTo }: { zoomTo: (kind: "100" | "onePage" | "multi" | "width") => void }) {
  const ctx = useBelge();
  const v = ctx.view;
  const darkPage = useDarkPage();
  return (
    <>
      <RibbonGroup label="Görünümler" icon={<DocumentOnePage32Regular />}>
        <RibbonButton size="large" icon={<ReadingModeMobile24Regular />} label="Okuma Modu" checked={v.mode === "read"} onClick={() => ctx.setView({ mode: "read" })} />
        <RibbonButton size="large" icon={<DocumentOnePage32Regular />} label="Yazdırma Düzeni" checked={v.mode === "print"} onClick={() => ctx.setView({ mode: "print" })} />
        <RibbonButton size="large" icon={<Globe32Regular />} label="Web Düzeni" checked={v.mode === "web"} onClick={() => ctx.setView({ mode: "web" })} />
        <RibbonStack>
          <RibbonButton icon={<TextBulletListTree24Regular />} label="Anahat" checked={v.mode === "outline"} onClick={() => ctx.setView({ mode: "outline" })} />
          <RibbonButton icon={<Drafts24Regular />} label="Taslak" checked={v.mode === "draft"} onClick={() => ctx.setView({ mode: "draft" })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tam Ekran" icon={<ScanText32Regular />}>
        <RibbonButton size="large" icon={<ScanText32Regular />} label="Odak" description="Dikkat dağıtıcı öğeleri kaldırıp belgenize odaklanın. Çıkmak için Esc." checked={v.focusMode} onClick={() => ctx.setView({ focusMode: !v.focusMode })} />
      </RibbonGroup>
      <RibbonGroup label="Koyu Mod" icon={<DarkTheme24Regular />}>
        <RibbonButton size="large" icon={<DarkTheme24Regular />} label="Koyu Sayfa" description="Sayfayı ekranda koyu zeminle gösterin. Yazdırma ve PDF etkilenmez, sayfa beyaz basılır." checked={darkPage} onClick={() => setDarkPage(!darkPage)} />
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
      <ReportBugButton app="belge" />
    </RibbonGroup>
  );
}

const CM_PER_TW = 2.54 / 1440;

export function HeaderFooterTab() {
  const ctx = useBelge();
  const { updateSettings, settings, mainEditor, controller } = ctx;
  const which = ctx.headerFooterEditing ?? "header";
  const pageIndex = ctx.headerFooterPage;
  const pg = controller.info.pages[pageIndex];
  const variant = hfVariantOf(pg, pageIndex, settings);
  const secIndex = pg?.section ?? 0;
  const page = sectionPageAt(mainEditor, settings.page, secIndex);
  const m = page.margin;
  const setSection = (fn: (p: PageSetup) => PageSetup) => applyPageSetup(mainEditor, settings, updateSettings, fn, "auto", secIndex);
  // Hazır içerik, düzenlenen türdeki (ilk/çift/varsayılan) üst ya da alt bilgiye uygulanır; düzenleyici yeniden kurulur.
  const setPart = (w: "header" | "footer", json: typeof settings.header) => {
    const key = hfKey(w, variant);
    updateSettings((s) => withHfContent(s, sectionIds(mainEditor.state.doc), secIndex, key, json));
    ctx.editHeaderFooterAt(w, pageIndex);
  };
  // Önceki/Sonraki: türü (bölüm, ilk/çift/tek) farklı olan önceki/sonraki üst/alt bilgiye geçer.
  const pages = controller.info.pages;
  const storyOf = (i: number) => `${pages[i]?.section ?? 0}:${hfVariantOf(pages[i], i, settings)}`;
  const stepTo = (dir: -1 | 1): number | null => {
    const cur = storyOf(pageIndex);
    for (let i = pageIndex + dir; i >= 0 && i < pages.length; i += dir) if (!pages[i]?.blank && storyOf(i) !== cur) return dir < 0 ? firstOfStory(i) : i;
    return null;
  };
  // Önceki türün ilk sayfası (Word gibi önceki bölümün başına gider).
  const firstOfStory = (i: number) => {
    const st = storyOf(i);
    let k = i;
    while (k - 1 >= 0 && storyOf(k - 1) === st) k--;
    return k;
  };
  const prev = stepTo(-1);
  const next = stepTo(1);
  const key = hfKey(which, variant);
  const linked = hfLinked(sectionIds(mainEditor.state.doc), settings, secIndex, key);
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
          <RibbonButton icon={<ArrowUp20Regular />} label="Üst Bilgiye Git" checked={which === "header"} onClick={() => ctx.editHeaderFooter("header")} />
          <RibbonButton icon={<ArrowDown20Regular />} label="Alt Bilgiye Git" checked={which === "footer"} onClick={() => ctx.editHeaderFooter("footer")} />
        </RibbonStack>
        <RibbonStack>
          <RibbonButton icon={<ArrowPrevious20Regular />} label="Önceki" tip="Önceki bölümün ya da sayfa türünün üst/alt bilgisine git" disabled={prev === null} onClick={() => prev !== null && ctx.editHeaderFooterAt(which, prev)} />
          <RibbonButton icon={<ArrowNext20Regular />} label="Sonraki" tip="Sonraki bölümün ya da sayfa türünün üst/alt bilgisine git" disabled={next === null} onClick={() => next !== null && ctx.editHeaderFooterAt(which, next)} />
          <RibbonButton
            icon={<Link20Regular />}
            label="Öncekine Bağla"
            tip="Bu bölümün üst/alt bilgisini önceki bölümdekiyle aynı yapar"
            checked={linked}
            disabled={secIndex === 0}
            onClick={() => {
              if (linked) setHfLink(mainEditor, updateSettings, settings, secIndex, key, false);
              else if (window.confirm("Bu üst/alt bilgi silinip önceki bölümdekine bağlansın mı?")) setHfLink(mainEditor, updateSettings, settings, secIndex, key, true);
            }}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Seçenekler" icon={<DocumentHeader32Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Farklı İlk Sayfa" checked={!!page.titlePage} onChange={(v) => setSection((p) => ({ ...p, titlePage: v || undefined }))} />
          <RibbonCheckbox label="Tek ve Çift Sayfalarda Farklı" checked={settings.oddEven} onChange={(v) => updateSettings((s) => ({ ...s, oddEven: v }))} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Konum" icon={<DocumentHeader32Regular />}>
        <RibbonStack>
          <RibbonSpinner label="Üst Bilgi Üstten:" value={m.header * CM_PER_TW} step={0.1} min={0} max={10} unit="cm" decimals={2} width={76} onChange={(v) => setSection((p) => ({ ...p, margin: { ...p.margin, header: Math.round(v / CM_PER_TW) } }))} />
          <RibbonSpinner label="Alt Bilgi Alttan:" value={m.footer * CM_PER_TW} step={0.1} min={0} max={10} unit="cm" decimals={2} width={76} onChange={(v) => setSection((p) => ({ ...p, margin: { ...p.margin, footer: Math.round(v / CM_PER_TW) } }))} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Kapat" icon={<Dismiss32Regular />}>
        <RibbonButton size="large" className="bl-wide" icon={<Dismiss32Regular />} label="Üst Bilgi ve Alt Bilgiyi Kapat" onClick={() => ctx.editHeaderFooter(null)} />
      </RibbonGroup>
    </>
  );
}
