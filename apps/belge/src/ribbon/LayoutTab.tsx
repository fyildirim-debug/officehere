// Düzen sekmesi: Sayfa Yapısı (bölüm başına), Kesmeler (sayfa/sütun/bölüm sonları) ve Paragraf.
import {
  DocumentMargins24Regular,
  DocumentLandscape24Regular,
  DocumentOnePage32Regular,
  DocumentPageBreak24Regular,
  TextColumnTwo24Regular,
  TextIndentIncreaseLtr20Regular,
  TextIndentDecreaseLtr20Regular,
  TextAddSpaceBefore20Regular,
  TextAddSpaceAfter20Regular,
  Orientation24Regular,
  TextNumberListLtr20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, MenuHeader, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack, RibbonLabel } from "@officehere/ui";
import type { Editor } from "@tiptap/core";
import { useBelge } from "../state/context";
import { MARGIN_PRESETS, PAGE_SIZE_LIST, fmtCm, marginPresetId, pageSizeId, samePage, type DocSettings, type PageSetup, type SectionAttrs } from "../model/document";
import { setParaAttrs } from "../editor/commands";
import type { LineNumbering } from "../model/document";
import { isolateSelection } from "../editor/sections";
import { applyPageSetup, sectionPageAt } from "../state/pageSetup";
import { ArrangeGroup } from "./ArrangeGroup";

const PT_PER_CM = 72 / 2.54;

export function setOrientation(p: PageSetup, o: "portrait" | "landscape"): PageSetup {
  if (p.orientation === o) return p;
  return { ...p, orientation: o, width: p.height, height: p.width };
}

export function setPaperSize(p: PageSetup, id: string): PageSetup {
  const s = PAGE_SIZE_LIST.find((x) => x.id === id);
  if (!s) return p;
  return p.orientation === "landscape" ? { ...p, width: s.h, height: s.w } : { ...p, width: s.w, height: s.h };
}

/**
 * Sütun sayısını uygular. Metin seçiliyse Word'deki gibi seçim sürekli bölüm sonlarıyla ayrı
 * bir bölüme alınır ve yalnız ona uygulanır; değilse imleçteki bölüme (tek bölümde tüm belgeye).
 */
export function applyColumns(editor: Editor, settings: DocSettings, updateSettings: (fn: (s: DocSettings) => DocSettings) => void, fn: (p: PageSetup) => PageSetup) {
  const { selection, doc } = editor.state;
  if (!selection.empty && doc.textBetween(selection.from, selection.to).trim()) {
    const secIdx = selection.$from.index(0);
    const own = (doc.child(secIdx).attrs.page as PageSetup | null) ?? null;
    const mid = isolateSelection(editor.state, editor.view.dispatch, own);
    if (mid !== null) {
      const d = editor.state.doc;
      const node = d.nodeAt(mid);
      if (node) {
        const next = fn((node.attrs.page as PageSetup | null) ?? settings.page);
        editor.view.dispatch(editor.state.tr.setNodeMarkup(mid, undefined, { ...(node.attrs as SectionAttrs), page: samePage(next, settings.page) ? null : next }));
      }
      return;
    }
  }
  applyPageSetup(editor, settings, updateSettings, fn);
}

export function LayoutTab() {
  const ctx = useBelge();
  const { settings, updateSettings, fmt, mainEditor: editor } = ctx;
  // İmleçteki bölümün etkin sayfa yapısı (fmt.sectionKey değişince yenilenir).
  void fmt.sectionKey;
  const page = sectionPageAt(editor, settings.page, fmt.section);
  const setPage = (fn: (p: PageSetup) => PageSetup) => applyPageSetup(editor, settings, updateSettings, fn);
  const mp = marginPresetId(page);
  const sz = pageSizeId(page);
  const chain = () => editor.chain().focus();
  const sectionBreak = (kind: SectionAttrs["kind"]) => chain().insertSectionBreak(kind).run();
  // Satır numaraları imleçteki bölüme (tek bölümde belgeye) uygulanır.
  const ln = page.lineNumbers ?? null;
  const setLineNumbers = (restart: LineNumbering["restart"] | null) => setPage((p) => ({ ...p, lineNumbers: restart ? { ...(p.lineNumbers ?? {}), restart } : null }));

  return (
    <>
      <RibbonGroup label="Sayfa Yapısı" onLauncher={() => ctx.openDialog("pageSetup")} launcherTip="Sayfa Yapısı" icon={<DocumentMargins24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DocumentMargins24Regular />}
          label="Kenar Boşlukları"
          menu={
            <Menu width={280}>
              {MARGIN_PRESETS.map((m) => (
                <MenuItem
                  key={m.id}
                  checked={mp === m.id}
                  label={m.label}
                  description={`Üst: ${fmtCm(m.m.top)}  Alt: ${fmtCm(m.m.bottom)}  Sol: ${fmtCm(m.m.left)}  Sağ: ${fmtCm(m.m.right)}`}
                  onClick={() => setPage((p) => ({ ...p, margin: { ...p.margin, ...m.m } }))}
                />
              ))}
              <MenuSeparator />
              <MenuItem label="Özel Kenar Boşlukları..." onClick={() => ctx.openDialog("pageSetup", "margins")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={page.orientation === "landscape" ? <DocumentLandscape24Regular /> : <Orientation24Regular />}
          label="Yönlendirme"
          menu={
            <Menu>
              <MenuItem icon={<DocumentOnePage32Regular />} label="Dikey" checked={page.orientation === "portrait"} onClick={() => setPage((p) => setOrientation(p, "portrait"))} />
              <MenuItem icon={<DocumentLandscape24Regular />} label="Yatay" checked={page.orientation === "landscape"} onClick={() => setPage((p) => setOrientation(p, "landscape"))} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentOnePage32Regular />}
          label="Boyut"
          menu={
            <Menu width={220}>
              {PAGE_SIZE_LIST.map((s) => (
                <MenuItem key={s.id} checked={sz === s.id} label={s.label} description={`${fmtCm(s.w, 1)} x ${fmtCm(s.h, 1)}`} onClick={() => setPage((p) => setPaperSize(p, s.id))} />
              ))}
              <MenuSeparator />
              <MenuItem label="Diğer Kağıt Boyutları..." onClick={() => ctx.openDialog("pageSetup", "paper")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<TextColumnTwo24Regular />}
          label="Sütunlar"
          menu={
            <Menu>
              {[
                [1, "Bir"],
                [2, "İki"],
                [3, "Üç"],
              ].map(([n, label]) => (
                <MenuItem key={n} label={label as string} checked={page.columns === n} onClick={() => applyColumns(editor, settings, updateSettings, (p) => ({ ...p, columns: n as number }))} />
              ))}
              <MenuSeparator />
              <MenuItem label="Diğer Sütunlar..." onClick={() => ctx.openDialog("columns")} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonMenuButton
            icon={<DocumentPageBreak24Regular />}
            label="Kesmeler"
            menu={
              <Menu width={330}>
                <MenuHeader>Sayfa Sonları</MenuHeader>
                <MenuItem label="Sayfa" description="Bir sayfanın nerede bitip sonraki sayfanın nerede başladığını işaretleyin." shortcut="Ctrl+Enter" disabled={!ctx.inBody} onClick={() => chain().setPageBreak().run()} />
                <MenuItem label="Sütun" description="Sütun sonundan sonraki metin sonraki sütunun başlangıcında yer alır." shortcut="Ctrl+Shift+Enter" disabled={!ctx.inBody} onClick={() => chain().insertColumnBreak().run()} />
                <MenuItem label="Metin Kaydırma" description="Web sayfalarındaki nesnelerin etrafındaki metni ayırın (satır sonu, Shift+Enter)." onClick={() => ctx.editor.chain().focus().setHardBreak().run()} />
                <MenuHeader>Bölüm Sonları</MenuHeader>
                <MenuItem label="Sonraki Sayfa" description="Bölüm sonu ekleyin ve yeni bölümü sonraki sayfada başlatın." disabled={!ctx.inBody} onClick={() => sectionBreak("nextPage")} />
                <MenuItem label="Sürekli" description="Bölüm sonu ekleyin ve yeni bölümü aynı sayfada başlatın." disabled={!ctx.inBody} onClick={() => sectionBreak("continuous")} />
                <MenuItem label="Çift Sayfa" description="Bölüm sonu ekleyin ve yeni bölümü sonraki çift numaralı sayfada başlatın." disabled={!ctx.inBody} onClick={() => sectionBreak("evenPage")} />
                <MenuItem label="Tek Sayfa" description="Bölüm sonu ekleyin ve yeni bölümü sonraki tek numaralı sayfada başlatın." disabled={!ctx.inBody} onClick={() => sectionBreak("oddPage")} />
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<TextNumberListLtr20Regular />}
            label="Satır Numaraları"
            menu={
              <Menu width={280}>
                <MenuItem label="Yok" checked={!ln} onClick={() => setLineNumbers(null)} />
                <MenuItem label="Sürekli" checked={ln?.restart === "continuous"} onClick={() => setLineNumbers("continuous")} />
                <MenuItem label="Her Sayfada Yeniden Başlat" checked={ln?.restart === "newPage"} onClick={() => setLineNumbers("newPage")} />
                <MenuItem label="Her Bölümde Yeniden Başlat" checked={ln?.restart === "newSection"} onClick={() => setLineNumbers("newSection")} />
                <MenuSeparator />
                <MenuItem label="Geçerli Paragrafta Gizle" checked={!!fmt.suppressLineNumbers} onClick={() => setParaAttrs(ctx.editor, { suppressLineNumbers: !fmt.suppressLineNumbers })} />
                <MenuSeparator />
                <MenuItem label="Satır Numaralandırma Seçenekleri..." onClick={() => ctx.openDialog("lineNumbers")} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Paragraf" onLauncher={() => ctx.openDialog("paragraph")} launcherTip="Paragraf Ayarları">
        <RibbonStack>
          <RibbonLabel>Girinti</RibbonLabel>
          <RibbonSpinner icon={<TextIndentIncreaseLtr20Regular />} label="Sol:" value={fmt.indentLeft / PT_PER_CM} step={0.1} min={-20} max={50} unit="cm" decimals={2} width={76} onChange={(v) => setParaAttrs(ctx.editor, { indentLeft: v ? +(v * PT_PER_CM).toFixed(2) : null })} />
          <RibbonSpinner icon={<TextIndentDecreaseLtr20Regular />} label="Sağ:" value={fmt.indentRight / PT_PER_CM} step={0.1} min={-20} max={50} unit="cm" decimals={2} width={76} onChange={(v) => setParaAttrs(ctx.editor, { indentRight: v ? +(v * PT_PER_CM).toFixed(2) : null })} />
        </RibbonStack>
        <RibbonStack>
          <RibbonLabel>Aralık</RibbonLabel>
          <RibbonSpinner icon={<TextAddSpaceBefore20Regular />} label="Önce:" value={fmt.spaceBefore ?? 0} step={6} min={0} max={1584} unit="nk" decimals={1} width={70} onChange={(v) => setParaAttrs(ctx.editor, { spaceBefore: v })} />
          <RibbonSpinner icon={<TextAddSpaceAfter20Regular />} label="Sonra:" value={fmt.spaceAfter ?? (fmt.styleId === "NoSpacing" ? 0 : 8)} step={6} min={0} max={1584} unit="nk" decimals={1} width={70} onChange={(v) => setParaAttrs(ctx.editor, { spaceAfter: v })} />
        </RibbonStack>
      </RibbonGroup>
      <ArrangeGroup />
    </>
  );
}
