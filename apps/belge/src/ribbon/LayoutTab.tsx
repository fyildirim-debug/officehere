// Düzen sekmesi: Sayfa Yapısı ve Paragraf (girinti/aralık).
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
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, MenuHeader, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack, RibbonLabel } from "@officehere/ui";
import { useBelge } from "../state/context";
import { MARGIN_PRESETS, PAGE_SIZE_LIST, fmtCm, marginPresetId, pageSizeId, type PageSetup } from "../model/document";
import { setParaAttrs } from "../editor/commands";

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

export function LayoutTab() {
  const ctx = useBelge();
  const { settings, updateSettings, fmt, editor } = ctx;
  const page = settings.page;
  const setPage = (fn: (p: PageSetup) => PageSetup) => updateSettings((s) => ({ ...s, page: fn(s.page) }));
  const mp = marginPresetId(page);
  const sz = pageSizeId(page);

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
                <MenuItem key={n} label={label as string} checked={page.columns === n} onClick={() => setPage((p) => ({ ...p, columns: n as number }))} />
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
              <Menu width={300}>
                <MenuHeader>Sayfa Sonları</MenuHeader>
                <MenuItem label="Sayfa" description="Bir sayfanın nerede bitip sonraki sayfanın nerede başladığını işaretleyin." shortcut="Ctrl+Enter" onClick={() => editor.chain().focus().setPageBreak().run()} />
                <MenuItem label="Metin Kaydırma" description="Satır sonu (Shift+Enter)." onClick={() => editor.chain().focus().setHardBreak().run()} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Paragraf" onLauncher={() => ctx.openDialog("paragraph")} launcherTip="Paragraf Ayarları">
        <RibbonStack>
          <RibbonLabel>Girinti</RibbonLabel>
          <RibbonSpinner icon={<TextIndentIncreaseLtr20Regular />} label="Sol:" value={fmt.indentLeft / PT_PER_CM} step={0.1} min={-20} max={50} unit="cm" decimals={2} width={76} onChange={(v) => setParaAttrs(editor, { indentLeft: v ? +(v * PT_PER_CM).toFixed(2) : null })} />
          <RibbonSpinner icon={<TextIndentDecreaseLtr20Regular />} label="Sağ:" value={fmt.indentRight / PT_PER_CM} step={0.1} min={-20} max={50} unit="cm" decimals={2} width={76} onChange={(v) => setParaAttrs(editor, { indentRight: v ? +(v * PT_PER_CM).toFixed(2) : null })} />
        </RibbonStack>
        <RibbonStack>
          <RibbonLabel>Aralık</RibbonLabel>
          <RibbonSpinner icon={<TextAddSpaceBefore20Regular />} label="Önce:" value={fmt.spaceBefore ?? 0} step={6} min={0} max={1584} unit="nk" decimals={1} width={70} onChange={(v) => setParaAttrs(editor, { spaceBefore: v })} />
          <RibbonSpinner icon={<TextAddSpaceAfter20Regular />} label="Sonra:" value={fmt.spaceAfter ?? (fmt.styleId === "NoSpacing" ? 0 : 8)} step={6} min={0} max={1584} unit="nk" decimals={1} width={70} onChange={(v) => setParaAttrs(editor, { spaceAfter: v })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
