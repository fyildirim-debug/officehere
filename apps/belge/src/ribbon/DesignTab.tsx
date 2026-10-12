// Tasarım sekmesi: Belge Biçimlendirme (temalar, stil kümeleri, renkler, yazı tipleri, paragraf aralığı) ve Sayfa Arka Planı.
import { Color32Regular, DocumentBorder32Regular, TextFont20Regular, Water32Regular, PaintBucket24Regular, TextParagraph20Regular, DesignIdeas32Regular } from "@fluentui/react-icons";
import { THEMES, fontStack } from "@officehere/core";
import { ColorPalette, Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonStack, useMenuClose } from "@officehere/ui";
import { useBelge } from "../state/context";
import { PARAGRAPH_SPACINGS, STYLE_SETS, resolveStyles, previewStyle } from "../model/styles";
import type { Watermark } from "../model/document";

export const WATERMARK_PRESETS: { group: string; text: string; layout: "diagonal" | "horizontal" }[] = [
  { group: "Gizli", text: "GİZLİ", layout: "diagonal" },
  { group: "Gizli", text: "GİZLİ", layout: "horizontal" },
  { group: "Gizli", text: "KOPYALAMAYIN", layout: "diagonal" },
  { group: "Feragatnameler", text: "TASLAK", layout: "diagonal" },
  { group: "Feragatnameler", text: "ÖRNEK", layout: "diagonal" },
  { group: "Acil", text: "ACİL", layout: "diagonal" },
  { group: "Acil", text: "ÇOK ACİL", layout: "horizontal" },
];

function StyleSetPreview({ id, theme, fontTheme }: { id: string; theme: string; fontTheme: string | null }) {
  const st = resolveStyles(id, theme, fontTheme);
  return (
    <span className="bl-setprev">
      <span style={{ ...previewStyle(st, "Title"), fontSize: 10 }}>Konu Başlığı</span>
      <span style={{ ...previewStyle(st, "Heading1"), fontSize: 7.5 }}>Başlık 1</span>
      <span className="bl-setprev-lines">
        <i />
        <i />
        <i />
      </span>
    </span>
  );
}

export function DesignTab() {
  const ctx = useBelge();
  const { settings, updateSettings } = ctx;
  const setWm = (wm: Watermark | null) => updateSettings((s) => ({ ...s, watermark: wm }));

  return (
    <>
      <RibbonGroup label="Belge Biçimlendirme" icon={<DesignIdeas32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DesignIdeas32Regular />}
          label="Temalar"
          menu={
            <Menu width={220}>
              <MenuHeader>OfficeHere</MenuHeader>
              {THEMES.map((t) => (
                <MenuItem
                  key={t.name}
                  checked={settings.theme === t.name}
                  label={
                    <span className="bl-theme-item">
                      <span className="bl-theme-swatches">
                        {t.colors.slice(4).map((c) => (
                          <i key={c} style={{ background: c }} />
                        ))}
                      </span>
                      <span style={{ fontFamily: fontStack(t.majorFont) }}>{t.name}</span>
                    </span>
                  }
                  onClick={() => updateSettings((s) => ({ ...s, theme: t.name, fontTheme: null }))}
                />
              ))}
            </Menu>
          }
        />
        <RibbonGallery
          columns={5}
          itemWidth={74}
          itemHeight={58}
          selectedId={settings.styleSet}
          onSelect={(id) => updateSettings((s) => ({ ...s, styleSet: id }))}
          items={STYLE_SETS.map((s) => ({ id: s.id, label: s.label, preview: <StyleSetPreview id={s.id} theme={settings.theme} fontTheme={settings.fontTheme} /> }))}
        />
        <RibbonStack>
          <RibbonMenuButton
            icon={<Color32Regular />}
            label="Renkler"
            menu={
              <Menu width={220}>
                {THEMES.map((t) => (
                  <MenuItem
                    key={t.name}
                    checked={settings.theme === t.name}
                    label={
                      <span className="bl-theme-item">
                        <span className="bl-theme-swatches">
                          {t.colors.slice(2).map((c, i) => (
                            <i key={`${c}${i}`} style={{ background: c }} />
                          ))}
                        </span>
                        {t.name}
                      </span>
                    }
                    onClick={() => updateSettings((s) => ({ ...s, fontTheme: s.fontTheme ?? s.theme, theme: t.name }))}
                  />
                ))}
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<TextFont20Regular />}
            label="Yazı Tipleri"
            menu={
              <Menu width={240}>
                {THEMES.map((t) => (
                  <MenuItem
                    key={t.name}
                    checked={(settings.fontTheme ?? settings.theme) === t.name}
                    label={
                      <span className="bl-font-item">
                        <b>{t.name}</b>
                        <span style={{ fontFamily: fontStack(t.majorFont), fontSize: 14 }}>{t.majorFont}</span>
                        <span style={{ fontFamily: fontStack(t.minorFont) }}>{t.minorFont}</span>
                      </span>
                    }
                    onClick={() => updateSettings((s) => ({ ...s, fontTheme: t.name }))}
                  />
                ))}
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<TextParagraph20Regular />}
            label="Paragraf Aralığı"
            menu={
              <Menu width={260}>
                <MenuHeader>Yerleşik</MenuHeader>
                {PARAGRAPH_SPACINGS.map((p) => (
                  <MenuItem key={p.id} label={p.label} description={p.desc} checked={settings.paragraphSpacing === p.id} onClick={() => updateSettings((s) => ({ ...s, paragraphSpacing: p.id }))} />
                ))}
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Sayfa Arka Planı" icon={<PaintBucket24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Water32Regular />}
          label="Filigran"
          menu={
            <div className="oh-menu bl-wm-menu">
              {["Gizli", "Feragatnameler", "Acil"].map((g) => (
                <div key={g}>
                  <MenuHeader>{g}</MenuHeader>
                  <div className="bl-wm-grid">
                    {WATERMARK_PRESETS.filter((w) => w.group === g).map((w, i) => (
                      <WmTile key={i} text={w.text} layout={w.layout} onPick={() => setWm({ text: w.text, layout: w.layout, color: "#C0C0C0", font: "Aptos", semitransparent: true })} />
                    ))}
                  </div>
                </div>
              ))}
              <MenuSeparator />
              <MenuItem label="Özel Filigran..." onClick={() => ctx.openDialog("watermark")} />
              <MenuItem label="Filigranı Kaldır" disabled={!settings.watermark} onClick={() => setWm(null)} />
            </div>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<PaintBucket24Regular />}
          label="Sayfa Rengi"
          menu={
            <div className="oh-menu oh-menu-palette">
              <ColorPalette noneLabel="Renk Yok" value={settings.pageColor} onSelect={(c) => updateSettings((s) => ({ ...s, pageColor: c }))} />
            </div>
          }
        />
        <RibbonButton size="large" icon={<DocumentBorder32Regular />} label="Sayfa Kenarlıkları" onClick={() => ctx.openDialog("pageBorders")} />
      </RibbonGroup>
    </>
  );
}

function WmTile({ text, layout, onPick }: { text: string; layout: "diagonal" | "horizontal"; onPick: () => void }) {
  return (
    <MenuItemTile onPick={onPick}>
      <span className={`bl-wm-prev is-${layout}`}>
        <span>{text}</span>
      </span>
    </MenuItemTile>
  );
}

function MenuItemTile({ onPick, children }: { onPick: () => void; children: React.ReactNode }) {
  const close = useMenuClose();
  return (
    <button
      type="button"
      className="bl-wm-tile"
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        onPick();
        close();
      }}
    >
      {children}
    </button>
  );
}
