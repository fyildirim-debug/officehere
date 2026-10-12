/** Asıl Slayt sekmesi (Asıl Slayt Görünümünde): düzen ekleme/silme, yer tutucular, tema, arka plan. */
import {
  Color24Regular,
  ContentView32Regular,
  Delete20Regular,
  DismissSquare32Regular,
  PaintBucket20Regular,
  Rename20Regular,
  SlideAdd32Regular,
  SlideLayout24Regular,
  TextFont24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { MASTER_SLIDE_ID } from "../model/masterView";
import { masterAddPlaceholder, masterHas, masterInsertLayout, masterToggle } from "../state/masterActions";
import { THEMES } from "../themes/themes";
import { BgStyleGrid, FONT_PAIRS, PaletteRow } from "./DesignTab";
import type { RibbonCtx } from "./types";

export function MasterTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const cur = state.pres.slides[state.slideIdx];
  const isMaster = cur?.id === MASTER_SLIDE_ID;
  const setTheme = (t: Partial<typeof state.pres.theme>) => a.store.commit((p) => ({ ...p, theme: { ...p.theme, ...t } }));
  return (
    <>
      <RibbonGroup label="Asıl Slaydı Düzenle" icon={<SlideLayout24Regular />}>
        <RibbonButton size="large" icon={<SlideAdd32Regular />} label="Düzen Ekle" description="Asıl slayda yeni, özel bir düzen ekleyin." onClick={() => masterInsertLayout(a)} />
        <RibbonStack>
          <RibbonButton icon={<Delete20Regular />} label="Sil" description="Seçili düzeni silin (kullanılan düzenler silinemez)." disabled={isMaster || !cur} onClick={() => a.deleteSlides()} />
          <RibbonButton icon={<Rename20Regular />} label="Yeniden Adlandır" description="Seçili düzeni yeniden adlandırın." disabled={isMaster || !cur} onClick={cmd.renameLayout} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Asıl Düzen" icon={<ContentView32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<ContentView32Regular />}
          label="Yer Tutucu Ekle"
          description="Düzene içerik, metin, resim ya da tablo yer tutucusu ekleyin."
          disabled={isMaster}
          menu={
            <Menu>
              <MenuItem label="İçerik" onClick={() => masterAddPlaceholder(a, "obj")} />
              <MenuItem label="Metin" onClick={() => masterAddPlaceholder(a, "body")} />
              <MenuItem label="Resim" onClick={() => masterAddPlaceholder(a, "pic")} />
              <MenuItem label="Tablo" onClick={() => masterAddPlaceholder(a, "tbl")} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonCheckbox label="Başlık" disabled={isMaster} checked={masterHas(a, "title")} onChange={(v) => masterToggle(a, "title", v)} />
          <RibbonCheckbox label="Alt Bilgiler" disabled={isMaster} checked={masterHas(a, "footers")} onChange={(v) => masterToggle(a, "footers", v)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tema Düzenle" icon={<Color24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Color24Regular />}
          label="Renkler"
          description="Tema renklerini değiştirin."
          menu={
            <Menu className="menu-wide">
              {THEMES.flatMap((t) =>
                t.variants.map((v) => (
                  <MenuItem
                    key={`${t.id}-${v.name}`}
                    icon={<span />}
                    label={
                      <span style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <PaletteRow c={v.colors} /> {v.name}
                      </span>
                    }
                    onClick={() => setTheme({ colors: { ...v.colors } })}
                  />
                )),
              )}
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<TextFont24Regular />}
          label="Yazı Tipleri"
          description="Tema yazı tiplerini değiştirin."
          menu={
            <Menu className="menu-wide">
              {FONT_PAIRS.map((f) => (
                <MenuItem
                  key={f.name}
                  checked={state.pres.theme.fonts.major === f.major && state.pres.theme.fonts.minor === f.minor}
                  label={
                    <span>
                      <span style={{ fontFamily: f.major, fontSize: 15 }}>Aa</span> {f.name}
                      <span className="oh-menu-desc">
                        {f.major} / {f.minor}
                      </span>
                    </span>
                  }
                  onClick={() => setTheme({ fonts: { major: f.major, minor: f.minor } })}
                />
              ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Arka Plan" icon={<PaintBucket20Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<PaintBucket20Regular />}
          label="Arka Plan Stilleri"
          menu={
            <Menu>
              <BgStyleGrid ctx={ctx} />
              <MenuItem label="Arka Planı Biçimlendir..." onClick={() => cmd.togglePane("background")} />
            </Menu>
          }
        />
        <RibbonCheckbox label="Arka Plan Grafiklerini Gizle" disabled={isMaster || !cur} checked={!!cur?.hideMasterShapes} onChange={(v) => a.store.commit((p) => ({ ...p, slides: p.slides.map((s) => (s.id === cur?.id ? { ...s, hideMasterShapes: v || undefined } : s)) }))} />
      </RibbonGroup>
      <RibbonGroup label="Kapat" icon={<DismissSquare32Regular />}>
        <RibbonButton size="large" className="rb-wide" icon={<DismissSquare32Regular />} label="Asıl Görünümü Kapat" description="Asıl Slayt Görünümünden çıkıp slaytlarınızı düzenlemeye dönün." onClick={() => cmd.setView("normal")} />
      </RibbonGroup>
    </>
  );
}
