/** Tasarım sekmesi: Temalar, Çeşitlemeler, Özelleştir. */
import { memo, useMemo } from "react";
import { PaintBucket20Regular, SlideSize24Regular,
  Color24Regular,
  PaintBrush24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGallery, RibbonGroup, RibbonMenuButton, useContextMenu, useMenuClose } from "@officehere/ui";
import { makeSp, textBody } from "../model/factory";
import { applyThemeToSlides } from "../model/designs";
import { applyTheme, applyVariant, defaultThemeId, newPresentation, setDefaultThemeId, SLIDE_16_9, SLIDE_4_3 } from "../model/presentation";
import type { Color, Fill, Presentation, Theme } from "../model/types";
import { fillToCss } from "../render/style";
import { ScaledSlide } from "../render/SlideView";
import { THEMES, getThemeDef, type ColorScheme } from "../themes/themes";
import type { RibbonCtx } from "./types";

const previewCache = new Map<string, Presentation>();

function previewPres(themeId: string, variant: number, ratio: number): Presentation {
  const key = `${themeId}:${variant}:${ratio.toFixed(3)}`;
  const hit = previewCache.get(key);
  if (hit) return hit;
  const size = ratio < 1.5 ? SLIDE_4_3 : SLIDE_16_9;
  const p = newPresentation(themeId, variant, size);
  p.slides[0].shapes = p.slides[0].shapes
    .filter((s) => !(s.kind === "sp" && s.ph?.type === "subTitle"))
    .map((s) => (s.kind === "sp" && s.ph?.type === "ctrTitle" ? { ...s, text: { ...textBody("Aa"), paragraphs: [{ runs: [{ text: "Aa", props: { size: 220 } }], level: 0, props: {} }] } } : s));
  // Vurgu renklerinden şerit (Office galerisindeki gibi).
  const bw = (p.width * 0.6) / 6;
  (["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).forEach((k, i) => {
    p.slides[0].shapes.push(makeSp("rect", p.width * 0.2 + i * bw, p.height * 0.8, bw, p.height * 0.08, { fill: { type: "solid", color: { scheme: k } }, line: { fill: { type: "none" }, width: 1, dash: "solid" } }));
  });
  previewCache.set(key, p);
  return p;
}

export const ThemeTile = memo(function ThemeTile({ themeId, variant, width, ratio }: { themeId: string; variant: number; width: number; ratio: number }) {
  const p = useMemo(() => previewPres(themeId, variant, ratio), [themeId, variant, ratio]);
  return <ScaledSlide pres={p} slide={p.slides[0]} slideNumber={1} width={width} />;
});

export function PaletteRow({ c }: { c: ColorScheme }) {
  return (
    <span style={{ display: "inline-flex", gap: 1 }}>
      {(["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).map((k) => (
        <span key={k} style={{ width: 10, height: 12, background: c[k], border: "1px solid rgba(0,0,0,.15)" }} />
      ))}
    </span>
  );
}

export const FONT_PAIRS: { name: string; major: string; minor: string }[] = [
  { name: "OfficeHere", major: "Aptos Display", minor: "Aptos" },
  { name: "Klasik", major: "Calibri Light", minor: "Calibri" },
  { name: "Arial", major: "Arial", minor: "Arial" },
  { name: "Cambria", major: "Cambria", minor: "Calibri" },
  { name: "Georgia", major: "Georgia", minor: "Verdana" },
  { name: "Trebuchet", major: "Trebuchet MS", minor: "Trebuchet MS" },
  { name: "Kâğıt", major: "Book Antiqua", minor: "Garamond" },
  { name: "Times", major: "Times New Roman", minor: "Times New Roman" },
  { name: "Segoe", major: "Segoe UI", minor: "Segoe UI" },
  { name: "Tahoma", major: "Verdana", minor: "Tahoma" },
];

function bgStyles(): Fill[] {
  const g = (a: Color, b: Color, kind: "linear" | "radial" = "linear"): Fill => ({ type: "gradient", kind, angle: 90, stops: [{ pos: 0, color: a }, { pos: 1, color: b }] });
  return [
    { type: "solid", color: { scheme: "lt1" } },
    { type: "solid", color: { scheme: "lt2" } },
    { type: "solid", color: { scheme: "dk2" } },
    { type: "solid", color: { scheme: "dk1" } },
    g({ scheme: "lt1" }, { scheme: "lt2" }),
    g({ scheme: "lt2" }, { scheme: "lt1" }, "radial"),
    g({ scheme: "dk2", lumMod: 0.75, lumOff: 0.25 }, { scheme: "dk2" }),
    g({ scheme: "dk2" }, { scheme: "dk1" }, "radial"),
    g({ scheme: "accent1", lumMod: 0.2, lumOff: 0.8 }, { scheme: "lt1" }),
    g({ scheme: "accent2", lumMod: 0.2, lumOff: 0.8 }, { scheme: "lt1" }),
    g({ scheme: "accent1" }, { scheme: "accent1", lumMod: 0.5 }),
    g({ scheme: "accent2", lumMod: 0.75 }, { scheme: "dk1" }, "radial"),
  ];
}

export function BgStyleGrid({ ctx }: { ctx: RibbonCtx }) {
  const close = useMenuClose();
  const theme = ctx.state.pres.theme;
  return (
    <div className="preset-grid" style={{ gridTemplateColumns: "repeat(4, 60px)" }}>
      {bgStyles().map((f, i) => (
        <button
          key={i}
          type="button"
          className="preset-btn"
          title={`Stil ${i + 1}`}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            if (ctx.state.view === "master")
              // Asıl görünümde: asıl slayta uygulanır, düzenlerin kendi arka planları kaldırılır.
              ctx.actions.store.commit((p) => ({ ...p, slides: p.slides.map((s, i) => ({ ...s, background: i === 0 ? f : undefined })) }));
            else ctx.actions.store.commit((p) => ({ ...p, master: { ...p.master, background: f }, slides: p.slides.map((s) => ({ ...s, background: undefined })) }));
            close();
          }}
        >
          <span style={{ display: "block", width: 54, height: 32, background: fillToCss(f, theme), border: "1px solid #ccc" }} />
        </button>
      ))}
    </div>
  );
}

export function DesignTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const pres = state.pres;
  const ratio = pres.width / pres.height;
  const curDef = getThemeDef(pres.theme.id);
  const setTheme = (t: Partial<Theme>) => a.store.commit((p) => ({ ...p, theme: { ...p.theme, ...t } }));
  const themeMenu = useContextMenu();
  // Tema galerisinde sağ tık: Tüm Slaytlara Uygula / Varsayılan Tema Olarak Ayarla.
  const onThemeMenu = (e: React.MouseEvent) => {
    const item = (e.target as Element).closest(".oh-gallery-item");
    const t = item ? THEMES.find((x) => x.name === item.getAttribute("aria-label")) : undefined;
    if (!t) return;
    e.preventDefault();
    const isDefault = defaultThemeId() === t.id;
    themeMenu.open(
      e,
      <Menu>
        <MenuItem label="Tüm Slaytlara Uygula" onClick={() => a.store.commit((p) => applyTheme(p, t.id, 0), { sel: [], editing: null })} />
        <MenuItem
          label="Seçili Slaytlara Uygula"
          onClick={() => {
            const cur = a.s.pres.slides[a.s.slideIdx]?.id;
            const ids = a.s.slideSel.length ? a.s.slideSel : cur ? [cur] : [];
            a.store.commit((p) => applyThemeToSlides(p, t.id, ids), { sel: [], editing: null });
          }}
        />
        <MenuItem label="Varsayılan Tema Olarak Ayarla" checked={isDefault} onClick={() => {
          setDefaultThemeId(t.id);
          a.message(`"${t.name}" yeni boş sunular için varsayılan tema olarak ayarlandı.`, "success");
        }} />
      </Menu>,
    );
  };
  return (
    <>
      <RibbonGroup label="Temalar" icon={<PaintBrush24Regular />}>
        <div className="theme-ctx-host" onContextMenu={onThemeMenu}>
        <RibbonGallery
          columns={6}
          itemWidth={78}
          itemHeight={ratio < 1.5 ? 60 : 46}
          expandedColumns={4}
          selectedId={pres.theme.id}
          items={THEMES.map((t) => ({ id: t.id, label: t.name, preview: <ThemeTile themeId={t.id} variant={0} width={74} ratio={ratio} /> }))}
          onSelect={(id) => a.store.commit((p) => applyTheme(p, id, 0), { sel: [], editing: null })}
          footer={
            <>
              <MenuSeparator />
              <MenuItem label="Temaları Sıfırla (Varsayılan Tema)" onClick={() => a.store.commit((p) => applyTheme(p, defaultThemeId(), 0))} />
            </>
          }
        />
        </div>
        {themeMenu.element}
      </RibbonGroup>
      <RibbonGroup label="Çeşitlemeler" icon={<Color24Regular />}>
        <RibbonGallery
          columns={4}
          itemWidth={78}
          itemHeight={ratio < 1.5 ? 60 : 46}
          selectedId={curDef ? `v${pres.theme.variant ?? 0}` : undefined}
          items={(curDef ?? THEMES[0]).variants.map((v, i) => ({ id: `v${i}`, label: v.name, preview: <ThemeTile themeId={(curDef ?? THEMES[0]).id} variant={i} width={74} ratio={ratio} /> }))}
          onSelect={(id) => {
            const v = Number(id.slice(1));
            if (curDef) a.store.commit((p) => applyVariant(p, v));
            else a.store.commit((p) => applyTheme(p, THEMES[0].id, v));
          }}
          footer={
            <>
              <MenuSeparator />
              <MenuItem
                label="Renkler"
                submenu={
                  <Menu className="menu-wide">
                    {THEMES.flatMap((t) =>
                      t.variants.map((v) => (
                        <MenuItem key={`${t.id}-${v.name}`} icon={<span />} label={<span style={{ display: "flex", gap: 8, alignItems: "center" }}><PaletteRow c={v.colors} /> {v.name}</span>} onClick={() => setTheme({ colors: { ...v.colors } })} />
                      )),
                    )}
                  </Menu>
                }
              />
              <MenuItem
                label="Yazı Tipleri"
                submenu={
                  <Menu className="menu-wide">
                    {FONT_PAIRS.map((f) => (
                      <MenuItem
                        key={f.name}
                        checked={pres.theme.fonts.major === f.major && pres.theme.fonts.minor === f.minor}
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
              <MenuItem
                label="Arka Plan Stilleri"
                submenu={
                  <Menu>
                    <BgStyleGrid ctx={ctx} />
                    <MenuSeparator />
                    <MenuItem label="Arka Planı Biçimlendir..." onClick={() => cmd.togglePane("background")} />
                  </Menu>
                }
              />
            </>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Özelleştir" icon={<SlideSize24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<SlideSize24Regular />}
          label="Slayt Boyutu"
          description="Sunudaki tüm slaytların boyutunu değiştirin."
          menu={
            <Menu>
              <MenuHeader>Slayt Boyutu</MenuHeader>
              <MenuItem label="Standart (4:3)" checked={Math.abs(ratio - 4 / 3) < 0.01} onClick={() => Math.abs(ratio - 4 / 3) >= 0.01 && cmd.setSlideSize(SLIDE_4_3.width, SLIDE_4_3.height)} />
              <MenuItem label="Geniş Ekran (16:9)" checked={Math.abs(ratio - 16 / 9) < 0.01} onClick={() => Math.abs(ratio - 16 / 9) >= 0.01 && cmd.setSlideSize(SLIDE_16_9.width, SLIDE_16_9.height)} />
              <MenuSeparator />
              <MenuItem label="Özel Slayt Boyutu..." onClick={() => ctx.openDialog("slideSize")} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<PaintBucket20Regular />} label="Arka Planı Biçimlendir" checked={state.pane === "background"} description="Slayt arka planını dolgu, gradyan veya resimle biçimlendirin." onClick={() => cmd.togglePane("background")} />
      </RibbonGroup>
    </>
  );
}
