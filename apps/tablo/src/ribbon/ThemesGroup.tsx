/** Sayfa Düzeni > Temalar grubu: Temalar galerisi, Renkler ve Yazı Tipleri menüleri. */
import { TextFont20Regular, Color20Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { fontStack } from "@officehere/core";
import { Menu, MenuHeader, MenuItem, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import { applyTheme } from "../actions/theme";
import { COLOR_SETS, FONT_SETS, THEME_LIST, type WorkbookTheme } from "../model/theme";

function ThemeTile({ t, active, onPick }: { t: WorkbookTheme; active: boolean; onPick: () => void }) {
  return (
    <button type="button" className={`tb-theme-tile${active ? " is-active" : ""}`} onMouseDown={(e) => e.preventDefault()} onClick={onPick} title={t.name} aria-label={t.name}>
      <span className="tb-theme-aa" style={{ fontFamily: fontStack(t.majorFont), color: t.colors[3] }}>
        Aa
      </span>
      <span className="tb-theme-bars">
        {t.colors.slice(4, 8).map((c, i) => (
          <span key={i} style={{ background: c }} />
        ))}
      </span>
      <span className="tb-theme-name">{t.name}</span>
    </button>
  );
}

function Swatches({ colors }: { colors: readonly string[] }) {
  return (
    <span className="tb-theme-swatches">
      {[colors[1], colors[3], ...colors.slice(4)].map((c, i) => (
        <span key={i} style={{ background: c }} />
      ))}
    </span>
  );
}

export function ThemesGroup({ ctl }: { ctl: Controller }) {
  const cur = ctl.wb.theme;
  return (
    <RibbonGroup label="Temalar" icon={<DesignIdeas24Regular />}>
      <RibbonMenuButton
        size="large"
        icon={<DesignIdeas24Regular />}
        label="Temalar"
        description="Belgenin genel tasarımını (renkler ve yazı tipleri) değiştirir."
        menu={(close) => (
          <div className="oh-menu tb-theme-menu" role="menu">
            <MenuHeader>Yerleşik</MenuHeader>
            <div className="tb-theme-grid">
              {THEME_LIST.map((t) => (
                <ThemeTile
                  key={t.name}
                  t={t}
                  active={cur.name === t.name && cur.colorsName === t.colorsName && cur.fontsName === t.fontsName}
                  onPick={() => {
                    close();
                    applyTheme(ctl, structuredClone(t), "Tema");
                  }}
                />
              ))}
            </div>
          </div>
        )}
      />
      <RibbonStack>
        <RibbonMenuButton
          icon={<Color20Regular />}
          label="Renkler"
          menu={
            <Menu width={290} className="tb-menu-scroll">
              <MenuHeader>Renkler</MenuHeader>
              {COLOR_SETS.map((c) => (
                <MenuItem key={c.name} checked={cur.colorsName === c.name} label={<span className="tb-theme-colorrow"><Swatches colors={c.colors} /> {c.name}</span>} onClick={() => applyTheme(ctl, { colors: [...c.colors] as WorkbookTheme["colors"], colorsName: c.name }, "Tema Renkleri")} />
              ))}
            </Menu>
          }
        />
        <RibbonMenuButton
          icon={<TextFont20Regular />}
          label="Yazı Tipleri"
          menu={
            <Menu width={260} className="tb-menu-scroll">
              <MenuHeader>Yazı Tipleri</MenuHeader>
              {FONT_SETS.map((f) => (
                <MenuItem
                  key={f.name}
                  checked={cur.fontsName === f.name}
                  label={
                    <span className="tb-theme-fontrow">
                      <span className="tb-theme-aa-small" style={{ fontFamily: fontStack(f.major) }}>
                        Aa
                      </span>
                      <span>
                        <strong>{f.name}</strong>
                        <span style={{ fontFamily: fontStack(f.major), display: "block" }}>{f.major}</span>
                        <span style={{ fontFamily: fontStack(f.minor), display: "block" }}>{f.minor}</span>
                      </span>
                    </span>
                  }
                  onClick={() => applyTheme(ctl, { majorFont: f.major, minorFont: f.minor, fontsName: f.name }, "Tema Yazı Tipleri")}
                />
              ))}
            </Menu>
          }
        />
      </RibbonStack>
    </RibbonGroup>
  );
}

