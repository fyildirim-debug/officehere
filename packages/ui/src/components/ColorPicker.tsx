import { useRef, type ReactElement } from "react";
import { DEFAULT_THEME, STANDARD_COLORS, HIGHLIGHT_COLORS, themeShades, type ThemePalette } from "@officehere/core";
import { RibbonSplitButton, type RibbonButtonSize } from "./Ribbon";
import { MenuHeader, MenuItem, MenuSeparator, useMenuClose } from "./Menu";

export interface ColorPaletteProps {
  onSelect: (color: string | null) => void;
  /** "Otomatik" / "Dolgu Yok" satırının metni; verilmezse gösterilmez. */
  noneLabel?: string;
  noneIcon?: ReactElement;
  theme?: ThemePalette;
  /** Vurgulama paleti (yalnız 15 vurgu rengi). */
  variant?: "full" | "highlight";
  value?: string | null;
}

/** Office tarzı renk paleti: tema renkleri (10×6), standart renkler, diğer renkler. */
export function ColorPalette({ onSelect, noneLabel, noneIcon, theme = DEFAULT_THEME, variant = "full", value }: ColorPaletteProps) {
  const close = useMenuClose();
  const inputRef = useRef<HTMLInputElement>(null);
  const pick = (c: string | null) => {
    onSelect(c);
    close();
  };
  const norm = value?.toUpperCase();

  if (variant === "highlight") {
    return (
      <div className="oh-palette">
        <div className="oh-palette-grid is-highlight">
          {HIGHLIGHT_COLORS.map((c) => (
            <Swatch key={c} color={c} selected={norm === c} onPick={pick} />
          ))}
        </div>
        <MenuSeparator />
        <MenuItem label={noneLabel ?? "Renk Yok"} onClick={() => onSelect(null)} />
      </div>
    );
  }

  const columns = theme.colors.map((c) => themeShades(c));
  return (
    <div className="oh-palette">
      {noneLabel && (
        <>
          <MenuItem icon={noneIcon ?? <span className="oh-swatch-none" />} label={noneLabel} onClick={() => onSelect(null)} />
          <MenuSeparator />
        </>
      )}
      <MenuHeader>Tema Renkleri</MenuHeader>
      <div className="oh-palette-theme">
        <div className="oh-palette-grid">
          {columns.map((col) => (
            <Swatch key={`base-${col[0]}`} color={col[0]} selected={norm === col[0].toUpperCase()} onPick={pick} />
          ))}
        </div>
        <div className="oh-palette-grid is-shades">
          {[1, 2, 3, 4, 5].map((row) => columns.map((col, i) => <Swatch key={`${row}-${i}`} color={col[row]} selected={norm === col[row].toUpperCase()} onPick={pick} />))}
        </div>
      </div>
      <MenuHeader>Standart Renkler</MenuHeader>
      <div className="oh-palette-grid">
        {STANDARD_COLORS.map((c) => (
          <Swatch key={c} color={c} selected={norm === c} onPick={pick} />
        ))}
      </div>
      <MenuSeparator />
      <MenuItem
        label="Diğer Renkler..."
        keepOpen
        icon={
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="5" cy="6" r="3" fill="#e81123" />
            <circle cx="11" cy="6" r="3" fill="#0078d4" />
            <circle cx="8" cy="11" r="3" fill="#16c60c" />
          </svg>
        }
        onClick={() => inputRef.current?.click()}
      />
      <input
        ref={inputRef}
        type="color"
        className="oh-hidden-input"
        defaultValue={value ?? "#000000"}
        onChange={(e) => pick(e.target.value.toUpperCase())}
      />
    </div>
  );
}

function Swatch({ color, selected, onPick }: { color: string; selected: boolean; onPick: (c: string) => void }) {
  return (
    <button
      type="button"
      className={`oh-swatch${selected ? " is-selected" : ""}`}
      style={{ background: color }}
      title={color}
      aria-label={color}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => onPick(color)}
    />
  );
}

export interface ColorSplitButtonProps {
  icon: ReactElement;
  label?: string;
  tip: string;
  description?: string;
  /** Son seçilen renk (ana düğme bunu uygular). */
  color: string | null;
  onApply: (color: string | null) => void;
  /** Paletten seçim yapılınca (son rengi güncellemek için). */
  onColorChange?: (color: string | null) => void;
  noneLabel?: string;
  variant?: "full" | "highlight";
  size?: RibbonButtonSize;
  theme?: ThemePalette;
  disabled?: boolean;
}

/** Yazı tipi rengi / vurgulama / dolgu gibi renk şeritli bölünmüş düğme. */
export function ColorSplitButton({ icon, label, tip, description, color, onApply, onColorChange, noneLabel, variant, size = "small", theme, disabled }: ColorSplitButtonProps) {
  return (
    <RibbonSplitButton
      icon={icon}
      label={label}
      size={size}
      tip={tip}
      description={description}
      colorBar={color}
      disabled={disabled}
      onClick={() => onApply(color)}
      menu={
        <div className="oh-menu oh-menu-palette">
          <ColorPalette
            variant={variant}
            noneLabel={noneLabel}
            theme={theme}
            value={color}
            onSelect={(c) => {
              onColorChange?.(c);
              onApply(c);
            }}
          />
        </div>
      }
    />
  );
}
