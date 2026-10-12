/** Renk seçicilerin tema paleti: çalışma kitabı temasından (Sayfa Düzeni > Temalar). */
import { createContext, useContext } from "react";
import { DEFAULT_THEME, type ThemePalette } from "@officehere/core";

export const ThemePaletteContext = createContext<ThemePalette>(DEFAULT_THEME);

export function useThemePalette(): ThemePalette {
  return useContext(ThemePaletteContext);
}
