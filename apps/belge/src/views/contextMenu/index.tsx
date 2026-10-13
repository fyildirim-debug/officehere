// Sağ tık menüsü: sağ tıklanan nesneye (resim, grafik, tablo, metin kutusu, denklem, İçindekiler) göre kendi
// menüsü; metinde Yazı Tipi/Paragraf menüsü. Hedef target.ts'te bulunur ve seçilir.
import type { MenuDeps } from "./common";
import type { MenuKind, MenuTarget } from "./target";
import { TextMenu } from "./text";
import { TableMenu } from "./table";
import { ImageMenu } from "./image";
import { ChartMenu } from "./chart";
import { TextBoxMenu } from "./textbox";
import { EquationMenu } from "./equation";
import { TocMenu } from "./toc";

export { menuPointForSelection, selectContextTarget, targetOfSelection, type MenuKind, type MenuTarget } from "./target";
export type { MenuDeps } from "./common";

/** Mini araç çubuğu yalnız yazı içeren hedeflerde menünün üstünde açılır. */
export const showsMiniBar = (kind: MenuKind) => kind === "text" || kind === "table";

export function ContextMenuContent({ target, deps }: { target: MenuTarget; deps: MenuDeps }) {
  switch (target.kind) {
    case "image":
      return <ImageMenu deps={deps} />;
    case "chart":
      return <ChartMenu deps={deps} />;
    case "table":
      return <TableMenu deps={deps} />;
    case "textBox":
      return <TextBoxMenu deps={deps} target={target} />;
    case "equation":
      return <EquationMenu deps={deps} />;
    case "toc":
    case "bibliography":
      return <TocMenu deps={deps} target={target} />;
    default:
      return <TextMenu deps={deps} target={target} />;
  }
}
