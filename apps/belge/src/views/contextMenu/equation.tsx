// Denklem bağlam menüsü: Kes, Kopyala, Denklemi Düzenle ve sil.
import { MathFormula20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { ClipboardItems, DeleteItem, type MenuDeps } from "./common";

export function EquationMenu({ deps }: { deps: MenuDeps }) {
  const { openDialog } = useBelge();
  const { editor } = deps;
  return (
    <Menu width={220} label="Denklem">
      <ClipboardItems editor={editor} paste={false} />
      <MenuSeparator />
      <MenuItem icon={<MathFormula20Regular />} label="Denklemi Düzenle..." onClick={() => openDialog("equation")} />
      <MenuSeparator />
      <DeleteItem editor={editor} />
    </Menu>
  );
}
