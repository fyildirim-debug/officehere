// Geçici iskelet: tablo uygulaması bunun yerine gelecek.
import { AppShell, TitleBar } from "@officehere/ui";

export function App() {
  return (
    <AppShell app="tablo" titleBar={<TitleBar app="tablo" title="Kitap1" />}>
      <div />
    </AppShell>
  );
}
