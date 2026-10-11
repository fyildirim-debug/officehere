// Geçici iskelet: sunum uygulaması bunun yerine gelecek.
import { AppShell, TitleBar } from "@officehere/ui";

export function App() {
  return (
    <AppShell app="sunum" titleBar={<TitleBar app="sunum" title="Sunu1" />}>
      <div />
    </AppShell>
  );
}
