import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import { App } from "./App";
import { ErrorBoundary } from "./views/ErrorBoundary";

createRoot(document.getElementById("root")!, {
  // Hata sınırlarının yakaladığı hatalar uyarı olarak günlüğe yazılır (kullanıcıya kurtarma ekranı gösterilir).
  onCaughtError: (error) => console.warn("[OfficeHere Belge] yakalanan hata:", error),
}).render(
  <StrictMode>
    <ErrorBoundary name="app" label="Uygulama" root>
      <App />
    </ErrorBoundary>
  </StrictMode>,
);
