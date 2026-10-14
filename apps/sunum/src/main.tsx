import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import { App } from "./App";
import { ErrorBoundary } from "./app/ErrorBoundary";
import { recoveryHandlers } from "./app/appState";
import { AudienceView, preloadChunks } from "./app/lazyChunks";

// "?izleyici=<kimlik>": Sunucu Görünümü'nün izleyici penceresi.
const audience = new URLSearchParams(window.location.search).get("izleyici");

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    {audience ? (
      <AudienceView id={audience} />
    ) : (
      <ErrorBoundary region="OfficeHere Sunum" root recovery={recoveryHandlers()}>
        <App />
      </ErrorBoundary>
    )}
  </StrictMode>,
);

// Ayrı parçalar (dosya okuyucu/yazıcıları vb.) ilk çizimden sonra boşta önceden yüklenir:
// ilk açılış hızlı kalır, kaydetme ve gösteri sonradan ağa gerek duymaz.
preloadChunks();
