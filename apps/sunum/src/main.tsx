import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import { App } from "./App";
import { ErrorBoundary } from "./app/ErrorBoundary";
import { recoveryHandlers } from "./app/appState";
import { AudienceView } from "./show/AudienceView";

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
