import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import "./styles/tablo.css";
import "./styles/app.css";
import "./styles/dialogs.css";
import "./styles/views.css";
import "./styles/pagelayout.css";
import "./styles/panes.css";
import "./styles/objects.css";
import { App } from "./App";
import { Controller } from "./controller";
import { ErrorBoundary } from "./app/ErrorBoundary";
import { recoveryHandlers } from "./app/recovery";

// Denetleyici kökün dışında tutulur: uygulama çökse de çalışma kitabı bellekte kalır
const ctl = new Controller();

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <ErrorBoundary region="OfficeHere Tablo" root recovery={recoveryHandlers(ctl)}>
      <App ctl={ctl} />
    </ErrorBoundary>
  </StrictMode>,
);
