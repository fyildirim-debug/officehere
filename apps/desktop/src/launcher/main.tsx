import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import "./launcher.css";
import { installErrorReporting } from "@officehere/ui";
import { Launcher } from "./Launcher";
import { LauncherErrorBoundary } from "./LauncherErrorBoundary";

// Yakalanmamış hatalar yerel günlüğe yazılır (belge içeriği olmadan).
installErrorReporting();

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <LauncherErrorBoundary>
      <Launcher />
    </LauncherErrorBoundary>
  </StrictMode>,
);
