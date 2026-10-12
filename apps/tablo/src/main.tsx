import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@officehere/ui/styles.css";
import "./styles/tablo.css";
import { App } from "./App";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
