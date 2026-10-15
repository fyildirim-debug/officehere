import { ipcRenderer } from "electron";

// PDF görüntüleyicisinin arama çubuğu (Ctrl+F). Sayfa yalnız HTML ve CSS'tir (betik yok); olaylar burada
// bağlanır. Arama ana süreçte PDF sayfasında yapılır (webContents.findInPage), sonuç "2/5" olarak gelir.
// Sayfaya hiçbir şey açılmaz (contextBridge kullanılmaz).

window.addEventListener("DOMContentLoaded", () => {
  const form = document.querySelector("form");
  const input = document.querySelector("input");
  const out = document.querySelector("output");
  if (!form || !input || !out) return;
  /** findNext: yeni arama oturumu (metin değişti); değilse sonraki/önceki eşleşme. */
  const find = (forward: boolean, findNext: boolean) => ipcRenderer.send("oh:pdf-bul", { text: input.value, forward, findNext });
  const close = () => ipcRenderer.send("oh:pdf-bul-kapat");
  form.addEventListener("submit", (e) => e.preventDefault());
  input.addEventListener("input", () => find(true, true));
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") find(!e.shiftKey, false);
    else if (e.key === "Escape") close();
    else return;
    e.preventDefault();
  });
  document.getElementById("onceki")?.addEventListener("click", () => find(false, false));
  document.getElementById("sonraki")?.addEventListener("click", () => find(true, false));
  document.getElementById("kapat")?.addEventListener("click", close);
  ipcRenderer.on("oh:pdf-bul-sonuc", (_e, r: { active: number; matches: number } | null) => {
    out.textContent = !input.value || !r ? "" : r.matches ? `${r.active}/${r.matches}` : "Bulunamadı";
  });
  ipcRenderer.on("oh:pdf-bul-odak", () => {
    input.focus();
    input.select();
  });
});
