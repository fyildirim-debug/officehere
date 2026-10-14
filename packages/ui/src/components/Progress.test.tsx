// @vitest-environment jsdom
/** İlerleme göstergesi: metin, erişilebilirlik öznitelikleri, belirsiz durum ve klavyeyle İptal. */
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Progress, progressPercent, type ProgressProps } from "./Progress";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe("Progress", () => {
  let host: HTMLDivElement;
  let root: Root;
  beforeEach(() => {
    host = document.createElement("div");
    document.body.appendChild(host);
    root = createRoot(host);
  });
  afterEach(() => {
    act(() => root.unmount());
    host.remove();
  });

  const render = (props: ProgressProps) => act(() => root.render(<Progress {...props} />));
  const bar = () => host.querySelector<HTMLElement>("[role=progressbar]")!;

  it('"Açılıyor… %40 İptal" biçiminde çizilir', () => {
    render({ label: "Açılıyor…", value: 0.4, onCancel: () => undefined });
    expect(host.textContent).toBe("Açılıyor…%40İptal");
    expect(host.querySelector<HTMLElement>(".oh-progress-fill")!.style.width).toBe("40%");
  });

  it("progressbar rolü, değer öznitelikleri ve etiket bağlantısı doğru", () => {
    render({ label: "Açılıyor…", value: 0.4, detail: "Sayfalar okunuyor" });
    const el = bar();
    expect(el.getAttribute("aria-valuenow")).toBe("40");
    expect(el.getAttribute("aria-valuemin")).toBe("0");
    expect(el.getAttribute("aria-valuemax")).toBe("100");
    expect(el.getAttribute("aria-valuetext")).toBe("%40, Sayfalar okunuyor");
    const label = document.getElementById(el.getAttribute("aria-labelledby")!);
    expect(label?.textContent).toBe("Açılıyor…");
    // Görünen yüzde ekran okuyucuda ikinci kez okunmaz.
    expect(host.querySelector(".oh-progress-pct")!.getAttribute("aria-hidden")).toBe("true");
    expect(el.getAttribute("data-tip")).toBe("Sayfalar okunuyor");
  });

  it("değer güncellenince çubuk ve yüzde güncellenir", () => {
    render({ label: "Açılıyor…", value: 0.1 });
    render({ label: "Açılıyor…", value: 0.735 });
    expect(bar().getAttribute("aria-valuenow")).toBe("74");
    expect(host.querySelector(".oh-progress-pct")!.textContent).toBe("%74");
  });

  it("oran verilmezse belirsizdir: aria-valuenow ve yüzde yok", () => {
    render({ label: "Kaydediliyor…" });
    expect(host.querySelector(".oh-progress")!.classList.contains("is-indeterminate")).toBe(true);
    expect(bar().hasAttribute("aria-valuenow")).toBe(false);
    expect(bar().hasAttribute("aria-valuetext")).toBe(false);
    expect(host.querySelector(".oh-progress-pct")).toBeNull();
    expect(host.textContent).toBe("Kaydediliyor…");
  });

  it("İptal düğmesi gerçek düğmedir: sekmeyle odaklanır, etkinleşince onCancel çağrılır", () => {
    const onCancel = vi.fn();
    render({ label: "Açılıyor…", value: 0.4, onCancel });
    const btn = host.querySelector("button")!;
    expect(btn.getAttribute("type")).toBe("button");
    expect(btn.tabIndex).toBe(0);
    btn.focus();
    expect(document.activeElement).toBe(btn);
    act(() => btn.click());
    expect(onCancel).toHaveBeenCalledTimes(1);
  });

  it("onCancel yoksa düğme yok; düğme metni değiştirilebilir", () => {
    render({ label: "Açılıyor…", value: 0.4 });
    expect(host.querySelector("button")).toBeNull();
    render({ label: "Dışa aktarılıyor…", value: 0.4, onCancel: () => undefined, cancelLabel: "Durdur" });
    expect(host.querySelector("button")!.textContent).toBe("Durdur");
  });

  it("yüzde sınırlanır ve yuvarlanır", () => {
    expect(progressPercent(-0.2)).toBe(0);
    expect(progressPercent(0.004)).toBe(0);
    expect(progressPercent(0.996)).toBe(100);
    expect(progressPercent(1.7)).toBe(100);
    expect(progressPercent(Number.NaN)).toBeNull();
    expect(progressPercent(undefined)).toBeNull();
    expect(progressPercent(null)).toBeNull();
  });
});
