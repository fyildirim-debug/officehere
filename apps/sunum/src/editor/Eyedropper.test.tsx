// @vitest-environment jsdom
/** Damlalık: kip yeniden açılınca önceki oturumun "hazır" durumu taşınmaz; çizimden önce gelen imleç konumu hazır olunca gösterilir. */
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { newPresentation } from "../model/presentation";
import { EyedropperOverlay, startEyedropper, stopEyedropper } from "./Eyedropper";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

/** 4×4 kırmızı tuval yerine geçen sahte çizim (jsdom'da tuval yok). */
const fakeCanvas = () =>
  ({
    width: 4,
    height: 4,
    getContext: () => ({ getImageData: () => ({ width: 4, height: 4, data: new Uint8ClampedArray(64).map((_, i) => (i % 4 === 0 || i % 4 === 3 ? 255 : 0)) }) }),
  }) as unknown as HTMLCanvasElement;

describe("EyedropperOverlay", () => {
  let host: HTMLDivElement;
  let root: Root;
  let pending: Array<() => void>;
  beforeEach(() => {
    host = document.createElement("div");
    document.body.appendChild(host);
    root = createRoot(host);
    pending = [];
  });
  afterEach(() => {
    act(() => root.unmount());
    host.remove();
    stopEyedropper();
  });

  const render = () => {
    const pres = newPresentation("varsayilan", 0);
    // Çizim elle bitirilir: "hazır" olmadan önceki durum sınanabilsin.
    const rasterize = () => new Promise<HTMLCanvasElement>((r) => pending.push(() => r(fakeCanvas())));
    act(() => root.render(<EyedropperOverlay pres={pres} slideIdx={0} toSlide={(x, y) => [x, y]} rasterize={rasterize} />));
  };
  const layer = () => host.querySelector(".eyedropper-layer");
  const finishRaster = async () => {
    await act(async () => {
      pending.shift()?.();
      await Promise.resolve();
    });
  };
  const move = (x: number, y: number) =>
    act(() => {
      layer()!.dispatchEvent(new MouseEvent("pointermove", { bubbles: true, clientX: x, clientY: y }));
    });

  it("ikinci açılışta çizim bitene dek hazır görünmez", async () => {
    render();
    act(() => startEyedropper(() => undefined));
    await finishRaster();
    expect(layer()!.classList.contains("is-ready")).toBe(true);
    act(() => stopEyedropper());
    expect(layer()).toBeNull();
    // Katman bir kare bile "hazır" çizilmemeli (etki sonradan düzeltse de test o anı yakalar).
    const seen = new MutationObserver(() => undefined);
    seen.observe(host, { subtree: true, childList: true, attributes: true, attributeOldValue: true });
    act(() => startEyedropper(() => undefined));
    const classes = seen.takeRecords().flatMap((r) => (r.type === "attributes" ? [r.oldValue ?? ""] : [...r.addedNodes].map((n) => (n as Element).className ?? "")));
    seen.disconnect();
    expect(classes.filter((c) => c.includes("is-ready"))).toEqual([]);
    expect(layer()!.classList.contains("is-ready")).toBe(false);
    expect(host.querySelector(".eyedropper-tip")).toBeNull();
    await finishRaster();
    expect(layer()!.classList.contains("is-ready")).toBe(true);
  });

  it("çizimden önce gelen imleç konumunun rengi hazır olunca gösterilir", async () => {
    render();
    act(() => startEyedropper(() => undefined));
    move(1, 1);
    expect(host.querySelector(".eyedropper-tip")).toBeNull();
    await finishRaster();
    expect(host.querySelector(".eyedropper-tip")?.getAttribute("data-color")).toBe("#FF0000");
  });
});
