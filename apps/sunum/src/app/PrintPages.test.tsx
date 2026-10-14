// @vitest-environment jsdom
/** Yazdırma kökü: StrictMode'un çift etki kurulumunda tek yazdırma, kapanınca iptal. */
import { StrictMode } from "react";
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_PRINT_SETTINGS } from "../io/presProps";
import { newPresentation } from "../model/presentation";
import { PrintRoot } from "./PrintPages";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

describe("PrintRoot", () => {
  let host: HTMLDivElement;
  beforeEach(() => {
    host = document.createElement("div");
    document.body.appendChild(host);
    if (!globalThis.requestAnimationFrame) vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => setTimeout(() => cb(0), 0));
  });
  afterEach(() => {
    host.remove();
    vi.restoreAllMocks();
  });

  const render = (props: { onDone: () => void; onReady?: () => Promise<void> }) => {
    const root = createRoot(host);
    const pres = newPresentation("varsayilan", 0);
    act(() =>
      root.render(
        <StrictMode>
          <PrintRoot pres={pres} pages={[]} size={{ w: 1280, h: 720 }} settings={DEFAULT_PRINT_SETTINGS} {...props} />
        </StrictMode>,
      ),
    );
    return root;
  };

  it("StrictMode'da etki iki kez kurulsa da bir kez yazdırır ve bir kez biter", async () => {
    const print = vi.spyOn(window, "print").mockImplementation(() => undefined);
    const onDone = vi.fn();
    const root = render({ onDone });
    // İki kare + yazı tipi beklemesi yavaş makinede 50 ms'yi aşabilir: sabit süre yerine koşul beklenir.
    await vi.waitFor(() => expect(print).toHaveBeenCalledTimes(1));
    window.dispatchEvent(new Event("afterprint"));
    await wait(600);
    // afterprint ve yedek zamanlayıcı birlikte gelse de tek bitiş.
    expect(onDone).toHaveBeenCalledTimes(1);
    act(() => root.unmount());
    expect(document.documentElement.classList.contains("is-printing")).toBe(false);
  });

  it("doğrudan PDF (onReady) da bir kez çağrılır; yazdırma penceresi açılmaz", async () => {
    const print = vi.spyOn(window, "print").mockImplementation(() => undefined);
    const onReady = vi.fn(async () => undefined);
    const onDone = vi.fn();
    const root = render({ onDone, onReady });
    await vi.waitFor(() => expect(onDone).toHaveBeenCalledTimes(1));
    expect(onReady).toHaveBeenCalledTimes(1);
    expect(print).not.toHaveBeenCalled();
    expect(onDone).toHaveBeenCalledTimes(1);
    act(() => root.unmount());
  });

  it("zincir bitmeden kapanırsa yazdırılmaz", async () => {
    const print = vi.spyOn(window, "print").mockImplementation(() => undefined);
    const root = render({ onDone: vi.fn() });
    act(() => root.unmount());
    await wait(200);
    expect(print).not.toHaveBeenCalled();
  });
});
