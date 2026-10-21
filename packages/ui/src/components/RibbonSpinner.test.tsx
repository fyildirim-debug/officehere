// @vitest-environment jsdom
/** Şerit sayı kutusu: yazılan değer Enter ya da odak kaybında bir kez uygulanır; yazılmadıysa odak kaybı değeri geri yazmaz. */
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { RibbonSpinner } from "./Ribbon";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe("RibbonSpinner", () => {
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

  const render = (value: number, onChange: (v: number) => void) => act(() => root.render(<RibbonSpinner label="Yükseklik:" value={value} onChange={onChange} unit="cm" decimals={2} />));
  const input = () => host.querySelector("input")!;
  /** React'in denetimli girdisine kullanıcı yazması. */
  const type = (v: string) =>
    act(() => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(input(), v);
      input().dispatchEvent(new Event("input", { bubbles: true }));
    });
  const key = (k: string) => act(() => void input().dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true })));
  const blur = () => act(() => void input().dispatchEvent(new FocusEvent("focusout", { bubbles: true })));

  it("yazılan değer Enter ile bir kez uygulanır; ardından gelen odak kaybı yinelemez", () => {
    const onChange = vi.fn();
    render(0.5, onChange);
    expect(input().value).toBe("0,5 cm");
    type("2");
    key("Enter");
    blur();
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange).toHaveBeenCalledWith(2);
  });

  it("yazılmadıysa odak kaybı eski değeri geri yazmaz (değeri başka bir komut değiştirse bile)", () => {
    const onChange = vi.fn();
    render(2, onChange);
    // Başka bir komut değeri değiştirir; kutu henüz eski metni gösterirken odağı kaybeder.
    render(1, onChange);
    blur();
    expect(onChange).not.toHaveBeenCalled();
    expect(input().value).toBe("1 cm");
  });

  it("geçersiz metin uygulanmaz ve kutu geçerli değere döner", () => {
    const onChange = vi.fn();
    render(1.25, onChange);
    type("abc");
    blur();
    expect(onChange).not.toHaveBeenCalled();
    expect(input().value).toBe("1,25 cm");
  });
});
