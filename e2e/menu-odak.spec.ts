import { expect, test } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Menü kapanınca odak düzenleyiciye dönerken, menü komutunun kurduğu seçim korunur.
// Yük altında animasyon karesi gecikir; tarayıcının selectionchange olayı önce gelir.
// Odak düzenleyiciye yerel focus() ile hemen verilirse yazı eski imlece giderdi.

test("Belge: menüden eklenen metin kutusuna yazılır (geciken karede de)", async ({ page }) => {
  const errors = collectErrors(page);
  await page.addInitScript(() => {
    const raf = window.requestAnimationFrame.bind(window);
    window.requestAnimationFrame = (cb) => {
      setTimeout(() => raf(cb), 60);
      return 0;
    };
  });
  await page.goto(URL.belge);
  const body = page.locator(".bl-body .ProseMirror");
  await expect(body).toBeFocused();
  await page.keyboard.type("Gövde paragrafı");
  await page.getByRole("tab", { name: "Ekle", exact: true }).click();
  await page.locator(".oh-ribbon-body").getByRole("button", { name: "Metin Kutusu", exact: true }).first().click();
  await page.getByRole("menuitem", { name: /^Basit Metin Kutusu(\s|$)/ }).first().click();
  await page.keyboard.type("Kutudaki yazı");
  await expect(body.locator(".bl-textbox")).toHaveText("Kutudaki yazı");
  await expect(body.locator("p").filter({ hasText: "Gövde paragrafı" }).last()).toHaveText("Gövde paragrafı");
  expect(errors).toEqual([]);
});
