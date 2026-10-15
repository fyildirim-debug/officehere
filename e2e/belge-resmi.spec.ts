import { readFileSync } from "node:fs";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, zipEntry } from "./ortak";

// OfficeHere Belge: Resmî Yazışma (2020) şablon kümesi ve Yönetmeliğe uygunluk denetimi. Şablon galerisi ve form
// (Dosya › Yeni › Resmî yazışma), ekrandaki ölçüler, Gözden Geçir › Resmî Yazı Denetimi, komut araması ve .docx kaydı.

const CM = 96 / 2.54;
const body = (page: Page) => page.locator(".bl-body .ProseMirror");
const pencere = (page: Page) => page.getByRole("dialog", { name: "Resmî Yazışma Şablonları" });
const denetim = (page: Page) => page.getByRole("dialog", { name: "Resmî Yazı Denetimi" });
const kural = (page: Page, k: string) => denetim(page).locator(`[data-kural="${k}"]`);

let errors: string[] = [];

test.beforeEach(async ({ page }) => {
  errors = collectErrors(page);
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
  await page.goto(URL.belge);
  await expect(body(page)).toBeFocused();
});

test.afterEach(() => {
  expect(errors).toEqual([]);
});

async function galeriyiAc(page: Page, sablon: string) {
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Yeni" }).first().click();
  await page.locator(".oh-template", { hasText: "Resmî yazışma" }).click();
  await expect(pencere(page)).toBeVisible();
  await pencere(page).getByRole("button", { name: sablon, exact: true }).click();
  await pencere(page).getByPlaceholder("GG.AA.YYYY").fill("15.10.2026");
}

async function olustur(page: Page, ad: string) {
  await pencere(page).getByRole("button", { name: "Oluştur", exact: true }).click();
  await expect(pencere(page)).toBeHidden();
  await expect(page.locator(".oh-titlebar")).toContainText(ad);
}

async function denetimiAc(page: Page) {
  await page.getByRole("tab", { name: "Gözden Geçir", exact: true }).click();
  await page.getByRole("button", { name: "Resmî Yazı Denetimi" }).click();
  await expect(denetim(page)).toBeVisible();
}

test("Dosya › Yeni › Resmî yazışma: galeri ve form; alanlar belgeye işlenir, ölçüler ekranda doğru", async ({ page }) => {
  await galeriyiAc(page, "Üst yazı");
  const p = pencere(page);
  // Form anında uyarır: konu sonunda noktalama ve küçük harf (m.13), muhatapta yönelme eki (m.14).
  await p.getByLabel("Konu", { exact: true }).fill("hizmet içi eğitim takvimi.");
  await expect(p.getByRole("note").filter({ hasText: "sonunda noktalama" })).toBeVisible();
  await expect(p.getByRole("note").filter({ hasText: "baş harfleri büyük" })).toBeVisible();
  await p.getByLabel("Konu", { exact: true }).fill("Hizmet İçi Eğitim Takvimi");
  await expect(p.getByRole("note")).toHaveCount(0);
  await p.getByLabel("Muhatap", { exact: true }).fill("Örnek Valiliği");
  await expect(p.getByRole("note").filter({ hasText: "yönelme ekiyle" })).toBeVisible();
  await p.getByLabel("Muhatap", { exact: true }).fill("örnek valiliğine");
  await expect(p.getByRole("note")).toHaveCount(0);
  await p.getByLabel("Başlık türü").selectOption({ label: "İlçe Millî Eğitim Müdürlüğü" });
  await expect(p.getByLabel("Kaymakamlık (2. satır)")).toHaveValue("ÖRNEK KAYMAKAMLIĞI");
  await olustur(page, "Üst yazı");

  const b = body(page);
  await expect(b.locator("p").nth(1)).toHaveText("ÖRNEK KAYMAKAMLIĞI");
  await expect(b.locator("p").nth(2)).toHaveText("İlçe Millî Eğitim Müdürlüğü");
  await expect(b.locator("p", { hasText: "ÖRNEK VALİLİĞİNE" })).toHaveCSS("text-align", "center");
  await expect(b.locator("p", { hasText: "Hizmet İçi Eğitim Takvimi" })).toBeVisible();

  const olcu = await page.evaluate(() => {
    const ps = [...document.querySelectorAll<HTMLElement>(".bl-body .ProseMirror p")];
    const bul = (re: RegExp) => ps.find((x) => re.test(x.textContent ?? ""))!;
    const sayfa = document.querySelector<HTMLElement>('.bl-page[data-page="0"]')!.getBoundingClientRect();
    const sayi = bul(/^Sayı/);
    // Tarih metninin kutusu (sayı satırının son metin düğümü).
    const r = document.createRange();
    const son = [...sayi.querySelectorAll("*")]
      .flatMap((e) => [...e.childNodes])
      .concat([...sayi.childNodes])
      .filter((n) => n.nodeType === 3 && /15\.10\.2026/.test(n.textContent ?? ""))[0]!;
    r.selectNodeContents(son);
    const tarih = r.getBoundingClientRect();
    const satir = sayi.getBoundingClientRect();
    const metin = bul(/^İlgi yazı ile/);
    const imza = ps.filter((x) => x.textContent === "Ad SOYAD")[0];
    const cs = getComputedStyle(metin);
    return {
      solKenar: satir.left - sayfa.left,
      sagKenar: sayfa.right - tarih.right,
      ayniSatir: Math.abs(tarih.top - (satir.top + (satir.height - tarih.height) / 2)) < 4 && satir.height < 30,
      girinti: parseFloat(cs.textIndent),
      yasla: cs.textAlign,
      font: cs.fontFamily,
      punto: parseFloat(cs.fontSize),
      imzaSol: imza.getBoundingClientRect().left + parseFloat(getComputedStyle(imza).paddingLeft) - sayfa.left,
      imzaHiza: getComputedStyle(imza).textAlign,
      imzaSag: sayfa.right - imza.getBoundingClientRect().right,
    };
  });
  // m.8: sol ve sağ kenar 1,5 cm; m.12: tarih sayı ile aynı satırda, yazı alanının en sağında.
  expect(olcu.solKenar / CM).toBeCloseTo(1.5, 1);
  expect(olcu.sagKenar / CM).toBeCloseTo(1.5, 1);
  expect(olcu.ayniSatir).toBe(true);
  // m.16/4: 1,25 cm girinti, iki yana yaslı; m.7: Times New Roman 12 (16 px).
  expect(olcu.girinti / CM).toBeCloseTo(1.25, 1);
  expect(olcu.yasla).toBe("justify");
  expect(olcu.font).toMatch(/^"?Times New Roman/);
  expect(olcu.punto).toBeCloseTo(16, 0);
  // m.17: imza bloğu sağda (blok 1,5 + 11 cm'den sağ kenara), ortalı.
  expect(olcu.imzaSol / CM).toBeCloseTo(12.5, 1);
  expect(olcu.imzaSag / CM).toBeCloseTo(1.5, 1);
  expect(olcu.imzaHiza).toBe("center");
});

test("Gözden Geçir › Resmî Yazı Denetimi: şablon uygun; bozulan konu hata olur, Göster bulguya gider; düzey seçimi", async ({ page }) => {
  await galeriyiAc(page, "Üst yazı");
  await olustur(page, "Üst yazı");
  await denetimiAc(page);
  const d = denetim(page);
  await expect(d.locator("[data-kural]")).toHaveCount(7);
  await expect(d.locator('[data-durum="hata"], [data-durum="uyari"]')).toHaveCount(0);
  await expect(d.getByRole("status")).toContainText("hata ya da uyarı yok");
  await expect(kural(page, "kenar")).toContainText("1,5 cm");
  await expect(kural(page, "yaziTipi")).toContainText("Times New Roman 12 punto");
  await d.getByRole("button", { name: "Kapat", exact: true }).last().click();
  await expect(d).toBeHidden();

  // Konunun sonuna nokta ekle.
  await body(page).locator("p", { hasText: "Hizmet İçi Eğitim Programı" }).click();
  await page.keyboard.press("End");
  await page.keyboard.type(".");
  await denetimiAc(page);
  await expect(kural(page, "konu")).toHaveAttribute("data-durum", "hata");
  await expect(kural(page, "konu")).toContainText("noktalama");
  await kural(page, "konu").getByRole("button", { name: "Göster" }).click();
  await expect(denetim(page)).toBeHidden();
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString() ?? "")).toContain("Hizmet İçi Eğitim Programı.");

  // Muhatap düzeyi "alt makam" seçilince "arz ederim" kapanışı uyumsuz.
  await denetimiAc(page);
  await expect(kural(page, "kapanis")).toHaveAttribute("data-durum", "bilgi");
  await denetim(page).getByLabel("Muhatap düzeyi (arz / rica uyumu için)").selectOption({ label: "Alt makam (rica ederim)" });
  await expect(kural(page, "kapanis")).toHaveAttribute("data-durum", "hata");
  await expect(kural(page, "kapanis")).toContainText("“… rica ederim.” olmalı");
});

test("Şablon penceresi İptal ile kapanır; Dosya › Yeni sayfası yerinde kalır, belge değişmez", async ({ page }) => {
  await body(page).pressSequentially("Mevcut metin");
  await galeriyiAc(page, "Dilekçe");
  await pencere(page).getByRole("button", { name: "İptal", exact: true }).click();
  await expect(pencere(page)).toBeHidden();
  await expect(page.locator(".oh-bs-h1")).toHaveText("Yeni");
  await expect(page.locator(".oh-template", { hasText: "Resmî yazışma" })).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(body(page)).toHaveText("Mevcut metin");
});

test("Denetim boş belgede kenar ve yazı tipi hatalarını gösterir; kenarlar tek tıkla düzelir", async ({ page }) => {
  await body(page).pressSequentially("Deneme metni");
  await denetimiAc(page);
  await expect(kural(page, "kenar")).toHaveAttribute("data-durum", "hata");
  await expect(kural(page, "yaziTipi")).toHaveAttribute("data-durum", "hata");
  await kural(page, "kenar").getByRole("button", { name: "Düzelt" }).click();
  await expect(kural(page, "kenar")).toHaveAttribute("data-durum", "uygun");
});

test("Komut araması 'resmî yazı' ile denetimi bulur ve açar", async ({ page }) => {
  await page.keyboard.press("Alt+q");
  await page.keyboard.type("resmî yazı");
  const ilk = page.getByRole("listbox", { name: "Arama sonuçları" }).getByRole("option").first();
  await expect(ilk).toContainText("Resmî Yazı Denetimi");
  await expect(ilk).toContainText("Gözden Geçir");
  await ilk.click();
  await expect(denetim(page)).toBeVisible();
});

test("Olur yazısı .docx olarak kaydedilir: kenar, girinti, yazı tipi ve alt bilgi; yeniden açılınca denetimden geçer", async ({ page }, testInfo) => {
  await galeriyiAc(page, "Olur yazısı");
  await olustur(page, "Olur yazısı");
  await expect(body(page).locator("p", { hasText: /^OLUR$/ })).toHaveCSS("text-align", "center");
  const indirme = page.waitForEvent("download");
  await page.keyboard.press("Control+s");
  const yol = testInfo.outputPath("olur.docx");
  await (await indirme).saveAs(yol);
  const dosya = readFileSync(yol);
  const belge = zipEntry(dosya, "word/document.xml") ?? "";
  expect(belge).toMatch(/<w:pgMar w:top="850" w:right="850" w:bottom="\d+" w:left="850"/);
  expect(belge).toContain('<w:ind w:firstLine="709"/>');
  expect(belge).toContain('<w:tab w:val="right" w:pos="10205"/>');
  expect(zipEntry(dosya, "word/styles.xml")).toContain('w:ascii="Times New Roman"');
  expect(zipEntry(dosya, "word/footer1.xml") ?? zipEntry(dosya, "word/footer2.xml")).toContain("ResmiAltBilgi");

  await page.setInputFiles("#bl-open-input", { name: "olur-yeniden.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: dosya });
  await expect(page.locator(".oh-titlebar")).toContainText("olur-yeniden");
  await denetimiAc(page);
  await expect(denetim(page).locator('[data-durum="hata"], [data-durum="uyari"]')).toHaveCount(0);
});
