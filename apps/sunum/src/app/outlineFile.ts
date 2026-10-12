/** Yeni Slayt > Anahattan Slaytlar: ana hat dosyasını seçip slaytları ekler. */
import { openFile } from "@officehere/core";
import { parseDocxOutline, parseTextOutline, slidesFromOutline } from "../model/outlineImport";
import type { Actions } from "../state/actions";

export async function importOutline(a: Actions): Promise<void> {
  const f = await openFile([{ name: "Tüm Ana Hatlar", extensions: ["txt", "md", "docx"] }]);
  if (!f) return;
  try {
    const lower = f.name.toLowerCase();
    const entries = lower.endsWith(".docx") ? await parseDocxOutline(f.data) : parseTextOutline(new TextDecoder("utf-8").decode(f.data));
    if (!entries.length) {
      a.message("Ana hatta slayt oluşturulacak başlık bulunamadı.", "warning");
      return;
    }
    const slides = slidesFromOutline(a.s.pres, entries);
    const at = Math.min(a.s.slideIdx + 1, a.s.pres.slides.length);
    a.store.commit((p) => ({ ...p, slides: [...p.slides.slice(0, at), ...slides, ...p.slides.slice(at)] }), { slideIdx: at, sel: [], editing: null });
    a.message(`Ana hattan ${slides.length} slayt eklendi.`, "success");
  } catch (e) {
    a.message(`Ana hat okunamadı: ${(e as Error).message}`, "error");
  }
}
