/**
 * Ekle › Yakınlaştırma: Slayt Yakınlaştırma (geçerli slayda seçili slaytların
 * canlı küçük resimleri) ve Özet Yakınlaştırma (seçili slaytlardan başlayan
 * bölümlere giden özet slaydı). Gösteride tıklanınca hedefe yakınlaşılır.
 */
import { slideToPng } from "../app/exporting";
import { uid } from "../model/factory";
import { updateSlideShapes } from "../model/ops";
import { newSlide } from "../model/presentation";
import { addSectionAt } from "../model/sections";
import type { PicShape, Presentation, Slide } from "../model/types";
import { slideTitle } from "../show/slideTitle";
import type { Actions } from "./actions";

const blobToDataUrl = (b: Blob) =>
  new Promise<string>((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = () => rej(r.error);
    r.readAsDataURL(b);
  });

/** Hedef slaydın PNG anlık görüntüsü (dosyaya yazılan yedek resim). */
async function snapshot(pres: Presentation, index: number): Promise<string> {
  try {
    const b = await slideToPng(pres, index, 0.5);
    if (b) return await blobToDataUrl(b);
  } catch {
    // tuval kullanılamıyorsa boş resim
  }
  return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4////fwAJ+wP9KobjigAAAABJRU5ErkJggg==";
}

/** Izgara yerleşimi: n kutu slaydın orta bölgesine sığdırılır. */
function grid(pres: Presentation, n: number, top: number): { x: number; y: number; w: number; h: number }[] {
  const cols = Math.min(n, n <= 4 ? n : Math.ceil(Math.sqrt(n * 1.6)));
  const rows = Math.ceil(n / cols);
  const gap = pres.width * 0.025;
  const availW = pres.width * 0.84;
  const availH = pres.height - top - pres.height * 0.08;
  let w = (availW - gap * (cols - 1)) / cols;
  let h = (w * pres.height) / pres.width;
  if (h * rows + gap * (rows - 1) > availH) {
    h = (availH - gap * (rows - 1)) / rows;
    w = (h * pres.width) / pres.height;
  }
  const x0 = (pres.width - (w * cols + gap * (cols - 1))) / 2;
  return Array.from({ length: n }, (_, i) => ({ x: x0 + (i % cols) * (w + gap), y: top + Math.floor(i / cols) * (h + gap), w, h }));
}

async function zoomPic(pres: Presentation, target: Slide, box: { x: number; y: number; w: number; h: number }, summary: boolean, n: number): Promise<PicShape> {
  const i = pres.slides.indexOf(target);
  return {
    kind: "pic",
    id: uid(),
    name: `Slayt Yakınlaştırma ${n}`,
    descr: slideTitle(target) || `Slayt ${i + 1}`,
    ...box,
    rot: 0,
    src: await snapshot(pres, i),
    line: { fill: { type: "solid", color: { scheme: "dk1", lumMod: 0.25, lumOff: 0.75 } }, width: 0.75, dash: "solid" },
    link: `#slide:${target.id}`,
    zoom: { slideId: target.id, summary: summary || undefined, ret: summary || undefined },
  };
}

/** Slayt Yakınlaştırma: seçili slaytların yakınlaştırmalarını geçerli slayda ekler. */
export async function insertSlideZoom(a: Actions, slideIds: string[]) {
  const pres = a.s.pres;
  const targets = slideIds.map((id) => pres.slides.find((s) => s.id === id)).filter((s): s is Slide => !!s);
  if (!targets.length || !a.slide()) return;
  // Yakınlaştırmalar slaydın alt bölümüne, içeriği örtmeyecek boyutta yerleşir.
  const bw = Math.min(pres.width * 0.22, (pres.width * 0.84) / targets.length - pres.width * 0.02);
  const bh = (bw * pres.height) / pres.width;
  const x0 = (pres.width - targets.length * bw - (targets.length - 1) * pres.width * 0.02) / 2;
  const boxes = targets.map((_, k) => ({ x: x0 + k * (bw + pres.width * 0.02), y: pres.height - bh - pres.height * 0.08, w: bw, h: bh }));
  const pics: PicShape[] = [];
  for (const [k, t] of targets.entries()) pics.push(await zoomPic(pres, t, boxes[k], false, a.nextShapeNumber() + k));
  const i = a.s.slideIdx;
  a.store.commit((p) => updateSlideShapes(p, i, (list) => [...list, ...pics]), { sel: pics.map((p) => p.id), editing: null });
}

/** Özet Yakınlaştırma: seçili slaytlar bölüm başı olur; önlerine özet slaydı eklenir. */
export async function insertSummaryZoom(a: Actions, slideIds: string[]) {
  let pres = a.s.pres;
  const ids = pres.slides.filter((s) => slideIds.includes(s.id)).map((s) => s.id);
  if (!ids.length) return;
  // Bölümler: her seçili slayt bir bölüm başlatır (adı slayt başlığı).
  let slides = pres.slides;
  for (const id of ids) {
    const idx = slides.findIndex((s) => s.id === id);
    if (slides[idx].section) continue;
    const r = addSectionAt(slides, idx, slideTitle(slides[idx]) || `Bölüm ${ids.indexOf(id) + 1}`);
    if (r) slides = r.slides;
  }
  pres = { ...pres, slides };
  // Özet slaydı: Yalnızca Başlık düzeninde, ilk seçili slayttan önce.
  const first = slides.findIndex((s) => s.id === ids[0]);
  const sum = newSlide(pres, "titleOnly");
  const title = sum.shapes.find((s) => s.kind === "sp" && (s.ph?.type === "title" || s.ph?.type === "ctrTitle"));
  if (title && title.kind === "sp") title.text = { ...(title.text ?? { paragraphs: [] }), paragraphs: [{ runs: [{ text: "Özet", props: {} }], level: 0, props: {} }] };
  const boxes = grid(pres, ids.length, pres.height * 0.3);
  const pics: PicShape[] = [];
  for (const [k, id] of ids.entries()) pics.push(await zoomPic(pres, slides.find((s) => s.id === id)!, boxes[k], true, k + 1));
  pics.forEach((p, k) => (p.name = `Özet Yakınlaştırma ${k + 1}`));
  sum.shapes = [...sum.shapes, ...pics];
  // Özet slaydı önceki bölümün sonunda kalır; ilk slayttan önceyse kendi bölümünü açar.
  const out = [...slides.slice(0, first), sum, ...slides.slice(first)];
  if (first === 0) {
    out[0] = { ...out[0], section: { id: uid("sec"), name: "Özet" } };
  }
  a.store.commit((p) => ({ ...p, slides: out }), { slideIdx: first, sel: pics.map((x) => x.id), editing: null });
}
