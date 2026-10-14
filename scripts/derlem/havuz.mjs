// Zaman aşımlı işçi havuzu.
//
// Her iş ayrı bir worker_threads işçisinde, süre sınırıyla çalışır. Sınır aşılırsa işçi sonlandırılır
// (senkron sonsuz döngü de böylece kesilir) ve yerine yenisi açılır. Bellek sınırı aşımı
// (ERR_WORKER_OUT_OF_MEMORY) ve işçinin beklenmedik çıkışı ayrıca raporlanır.
import { Worker } from "node:worker_threads";

const ISCI = new URL("./isci.mjs", import.meta.url);

export class Havuz {
  /** @param {{boyut?: number, bellekMb?: number, veri?: object}} o veri: işçiye workerData olarak geçer (paket dizini vb.) */
  constructor({ boyut = 3, bellekMb = 3072, veri = {} } = {}) {
    this.boyut = boyut;
    this.bellekMb = bellekMb;
    this.veri = veri;
    this.yuvalar = [];
    this.kuyruk = [];
    this.sira = 0;
  }

  async ac() {
    const w = new Worker(ISCI, { workerData: this.veri, resourceLimits: { maxOldGenerationSizeMb: this.bellekMb }, stderr: true, stdout: true });
    // Okuyucuların console.warn çıktıları rapor gürültüsü olmasın diye yutulur.
    w.stderr.on("data", () => {});
    w.stdout.on("data", () => {});
    await new Promise((res, rej) => {
      const on = (m) => {
        if (m?.hazir) {
          w.off("message", on);
          res();
        } else if (m?.yuklemeHatasi) rej(new Error(`İşçi başlatılamadı: ${m.yuklemeHatasi}`));
      };
      w.on("message", on);
      w.once("error", rej);
    });
    return w;
  }

  /** İşi çalıştırır; sonuç: {durum: "tamam"|"zaman-asimi"|"bellek"|"isci-hatasi"|"cikis", cikti?, hata?, sureMs} */
  async calistir(is, sinirMs) {
    let yuva = this.yuvalar.find((s) => !s.mesgul);
    if (!yuva && this.yuvalar.length < this.boyut) {
      yuva = { w: null, mesgul: true };
      this.yuvalar.push(yuva);
      yuva.w = await this.ac();
    } else if (!yuva) {
      await new Promise((r) => this.kuyruk.push(r));
      return this.calistir(is, sinirMs);
    }
    yuva.mesgul = true;
    const id = ++this.sira;
    const w = yuva.w;
    const t0 = performance.now();
    const sonuc = await new Promise((resolve) => {
      let bitti = false;
      const bitir = (v) => {
        if (bitti) return;
        bitti = true;
        clearTimeout(zamanlayici);
        w.off("message", onMesaj);
        w.off("error", onHata);
        w.off("exit", onCikis);
        resolve(v);
      };
      const onMesaj = (m) => m?.id === id && bitir({ durum: "tamam", cikti: m.cikti });
      const onHata = (e) =>
        bitir({
          durum: e?.code === "ERR_WORKER_OUT_OF_MEMORY" ? "bellek" : "isci-hatasi",
          hata: { ileti: String(e?.message ?? e), yigin: String(e?.stack ?? "").split("\n").slice(0, 8).join("\n"), kod: e?.code },
        });
      const onCikis = (c) => bitir({ durum: "cikis", kod: c });
      const zamanlayici = setTimeout(() => bitir({ durum: "zaman-asimi" }), sinirMs);
      w.on("message", onMesaj);
      w.once("error", onHata);
      w.once("exit", onCikis);
      w.postMessage({ ...is, id });
    });
    sonuc.sureMs = Math.round(performance.now() - t0);
    if (sonuc.durum !== "tamam") {
      // İşçi güvenilmez durumda: sonlandır, yenisini aç.
      try {
        await w.terminate();
      } catch {
        /* zaten kapanmış */
      }
      yuva.w = await this.ac();
    }
    yuva.mesgul = false;
    this.kuyruk.shift()?.();
    return sonuc;
  }

  async kapat() {
    await Promise.all(this.yuvalar.map((s) => s.w?.terminate()));
  }
}
