// Makine hızı başvuru ölçümü.
//
// Süre sınırlarını (bulanık testte örnek başına 5 sn) makineye göre ölçeklemek için uygulama kodundan
// bağımsız, sabit bir CPU iş yükü (sıralama, JSON, dize ve düzenli ifade) 5 kez koşulur; en kısa süre
// alınır. Temel çizgiyi üreten makinedeki süreyle oranlanır: oran = ölçülen / temel (1–5 arasına kırpılır).
export function hizOlc(tekrar = 5) {
  let enKisa = Infinity;
  for (let t = 0; t < tekrar; t++) {
    const t0 = performance.now();
    let tohum = 12345;
    const rnd = () => (tohum = (tohum * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const sayilar = Array.from({ length: 200_000 }, rnd);
    sayilar.sort((a, b) => a - b);
    const nesne = Array.from({ length: 20_000 }, (_, i) => ({ i, ad: `Hücre ${i} ığüşöç`, d: sayilar[i * 10], e: [i, i * 2, i * 3] }));
    const geri = JSON.parse(JSON.stringify(nesne));
    let s = "";
    for (let i = 0; i < 50_000; i++) s += (i % 97).toString(36);
    const eslesme = s.match(/[a-f]\d[a-z]/g)?.length ?? 0;
    const m = new Map();
    for (let i = 0; i < 100_000; i++) m.set(`k${i % 5000}`, (m.get(`k${i % 5000}`) ?? 0) + i);
    if (geri.length + eslesme + m.size < 0) console.log("olanaksız");
    enKisa = Math.min(enKisa, performance.now() - t0);
  }
  return Math.round(enKisa * 10) / 10;
}

/** Ölçülen süreyi temel makinedekine oranlar; 1'den küçükse 1 (hızlı makinede sınır daralmaz), en çok 5. */
export function hizOrani(olculen, temel) {
  if (!temel || !olculen) return 1;
  return Math.min(5, Math.max(1, olculen / temel));
}
