#!/usr/bin/env python3
"""OfficeHere'in kaydettiği dosyaları (ve özgünlerini) LibreOffice ile PDF'e çevirir.

Derlem turunun (pnpm derlem) çıktıları .derlem/rapor/derlem.json'dan okunur. Her çıktı ve özgünü
`soffice --headless --convert-to pdf` ile açılır; şunlar raporlanır:
  - LibreOffice'in açamadığı (PDF üretemediği) çıktılar
  - sayfa sayısı ve sayfa boyutu (özgün → OfficeHere)
  - boş sayfalar (metni ve görüntüsü olmayan); özgünden FAZLA boş sayfa = sorun
  - PDF metnindeki kayıp kelime oranı
  - piksel farkı: özgünün ve çıktının LO çizimi `pdftoppm -r 60` ile sayfa sayfa karşılaştırılır (ilk
    LO_PIKSEL_SAYFA sayfa; bir kanalı 64'ten çok değişen piksel farklı sayılır, boyut/sayfa sayısı farkı
    tümüyle farklı). Oran ≤ %0,5 ise uyumlu. Temel çizgi scripts/derlem/lo_piksel_temel.json: uyumluyken
    eşiği aşan ya da temelden 0,5 puandan çok artan fark = gerileme → çıkış kodu 3 (raporlar yine yazılır).
    Temel yoksa yalnız rapor. Aday temel her koşuda .derlem/rapor/lo_piksel_temel-aday.json'a yazılır
    (CI artefaktı; yazı tipi farkı yüzünden temel, CI'nın kendi koşusundan alınmalıdır).
Ön koşul: soffice, pdfinfo, pdftotext, pdfimages, pdftoppm (poppler-utils). pdftoppm yoksa piksel farkı atlanır.
Kullanım: python3 scripts/derlem/lo_pdf.py   (pnpm derlem:lo bunu çağırır) → .derlem/rapor/lo_pdf.json, lo_piksel.md
Ortam: DERLEM_CIKTI (varsayılan <kök>/.derlem), LO_PDF_PARALEL (3), LO_PDF_SINIR_SN (120),
       LO_PIKSEL=0 (piksel farkını kapatır), LO_PIKSEL_SAYFA (10), LO_PIKSEL_TEMEL_YAZ=1 (temel çizgiyi yazar)
"""
import concurrent.futures as cf
import json
import operator
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

KOK = Path(__file__).resolve().parents[2]
CIKTI = Path(os.environ.get("DERLEM_CIKTI", KOK / ".derlem")).resolve()
PDF = CIKTI / "pdf"
PARALEL = int(os.environ.get("LO_PDF_PARALEL", "3"))
SINIR = int(os.environ.get("LO_PDF_SINIR_SN", "120"))
ortam = dict(os.environ, LC_ALL="C.UTF-8", LANG="C.UTF-8")
RAPOR = CIKTI / "rapor"
PIKSEL = os.environ.get("LO_PIKSEL", "1") != "0" and shutil.which("pdftoppm") is not None
PIKSEL_SAYFA = int(os.environ.get("LO_PIKSEL_SAYFA", "10"))
PIKSEL_DPI = 60
PIKSEL_ESIK = 0.5  # %: özgün ↔ çıktı farkı bu kadar ya da azsa uyumlu; temelden bu kadar puan artış = gerileme
PIKSEL_KANAL = 64  # bir kanaldaki fark bundan büyükse piksel farklı (kenar yumuşatma gürültüsü sayılmaz)
PIKSEL_TEMEL = KOK / "scripts" / "derlem" / "lo_piksel_temel.json"

tur = json.loads((CIKTI / "rapor" / "derlem.json").read_text(encoding="utf-8"))
isler = []  # (anahtar, özgün yol, çıktı yolu)
for f in tur["dosyalar"]:
    for t in f.get("turlar") or []:
        if t.get("ciktiYolu") and not t.get("hata"):
            isler.append((f'{f["anahtar"]}→{t["bicim"]}', f["yol"], t["ciktiYolu"]))
ozgunler = sorted({o for _, o, _ in isler})


def sayfa_bilgisi(pdf):
    info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
    sayfa = re.search(r"Pages:\s+(\d+)", info)
    boyut = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", info)
    n = int(sayfa.group(1)) if sayfa else 0
    metin = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, errors="replace").stdout
    sayfalar = metin.split("\f")
    resimli = set()
    liste = subprocess.run(["pdfimages", "-list", str(pdf)], capture_output=True, text=True, errors="replace").stdout.splitlines()
    for satir in liste[2:]:
        p = satir.split()
        if p and p[0].isdigit():
            resimli.add(int(p[0]))
    bos = sum(1 for i in range(n) if not (sayfalar[i].strip() if i < len(sayfalar) else "") and (i + 1) not in resimli)
    return {
        "sayfa": n,
        "boyut": f"{float(boyut.group(1)):.0f}x{float(boyut.group(2)):.0f}" if boyut else "",
        "bos": bos,
        "metin": len(re.sub(r"\s+", "", metin)),
        "kelimeler": re.findall(r"\w+", metin.lower()),
    }


def cevir(yol, yuva):
    yol = Path(yol)
    etiket = "ozgun" if not str(yol).startswith(str(CIKTI / "cikti")) else "oh"
    hedef = PDF / etiket / (re.sub(r"[^\w.-]", "_", str(yol.relative_to(KOK) if str(yol).startswith(str(KOK)) else yol.name)))
    hedef.mkdir(parents=True, exist_ok=True)
    t = time.time()
    profil = (CIKTI / f"lo-pdf-{yuva}").as_uri()
    try:
        p = subprocess.run(["timeout", str(SINIR), "soffice", "--headless", "--norestore", f"-env:UserInstallation={profil}", "--convert-to", "pdf", "--outdir", str(hedef), str(yol)], capture_output=True, text=True, env=ortam)
        rc, err = p.returncode, (p.stderr + p.stdout)[-400:]
    except Exception as e:  # noqa: BLE001
        rc, err = -1, str(e)
    pdf = hedef / (yol.stem + ".pdf")
    sonuc = {"yol": str(yol), "rc": rc, "sn": round(time.time() - t, 1)}
    if not pdf.exists() or pdf.stat().st_size == 0:
        sonuc["hata"] = ("süre aşımı" if rc == 124 else "PDF üretilmedi") + ": " + err.strip().replace("\n", " ")[-250:]
        return sonuc
    sonuc.update(sayfa_bilgisi(pdf))
    sonuc["pdf"] = str(pdf)
    return sonuc


# --- Piksel farkı -------------------------------------------------------------------------------------

# map(operator.sub) farkı -255..255; negatif dizin sondan sayar, böylece tablo[d] = |d| > PIKSEL_KANAL.
_FARK_TABLO = bytes(1 if min(i, 511 - i) > PIKSEL_KANAL else 0 for i in range(511))


def goruntule(pdf):
    """PDF'in ilk PIKSEL_SAYFA sayfasını PIKSEL_DPI çözünürlükte RGB'ye çizer: [(genişlik, yükseklik, bayt)]."""
    with tempfile.TemporaryDirectory(dir=CIKTI) as d:
        subprocess.run(["pdftoppm", "-r", str(PIKSEL_DPI), "-l", str(PIKSEL_SAYFA), str(pdf), str(Path(d) / "s")], capture_output=True, timeout=SINIR)
        sayfalar = []
        for f in sorted(Path(d).glob("s-*.ppm"), key=lambda p: int(p.stem.rsplit("-", 1)[1])):
            b = f.read_bytes()
            m = re.match(rb"P6\s+(\d+)\s+(\d+)\s+(\d+)\s", b)
            if m and m.group(3) == b"255":
                sayfalar.append((int(m.group(1)), int(m.group(2)), b[m.end():]))
        return sayfalar


def farkli_piksel(a, b, genislik):
    """Aynı genişlikte iki RGB satır bloğunda farklı piksel sayısı."""
    toplam = 0
    adim = genislik * 3
    for i in range(0, len(a), adim):
        ra, rb = a[i:i + adim], b[i:i + adim]
        if ra == rb:
            continue
        bayrak = bytes(map(_FARK_TABLO.__getitem__, map(operator.sub, ra, rb)))
        # Her bayt 0/1: üç kanalın bayrakları VEYA'lanınca bir pikselin bayrağı olur.
        toplam += (int.from_bytes(bayrak[0::3], "big") | int.from_bytes(bayrak[1::3], "big") | int.from_bytes(bayrak[2::3], "big")).bit_count()
    return toplam


def piksel_farki(ozgun, cikti):
    """İki sayfa listesi arasında farklı piksel oranı (%); eksik sayfa ve taşan alan tümüyle farklı sayılır."""
    farkli = toplam = 0
    for i in range(max(len(ozgun), len(cikti))):
        if i >= len(ozgun) or i >= len(cikti):
            w, h, _ = (ozgun if i < len(ozgun) else cikti)[i]
            farkli += w * h
            toplam += w * h
            continue
        (wa, ha, a), (wb, hb, b) = ozgun[i], cikti[i]
        w, h = min(wa, wb), min(ha, hb)
        toplam += max(wa, wb) * max(ha, hb)
        farkli += max(wa, wb) * max(ha, hb) - w * h
        if wa != wb:
            a = b"".join(a[y * wa * 3:(y * wa + w) * 3] for y in range(h))
            b = b"".join(b[y * wb * 3:(y * wb + w) * 3] for y in range(h))
        farkli += farkli_piksel(a[: w * h * 3], b[: w * h * 3], w)
    return round(100 * farkli / toplam, 2) if toplam else 0.0


def piksel_olc(isler, sonuclar):
    """Özgün başına bir kez çizer; {çıktı yolu: (oran %, karşılaştırılan sayfa)}."""
    gruplar = {}
    for _, ozgun, cikti in isler:
        if sonuclar.get(ozgun, {}).get("pdf") and sonuclar.get(cikti, {}).get("pdf"):
            gruplar.setdefault(ozgun, []).append(cikti)

    def olc(ozgun):
        a = goruntule(sonuclar[ozgun]["pdf"])
        out = {}
        for cikti in gruplar[ozgun]:
            b = goruntule(sonuclar[cikti]["pdf"])
            out[cikti] = (piksel_farki(a, b), max(len(a), len(b)))
        return out

    sonuc = {}
    with cf.ThreadPoolExecutor(max_workers=PARALEL) as ex:
        for r in ex.map(olc, sorted(gruplar)):
            sonuc.update(r)
    return sonuc


def yuzde(v):
    return f"%{v:.2f}".replace(".", ",")


def piksel_karsilastir(simdi, temel):
    """simdi/temel: {anahtar: oran %}. (gerilemeler, iyileşmeler, temelde olmayanlar)."""
    gerileme, iyilesme, yeni = [], [], []
    for k, s in sorted(simdi.items()):
        if k not in temel:
            yeni.append(k)
            continue
        t = temel[k]
        if (t <= PIKSEL_ESIK < s) or s > t + PIKSEL_ESIK:
            gerileme.append(f"{k}: {yuzde(t)} → {yuzde(s)}")
        elif (s <= PIKSEL_ESIK < t) or s < t - PIKSEL_ESIK:
            iyilesme.append(f"{k}: {yuzde(t)} → {yuzde(s)}")
    return gerileme, iyilesme, yeni


def piksel_raporu(satirlar, sure):
    """lo_piksel.md ve aday temeli yazar; temel varsa karşılaştırır. Gerileme sayısını döndürür."""
    olculen = {s["anahtar"]: s["pikselFarki"] for s in satirlar if s.get("pikselFarki") is not None}
    aday = {
        "aciklama": "LibreOffice piksel farkı temel çizgisi (özgün ↔ OfficeHere çıktısı, pdftoppm -r 60, %). scripts/derlem/lo_pdf.py yazar; CI'daki saglamlik koşusunun lo_piksel_temel-aday.json'undan alınır.",
        "esik": PIKSEL_ESIK,
        "guncelleme": time.strftime("%Y-%m-%d"),
        "farklar": dict(sorted(olculen.items())),
    }
    (RAPOR / "lo_piksel_temel-aday.json").write_text(json.dumps(aday, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    if os.environ.get("LO_PIKSEL_TEMEL_YAZ") == "1":
        PIKSEL_TEMEL.write_text(json.dumps(aday, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        print(f"piksel temel çizgisi yazıldı: {PIKSEL_TEMEL} ({len(olculen)} çıktı)")
    uyumlu = sum(1 for v in olculen.values() if v <= PIKSEL_ESIK)
    L = ["## LibreOffice piksel farkı (özgün ↔ OfficeHere çıktısı, `pdftoppm -r 60`)", ""]
    L.append(f"{len(olculen)} çıktı ölçüldü ({sure} sn): fark ≤ %0,5 (uyumlu) **{uyumlu}**, üstünde **{len(olculen) - uyumlu}**.")
    smartart = [(k, v) for k, v in sorted(olculen.items()) if "smartart" in k.lower()]
    if smartart:
        L.append("SmartArt dosyaları (geçiş katmanı kabulü için temel): " + ", ".join(f"`{k}` {yuzde(v)}" for k, v in smartart) + ".")
    gerileme = []
    if PIKSEL_TEMEL.exists():
        temel = json.loads(PIKSEL_TEMEL.read_text(encoding="utf-8")).get("farklar", {})
        gerileme, iyilesme, yeni = piksel_karsilastir(olculen, temel)
        L += ["", f"Temel çizgiye göre: gerileme **{len(gerileme)}**, iyileşme {len(iyilesme)}, temelde olmayan {len(yeni)}."]
        for baslik, liste in (("Gerileme (HATA)", gerileme), ("İyileşme (temeli güncelleyin: LO_PIKSEL_TEMEL_YAZ=1 ya da CI aday temeli)", iyilesme)):
            if liste:
                L += ["", f"### {baslik}", ""] + [f"- {x}" for x in liste[:60]]
    else:
        L += ["", f"Temel çizgi yok ({PIKSEL_TEMEL.relative_to(KOK)}); gerileme denetlenmedi. Aday: .derlem/rapor/lo_piksel_temel-aday.json"]
    L += ["", "| çıktı | piksel farkı | sayfa |", "|---|---|---|"]
    for s in sorted((s for s in satirlar if s.get("pikselFarki") is not None), key=lambda s: -s["pikselFarki"]):
        isaret = "**" if s["pikselFarki"] > PIKSEL_ESIK else ""
        L.append(f"| {s['anahtar']} | {isaret}{yuzde(s['pikselFarki'])}{isaret} | {s['pikselSayfa']} |")
    (RAPOR / "lo_piksel.md").write_text("\n".join(L) + "\n", encoding="utf-8")
    return len(gerileme)


tum = ozgunler + [c for _, _, c in isler]
sonuclar = {}
t0 = time.time()
with cf.ThreadPoolExecutor(max_workers=PARALEL) as ex:
    gelecek = {ex.submit(cevir, f, i % PARALEL): f for i, f in enumerate(tum)}
    for g in cf.as_completed(gelecek):
        r = g.result()
        sonuclar[r["yol"]] = r
        print(("HATA " if "hata" in r else "tamam ") + os.path.basename(r["yol"]), r.get("sayfa", ""), r.get("sn"), flush=True)

sure = round(time.time() - t0, 1)
t1 = time.time()
piksel = piksel_olc(isler, sonuclar) if PIKSEL else {}
piksel_sn = round(time.time() - t1, 1)

satirlar = []
for anahtar, ozgun, cikti in isler:
    ro, rs = sonuclar.get(cikti, {}), sonuclar.get(ozgun, {})
    s = {
        "anahtar": anahtar,
        "ozgun": os.path.relpath(ozgun, KOK),
        "cikti": os.path.relpath(cikti, CIKTI),
        "ozgunHata": rs.get("hata"),
        "ohHata": ro.get("hata"),
        "ozgunSayfa": rs.get("sayfa"),
        "ohSayfa": ro.get("sayfa"),
        "ozgunBoyut": rs.get("boyut"),
        "ohBoyut": ro.get("boyut"),
        "ozgunBos": rs.get("bos"),
        "ohBos": ro.get("bos"),
    }
    if ro.get("bos") is not None:
        s["bosFazla"] = max(0, ro["bos"] - (rs.get("bos") or 0))
    if ro.get("kelimeler") is not None and rs.get("kelimeler") is not None:
        a, b = Counter(rs["kelimeler"]), Counter(ro["kelimeler"])
        top = sum(a.values()) or 1
        kayip = sum((a - b).values())
        s["kayipKelime"] = kayip
        s["kayipOrani"] = round(kayip / top, 4)
    if cikti in piksel:
        s["pikselFarki"], s["pikselSayfa"] = piksel[cikti]
    satirlar.append(s)
RAPOR.mkdir(parents=True, exist_ok=True)
(RAPOR / "lo_pdf.json").write_text(json.dumps({"sureSn": sure, "pikselSn": piksel_sn, "satirlar": satirlar}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"{len(tum)} dosya PDF'e çevrildi ({round(sure)} sn)")
if PIKSEL:
    gerileme = piksel_raporu(satirlar, piksel_sn)
    print(f"piksel farkı: {len(piksel)} çıktı ({piksel_sn} sn), temel çizgiye göre {gerileme} gerileme → {RAPOR / 'lo_piksel.md'}")
    if gerileme:
        sys.exit(3)
else:
    (RAPOR / "lo_piksel.md").unlink(missing_ok=True)
    print("piksel farkı atlandı (" + ("LO_PIKSEL=0" if os.environ.get("LO_PIKSEL") == "0" else "pdftoppm yok") + ")")
