#!/usr/bin/env python3
"""OfficeHere'in kaydettiği dosyaları (ve özgünlerini) LibreOffice ile PDF'e çevirir.

Derlem turunun (pnpm derlem) çıktıları .derlem/rapor/derlem.json'dan okunur. Her çıktı ve özgünü
`soffice --headless --convert-to pdf` ile açılır; şunlar raporlanır:
  - LibreOffice'in açamadığı (PDF üretemediği) çıktılar
  - sayfa sayısı ve sayfa boyutu (özgün → OfficeHere)
  - boş sayfalar (metni ve görüntüsü olmayan); özgünden FAZLA boş sayfa = sorun
  - PDF metnindeki kayıp kelime oranı
Ön koşul: soffice, pdfinfo, pdftotext, pdfimages (poppler-utils).
Kullanım: python3 scripts/derlem/lo_pdf.py   (pnpm derlem:lo bunu çağırır) → .derlem/rapor/lo_pdf.json
Ortam: DERLEM_CIKTI (varsayılan <kök>/.derlem), LO_PDF_PARALEL (3), LO_PDF_SINIR_SN (120)
"""
import concurrent.futures as cf
import json
import os
import re
import subprocess
import time
from collections import Counter
from pathlib import Path

KOK = Path(__file__).resolve().parents[2]
CIKTI = Path(os.environ.get("DERLEM_CIKTI", KOK / ".derlem")).resolve()
PDF = CIKTI / "pdf"
PARALEL = int(os.environ.get("LO_PDF_PARALEL", "3"))
SINIR = int(os.environ.get("LO_PDF_SINIR_SN", "120"))
ortam = dict(os.environ, LC_ALL="C.UTF-8", LANG="C.UTF-8")

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
    return sonuc


tum = ozgunler + [c for _, _, c in isler]
sonuclar = {}
t0 = time.time()
with cf.ThreadPoolExecutor(max_workers=PARALEL) as ex:
    gelecek = {ex.submit(cevir, f, i % PARALEL): f for i, f in enumerate(tum)}
    for g in cf.as_completed(gelecek):
        r = g.result()
        sonuclar[r["yol"]] = r
        print(("HATA " if "hata" in r else "tamam ") + os.path.basename(r["yol"]), r.get("sayfa", ""), r.get("sn"), flush=True)

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
    satirlar.append(s)
(CIKTI / "rapor").mkdir(parents=True, exist_ok=True)
(CIKTI / "rapor" / "lo_pdf.json").write_text(json.dumps({"sureSn": round(time.time() - t0, 1), "satirlar": satirlar}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"{len(tum)} dosya PDF'e çevrildi ({round(time.time() - t0)} sn)")
