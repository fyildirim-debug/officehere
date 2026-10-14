#!/usr/bin/env python3
"""Derlem üreticilerinin belgelere gömdüğü görselleri üretir (bizim içeriğimiz; Pillow ile çizilir).

- grafik.png : 640x400 sütun grafiği görüntüsü
- foto.jpg   : 1024x768 renk geçişi (fotoğraf yerine; JPEG yolunu sınar)
- logo.png   : 200x200 saydam zeminli daire (RGBA)

Kullanım: python3 scripts/derlem/uret_gorsel.py [çıktı_dizini]   (varsayılan: test-derlem/gorsel)
"""
import os
import sys

from PIL import Image, ImageDraw

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(KOK, "test-derlem", "gorsel")
os.makedirs(OUT, exist_ok=True)

# Sütun grafiği
g = Image.new("RGB", (640, 400), "white")
d = ImageDraw.Draw(g)
renkler = [(0x44, 0x72, 0xC4), (0xED, 0x7D, 0x31), (0xA5, 0xA5, 0xA5), (0xFF, 0xC0, 0x00), (0x5B, 0x9B, 0xD5)]
yukseklik = [120, 260, 180, 330, 220]
for i, (r, h) in enumerate(zip(renkler, yukseklik)):
    x = 60 + i * 110
    d.rectangle([x, 380 - h, x + 80, 380], fill=r)
d.line([40, 380, 620, 380], fill="black", width=3)
g.save(os.path.join(OUT, "grafik.png"), optimize=True)

# Renk geçişi (JPEG)
w, h = 1024, 768
f = Image.new("RGB", (w, h))
px = f.load()
for y in range(h):
    for x in range(w):
        px[x, y] = (x * 255 // w, y * 255 // h, (x * y * 128) // (w * h))
f.save(os.path.join(OUT, "foto.jpg"), quality=85)

# Saydam zeminli daire
lg = Image.new("RGBA", (200, 200), (0, 0, 0, 0))
ImageDraw.Draw(lg).ellipse([10, 10, 190, 190], fill=(0xE0, 0x5A, 0x5A, 255))
lg.save(os.path.join(OUT, "logo.png"), optimize=True)
print("görseller:", OUT, sorted(os.listdir(OUT)))
