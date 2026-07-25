#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Paragraf birleştirmeyi ZORLAYAN test PDF'i.

"Gladyo ve Ergenekon" çıktısında görülen bozulmayı üretir: sayfa geometrisi
paragraf sezgisellerini yanıltır (satır aralığı düzensiz, satır başları x
ekseninde oynak), satır sonlarında tireli bölünmeler vardır. Doğru davranış:
satırların TEK paragrafta birleşmesi ve tireli kelimelerin onarılması.
"""
import os
import random
import fitz

def _find_font(cands):
    for p in cands:
        if os.path.exists(p):
            return p
    raise SystemExit("Türkçe destekli serif font bulunamadı")

REG = _find_font([
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
])

W, H = 420.0, 640.0
ML, TOP, BOTTOM = 60.0, 90.0, 570.0
FS = 11.0

# Elle satırlanmış paragraflar; "-" ile biten satırlar kelime ortasından bölünmüş.
PARAGRAPHS = [
    [
        "riydi. Bu açıdan Doğan Öz, Muammer Aksoy, Uğur Mumcu su-",   # suikastlarıyla
        "ikastlarıyla aynı türdendi. Tayyip Erdoğan, suikasttan iki",
        "gün sonra MİT Müsteşarı, Emniyet Genel Müdürü'nü toplaya-",   # toplayarak
        "rak talimat verdi.",
    ],
    [
        "Bu talimat, ancak düzmece senaryolarla ve tertiplerle uygu-",  # uygulanabilirdi
        "lanabilirdi. Danıştay saldırısından hemen sonra Dışişleri",
        "Bakanı, Başbakan Yardımcısı sıfatıyla Emniyet Genel Müdür-",   # Müdürlüğü'nden
        "lüğü'nden ve MİT'ten brifing istedi.",
    ],
    [
        "Bizlerin Gladyo ile çarpışmaları, 1971 12 Mart darbesinin",
        "işkencehane ve hapishanelerinde başlar; bugünlere",          # noktalamasız kesik
        "gelir.",
    ],
    [
        "İkinci Dünya Savaşı'ndan sonra iki süper devlet arasındaki",
        "rekabetin odağı Avrupa idi. Yani dünyada üstünlük için mü-",  # mücadele
        "cadele, sonuç olarak Avrupa'yı denetlemek için yapılıyordu.",
    ],
]

doc = fitz.open()
page = doc.new_page(width=W, height=H)
y = TOP
rnd = random.Random(7)

for pi, para in enumerate(PARAGRAPHS):
    for li, line in enumerate(para):
        if y > BOTTOM:
            page = doc.new_page(width=W, height=H)
            y = TOP
        # Geometrik gürültü: satır başları oynak (yanlış "girinti" sinyali),
        # satır aralığı düzensiz (yanlış "paragraf boşluğu" sinyali).
        x = ML + (14.0 if li == 0 else rnd.choice([0.0, 9.0, 13.0]))
        page.insert_text((x, y), line, fontsize=FS, fontname="sr", fontfile=REG)
        y += rnd.choice([17.0, 24.0, 29.0])       # bazen 1.7x'ten büyük "boşluk"
    y += 8

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_zor.pdf")
doc.save(out)
print(f"Yazıldı: {out} · {doc.page_count} sayfa · {len(PARAGRAPHS)} paragraf bekleniyor")
