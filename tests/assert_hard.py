#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""test_zor.pdf çıktısının denetimi: satırlar tek paragrafta birleşmeli ve
satır sonu tireleri onarılmalı (Gladyo/Ergenekon çıktısındaki bozulma sınıfı)."""
import sys
import re
import docx

paras = [p.text for p in docx.Document(sys.argv[1]).paragraphs if p.text.strip()]
allt = " ".join(paras)

checks = [
    ("paragraf sayısı doğru (4)", len(paras) == 4),
    ("tire onarıldı: suikastlarıyla", "suikastlarıyla" in allt),
    ("tire onarıldı: toplayarak", "toplayarak" in allt),
    ("tire onarıldı: uygulanabilirdi", "uygulanabilirdi" in allt),
    ("tire onarıldı: Müdürlüğü'nden", "Müdürlüğü'nden" in allt),
    ("tire onarıldı: mücadele", "mücadele," in allt),
    ("noktalamasız kesik birleşti", "bugünlere gelir." in allt),
    ("kelime ortası boşluk kalmadı",
     not re.search(r"[a-zçğıöşü]- [a-zçğıöşü]", allt)),
    ("hiçbir paragraf küçük harfle başlamıyor (ilki hariç)",
     all(not re.match(r"^[a-zçğıöşü]", p.strip()) for p in paras[1:])),
]

print("--- ZOR TEST DENETİMLERİ ---")
fails = 0
for n, c in checks:
    print(("  ✓  " if c else "  ✗  ") + n)
    fails += (not c)
print("SONUÇ:", "GEÇTİ ✓" if not fails else f"{fails} BAŞARISIZ ✗")
if fails:
    print("\n--- üretilen paragraflar ---")
    for i, t in enumerate(paras):
        print(f"  [{i}] {t[:120]}")
sys.exit(1 if fails else 0)
