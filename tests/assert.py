#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dönüştürme çıktısını denetler: paras.json (motorun ürettiği paragraflar) +
out.docx (geçerli Word dosyası mı)."""
import json, sys, zipfile, re

paras = json.load(open('paras.json', encoding='utf-8'))
texts = [p['text'] for p in paras]
alltext = "\n".join(texts)

checks = []
def check(name, cond): checks.append((name, bool(cond)))

# 1) Satır sonu tireleri doğru birleşti (küçük harf devam)
check("tire: 'evrelerinde' birlesti",
      "evrelerinde" in alltext and "evre- lerinde" not in alltext and "evre lerinde" not in alltext)
check("tire: 'verememeye' birlesti",
      "verememeye" in alltext and "vereme- meye" not in alltext and "vereme meye" not in alltext)
check("tire: 'dogrudan' birlesti",
      "doğrudan" in alltext and "doğru- dan" not in alltext and "doğru dan" not in alltext)

# 2) Büyük harfle devam eden bileşikte tire KORUNDU (boşluksuz)
check("tire: 'Sovyet-Rusya' korundu",
      "Sovyet-Rusya" in alltext and "Sovyet- Rusya" not in alltext and "Sovyet Rusya" not in alltext)

# 3) Ayrı satırlar tek paragrafta birleşti (boşlukla)
LONG = "üretici güçlerin gelişme düzeyi ile üretim ilişkileri arasında zamanla biriken çelişkidir"
check("satir birlestirme (paragraf butunlugu)", any(LONG in t for t in texts))

# 3b) Girintili ilk satır paragraftan KOPMADI (regresyon koruması)
INTRO2 = "Sovyetler Birliği'nin tarih sahnesinden çekilişi, yirminci yüzyılın en tartışmalı olaylarından"
check("girintili ilk satir paragrafta kaldi", any(INTRO2 in t for t in texts))

# 4) Sayfa sınırını aşan paragraf TEK paragraf oldu
check("sayfa sinirini asan paragraf birlesti",
      any(("üretici güçlerin" in t and "gözler önüne serdi" in t) for t in texts))

# 5) Üstbilgi (her sayfada tekrar eden başlık) temizlendi
check("ustbilgi temizlendi",
      "PERİNÇEK" not in alltext and "STALİN'DEN" not in alltext)

# 6) Sayfa numaraları temizlendi (salt sayı paragrafı kalmadı)
check("sayfa numaralari temizlendi",
      not any(re.fullmatch(r"\d{1,4}", t.strip()) for t in texts))

# 7) Başlık algılandı (Heading olarak biçimlendi)
check("baslik algilandi",
      any(p.get('type') in ('h1', 'h2') and 'BÖLÜM' in p['text'] for p in paras))

# 8) Türkçe karakterler korundu
check("turkce korundu",
      all(w in alltext for w in ("çözülüş", "işlevsizleşmiştir", "değişen", "İşte", "yüzyılın")))

# 9) Kelimeler/cümleler bozulmadı (bütün cümleler yerinde)
check("cumleler butun",
      "Merkezî planlama" in alltext and "uzun erimli bir tarihsel süreçtir" in alltext)

# --- .docx geçerliliği ---
try:
    z = zipfile.ZipFile('out.docx')
    names = z.namelist()
    docxml = z.read('word/document.xml').decode('utf-8')
    check("docx: gerekli parcalar var",
          '[Content_Types].xml' in names and 'word/document.xml' in names and 'word/styles.xml' in names)
    check("docx: Heading stili gomuldu", ('Heading1' in docxml or 'Heading2' in docxml))
    check("docx: metin gomuldu", 'evrelerinde' in docxml and 'Sovyet-Rusya' in docxml)
    check("docx: gecersiz XML karakteri yok", True)  # parse edilebildiyse tamam
    import xml.dom.minidom as MD
    MD.parseString(docxml)  # XML gerçekten geçerli mi
    check("docx: document.xml gecerli XML", True)
except Exception as e:
    check("docx: acildi", False)
    print("docx HATASI:", repr(e))

print("\n--- DENETİMLER ---")
fails = 0
for n, c in checks:
    print(("  ✓  " if c else "  ✗  ") + n)
    if not c: fails += 1
print("\nSONUÇ:", "TÜM TESTLER GEÇTİ ✓✓✓" if fails == 0 else f"{fails} TEST BAŞARISIZ ✗")

print("\n--- İlk 6 paragraf (önizleme) ---")
for p in paras[:6]:
    print(f"[{p.get('type')}] {p['text'][:110]}{'...' if len(p['text'])>110 else ''}")

sys.exit(1 if fails else 0)
