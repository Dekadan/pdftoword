#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gerçekçi bir kitap PDF'i üretir: her satır ayrı yerleştirilir (kullanıcının
yaşadığı 'her satır ayrı oturuyor' sorunu), tireli bölünmeler, her sayfada
tekrar eden üstbilgi + sayfa numarası, girintili paragraflar, bir başlık ve
sayfa sınırını aşan bir paragraf. Türkçe karakterler baştan sona kullanılır."""
import fitz, os

def _find_font(cands):
    for p in cands:
        if os.path.exists(p):
            return p
    raise SystemExit("Türkçe destekli serif font bulunamadı: " + ", ".join(cands))

REG = _find_font([
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
])
BOLD = _find_font([
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
])

W, H = 595.0, 842.0
ML = 78.0
TOP, BOTTOM = 118.0, 772.0
LHB, LHH = 24.0, 32.0
INDENT = 30.0
BODY_FS, HEAD_FS, HEADING_FS = 13.0, 9.0, 17.0
HEADER = "STALİN'DEN GORBAÇOV'A · DOĞU PERİNÇEK"

def wrap(text, width=64):
    """Metni ~width karakterlik satırlara böler (tire yok, sözcükler bütün)."""
    out, cur = [], ""
    for w in text.split():
        if cur and len(cur) + 1 + len(w) > width:
            out.append(cur); cur = w
        else:
            cur = (cur + " " + w) if cur else w
    if cur:
        out.append(cur)
    return out

HEADING = "BİRİNCİ BÖLÜM — ÇÖZÜLÜŞÜN KÖKLERİ"

INTRO = ("Sovyetler Birliği'nin tarih sahnesinden çekilişi, yirminci yüzyılın en "
    "tartışmalı olaylarından biri olarak kabul edilir. Bu büyük altüst oluşun "
    "köklerini anlamak isteyen bir okur, olguları yüzeysel gözlemlerin ötesine "
    "geçerek incelemek zorundadır. Ekonomik yapı, siyasal önderlik ve toplumsal "
    "bilinç arasındaki karmaşık ilişki, çözülüş sürecinin yönünü belirleyen temel "
    "etmenlerin başında gelir. Tarihsel maddecilik, bu ilişkiyi kavramak için "
    "elimizdeki en güçlü yöntemdir. İşte bu bölümde, adım adım bu etmenleri ele "
    "alacağız ve her birini kendi bağlamı içinde değerlendireceğiz.")

# Sayfa sınırını aşacak, tireli bölünmeler içeren TANINIR paragraf (elle satırlanmış)
CROSS = [
    "Çözülüşün ardında yatan en önemli etmenlerden biri, üretici",
    "güçlerin gelişme düzeyi ile üretim ilişkileri arasında zamanla",
    "biriken çelişkidir. Merkezî planlama, sanayileşmenin ilk evre-",   # evrelerinde
    "lerinde büyük başarılar sağlamış olsa da, ilerleyen yıllarda",
    "toplumun değişen gereksinimlerine yeterince yanıt vereme-",        # verememeye
    "meye başlamıştır. Bu tıkanıklık, yalnızca ekonomik bir sorun",
    "olmakla kalmamış, siyasal önderliğin niteliğiyle de doğru-",       # doğrudan
    "dan bağlantılı hâle gelmiştir. Önderlik kademelerindeki",
    "bürokratikleşme, halkın inisiyatifini giderek zayıflatmış;",
    "böylece toplumsal denetim düzenekleri işlevsizleşmiştir. Bu",
    "koşullarda yeniden alevlenen Sovyet-",                             # Sovyet-Rusya (tire korunur)
    "Rusya tartışmaları, sorunun derinliğini gözler önüne serdi.",
    "Kısacası çözülüş, tek bir nedene değil, iç içe geçmiş birçok",
    "etmenin bütününe dayanan uzun erimli bir tarihsel süreçtir.",
]

FILL_A = ("Bu noktada bir uyarı gereklidir: karmaşık bir tarihsel olguyu tek bir "
    "cümleyle açıklamaya kalkışmak, çoğu zaman yanıltıcı sonuçlara götürür. "
    "Olaylar zincirini oluşturan halkaları tek tek görmeden bütünü kavramak "
    "olanaksızdır. Bu yüzden ilerleyen sayfalarda ekonomik göstergeleri, siyasal "
    "kararları ve toplumsal tepkileri ayrı ayrı ele alacak, sonra bunları ortak "
    "bir çerçevede birleştireceğiz. Yöntemimiz, olguları konuşturmak olacaktır.")

FILL_B = ("Ekonomik göstergeler, uzun bir durgunluk döneminin sinyallerini çok "
    "önceden vermeye başlamıştı. Verimlilik artışı yavaşlamış, teknolojik "
    "yenilenme sekteye uğramış ve tüketim mallarındaki darlık gündelik yaşamı "
    "doğrudan etkilemişti. Bu tablo, halkın sisteme olan güvenini adım adım "
    "aşındırdı. Güven kaybı ise, hiçbir istatistiğe tam olarak yansımayan ama "
    "sonuçları en ağır olan etmenlerden biriydi.")

FILL_C = ("Siyasal önderlik ise bu sorunların üstesinden gelecek araçlardan yoksundu. "
    "Karar süreçleri ağırlaşmış, sorumluluk almaktan kaçınan bir bürokrasi "
    "kademelenmişti. Reform girişimleri çoğu zaman yarım kaldı; kimi zaman da "
    "amaçlananın tam tersi sonuçlar doğurdu. Böyle bir ortamda, köklü bir "
    "dönüşümün sancısız gerçekleşmesini beklemek gerçekçi değildi.")

FILL_D = ("Toplumsal bilinç düzeyinde ise çelişkili eğilimler bir arada yaşıyordu. "
    "Bir yanda eşitlikçi değerlere bağlılık sürüyor, öte yanda tüketim "
    "toplumunun cazibesi genç kuşakları etkiliyordu. Bu gerilim, dönüşüm "
    "sürecinin yönünü belirleyen sessiz ama güçlü bir etmen olarak işledi. "
    "Sonraki bölümlerde bu bilinç değişiminin izlerini ayrıntısıyla süreceğiz.")

PRE = ("Yöntem sorununa açıklık getirmeden ilerlemek, bizi daha başında yanlış "
    "yollara sürükleyebilir. Bir olguyu değerlendirirken, onu çevreleyen koşulları "
    "ve zaman içindeki hareketini birlikte görmek gerekir. Durağan bir fotoğraf "
    "yerine, sürecin filmini izlemek; nedenleri sonuçlardan, biçimi özden ayırmadan "
    "bütünü kavramak zorundayız. Aşağıdaki çözümleme boyunca bu ilkeye sadık "
    "kalacağız ve her adımda olguların kendisine başvuracağız.")

# --- Blok listesi --- (CROSS'u sayfa sınırına denk getirmek için önüne dolgu,
#     4 sayfaya ulaşmak için ardına tekrarlı dolgu)
blocks = [("h", HEADING)]
blocks.append(("p", wrap(INTRO)))
blocks.append(("p", wrap(PRE)))
blocks.append(("p", CROSS))
for f in (FILL_A, FILL_B, FILL_C, FILL_D, FILL_A, FILL_B, FILL_C, FILL_D):
    blocks.append(("p", wrap(f)))

# --- Yerleşim ---
doc = fitz.open()
pages = []
def new_page():
    p = doc.new_page(width=W, height=H); pages.append(p); return p
cur = new_page(); y = TOP
line_pages = {}   # (blok_idx, satır_idx) -> sayfa no

def put(text, x, fs, bold=False):
    cur.insert_text((x, y), text, fontsize=fs,
                    fontname=("sb" if bold else "sr"),
                    fontfile=(BOLD if bold else REG))

for bi, (kind, payload) in enumerate(blocks):
    if kind == "h":
        if y + LHH > BOTTOM:
            cur = new_page(); y = TOP
        y += 6
        put(payload, ML, HEADING_FS, bold=True)
        y += LHH
    else:
        for li, line in enumerate(payload):
            if y + LHB > BOTTOM:
                cur = new_page(); y = TOP
            x = ML + (INDENT if li == 0 else 0.0)
            put(line, x, BODY_FS)
            line_pages[(bi, li)] = len(pages)
            y += LHB
        y += 7  # paragraf arası boşluk

# --- Üstbilgi + sayfa numarası (her sayfada) --- taze sayfa referansı kullan
for idx in range(1, doc.page_count + 1):
    p = doc[idx - 1]
    p.insert_text((ML, 74.0), HEADER, fontsize=HEAD_FS, fontname="sr", fontfile=REG)
    num = str(idx)
    p.insert_text((W/2 - 4, 806.0), num, fontsize=11.0, fontname="sr", fontfile=REG)

out = "test_kitap.pdf"
doc.save(out)
print(f"Yazıldı: {out} · {doc.page_count} sayfa")

# CROSS paragrafının sayfa sınırını gerçekten aştığını doğrula
cross_idx = 3
cp = sorted({line_pages[(cross_idx, i)] for i in range(len(CROSS))})
print("CROSS paragrafının bulunduğu sayfalar:", cp, "-> sınır aşıldı mı:", len(cp) > 1)

# fitz ile çıkarımı doğrula (Türkçe karakter + üstbilgi/altbilgi var mı)
d = fitz.open(out)
t0 = d[0].get_text()
print("Sayfa 1 üstbilgi görülüyor mu:", "PERİNÇEK" in t0)
print("Türkçe örnek ('çözülüş') sayfa metninde:", "çözülüş" in d[0].get_text() or "çözülüş" in d[1].get_text())
