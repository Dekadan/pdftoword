#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf2word.py — Kitap PDF'lerini "baskı öncesi" düzeninde, düzenlenebilir Word
(.docx) dosyasına çevirir.

Neler yapar:
  * Satırları doğru paragraflara birleştirir (sayfa sınırlarını da aşarak).
  * Satır sonu tirelerini (yumuşak tire dahil) doğru birleştirir; büyük harfle
    başlayan bileşiklerde tireyi korur (Sovyet-Rusya).
  * Apostrof sonrası bölünen ekleri onarır (Türkiye' nin -> Türkiye'nin).
  * DİPNOTLAR: gövdedeki üst simge rakamları ve sayfa altındaki küçük puntolu
    numaralı notları eşleştirip GERÇEK Word dipnotu yapar (sayfa altında,
    otomatik numaralı). Eşleşmeyenler kaybolmaz; küçük italik paragraf kalır.
  * İÇİNDEKİLER: PDF'teki içindekiler sayfalarını tanır, yerine Word'ün CANLI
    İçindekiler alanını koyar (başlıklardan üretilir; sayfa numaralarını Word
    kendisi doğru basar, belge açılınca güncellenir).
  * Sayfa altbilgisine SAYFA NUMARASI alanı ekler; kâğıt A4, kitap marjları.
  * Üstbilgi/altbilgi tekrarlarını ve salt sayfa numarası satırlarını atar.
  * Başlıkları tanır (punto + "V.2." / "BÖLÜM" / BÜYÜK HARF desenleri) ve
    Word Heading 1-3 stilleriyle işaretler (canlı içindekiler bunlardan beslenir).
  * Kenar notlarını gövdeyi bölmeden ayırır; bozuk metin katmanlı sayfaları raporlar.

Kelimeleri/cümleleri DEĞİŞTİRMEZ; yalnızca doğru yere ve doğru yapıya koyar.

Kurulum:  pip install pymupdf python-docx
Kullanım: python pdf2word.py "kitap.pdf" [cikti.docx] [--no-footnotes --no-toc ...]
          python pdf2word.py klasor/           (toplu çevirim)
"""
import argparse
import os
import re
import statistics
import sys
import zipfile
from xml.sax.saxutils import escape as xesc

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF gerekli:  pip install pymupdf")
try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx gerekli:  pip install python-docx")

SOFT = "­"                     # yumuşak tire (U+00AD)
LOWER = "a-zçğıöşü"
UPPER = "A-ZÇĞİÖŞÜ"
LETTER = LOWER + UPPER
VOWELS = set("aeıioöuüAEIİOÖUÜ")

RE_LETTER_HYPHEN = re.compile("[" + LETTER + "]-$")
RE_STARTS_LOWER = re.compile("^[" + LOWER + "]")
RE_PAGENUM = re.compile(r"^[\s.\-–—·]*\d{1,4}[\s.\-–—·]*$")
RE_ROMAN = re.compile(r"^[\s.\-–—·]*[ivxlcdmIVXLCDM]{1,7}[\s.\-–—·]*$")
RE_STRONG_END = re.compile(r"[.!?…][\"»”’')\]]?$")
RE_MIDCAP = re.compile("[" + LOWER + "][" + UPPER + "]")
RE_LETDIG = re.compile("[" + LETTER + r"]\d|\d[" + LETTER + "]")
RARE_CHARS = set("ąăĄ*=>~^`|ª•■□▪")

FN_TOKEN = "⟦FN%d⟧"                    # ⟦FN7⟧ — dipnot yer tutucusu
RE_FN_TOKEN = re.compile("⟦FN(\\d+)⟧")
TOC_TOKEN = "⟦TOCFIELD⟧"               # içindekiler alanı yer tutucusu
# Dipnot metni başı: "12 Doğu Perinçek…" ya da boşluksuz "12Doğu Perinçek…"
RE_NOTE_START = re.compile(r"^(\d{1,3})[.)]?(?:\s+(\S.*)|([" + UPPER + r"\"«“].*))$")
RE_TRAIL_NUM = re.compile(r"\s(\d{1,3})\s*$")     # satır sonuna düşmüş sonraki not numarası
RE_TOC_ENTRY = re.compile(r".{3,}\s\d{1,4}\s*$")
RE_HEAD_NUM = re.compile(r"^((?:[IVXLCDM]+|\d+)(?:\.(?:\d+|[IVXLCDM]+))*)[.)]?\s+\S")
RE_CHAPTER_WORD = re.compile(
    r"(BÖLÜM|GİRİŞ|GIRIŞ|SONUÇ|ÖNSÖZ|SUNUŞ|ÖNDEYİŞ|KAYNAKÇA|DİZİN|EKLER?\b|İÇİNDEKİLER)")

# Apostrof sonrası bölünmüş Türkçe ekler (yalnızca bunlar birleştirilir)
_SUFFIXES = """
nin nın nun nün in ın un ün e a ye ya na ne i ı u ü yi yı yu yü
de da te ta nde nda den dan ten tan nden ndan le la yle yla
deki daki ndeki ndaki teki taki li lı lu lü lik lık luk lük
ler lar leri ları lere lara lerde larda lerden lardan lerin ların
si sı su sü dir dır dur dür tir tır tur tür
""".split()
SUFFIX_SET = set(s for s in _SUFFIXES if s.isalpha())


# ---------------------------------------------------------------- yardımcılar
def median(a):
    return statistics.median(a) if a else 0


def mode_round(vals, step):
    if not vals:
        return 0
    from collections import Counter
    c = Counter(round(v / step) * step for v in vals)
    return c.most_common(1)[0][0]


def percentile(a, p):
    if not a:
        return 0
    b = sorted(a)
    return b[min(len(b) - 1, max(0, round((len(b) - 1) * p)))]


def word_count(s):
    return len(s.split())


def _token_weird(w):
    core = w.strip(".,;:!?()[]{}\"'«»…-/")
    letters = [c for c in core if c.isalpha()]
    if not letters:
        return False
    if all(c.isupper() for c in letters):        # akronim (ABD, FETÖ) -> normal
        return False
    if RE_MIDCAP.search(core):                    # kelime ortası küçük->BÜYÜK
        return True
    if len(letters) >= 5 and not any(c in VOWELS for c in core):
        return True
    return False


def _token_garbage(w):
    return _token_weird(w) or bool(RE_LETDIG.search(w)) or any(c in RARE_CHARS for c in w)


def garbage_line(line):
    toks = line.split()
    if len(toks) < 3:
        return False
    bad = sum(1 for w in toks if _token_garbage(w))
    return bad / len(toks) >= 0.30


def ocr_junk(text):
    """OCR'lanan kapak/görsel sayfalarından gelen anlamsız kırıntı satırı
    ("|", "kz", "~-" gibi). Yalnızca OCR'lı satırlara uygulanır."""
    s = text.strip()
    if not s:
        return True
    for tok in s.split():
        core = "".join(c for c in tok if c.isalpha())
        if len(core) >= 2 and any(c in VOWELS for c in core):
            return False          # sesli harfli, gerçek kelimeye benzer parça var
        if tok.isdigit():
            return False          # sayfa numarası vb. — burada eleme
    return True


def weird_line(text):
    s = text.strip()
    if not s:
        return True
    if any(_token_weird(w) for w in s.split()):
        return True
    nonspace = [c for c in s if not c.isspace()]
    alpha = sum(1 for c in nonspace if c.isalpha())
    return alpha / max(1, len(nonspace)) < 0.7


# ============================== 0) OCR desteği ==============================
def find_tessdata():
    """Türkçe OCR dil dosyasını (tur.traineddata) içeren tessdata klasörünü bul:
    paketlenmiş uygulama içi -> ortam değişkeni -> sistem yolları."""
    import glob
    cands = []
    if getattr(sys, "_MEIPASS", None):                       # PyInstaller onefile
        cands.append(os.path.join(sys._MEIPASS, "tessdata"))
    cands.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "tessdata"))
    cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "tessdata"))
    env = os.environ.get("TESSDATA_PREFIX")
    if env:
        cands += [env, os.path.join(env, "tessdata")]
    cands += glob.glob("/usr/share/tesseract-ocr/*/tessdata")
    cands += ["/usr/share/tessdata", "/usr/local/share/tessdata",
              "/opt/homebrew/share/tessdata",
              r"C:\Program Files\Tesseract-OCR\tessdata"]
    for c in cands:
        if c and os.path.exists(os.path.join(c, "tur.traineddata")):
            return c
    return None


def choose_ocr_pages(doc, mode):
    """OCR uygulanacak sayfaları seç.
    auto: metin katmanı bozuk (çöp) sayfalar + metinsiz taranmış sayfalar.
    full: tüm sayfalar.  off: hiçbiri.
    Döner: (sayfa_indeksleri, çöp_sayfa_numaraları)"""
    garbage = []
    scanned = []
    for pno in range(doc.page_count):
        t = doc[pno].get_text()
        lines = [l for l in t.split("\n") if l.strip()]
        g = sum(1 for l in lines if garbage_line(l))
        if g >= 2:
            garbage.append(pno)
        if len(t.strip()) < 25 and doc[pno].get_images():
            scanned.append(pno)
    if mode == "off":
        return set(), garbage
    if mode == "full":
        return set(range(doc.page_count)), garbage
    # auto: taranmış kitap sezgisi — sayfaların çoğu metinsiz görüntüyse hepsini OCR'la
    if len(scanned) >= max(3, doc.page_count * 0.6):
        return set(range(doc.page_count)), garbage
    return set(garbage) | set(scanned), garbage


# ============================== 1) satır + dipnot işareti çıkarımı ==========
def extract_pages(doc, ocr_pages=None, tessdata=None, on_page=None):
    """Sayfa sayfa satırlar; gövdedeki üst simge rakamlar ⟦FNk⟧ olarak işaretlenir.
    ocr_pages içindeki sayfaların metni OCR ile (görüntüden) okunur.
    on_page(okunan, toplam, ocr_mu) çağrılırsa ilerleme bildirilir."""
    ocr_pages = ocr_pages or set()
    ocr_done = []
    pages = []
    markers = []          # k sırayla: {"k", "page", "num"}
    for pno in range(doc.page_count):
        if on_page:
            on_page(pno, doc.page_count, pno in ocr_pages and bool(tessdata))
        page = doc[pno]
        W, H = page.rect.width, page.rect.height
        d = None
        if pno in ocr_pages and tessdata:
            try:
                tp = page.get_textpage_ocr(language="tur", dpi=300, full=True,
                                           tessdata=tessdata)
                d = page.get_text("dict", textpage=tp)
                ocr_done.append(pno)
            except Exception:
                d = None                      # OCR başarısızsa normal çıkarıma düş
        if d is None:
            d = page.get_text("dict")
        lines = []
        for block in d["blocks"]:
            if block.get("type", 0) != 0:
                continue
            for ln in block["lines"]:
                spans = [s for s in ln["spans"] if s["text"]]
                if not spans:
                    continue
                nondigit = [s["size"] for s in spans
                            if s["text"].strip() and not s["text"].strip().isdigit()]
                main_size = median(nondigit) if nondigit else median(
                    [s["size"] for s in spans if s["text"].strip()])

                parts = []
                sizes = []
                prev_end = None
                pend_num = None      # bitişik üst simge rakamları birleştir (1+0 -> 10)

                def flush_pend():
                    nonlocal pend_num
                    if pend_num is not None:
                        k = len(markers)
                        markers.append({"k": k, "page": pno, "num": pend_num})
                        parts.append(FN_TOKEN % k)
                        pend_num = None

                for s in spans:
                    t = s["text"]
                    st = t.strip()
                    x0s, _, x1s, _ = s["bbox"]
                    is_sup = (st.isdigit() and 1 <= len(st) <= 3 and main_size and
                              ((s["flags"] & 1) or s["size"] <= main_size * 0.82) and
                              len(spans) > 1)
                    if is_sup:
                        pend_num = (pend_num or "") + st
                        prev_end = x1s
                        continue
                    # İşaretin devam rakamları sonraki span'e yapışabiliyor
                    # ("başlar." + ¹ + "0 Dosyada" -> işaret 10). Baştaki rakamları çek.
                    if pend_num is not None:
                        mdig = re.match(r"(\d{1,2})(?!\d)", t)
                        if mdig and len(pend_num) + len(mdig.group(1)) <= 3:
                            pend_num += mdig.group(1)
                            t = t[mdig.end():]
                            st = t.strip()
                            if not st and not t:
                                prev_end = x1s
                                continue
                    flush_pend()
                    if prev_end is not None:
                        gap = x0s - prev_end
                        if gap > s["size"] * 0.25 and parts and \
                                not parts[-1].endswith(" ") and not t.startswith(" "):
                            parts.append(" ")
                    parts.append(t)
                    if st:
                        sizes.append(s["size"])
                    prev_end = x1s
                flush_pend()

                text = re.sub(r"[ \t ]+", " ", "".join(parts)).strip()
                if not text:
                    continue
                # Kalin yazi orani (baslik taniminda kullanilir; flags bit 4 = bold)
                nchar = sum(len(s["text"].strip()) for s in spans if s["text"].strip())
                nbold = sum(len(s["text"].strip()) for s in spans
                            if s["text"].strip() and (s["flags"] & 16))
                x0, y0, x1, y1 = ln["bbox"]
                lines.append({
                    "text": text, "x0": x0, "x1": x1, "y0": y0, "y1": y1,
                    "size": median(sizes) if sizes else main_size,
                    "page": pno, "W": W, "H": H,
                    "ocr": pno in ocr_pages and bool(tessdata),
                    "bold": nchar > 0 and nbold / nchar >= 0.8,
                })
        # Blok sırası PDF içeriğinde görsel sırayla aynı olmak zorunda değil;
        # satırları okuma sırasına (üstten alta, soldan sağa) diz.
        lines.sort(key=lambda l: (round(l["y0"], 1), l["x0"]))
        if pno in ocr_pages and tessdata:      # kapak/görsel OCR kırıntılarını at
            lines = [l for l in lines if not ocr_junk(l["text"])]
        pages.append({"W": W, "H": H, "lines": lines})
    return pages, markers, ocr_done


# ============================== 2) dipnot metinleri =========================
def collect_footnotes(pages, body_size, dehyphen):
    """Sayfa altındaki küçük puntolu numaralı notları toplar.
    Dönen: notes[page][num] = {"text", "lines": [line objesi...]}"""
    notes = {}
    for pi, pg in enumerate(pages):
        cur = None
        page_notes = {}
        for ln in sorted(pg["lines"], key=lambda l: (l["y0"], l["x0"])):
            small = ln["size"] <= body_size * 0.90
            bottom = ln["y0"] > pg["H"] * 0.50
            m = RE_NOTE_START.match(ln["text"]) if (small and bottom) else None
            if m and (cur is None or ln["y0"] >= cur["last_y"]):
                num = m.group(1)
                cur = {"num": num, "text": (m.group(2) or m.group(3) or "").strip(),
                       "lines": [ln], "last_y": ln["y0"]}
                page_notes.setdefault(num, cur)
            elif cur is not None and small and bottom and ln["y0"] >= cur["last_y"]:
                # Bir sonraki notun numarası, bu notun son satırının SONUNA
                # düşmüş olabilir ("… s.273. 16" + sonraki satır "Adnan Akfırat…").
                m2 = RE_TRAIL_NUM.search(cur["text"])
                starts_new = (m2 and int(m2.group(1)) == int(cur["num"]) + 1
                              and re.match("^[" + UPPER + r"\"«“]", ln["text"].strip()))
                if starts_new:
                    nxt = m2.group(1)
                    cur["text"] = cur["text"][:m2.start()].rstrip()
                    cur = {"num": nxt, "text": ln["text"].strip(),
                           "lines": [ln], "last_y": ln["y0"]}
                    page_notes.setdefault(nxt, cur)
                else:
                    _join(cur, ln["text"], dehyphen)
                    cur["lines"].append(ln)
                    cur["last_y"] = ln["y0"]
            else:
                cur = None
        if page_notes:
            notes[pi] = page_notes
    return notes


RE_OPEN_END = re.compile(r"[,;:(\[«“„\-–—]$")     # açık uçlu bitiş (virgül, tire…)


def all_caps(s):
    letters = [c for c in s if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)


def is_continuation(prev_text, next_text):
    """Önceki satır cümle ortasında bitip bu satır küçük harfle başlıyorsa,
    bu bir paragraf devamıdır — sayfa geometrisi (girinti/satır aralığı) ne
    derse desin. Kitaplarda gerçek bir paragraf küçük harfle başlamaz."""
    a = (prev_text or "").strip()
    b = (next_text or "").strip()
    if not a or not b:
        return False
    b0 = RE_FN_TOKEN.sub("", b).lstrip()          # dipnot işaretini atla
    if not b0 or not RE_STARTS_LOWER.match(b0):
        return False
    if a.endswith(SOFT) or RE_LETTER_HYPHEN.search(a):   # kelime ortasından bölünmüş
        return True
    if RE_STRONG_END.search(a):                          # cümle gerçekten bitmiş
        return False
    return True                                          # noktalamasız kesilmiş


def _join(par, add, dehyphen):
    add = add.strip()
    if not add:
        return
    prev = par["text"].rstrip()
    if prev.endswith(SOFT):
        par["text"] = prev[:-1] + add
    elif RE_LETTER_HYPHEN.search(prev):
        if dehyphen and RE_STARTS_LOWER.match(add):
            par["text"] = prev[:-1] + add
        else:
            par["text"] = prev + add
    else:
        par["text"] = (prev + " " + add) if prev else add


RE_GLUE = re.compile(r"([.!?…\"»”')\]])(\d{1,3})(?=\s|$)")


def match_footnotes(pages, markers, notes, body_size):
    """İşaret <-> not eşleşmesi (iki geçiş):
    1) Üst simge/küçük punto span işaretleri.
    2) Gövde metnine yapışık işaretler: sahipsiz notun numarası, aynı sayfada
       noktalamaya bitişik geçiyorsa (başlar.10 / dedi."13) işarete çevrilir.
    Eşleşen not satırları gövdeden çıkar; eşleşmeyen işaret düz rakama döner;
    eşleşmeyen not küçük paragraf olarak kalır (hiçbir şey kaybolmaz)."""
    fn_list = []                     # [(id, text)]
    remove = set()
    remap = {}

    def claim(num, page):
        for p in (page, page + 1):
            note = notes.get(p, {}).get(num)
            if note is not None:
                notes[p][num] = None
                return note
        return None

    for mk in markers:
        note = claim(mk["num"], mk["page"])
        if note is None:
            remap[mk["k"]] = None
        else:
            fid = len(fn_list) + 1
            fn_list.append((fid, note["text"]))
            remap[mk["k"]] = fid
            for ln in note["lines"]:
                remove.add(id(ln))

    def fix_text(t):
        def sub(m):
            k = int(m.group(1))
            new = remap.get(k)
            if new is None:
                for mk in markers:
                    if mk["k"] == k:
                        return mk["num"]
                return ""
            return FN_TOKEN % new
        return RE_FN_TOKEN.sub(sub, t)

    for pg in pages:
        for ln in pg["lines"]:
            ln["text"] = fix_text(ln["text"])

    # --- 2. geçiş: metne yapışık işaretler ---
    for p in sorted(notes):
        for num, note in list(notes[p].items()):
            if note is None:
                continue
            placed = False
            for tp in (p, p - 1):
                if placed or not (0 <= tp < len(pages)):
                    continue
                for ln in pages[tp]["lines"]:
                    if id(ln) in remove or ln["size"] <= body_size * 0.90:
                        continue      # not satırlarının kendisine dokunma
                    hit = {"done": False}

                    def sub(m):
                        if hit["done"] or m.group(2) != num:
                            return m.group(0)
                        hit["done"] = True
                        fid = len(fn_list) + 1
                        fn_list.append((fid, note["text"]))
                        return m.group(1) + (FN_TOKEN % fid)

                    new_text = RE_GLUE.sub(sub, ln["text"])
                    if hit["done"]:
                        ln["text"] = new_text
                        notes[p][num] = None
                        for l2 in note["lines"]:
                            remove.add(id(l2))
                        placed = True
                        break

    for pg in pages:
        pg["lines"] = [ln for ln in pg["lines"] if id(ln) not in remove]
    return fn_list


# ============================== 3) içindekiler ==============================
def _tr_fold(s):
    """Türkçe güvenli küçük harf karşılaştırma (İ -> i̇ birleşik nokta sorunu)."""
    return s.casefold().replace("̇", "")


def replace_toc(pages, body_size):
    """Kitabın başındaki içindekiler sayfalarını bulur; satırlarını atıp yerine
    tek bir ⟦TOCFIELD⟧ satırı koyar (Word'de canlı alana dönüşür).
    İki düzeni tanır: girdi + sayfa no AYNI satırda ya da AYRI satırlarda."""
    n = len(pages)
    limit = max(12, int(n * 0.25))
    cand = set()
    titled = set()
    for pi in range(min(limit, n)):
        lines = pages[pi]["lines"]
        if any(_tr_fold(l["text"]).strip() == "içindekiler" for l in lines):
            titled.add(pi)
        if len(lines) < 5:
            continue
        inline = sum(1 for l in lines
                     if word_count(l["text"]) >= 2 and RE_TOC_ENTRY.search(l["text"]))
        numonly = sum(1 for l in lines if RE_PAGENUM.match(l["text"]))
        if (inline >= 6 and inline / len(lines) >= 0.55) or \
           (numonly >= 5 and numonly / len(lines) >= 0.25):
            cand.add(pi)
    cand |= titled
    if not cand:
        return False
    # başlıklı sayfadan başlayan bitişik grubu al; başlık yoksa >=2 sayfalık grup iste
    groups = []
    for pi in sorted(cand):
        if groups and pi == groups[-1][-1] + 1:
            groups[-1].append(pi)
        else:
            groups.append([pi])
    chosen = None
    for g in groups:
        if any(pi in titled for pi in g):
            chosen = g
            break
    if chosen is None:
        chosen = next((g for g in groups if len(g) >= 2), None)
    if not chosen:
        return False
    first = chosen[0]
    for pi in chosen:
        pages[pi]["lines"] = []
    pages[first]["lines"] = [{
        "text": TOC_TOKEN, "x0": 0, "x1": 100, "y0": 0, "y1": 10,
        "size": body_size, "page": first, "W": pages[first]["W"], "H": pages[first]["H"],
    }]
    return True


# ============================== 4) üstbilgi/altbilgi ========================
def norm_hf(s):
    return re.sub(r"\s+", " ", re.sub(r"\d+", "#", s.lower())).strip()


def strip_headers_footers(pages):
    n = len(pages)
    if n < 3:
        return
    from collections import Counter
    top, bot = Counter(), Counter()
    for pg in pages:
        lines = pg["lines"]
        if not lines:
            continue
        if word_count(lines[0]["text"]) <= 8:
            top[norm_hf(lines[0]["text"])] += 1
        if word_count(lines[-1]["text"]) <= 8:
            bot[norm_hf(lines[-1]["text"])] += 1
    rep = max(3, round(n * 0.3))

    def is_num(s):
        return bool(RE_PAGENUM.match(s) or RE_ROMAN.match(s.strip()))

    for pg in pages:
        lines = pg["lines"]
        if len(lines) <= 1:
            if lines and RE_PAGENUM.match(lines[0]["text"]):
                pg["lines"] = []
            continue
        a, b = 0, len(lines)
        t, bt = lines[0], lines[-1]
        if is_num(t["text"]) or (word_count(t["text"]) <= 8 and top[norm_hf(t["text"])] >= rep):
            a = 1
        if b - a > 0 and (is_num(bt["text"]) or
                          (word_count(bt["text"]) <= 8 and bot[norm_hf(bt["text"])] >= rep)):
            b -= 1
        pg["lines"] = lines[a:b]


# ============================== 5) başlık tanıma ============================
def heading_level(line, body_size):
    t = RE_FN_TOKEN.sub("", line["text"]).strip()
    if not t or t == TOC_TOKEN or word_count(t) > 14 or weird_line(t):
        return 0
    letters = [c for c in t if c.isalpha()]
    allcaps = letters and all(c.isupper() for c in letters)
    if allcaps and word_count(t) <= 8:
        if RE_CHAPTER_WORD.search(t) or line["size"] >= body_size * 1.25:
            return 1
        if line["size"] >= body_size * 1.02:
            return 2
        return 0
    if not (t[0].isupper() or t[0].isdigit()):
        return 0
    if line["size"] >= body_size * 1.6:
        return 1
    if line["size"] >= body_size * 1.35:
        return 2
    m = RE_HEAD_NUM.match(t)
    if m and line["size"] >= body_size * 1.02 and word_count(t) <= 12:
        depth = m.group(1).count(".")
        return 2 if depth == 0 else 3
    return 0


# ============================== 6) paragraf birleştirme =====================
def assemble(pages, body_size, opts, scanned=False):
    """scanned=True: kitabın TAMAMI taranmış (metin OCR'dan geliyor). Bu durumda
    punto ölçüleri oynak olduğu için küçük-punto ayrımına (italik kenar notu)
    güvenilmez — her şey düz gövde metni olur; başlıklar ise OCR'dan gelmek
    zorunda olduğu için başlık tanıma AÇIK kalır."""
    flat = [ln for pg in pages for ln in pg["lines"]]
    if not flat:
        return []

    body_lines = [l for l in flat if abs(l["size"] - body_size) <= 0.6]
    ref = body_lines if len(body_lines) >= 5 else flat
    body_left = mode_round([l["x0"] for l in ref], 3)
    right_edge = percentile([l["x1"] for l in ref], 0.9)
    body_width = max(1.0, right_edge - body_left)

    gaps = []
    for pg in pages:
        body = [l for l in pg["lines"] if abs(l["size"] - body_size) <= 0.6]
        for i in range(1, len(body)):
            g = body[i]["y0"] - body[i - 1]["y0"]
            if 0 < g < body_size * 4:
                gaps.append(g)
    line_gap = median(gaps) or body_size * 1.4
    indent_thr = body_left + max(6, body_size * 0.9)
    short_if = right_edge - body_size * 3.5

    def is_margin_note(ln):
        if not opts.asides:
            return False
        narrow = (ln["x1"] - ln["x0"]) < 0.55 * body_width
        offset = ln["x0"] > body_left + 3.0 * body_size
        return narrow and offset and ln["size"] <= body_size + 0.5

    def is_small(ln):
        if scanned:              # OCR punto oynaklığı: düz metin, tek punto
            return False
        return opts.asides and body_size and ln["size"] <= body_size * 0.86

    first_page = min((ln["page"] for ln in flat), default=0)

    def front_matter(ln):        # kapak / iç kapak / künye bölgesi
        return ln["page"] <= first_page + 3

    paras = []
    cur = None
    prev = None
    prev_heading = False
    cur_aside = None
    pending = []

    def flush_aside():
        nonlocal cur_aside
        if cur_aside and cur_aside["text"].strip():
            pending.append(cur_aside)
        cur_aside = None

    def push_body():
        nonlocal cur
        if cur and cur["text"].strip():
            paras.append(cur)
        cur = None
        paras.extend(pending)
        pending.clear()

    for ln in flat:
        if ln["text"] == TOC_TOKEN:
            flush_aside()
            push_body()
            paras.append({"type": "toc", "text": TOC_TOKEN})
            prev = ln
            prev_heading = True
            continue

        # Kitabın tümü taranmışsa başlıklar zaten OCR'dan gelmek zorunda: tanıma açık.
        # Yalnızca "birkaç bozuk sayfa onarıldı" durumunda (kupür/şema) kapatılır ki
        # gazete manşetleri başlık/içindekiler'e sızmasın.
        allow_heading = opts.headings and (
            scanned or not ln.get("ocr") or getattr(opts, "ocr", "auto") == "full")
        hl = heading_level(ln, body_size) if allow_heading else 0
        is_heading = hl > 0

        if not is_heading and cur is not None and is_margin_note(ln):
            if cur_aside is None:
                cur_aside = {"type": "aside", "text": ln["text"].strip()}
            else:
                _join(cur_aside, ln["text"], opts.dehyphen)
            continue
        flush_aside()

        is_sm = (not is_heading) and is_small(ln)
        kind = "h%d" % hl if is_heading else ("aside" if is_sm else "p")

        new_par = False
        if cur is None:
            new_par = True
        elif kind != cur["type"] or is_heading or prev_heading:
            new_par = True
        elif kind == "p":
            same_page = prev and prev["page"] == ln["page"]
            gap = (ln["y0"] - prev["y0"]) if same_page else None
            if gap is not None and gap > line_gap * 1.7:
                new_par = True
            if ln["x0"] > indent_thr:
                new_par = True
            if prev and prev["x1"] < short_if and RE_STRONG_END.search(prev["text"].strip()):
                new_par = True

        # PARAGRAF DEVAMI kuralları — geometrik tahminleri ezer. Kitapta bir
        # paragraf cümle ortasında bitmez; bittiği görünüyorsa satır bölünmesidir.
        if (new_par and cur is not None and not is_heading and not prev_heading
                and cur["type"] in ("p", "aside") and kind in ("p", "aside")):
            a = cur["text"].strip()
            closed = bool(RE_STRONG_END.search(a))          # gerçekten cümle bitmiş mi
            # (1) küçük harfle devam / tireyle bölünmüş kelime
            if is_continuation(a, ln["text"]):
                new_par = False
            # (2) önceki satır virgül, noktalı virgül, tire vb. ile bitmiş:
            #     sonraki satır BÜYÜK harfle başlasa da (özel ad) devamdır
            elif RE_OPEN_END.search(a):
                new_par = False
            # (3) önceki satır sağ kenara kadar DOLU ve cümle kapanmamış:
            #     paragrafın son satırı kısa olur; dolu satır demek ki devam ediyor
            elif (not closed and prev is not None
                  and prev["x1"] >= right_edge - body_size * 1.2
                  and ln["x0"] <= indent_thr):
                new_par = False
        # (4) KAPAK/İÇ KAPAK: art arda gelen kısa BÜYÜK HARF satırları tek başlık
        #     parçasıdır ("GLADYO" + "VE" + "ERGENEKON" -> "GLADYO VE ERGENEKON")
        elif (new_par and cur is not None and front_matter(ln)
                and cur["type"] == kind and kind in ("p", "h1", "h2", "h3")
                and all_caps(cur["text"]) and all_caps(ln["text"])
                and word_count(cur["text"]) <= 6 and word_count(ln["text"]) <= 6
                and not RE_STRONG_END.search(cur["text"].strip())):
            new_par = False

        if new_par:
            push_body()
            cur = {"type": kind, "text": ln["text"].strip()}
        else:
            _join(cur, ln["text"], opts.dehyphen)

        prev = ln
        prev_heading = is_heading
    flush_aside()
    push_body()
    return paras


# ============================== 7) metin temizleme ==========================
def fix_apostrophe(text):
    def repl(m):
        suf = m.group(2)
        return (m.group(1) + suf) if suf in SUFFIX_SET else m.group(0)
    return re.sub(r"([’'])\s+([" + LOWER + r"]+)", repl, text)


def clean_text(text):
    text = text.replace(SOFT, "")
    text = fix_apostrophe(text)
    text = re.sub(r"[ \t ]+", " ", text).strip()
    return text


# ============================== 8) docx yazımı ==============================
def _field(par, instr, placeholder):
    r = par.add_run()
    f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "begin"); f.set(qn("w:dirty"), "true")
    r._r.append(f)
    r = par.add_run()
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = instr
    r._r.append(i)
    r = par.add_run()
    f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "separate")
    r._r.append(f)
    par.add_run(placeholder)
    r = par.add_run()
    f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "end")
    r._r.append(f)


def _add_runs_with_tokens(par, text):
    """Metni yazar; ⟦FN7⟧ tokenlarını KENDİ başına birer run yapar (sonradan
    gerçek dipnot referansıyla değiştirilir)."""
    for part in re.split("(⟦FN\\d+⟧)", text):
        if not part:
            continue
        par.add_run(part)


def write_docx(paras, opts, out_path, title, fn_list, toc_found):
    doc = Document()

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, m, Cm(2.5))

    st = doc.styles["Normal"]
    st.font.name = opts.font
    st.font.size = Pt(opts.size)
    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if opts.align == "both" else WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = opts.spacing
    pf.space_after = Pt(opts.para_space)
    if opts.indent:
        pf.first_line_indent = Cm(1.0)

    if opts.pagenum:
        fpar = sec.footer.paragraphs[0]
        fpar.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _field(fpar, " PAGE ", "1")

    if toc_found or fn_list:
        try:
            el = OxmlElement("w:updateFields"); el.set(qn("w:val"), "true")
            doc.settings.element.insert(0, el)   # açılışta alanları (İçindekiler) güncelle
        except Exception:
            pass

    for p in paras:
        if p["type"] == "toc":
            h = doc.add_heading("İçindekiler", level=1)
            h.paragraph_format.first_line_indent = Cm(0)
            tp = doc.add_paragraph()
            tp.paragraph_format.first_line_indent = Cm(0)
            _field(tp, ' TOC \\o "1-3" \\h \\z \\u ',
                   "İçindekiler — Word'de açınca kendiliğinden dolar (gerekirse F9).")
            continue
        txt = clean_text(p["text"])
        if not txt:
            continue
        if p["type"] in ("h1", "h2", "h3"):
            par = doc.add_heading("", level=int(p["type"][1]))
            par.paragraph_format.first_line_indent = Cm(0)
            _add_runs_with_tokens(par, RE_FN_TOKEN.sub("", txt))
        elif p["type"] == "aside":
            par = doc.add_paragraph()
            par.paragraph_format.first_line_indent = Cm(0)
            _add_runs_with_tokens(par, txt)
            for r in par.runs:
                r.italic = True
                r.font.size = Pt(max(8, opts.size - 1.5))
        else:
            par = doc.add_paragraph()
            _add_runs_with_tokens(par, txt)

    doc.core_properties.title = title
    doc.core_properties.author = "pdf2word"
    doc.save(out_path)

    if fn_list:
        inject_footnotes(out_path, fn_list, opts)


# --------- gerçek Word dipnotları: paket düzeyinde OOXML enjeksiyonu --------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _footnotes_xml(fn_list, opts):
    sz = int(max(8, opts.size - 2) * 2)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<w:footnotes xmlns:w="{W_NS}">',
        '<w:footnote w:type="separator" w:id="-1"><w:p>'
        '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0"/></w:pPr>'
        '<w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p>'
        '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0"/></w:pPr>'
        '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for fid, text in fn_list:
        t = xesc(clean_text(text))
        parts.append(
            f'<w:footnote w:id="{fid}"><w:p>'
            '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>'
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve"> {t}</w:t></w:r>'
            '</w:p></w:footnote>')
    parts.append("</w:footnotes>")
    return "".join(parts).encode("utf-8")


def inject_footnotes(path, fn_list, opts):
    from lxml import etree
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}

    # 1) document.xml: ⟦FN7⟧ tokenlı runları gerçek referansla değiştir
    root = etree.fromstring(items["word/document.xml"])
    ns = {"w": W_NS}
    for t in root.iter(f"{{{W_NS}}}t"):
        m = RE_FN_TOKEN.fullmatch(t.text or "")
        if not m:
            continue
        fid = int(m.group(1))
        run = t.getparent()
        new = etree.SubElement(run.getparent(), f"{{{W_NS}}}r")
        rpr = etree.SubElement(new, f"{{{W_NS}}}rPr")
        va = etree.SubElement(rpr, f"{{{W_NS}}}vertAlign")
        va.set(f"{{{W_NS}}}val", "superscript")
        ref = etree.SubElement(new, f"{{{W_NS}}}footnoteReference")
        ref.set(f"{{{W_NS}}}id", str(fid))
        run.getparent().replace(run, new)
    items["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 2) footnotes.xml ekle
    items["word/footnotes.xml"] = _footnotes_xml(fn_list, opts)

    # 3) [Content_Types].xml
    ct = items["[Content_Types].xml"].decode("utf-8")
    if "footnotes+xml" not in ct:
        ct = ct.replace("</Types>",
            '<Override PartName="/word/footnotes.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
            "</Types>")
    items["[Content_Types].xml"] = ct.encode("utf-8")

    # 4) document.xml.rels
    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    if "relationships/footnotes" not in rels:
        rels = rels.replace("</Relationships>",
            '<Relationship Id="rIdFn1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
            "</Relationships>")
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)


# ============================== ana akış ====================================
def convert(pdf_path, out_path, opts, on_page=None):
    """on_page(okunan_sayfa, toplam, ocr_mu) verilirse ilerleme bildirilir."""
    doc = fitz.open(pdf_path)

    ocr_mode = getattr(opts, "ocr", "auto")
    want_ocr, garbage_idx = choose_ocr_pages(doc, ocr_mode)
    tessdata = find_tessdata() if want_ocr else None
    ocr_missing = bool(want_ocr) and tessdata is None
    pages, markers, ocr_done = extract_pages(
        doc, want_ocr if tessdata else set(), tessdata, on_page)

    all_sizes = [l["size"] for pg in pages for l in pg["lines"]]
    body_size = mode_round(all_sizes, 0.5) or median(all_sizes)

    fn_list = []
    if opts.footnotes:
        notes = collect_footnotes(pages, body_size, opts.dehyphen)
        fn_list = match_footnotes(pages, markers, notes, body_size)
    else:
        for pg in pages:                       # tokenları düz rakama çevir
            for ln in pg["lines"]:
                ln["text"] = RE_FN_TOKEN.sub(
                    lambda m: next((mk["num"] for mk in markers
                                    if mk["k"] == int(m.group(1))), ""), ln["text"])

    toc_found = replace_toc(pages, body_size) if opts.toc else False
    if opts.headers:
        strip_headers_footers(pages)
    scanned = len(want_ocr) >= max(3, doc.page_count * 0.6) and bool(tessdata)
    paras = assemble(pages, body_size, opts, scanned=scanned)

    ocr_set = set(ocr_done)
    garbage_pages = [p + 1 for p in garbage_idx if p not in ocr_set]

    title = os.path.splitext(os.path.basename(pdf_path))[0]
    write_docx(paras, opts, out_path, title, fn_list, toc_found)
    return {
        "pages": doc.page_count,
        "paragraphs": len([p for p in paras if p["text"].strip()]),
        "footnotes": len(fn_list),
        "toc": toc_found,
        "ocr_pages": [p + 1 for p in ocr_done],
        "ocr_missing": ocr_missing,
        "garbage_pages": garbage_pages,
    }


def main():
    ap = argparse.ArgumentParser(description="PDF kitabı 'baskı öncesi' Word'e çevirir.")
    ap.add_argument("input", help="PDF dosyası veya klasör")
    ap.add_argument("output", nargs="?", help="Çıktı .docx (tek dosya için)")
    ap.add_argument("--font", default="Times New Roman")
    ap.add_argument("--size", type=float, default=12.0)
    ap.add_argument("--align", choices=["both", "left"], default="both")
    ap.add_argument("--spacing", type=float, default=1.15, help="Satır aralığı çarpanı")
    ap.add_argument("--para-space", type=float, default=6.0,
                    help="Paragraf arası boşluk (punto, varsayılan 6)")
    ap.add_argument("--no-indent", dest="indent", action="store_false")
    ap.add_argument("--no-dehyphen", dest="dehyphen", action="store_false")
    ap.add_argument("--no-headers", dest="headers", action="store_false")
    ap.add_argument("--no-headings", dest="headings", action="store_false")
    ap.add_argument("--no-asides", dest="asides", action="store_false")
    ap.add_argument("--no-footnotes", dest="footnotes", action="store_false",
                    help="Dipnotları gerçek Word dipnotu yapma")
    ap.add_argument("--no-toc", dest="toc", action="store_false",
                    help="İçindekileri canlı alana çevirme")
    ap.add_argument("--no-pagenum", dest="pagenum", action="store_false",
                    help="Altbilgiye sayfa numarası koyma")
    ap.add_argument("--ocr", choices=["auto", "full", "off"], default="auto",
                    help="OCR: auto=yalnız bozuk/taranmış sayfalar (varsayılan), "
                         "full=tüm sayfalar (taranmış kitap), off=kapalı")
    ap.set_defaults(indent=True, dehyphen=True, headers=True, headings=True,
                    asides=True, footnotes=True, toc=True, pagenum=True)
    opts = ap.parse_args()

    def report(name, r):
        extra = []
        if r["footnotes"]:
            extra.append(f"{r['footnotes']} gerçek dipnot")
        if r["toc"]:
            extra.append("canlı içindekiler")
        if r.get("ocr_pages"):
            extra.append(f"OCR uygulanan sayfalar: {r['ocr_pages']}")
        if r.get("ocr_missing"):
            extra.append("OCR gerekliydi ama tur.traineddata bulunamadı")
        if r["garbage_pages"]:
            extra.append(f"hâlâ bozuk görünen sayfalar: {r['garbage_pages']}")
        print(f"✓ {name}: {r['pages']} sayfa -> {r['paragraphs']} paragraf"
              + ("  [" + " · ".join(extra) + "]" if extra else ""))

    if os.path.isdir(opts.input):
        pdfs = sorted(f for f in os.listdir(opts.input) if f.lower().endswith(".pdf"))
        if not pdfs:
            sys.exit("Klasörde PDF yok.")
        for f in pdfs:
            src = os.path.join(opts.input, f)
            dst = os.path.splitext(src)[0] + ".docx"
            report(f, convert(src, dst, opts))
    else:
        out = opts.output or (os.path.splitext(opts.input)[0] + ".docx")
        report(os.path.basename(out), convert(opts.input, out, opts))


if __name__ == "__main__":
    main()
