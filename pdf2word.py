#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf2word.py — Kitap PDF'lerini düzenlenebilir Word (.docx) dosyasına çevirir.

Neden bu araç (tarayıcı sürümünden farkı):
  Türkçe kitaplarda satır sonları çoğunlukla YUMUŞAK TİRE (U+00AD) ile bölünür
  (ör. "dışlaya­\nrak"). Tarayıcıdaki pdf.js bu tireyi SİLER; o yüzden "dışlaya"
  ile "rak" arasına yanlışlıkla boşluk girer ("dışlaya rak"). PyMuPDF ise
  yumuşak tireyi KORUR — böylece kelimeler doğru birleşir ("dışlayarak").

Ne yapar:
  * Satırları doğru paragraflara birleştirir (sayfa sınırlarını da aşarak).
  * Satır sonu tirelerini (yumuşak tire ve normal tire) doğru birleştirir;
    büyük harfle başlayan bileşiklerde tireyi korur (Sovyet-Rusya).
  * Apostrof sonrası bölünen ekleri onarır (Türkiye' nin -> Türkiye'nin).
  * Her sayfada tekrar eden üstbilgi/altbilgi ve salt sayfa numaralarını atar.
  * Büyük puntolu başlıkları Word'de Heading olarak biçimler.
  * Küçük puntolu kenar notu/dipnotları gövdeye karıştırmaz (ayrı paragraf).
  * Metin katmanı bozuk (çöp) sayfaları raporlar (OCR gerekebilir).

Kelimeleri/cümleleri DEĞİŞTİRMEZ; yalnızca doğru yere yerleştirir.

Kurulum:  pip install pymupdf python-docx
Kullanım: python pdf2word.py "kitap.pdf" [cikti.docx]
          python pdf2word.py "kitap.pdf" --no-dehyphen --font "Georgia" --size 12
          python pdf2word.py klasor/   (klasördeki tüm PDF'leri toplu çevirir)
"""
import argparse
import os
import re
import statistics
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF gerekli:  pip install pymupdf")
try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    sys.exit("python-docx gerekli:  pip install python-docx")

SOFT = "­"                     # yumuşak tire
LOWER = "a-zçğıöşü"
UPPER = "A-ZÇĞİÖŞÜ"
LETTER = LOWER + UPPER
VOWELS = set("aeıioöuüAEIİOÖUÜ")

RE_LETTER_HYPHEN = re.compile("[" + LETTER + "]-$")
RE_STARTS_LOWER = re.compile("^[" + LOWER + "]")
RE_PAGENUM = re.compile(r"^[\s.\-–—·]*\d{1,4}[\s.\-–—·]*$")
RE_ROMAN = re.compile(r"^[\s.\-–—·]*[ivxlcdmIVXLCDM]{1,7}[\s.\-–—·]*$")
RE_STRONG_END = re.compile(r"[.!?…][\"»”’')\]]?$")

# Apostrof sonrası bölünmüş Türkçe ekler (yalnızca bunlar birleştirilir; tırnak
# kapanışları -Kürdistan' adı- bozulmasın diye TAM eşleşme aranır).
_SUFFIXES = """
nin nın nun nün in ın un ün
e a ye ya na ne
i ı u ü yi yı yu yü
de da te ta nde nda
den dan ten tan nden ndan
le la yle yla
deki daki ndeki ndaki teki taki
li lı lu lü lik lık luk lük
ler lar leri ları lere lara lerde larda lerden lardan lerin ların
si sı su sü nin
dir dır dur dür tir tır tur tür
dan'sonra
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


RE_MIDCAP = re.compile("[" + LOWER + "][" + UPPER + "]")


def _token_weird(w):
    core = w.strip(".,;:!?()[]{}\"'«»…-/")
    letters = [c for c in core if c.isalpha()]
    if not letters:
        return False
    if all(c.isupper() for c in letters):        # akronim (ABD, FETÖ, PKK) -> normal
        return False
    if RE_MIDCAP.search(core):                    # kelime ortası küçük->BÜYÜK -> OCR çöpü
        return True
    if len(letters) >= 5 and not any(c in VOWELS for c in core):  # uzun, seslisiz
        return True
    return False


def looks_garbage(text):
    """Bozuk metin katmanı işareti: tokenlerin çoğu 'tuhaf' (kelime-içi büyük
    harf, sesli harfsiz uzun dizi, aşırı noktalama)."""
    toks = text.split()
    if len(toks) < 4:
        return False
    bad = sum(1 for w in toks if _token_weird(w))
    return bad / len(toks) >= 0.35


RE_LETDIG = re.compile("[" + LETTER + r"]\d|\d[" + LETTER + "]")   # harf-rakam bitişik
RARE_CHARS = set("ąăĄ*=>~^`|ª•■□▪")                                # OCR çöpünde sık


def _token_garbage(w):
    return _token_weird(w) or bool(RE_LETDIG.search(w)) or any(c in RARE_CHARS for c in w)


def garbage_line(line):
    """Bozuk metin katmanı satırı (kupür/şema OCR'ı) — rapor için."""
    toks = line.split()
    if len(toks) < 3:
        return False
    bad = sum(1 for w in toks if _token_garbage(w))
    return bad / len(toks) >= 0.30


def weird_line(text):
    """Kısa da olsa 'başlık olamayacak kadar bozuk' satır (kupür/şema metni)."""
    s = text.strip()
    if not s:
        return True
    if any(_token_weird(w) for w in s.split()):
        return True
    nonspace = [c for c in s if not c.isspace()]
    if not nonspace:
        return True
    alpha = sum(1 for c in nonspace if c.isalpha())
    return alpha / len(nonspace) < 0.7    # rakam/noktalama ağırlıklı -> başlık değil


# ------------------------------------------------------------- 1) satır çıkarımı
def extract_lines(doc):
    """Her sayfadan satırları (metin + geometri) çıkarır. Yumuşak tire korunur."""
    pages = []
    for pno in range(doc.page_count):
        page = doc[pno]
        W, H = page.rect.width, page.rect.height
        d = page.get_text("dict")
        lines = []
        for block in d["blocks"]:
            if block.get("type", 0) != 0:      # 0 = metin bloğu (resimleri atla)
                continue
            for ln in block["lines"]:
                spans = ln["spans"]
                text = "".join(s["text"] for s in spans)
                if not text.strip():
                    continue
                sizes = [s["size"] for s in spans if s["text"].strip()]
                x0, y0, x1, y1 = ln["bbox"]
                lines.append({
                    "text": text, "x0": x0, "x1": x1, "y0": y0, "y1": y1,
                    "size": median(sizes) if sizes else 0,
                    "page": pno, "W": W, "H": H,
                })
        pages.append(lines)
    return pages


# ---------------------------------------------- 2) üstbilgi/altbilgi temizliği
def norm_hf(s):
    return re.sub(r"\s+", " ", re.sub(r"\d+", "#", s.lower())).strip()


def strip_headers_footers(pages):
    n = len(pages)
    if n < 3:
        return pages
    from collections import Counter
    top, bot = Counter(), Counter()
    for lines in pages:
        if not lines:
            continue
        if word_count(lines[0]["text"]) <= 8:
            top[norm_hf(lines[0]["text"])] += 1
        if word_count(lines[-1]["text"]) <= 8:
            bot[norm_hf(lines[-1]["text"])] += 1
    rep = max(3, round(n * 0.3))

    def is_num(s):
        return bool(RE_PAGENUM.match(s) or RE_ROMAN.match(s.strip()))

    out = []
    for lines in pages:
        if len(lines) <= 1:
            if lines and RE_PAGENUM.match(lines[0]["text"]):
                out.append([])
            else:
                out.append(lines)
            continue
        a, b = 0, len(lines)
        t, bt = lines[0], lines[-1]
        if is_num(t["text"]) or (word_count(t["text"]) <= 8 and top[norm_hf(t["text"])] >= rep):
            a = 1
        if b - a > 0 and (is_num(bt["text"]) or (word_count(bt["text"]) <= 8 and bot[norm_hf(bt["text"])] >= rep)):
            b -= 1
        out.append(lines[a:b])
    return out


# --------------------------------------------------- 3) paragraf birleştirme
def heading_level(line, body_size):
    t = line["text"].strip()
    if not t or word_count(t) > 14 or weird_line(t):   # kupür/şema iri puntosunu başlık sayma
        return 0
    if not (t[0].isupper() or t[0].isdigit()):         # başlık büyük harf/rakamla başlar
        return 0
    if line["size"] >= body_size * 1.6:
        return 1
    if line["size"] >= body_size * 1.35:
        return 2
    return 0


def append_line(par, line_text, dehyphen):
    add = line_text.strip()
    if not add:
        return
    if not par["text"]:
        par["text"] = add
        return
    prev = par["text"].rstrip()
    if prev.endswith(SOFT):
        par["text"] = prev[:-1] + add                       # yumuşak tire: birleştir
    elif RE_LETTER_HYPHEN.search(prev):
        if dehyphen and RE_STARTS_LOWER.match(add):
            par["text"] = prev[:-1] + add                   # satır sonu tiresi: birleştir
        else:
            par["text"] = prev + add                        # büyük harf/kapalı: tireyi koru
    else:
        par["text"] = prev + " " + add


def assemble(pages, opts):
    flat = [ln for lines in pages for ln in lines]
    if not flat:
        return []

    sizes = [l["size"] for l in flat]
    body_size = mode_round(sizes, 0.5) or median(sizes)
    body_lines = [l for l in flat if abs(l["size"] - body_size) <= 0.6]
    ref = body_lines if len(body_lines) >= 5 else flat
    body_left = mode_round([l["x0"] for l in ref], 3)
    right_edge = percentile([l["x1"] for l in ref], 0.9)
    body_width = max(1.0, right_edge - body_left)

    gaps = []
    for lines in pages:
        body = [l for l in lines if abs(l["size"] - body_size) <= 0.6]
        for i in range(1, len(body)):
            g = body[i]["y0"] - body[i - 1]["y0"]
            if 0 < g < body_size * 4:
                gaps.append(g)
    line_gap = median(gaps) or body_size * 1.4
    indent_thr = body_left + max(6, body_size * 0.9)
    short_if = right_edge - body_size * 3.5

    def is_margin_note(ln):
        """Sayfa kenarındaki dar, gövdeyle aynı ya da küçük puntolu not/kaynak."""
        if not opts.asides:
            return False
        narrow = (ln["x1"] - ln["x0"]) < 0.55 * body_width
        offset = ln["x0"] > body_left + 3.0 * body_size
        return narrow and offset and ln["size"] <= body_size + 0.5

    def is_small(ln):     # dipnot / küçük puntolu alıntı
        return opts.asides and body_size and ln["size"] <= body_size * 0.86

    paras = []
    cur = None            # açık gövde paragrafı
    prev = None
    prev_heading = False
    cur_aside = None      # oluşturulmakta olan kenar notu
    pending = []          # gövde paragrafı bitince yayımlanacak notlar

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
        paras.extend(pending)     # notları, kesintiye uğrattıkları paragraftan SONRA yay
        pending.clear()

    for ln in flat:
        hl = heading_level(ln, body_size) if opts.headings else 0
        is_heading = hl > 0

        # 1) Açık bir gövde paragrafını kesen kenar notu: paragrafı BÖLME,
        #    notu tamponla; gövde sonraki satırla birleşmeye devam etsin.
        if not is_heading and cur is not None and is_margin_note(ln):
            if cur_aside is None:
                cur_aside = {"type": "aside", "text": ln["text"].strip()}
            else:
                append_line(cur_aside, ln["text"], opts.dehyphen)
            continue
        flush_aside()

        # 2) Küçük puntolu satır (dipnot): yerinde ayrı paragraf
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

        if new_par:
            push_body()
            cur = {"type": kind, "text": ln["text"].strip()}
        else:
            append_line(cur, ln["text"], opts.dehyphen)

        prev = ln
        prev_heading = is_heading
    flush_aside()
    push_body()
    return paras


# --------------------------------------------------------- 4) metin temizleme
def fix_apostrophe(text):
    """Türkiye' nin -> Türkiye'nin  (yalnızca bilinen Türkçe ekler; tırnaklar korunur)."""
    def repl(m):
        suf = m.group(2)
        return (m.group(1) + suf) if suf in SUFFIX_SET else m.group(0)
    return re.sub(r"([’'])\s+([" + LOWER + r"]+)", repl, text)


def clean_text(text):
    text = text.replace(SOFT, "")
    text = fix_apostrophe(text)
    text = re.sub(r"[ \t ]+", " ", text).strip()
    return text


# --------------------------------------------------------------- 5) docx yaz
def write_docx(paras, opts, out_path, title):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = opts.font
    st.font.size = Pt(opts.size)
    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if opts.align == "both" else WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = opts.spacing
    pf.space_after = Pt(0)
    if opts.indent:
        pf.first_line_indent = Cm(1.0)

    for p in paras:
        txt = clean_text(p["text"])
        if not txt:
            continue
        if p["type"] == "h1":
            doc.add_heading(txt, level=1)
        elif p["type"] == "h2":
            doc.add_heading(txt, level=2)
        elif p["type"] == "aside":
            par = doc.add_paragraph(txt)
            par.paragraph_format.first_line_indent = Cm(0)
            for r in par.runs:
                r.italic = True
                r.font.size = Pt(max(8, opts.size - 1.5))
        else:
            doc.add_paragraph(txt)

    doc.core_properties.title = title
    doc.core_properties.author = "pdf2word.py"
    doc.save(out_path)


# ---------------------------------------------------------------- ana akış
def convert(pdf_path, out_path, opts):
    doc = fitz.open(pdf_path)
    pages = extract_lines(doc)
    if opts.headers:
        pages = strip_headers_footers(pages)
    paras = assemble(pages, opts)

    # çöp (bozuk metin katmanı) sayfa raporu — satır düzeyinde say
    garbage_pages = []
    for pno in range(doc.page_count):
        lines = [l for l in doc[pno].get_text().split("\n") if l.strip()]
        g = sum(1 for l in lines if garbage_line(l))
        if g >= 2:
            garbage_pages.append(pno + 1)

    title = os.path.splitext(os.path.basename(pdf_path))[0]
    write_docx(paras, opts, out_path, title)
    return {
        "pages": doc.page_count,
        "paragraphs": len([p for p in paras if p["text"].strip()]),
        "garbage_pages": sorted(garbage_pages),
    }


def main():
    ap = argparse.ArgumentParser(description="PDF kitabı düzenlenebilir Word'e çevirir.")
    ap.add_argument("input", help="PDF dosyası veya klasör")
    ap.add_argument("output", nargs="?", help="Çıktı .docx (tek dosya için)")
    ap.add_argument("--font", default="Times New Roman")
    ap.add_argument("--size", type=float, default=12.0)
    ap.add_argument("--align", choices=["both", "left"], default="both")
    ap.add_argument("--spacing", type=float, default=1.15)
    ap.add_argument("--no-indent", dest="indent", action="store_false")
    ap.add_argument("--no-dehyphen", dest="dehyphen", action="store_false",
                    help="Satır sonu tirelerini birleştirme (metni birebir koru)")
    ap.add_argument("--no-headers", dest="headers", action="store_false",
                    help="Üstbilgi/altbilgi ve sayfa no temizliğini kapat")
    ap.add_argument("--no-headings", dest="headings", action="store_false")
    ap.add_argument("--no-asides", dest="asides", action="store_false",
                    help="Küçük puntolu kenar notu/dipnot ayrımını kapat")
    ap.set_defaults(indent=True, dehyphen=True, headers=True, headings=True, asides=True)
    opts = ap.parse_args()

    if os.path.isdir(opts.input):
        pdfs = sorted(f for f in os.listdir(opts.input) if f.lower().endswith(".pdf"))
        if not pdfs:
            sys.exit("Klasörde PDF yok.")
        for f in pdfs:
            src = os.path.join(opts.input, f)
            dst = os.path.splitext(src)[0] + ".docx"
            r = convert(src, dst, opts)
            print(f"✓ {f}: {r['pages']} sayfa -> {r['paragraphs']} paragraf"
                  + (f"  [çöp sayfa: {r['garbage_pages']}]" if r["garbage_pages"] else ""))
    else:
        out = opts.output or (os.path.splitext(opts.input)[0] + ".docx")
        r = convert(opts.input, out, opts)
        print(f"✓ {os.path.basename(out)} yazıldı: {r['pages']} sayfa -> "
              f"{r['paragraphs']} paragraf")
        if r["garbage_pages"]:
            print(f"⚠ Metin katmanı bozuk görünen sayfalar (OCR gerekebilir): {r['garbage_pages']}")


if __name__ == "__main__":
    main()
