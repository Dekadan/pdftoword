#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF'ten Word'e — Kitap Dönüştürücü (masaüstü uygulaması)

Çift tıklayınca açılan pencereli uygulama: PDF seç -> "Word'e Çevir" -> temiz .docx.
Kurulum/komut satırı gerektirmez (GitHub tarafından .exe/.app olarak derlenir).
Dönüştürme motoru pdf2word.py'dir (satır sonu tirelerini doğru birleştirir).

Geliştirici notu:
  Görünmez test için:  python pdf2word_gui.py --selftest girdi.pdf cikti.docx
  (bu mod tkinter'ı yüklemez; paketlemenin/bağımlılıkların çalıştığını doğrular)
"""
import os
import sys
import threading
import queue
import time
from types import SimpleNamespace

from pdf2word import convert   # dönüştürme motoru (fitz + python-docx)

FONTS = ["Times New Roman", "Georgia", "Cambria", "Calibri", "Arial", "Book Antiqua"]


def build_opts(font="Times New Roman", size=12.0, align="both", spacing=1.15,
               para_space=6.0, indent=True, dehyphen=True, headers=True,
               headings=True, asides=True, footnotes=True, toc=True, pagenum=True,
               ocr="auto"):
    return SimpleNamespace(font=font, size=float(size), align=align, spacing=float(spacing),
                           para_space=float(para_space), indent=indent, dehyphen=dehyphen,
                           headers=headers, headings=headings, asides=asides,
                           footnotes=footnotes, toc=toc, pagenum=pagenum, ocr=ocr)


def unique_path(path):
    """Var olan dosyanın üzerine yazma: "Kitap.docx" -> "Kitap (2).docx"."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def convert_files(files, opts, outdir=None, progress=None, page_progress=None):
    """Dosyaları sırayla çevirir.
    progress(i, n, name, result|None, error|None) — dosya düzeyinde,
    page_progress(i, n, name, sayfa, toplam_sayfa, ocr_mu) — sayfa düzeyinde."""
    results = []
    n = len(files)
    if outdir:
        os.makedirs(outdir, exist_ok=True)
    for i, f in enumerate(files):
        name = os.path.basename(f)
        if progress:
            progress(i, n, name, None, None)
        try:
            base = os.path.splitext(os.path.basename(f))[0] + ".docx"
            out = os.path.join(outdir, base) if outdir else (os.path.splitext(f)[0] + ".docx")
            out = unique_path(out)
            cb = None
            if page_progress:
                def cb(p, tot, is_ocr, _i=i, _n=n, _name=name):
                    page_progress(_i, _n, _name, p, tot, is_ocr)
            r = convert(f, out, opts, on_page=cb)
            r["file"] = f
            r["out"] = out
            r["ok"] = True
            results.append(r)
            if progress:
                progress(i + 1, n, name, r, None)
        except Exception as e:   # bir dosya hata verse de diğerlerine devam
            results.append({"file": f, "ok": False, "error": str(e)})
            if progress:
                progress(i + 1, n, name, None, str(e))
    return results


# ----------------------------------------------------------------- GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("PDF'ten Word'e — Kitap Dönüştürücü")
    root.geometry("740x820")
    root.minsize(700, 730)

    state = {"files": [], "outdir": None, "last_dir": None}
    q = queue.Queue()

    pad = {"padx": 12, "pady": 6}
    main = ttk.Frame(root, padding=14)
    main.pack(fill="both", expand=True)

    ttk.Label(main, text="📖 PDF'ten Word'e — Kitap Dönüştürücü",
              font=("Segoe UI", 15, "bold")).pack(anchor="w")
    ttk.Label(main, text="PDF kitaplarını, satırları doğru paragraflara birleştirerek düzenlenebilir "
                         "Word'e çevirir. Her şey bu bilgisayarda çalışır; internet/ücret gerekmez.",
              wraplength=620, foreground="#555").pack(anchor="w", pady=(2, 10))

    # Durum/işlem alanı EN ALTA sabitlenir: pencere küçük olsa da hep görünür kalır
    # (aksi halde "başladı mı?" belirsizliği doğuyordu).
    bottom = ttk.Frame(main)
    bottom.pack(side="bottom", fill="x", pady=(10, 0))

    # --- Dosya seçimi ---
    files_box = ttk.LabelFrame(main, text="1) PDF dosyaları", padding=10)
    files_box.pack(fill="both", expand=True)
    lst = tk.Listbox(files_box, height=6, activestyle="none", selectmode="extended")
    lst.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(files_box, command=lst.yview)
    sb.pack(side="left", fill="y")
    lst.config(yscrollcommand=sb.set)
    fbtns = ttk.Frame(files_box)
    fbtns.pack(side="left", fill="y", padx=(10, 0))

    count_lbl = ttk.Label(main, text="", foreground="#555", font=("Segoe UI", 9))

    def refresh_list():
        lst.delete(0, "end")
        for f in state["files"]:
            lst.insert("end", "  " + os.path.basename(f))
        n = len(state["files"])
        count_lbl.config(text=(f"{n} dosya seçili · listeden bir dosyaya çift tıklayarak "
                               f"çıkarabilirsiniz" if n else "Henüz PDF seçilmedi."))

    def add_paths(paths):
        for f in paths:
            if f and f not in state["files"]:
                state["files"].append(f)
                state["last_dir"] = os.path.dirname(f)
        refresh_list()
        syncgo()

    def add_files():
        add_paths(filedialog.askopenfilenames(
            title="PDF seç (birden çok seçebilirsiniz)",
            filetypes=[("PDF dosyaları", "*.pdf")],
            initialdir=state["last_dir"] or os.path.expanduser("~")))

    def remove_selected(event=None):
        sel = list(lst.curselection())
        if not sel:
            messagebox.showinfo("Seçim yok",
                                "Önce listeden çıkarmak istediğiniz dosyayı tıklayın.")
            return
        for i in sorted(sel, reverse=True):
            del state["files"][i]
        refresh_list()
        syncgo()

    def clear_files():
        if state["files"] and messagebox.askyesno("Listeyi temizle",
                                                  "Listedeki tüm dosyalar çıkarılsın mı?"):
            state["files"].clear()
            refresh_list()
            syncgo()

    ttk.Button(fbtns, text="➕ PDF Ekle…", command=add_files).pack(fill="x", pady=2)
    ttk.Button(fbtns, text="➖ Seçileni Çıkar", command=remove_selected).pack(fill="x", pady=2)
    ttk.Button(fbtns, text="Tümünü Temizle", command=clear_files).pack(fill="x", pady=2)
    lst.bind("<Double-Button-1>", remove_selected)
    lst.bind("<Delete>", remove_selected)
    count_lbl.pack(anchor="w", pady=(4, 0))

    # --- Kayıt klasörü ---
    out_box = ttk.LabelFrame(main, text="2) Word dosyaları nereye kaydedilsin?", padding=10)
    out_box.pack(fill="x", pady=(10, 0))
    out_lbl = ttk.Label(out_box, text="PDF'in bulunduğu klasöre kaydedilecek",
                        foreground="#555")
    out_lbl.pack(side="left", fill="x", expand=True)

    def pick_outdir():
        d = filedialog.askdirectory(title="Kayıt klasörünü seç",
                                    initialdir=state["outdir"] or state["last_dir"]
                                    or os.path.expanduser("~"))
        if d:
            state["outdir"] = d
            out_lbl.config(text=d, foreground="#111")

    def reset_outdir():
        state["outdir"] = None
        out_lbl.config(text="PDF'in bulunduğu klasöre kaydedilecek", foreground="#555")

    ttk.Button(out_box, text="Klasör Seç…", command=pick_outdir).pack(side="left", padx=(8, 0))
    ttk.Button(out_box, text="Varsayılan", command=reset_outdir).pack(side="left", padx=(6, 0))

    # --- Seçenekler ---
    opt = ttk.LabelFrame(main, text="3) Word biçimi ve metin işleme", padding=10)
    opt.pack(fill="x", pady=(10, 0))
    v_font = tk.StringVar(value="Times New Roman")
    v_size = tk.DoubleVar(value=12.0)
    v_align = tk.StringVar(value="İki yana yasla")
    v_dehyph = tk.BooleanVar(value=True)
    v_head = tk.BooleanVar(value=True)
    v_heading = tk.BooleanVar(value=True)

    v_pspace = tk.DoubleVar(value=6.0)
    v_fn = tk.BooleanVar(value=True)
    v_toc = tk.BooleanVar(value=True)
    v_pgnum = tk.BooleanVar(value=True)

    row1 = ttk.Frame(opt); row1.pack(fill="x", pady=2)
    ttk.Label(row1, text="Yazı tipi:").pack(side="left")
    ttk.Combobox(row1, textvariable=v_font, values=FONTS, width=18, state="readonly").pack(side="left", padx=(4, 14))
    ttk.Label(row1, text="Punto:").pack(side="left")
    ttk.Spinbox(row1, from_=8, to=20, increment=0.5, textvariable=v_size, width=5).pack(side="left", padx=(4, 14))
    ttk.Label(row1, text="Hizalama:").pack(side="left")
    ttk.Combobox(row1, textvariable=v_align, values=["İki yana yasla", "Sola yasla"], width=13,
                 state="readonly").pack(side="left", padx=(4, 0))

    row1b = ttk.Frame(opt); row1b.pack(fill="x", pady=(6, 2))
    ttk.Label(row1b, text="Paragraf arası boşluk (punto):").pack(side="left")
    ttk.Spinbox(row1b, from_=0, to=24, increment=1, textvariable=v_pspace, width=5).pack(side="left", padx=(4, 0))

    row2 = ttk.Frame(opt); row2.pack(fill="x", pady=(6, 2))
    ttk.Checkbutton(row2, text="Satır sonu tirelerini birleştir", variable=v_dehyph).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row2, text="Üstbilgi/sayfa no temizle", variable=v_head).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row2, text="Başlıkları algıla", variable=v_heading).pack(side="left")

    v_ocr = tk.BooleanVar(value=True)

    row3 = ttk.Frame(opt); row3.pack(fill="x", pady=(6, 2))
    ttk.Checkbutton(row3, text="Dipnotları gerçek Word dipnotu yap", variable=v_fn).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row3, text="Canlı İçindekiler alanı", variable=v_toc).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row3, text="Sayfa numarası (altbilgi)", variable=v_pgnum).pack(side="left")

    row4 = ttk.Frame(opt); row4.pack(fill="x", pady=(6, 2))
    ttk.Checkbutton(row4, text="Bozuk/taranmış sayfalara OCR uygula (Türkçe; biraz yavaşlatır)",
                    variable=v_ocr).pack(side="left")

    # --- Çevir + durum (en altta sabit) ---
    act = ttk.Frame(bottom); act.pack(fill="x")
    go_btn = ttk.Button(act, text="▶  Word'e Çevir")
    go_btn.pack(side="left")
    open_btn = ttk.Button(act, text="📂 Kayıt klasörünü aç", state="disabled")
    open_btn.pack(side="left", padx=8)

    prog = ttk.Progressbar(bottom, mode="determinate")
    prog.pack(fill="x", pady=(10, 4))
    status = ttk.Label(bottom, text="Hazır. PDF ekleyip “Word'e Çevir”e basın.",
                       foreground="#333", font=("Segoe UI", 10, "bold"))
    status.pack(anchor="w")
    substatus = ttk.Label(bottom, text="", foreground="#555", font=("Segoe UI", 9))
    substatus.pack(anchor="w")

    def syncgo():
        go_btn.config(state=("normal" if state["files"] and not state.get("busy") else "disabled"))

    def open_folder(path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess; subprocess.Popen(["open", path])
            else:
                import subprocess; subprocess.Popen(["xdg-open", path])
        except Exception:
            pass

    def worker(files, opts, outdir):
        def prog_cb(i, n, name, r, err):
            q.put(("progress", i, n, name, r, err))

        def page_cb(i, n, name, p, tot, is_ocr):
            q.put(("page", i, n, name, p, tot, is_ocr))
        results = convert_files(files, opts, outdir=outdir,
                                progress=prog_cb, page_progress=page_cb)
        q.put(("done", results))

    def start():
        if not state["files"]:
            messagebox.showinfo("PDF yok", "Önce “➕ PDF Ekle…” ile en az bir dosya seçin.")
            return
        opts = build_opts(
            font=v_font.get(),
            size=v_size.get(),
            align="both" if v_align.get().startswith("İki") else "left",
            para_space=v_pspace.get(),
            dehyphen=v_dehyph.get(), headers=v_head.get(), headings=v_heading.get(),
            footnotes=v_fn.get(), toc=v_toc.get(), pagenum=v_pgnum.get(),
            ocr="auto" if v_ocr.get() else "off",
        )
        state["busy"] = True
        state["t0"] = time.time()
        go_btn.config(state="disabled", text="⏳ Çevriliyor…")
        open_btn.config(state="disabled")
        prog.config(maximum=100, value=0)
        n = len(state["files"])
        status.config(text=f"▶ Başladı — {n} dosya çevriliyor…", foreground="#0a6")
        substatus.config(text="Dosya okunuyor…")
        root.update_idletasks()          # düğmeye basar basmaz ekrana yansısın
        threading.Thread(target=worker,
                         args=(list(state["files"]), opts, state["outdir"]),
                         daemon=True).start()
        root.after(80, poll)

    def poll():
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "page":
                    _, i, n, name, p, tot, is_ocr = msg
                    frac = (i + (p / max(1, tot))) / max(1, n)
                    prog.config(value=frac * 100)
                    el = int(time.time() - state.get("t0", time.time()))
                    tag = " · OCR (görüntüden okunuyor)" if is_ocr else ""
                    status.config(text=f"⏳ Çevriliyor: {name}   ({i + 1}/{n})",
                                  foreground="#0a6")
                    substatus.config(text=f"sayfa {p + 1}/{tot}{tag} · geçen süre {el} sn")
                elif msg[0] == "progress":
                    _, i, n, name, r, err = msg
                    if err:
                        status.config(text=f"✗ HATA: {name} — {err}", foreground="#a11")
                    elif r is not None:
                        prog.config(value=(i / max(1, n)) * 100)
                        substatus.config(text=f"✓ {name} bitti — Word dosyası kaydedildi")
                elif msg[0] == "done":
                    finish(msg[1])
                    return
        except queue.Empty:
            pass
        root.after(80, poll)

    def finish(results):
        state["busy"] = False
        go_btn.config(state="normal", text="▶  Word'e Çevir")
        prog.config(value=100)
        ok = [r for r in results if r.get("ok")]
        bad = [r for r in results if not r.get("ok")]
        lines = [f"{len(ok)} dosya başarıyla çevrildi." if ok else "Hiç dosya çevrilemedi."]
        for r in ok:
            gp = r.get("garbage_pages") or []
            bits = [f"{r['paragraphs']} paragraf"]
            if r.get("footnotes"):
                bits.append(f"{r['footnotes']} gerçek dipnot")
            if r.get("toc"):
                bits.append("canlı içindekiler")
            if r.get("ocr_pages"):
                bits.append(f"{len(r['ocr_pages'])} sayfaya OCR")
            extra = f"  ⚠ OCR gerekebilen sayfalar: {gp}" if gp else ""
            lines.append(f"• {os.path.basename(r['out'])}  ({', '.join(bits)}){extra}")
        for r in bad:
            lines.append(f"✗ {os.path.basename(r['file'])} — {r.get('error')}")
        el = int(time.time() - state.get("t0", time.time()))
        status.config(text=("✓ " + lines[0]) if ok else ("✗ " + lines[0]),
                      foreground=("#0a6" if ok else "#a11"))
        if ok:
            folder = os.path.dirname(ok[-1]["out"])
            substatus.config(text=f"Kayıt yeri: {folder}  ·  toplam {el} sn")
            open_btn.config(state="normal", command=lambda: open_folder(folder))
            lines.append("")
            lines.append(f"Kayıt yeri: {folder}")
        else:
            substatus.config(text="")
        messagebox.showinfo("Bitti", "\n".join(lines))

    go_btn.config(command=start)
    refresh_list()
    syncgo()

    # Otomatik arayüz testi için kanca (normal kullanımda dokunulmaz)
    root.test_hooks = {"state": state, "add_paths": add_paths, "list": lst,
                       "remove_selected": remove_selected, "clear": clear_files,
                       "outdir_label": out_lbl, "reset_outdir": reset_outdir,
                       "go": go_btn, "status": status, "substatus": substatus,
                       "count": count_lbl}
    if os.environ.get("PDF2WORD_GUI_SELFCLOSE"):   # test: pencereyi kur, kısa süre sonra kapat
        root.after(400, root.destroy)
    root.mainloop()


def main():
    # Görünmez test/otomasyon modu (tkinter yüklemez)
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        opts = build_opts()
        res = convert_files([sys.argv[2]], opts, outdir=None)
        r = res[0]
        if r.get("ok"):
            print(f"SELFTEST OK: {r['out']}  ({r['paragraphs']} paragraf, {r['pages']} sayfa, "
                  f"{r.get('footnotes', 0)} dipnot, içindekiler={r.get('toc')}, "
                  f"OCR sayfaları={r.get('ocr_pages')})")
            # istenen çıktı adına taşı
            if sys.argv[3] and sys.argv[3] != r["out"]:
                os.replace(r["out"], sys.argv[3])
                print("->", sys.argv[3])
            return 0
        print("SELFTEST FAIL:", r.get("error"))
        return 1
    run_gui()
    return 0


if __name__ == "__main__":
    sys.exit(main())
