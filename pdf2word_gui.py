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
from types import SimpleNamespace

from pdf2word import convert   # dönüştürme motoru (fitz + python-docx)

FONTS = ["Times New Roman", "Georgia", "Cambria", "Calibri", "Arial", "Book Antiqua"]


def build_opts(font="Times New Roman", size=12.0, align="both", spacing=1.15,
               indent=True, dehyphen=True, headers=True, headings=True, asides=True):
    return SimpleNamespace(font=font, size=float(size), align=align, spacing=float(spacing),
                           indent=indent, dehyphen=dehyphen, headers=headers,
                           headings=headings, asides=asides)


def convert_files(files, opts, outdir=None, progress=None):
    """Dosyaları sırayla çevirir. progress(i, n, name, result|None, error|None)."""
    results = []
    n = len(files)
    for i, f in enumerate(files):
        name = os.path.basename(f)
        if progress:
            progress(i, n, name, None, None)
        try:
            base = os.path.splitext(os.path.basename(f))[0] + ".docx"
            out = os.path.join(outdir, base) if outdir else (os.path.splitext(f)[0] + ".docx")
            r = convert(f, out, opts)
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
    root.geometry("660x600")
    root.minsize(560, 520)

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

    # --- Dosya seçimi ---
    files_box = ttk.LabelFrame(main, text="1) PDF dosyaları", padding=10)
    files_box.pack(fill="both", expand=True)
    lst = tk.Listbox(files_box, height=6, activestyle="none")
    lst.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(files_box, command=lst.yview)
    sb.pack(side="left", fill="y")
    lst.config(yscrollcommand=sb.set)
    fbtns = ttk.Frame(files_box)
    fbtns.pack(side="left", fill="y", padx=(10, 0))

    def refresh_list():
        lst.delete(0, "end")
        for f in state["files"]:
            lst.insert("end", "  " + os.path.basename(f))

    def add_files():
        fs = filedialog.askopenfilenames(title="PDF seç", filetypes=[("PDF", "*.pdf")],
                                         initialdir=state["last_dir"] or os.path.expanduser("~"))
        for f in fs:
            if f not in state["files"]:
                state["files"].append(f)
                state["last_dir"] = os.path.dirname(f)
        refresh_list()

    def clear_files():
        state["files"].clear()
        refresh_list()

    ttk.Button(fbtns, text="PDF Seç…", command=add_files).pack(fill="x", pady=2)
    ttk.Button(fbtns, text="Listeyi Temizle", command=clear_files).pack(fill="x", pady=2)

    # --- Seçenekler ---
    opt = ttk.LabelFrame(main, text="2) Word biçimi", padding=10)
    opt.pack(fill="x", pady=(10, 0))
    v_font = tk.StringVar(value="Times New Roman")
    v_size = tk.DoubleVar(value=12.0)
    v_align = tk.StringVar(value="İki yana yasla")
    v_dehyph = tk.BooleanVar(value=True)
    v_head = tk.BooleanVar(value=True)
    v_heading = tk.BooleanVar(value=True)

    row1 = ttk.Frame(opt); row1.pack(fill="x", pady=2)
    ttk.Label(row1, text="Yazı tipi:").pack(side="left")
    ttk.Combobox(row1, textvariable=v_font, values=FONTS, width=18, state="readonly").pack(side="left", padx=(4, 14))
    ttk.Label(row1, text="Punto:").pack(side="left")
    ttk.Spinbox(row1, from_=8, to=20, increment=0.5, textvariable=v_size, width=5).pack(side="left", padx=(4, 14))
    ttk.Label(row1, text="Hizalama:").pack(side="left")
    ttk.Combobox(row1, textvariable=v_align, values=["İki yana yasla", "Sola yasla"], width=13,
                 state="readonly").pack(side="left", padx=(4, 0))

    row2 = ttk.Frame(opt); row2.pack(fill="x", pady=(6, 2))
    ttk.Checkbutton(row2, text="Satır sonu tirelerini birleştir", variable=v_dehyph).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row2, text="Üstbilgi/sayfa no temizle", variable=v_head).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(row2, text="Başlıkları algıla", variable=v_heading).pack(side="left")

    # --- Çevir + durum ---
    act = ttk.Frame(main); act.pack(fill="x", pady=(12, 0))
    go_btn = ttk.Button(act, text="Word'e Çevir")
    go_btn.pack(side="left")
    open_btn = ttk.Button(act, text="Kaydedilen klasörü aç", state="disabled")
    open_btn.pack(side="left", padx=8)

    prog = ttk.Progressbar(main, mode="determinate")
    prog.pack(fill="x", pady=(12, 4))
    status = ttk.Label(main, text="Hazır. PDF seçip 'Word'e Çevir'e basın.", foreground="#333")
    status.pack(anchor="w")

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

    def worker(files, opts):
        def prog_cb(i, n, name, r, err):
            q.put(("progress", i, n, name, r, err))
        results = convert_files(files, opts, outdir=state["outdir"], progress=prog_cb)
        q.put(("done", results))

    def start():
        if not state["files"]:
            messagebox.showinfo("PDF yok", "Önce en az bir PDF dosyası seçin.")
            return
        opts = build_opts(
            font=v_font.get(),
            size=v_size.get(),
            align="both" if v_align.get().startswith("İki") else "left",
            dehyphen=v_dehyph.get(), headers=v_head.get(), headings=v_heading.get(),
        )
        go_btn.config(state="disabled")
        open_btn.config(state="disabled")
        prog.config(maximum=len(state["files"]), value=0)
        status.config(text="Başlıyor…")
        threading.Thread(target=worker, args=(list(state["files"]), opts), daemon=True).start()
        root.after(120, poll)

    def poll():
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "progress":
                    _, i, n, name, r, err = msg
                    prog.config(value=i)
                    if r is None and err is None:
                        status.config(text=f"İşleniyor: {name}  ({i + 1}/{n})")
                    elif err:
                        status.config(text=f"HATA: {name} — {err}")
                    else:
                        status.config(text=f"Bitti: {name}  ({i}/{n})")
                elif msg[0] == "done":
                    finish(msg[1])
                    return
        except queue.Empty:
            pass
        root.after(120, poll)

    def finish(results):
        go_btn.config(state="normal")
        ok = [r for r in results if r.get("ok")]
        bad = [r for r in results if not r.get("ok")]
        lines = [f"{len(ok)} dosya başarıyla çevrildi." if ok else "Hiç dosya çevrilemedi."]
        for r in ok:
            gp = r.get("garbage_pages") or []
            extra = f"  ⚠ OCR gerekebilen sayfalar: {gp}" if gp else ""
            lines.append(f"• {os.path.basename(r['out'])}  ({r['paragraphs']} paragraf){extra}")
        for r in bad:
            lines.append(f"✗ {os.path.basename(r['file'])} — {r.get('error')}")
        status.config(text=lines[0])
        if ok:
            state["outdir_last"] = os.path.dirname(ok[-1]["out"])
            open_btn.config(state="normal",
                            command=lambda: open_folder(os.path.dirname(ok[-1]["out"])))
        messagebox.showinfo("Bitti", "\n".join(lines))

    go_btn.config(command=start)
    refresh_list()
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
            print(f"SELFTEST OK: {r['out']}  ({r['paragraphs']} paragraf, {r['pages']} sayfa)")
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
