"""طابعة الستكرات — استيراد من إكسل وطباعة على طابعة حرارية (ويندوز)."""
import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import ImageTk

import core

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
    from PIL import ImageWin

    HAS_WIN = True
except ImportError:  # للتجربة خارج ويندوز
    HAS_WIN = False

def _settings_path():
    """وضع محمول: الإعدادات بجانب البرنامج، وإن لم يُسمح بالكتابة هناك فمجلد المستخدم."""
    base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    if os.access(base, os.W_OK):
        return os.path.join(base, "StickerPrinter.json")
    return os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "StickerPrinter.json")


SETTINGS_FILE = _settings_path()
NONE = "— بدون —"
QR_POS = {"يسار": "left", "يمين": "right", "أعلى": "top"}
ALIGN = {"وسط": "center", "يمين": "right", "يسار": "left"}


# ---------- الطباعة ----------
def list_printers():
    if not HAS_WIN:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]


def default_printer():
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def make_dc(printer, w_mm, h_mm):
    """ينشئ DC للطابعة بمقاس الورق المطلوب (وإن فشل يستعمل إعدادات الطابعة)."""
    try:
        h = win32print.OpenPrinter(printer)
        try:
            dm = win32print.GetPrinter(h, 2)["pDevMode"]
        finally:
            win32print.ClosePrinter(h)
        dm.PaperSize = 256  # DMPAPER_USER
        dm.PaperWidth = int(round(w_mm * 10))
        dm.PaperLength = int(round(h_mm * 10))
        dm.Fields |= win32con.DM_PAPERSIZE | win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH
        return win32ui.CreateDCFromHandle(win32gui.CreateDC("WINSPOOL", printer, dm))
    except Exception:
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer)
        return dc


def print_labels(printer, items, s, offx_mm, offy_mm, progress=None):
    """items = [(lines, qr, copies), ...]"""
    dc = make_dc(printer, s.width_mm, s.height_mm)
    dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
    dpi = dpi_x or 203
    phys_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
    phys_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
    x = core.mm_to_px(offx_mm, dpi) - phys_x
    y = core.mm_to_px(offy_mm, dpi) - phys_y
    dc.StartDoc("Stickers")
    try:
        total = sum(c for _, _, c in items)
        done = 0
        for lines, qr, copies in items:
            img = core.render_label(lines, qr, s, dpi).convert("RGB")
            for _ in range(copies):
                dc.StartPage()
                ImageWin.Dib(img).draw(dc.GetHandleOutput(), (x, y, x + img.width, y + img.height))
                dc.EndPage()
                done += 1
                if progress:
                    progress(done, total)
        dc.EndDoc()
    except Exception:
        dc.AbortDoc()
        raise


# ---------- الواجهة ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("طابعة الستكرات")
        self.geometry("1020x700")
        self.minsize(900, 620)
        self.headers, self.rows = [], []
        self.preview_img = None

        v = self.v = {}
        defaults = dict(width=50, height=30, margin=2, font=12, qr=20, qrpos="يسار",
                        align="وسط", bold=True, border=False, offx=0, offy=0, printer="")
        saved = {}
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                saved = json.load(f)
        except Exception:
            pass
        defaults.update({k: x for k, x in saved.items() if k in defaults})
        for k, val in defaults.items():
            v[k] = tk.BooleanVar(value=val) if isinstance(val, bool) else tk.StringVar(value=str(val))
        if not v["printer"].get():
            v["printer"].set(default_printer())
        for var in v.values():
            var.trace_add("write", lambda *_: self.after_idle(self.update_preview))

        self.build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --- بناء الواجهة ---
    def build_ui(self):
        left = ttk.Frame(self, padding=10)
        left.pack(side="left", fill="y")
        right = ttk.Frame(self, padding=10)
        right.pack(side="left", fill="both", expand=True)

        # 1) الملف
        f = ttk.LabelFrame(left, text="1) ملف الإكسل", padding=8)
        f.pack(fill="x")
        ttk.Button(f, text="اختيار ملف…", command=self.pick_file).pack(fill="x")
        self.file_lbl = ttk.Label(f, text="لم يتم اختيار ملف", wraplength=300)
        self.file_lbl.pack(fill="x", pady=(4, 0))

        # 2) الأعمدة
        f = ttk.LabelFrame(left, text="2) الأعمدة", padding=8)
        f.pack(fill="x", pady=8)
        ttk.Label(f, text="أعمدة النص (اختر واحداً أو أكثر، كل عمود سطر):").pack(anchor="w")
        self.cols = tk.Listbox(f, selectmode="extended", height=5, exportselection=False)
        self.cols.pack(fill="x")
        self.cols.bind("<<ListboxSelect>>", lambda e: self.update_preview())
        self.qr_col = self._combo(f, "عمود QR:")
        self.cnt_col = self._combo(f, "عمود عدد النسخ:")

        # 3) التصميم
        f = ttk.LabelFrame(left, text="3) حجم وتصميم الملصق", padding=8)
        f.pack(fill="x")
        grid = [("العرض (مم)", "width"), ("الارتفاع (مم)", "height"), ("الهامش (مم)", "margin"),
                ("حجم الخط (pt)", "font"), ("حجم QR (مم)", "qr")]
        for i, (t, k) in enumerate(grid):
            ttk.Label(f, text=t).grid(row=i, column=0, sticky="w", pady=2)
            ttk.Spinbox(f, from_=0, to=500, increment=1, width=8, textvariable=self.v[k]).grid(row=i, column=1, sticky="e")
        r = len(grid)
        ttk.Label(f, text="مكان QR").grid(row=r, column=0, sticky="w", pady=2)
        ttk.Combobox(f, values=list(QR_POS), state="readonly", width=8, textvariable=self.v["qrpos"]).grid(row=r, column=1, sticky="e")
        ttk.Label(f, text="محاذاة النص").grid(row=r + 1, column=0, sticky="w", pady=2)
        ttk.Combobox(f, values=list(ALIGN), state="readonly", width=8, textvariable=self.v["align"]).grid(row=r + 1, column=1, sticky="e")
        ttk.Checkbutton(f, text="السطر الأول بخط عريض", variable=self.v["bold"]).grid(row=r + 2, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(f, text="رسم إطار حول الملصق", variable=self.v["border"]).grid(row=r + 3, column=0, columnspan=2, sticky="w")
        f.columnconfigure(0, weight=1)

        # 4) الطابعة
        f = ttk.LabelFrame(left, text="4) الطابعة", padding=8)
        f.pack(fill="x", pady=8)
        ttk.Combobox(f, values=list_printers(), textvariable=self.v["printer"]).pack(fill="x")
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(4, 0))
        ttk.Label(row, text="إزاحة أفقية/رأسية (مم):").pack(side="left")
        ttk.Spinbox(row, from_=-20, to=20, increment=0.5, width=5, textvariable=self.v["offx"]).pack(side="left", padx=2)
        ttk.Spinbox(row, from_=-20, to=20, increment=0.5, width=5, textvariable=self.v["offy"]).pack(side="left")

        # المعاينة
        top = ttk.Frame(right)
        top.pack(fill="x")
        ttk.Label(top, text="معاينة الصف رقم:").pack(side="left")
        self.row_no = tk.IntVar(value=1)
        self.row_spin = ttk.Spinbox(top, from_=1, to=1, width=6, textvariable=self.row_no, command=self.update_preview)
        self.row_spin.pack(side="left", padx=4)
        self.row_spin.bind("<KeyRelease>", lambda e: self.update_preview())
        self.count_lbl = ttk.Label(top, text="")
        self.count_lbl.pack(side="left", padx=8)

        self.canvas = tk.Canvas(right, bg="#d9d9d9", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, pady=8)
        self.canvas.bind("<Configure>", lambda e: self.update_preview())

        rng = ttk.Frame(right)
        rng.pack(fill="x")
        ttk.Label(rng, text="طباعة من صف").pack(side="left")
        self.r_from = tk.StringVar(value="1")
        self.r_to = tk.StringVar(value="")
        ttk.Entry(rng, width=6, textvariable=self.r_from).pack(side="left", padx=4)
        ttk.Label(rng, text="إلى").pack(side="left")
        ttk.Entry(rng, width=6, textvariable=self.r_to).pack(side="left", padx=4)
        ttk.Label(rng, text="(اتركه فارغاً = للنهاية)").pack(side="left")

        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=8)
        ttk.Button(btns, text="🖨  طباعة", command=self.do_print).pack(side="left", fill="x", expand=True, padx=2, ipady=6)
        ttk.Button(btns, text="طباعة ملصق تجريبي", command=lambda: self.do_print(test=True)).pack(side="left", padx=2, ipady=6)
        ttk.Button(btns, text="حفظ كصور PNG", command=self.save_png).pack(side="left", padx=2, ipady=6)
        self.bar = ttk.Progressbar(right)
        self.bar.pack(fill="x")
        self.status = ttk.Label(right, text="")
        self.status.pack(fill="x")

    def _combo(self, parent, label):
        ttk.Label(parent, text=label).pack(anchor="w", pady=(6, 0))
        c = ttk.Combobox(parent, state="readonly", values=[NONE])
        c.set(NONE)
        c.pack(fill="x")
        c.bind("<<ComboboxSelected>>", lambda e: self.update_preview())
        return c

    # --- الملف ---
    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            self.headers, self.rows = core.read_excel(path)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذّرت قراءة الملف:\n{e}")
            return
        self.file_lbl.config(text=f"{os.path.basename(path)} — {len(self.rows)} صف")
        self.cols.delete(0, "end")
        for h in self.headers:
            self.cols.insert("end", h)
        if self.headers:
            self.cols.selection_set(0)
        vals = [NONE] + self.headers
        for c in (self.qr_col, self.cnt_col):
            c["values"] = vals
            c.set(NONE)
        low = [h.lower() for h in self.headers]
        for i, h in enumerate(low):  # تخمين ذكي للأعمدة
            if h in ("qr", "رمز qr", "qr code", "باركود", "رابط") or "qr" in h:
                self.qr_col.set(self.headers[i])
            if h in ("copies", "count", "qty", "عدد", "الكمية", "عدد النسخ", "نسخ"):
                self.cnt_col.set(self.headers[i])
        self.row_spin.config(to=max(1, len(self.rows)))
        self.row_no.set(1)
        self.r_to.set(str(len(self.rows)))
        self.update_preview()

    # --- القراءة من الواجهة ---
    def settings(self):
        def num(k, d):
            try:
                return float(self.v[k].get().replace(",", "."))
            except ValueError:
                return d
        return core.LabelSettings(
            width_mm=max(5, num("width", 50)), height_mm=max(5, num("height", 30)),
            margin_mm=max(0, num("margin", 2)), font_pt=max(5, num("font", 12)),
            qr_mm=max(5, num("qr", 20)), qr_pos=QR_POS.get(self.v["qrpos"].get(), "left"),
            align=ALIGN.get(self.v["align"].get(), "center"),
            bold_first=self.v["bold"].get(), border=self.v["border"].get())

    def row_content(self, i):
        row = self.rows[i]
        lines = [row[j] for j in self.cols.curselection() if j < len(row) and row[j]]
        qr = ""
        if self.qr_col.get() != NONE:
            qr = row[self.headers.index(self.qr_col.get())]
        copies = 1
        if self.cnt_col.get() != NONE:
            try:
                copies = max(0, int(float(row[self.headers.index(self.cnt_col.get())] or 1)))
            except ValueError:
                copies = 1
        return lines, qr, copies

    # --- المعاينة ---
    def update_preview(self):
        self.canvas.delete("all")
        if not self.rows:
            self.canvas.create_text(20, 20, anchor="nw", text="اختر ملف إكسل لتظهر المعاينة هنا")
            return
        try:
            i = min(max(int(self.row_no.get()), 1), len(self.rows)) - 1
        except (tk.TclError, ValueError):
            return
        lines, qr, copies = self.row_content(i)
        s = self.settings()
        img = core.render_label(lines, qr, s, 203)
        cw, ch = max(self.canvas.winfo_width(), 100), max(self.canvas.winfo_height(), 100)
        k = min((cw - 30) / img.width, (ch - 30) / img.height)
        shown = img.convert("RGB").resize((max(1, int(img.width * k)), max(1, int(img.height * k))))
        self.preview_img = ImageTk.PhotoImage(shown)
        self.canvas.create_image(cw // 2, ch // 2, image=self.preview_img)
        self.count_lbl.config(text=f"من {len(self.rows)} صف — النسخ لهذا الصف: {copies}")

    # --- الطباعة والحفظ ---
    def _items(self, test=False):
        if not self.rows:
            messagebox.showwarning("تنبيه", "اختر ملف إكسل أولاً")
            return None
        if not self.cols.curselection() and self.qr_col.get() == NONE:
            messagebox.showwarning("تنبيه", "اختر عموداً واحداً على الأقل للنص أو QR")
            return None
        if test:
            lines, qr, _ = self.row_content(max(int(self.row_no.get()) - 1, 0))
            return [(lines, qr, 1)]
        try:
            a = max(int(self.r_from.get() or 1), 1)
            b = min(int(self.r_to.get() or len(self.rows)), len(self.rows))
        except ValueError:
            messagebox.showwarning("تنبيه", "أرقام الصفوف غير صحيحة")
            return None
        return [self.row_content(i) for i in range(a - 1, b)]

    def do_print(self, test=False):
        if not HAS_WIN:
            messagebox.showinfo("تنبيه", "الطباعة تعمل على ويندوز فقط. استعمل «حفظ كصور PNG» للتجربة.")
            return
        items = self._items(test)
        if not items:
            return
        printer = self.v["printer"].get()
        if not printer:
            messagebox.showwarning("تنبيه", "اختر الطابعة")
            return
        total = sum(c for _, _, c in items)
        if total == 0:
            messagebox.showinfo("تنبيه", "لا توجد ملصقات للطباعة")
            return
        if not test and not messagebox.askyesno("تأكيد", f"سيتم طباعة {total} ملصق على:\n{printer}"):
            return
        self.bar.config(maximum=total, value=0)

        def prog(d, t):
            self.bar.config(value=d)
            self.status.config(text=f"تمت طباعة {d} من {t}")
            self.update()
        try:
            print_labels(printer, items, self.settings(), self._f("offx"), self._f("offy"), prog)
            self.status.config(text=f"✔ اكتملت الطباعة ({total} ملصق)")
        except Exception as e:
            messagebox.showerror("خطأ في الطباعة", str(e))

    def save_png(self):
        items = self._items()
        if not items:
            return
        folder = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if not folder:
            return
        s, n = self.settings(), 0
        for lines, qr, copies in items:
            if copies:
                n += 1
                core.render_label(lines, qr, s, 300).save(os.path.join(folder, f"sticker_{n:04d}.png"), dpi=(300, 300))
        self.status.config(text=f"تم حفظ {n} صورة في {folder}")

    def _f(self, k):
        try:
            return float(self.v[k].get().replace(",", "."))
        except ValueError:
            return 0.0

    def on_close(self):
        try:
            data = {k: (x.get()) for k, x in self.v.items()}
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
