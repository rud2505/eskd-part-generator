"""Генератор деталей по ЕСКД — окно программы.
Запуск: двойной щелчок по «Запуск.bat» или  python app.py"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import deps  # noqa: E402  — библиотеки из папки lib
import os, sys, threading, traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shaft import Shaft, parse_steps, MATERIALS, fmt, chamfered
from flange import Flange

APP = "Генератор деталей ЕСКД"
DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
RA = ["0,8", "1,6", "3,2", "6,3", "12,5"]
KINDS = ["Вал ступенчатый", "Фланец", "Втулка (скоро)"]
# значения по умолчанию для штампа: тип -> (наименование, обозначение, материал)
DEFAULTS = {
    "Вал ступенчатый": ("Вал", "АБВГ.711111.001", "Сталь 45 ГОСТ 1050-2013"),
    "Фланец": ("Фланец", "АБВГ.711111.002", "Сталь 20 ГОСТ 1050-2013"),
}
# поля фланца: (атрибут, подпись, значение по умолчанию)
FLANGE_FIELDS = [
    ("D", "Наружный диаметр D", "160"),
    ("d", "Центральное отверстие d", "60"),
    ("H", "Толщина диска H", "20"),
    ("D1", "Окружность отверстий D1", "125"),
    ("n", "Число отверстий n", "4"),
    ("d0", "Диаметр отверстий d0", "18"),
    ("Dh", "Ступица: диаметр (0 — нет)", "0"),
    ("Hh", "Ступица: длина", "0"),
]


def num(s):
    return float(str(s).strip().replace(",", ".") or 0)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP)
        self.geometry("1060x800")
        self.minsize(980, 740)
        self.result = None
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        self._build_ui()
        self.on_kind()

    # ------------------------------------------------------------------ окно
    def _build_ui(self):
        left = ttk.Frame(self, padding=12)
        left.pack(side="left", fill="y")
        right = ttk.Frame(self, padding=(0, 12, 12, 12))
        right.pack(side="left", fill="both", expand=True)

        box = ttk.LabelFrame(left, text=" 1. Программа САПР ", padding=8)
        box.pack(fill="x")
        self.cad = tk.StringVar(value="SolidWorks")
        ttk.Radiobutton(box, text="SolidWorks", value="SolidWorks", variable=self.cad).pack(anchor="w")
        ttk.Radiobutton(box, text="КОМПАС-3D", value="КОМПАС", variable=self.cad).pack(anchor="w")

        box = ttk.LabelFrame(left, text=" 2. Тип детали ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        self.kind = ttk.Combobox(box, values=KINDS, state="readonly", width=30)
        self.kind.current(0)
        self.kind.pack(fill="x")
        self.kind.bind("<<ComboboxSelected>>", lambda e: self.on_kind())
        self._kind_prev = KINDS[0]

        # 3. Параметры: две панели, показывается нужная
        self.params = ttk.Frame(left)
        self.params.pack(fill="x", pady=(8, 0))
        self.p_shaft = ttk.LabelFrame(self.params, text=" 3. Ступени слева направо: диаметр × длина, мм ",
                                      padding=8)
        self.steps = tk.Text(self.p_shaft, width=34, height=6, font=("Consolas", 11))
        self.steps.insert("1.0", "20 x 100\n30 x 100\n40 x 50")
        self.steps.pack(fill="x")
        self.steps.bind("<KeyRelease>", lambda e: self.redraw())
        ttk.Label(self.p_shaft, text="Одна строка — одна ступень. Например: 25 x 60",
                  foreground="#666").pack(anchor="w", pady=(4, 0))

        self.p_flange = ttk.LabelFrame(self.params, text=" 3. Размеры фланца, мм ", padding=8)
        self.fl = {}
        for i, (key, label, val) in enumerate(FLANGE_FIELDS):
            ttk.Label(self.p_flange, text=label).grid(row=i, column=0, sticky="w", pady=1, padx=(0, 8))
            v = tk.StringVar(value=val)
            e = ttk.Entry(self.p_flange, textvariable=v, width=10)
            e.grid(row=i, column=1, sticky="w", pady=1)
            e.bind("<KeyRelease>", lambda ev: self.redraw())
            self.fl[key] = v

        box = ttk.LabelFrame(left, text=" 4. Основная надпись и параметры ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        self.f_name = self._field(box, 0, "Наименование", "Вал")
        self.f_desig = self._field(box, 1, "Обозначение", "АБВГ.711111.001")
        self.f_dev = self._field(box, 2, "Разработал", "")
        ttk.Label(box, text="Материал").grid(row=3, column=0, sticky="w", pady=2)
        self.f_mat = ttk.Combobox(box, values=list(MATERIALS), state="readonly", width=26)
        self.f_mat.current(0)
        self.f_mat.grid(row=3, column=1, sticky="we", pady=2)
        ttk.Label(box, text="Фаски, мм").grid(row=4, column=0, sticky="w", pady=2)
        self.f_ch = ttk.Combobox(box, values=["авто", "0", "0,5", "1", "1,6", "2", "2,5"], width=26)
        self.f_ch.set("авто")
        self.f_ch.grid(row=4, column=1, sticky="we", pady=2)
        self.f_ch.bind("<<ComboboxSelected>>", lambda e: self.redraw())
        self.f_ch.bind("<KeyRelease>", lambda e: self.redraw())
        ttk.Label(box, text="Шероховатость Ra").grid(row=5, column=0, sticky="w", pady=2)
        self.f_ra = ttk.Combobox(box, values=RA, state="readonly", width=26)
        self.f_ra.set("6,3")
        self.f_ra.grid(row=5, column=1, sticky="we", pady=2)

        box = ttk.LabelFrame(left, text=" 5. Папка для файлов ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        self.f_out = tk.StringVar(value=DEFAULT_OUT)
        ttk.Entry(box, textvariable=self.f_out, width=28).pack(side="left", fill="x", expand=True)
        ttk.Button(box, text="…", width=3, command=self.choose_dir).pack(side="left", padx=(4, 0))

        box = ttk.LabelFrame(left, text=" 6. Пакет деталей из Excel ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        ttk.Button(box, text="Создать шаблон…", command=self.new_template).pack(side="left")
        self.btn_batch = ttk.Button(box, text="Построить из Excel…", command=self.run_batch)
        self.btn_batch.pack(side="left", padx=(6, 0))

        self.title_lbl = ttk.Label(right, text="Эскиз", font=("Segoe UI", 10, "bold"))
        self.title_lbl.pack(anchor="w")
        self.canvas = tk.Canvas(right, height=270, bg="white", highlightthickness=1,
                                highlightbackground="#bbb")
        self.canvas.pack(fill="x")
        self.canvas.bind("<Configure>", lambda e: self.redraw())
        self.info = ttk.Label(right, text="", foreground="#333")
        self.info.pack(anchor="w", pady=(4, 8))
        self.btn = ttk.Button(right, text="Построить модель и чертёж", style="Big.TButton",
                              command=self.run)
        self.btn.pack(fill="x", pady=(0, 4))
        self.prog = ttk.Progressbar(right, mode="determinate")
        self.prog.pack(fill="x", pady=(0, 8))

        ttk.Label(right, text="Ход работы", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.log_box = tk.Text(right, height=10, font=("Consolas", 10), state="disabled", bg="#fafafa")
        self.log_box.pack(fill="both", expand=True)
        bar = ttk.Frame(right)
        bar.pack(fill="x", pady=(8, 0))
        self.btn_pdf = ttk.Button(bar, text="Открыть PDF", state="disabled",
                                  command=lambda: os.startfile(self.result["pdf"]))
        self.btn_pdf.pack(side="left")
        self.btn_dir = ttk.Button(bar, text="Открыть папку", state="disabled",
                                  command=lambda: os.startfile(os.path.dirname(self.result["pdf"])))
        self.btn_dir.pack(side="left", padx=6)

    def _field(self, box, row, label, value):
        ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", pady=2, padx=(0, 8))
        v = tk.StringVar(value=value)
        ttk.Entry(box, textvariable=v, width=28).grid(row=row, column=1, sticky="we", pady=2)
        return v

    def choose_dir(self):
        d = filedialog.askdirectory(initialdir=self.f_out.get())
        if d:
            self.f_out.set(os.path.normpath(d))

    def on_kind(self):
        k = self.kind.get()
        if k not in DEFAULTS:               # «скоро» — вернуть прежний выбор
            self.kind.set(self._kind_prev)
            return
        # подставить значения штампа по умолчанию, если пользователь их не менял
        old = DEFAULTS[self._kind_prev]
        new = DEFAULTS[k]
        if self.f_name.get() in (old[0], ""):
            self.f_name.set(new[0])
        if self.f_desig.get() in (old[1], ""):
            self.f_desig.set(new[1])
        if self.f_mat.get() == old[2]:
            self.f_mat.set(new[2])
        self._kind_prev = k
        self.p_shaft.pack_forget()
        self.p_flange.pack_forget()
        (self.p_flange if k == "Фланец" else self.p_shaft).pack(fill="x")
        self.title_lbl.config(text="Эскиз фланца: разрез и вид слева" if k == "Фланец" else "Эскиз вала")
        self.redraw()

    # ------------------------------------------------------------------ данные
    def make_part(self):
        ch = self.f_ch.get().strip().replace(",", ".")
        if ch != "авто":
            float(ch)
        common = dict(name=self.f_name.get().strip(), designation=self.f_desig.get().strip(),
                      material=self.f_mat.get(), developer=self.f_dev.get().strip(),
                      chamfer=ch, ra=self.f_ra.get())
        if self.kind.get() == "Фланец":
            vals = {k: num(v.get()) for k, v in self.fl.items()}
            vals["n"] = int(vals["n"])
            return Flange(**vals, **common)
        return Shaft(parse_steps(self.steps.get("1.0", "end")), **common)

    # ------------------------------------------------------------------ эскиз
    def redraw(self):
        c = self.canvas
        c.delete("all")
        W, H = c.winfo_width() or 600, c.winfo_height() or 270
        try:
            p = self.make_part()
        except Exception as e:
            c.create_text(W / 2, H / 2, text=f"⚠ {e}", fill="#c00", font=("Segoe UI", 10), width=W - 40)
            self.info.config(text="")
            return
        line, fill = "#1f3b63", "#e8eef7"
        dash = (12, 3, 2, 3)
        prof = chamfered(p.profile(), p.chamfers)
        sa, sb = p.pick_scale()
        if isinstance(p, Flange):
            L, D = p.length, p.D
            k = min((W - 120) / (L + D), (H - 50) / D)
            x0 = (W - (L + D) * k - 60) / 2
            y0 = H / 2
            for sgn in (1, -1):
                pts = [v for x, y in prof for v in (x0 + x * k, y0 - sgn * y * k)]
                c.create_polygon(pts, fill="#c9d6ea", outline=line, width=2, stipple="gray50")
            for y, z in p.holes():          # отверстия, попавшие в разрез (на оси Y)
                if abs(z) < 1e-6:
                    for e in (-1, 1):
                        yy = y0 - (y + e * p.d0 / 2) * k
                        c.create_line(x0, yy, x0 + p.H * k, yy, fill=line, width=2)
            c.create_line(x0 - 10, y0, x0 + L * k + 10, y0, dash=dash, fill="#888")
            cx, R = x0 + L * k + 60 + D * k / 2, D / 2 * k
            c.create_oval(cx - R, y0 - R, cx + R, y0 + R, fill=fill, outline=line, width=2)
            r = p.d / 2 * k
            c.create_oval(cx - r, y0 - r, cx + r, y0 + r, fill="white", outline=line, width=2)
            R1 = p.D1 / 2 * k
            c.create_oval(cx - R1, y0 - R1, cx + R1, y0 + R1, outline="#888", dash=dash)
            for y, z in p.holes():
                hx, hy, hr = cx + z * k, y0 - y * k, p.d0 / 2 * k
                c.create_oval(hx - hr, hy - hr, hx + hr, hy + hr, fill="white", outline=line, width=2)
            c.create_line(cx - R - 8, y0, cx + R + 8, y0, dash=dash, fill="#888")
            c.create_line(cx, y0 - R - 8, cx, y0 + R + 8, dash=dash, fill="#888")
            c.create_text(cx, y0 + R + 14, text=f"{p.n} отв. Ø{fmt(p.d0)} на Ø{fmt(p.D1)}",
                          fill="#555", font=("Segoe UI", 9))
            info = f"Ø{fmt(D)} × {fmt(L)} мм · отверстий: {p.n} · фасок: {len(p.chamfers)}"
        else:
            L, D = p.length, p.dmax
            k = min((W - 80) / L, (H - 70) / D)
            x0, y0 = (W - L * k) / 2, H / 2
            pts = [(x0 + x * k, y0 - y * k) for x, y in prof]
            pts += [(x0 + x * k, y0 + y * k) for x, y in reversed(prof)]
            c.create_polygon([v for pt in pts for v in pt], fill=fill, outline=line, width=2)
            xs = p.x_edges()
            for i, (d, l) in enumerate(p.steps):
                xa, xb = x0 + xs[i] * k, x0 + xs[i + 1] * k
                if i:
                    rr = min(d, p.steps[i - 1][0]) / 2 * k
                    c.create_line(xa, y0 - rr, xa, y0 + rr, fill=line, width=2)
                if xb - xa > 28:
                    c.create_text((xa + xb) / 2, y0 - 9, text=f"Ø{fmt(d)}", fill=line,
                                  font=("Segoe UI", 9, "bold"))
                c.create_text((xa + xb) / 2, y0 + D / 2 * k + 14, text=fmt(l), fill="#555",
                              font=("Segoe UI", 9))
            c.create_line(x0 - 15, y0, x0 + L * k + 15, y0, dash=dash, fill="#888")
            info = (f"Длина {fmt(L)} мм · Ømax {fmt(D)} мм · ступеней: {len(p.steps)} · "
                    f"фасок: {len(p.chamfers)}")
        self.info.config(text=f"{info} · лист А3, масштаб {fmt(sa)}:{fmt(sb)}")

    # ------------------------------------------------------------------ запуск
    def log(self, text):
        def _add():
            self.log_box.config(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _add)

    def run(self):
        try:
            p = self.make_part()
        except Exception as e:
            messagebox.showerror(APP, f"Проверьте данные:\n{e}")
            return
        self.btn.config(state="disabled", text="Строю…")
        self.btn_pdf.config(state="disabled")
        self.btn_dir.config(state="disabled")
        self.log(f"— {self.cad.get()}: {p.summary()}")
        threading.Thread(target=self._work, args=(p, self.f_out.get(), self.cad.get()),
                         daemon=True).start()

    def _work(self, p, out, cad):
        try:
            if cad == "КОМПАС":
                from ks_driver import KompasDriver as Driver
            else:
                from sw_driver import SolidWorksDriver as Driver
            self.result = Driver(log=self.log).build(p, out)
            self.log("Готово ✓")
            self.after(0, lambda: (self.btn_pdf.config(state="normal"),
                                   self.btn_dir.config(state="normal")))
        except Exception as e:
            self.log(f"Ошибка: {e}")
            self.log(traceback.format_exc(limit=2))
        finally:
            self.after(0, lambda: self.btn.config(state="normal", text="Построить модель и чертёж"))

    # ------------------------------------------------------------------ пакет из Excel
    def _driver(self, cad):
        if cad == "КОМПАС":
            from ks_driver import KompasDriver as Driver
        else:
            from sw_driver import SolidWorksDriver as Driver
        return Driver(log=self.log)

    def new_template(self):
        from batch import make_template
        path = filedialog.asksaveasfilename(
            title="Куда сохранить шаблон", defaultextension=".xlsx",
            initialfile="Список деталей.xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            make_template(path)
            self.log(f"Шаблон сохранён: {path}")
            os.startfile(path)

    def run_batch(self):
        path = filedialog.askopenfilename(title="Список деталей (Excel)",
                                          filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        for b in (self.btn, self.btn_batch, self.btn_pdf, self.btn_dir):
            b.config(state="disabled")
        self.prog.config(value=0)
        self.log(f"— Пакет ({self.cad.get()}): {os.path.basename(path)}")
        threading.Thread(target=self._work_batch, args=(path, self.f_out.get(), self.cad.get()),
                         daemon=True).start()

    def _work_batch(self, path, out, cad):
        from batch import run
        try:
            def progress(i, n):
                self.after(0, lambda: self.prog.config(maximum=max(n, 1), value=i))
            res = run(path, self._driver(cad), out, log=self.log, progress=progress)
            if res["book"]:
                self.result = {"pdf": res["book"]}
                self.after(0, lambda: (self.btn_pdf.config(state="normal"),
                                       self.btn_dir.config(state="normal")))
            os.startfile(res["report"])
        except Exception as e:
            self.log(f"Ошибка: {e}")
            self.log(traceback.format_exc(limit=2))
        finally:
            self.after(0, lambda: (self.btn.config(state="normal"), self.btn_batch.config(state="normal")))


if __name__ == "__main__":
    App().mainloop()
