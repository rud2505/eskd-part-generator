"""Пакетный режим: список деталей в Excel -> модели, чертежи, PDF, отчёт и общий PDF-комплект.

Excel-файл: два листа — «Валы» и «Фланцы», одна строка = одна деталь.
Запуск без окна:  python batch.py sw "Список.xlsx"   (или ks — для КОМПАСа)"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import deps  # noqa: E402  — библиотеки из папки lib
import os, sys, time
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

from shaft import Shaft, parse_steps, MATERIALS, fmt
from flange import Flange

COMMON = ["Обозначение", "Наименование", "Материал", "Фаски", "Ra", "Разработал"]
SHAFT_COLS = ["Обозначение", "Наименование", "Ступени (Ø×L через ;)", "Материал", "Фаски", "Ra",
              "Разработал"]
FLANGE_COLS = ["Обозначение", "Наименование", "D", "d", "H", "D1", "n", "d0", "Dст", "Lст",
               "Материал", "Фаски", "Ra", "Разработал"]
RESULT_COLS = ["Статус", "Масса, кг", "Файл PDF"]

HEAD = PatternFill("solid", fgColor="1F3B63")
OK = PatternFill("solid", fgColor="E2F0D9")
BAD = PatternFill("solid", fgColor="F8D7DA")
THIN = Border(*(Side(style="thin", color="BBBBBB"),) * 4)


# ------------------------------------------------------------------ шаблон
def make_template(path):
    wb = Workbook()
    examples = {
        "Валы": (SHAFT_COLS, [
            ["АБВГ.711111.001", "Вал", "20x100; 30x100; 40x50", "Сталь 45 ГОСТ 1050-2013", "авто", "6,3", ""],
            ["АБВГ.711111.003", "Вал", "25x40; 35x120; 30x60", "Сталь 40Х ГОСТ 4543-2016", "авто", "3,2", ""],
            ["АБВГ.711111.005", "Ось", "16x30; 20x150; 16x30", "Сталь 45 ГОСТ 1050-2013", "1", "6,3", ""],
        ], {"C": "Ступени слева направо: диаметр x длина, через точку с запятой.\nПример: 20x100; 30x100; 40x50"}),
        "Фланцы": (FLANGE_COLS, [
            ["АБВГ.711111.002", "Фланец", 160, 60, 20, 125, 4, 18, 0, 0, "Сталь 20 ГОСТ 1050-2013", "авто", "6,3", ""],
            ["АБВГ.711111.004", "Фланец", 200, 80, 24, 160, 8, 18, 120, 40, "Сталь 20 ГОСТ 1050-2013", "авто", "6,3", ""],
        ], {"C": "Наружный диаметр", "D": "Центральное отверстие", "E": "Толщина диска",
            "F": "Диаметр окружности центров отверстий", "G": "Число отверстий",
            "H": "Диаметр отверстий под крепёж", "I": "Диаметр ступицы (0 — без ступицы)",
            "J": "Длина ступицы"}),
    }
    first = True
    for title, (cols, rows, notes) in examples.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title
        ws.append(cols)
        for r in rows:
            ws.append(r)
        for i, c in enumerate(ws[1], 1):
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = HEAD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[c.column_letter].width = max(9, min(28, len(cols[i - 1]) + 4))
        for col, text in notes.items():
            ws[f"{col}1"].comment = Comment(text, "Генератор ЕСКД")
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"
        mat_col = cols.index("Материал") + 1
        letter = ws.cell(row=1, column=mat_col).column_letter
        ws.column_dimensions[letter].width = 28
        dv = DataValidation(type="list", formula1='"' + ",".join(MATERIALS) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{letter}2:{letter}500")
    wb.save(path)
    return path


# ------------------------------------------------------------------ чтение
def _num(v, name):
    if v is None or str(v).strip() == "":
        raise ValueError(f"не заполнено «{name}»")
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise ValueError(f"«{name}» должно быть числом, а не «{v}»")


def _common(row):
    ch = str(row.get("Фаски") or "авто").strip().replace(",", ".")
    mat = str(row.get("Материал") or "").strip()
    if mat not in MATERIALS:
        raise ValueError(f"неизвестный материал «{mat}» — выберите из списка")
    ra = row.get("Ra")
    ra = fmt(float(str(ra).replace(",", "."))) if ra not in (None, "") else "6,3"
    return dict(designation=str(row.get("Обозначение") or "").strip(),
                name=str(row.get("Наименование") or "").strip(),
                material=mat, chamfer=ch, ra=ra,
                developer=str(row.get("Разработал") or "").strip())


def read_excel(path):
    """-> список задач: {sheet, row, obj или error}. Пустые строки пропускаются."""
    wb = load_workbook(path, data_only=True)
    jobs = []
    for sheet, kind in (("Валы", "shaft"), ("Фланцы", "flange")):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        head = [str(c.value or "").strip() for c in ws[1]]
        for r in range(2, ws.max_row + 1):
            vals = [ws.cell(row=r, column=i + 1).value for i in range(len(head))]
            if all(v in (None, "") for v in vals[:3]):
                continue
            row = dict(zip(head, vals))
            job = {"sheet": sheet, "row": r, "label": f"{sheet}, строка {r}"}
            try:
                common = _common(row)
                if not common["designation"]:
                    raise ValueError("не заполнено «Обозначение»")
                if kind == "shaft":
                    steps = str(row.get("Ступени (Ø×L через ;)") or "").replace(";", "\n")
                    common["name"] = common["name"] or "Вал"
                    job["obj"] = Shaft(parse_steps(steps), **common)
                else:
                    common["name"] = common["name"] or "Фланец"
                    v = {k: _num(row.get(c), c) for k, c in
                         [("D", "D"), ("d", "d"), ("H", "H"), ("D1", "D1"), ("n", "n"), ("d0", "d0")]}
                    v["Dh"] = _num(row.get("Dст") or 0, "Dст")
                    v["Hh"] = _num(row.get("Lст") or 0, "Lст")
                    v["n"] = int(v["n"])
                    job["obj"] = Flange(**v, **common)
            except Exception as e:
                job["error"] = str(e)
            jobs.append(job)
    # одинаковые обозначения затёрли бы файлы друг друга
    seen = {}
    for j in jobs:
        if "obj" in j:
            d = j["obj"].designation
            if d in seen:
                j.pop("obj")
                j["error"] = f"обозначение {d} уже есть ({seen[d]})"
            else:
                seen[d] = j["label"]
    return jobs


# ------------------------------------------------------------------ запуск
def run(path, driver, out_dir, log=print, progress=None):
    jobs = read_excel(path)
    good = [j for j in jobs if "obj" in j]
    log(f"Найдено деталей: {len(jobs)} (к построению: {len(good)}, с ошибками в данных: {len(jobs) - len(good)})")
    for j in jobs:
        if "error" in j:
            log(f"  ✗ {j['label']}: {j['error']}")
    t0 = time.time()
    for i, j in enumerate(good, 1):
        o = j["obj"]
        log(f"[{i}/{len(good)}] {o.designation} — {o.summary()}")
        if progress:
            progress(i - 1, len(good))
        try:
            j["result"] = driver.build(o, out_dir)
            if j["result"].get("problems"):
                j["warn"] = "проверить: " + ", ".join(j["result"]["problems"])
        except Exception as e:
            j["error"] = f"ошибка САПР: {e}"
            log(f"  ✗ {e}")
    if progress:
        progress(len(good), len(good))
    secs = time.time() - t0
    report, book = write_report(path, jobs, out_dir, driver.name)
    done = sum(1 for j in jobs if "result" in j)
    log(f"Готово: {done} из {len(jobs)} за {secs:.0f} с (≈{secs / max(done, 1):.0f} с на деталь)")
    log(f"Отчёт: {report}")
    if book:
        log(f"Комплект PDF: {book}")
    return {"jobs": jobs, "report": report, "book": book, "done": done}


def write_report(src, jobs, out_dir, cad):
    """Копия исходной таблицы + колонки «Статус / Масса / Файл PDF», плюс общий PDF."""
    wb = load_workbook(src)
    for sheet in ("Валы", "Фланцы"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        base = ws.max_column + 1
        for k, title in enumerate(RESULT_COLS):
            c = ws.cell(row=1, column=base + k, value=title)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = HEAD
            ws.column_dimensions[c.column_letter].width = [34, 11, 60][k]
        for j in (j for j in jobs if j["sheet"] == sheet):
            r = j["row"]
            if "result" in j:
                st, fill = ("✓ " + j.get("warn", "готово")), OK
                ws.cell(row=r, column=base + 1, value=round(j["result"]["mass"], 2))
                ws.cell(row=r, column=base + 2, value=j["result"]["pdf"])
            else:
                st, fill = "✗ " + j.get("error", "?"), BAD
            ws.cell(row=r, column=base, value=st)
            for col in range(1, base + 3):
                ws.cell(row=r, column=col).fill = fill
                ws.cell(row=r, column=col).border = THIN
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    name = os.path.splitext(os.path.basename(src))[0]
    report = os.path.join(out_dir, f"{name} — отчёт {cad} {stamp}.xlsx")
    wb.save(report)

    pdfs = [j["result"]["pdf"] for j in jobs if "result" in j and os.path.exists(j["result"]["pdf"])]
    book = None
    if pdfs:
        try:
            import pymupdf
            doc = pymupdf.open()
            for p in pdfs:
                doc.insert_pdf(pymupdf.open(p))
            book = os.path.join(out_dir, f"{name} — комплект {cad} {stamp}.pdf")
            doc.save(book)
        except ImportError:
            pass
    return report, book


if __name__ == "__main__":
    cad, xlsx = sys.argv[1], sys.argv[2]
    if cad == "ks":
        from ks_driver import KompasDriver as D
    else:
        from sw_driver import SolidWorksDriver as D
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
    run(xlsx, D(), out)
