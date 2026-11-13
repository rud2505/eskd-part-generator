"""«Водитель» SolidWorks: превращает описание детали (вал, фланец) в модель,
чертёж по ЕСКД и PDF через API SolidWorks (COM)."""
import os, math
import pythoncom
from swlib import connect, cast, swconst
from shaft import MATERIALS, DENSITY, fmt, chamfered

C = swconst.constants
FONT = "GOST type B"
A3W, A3H = 0.420, 0.297
DRW_TPL = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\data\templates\gost.drwdot"
SHEET_FMT = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2026\lang\russian\sheetformat\a3 - gost_sh1_land.slddrt"
VIEW_CENTER = (0.200, 0.175)   # центр главного вида на листе, м
m = lambda v: v / 1000.0       # мм -> м


def g(o, n, *a):
    """Свойство или метод COM-объекта — без разницы, как он описан в справочнике."""
    x = getattr(o, n)
    return x(*a) if callable(x) else x


def kind(obj):
    return getattr(obj, "kind", "Вал")


class SolidWorksDriver:
    name = "SolidWorks"

    def __init__(self, log=print):
        self.log = log
        self.sw = None

    # ================================================================== общее
    def connect(self):
        pythoncom.CoInitialize()
        self.sw = connect()
        self.log(f"Подключились к SolidWorks {g(self.sw, 'RevisionNumber')}")

    def _mm_units(self, ext):
        ext.SetUserPreferenceInteger(C.swUnitSystem, 0, C.swUnitSystem_MMGS)
        ext.SetUserPreferenceInteger(C.swUnitsLinearDecimalPlaces, 0, 0)
        ext.SetUserPreferenceToggle(C.swDetailingDualDimensions, 0, False)

    def _props(self, doc, props):
        cpm = cast(doc.Extension.CustomPropertyManager(""), "CustomPropertyManager")
        for k, v in props.items():
            cpm.Add3(k, 30, v, 2)

    def _title_props(self, obj, mass):
        return {"Наименование": obj.name, "Обозначение": obj.designation, "Код_документа": "",
                "Материал": obj.material, "Масса": f"{mass:.2f}".replace(".", ","),
                "п_Разраб": obj.developer}

    def _close_and_remove(self, path):
        """Закрыть файл в SolidWorks (если открыт) и удалить старую версию."""
        d = self.sw.GetOpenDocumentByName(path)
        if d:
            title = g(cast(d, "ModelDoc2"), "GetTitle")
            self.sw.QuitDoc(title)
            self.sw.CloseDoc(title)
        if os.path.exists(path):
            try:
                os.remove(path)
            except PermissionError:
                raise RuntimeError(f"Файл занят (закройте его в SolidWorks): {path}")

    def build(self, obj, out_dir):
        if self.sw is None:
            self.connect()
        out_dir = os.path.join(out_dir, self.name)   # out\SolidWorks
        os.makedirs(out_dir, exist_ok=True)
        base = f"{obj.designation} {obj.name}".strip() or kind(obj)
        part = os.path.join(out_dir, base + ".SLDPRT")
        drw = os.path.join(out_dir, base + ".SLDDRW")
        pdf = os.path.join(out_dir, base + ".pdf")
        for p in (drw, part, pdf):
            self._close_and_remove(p)
        if kind(obj) == "Фланец":
            mass = self.flange_part(obj, part)
            problems = self.flange_drawing(obj, part, drw, pdf, mass)
        else:
            mass = self.shaft_part(obj, part)
            problems = self.shaft_drawing(obj, part, drw, pdf, mass)
        return {"part": part, "drawing": drw, "pdf": pdf, "mass": mass, "problems": problems}

    # ================================================================== модель: общие шаги
    def _new_part(self):
        sw = self.sw
        tpl = sw.GetUserPreferenceStringValue(C.swDefaultTemplatePart)
        doc = cast(sw.NewDocument(tpl, 0, 0, 0), "ModelDoc2")
        ext = cast(doc.Extension, "ModelDocExtension")
        self._mm_units(ext)
        sw.SetUserPreferenceToggle(C.swInputDimValOnCreate, False)
        return doc, ext, cast(doc.SketchManager, "SketchManager"), cast(doc.FeatureManager, "FeatureManager")

    def _plane(self, doc, index):
        """index: 0 — Спереди (XY), 1 — Сверху (XZ), 2 — Справа (YZ)."""
        f, i = cast(g(doc, "FirstFeature"), "Feature"), -1
        while f:
            if g(f, "GetTypeName2") == "RefPlane":
                i += 1
                if i == index:
                    doc.ClearSelection2(True)
                    f.Select2(False, 0)
                    return f
            nf = g(f, "GetNextFeature")
            f = cast(nf, "Feature") if nf else None
        raise RuntimeError("Не найдена опорная плоскость")

    def _select_last_feature(self, doc):
        doc.ClearSelection2(True)
        cast(doc.FeatureByPositionReverse(0), "Feature").Select2(False, 0)

    def _revolve(self, doc, sm, fm, pts, length):
        """Эскиз замкнутого контура на плоскости «Спереди» + вращение вокруг оси X."""
        self._plane(doc, 0)
        sm.InsertSketch(True)
        sm.AddToDB = True
        for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1]):
            sm.CreateLine(m(x1), m(y1), 0, m(x2), m(y2), 0)
        sm.CreateCenterLine(0, 0, 0, m(length), 0, 0)
        sm.AddToDB = False
        sm.InsertSketch(True)
        self._select_last_feature(doc)
        rev = fm.FeatureRevolve2(True, True, False, False, False, False, 0, 0,
                                 2 * math.pi, 0, False, False, 0, 0, 0, 0, 0, True, True, True)
        if rev is None:
            raise RuntimeError("SolidWorks не смог выполнить вращение эскиза")

    def _chamfers(self, doc, ext, fm, chamfers):
        """Каждая фаска — отдельной операцией: так SolidWorks не «теряет» кромки."""
        n, s45 = 0, math.sqrt(0.5)
        for ch in chamfers:
            doc.ClearSelection2(True)
            # точка на окружности под 45° — подальше от шва тела вращения (плоскость z = 0)
            if ext.SelectByID2("", "EDGE", m(ch.x), m(ch.r * s45), m(ch.r * s45), False, 0, None, 0):
                n += fm.InsertFeatureChamfer(4, 1, m(ch.c), math.pi / 4, 0, 0, 0, 0) is not None
        doc.ClearSelection2(True)
        self.log(f"  фаски: {n} из {len(chamfers)} {'✓' if n == len(chamfers) else '⚠'}")

    def _finish_part(self, doc, obj, path):
        cast(doc, "PartDoc").SetMaterialPropertyName2("", "solidworks materials", MATERIALS[obj.material])
        doc.ForceRebuild3(False)
        vol = cast(g(doc.Extension, "CreateMassProperty"), "MassProperty").Volume   # м³
        mass = vol * DENSITY[obj.material] * 1000                                   # кг
        self._props(doc, self._title_props(obj, mass))
        doc.SaveAs3(path, 0, 1)
        self.log(f"  модель сохранена, масса {fmt(round(mass, 2))} кг")
        return mass

    # ================================================================== чертёж: общие шаги
    def _new_drawing(self, obj, mass, sa, sb):
        d = self.sw.NewDocument(DRW_TPL, 0, A3W, A3H)
        doc, drw = cast(d, "ModelDoc2"), cast(d, "DrawingDoc")
        ext = cast(doc.Extension, "ModelDocExtension")
        self._mm_units(ext)
        ext.SetUserPreferenceToggle(C.swDetailingDimsShowParenthesisByDefault, 0, False)
        ext.SetUserPreferenceToggle(C.swDetailingAutoInsertCenterLines, 0, True)
        # подпись разреза по ГОСТ 2.305: только «А–А», без слова «Сечение» и масштаба
        ext.SetUserPreferenceToggle(C.swDetailingSectionViewLabels_PerStandard, 0, False)
        ext.SetUserPreferenceInteger(C.swDetailingSectionViewLabels_Name, 0, 0)
        ext.SetUserPreferenceInteger(C.swDetailingSectionViewLabels_Scale, 0, 2)
        for nm in [n for n in vars(C) if n.startswith("swDetailing") and n.endswith("TextFormat")]:
            try:
                tf = cast(ext.GetUserPreferenceTextFormat(getattr(C, nm), 0), "TextFormat")
                tf.TypeFaceName = FONT
                tf.Italic = True             # ГОСТ 2.304, тип Б с наклоном — как в КОМПАСе
                tf.CharHeight = 0.005 if nm == "swDetailingNoteTextFormat" else 0.0035
                ext.SetUserPreferenceTextFormat(getattr(C, nm), 0, tf)
            except Exception:
                pass
        # линии по ГОСТ 2.303: основная 0,5, тонкие 0,25
        ext.SetUserPreferenceInteger(C.swLineFontVisibleEdgesThickness, 0, C.swLW_CUSTOM)
        ext.SetUserPreferenceDouble(C.swLineFontVisibleEdgesThicknessCustom, 0, 0.0005)
        ext.SetUserPreferenceInteger(C.swLineFontDimensionsThickness, 0, C.swLW_CUSTOM)
        ext.SetUserPreferenceDouble(C.swLineFontDimensionsThicknessCustom, 0, 0.00025)
        self._props(doc, self._title_props(obj, mass))
        sheet = cast(g(drw, "GetCurrentSheet"), "Sheet")
        drw.SetupSheet5(g(sheet, "GetName"), 12, 12, sa, sb, True, SHEET_FMT,
                        A3W, A3H, "По умолчанию", True)
        # знак шероховатости рисуем сразу, пока на листе нет видов (иначе линии уйдут в активный вид)
        self._roughness_sign(doc, drw, obj.ra, sa / sb)
        return doc, drw, ext

    def _view(self, drw, part_path, orient, x, y, k):
        v = cast(drw.CreateDrawViewFromModelView3(part_path, orient, x, y, 0), "View")
        v.UseSheetScale = False
        v.ScaleDecimal = k
        v.SetDisplayMode3(False, 2, False, True)
        return v

    def _circles(self, doc, view, part_path):
        """Помощники выбора круговых кромок вида: сначала видимые, потом из модели."""
        selmgr = cast(doc.SelectionManager, "SelectionMgr")
        vis = []
        for e in view.GetVisibleEntities2(None, C.swViewEntityType_Edge) or []:
            cv = cast(g(cast(e, "Edge"), "GetCurve"), "Curve")
            if g(cv, "IsCircle"):
                p = cv.CircleParams
                vis.append((p[0] * 1000, p[6] * 1000, e))
        model = []
        pdoc = self.sw.GetOpenDocumentByName(part_path)
        if pdoc:
            body = cast(cast(pdoc, "PartDoc").GetBodies2(0, True)[0], "Body2")
            for e in g(body, "GetEdges") or []:
                cv = cast(g(cast(e, "Edge"), "GetCurve"), "Curve")
                if g(cv, "IsCircle"):
                    p = cv.CircleParams
                    ax = abs(p[3])   # ось окружности вдоль X — кромки тел вращения
                    model.append((p[0] * 1000, p[6] * 1000, e, ax > 0.99, p[1] * 1000, p[2] * 1000))

        def pick(cond, append=False):
            found = [c for c in vis if cond(c[0], c[1])]
            if found:
                x, r, e = max(found, key=lambda c: c[1])
                sd = cast(selmgr.CreateSelectData(), "SelectData")
                sd.View = view
                if cast(e, "Entity").Select4(append, sd):
                    return True
            for x, r, e, axial, *_ in model:
                if axial and cond(x, r):
                    return view.SelectEntity(e, append)
            return False

        def pick_hole(r0, angle, append=False):
            """Отверстие под крепёж, ближайшее к направлению angle (на виде слева:
            вверх — ось Y модели, по горизонтали — ось Z, узор симметричен)."""
            holes = [(e, cy, cz) for x, r, e, axial, cy, cz in model
                     if abs(r - r0) < 0.01 and (abs(cy) > 0.1 or abs(cz) > 0.1)]
            if not holes:
                return False
            ty, tz = math.sin(angle), abs(math.cos(angle))
            e = min(holes, key=lambda h: (h[1] / math.hypot(h[1], h[2]) - ty) ** 2 +
                    (abs(h[2]) / math.hypot(h[1], h[2]) - tz) ** 2)[0]
            return view.SelectEntity(e, append)
        return pick, pick_hole

    @staticmethod
    def _fix(dd, prefix=None):
        dd = cast(dd, "DisplayDimension")
        try:
            dd.ShowParenthesis = False
        except Exception:
            pass
        if prefix:
            dd.SetText(C.swDimensionTextPrefix, prefix)
        return dd

    def _leader(self, doc, text, attach, text_pt, shelf_right):
        """Выноска «подчёркнутый текст»: надпись стоит на полке (ГОСТ 2.307).
        attach, text_pt — точки листа в мм."""
        doc.ClearSelection2(True)
        n = doc.InsertNote("<FONT size=3.5mm>" + text)
        if not n:
            return False
        n = cast(n, "Note")
        w = 3.0 * len(text)
        tx = text_pt[0] if shelf_right else text_pt[0] - w
        n.SetTextPoint(m(tx), m(text_pt[1] + 5), 0)
        ann = cast(n.GetAnnotation(), "Annotation")
        ann.SetLeader3(C.swUNDERLINED, C.swLS_SMART, False, False, False, False)
        if ann.GetLeaderCount():
            ann.SetLeaderAttachmentPointAtIndex(0, m(attach[0]), m(attach[1]), 0)
        doc.ClearSelection2(True)
        return True

    # ------------------------------------------------------------------ \u0448\u0442\u0430\u043C\u043F \u043F\u043E \u0413\u041E\u0421\u0422
    def _title_block_gost(self, doc, drw):
        """\u0412\u0441\u0435 \u043D\u0430\u0434\u043F\u0438\u0441\u0438 \u043E\u0441\u043D\u043E\u0432\u043D\u043E\u0439 \u043D\u0430\u0434\u043F\u0438\u0441\u0438 \u2014 \u0448\u0440\u0438\u0444\u0442\u043E\u043C \u0413\u041E\u0421\u0422 \u0442\u0438\u043F \u0411; \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B (\u0433\u0440\u0430\u0444\u0430 3) \u2014 5 \u043C\u043C,
        \u0431\u0435\u0437 \u043F\u043E\u0434\u0447\u0451\u0440\u043A\u0438\u0432\u0430\u043D\u0438\u044F, \u043F\u043E \u0446\u0435\u043D\u0442\u0440\u0443 \u0433\u0440\u0430\u0444\u044B. \u0412 \u0448\u0430\u0431\u043B\u043E\u043D\u0435 SolidWorks \u043E\u043D\u0438 \u0431\u044B\u043B\u0438 Arial, \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B 2,5 \u043C\u043C."""
        drw.EditTemplate()
        sv = cast(g(drw, "GetFirstView"), "View")
        name_x, mat = None, None
        n = g(sv, "GetFirstNote")
        while n:
            n = cast(n, "Note")
            lt = g(n, "PropertyLinkedText") or ""
            ann = cast(n.GetAnnotation(), "Annotation")
            for i in range(max(ann.GetTextFormatCount(), 1)):
                tf = cast(ann.GetTextFormat(i), "TextFormat")
                tf.TypeFaceName = FONT
                tf.Italic = True             # \u0442\u0438\u043F \u0411 \u0441 \u043D\u0430\u043A\u043B\u043E\u043D\u043E\u043C 75\u00B0, \u043A\u0430\u043A \u0432 \u041A\u041E\u041C\u041F\u0410\u0421\u0435
                tf.Underline = False
                if '"\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B"' in lt:
                    tf.CharHeight = 0.005
                    tf.WidthFactor = 0.75
                ann.SetTextFormat(i, False, tf)
            if lt.strip() == '$PRP:"\u041D\u0430\u0438\u043C\u0435\u043D\u043E\u0432\u0430\u043D\u0438\u0435"':
                e = list(n.GetExtent())
                name_x = (e[0] + e[3]) / 2           # \u0446\u0435\u043D\u0442\u0440 \u0433\u0440\u0430\u0444\u044B 1 = \u0446\u0435\u043D\u0442\u0440 \u0433\u0440\u0430\u0444\u044B 3 \u043F\u043E \u0448\u0438\u0440\u0438\u043D\u0435
            if '"\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B"' in lt:
                mat = (n, ann)
            n = g(n, "GetNext")
        if mat and name_x:
            n, ann = mat
            n.SetTextJustification(C.swTextJustificationCenter)
            # \u0446\u0435\u043D\u0442\u0440 \u0433\u0440\u0430\u0444\u044B 3: \u043F\u043E X \u2014 \u043A\u0430\u043A \u0443 \u043D\u0430\u0438\u043C\u0435\u043D\u043E\u0432\u0430\u043D\u0438\u044F, \u043F\u043E Y \u2014 \u0441\u0435\u0440\u0435\u0434\u0438\u043D\u0430 \u043D\u0438\u0436\u043D\u0435\u0439 \u0441\u0442\u0440\u043E\u043A\u0438 \u0448\u0442\u0430\u043C\u043F\u0430 (5\u202620 \u043C\u043C).
            # \u041F\u0440\u0438 \u0432\u044B\u0440\u0430\u0432\u043D\u0438\u0432\u0430\u043D\u0438\u0438 \u043F\u043E \u0446\u0435\u043D\u0442\u0440\u0443 SetTextPoint \u0437\u0430\u0434\u0430\u0451\u0442 \u0446\u0435\u043D\u0442\u0440 \u043F\u043E \u0448\u0438\u0440\u0438\u043D\u0435 \u0438 \u0412\u0415\u0420\u0425 \u0442\u0435\u043A\u0441\u0442\u0430.
            n.SetTextPoint(name_x, 0.015, 0)
            doc.EditRebuild3()
            e = list(n.GetExtent())
            n.SetTextPoint(name_x, 0.0125 + (e[4] - e[1]) / 2, 0)
            doc.EditRebuild3()
        drw.EditSheet()

    # ------------------------------------------------------------------ \u0448\u0435\u0440\u043E\u0445\u043E\u0432\u0430\u0442\u043E\u0441\u0442\u044C
    def _roughness_sign(self, doc, drw, ra, k=1.0):
        """\u041D\u0435\u0443\u043A\u0430\u0437\u0430\u043D\u043D\u0430\u044F \u0448\u0435\u0440\u043E\u0445\u043E\u0432\u0430\u0442\u043E\u0441\u0442\u044C \u0432 \u043F\u0440\u0430\u0432\u043E\u043C \u0432\u0435\u0440\u0445\u043D\u0435\u043C \u0443\u0433\u043B\u0443 \u043F\u043E \u0413\u041E\u0421\u0422 2.309:
        \u0437\u043D\u0430\u043A \u221A \u0441 \u043F\u043E\u043B\u043A\u043E\u0439, \u043F\u043E\u0434 \u043F\u043E\u043B\u043A\u043E\u0439 \u00ABRa 6,3 (\u221A)\u00BB. \u0420\u0438\u0441\u0443\u0435\u0442\u0441\u044F \u043B\u0438\u043D\u0438\u044F\u043C\u0438 \u2014 \u0432 \u0448\u0440\u0438\u0444\u0442\u0435 \u0413\u041E\u0421\u0422 \u043D\u0435\u0442 \u0441\u0438\u043C\u0432\u043E\u043B\u0430 \u221A."""
        doc.ClearSelection2(True)
        sm = cast(doc.SketchManager, "SketchManager")
        t60 = math.tan(math.radians(60))

        def line(x1, y1, x2, y2):
            sm.AddToDB = True            # без «притягивания» концов к соседним объектам
            # эскиз листа SolidWorks живёт в масштабе листа: делим на k, чтобы попасть в мм бумаги
            s = sm.CreateLine(m(x1 / k), m(y1 / k), 0, m(x2 / k), m(y2 / k), 0)
            sm.AddToDB = False
            if s:
                cast(s, "SketchSegment").Width = C.swLW_THIN

        def note(text, x, top):
            nt = cast(doc.InsertNote(text), "Note")
            nt.SetTextPoint(m(x), m(top), 0)
            doc.ClearSelection2(True)
            return nt

        def ext(nt):
            e = [v * 1000 for v in nt.GetExtent()]
            return e[0], e[1], e[3], e[4]                 # x_left, y_bottom, x_right, y_top

        shelf_y, top = 287.0, 285.5                       # \u043F\u043E\u043B\u043A\u0430 \u0438 \u0432\u0435\u0440\u0445 \u0442\u0435\u043A\u0441\u0442\u0430, \u043C\u043C \u043B\u0438\u0441\u0442\u0430
        hs, Hs = 6.0, 12.0                                # \u043A\u043E\u0440\u043E\u0442\u043A\u043E\u0435 \u0438 \u0434\u043B\u0438\u043D\u043D\u043E\u0435 \u043F\u043B\u0435\u0447\u043E \u0437\u043D\u0430\u043A\u0430
        sm_h, sm_H = 2.5, 5.0                             # \u043C\u0430\u043B\u0435\u043D\u044C\u043A\u0438\u0439 \u0437\u043D\u0430\u043A \u0432 \u0441\u043A\u043E\u0431\u043A\u0430\u0445
        n1 = note(f"Ra {ra} (", 300, top)
        n2 = note(")", 300, top)
        doc.EditRebuild3()
        w1 = ext(n1)[2] - ext(n1)[0]
        w2 = ext(n2)[2] - ext(n2)[0]
        ws = sm_h / t60 + sm_H / t60
        Ls = 1.0 + w1 + 0.8 + ws + 0.8 + w2 + 1.0         # \u0434\u043B\u0438\u043D\u0430 \u043F\u043E\u043B\u043A\u0438 \u2014 \u043D\u0430\u0434 \u0432\u0441\u0435\u0439 \u043D\u0430\u0434\u043F\u0438\u0441\u044C\u044E
        tx = 405.0 - Ls                                   # \u043A\u043E\u043D\u0435\u0446 \u0434\u043B\u0438\u043D\u043D\u043E\u0433\u043E \u043F\u043B\u0435\u0447\u0430 (\u043D\u0430\u0447\u0430\u043B\u043E \u043F\u043E\u043B\u043A\u0438)
        vx, vy = tx - Hs / t60, shelf_y - Hs              # \u0432\u0435\u0440\u0448\u0438\u043D\u0430 \u0437\u043D\u0430\u043A\u0430
        line(vx - hs / t60, vy + hs, vx, vy)
        line(vx, vy, tx, shelf_y)
        line(tx, shelf_y, tx + Ls, shelf_y)
        n1.SetTextPoint(m(tx + 1.0), m(top), 0)
        doc.EditRebuild3()
        x1l, y1b, x1r, y1t = ext(n1)
        svx = x1r + 0.8 + sm_h / t60                      # \u0432\u0435\u0440\u0448\u0438\u043D\u0430 \u043C\u0430\u043B\u0435\u043D\u044C\u043A\u043E\u0433\u043E \u0437\u043D\u0430\u043A\u0430
        svy = y1b + 0.4
        line(svx - sm_h / t60, svy + sm_h, svx, svy)
        line(svx, svy, svx + sm_H / t60, svy + sm_H)
        n2.SetTextPoint(m(svx + sm_H / t60 + 0.8), m(top), 0)
        doc.ClearSelection2(True)

    def _finish_drawing(self, doc, drw, obj, sa, sb, drw_path, pdf_path, problems, extra_tt=()):
        self._title_block_gost(doc, drw)
        doc.ClearSelection2(True)
        items = ["Неуказанные предельные отклонения размеров:\n    H14, h14, \u00B1IT14/2.",
                 "Острые кромки притупить.", *extra_tt]
        tt = "\n".join(f"{i}. {t}" for i, t in enumerate(items, 1))
        n = doc.InsertNote(tt)
        if n:
            cast(n, "Note").SetTextPoint(0.240, 0.105 + 0.006 * len(extra_tt), 0)
        doc.ClearSelection2(True)
        sheet = cast(g(drw, "GetCurrentSheet"), "Sheet")
        sheet.SetProperties2(12, 12, sa, sb, True, A3W, A3H, True)
        self.sw.SetUserPreferenceToggle(C.swPDFExportInColor, False)
        doc.EditRebuild3()
        doc.ForceRebuild3(False)
        doc.SaveAs3(drw_path, 0, 1)
        doc.SaveAs3(pdf_path, 0, 2)
        self.log(f"  чертёж А3, масштаб {fmt(sa)}:{fmt(sb)} — сохранён")
        if problems:
            self.log("  ⚠ не удалось поставить: " + ", ".join(problems) + " — проверьте вручную")
        return problems

    # ================================================================== ВАЛ
    def shaft_part(self, shaft, path):
        self.log("Строю 3D-модель вала…")
        doc, ext, sm, fm = self._new_part()
        self._revolve(doc, sm, fm, chamfered(shaft.profile(), shaft.chamfers), shaft.length)
        return self._finish_part(doc, shaft, path)

    def shaft_drawing(self, shaft, part_path, drw_path, pdf_path, mass):
        self.log("Делаю чертёж по ЕСКД…")
        sa, sb = shaft.pick_scale()
        k = sa / sb
        doc, drw, ext = self._new_drawing(shaft, mass, sa, sb)
        view = self._view(drw, part_path, "*Спереди", *VIEW_CENTER, k)
        doc.ForceRebuild3(False)
        ol = list(g(view, "GetOutline"))
        X0 = (ol[0] + ol[2]) / 2 * 1000 - shaft.length * k / 2      # мм листа
        Y0 = (ol[1] + ol[3]) / 2 * 1000
        P = lambda x, y: (X0 + x * k, Y0 + y * k)                   # модель -> лист, мм
        pick, _ = self._circles(doc, view, part_path)
        problems, xs, R = [], shaft.x_edges(), shaft.dmax / 2 * k

        for i, (dia, l) in enumerate(shaft.steps):
            r, a, b = dia / 2, xs[i], xs[i + 1]
            doc.ClearSelection2(True)
            ok = pick(lambda x, rr: abs(rr - r) < 0.01 and a - 0.01 <= x <= b + 0.01)
            dd = doc.AddVerticalDimension2(m(P((a + b) / 2, 0)[0]), m(Y0), 0) if ok else None
            if dd:
                self._fix(dd, "<MOD-DIAM>")
            else:
                problems.append(f"Ø{fmt(dia)}")

        def length_dim(x2, dy):
            doc.ClearSelection2(True)
            a = pick(lambda x, r: abs(x) < 0.01)
            b = pick(lambda x, r: abs(x - x2) < 0.01, True)
            dd = doc.AddHorizontalDimension2(m(P(x2 / 2, 0)[0]), m(Y0 + dy), 0) if a and b else None
            if dd:
                self._fix(dd)
            else:
                problems.append(fmt(x2))

        for j, x2 in enumerate(xs[1:-1]):
            length_dim(x2, -(R + 12 + 9 * j))
        length_dim(xs[-1], R + 30)

        for j, ch in enumerate(shaft.chamfers):
            side = 1 if ch.x < shaft.length / 2 else -1
            ax, ay = P(*ch.mid)
            if not self._leader(doc, shaft.chamfer_text(ch), (ax, ay),
                                (ax + side * 8, Y0 + R + 9 + (j % 2) * 7), side > 0):
                problems.append(shaft.chamfer_text(ch))
        return self._finish_drawing(doc, drw, shaft, sa, sb, drw_path, pdf_path, problems)

    # ================================================================== ФЛАНЕЦ
    def flange_part(self, fl, path):
        self.log("Строю 3D-модель фланца…")
        doc, ext, sm, fm = self._new_part()
        self._revolve(doc, sm, fm, chamfered(fl.profile(), fl.chamfers), fl.length)
        # отверстия под крепёж: эскиз на плоскости «Справа» (YZ), вырез насквозь в обе стороны
        self._plane(doc, 2)
        sm.InsertSketch(True)
        sm.AddToDB = True
        for y, z in fl.holes():
            sm.CreateCircleByRadius(m(z), m(y), 0, m(fl.d0 / 2))
        sm.AddToDB = False
        sm.InsertSketch(True)
        self._select_last_feature(doc)
        cut = fm.FeatureCut4(False, False, False, 1, 1, 0, 0, False, False, False, False, 0, 0,
                             False, False, False, False, False, True, True, True, True, False,
                             0, 0, False, False)
        self.log(f"  отверстия {fl.n}×Ø{fmt(fl.d0)}: {'✓' if cut else '✗'}")
        return self._finish_part(doc, fl, path)

    def flange_drawing(self, fl, part_path, drw_path, pdf_path, mass):
        self.log("Делаю чертёж по ЕСКД (разрез + вид слева)…")
        sa, sb = fl.pick_scale()
        k = sa / sb
        doc, drw, ext = self._new_drawing(fl, mass, sa, sb)
        L, D = fl.length * k, fl.D * k                      # мм листа
        gap = 60
        left = 200 - (L + gap + D) / 2
        yc = 185          # выше — чтобы над штампом осталось место для техтребований
        xs_c = left + L / 2                                  # центр разреза
        xl_c = left + L + gap + D / 2                        # центр вида слева
        problems = []

        # 1) вид слева (окружность с отверстиями)
        lv = self._view(drw, part_path, "*Слева", m(xl_c), m(yc), k)
        doc.ForceRebuild3(False)
        ol = list(g(lv, "GetOutline"))
        cx, cy = (ol[0] + ol[2]) / 2 * 1000, (ol[1] + ol[3]) / 2 * 1000

        # 2) линия разреза через ось и отверстия (вертикальная), затем разрез А-А слева
        drw.ActivateView(g(lv, "GetName2"))
        sm = cast(doc.SketchManager, "SketchManager")
        doc.ClearSelection2(True)
        # внутри активного вида координаты эскиза — от центра вида, в размерах модели
        R0 = fl.D / 2 + 6 / k
        ln = sm.CreateLine(0, m(R0), 0, 0, -m(R0), 0)
        sv = drw.CreateSectionViewAt5(m(xs_c), m(yc), 0, "А", 0, None, 0) if ln else None
        if not sv:
            raise RuntimeError("Не удалось построить разрез")
        sv = cast(sv, "View")
        doc.ForceRebuild3(False)

        # 3) размеры на разрезе
        ol = list(g(sv, "GetOutline"))
        X0 = (ol[0] + ol[2]) / 2 * 1000 - L / 2
        Y0 = (ol[1] + ol[3]) / 2 * 1000
        P = lambda x, y: (X0 + x * k, Y0 + y * k)
        pick, _ = self._circles(doc, sv, part_path)
        R, r, H, Lm = fl.D / 2, fl.d / 2, fl.H, fl.length
        top = Y0 + R * k

        # подпись разреза: только «А–А» (без слова «Разрез» и масштаба), над размерами
        lab = g(sv, "GetFirstNote")
        if lab:
            lab = cast(lab, "Note")
            try:
                lab.PropertyLinkedText = "<VLLABEL>"
            except Exception:
                lab.SetText("<VLLABEL>")
            lab.SetTextPoint(m(X0 + L / 2 - 4), m(min(top + 40, 286)), 0)

        def vdim(cond, x_sheet, label):
            doc.ClearSelection2(True)
            dd = doc.AddVerticalDimension2(m(x_sheet), m(Y0), 0) if pick(cond) else None
            if dd:
                self._fix(dd, "<MOD-DIAM>")
            else:
                problems.append(label)

        def hdim(xa, xb, y_sheet, label):
            doc.ClearSelection2(True)
            a = pick(lambda x, rr: abs(x - xa) < 0.01)
            b = pick(lambda x, rr: abs(x - xb) < 0.01, True)
            dd = doc.AddHorizontalDimension2(m(P((xa + xb) / 2, 0)[0]), m(y_sheet), 0) if a and b else None
            if dd:
                self._fix(dd)
            else:
                problems.append(label)

        vdim(lambda x, rr: abs(rr - R) < 0.01, X0 - 14, f"Ø{fmt(fl.D)}")
        vdim(lambda x, rr: abs(rr - r) < 0.01, P(Lm / 2, 0)[0], f"Ø{fmt(fl.d)}")
        if fl.has_hub:
            vdim(lambda x, rr: abs(rr - fl.Dh / 2) < 0.01 and x > H, P(Lm, 0)[0] + 14, f"Ø{fmt(fl.Dh)}")
            hdim(0, H, top + 16, fmt(H))
            hdim(0, Lm, top + 26, fmt(Lm))
        else:
            hdim(0, H, top + 16, fmt(H))

        # фаски: наружные — вверх, входы отверстия — вниз
        for ch in fl.chamfers:
            mx, my = fl.chamfer_mid(ch)
            right = ch.side < 0
            if fl.is_bore(ch):
                ax, ay = P(mx, -my)
                tp = (ax + (8 if right else -8), Y0 - R * k - 8)
            else:
                ax, ay = P(mx, my)
                tp = (ax + (8 if right else -8), ay + 8)
            if not self._leader(doc, fl.chamfer_text(ch), (ax, ay), tp, right):
                problems.append(fl.chamfer_text(ch))

        # 4) вид слева: окружность центров отверстий ØD1 и «n отв. Ød0»
        lol = list(g(lv, "GetOutline"))
        lcx, lcy = (lol[0] + lol[2]) / 2 * 1000, (lol[1] + lol[3]) / 2 * 1000
        drw.ActivateView(g(lv, "GetName2"))
        doc.ClearSelection2(True)
        circ = sm.CreateCircleByRadius(0, 0, 0, m(fl.D1 / 2))
        if circ:
            doc.ClearSelection2(True)
            seg = cast(circ, "SketchSegment")
            seg.Select4(False, None)
            try:
                doc.SetLineStyle("CENTER")
            except Exception:
                pass
            doc.ClearSelection2(True)
            seg.Select4(False, None)
            a = math.radians(45)
            rr = fl.D1 / 2 * k
            dd = doc.AddDiameterDimension2(m(lcx + (rr + 18) * math.cos(a)), m(lcy + (rr + 18) * math.sin(a)), 0)
            if dd:
                self._fix(dd)
            else:
                problems.append(f"Ø{fmt(fl.D1)}")
        _, pick_hole = self._circles(doc, lv, part_path)
        doc.ClearSelection2(True)
        a = -math.pi / 4 if fl.n % 4 == 0 and fl.n > 4 else -math.pi / 2 + 2 * math.pi / fl.n * (fl.n // 4)
        a = math.atan2(math.sin(a), abs(math.cos(a)))      # отверстие справа-снизу
        if pick_hole(fl.d0 / 2, a):
            rr = fl.D / 2 * k
            dd = doc.AddDiameterDimension2(m(lcx + (rr + 10) * math.cos(a) + 10),
                                           m(lcy + (rr + 10) * math.sin(a)), 0)
            if dd:
                self._fix(dd, f"{fl.n} отв. <MOD-DIAM>")
            else:
                problems.append(f"{fl.n} отв.")
        else:
            problems.append(f"{fl.n} отв.")
        doc.ClearSelection2(True)
        return self._finish_drawing(doc, drw, fl, sa, sb, drw_path, pdf_path, problems)
