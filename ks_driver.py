"""«Водитель» КОМПАС-3D: превращает описание детали (вал, фланец) в модель .m3d,
ассоциативный чертёж .cdw по ЕСКД и PDF через API КОМПАСа (COM)."""
import os, math
import pythoncom
import win32com.client
from kslib import connect, q, q5, const, const3d, KDIR
from shaft import fmt, DENSITY, chamfered

DIAMETER_SIGN = 1              # IDimensionText.Sign: 1 — Ø (2 — это квадрат □)
VIEW_CENTER = (200.0, 175.0)   # центр главного вида на листе А3, мм


def kind(obj):
    return getattr(obj, "kind", "Вал")


class KompasDriver:
    name = "КОМПАС-3D"

    def __init__(self, log=print):
        self.log = log
        self.ko = self.app = None

    def connect(self):
        pythoncom.CoInitialize()
        self.ko, self.app = connect()
        self.log(f"Подключились к {self.app.ApplicationName(True)}")

    # ================================================================== файлы
    def _close_and_remove(self, path):
        docs = q(self.app.Documents, "IDocuments")
        for i in reversed(range(docs.Count)):
            d = q(docs.Item(i), "IKompasDocument")
            if os.path.normcase(d.PathName or "") == os.path.normcase(path):
                d.Close(const.kdDoNotSaveChanges)
        if os.path.exists(path):
            try:
                os.remove(path)
            except PermissionError:
                raise RuntimeError(f"Файл занят (закройте его): {path}")

    def build(self, obj, out_dir):
        if self.app is None:
            self.connect()
        out_dir = os.path.join(out_dir, self.name)   # out\КОМПАС-3D
        os.makedirs(out_dir, exist_ok=True)
        base = f"{obj.designation} {obj.name}".strip() or kind(obj)
        part = os.path.join(out_dir, base + ".m3d")
        drw = os.path.join(out_dir, base + ".cdw")
        pdf = os.path.join(out_dir, base + ".pdf")
        for p in (drw, part, pdf):
            self._close_and_remove(p)
        if kind(obj) == "Фланец":
            mass = self.flange_part(obj, part)
            problems = self.flange_drawing(obj, part, drw, pdf)
        else:
            mass = self.shaft_part(obj, part)
            problems = self.shaft_drawing(obj, part, drw, pdf)
        return {"part": part, "drawing": drw, "pdf": pdf, "mass": mass, "problems": problems}

    # ================================================================== модель: общие шаги
    def _new_part(self):
        doc = q5(self.ko.Document3D(), "ksDocument3D")
        doc.Create(False, True)
        return q5(doc.GetPart(const3d.pTop_Part), "ksPart")

    def _sketch(self, part, plane):
        sk = q5(part.NewEntity(const3d.o3d_sketch), "ksEntity")
        sdef = q5(sk.GetDefinition(), "ksSketchDefinition")
        sdef.SetPlane(part.GetDefaultEntity(plane))
        sk.Create()
        return sk, sdef, q5(sdef.BeginEdit(), "ksDocument2D")

    def _revolve(self, part, pts, length):
        """Замкнутый контур на XOY + осевая линия, вращение на 360°."""
        sk, sdef, d2 = self._sketch(part, const3d.o3d_planeXOY)
        for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1]):
            d2.ksLineSeg(x1, y1, x2, y2, 1)
        d2.ksLineSeg(-5, 0, length + 5, 0, 3)
        sdef.EndEdit()
        rot = q5(part.NewEntity(const3d.o3d_bossRotated), "ksEntity")
        rdef = q5(rot.GetDefinition(), "ksBossRotatedDefinition")
        rdef.SetSideParam(True, 360)
        rdef.SetSketch(sk)
        if not rot.Create():
            raise RuntimeError("КОМПАС не смог выполнить вращение эскиза")

    def _chamfers(self, part, chamfers):
        n = 0
        for ch in chamfers:
            ent = q5(part.NewEntity(const3d.o3d_chamfer), "ksEntity")
            cdef = q5(ent.GetDefinition(), "ksChamferDefinition")
            cdef.tangent = False
            cdef.SetChamferParam(True, ch.c, ch.c)
            edges = q5(part.EntityCollection(const3d.o3d_edge), "ksEntityCollection")
            s45 = math.sqrt(0.5)       # точка под 45° — подальше от шва тела вращения
            edges.SelectByPoint(ch.x, ch.r * s45, ch.r * s45)
            q5(cdef.array(), "ksEntityCollection").Add(edges.First())
            n += bool(ent.Create())
        self.log(f"  фаски: {n} из {len(chamfers)} {'✓' if n == len(chamfers) else '⚠'}")

    def _finish_part(self, obj, path):
        d7 = self.app.ActiveDocument
        tp = q(q(d7, "IKompasDocument3D").TopPart, "IPart7")
        tp.Marking = obj.designation
        tp.Name = obj.name
        tp.SetMaterial(obj.material, DENSITY[obj.material])
        tp.Update()
        mass = tp.Mass / 1000.0
        q(d7, "IKompasDocument").SaveAs(path)
        self.log(f"  модель сохранена, масса {fmt(round(mass, 2))} кг")
        return mass

    # ================================================================== чертёж: общие шаги
    def _new_drawing(self):
        d = self.app.Documents.Add(const.ksDocumentDrawing, True)
        kd = q(d, "IKompasDocument2D")
        sheet = q(kd.LayoutSheets.Item(0), "ILayoutSheet")
        fm = q(sheet.Format, "ISheetFormat")
        fm.Format = const.ksFormatA3
        fm.VerticalOrientation = False
        sheet.Update()
        return d, kd, sheet, q(kd.ViewsAndLayersManager.Views, "IViews")

    def _assoc_view(self, views, part_path, projection, k, x, y, view_type=None):
        av = q(views.Add(view_type if view_type is not None else const.vt_Arbitrary), "IAssociationView")
        if projection:
            av.SourceFileName = part_path
            av.ProjectionName = projection
        av.Scale = k
        av.X, av.Y = x, y
        av.HiddenLinesVisible = False
        vd = q(av, "IViewDesignation")
        vd.ShowName = False
        vd.ShowScale = False
        return av

    @staticmethod
    def _proj(av):
        """Точка модели (мм) -> координаты вида."""
        def V(x, y, z=0.0):
            r = av.GetProjectionPoint(x, y, z)
            return (r[1], r[2]) if isinstance(r, tuple) and len(r) >= 3 else (x, y)
        return V

    @staticmethod
    def _lin(sym, p1, p2, p3, orient, diam=False):
        ld = q(sym.LineDimensions.Add(), "ILineDimension")
        ld.X1, ld.Y1 = p1
        ld.X2, ld.Y2 = p2
        ld.X3, ld.Y3 = p3
        ld.Orientation = orient
        if diam:
            q(ld, "IDimensionText").Sign = DIAMETER_SIGN
        return ld.Update()

    @staticmethod
    def _leader(sym, text, attach, shelf_start, right):
        bl = sym.Leaders.Add(const.ksDrLeader)
        br = q(bl, "IBranchs")
        br.X0, br.Y0 = shelf_start
        br.AddBranchByPoint(-1, *attach)
        ld = q(bl, "ILeader")
        ld.ShelfDirection = const.ksLSRight if right else const.ksLSLeft
        q(ld.TextOnShelf, "IText").Str = text.replace("x", "×")
        base = q(bl, "IBaseLeader")
        base.ArrowType = const.ksLeaderArrow
        return base.Update()

    @staticmethod
    def _axis(dc, x1, y1, x2, y2):
        ls = q(dc.LineSegments.Add(), "ILineSegment")
        ls.X1, ls.Y1, ls.X2, ls.Y2 = x1, y1, x2, y2
        ls.Style = const.ksCSAxial
        ls.Update()

    def _finish_drawing(self, d, sheet, obj, sa, sb, drw_path, pdf_path, problems):
        dd = q(d, "IDrawingDocument")
        sr = q(dd.SpecRough, "ISpecRough")
        sr.AddSign = True
        sr.Text = f"Ra {obj.ra}"
        sr.Update()
        td = q(dd.TechnicalDemand, "ITechnicalDemand")
        t = q(td.Text, "IText")
        for line in ["Неуказанные предельные отклонения размеров: H14, h14, ±IT14/2.",
                     "Острые кромки притупить."]:
            tl = q(t.Add(), "ITextLine")
            tl.Str = line
            tl.Numbering = 1
        td.Update()
        st = q(sheet.Stamp, "IStamp")     # наименование, обозначение, материал, масса — из модели
        if obj.developer:
            q(st.Text(110), "IText").Str = obj.developer
        st.Update()
        q(d, "IKompasDocument").SaveAs(drw_path)
        cv = q(self.app.Converter(KDIR + r"\Pdf2d.dll"), "IConverter")
        cv.Convert(drw_path, pdf_path, 0, False)
        self.log(f"  чертёж А3, масштаб {fmt(sa)}:{fmt(sb)} — сохранён")
        if problems:
            self.log("  ⚠ не удалось поставить: " + ", ".join(problems) + " — проверьте вручную")
        return problems

    # ================================================================== ВАЛ
    def shaft_part(self, shaft, path):
        self.log("Строю 3D-модель вала…")
        part = self._new_part()
        self._revolve(part, chamfered(shaft.profile(), shaft.chamfers), shaft.length)
        return self._finish_part(shaft, path)

    def shaft_drawing(self, shaft, part_path, drw_path, pdf_path):
        self.log("Делаю чертёж по ЕСКД…")
        sa, sb = shaft.pick_scale()
        k = sa / sb
        d, kd, sheet, views = self._new_drawing()
        av = self._assoc_view(views, part_path, "#Спереди", k,
                              VIEW_CENTER[0] - shaft.length * k / 2, VIEW_CENTER[1])
        av.Update()
        sym, dc = q(av, "ISymbols2DContainer"), q(av, "IDrawingContainer")
        problems = []
        self._axis(dc, -5 / k, 0, shaft.length + 5 / k, 0)

        xs, R = shaft.x_edges(), shaft.dmax / 2
        for i, (dia, l) in enumerate(shaft.steps):
            xm = (xs[i] + xs[i + 1]) / 2
            if not self._lin(sym, (xm, dia / 2), (xm, -dia / 2), (xm, 0), const.ksLinDVertical, True):
                problems.append(f"Ø{fmt(dia)}")
        for j, x2 in enumerate(xs[1:-1]):
            y2 = -max(shaft.steps[j][0], shaft.steps[j + 1][0]) / 2
            if not self._lin(sym, (0, -shaft.steps[0][0] / 2), (x2, y2),
                             (x2 / 2, -(R + (12 + 9 * j) / k)), const.ksLinDHorizontal):
                problems.append(fmt(x2))
        if not self._lin(sym, (0, shaft.steps[0][0] / 2), (shaft.length, shaft.steps[-1][0] / 2),
                         (shaft.length / 2, R + 30 / k), const.ksLinDHorizontal):
            problems.append(fmt(shaft.length))

        for j, ch in enumerate(shaft.chamfers):
            mx, my = ch.mid
            right = ch.x < shaft.length / 2
            if not self._leader(sym, shaft.chamfer_text(ch), (mx, my),
                                (mx + (8 if right else -8) / k, R + (9 + (j % 2) * 7) / k), right):
                problems.append(shaft.chamfer_text(ch))
        return self._finish_drawing(d, sheet, shaft, sa, sb, drw_path, pdf_path, problems)

    # ================================================================== ФЛАНЕЦ
    def flange_part(self, fl, path):
        self.log("Строю 3D-модель фланца…")
        part = self._new_part()
        self._revolve(part, chamfered(fl.profile(), fl.chamfers), fl.length)
        # отверстия: эскиз на плоскости YOZ, вырез насквозь в обе стороны
        sk, sdef, d2 = self._sketch(part, const3d.o3d_planeYOZ)
        for y, z in fl.holes():
            d2.ksCircle(y, z, fl.d0 / 2, 1)
        sdef.EndEdit()
        cut = q5(part.NewEntity(const3d.o3d_cutExtrusion), "ksEntity")
        cdef = q5(cut.GetDefinition(), "ksCutExtrusionDefinition")
        cdef.directionType = const3d.dtBoth
        cdef.SetSideParam(True, const3d.etThroughAll, 0, 0, False)
        cdef.SetSideParam(False, const3d.etThroughAll, 0, 0, False)
        cdef.SetSketch(sk)
        self.log(f"  отверстия {fl.n}×Ø{fmt(fl.d0)}: {'✓' if cut.Create() else '✗'}")
        return self._finish_part(fl, path)

    def flange_drawing(self, fl, part_path, drw_path, pdf_path):
        self.log("Делаю чертёж по ЕСКД (разрез + вид слева)…")
        sa, sb = fl.pick_scale()
        k = sa / sb
        d, kd, sheet, views = self._new_drawing()
        Ls, Ds, gap = fl.length * k, fl.D * k, 60
        left = 200 - (Ls + gap + Ds) / 2
        yc = 185
        problems = []

        # 1) вид слева
        lv = self._assoc_view(views, part_path, "#Слева", k, left + Ls + gap + Ds / 2, yc)
        lv.Update()
        Vl = self._proj(lv)
        lsym, ldc = q(lv, "ISymbols2DContainer"), q(lv, "IDrawingContainer")

        # 2) линия разреза через ось и отверстия + разрез на месте главного вида
        R = fl.D / 2
        top, bot = Vl(0, R + 6 / k), Vl(0, -R - 6 / k)
        # точки линии разреза КОМПАС берёт из массива Points (иначе — «вырожденный объект»)
        # ГОСТ 2.305: разрез по плоскости симметрии в проекционной связи не обозначают —
        # линию разреза кладём на невидимый и непечатаемый слой, подпись «А–А» скрываем
        hidden = q(q(lv.Layers, "ILayers").Add(), "ILayer")
        hidden.Name = "Линия разреза (скрыта по ГОСТ 2.305)"
        hidden.Visible = False
        hidden.Printable = False
        hidden.Update()
        cl = q(lsym.CutLines.Add(), "ICutLine")
        cl.Points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [*top, *bot])
        cl.LayerNumber = hidden.LayerNumber
        if not cl.Update():
            raise RuntimeError("КОМПАС: не удалось создать линию разреза")
        sv = q(views.Add(const.vt_Section), "IAssociationView")
        sv.SourceFileName = part_path
        sv.BaseObject = cl
        sv.Section = False                      # разрез (не сечение)
        sv.Scale = k
        sv.X, sv.Y = left + Ls / 2, yc
        q(sv, "IViewDesignation").ShowScale = False
        q(sv, "IViewDesignation").ShowName = False
        if not sv.Update():
            raise RuntimeError("КОМПАС: не удалось построить разрез")
        # центр разреза — посередине длины фланца
        V = self._proj(sv)
        o, e = V(0, 0), V(fl.length, 0)
        sv.X += (left + Ls / 2) - (sv.X + (o[0] + e[0]) / 2 * k)
        sv.Update()
        V = self._proj(sv)
        s = 1 if V(fl.length, 0)[0] > V(0, 0)[0] else -1     # куда смотрит ось X на разрезе
        sym = q(sv, "ISymbols2DContainer")

        # 3) размеры на разрезе (в координатах вида; смещения на листе делим на масштаб)
        r, H, Lm = fl.d / 2, fl.H, fl.length
        u0 = V(0, 0)[0]
        cR = next((ch.c for ch in fl.chamfers if ch.r == R and ch.x == 0), 0)
        self._lin(sym, V(cR, R), V(cR, -R), (u0 - s * 14 / k, 0), const.ksLinDVertical, True) \
            or problems.append(f"Ø{fmt(fl.D)}")
        self._lin(sym, V(Lm / 2, r), V(Lm / 2, -r), (V(Lm / 2, 0)[0], 0), const.ksLinDVertical, True) \
            or problems.append(f"Ø{fmt(fl.d)}")
        if fl.has_hub:
            Rh = fl.Dh / 2
            self._lin(sym, V(Lm - 3, Rh), V(Lm - 3, -Rh), (V(Lm, 0)[0] + s * 14 / k, 0),
                      const.ksLinDVertical, True) or problems.append(f"Ø{fmt(fl.Dh)}")
            self._lin(sym, V(0, R - cR), V(Lm, Rh), (V(Lm / 2, 0)[0], R + 26 / k),
                      const.ksLinDHorizontal) or problems.append(fmt(Lm))
        self._lin(sym, V(0, R - cR), V(H, R - cR), (V(H / 2, 0)[0], R + 16 / k),
                  const.ksLinDHorizontal) or problems.append(fmt(H))

        # фаски: наружные — выноской вверх, входы отверстия — вниз
        for ch in fl.chamfers:
            mx, my = fl.chamfer_mid(ch)
            right = (ch.side < 0) == (s > 0)
            dx = (8 if right else -8) / k
            if fl.is_bore(ch):
                a = V(mx, -my)
                shelf = (a[0] + dx, -(R + 8 / k))
            else:
                a = V(mx, my)
                shelf = (a[0] + dx, a[1] + 8 / k)
            self._leader(sym, fl.chamfer_text(ch), a, shelf, right) or problems.append(fl.chamfer_text(ch))

        # 4) вид слева: окружность центров отверстий (осевая) с ØD1 и «n отв. Ød0»
        circ = q(ldc.Circles.Add(), "ICircle")
        circ.Xc, circ.Yc = Vl(0, 0)
        circ.Radius = fl.D1 / 2
        circ.Style = const.ksCSAxial
        circ.Update()
        dd = q(lsym.DiametralDimensions.Add(), "IDiametralDimension")
        dd.Xc, dd.Yc = Vl(0, 0)
        dd.Radius = fl.D1 / 2
        dd.Angle = 45
        dd.Update() or problems.append(f"Ø{fmt(fl.D1)}")
        hy, hz = min(fl.holes(), key=lambda h: (h[0] + h[1] * 0.3))      # нижнее-правое
        hc = Vl(0, hy, hz)
        hd = q(lsym.DiametralDimensions.Add(), "IDiametralDimension")
        hd.Xc, hd.Yc = hc
        hd.Radius = fl.d0 / 2
        hd.Angle = -45
        q(q(hd, "IDimensionText").Prefix, "ITextLine").Str = f"{fl.n} отв. "
        hd.Update() or problems.append(f"{fl.n} отв.")
        return self._finish_drawing(d, sheet, fl, sa, sb, drw_path, pdf_path, problems)
