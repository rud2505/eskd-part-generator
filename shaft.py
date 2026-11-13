"""«Мозг»: что такое ступенчатый вал. Ничего не знает про КОМПАС или SolidWorks,
только геометрию и правила ЕСКД. Все размеры — в миллиметрах."""
from dataclasses import dataclass, field

# Материалы: название для штампа -> материал из библиотеки SolidWorks (для массы)
MATERIALS = {
    "Сталь 45 ГОСТ 1050-2013": "AISI 1045 Сталь, холоднотянутая",
    "Сталь 20 ГОСТ 1050-2013": "AISI 1020",
    "Сталь 40Х ГОСТ 4543-2016": "AISI 1045 Сталь, холоднотянутая",  # плотность та же
    "12Х18Н10Т ГОСТ 5632-2014": "AISI 321 Отожженная нержавеющая сталь (SS)",
}

# Плотность, г/см³ — для массы в штампе (у аналогов из библиотеки SolidWorks она чуть другая)
DENSITY = {
    "Сталь 45 ГОСТ 1050-2013": 7.85,
    "Сталь 20 ГОСТ 1050-2013": 7.85,
    "Сталь 40Х ГОСТ 4543-2016": 7.85,
    "12Х18Н10Т ГОСТ 5632-2014": 7.90,
}

# Масштабы по ГОСТ 2.302 (уменьшения), по порядку предпочтения
SCALES = [(1, 1), (1, 2), (1, 2.5), (1, 4), (1, 5), (1, 10)]


def gost_chamfer(d):
    """Рекомендуемая фаска c×45° по диаметру (ГОСТ 10948)."""
    for dmax, c in [(10, 0.6), (30, 1.0), (50, 1.6), (80, 2.0), (120, 2.5)]:
        if d <= dmax:
            return c
    return 3.0


def chamfered(points, chamfers):
    """Срезать углы замкнутого контура под фаски c×45°: вершина (x, r) заменяется двумя
    точками на соседних сторонах. Фаски получаются прямо в эскизе — надёжнее операции «Фаска»."""
    marks = {(round(ch.x, 6), round(ch.r, 6)): ch.c for ch in chamfers}
    out, n = [], len(points)
    for i, (x, y) in enumerate(points):
        c = marks.get((round(x, 6), round(y, 6)))
        if not c:
            out.append((x, y))
            continue
        for px, py in (points[i - 1], points[(i + 1) % n]):
            d = ((px - x) ** 2 + (py - y) ** 2) ** 0.5
            pt = (x + (px - x) / d * c, y + (py - y) / d * c)
            if pt not in out:
                out.append(pt)
    return out


def fmt(v):
    """Число по-русски: 1.6 -> '1,6', 2.0 -> '2'."""
    s = f"{v:g}"
    return s.replace(".", ",")


@dataclass
class Chamfer:
    x: float        # где кромка (по оси вала)
    r: float        # радиус кромки
    c: float        # катет фаски
    side: int       # +1 — металл справа от кромки, -1 — слева

    @property
    def mid(self):
        """Середина образующей фаски (для выноски)."""
        return self.x + self.side * self.c / 2, self.r - self.c / 2


@dataclass
class Shaft:
    steps: list                     # [(диаметр, длина), ...] слева направо
    name: str = "Вал"
    designation: str = "АБВГ.711111.001"
    material: str = "Сталь 45 ГОСТ 1050-2013"
    developer: str = ""
    chamfer: object = "авто"        # "авто", 0 (без фасок) или число мм
    ra: str = "6,3"
    chamfers: list = field(default_factory=list, init=False)

    def __post_init__(self):
        self.steps = [(float(d), float(l)) for d, l in self.steps]
        self.validate()
        self.chamfers = self._make_chamfers()

    # ---------- проверки ----------
    def validate(self):
        if not self.steps:
            raise ValueError("Нужна хотя бы одна ступень")
        for i, (d, l) in enumerate(self.steps, 1):
            if d <= 0 or l <= 0:
                raise ValueError(f"Ступень {i}: диаметр и длина должны быть больше нуля")
        for i in range(len(self.steps) - 1):
            if self.steps[i][0] == self.steps[i + 1][0]:
                raise ValueError(f"Ступени {i+1} и {i+2} одного диаметра — объедините их")

    # ---------- геометрия ----------
    @property
    def length(self):
        return sum(l for _, l in self.steps)

    @property
    def dmax(self):
        return max(d for d, _ in self.steps)

    def x_edges(self):
        """Координаты торцов и уступов: [0, x1, x2, ..., L]."""
        xs, x = [0.0], 0.0
        for _, l in self.steps:
            x += l
            xs.append(x)
        return xs

    def profile(self):
        """Контур половины вала над осью (для эскиза вращения)."""
        pts, xs = [(0.0, 0.0)], self.x_edges()
        for i, (d, _) in enumerate(self.steps):
            pts += [(xs[i], d / 2), (xs[i + 1], d / 2)]
        pts.append((xs[-1], 0.0))
        return pts

    def _c(self, d):
        if self.chamfer in ("авто", None, ""):
            return gost_chamfer(d)
        return float(self.chamfer)

    def _make_chamfers(self):
        """Фаски на всех наружных кромках: оба торца и наружная кромка каждого уступа."""
        if self.chamfer not in ("авто", None, "") and float(self.chamfer) <= 0:
            return []
        xs, st, res = self.x_edges(), self.steps, []
        d0 = st[0][0]
        res.append(Chamfer(0.0, d0 / 2, self._c(d0), +1))
        for i in range(len(st) - 1):
            a, b = st[i][0], st[i + 1][0]
            big = max(a, b)
            res.append(Chamfer(xs[i + 1], big / 2, self._c(big), +1 if b > a else -1))
        dn = st[-1][0]
        res.append(Chamfer(xs[-1], dn / 2, self._c(dn), -1))
        # фаска не может быть больше высоты уступа / длины ступени
        for ch in res:
            ch.c = min(ch.c, ch.r * 0.5)
        return res

    def chamfer_text(self, ch):
        return f"{fmt(ch.c)}x45°"

    def pick_scale(self, field_w=300.0, field_h=150.0):
        """Подбор масштаба по ГОСТ 2.302, чтобы вал влез в поле листа А3."""
        for a, b in SCALES:
            k = a / b
            if self.length * k <= field_w and self.dmax * k <= field_h:
                return a, b
        return SCALES[-1]

    def summary(self):
        s = " – ".join(f"Ø{fmt(d)}×{fmt(l)}" for d, l in self.steps)
        return f"{self.name}: {s}, L={fmt(self.length)}"


def parse_steps(text):
    """Строки вида '20 100', '20x100', '20×100', 'Ø20 x 100' -> [(20,100), ...]."""
    import re
    steps = []
    for n, line in enumerate(text.strip().splitlines(), 1):
        line = line.strip()
        if not line:
            continue
        nums = re.findall(r"\d+(?:[.,]\d+)?", line)
        if len(nums) != 2:
            raise ValueError(f"Строка {n}: «{line}» — нужно два числа: диаметр и длина")
        steps.append(tuple(float(v.replace(",", ".")) for v in nums))
    return steps
