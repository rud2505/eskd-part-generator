"""«Мозг» фланца: круглый плоский фланец с центральным отверстием, отверстиями
под крепёж по окружности и необязательной ступицей. Все размеры — в мм.
Ось фланца — ось X, привалочная (опорная) плоскость — x = 0."""
import math
from dataclasses import dataclass, field
from shaft import Chamfer, gost_chamfer, fmt, SCALES

PART_TYPE = "Фланец"


@dataclass
class Flange:
    D: float = 160          # наружный диаметр
    d: float = 60           # центральное отверстие
    H: float = 20           # толщина диска
    D1: float = 125         # диаметр окружности центров отверстий
    n: int = 4              # число отверстий
    d0: float = 18          # диаметр отверстий под крепёж
    Dh: float = 0           # диаметр ступицы (0 — без ступицы)
    Hh: float = 0           # длина ступицы (выступ над диском)
    name: str = "Фланец"
    designation: str = "АБВГ.711111.002"
    material: str = "Сталь 20 ГОСТ 1050-2013"
    developer: str = ""
    chamfer: object = "авто"
    ra: str = "6,3"
    chamfers: list = field(default_factory=list, init=False)

    kind = PART_TYPE

    def __post_init__(self):
        for k in ("D", "d", "H", "D1", "d0", "Dh", "Hh"):
            setattr(self, k, float(getattr(self, k)))
        self.n = int(self.n)
        self.validate()
        self.chamfers = self._make_chamfers()

    # ---------- проверки: всё должно помещаться ----------
    def validate(self):
        D, d, D1, d0 = self.D, self.d, self.D1, self.d0
        if min(D, d, self.H, D1, d0) <= 0:
            raise ValueError("Все размеры должны быть больше нуля")
        if self.n < 2:
            raise ValueError("Отверстий должно быть не меньше 2")
        if d >= D:
            raise ValueError("Центральное отверстие d должно быть меньше D")
        if self.has_hub:
            if not d < self.Dh < D:
                raise ValueError("Диаметр ступицы должен быть между d и D")
            if self.Hh <= 0:
                raise ValueError("Укажите длину ступицы или поставьте её диаметр 0")
        inner = (self.Dh if self.has_hub else d) / 2
        if D1 / 2 - d0 / 2 < inner + 2:
            raise ValueError("Отверстия залезают на ступицу/центральное отверстие — увеличьте D1")
        if D1 / 2 + d0 / 2 > D / 2 - 2:
            raise ValueError("Отверстия выходят к краю фланца — уменьшите D1 или d0")
        gap = D1 * math.sin(math.pi / self.n) - d0
        if gap < 2:
            raise ValueError("Отверстия слишком близко друг к другу — уменьшите n или d0")

    # ---------- геометрия ----------
    @property
    def has_hub(self):
        return self.Dh > 0

    @property
    def length(self):
        return self.H + (self.Hh if self.has_hub else 0)

    @property
    def dmax(self):
        return self.D

    def profile(self):
        """Контур половины сечения над осью (для эскиза вращения)."""
        r, R = self.d / 2, self.D / 2
        if self.has_hub:
            Rh, L = self.Dh / 2, self.length
            return [(0, r), (0, R), (self.H, R), (self.H, Rh), (L, Rh), (L, r)]
        return [(0, r), (0, R), (self.H, R), (self.H, r)]

    def holes(self):
        """Центры отверстий (y, z) — первое отверстие сверху, на оси Y."""
        R1 = self.D1 / 2
        return [(R1 * math.cos(math.pi / 2 + 2 * math.pi * i / self.n),
                 R1 * math.sin(math.pi / 2 + 2 * math.pi * i / self.n)) for i in range(self.n)]

    def _c(self, d):
        if self.chamfer in ("авто", None, ""):
            return gost_chamfer(d)
        return float(self.chamfer)

    def _make_chamfers(self):
        """Фаски: наружные кромки диска, торец ступицы, входы центрального отверстия."""
        if self.chamfer not in ("авто", None, "") and float(self.chamfer) <= 0:
            return []
        R, r, H, L = self.D / 2, self.d / 2, self.H, self.length
        res = [Chamfer(0, R, self._c(self.D), +1),
               Chamfer(H, R, self._c(self.D), -1)]
        if self.has_hub:
            res.append(Chamfer(L, self.Dh / 2, self._c(self.Dh), -1))
        # фаски на входе в отверстие (у них «металл» снаружи, радиус растёт)
        res += [Chamfer(0, r, self._c(self.d), +1), Chamfer(L, r, self._c(self.d), -1)]
        # не больше 1/8 толщины диска, со значением из стандартного ряда
        std = [0.5, 0.6, 1.0, 1.6, 2.0, 2.5, 3.0]
        limit = min(self.H / 8, (R - r) / 6)
        for ch in res:
            ch.c = max([v for v in std if v <= min(ch.c, limit)] or [0.5])
        return res

    def chamfer_text(self, ch):
        return f"{fmt(ch.c)}x45°"

    def is_bore(self, ch):
        return abs(ch.r - self.d / 2) < 1e-6

    def chamfer_mid(self, ch):
        """Середина образующей фаски в сечении (для выноски)."""
        if self.is_bore(ch):
            return ch.x + ch.side * ch.c / 2, ch.r + ch.c / 2
        return ch.mid

    def pick_scale(self, field_w=330.0, field_h=135.0):
        """Два вида рядом: разрез (ширина L) и вид слева (ширина D)."""
        for a, b in SCALES:
            k = a / b
            if (self.length + self.D) * k + 60 <= field_w and self.D * k <= field_h:
                return a, b
        return SCALES[-1]

    def summary(self):
        hub = f", ступица Ø{fmt(self.Dh)}×{fmt(self.Hh)}" if self.has_hub else ""
        return (f"{self.name}: Ø{fmt(self.D)}/Ø{fmt(self.d)}×{fmt(self.H)}, "
                f"{self.n} отв. Ø{fmt(self.d0)} на Ø{fmt(self.D1)}{hub}")
