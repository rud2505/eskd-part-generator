"""Проверка без окна:
python test_build.py sw 1 2 3     — валы в SolidWorks
python test_build.py ks f1 f2     — фланцы в КОМПАСе"""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import deps  # noqa: E402  — библиотеки из папки lib
import sys
from shaft import Shaft, parse_steps
from flange import Flange

cases = {
    "1": lambda p: Shaft(parse_steps("20x100\n30x100\n40x50"), designation=f"{p}.001"),
    "2": lambda p: Shaft(parse_steps("40x50\n25x30\n50x200\n30x120"), designation=f"{p}.002"),
    "3": lambda p: Shaft(parse_steps("16x40\n20x60\n24x30"), designation=f"{p}.003"),
    "f1": lambda p: Flange(designation=f"{p}.101"),
    "f2": lambda p: Flange(D=200, d=80, H=24, D1=160, n=8, d0=18, Dh=120, Hh=40,
                           designation=f"{p}.102"),
}
cad, *which = sys.argv[1:] or ["sw", "1"]
if cad == "ks":
    from ks_driver import KompasDriver as Driver
    prefix = "КОМП"
else:
    from sw_driver import SolidWorksDriver as Driver
    prefix = "ТЕСТ"
drv = Driver()
for c in which or ["1"]:
    obj = cases[c](prefix)
    obj.developer = "Иванов"
    print(obj.summary(), "масштаб", obj.pick_scale())
    res = drv.build(obj, r"D:\Claud\cadgen\out")
    print(res, getattr(drv, "debug", ""))
