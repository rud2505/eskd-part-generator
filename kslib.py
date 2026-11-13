"""Подключение к КОМПАС-3D через COM. У КОМПАСа два API:
API5 (KompasObject) — 3D-моделирование и 2D-примитивы,
API7 (IApplication) — документы, виды, штамп, размеры."""
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import deps  # noqa: E402  — библиотеки из папки lib
import pythoncom
import win32com.client
from win32com.client import gencache

KDIR = r"C:\Program Files\ASCON\KOMPAS-3D v24\Bin"


def _load(tlb):
    t = pythoncom.LoadTypeLib(KDIR + "\\" + tlb)
    a = t.GetLibAttr()
    return gencache.EnsureModule(str(a[0]), a[1], a[3], a[4])


api5 = _load("kAPI5.tlb")
api7 = _load("kAPI7.tlb")
const = _load("ksConstants.tlb").constants
const3d = _load("ksConstants3D.tlb").constants


def connect(visible=True):
    raw5 = win32com.client.Dispatch("Kompas.Application.5")
    ko = api5.KompasObject(raw5._oleobj_.QueryInterface(api5.KompasObject.CLSID, pythoncom.IID_IDispatch))
    raw7 = win32com.client.Dispatch("Kompas.Application.7")
    app = api7.IApplication(raw7._oleobj_.QueryInterface(api7.IApplication.CLSID, pythoncom.IID_IDispatch))
    app.Visible = visible
    app.HideMessage = const.ksHideMessageNo   # не показывать диалоги «Сохранить?» и т.п.
    return ko, app


def q(obj, iface):
    """Привести объект к интерфейсу API7 (например 'IPart7')."""
    cls = getattr(api7, iface)
    o = getattr(obj, "_oleobj_", obj)
    return cls(o.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch))


def q5(obj, iface):
    cls = getattr(api5, iface)
    o = getattr(obj, "_oleobj_", obj)
    return cls(o.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch))
