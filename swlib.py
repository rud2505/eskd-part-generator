import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import deps  # noqa: E402  — библиотеки из папки lib
import win32com.client, pythoncom
from win32com.client import gencache, CastTo

import pythoncom as _pc
SWDIR = r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS"
def _load(tlb):
    t = _pc.LoadTypeLib(SWDIR + "\\" + tlb)
    a = t.GetLibAttr()
    return gencache.EnsureModule(str(a[0]), a[1], a[3], a[4])
swmod = _load("sldworks.tlb")
swconst = _load("swconst.tlb")

def connect():
    raw = win32com.client.Dispatch("SldWorks.Application")
    sw = swmod.ISldWorks(raw._oleobj_)
    sw.Visible = True
    return sw

def cast(obj, name):
    if obj is None:
        return None
    cls = getattr(swmod, "I" + name, None) or getattr(swmod, name)
    o = getattr(obj, "_oleobj_", obj)
    o = o.QueryInterface(pythoncom.IID_IDispatch)
    return cls(o)

def ref_i4(v=0):
    return win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, v)

def nothing():
    return win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

