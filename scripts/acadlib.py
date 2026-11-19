"""Thu vien helper dieu khien AutoCAD qua COM cho script viet tay (khong phai MCP tool).

Chay bang Python cua .venv trong thu muc repo (da co pywin32):

    & "<repo>\\.venv\\Scripts\\python.exe" my_script.py

Trong script:

    import sys; sys.path.insert(0, r"<repo>\\scripts")
    import acadlib as A
    app, doc = A.connect("ten-ban-ve.dwg")
    with A.undo_group(doc):
        A.setup_sheet_a3(doc, "Layout1", "DN-01", {"SO_BAN_VE": "DN-01", ...})

Cac loi COM da gap va cach tranh duoc gom o day (xem tung ham):
- sysvar kieu short phai truyen VT_I2            -> set_var()
- setattr loi khi AutoCAD ban                    -> S(), retry()
- SendCommand bat dong bo                        -> send()
- Layout.StandardScale (AcPlotScale, 1:1 = 16) khac PViewport.StandardScale (1:1 = 2)
  -> viewport dung CustomScale
- Layout PlotType = acLayout thi khong gan CenterPlot duoc
- print tieng Viet tren console cp1252           -> stdout doi sang utf-8 khi import
"""
import contextlib
import os
import sys
import time

import pythoncom
import pywintypes
import win32com.client
from win32com.client import VARIANT

for _s in (sys.stdout, sys.stderr):
    with contextlib.suppress(AttributeError, ValueError):
        _s.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = os.path.normpath(os.path.join(HERE, "..", "templates"))
KHUNG_A3 = os.path.join(TEMPLATES, "RIK-KHUNG-A3.dwg")

# Khung A3 ngang: mep giay 420x297, khung trong le trai 20 / con lai 5, khung ten 180x55 goc phai duoi.
# Bien viewport thut 2 mm vao trong khung va chua goc khung ten.
A3_VIEWPORT_CLIP = [(22, 7), (233, 7), (233, 62), (413, 62), (413, 290), (22, 290)]
A3_TAGS = ("CHU_TRI", "THIET_KE", "VE", "KIEM_TRA", "PHE_DUYET", "CHU_DAU_TU", "DU_AN",
           "HANG_MUC", "TEN_BAN_VE", "TY_LE", "NGAY", "SO_BAN_VE", "LAN_SUA")

ALIGN = {"L": 0, "C": 1, "R": 2, "M": 4, "TL": 6, "TC": 7, "TR": 8,
         "ML": 9, "MC": 10, "MR": 11, "BL": 12, "BC": 13, "BR": 14}
LW = {"0.13": 13, "0.18": 18, "0.25": 25, "0.35": 35, "0.50": 50, "0.70": 70}


# ------------------------------------------------------------------ ket noi, retry
RETRY_MS = 60000
_filter = None  # giu tham chieu de ctypes khong giai phong vtable


def _register_filter():
    """IMessageFilter: AutoCAD ban (dang nap file, regen, chay lenh) thi COM tra 'Call was rejected
    by callee' - ke ca o buoc tra ten phuong thuc, lam win32com dynamic nem AttributeError.
    Filter nay bao COM tu goi lai moi 100 ms trong toi da RETRY_MS, ap dung cho MOI loi goi tu
    luong nay. pywin32 khong boc CoRegisterMessageFilter nen dung ctypes dung vtable tay."""
    global _filter
    if _filter is not None:
        return
    import ctypes
    from ctypes import POINTER, WINFUNCTYPE, Structure, c_long, c_ulong, c_void_p, byref

    iid_unknown = "{00000000-0000-0000-C000-000000000046}"
    iid_filter = "{00000016-0000-0000-C000-000000000046}"

    class GUID(Structure):
        _fields_ = [("d1", c_ulong), ("d2", ctypes.c_ushort), ("d3", ctypes.c_ushort),
                    ("d4", ctypes.c_ubyte * 8)]

        def __str__(self):
            b = bytes(self.d4)
            return "{%08X-%04X-%04X-%s-%s}" % (self.d1, self.d2, self.d3, b[:2].hex().upper(),
                                               b[2:].hex().upper())

    QI = WINFUNCTYPE(c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))
    REF = WINFUNCTYPE(c_ulong, c_void_p)
    HIC = WINFUNCTYPE(c_ulong, c_void_p, c_ulong, c_void_p, c_ulong, c_void_p)
    RRC = WINFUNCTYPE(c_ulong, c_void_p, c_void_p, c_ulong, c_ulong)
    MP = WINFUNCTYPE(c_ulong, c_void_p, c_void_p, c_ulong, c_ulong)

    def query_interface(this, riid, ppv):
        if str(riid.contents) in (iid_unknown, iid_filter):
            ppv[0] = this
            return 0
        ppv[0] = None
        return c_long(0x80004002).value  # E_NOINTERFACE

    def retry_rejected(this, task, tick, reject_type):
        # SERVERCALL_RETRYLATER = 2 -> goi lai sau 100 ms; -1 (0xFFFFFFFF) = bo cuoc
        return 100 if reject_type == 2 and tick < RETRY_MS else 0xFFFFFFFF

    class VTable(Structure):
        _fields_ = [("qi", QI), ("addref", REF), ("release", REF),
                    ("hic", HIC), ("rrc", RRC), ("mp", MP)]

    class Obj(Structure):
        _fields_ = [("vtbl", POINTER(VTable))]

    vt = VTable(QI(query_interface), REF(lambda this: 1), REF(lambda this: 1),
                HIC(lambda this, ct, task, tick, info: 0),   # SERVERCALL_ISHANDLED
                RRC(retry_rejected),
                MP(lambda this, task, tick, pt_: 2))          # PENDINGMSG_WAITDEFPROCESS
    obj = Obj(ctypes.pointer(vt))
    old = c_void_p()
    hr = ctypes.oledll.ole32.CoRegisterMessageFilter(byref(obj), byref(old))
    _filter = (vt, obj, hr)


def connect(doc_name=None):
    """Tra ve (app, doc). doc_name: chi chap nhan dung ban ve nay (tranh ve nham file).
    Dang ky message filter de moi loi goi COM tu cho khi AutoCAD ban."""
    pythoncom.CoInitialize()
    _register_filter()
    app = win32com.client.GetActiveObject("AutoCAD.Application")
    if app.Documents.Count == 0:
        raise RuntimeError("AutoCAD chua mo ban ve nao")
    doc = app.ActiveDocument
    if doc_name:
        if doc.Name.lower() != doc_name.lower():
            for d in app.Documents:
                if d.Name.lower() == doc_name.lower():
                    return app, d
            raise RuntimeError(f"Khong thay ban ve {doc_name}; dang mo: {doc.Name}")
    return app, doc


def retry(fn, *args, tries=40, delay=0.25):
    """Goi lai khi AutoCAD ban (com_error 'call rejected' hoac AttributeError cua dynamic dispatch)."""
    last = None
    for _ in range(tries):
        try:
            return fn(*args)
        except (pywintypes.com_error, AttributeError) as exc:
            last = exc
            time.sleep(delay)
    raise last


def S(obj, attr, value):
    """setattr co retry."""
    retry(setattr, obj, attr, value)


def call(obj, method, *args):
    """Goi phuong thuc co retry ca buoc lay thuoc tinh: doi tuong vua tao (vd Documents.Add) hay
    AutoCAD dang ban thi win32com dynamic nem AttributeError ngay khi tra ten phuong thuc."""
    return retry(lambda: getattr(obj, method)(*args))


def pt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def pt2(x, y):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y)))


def dbl_array(values):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values])


def obj_array(objs):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, list(objs))


def set_var(doc, name, value):
    """SetVariable; so nguyen thu VT_I2 truoc (sysvar short) roi moi VT_I4."""
    if isinstance(value, bool) or not isinstance(value, int):
        return call(doc, "SetVariable", name, value)
    try:
        return call(doc, "SetVariable", name, VARIANT(pythoncom.VT_I2, value))
    except pywintypes.com_error:
        return call(doc, "SetVariable", name, VARIANT(pythoncom.VT_I4, value))


@contextlib.contextmanager
def temp_var(doc, name, value):
    old = doc.GetVariable(name)
    set_var(doc, name, value)
    try:
        yield
    finally:
        set_var(doc, name, old)


@contextlib.contextmanager
def undo_group(doc):
    """Gom moi thay doi thanh mot buoc Undo (nguoi dung go U mot lan la hoan tac)."""
    doc.StartUndoMark()
    try:
        yield
    finally:
        retry(doc.EndUndoMark)


def wait_idle(doc, timeout=30.0):
    """Cho AutoCAD chay xong lenh (CMDNAMES rong). Trong luc lenh chay, duyet collection se loi."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            if doc.GetVariable("CMDNAMES") == "":
                return
        except (pywintypes.com_error, AttributeError):
            pass
        time.sleep(0.2)
    raise TimeoutError("AutoCAD van dang chay lenh sau %.0f s" % timeout)


def send(doc, command, timeout=30.0):
    """SendCommand roi cho xong. Nho ket thuc chuoi lenh bang dau cach/xuong dong."""
    wait_idle(doc, timeout)
    doc.SendCommand(command)
    time.sleep(0.2)
    wait_idle(doc, timeout)


def handent(ent):
    """Bieu thuc LISP chon doi tuong theo handle, dung trong SendCommand."""
    return f'(handent "{ent.Handle}")'


# ------------------------------------------------------------------ bang ky hieu
def ensure_layer(doc, name, color=7, plottable=True, lineweight=None):
    lay = retry(doc.Layers.Add, name)  # Add tra ve layer co san neu trung ten
    S(lay, "color", color)
    S(lay, "Plottable", plottable)
    if lineweight is not None:
        S(lay, "Lineweight", lineweight)
    return lay


def ensure_text_style(doc, name, font="Arial", bold=False, italic=False):
    """Text style TTF (SHX khong hien duoc tieng Viet co dau)."""
    st = retry(doc.TextStyles.Add, name)
    retry(st.SetFont, font, bold, italic, 0, 34)
    S(st, "Height", 0.0)
    return st


def true_color(app, rgb):
    major = app.Version.split(".")[0]
    clr = app.GetInterfaceObject(f"AutoCAD.AcCmColor.{major}")
    clr.SetRGB(*rgb)
    return clr


def find_block(doc, name):
    for b in doc.Blocks:
        if b.Name.upper() == name.upper():
            return b
    return None


# ------------------------------------------------------------------ ve (space = ModelSpace, PaperSpace hoac Block)
def _finish(ent, layer=None, lw=None):
    if layer:
        S(ent, "Layer", layer)
    if lw is not None:
        S(ent, "Lineweight", lw)
    return ent


def poly(space, pts, closed=True, layer=None, lw=None):
    pl = retry(space.AddLightWeightPolyline, dbl_array([c for p in pts for c in p]))
    S(pl, "Closed", closed)
    return _finish(pl, layer, lw)


def rect(space, x0, y0, x1, y1, layer=None, lw=None):
    return poly(space, [(x0, y0), (x1, y0), (x1, y1), (x0, y1)], True, layer, lw)


def line(space, x1, y1, x2, y2, layer=None, lw=None):
    return _finish(retry(space.AddLine, pt(x1, y1), pt(x2, y2)), layer, lw)


def _align(t, align, x, y):
    if align != "L":
        S(t, "Alignment", ALIGN[align])  # Alignment truoc, TextAlignmentPoint sau
        S(t, "TextAlignmentPoint", pt(x, y))


def text(space, s, x, y, h, style=None, align="L", layer=None):
    t = retry(space.AddText, s, pt(x, y), h)
    if style:
        S(t, "StyleName", style)
    _align(t, align, x, y)
    return _finish(t, layer)


def attdef(space, tag, prompt, default, x, y, h, style=None, align="L", mode=0):
    a = retry(space.AddAttribute, h, mode, prompt, pt(x, y), tag, default)
    if style:
        S(a, "StyleName", style)
    _align(a, align, x, y)
    return a


# ------------------------------------------------------------------ block
def load_block(doc, path, name=None):
    """Nap dinh nghia block tu file .dwg (ten block = ten file) neu ban ve chua co. Khong de lai doi tuong."""
    name = name or os.path.splitext(os.path.basename(path))[0]
    blk = find_block(doc, name)
    if blk is not None:
        return blk
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    ref = retry(doc.ModelSpace.InsertBlock, pt(0, 0), path, 1.0, 1.0, 1.0, 0.0)
    retry(ref.Delete)
    blk = find_block(doc, name)
    # Khung ten ve bang mm giay: bo don vi block de AutoCAD khong tu quy doi theo INSUNITS ban ve dich
    S(blk, "Units", 0)
    return blk


def set_attributes(ref, values):
    """Gan gia tri attribute theo tag (khong phan biet hoa thuong). Tra ve dict tag -> gia tri hien tai."""
    want = {k.upper(): v for k, v in (values or {}).items()}
    out = {}
    for a in retry(ref.GetAttributes):
        tag = a.TagString.upper()
        if tag in want:
            S(a, "TextString", want[tag])
        out[tag] = a.TextString
    return out


def insert_block(space, name, x=0.0, y=0.0, scale=1.0, rot=0.0, layer=None, values=None):
    ref = retry(space.InsertBlock, pt(x, y), name, scale, scale, scale, rot)
    if layer:
        S(ref, "Layer", layer)
    if values:
        set_attributes(ref, values)
    return ref


def block_refs(space, name):
    return [e for e in space if e.ObjectName == "AcDbBlockReference"
            and e.EffectiveName.upper() == name.upper()]


# ------------------------------------------------------------------ layout / viewport
def get_layout(doc, *names):
    for lo in doc.Layouts:
        if lo.Name in names:
            return lo
    raise KeyError(f"Khong co layout {names}; dang co: {[l.Name for l in doc.Layouts]}")


def activate_layout(doc, layout):
    """Chuyen sang layout ma khong de AutoCAD tu tao viewport mac dinh."""
    with temp_var(doc, "LAYOUTCREATEVIEWPORT", 0):
        S(doc, "ActiveLayout", layout)
    wait_idle(doc)


def page_setup(layout, pc3="DWG To PDF.pc3", media="ISO_full_bleed_A3_(420.00_x_297.00_MM)",
               ctb="monochrome.ctb", rotation=0):
    """Page setup in 1:1 theo layout, goc in 0,0, mm."""
    S(layout, "ConfigName", pc3)
    S(layout, "CanonicalMediaName", media)
    S(layout, "PaperUnits", 1)       # acMillimeters
    S(layout, "PlotRotation", rotation)
    S(layout, "PlotType", 5)         # acLayout - khong gan CenterPlot (Invalid input)
    S(layout, "StandardScale", 16)   # AcPlotScale ac1_1
    S(layout, "PlotOrigin", pt2(0, 0))
    S(layout, "StyleSheet", ctb)
    S(layout, "PlotWithPlotStyles", True)


def viewports(doc):
    """Viewport tren paperspace cua layout hien hanh, bo viewport tong (cai dau tien)."""
    wait_idle(doc)
    vps = [e for e in doc.PaperSpace if e.ObjectName == "AcDbViewport"]
    return vps[1:]


def _set_view(app, doc, vp, scale, model_center):
    """scale = don vi giay / don vi model (1:100 ve bang mm tren giay mm -> 0.01)."""
    S(vp, "ViewportOn", True)
    S(vp, "DisplayLocked", False)
    S(doc, "MSpace", True)
    S(doc, "ActivePViewport", vp)
    if model_center is None:
        model_center = (vp.Center[0] / scale, vp.Center[1] / scale)
    retry(app.ZoomCenter, pt(*model_center), vp.Height / scale)
    S(vp, "CustomScale", float(scale))  # PViewport.StandardScale dung enum khac Layout, nen dung CustomScale
    S(doc, "MSpace", False)
    S(vp, "DisplayLocked", True)


def rect_viewport(app, doc, x0, y0, x1, y1, scale=1.0, model_center=None, layer="A-VIEWPORT"):
    vp = retry(doc.PaperSpace.AddPViewport, pt((x0 + x1) / 2, (y0 + y1) / 2), x1 - x0, y1 - y0)
    S(vp, "Layer", layer)
    _set_view(app, doc, vp, scale, model_center)
    return vp


def poly_viewport(app, doc, pts, scale=1.0, model_center=None, layer="A-VIEWPORT"):
    """Viewport cat theo da giac (COM khong co ham nay -> -VPORTS _Object qua SendCommand).
    model_center = None: tam viewport tren giay chia scale (1:1 -> toa do giay trung toa do model)."""
    before = {v.Handle for v in viewports(doc)}
    clip = poly(doc.PaperSpace, pts, True, layer)
    send(doc, f"_.-VPORTS _O {handent(clip)} ")
    new = [v for v in viewports(doc) if v.Handle not in before]
    if not new:
        raise RuntimeError("-VPORTS khong tao duoc viewport")
    vp = new[0]
    S(vp, "Layer", layer)  # lenh tao tren layer hien hanh
    _set_view(app, doc, vp, scale, model_center)
    return vp


# ------------------------------------------------------------------ quy trinh tron goi
def setup_sheet_a3(app, doc, layout_name, new_name=None, values=None, scale=1.0, model_center=None,
                   template=KHUNG_A3, viewport=True):
    """Dung layout A3 hoan chinh: nap block khung ten, page setup, chen khung, viewport cat goc khung ten.
    Chay lai an toan: khung ten da co thi chi cap nhat attribute, viewport da co thi giu nguyen.
    Tra ve dict tom tat de kiem tra (khong xuat file)."""
    block = os.path.splitext(os.path.basename(template))[0]
    ensure_layer(doc, "A-KHUNG-TEN", 7)
    ensure_layer(doc, "A-VIEWPORT", 8, plottable=False)
    load_block(doc, template, block)

    lo = get_layout(doc, layout_name, *( [new_name] if new_name else []))
    if new_name and lo.Name != new_name:
        S(lo, "Name", new_name)
    activate_layout(doc, lo)
    page_setup(lo)

    ps = doc.PaperSpace
    refs = block_refs(ps, block)
    ref = refs[0] if refs else insert_block(ps, block, 0, 0, layer="A-KHUNG-TEN")
    attrs = set_attributes(ref, values)

    vps = [v for v in viewports(doc) if v.Layer.upper() == "A-VIEWPORT"]
    if viewport and not vps:
        vps = [poly_viewport(app, doc, A3_VIEWPORT_CLIP, scale, model_center)]
    retry(app.ZoomExtents)
    return {
        "layout": lo.Name,
        "paper": lo.GetPaperSize(),
        "block": block,
        "attributes": attrs,
        "viewports": [{"handle": v.Handle, "center": tuple(round(c, 2) for c in v.Center[:2]),
                       "w": round(v.Width, 2), "h": round(v.Height, 2),
                       "scale": v.CustomScale, "clipped": v.Clipped, "locked": v.DisplayLocked}
                      for v in vps],
    }
