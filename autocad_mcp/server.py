"""MCP server ket noi AutoCAD qua COM (ActiveX).

Yeu cau: AutoCAD ban full (khong phai LT) dang mo it nhat mot ban ve.
Viet cho goi mcp 2.x (FastMCP da doi ten thanh MCPServer).

Chay:  autocad-mcp            (MCP server qua stdio - de Claude goi)
       autocad-mcp --check    (kiem tra ket noi AutoCAD roi thoat)
"""
import argparse
import functools
import sys
from importlib.metadata import PackageNotFoundError, version

import pythoncom
import pywintypes
import win32com.client
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from . import __version__

mcp = MCPServer("autocad")


def acad_tool(fn):
    """Dang ky tool va doi loi COM thanh ToolError.
    mcp 2.x chi chuyen nguyen van ToolError toi Claude, loi khac bi an thanh cau chung chung."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except pywintypes.com_error as exc:
            raise ToolError(
                f"AutoCAD bao loi: {exc}. Neu CAD dang chay do lenh, nhan Esc roi thu lai."
            ) from exc
    return mcp.tool()(wrapper)


def _doc():
    """Lay ban ve dang mo. Moi lan goi deu ket noi lai cho an toan voi COM."""
    # Tool chay tren worker thread (anyio), moi thread phai CoInitialize
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("AutoCAD.Application")
    except pywintypes.com_error:
        raise ToolError("Khong tim thay AutoCAD dang mo. Hay mo AutoCAD va mot ban ve.")
    if app.Documents.Count == 0:
        raise ToolError("AutoCAD chua mo ban ve nao.")
    return app.ActiveDocument


def _pt(x, y, z=0.0):
    # AutoCAD COM bat buoc toa do la mang double, khong nhan tuple thuong
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z))
    )


@acad_tool
def drawing_info() -> dict:
    """Thong tin ban ve dang mo: ten, duong dan, layer hien hanh, so doi tuong."""
    doc = _doc()
    return {
        "name": doc.Name,
        "path": doc.FullName,
        "active_layer": doc.ActiveLayer.Name,
        "modelspace_objects": doc.ModelSpace.Count,
    }


@acad_tool
def list_layers() -> list[dict]:
    """Liet ke tat ca layer cung mau, linetype, trang thai on/freeze/lock."""
    doc = _doc()
    return [
        {"name": l.Name, "color": l.color, "linetype": l.Linetype,
         "on": l.LayerOn, "frozen": l.Freeze, "locked": l.Lock}
        for l in doc.Layers
    ]


@acad_tool
def create_layer(name: str, color: int = 7) -> str:
    """Tao layer moi (hoac cap nhat mau neu da co). color la ma ACI 1-255."""
    layer = _doc().Layers.Add(name)
    layer.color = color
    return f"Layer '{name}' mau {color}"


def _require_layer(doc, layer):
    """Bao loi truoc khi ve neu layer chua co, de khong de lai doi tuong tren layer hien hanh."""
    if not layer:
        return
    try:
        doc.Layers.Item(layer)
    except pywintypes.com_error:
        raise ToolError(
            f"Layer '{layer}' chua co trong ban ve, chua ve gi. Tao bang create_layer roi goi lai."
        ) from None


def _finish(ent, layer):
    """Gan layer cho doi tuong vua tao. Gan loi thi xoa doi tuong roi moi bao loi."""
    if layer:
        try:
            ent.Layer = layer
        except pywintypes.com_error:
            ent.Delete()
            raise
    return ent.Handle


@acad_tool
def draw_line(x1: float, y1: float, x2: float, y2: float, layer: str = "") -> str:
    """Ve line trong ModelSpace. layer phai co san (tao bang create_layer); de trong = layer hien hanh.
    Tra ve handle cua doi tuong."""
    doc = _doc()
    _require_layer(doc, layer)
    return _finish(doc.ModelSpace.AddLine(_pt(x1, y1), _pt(x2, y2)), layer)


@acad_tool
def draw_circle(cx: float, cy: float, radius: float, layer: str = "") -> str:
    """Ve circle trong ModelSpace. layer phai co san (tao bang create_layer); de trong = layer hien hanh.
    Tra ve handle cua doi tuong."""
    doc = _doc()
    _require_layer(doc, layer)
    return _finish(doc.ModelSpace.AddCircle(_pt(cx, cy), radius), layer)


@acad_tool
def add_text(text: str, x: float, y: float, height: float = 2.5, layer: str = "") -> str:
    """Them single-line text. layer phai co san (tao bang create_layer); de trong = layer hien hanh.
    Tra ve handle cua doi tuong."""
    doc = _doc()
    _require_layer(doc, layer)
    return _finish(doc.ModelSpace.AddText(text, _pt(x, y), height), layer)


_GET = pythoncom.DISPATCH_PROPERTYGET
# ObjectName -> thuoc tinh doc them ngoai Name/Layer/EffectiveName. MINSERT chen Rows x Columns ban
_REF_KINDS = {"AcDbBlockReference": (), "AcDbMInsertBlock": ("Rows", "Columns")}


def _block_refs(container):
    """Liet ke (ten dinh nghia, layer, ten hieu dung, so ban chen) cua block reference va MINSERT
    trong ModelSpace hoac trong mot dinh nghia block. Goi IDispatch truc tiep voi dispid da cache,
    nhanh hon ~3 lan so voi win32com dynamic (do tren FF-F1: 15.330 doi tuong 21,7 s -> 6,8 s)."""
    en = container._oleobj_.InvokeTypes(
        pythoncom.DISPID_NEWENUM, 0, _GET | pythoncom.DISPATCH_METHOD, (13, 10), ()
    ).QueryInterface(pythoncom.IID_IEnumVARIANT)
    id_type = None
    ref_ids = {}  # ObjectName -> dispid, moi loai doi tuong cache rieng
    refs = []
    while batch := en.Next(500):
        for d in batch:
            if id_type is None:
                id_type = d.GetIDsOfNames("ObjectName")
            kind = d.Invoke(id_type, 0, _GET, 1)
            if kind not in _REF_KINDS:
                continue
            if kind not in ref_ids:
                ref_ids[kind] = [d.GetIDsOfNames(n)
                                 for n in ("Name", "Layer", "EffectiveName", *_REF_KINDS[kind])]
            name, lay, eff, *grid = (d.Invoke(i, 0, _GET, 1) for i in ref_ids[kind])
            refs.append((name, lay, eff, grid[0] * grid[1] if grid else 1))
    return refs


def _merge(into, other, times=1):
    for name, n in other.items():
        into[name] = into.get(name, 0) + n * times


@acad_tool
def count_blocks(layer: str = "", include_nested: bool = False,
                 skip_hidden_layers: bool = False) -> dict:
    """Dem block theo ten trong ModelSpace (vd dem den, sprinkler, o cam).
    MINSERT tinh du Rows x Columns ban. Block an danh (*U..., vd ARRAY lien ket) khong dem chinh no
    ma luon dem cac block ben trong; dynamic block duoc dem theo ten goc.
    layer: chi dem block tren layer nay. Block long nam tren layer 0 duoc tinh theo layer cua block cha.
    include_nested: dem ca block long ben trong block khac (moi lan chen block cha deu tinh). Khong di vao xref.
    skip_hidden_layers: bo qua block tren layer dang freeze hoac off, va moi thu ben trong block bi freeze.
    Tra ve total, skipped_hidden (so block bi bo qua vi an), anonymous_expanded (so block an danh da
    duyet vao), counts (ten -> so luong).
    Ban ve lon co the cham vi moi doi tuong la mot lan goi COM."""
    doc = _doc()
    states = {l.Name.upper(): (l.LayerOn, l.Freeze) for l in doc.Layers}
    defs = {}  # ten dinh nghia -> [(ten, layer, ten hieu dung, so ban chen) cua block long]
    memo = {}

    def nested_refs(defname):
        # Moi dinh nghia chi doc qua COM mot lan; goi COM la phan cham nhat
        if defname not in defs:
            bdef = doc.Blocks.Item(defname)
            defs[defname] = [] if bdef.IsXRef else _block_refs(bdef)
        return defs[defname]

    def visit(ref, parent_layer, parent_frozen):
        # Ket qua cho mot ban chen; nguoi goi nhan voi so ban chen cua MINSERT
        name, ent_layer, eff_name, _ = ref
        counts, skipped, anon = {}, 0, 0
        # Doi tuong tren layer 0 trong block lay layer cua block cha
        lay = parent_layer if parent_layer and ent_layer == "0" else ent_layer
        on, frozen = states.get(lay.upper(), (True, False))
        # Block an danh chi la vo boc (ARRAY...): dem thu ben trong, khong dem chinh no.
        # Dynamic block cung co ten *U nhung EffectiveName la ten goc nen van dem binh thuong
        anonymous = eff_name.startswith("*")
        if anonymous:
            anon = 1
        elif skip_hidden_layers and (parent_frozen or frozen or not on):
            skipped = 1
        elif not layer or lay.upper() == layer.upper():
            counts[eff_name] = 1
        if include_nested or anonymous:
            c, s, a = walk_def(name, lay, parent_frozen or frozen)
            _merge(counts, c)
            skipped += s
            anon += a
        return counts, skipped, anon

    def walk_def(defname, parent_layer, parent_frozen):
        # Ket qua cua mot dinh nghia trong mot ngu canh layer dung lai cho moi lan chen
        key = (defname, parent_layer.upper(), parent_frozen)
        if key not in memo:
            memo[key] = walk(nested_refs(defname), parent_layer, parent_frozen)
        return memo[key]

    def walk(refs, parent_layer, parent_frozen):
        counts, skipped, anon = {}, 0, 0
        for ref in refs:
            n = ref[3]
            c, s, a = visit(ref, parent_layer, parent_frozen)
            _merge(counts, c, n)
            skipped += s * n
            anon += a * n
        return counts, skipped, anon

    counts, skipped, anon = walk(_block_refs(doc.ModelSpace), "", False)
    return {
        "total": sum(counts.values()),
        "skipped_hidden": skipped,
        "anonymous_expanded": anon,
        "counts": dict(sorted(counts.items(), key=lambda kv: -kv[1])),
    }


def _check():
    """In thong tin moi truong va ban ve dang mo. Tra ve ma thoat 0 neu ket noi duoc."""
    try:
        mcp_ver = version("mcp")
    except PackageNotFoundError:
        mcp_ver = "?"
    print(f"autocad-mcp {__version__} | Python {sys.version.split()[0]} | mcp {mcp_ver}")
    try:
        doc = _doc()
    except ToolError as exc:
        print("LOI:", exc)
        return 1
    app = doc.Application
    print(f"AutoCAD {app.Version} | so ban ve dang mo: {app.Documents.Count}")
    print(f"Ban ve hien hanh: {doc.FullName or doc.Name}")
    print(f"So layer: {doc.Layers.Count} | so doi tuong ModelSpace: {doc.ModelSpace.Count}")
    print("KET NOI OK")
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(prog="autocad-mcp",
                                     description="MCP server ket noi AutoCAD qua COM (ActiveX).")
    parser.add_argument("--check", action="store_true", help="kiem tra ket noi AutoCAD roi thoat")
    parser.add_argument("--version", action="version", version=f"autocad-mcp {__version__}")
    args = parser.parse_args(argv)
    if args.check:
        sys.exit(_check())
    mcp.run()  # giao tiep qua stdio: khong duoc print ra stdout trong che do nay


if __name__ == "__main__":
    main()
