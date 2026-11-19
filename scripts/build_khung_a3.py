"""Dung file mau templates/RIK-KHUNG-A3.dwg (khung A3 ngang + khung ten RIKCO co attribute).

Muon doi mau khung ten: sua ham draw() roi chay lai. Script mo mot ban ve trang tam thoi trong AutoCAD,
ve vao ModelSpace, luu de file mau roi dong ban ve tam - khong dung toi ban ve dang mo cua nguoi dung.

    & ".\\.venv\\Scripts\\python.exe" scripts\\build_khung_a3.py      (chay trong thu muc repo)

Ban ve chen file nay se co block RIK-KHUNG-A3 (ten = ten file), chen tai 0,0 tren layout A3 ti le 1:1.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import acadlib as A  # noqa: E402

NAVY = (40, 56, 144)  # RIKCO #283890
ST, STB = "RIK-ARIAL", "RIK-ARIAL-B"


def draw(app, doc, sp):
    A.ensure_text_style(doc, ST, bold=False)
    A.ensure_text_style(doc, STB, bold=True)
    thin, mid, thick = 18, 35, 50

    # Mep giay A3 va khung ve: le trai 20 (dong gay), cac le con lai 5
    A.rect(sp, 0, 0, 420, 297, lw=13)
    A.rect(sp, 20, 5, 415, 292, lw=thick)

    # Khung ten 180 x 55 goc phai duoi; XS = ranh gioi o ky ten | thong tin du an
    X0, X1, Y0, Y1, XS = 235, 415, 5, 60, 300
    A.rect(sp, X0, Y0, X1, Y1, lw=mid)
    A.line(sp, XS, Y0, XS, Y1, lw=mid)

    # Trai: 5 dong ky ten cao 11 - cot chuc danh | ho ten | chu ky
    for y in (16, 27, 38, 49):
        A.line(sp, X0, y, XS, y, lw=thin)
    A.line(sp, 258, Y0, 258, Y1, lw=thin)
    A.line(sp, 283, Y0, 283, Y1, lw=thin)
    roles = [("CHỦ TRÌ", "CHU_TRI"), ("THIẾT KẾ", "THIET_KE"), ("VẼ", "VE"),
             ("KIỂM TRA", "KIEM_TRA"), ("PHÊ DUYỆT", "PHE_DUYET")]
    for i, (role, tag) in enumerate(roles):
        yc = 54.5 - 11 * i
        A.text(sp, role, 237, yc, 2.0, ST, "ML")
        A.attdef(sp, tag, f"Họ tên {role.lower()}", "", 260, yc, 2.0, ST, "ML")

    # Phai: cong ty | CDT | du an | hang muc | ten ban ve | o so lieu
    for y in (16, 27, 49):
        A.line(sp, XS, y, X1, y, lw=mid)
    for y in (34.333, 41.667):
        A.line(sp, XS, y, X1, y, lw=thin)
    logo = A.text(sp, "RIKCO", 302.5, 54.5, 5.0, STB, "ML")
    A.S(logo, "TrueColor", A.true_color(app, NAVY))
    A.text(sp, "CÔNG TY CP THƯƠNG MẠI VÀ", 328, 56.2, 2.0, STB, "ML")
    A.text(sp, "XÂY LẮP CƠ ĐIỆN RIK", 328, 52.6, 2.0, STB, "ML")

    for label, tag, yc in [("CHỦ ĐẦU TƯ:", "CHU_DAU_TU", 45.333), ("DỰ ÁN:", "DU_AN", 38.0),
                           ("HẠNG MỤC:", "HANG_MUC", 30.667)]:
        A.text(sp, label, 302, yc, 1.8, ST, "ML")
        A.attdef(sp, tag, label.rstrip(":").capitalize(), "", 322, yc, 2.5, STB, "ML")

    A.text(sp, "TÊN BẢN VẼ:", 302, 24.6, 1.6, ST, "ML")
    A.attdef(sp, "TEN_BAN_VE", "Tên bản vẽ", "", 357.5, 20.4, 3.5, STB, "MC")

    for label, tag, xa, xb in [("TỶ LỆ", "TY_LE", 300, 325), ("NGÀY", "NGAY", 325, 352),
                               ("SỐ BẢN VẼ", "SO_BAN_VE", 352, 392), ("LẦN SỬA", "LAN_SUA", 392, 415)]:
        if xa != XS:
            A.line(sp, xa, Y0, xa, 16, lw=thin)
        A.text(sp, label, xa + 1.5, 13.6, 1.6, ST, "ML")
        A.attdef(sp, tag, label.capitalize(), "", (xa + xb) / 2, 8.6, 3.0, STB, "MC")


def main(out=A.KHUNG_A3):
    app, user_doc = A.connect()
    A.wait_idle(user_doc)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    tmp = A.call(app.Documents, "Add")
    A.wait_idle(tmp)
    try:
        A.set_var(tmp, "INSUNITS", 4)  # mm
        draw(app, tmp, A.retry(lambda: tmp.ModelSpace))
        A.call(app, "ZoomExtents")
        if os.path.exists(out):
            os.remove(out)
        A.call(tmp, "SaveAs", out)
        n = tmp.ModelSpace.Count
    finally:
        A.call(tmp, "Close", False)
        A.call(user_doc, "Activate")
    print(f"Da luu {out} ({n} doi tuong, {os.path.getsize(out)} byte)")


if __name__ == "__main__":
    main()
