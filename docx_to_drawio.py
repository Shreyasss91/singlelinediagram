"""Convert a legacy Word single line diagram (.docx, Word shapes) to .drawio.

The hand-made sheets predate this tool and live in Word as DrawingML shapes - 400-odd rectangles,
frames, lines and text boxes per page. This reads `word/document.xml` and writes one drawio page per
Word page, keeping every shape editable rather than flattening the page to an image.

    python docx_to_drawio.py "field_data/test/SLD_MRSF1_Feeder_and_BHRoad_Feeder.docx"
    python docx_to_drawio.py input.docx --out outputs/legacy/MRS F1 legacy.drawio

What it maps, and how faithfully:

  * rect / roundRect          -> a drawio rectangle (a filled one keeps its fill)
  * frame                     -> two concentric rectangles, the DTC symbol of these sheets
  * line                      -> a straight drawio edge, direction from flipH/flipV
  * text boxes                -> drawio text, one box per Word paragraph line
  * a number in a box         -> text, not a box (the kVA rating sits inside the symbol)

Fidelity: geometry is 1:1 with the page (1 in = 100 px, the same A3 units the generator uses) and
text keeps its size and alignment. It is a redraw, not a facsimile: Word's frame inset, text box
insets and any autofit are approximated, and the page is re-centred rather than placed at Word's
own offset. Compare it with the PNG export of the same DOCX before relying on it.

Pages are separated where a drawing whose text starts "Single line diagram" begins; anchors placed
outside the page (a stray shape Word clips) are dropped and counted.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"

PAGE_W, PAGE_H = 1654, 1169      # A3 landscape in drawio units (1 in = 100 px)
RULES_W, RULES_H = 15124000, 10692000   # the same page in EMU, for the off-page test
MARGIN = 34
PT_TO_PX = 1.6                   # measured against the rendered sheets, not Word's 1.389

THEME: dict[str, str] = {}


def q(tag: str, ns: str) -> str:
    return "{%s}%s" % (ns, tag)


def wval(el, name: str = "val"):
    """Attribute by local name: Word writes some attributes namespace-qualified."""
    if el is None:
        return None
    for key, value in el.attrib.items():
        if key.split("}")[-1] == name:
            return value
    return None


def theme_colours(z: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(z.read("word/theme/theme1.xml"))
    except KeyError:
        return {}
    scheme = root.find(".//" + q("clrScheme", A))
    out = {}
    if scheme is None:
        return out
    for el in list(scheme):
        colour = list(el)[0]
        value = colour.get("val") or colour.get("lastClr")
        if value and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
            out[el.tag.split("}")[1]] = "#" + value.upper()
    return out


def resolve_colour(node, default=None):
    """`a:solidFill` -> #RRGGBB, honouring srgbClr / schemeClr and a shade or lumMod."""
    if node is None or not len(node):
        return default
    colour = list(node)[0]
    tag = colour.tag.split("}")[1]
    if tag == "srgbClr":
        return "#" + colour.get("val").upper()
    if tag == "sysClr":
        return "#" + (colour.get("lastClr") or "000000").upper()
    if tag != "schemeClr":
        return default
    base = THEME.get(colour.get("val"), "#000000")
    for name, clamp in (("shade", False), ("lumMod", True)):
        el = colour.find(q(name, A))
        if el is None:
            continue
        factor = int(wval(el)) / 100000.0
        rgb = []
        for i in (1, 3, 5):
            v = int(base[i:i + 2], 16) * factor
            rgb.append("%02X" % (min(255, int(v)) if clamp else int(v)))
        base = "#" + "".join(rgb)
    return base


def line_of(sp, style):
    """(colour, width_px) of a shape's outline; colour None means no visible stroke."""
    ln = sp.find(q("ln", A))
    colour, width = None, None
    if ln is not None:
        width = int(ln.get("w", "9525")) / 12700.0 * 1.389
        fill = ln.find(q("solidFill", A))
        if fill is not None:
            colour = resolve_colour(fill)
        elif ln.find(q("noFill", A)) is not None:
            return None, 0.0
    if colour is None and style is not None:
        ref = style.find(q("lnRef", A))
        if ref is not None:
            if wval(ref, "idx") == "0":
                return None, 0.0
            colour = resolve_colour(ref, "#1F497D")
    return colour, max(0.8, width or 1.0)


def fill_of(sp):
    fill = sp.find(q("solidFill", A))
    if fill is not None:
        return resolve_colour(fill, "#FFFFFF")
    return None


def body_text(shape):
    """Each paragraph of a text box: text, size in points, bold, colour, alignment."""
    paras = []
    for p in shape.iter(q("p", W)):
        text = "".join(t.text or "" for t in p.iter(q("t", W))).strip()
        if not text:
            continue
        sz, bold, colour = 10, False, "#000000"
        mark = p.find(q("pPr", W) + "/" + q("rPr", W))
        if mark is not None:
            size_el = mark.find(q("sz", W))
            if size_el is not None and wval(size_el):
                sz = int(wval(size_el))
            bold = mark.find(q("b", W)) is not None
            col = mark.find(q("color", W))
            if col is not None:
                colour = "#" + (wval(col) or "000000").upper()
        for run in p.findall(q("r", W)):
            rpr = run.find(q("rPr", W))
            if rpr is None:
                continue
            size_el = rpr.find(q("sz", W))
            if size_el is not None and wval(size_el):
                sz = max(sz, int(wval(size_el)))
            bold = bold or rpr.find(q("b", W)) is not None
            col = rpr.find(q("color", W))
            if col is not None:
                colour = "#" + (wval(col) or "000000").upper()
        jc = p.find(q("pPr", W) + "/" + q("jc", W))
        align = {"ctr": "center", "r": "right", "both": "center"}.get(
            wval(jc) or "ctr", "center")
        paras.append(dict(text=text, pt=sz / 2.0, bold=bold, colour=colour,
                          align=align))
    return paras


def xfrm_of(sp):
    xfrm = sp.find(".//" + q("xfrm", A))
    if xfrm is None:
        return None
    off, ext = xfrm.find(q("off", A)), xfrm.find(q("ext", A))
    if off is None or ext is None:
        return None
    return ((float(off.get("x")), float(off.get("y"))),
            (float(ext.get("cx")), float(ext.get("cy"))),
            xfrm.get("flipH") == "1", xfrm.get("flipV") == "1")


def read_shape(shape):
    sp = shape.find(q("spPr", WPS))
    if sp is None:
        return None
    geo = sp.find(q("prstGeom", A))
    preset = geo.get("prst") if geo is not None else "rect"
    box = xfrm_of(sp)
    if box is None:
        return None
    (x, y), (w, h), fliph, flipv = box
    style = shape.find(q("style", WPS))
    stroke, sw = line_of(sp, style)
    paras = body_text(shape)
    if preset == "line":
        x0, y0, x1, y1 = x, y, x + w, y + h
        if fliph:
            x0, x1 = x1, x0
        if flipv:
            y0, y1 = y1, y0
        return dict(kind="line", x=x0, y=y0, w=x1 - x0, h=y1 - y0,
                    pts=((x0, y0), (x1, y1)), stroke=stroke or "#1F497D",
                    sw=sw or 1.0, preset=preset, paras=[], fill=None)
    return dict(kind="shape", x=x, y=y, w=w, h=h, preset=preset, stroke=stroke,
                sw=sw or 1.0, fill=fill_of(sp), paras=paras)


def collect(group, parent_tr, out):
    """Flatten a `wpg:wgp` / `wpg:grpSp` tree, mapping child coordinates to page space."""
    gx = group.find(q("grpSpPr", WPG) + "/" + q("xfrm", A))
    tr = parent_tr
    if gx is not None:
        off, ext = gx.find(q("off", A)), gx.find(q("ext", A))
        coff, cext = gx.find(q("chOff", A)), gx.find(q("chExt", A))
        if None not in (off, ext, coff, cext):
            ox, oy = float(off.get("x")), float(off.get("y"))
            ex, ey = float(ext.get("cx")), float(ext.get("cy"))
            cx, cy = float(coff.get("x")), float(coff.get("y"))
            cw, ch = float(cext.get("cx")), float(cext.get("cy"))
            sx = ex / cw if cw else 1.0
            sy = ey / ch if ch else 1.0
            bx, by = ox - cx * sx, oy - cy * sy

            def tr(px, py, _p=parent_tr, bx=bx, by=by, sx=sx, sy=sy):
                return _p(px * sx + bx, py * sy + by)

    for child in list(group):
        if child.tag in (q("wgp", WPG), q("grpSp", WPG)):
            collect(child, tr, out)
        elif child.tag == q("wsp", WPS):
            shape = read_shape(child)
            if shape is None:
                continue
            if shape["kind"] == "line":
                shape["pts"] = tuple(tr(px, py) for px, py in shape["pts"])
                (x0, y0), (x1, y1) = shape["pts"]
                shape.update(x=min(x0, x1), y=min(y0, y1), w=abs(x1 - x0),
                             h=abs(y1 - y0))
            else:
                x0, y0 = tr(shape["x"], shape["y"])
                x1, y1 = tr(shape["x"] + shape["w"], shape["y"] + shape["h"])
                shape.update(x=min(x0, x1), y=min(y0, y1), w=abs(x1 - x0),
                             h=abs(y1 - y0))
            out.append(shape)


def anchor_offset(anchor):
    def off(side):
        el = anchor.find(q(side, WP))
        pos = el.find(q("posOffset", WP)) if el is not None else None
        return float(pos.text) if pos is not None and pos.text else 0.0
    return off("positionH"), off("positionV")


def page_shapes(anchors):
    """Split the anchored drawings into pages at each 'Single line diagram' title."""
    pages, current, dropped = [], None, 0
    for anchor in anchors:
        hx, vy = anchor_offset(anchor)
        if hx > RULES_W or vy > RULES_H:
            dropped += 1
            continue
        text = " ".join(t.text or "" for t in anchor.iter(q("t", W))).strip()
        shapes = []
        group = anchor.find(".//" + q("wgp", WPG))
        if group is not None:
            collect(group, lambda px, py: (px, py), shapes)
        else:
            for wsp in anchor.iter(q("wsp", WPS)):
                shape = read_shape(wsp)
                if shape is not None:
                    shapes.append(shape)
        if text.lower().startswith("single line diagram") or current is None:
            current = {"title": text, "shapes": []}
            pages.append(current)
        current["shapes"].extend(shapes)
    if dropped:
        print(f"  dropped {dropped} anchor(s) placed off the page")
    for page in pages:
        for shape in page["shapes"]:
            if shape["paras"] and shape["paras"][0]["text"].lower().startswith(
                    "single line diagram"):
                page["title"] = shape["paras"][0]["text"]
                break
    return pages


def bbox(shapes):
    xs, ys = [], []
    for shape in shapes:
        if shape["kind"] == "line":
            for px, py in shape["pts"]:
                xs.append(px)
                ys.append(py)
        else:
            xs += [shape["x"], shape["x"] + shape["w"]]
            ys += [shape["y"], shape["y"] + shape["h"]]
    return min(xs), min(ys), max(xs), max(ys)


def fit_to_page(shapes):
    """Scale 1:1 into A3 where it fits, then centre; never enlarge a drawing."""
    x0, y0, x1, y1 = bbox(shapes)
    sw, sh = max(x1 - x0, 1.0), max(y1 - y0, 1.0)
    avail_w, avail_h = PAGE_W - 2 * MARGIN, PAGE_H - 2 * MARGIN
    k = min(1.0, avail_w / sw, avail_h / sh)
    ox = MARGIN + (avail_w - sw * k) / 2
    oy = MARGIN + (avail_h - sh * k) / 2

    def m(px, py):
        return ((px - x0) * k + ox, (py - y0) * k + oy)

    out = []
    for shape in shapes:
        if shape["kind"] == "line":
            shape = dict(shape, pts=tuple(m(px, py) for px, py in shape["pts"]))
        else:
            x, y = m(shape["x"], shape["y"])
            shape = dict(shape, x=x, y=y, w=max(1.0, shape["w"] * k),
                         h=max(1.0, shape["h"] * k))
        out.append(shape)
    return out


def attr(text: str) -> str:
    return escape(text, {'"': "&quot;", "\n": "&#10;"})


def page_xml(shapes, name: str) -> str:
    cells, n = [], 2
    edges = [s for s in shapes if s["kind"] == "line"]

    def numeric(shape):
        return bool(shape["paras"]) and all(
            re.fullmatch(r"[\d.]+", p["text"]) for p in shape["paras"])

    # a number in a box is a rating label, not a symbol: draw the text only
    texts = [s for s in shapes if s["kind"] == "shape" and s["paras"]
             and (not s["stroke"] or numeric(s))]
    boxes = [s for s in shapes if s["kind"] == "shape" and s["paras"]
             and s not in texts]
    symbols = [s for s in shapes if s["kind"] == "shape" and not s["paras"]]

    for shape in edges:
        (ax, ay), (bx, by) = shape["pts"]
        cells.append(
            f'<mxCell id="e{n}" style="endArrow=none;html=0;rounded=0;'
            f'strokeColor={shape["stroke"]};strokeWidth={max(0.8, shape["sw"]):.1f};" '
            f'edge="1" parent="1"><mxGeometry relative="1" as="geometry">'
            f'<mxPoint x="{ax:.1f}" y="{ay:.1f}" as="sourcePoint"/>'
            f'<mxPoint x="{bx:.1f}" y="{by:.1f}" as="targetPoint"/>'
            f'</mxGeometry></mxCell>')
        n += 1

    for shape in symbols:
        pieces = [(shape["x"], shape["y"], shape["w"], shape["h"])]
        if shape["preset"] == "frame":     # Word's frame is two concentric squares
            inset = 0.13 * min(shape["w"], shape["h"])
            pieces.append((shape["x"] + inset, shape["y"] + inset,
                           shape["w"] - 2 * inset, shape["h"] - 2 * inset))
        for gx, gy, gw, gh in pieces:
            style = (f'rounded=0;html=0;fillColor=none;strokeColor={shape["stroke"]};'
                     f'strokeWidth={shape["sw"]:.1f};')
            cells.append(
                f'<mxCell id="v{n}" value="" style="{style}" vertex="1" parent="1">'
                f'<mxGeometry x="{gx:.1f}" y="{gy:.1f}" width="{max(1.0, gw):.1f}" '
                f'height="{max(1.0, gh):.1f}" as="geometry"/></mxCell>')
            n += 1

    for shape in boxes:
        value = "<br>".join(escape(p["text"]) for p in shape["paras"])
        first = shape["paras"][0]
        extra = "rounded=1;arcSize=18;" if shape["preset"] == "roundRect" else "rounded=0;"
        style = (f'{extra}whiteSpace=wrap;html=0;fillColor={shape["fill"] or "none"};'
                 f'strokeColor={shape["stroke"] or "none"};'
                 f'strokeWidth={shape["sw"]:.1f};fontFamily=Arial;align={first["align"]};'
                 f'verticalAlign=middle;fontSize={max(5.0, first["pt"] * PT_TO_PX):.1f};'
                 f'fontColor={first["colour"]};fontStyle={1 if first["bold"] else 0};')
        cells.append(
            f'<mxCell id="v{n}" value="{attr(value)}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{shape["x"]:.1f}" y="{shape["y"]:.1f}" '
            f'width="{shape["w"]:.1f}" height="{shape["h"]:.1f}" as="geometry"/></mxCell>')
        n += 1

    for shape in texts:
        value = "<br>".join(escape(p["text"]) for p in shape["paras"])
        first = shape["paras"][0]
        style = (f'text;html=0;whiteSpace=nowrap;align={first["align"]};'
                 f'verticalAlign=middle;fontFamily=Arial;'
                 f'fontSize={max(5.0, first["pt"] * PT_TO_PX):.1f};'
                 f'fontColor={first["colour"]};'
                 f'fontStyle={1 if first["bold"] else 0};'
                 f'strokeColor=none;fillColor=none;')
        cells.append(
            f'<mxCell id="v{n}" value="{attr(value)}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{shape["x"]:.1f}" y="{shape["y"]:.1f}" '
            f'width="{shape["w"]:.1f}" height="{shape["h"]:.1f}" as="geometry"/></mxCell>')
        n += 1

    return (f'<diagram name="{attr(name)}"><mxGraphModel dx="1400" dy="1000" grid="1" '
            f'gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" '
            f'page="1" pageScale="1" pageWidth="{PAGE_W}" pageHeight="{PAGE_H}" '
            f'math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/>'
            + "".join(cells) + "</root></mxGraphModel></diagram>")


def convert(src: str, dst: str) -> list[str]:
    global THEME
    with zipfile.ZipFile(src) as z:
        THEME = theme_colours(z)
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(q("body", W))
    if body is None:
        raise SystemExit("no word/body in the document")
    anchors = body.findall(".//" + q("anchor", WP))
    if not anchors:
        raise SystemExit("no anchored drawings found - is this an inline-shape document?")
    pages = page_shapes(anchors)
    titles = []
    parts = []
    for i, page in enumerate(pages, 1):
        name = page["title"] or f"Page {i}"
        shapes = fit_to_page(page["shapes"])
        parts.append(page_xml(shapes, name))
        titles.append(name)
        print(f"  page {i}: {len(shapes)} shapes  {name[:60]}")
    with open(dst, "w", encoding="utf-8") as f:
        f.write('<mxfile host="app.diagrams.net" type="device">'
                + "".join(parts) + "</mxfile>")
    return titles


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("docx", help="the .docx holding the Word drawing")
    ap.add_argument("--out", help="output .drawio (default: beside the docx, same name)")
    args = ap.parse_args(argv)
    out = args.out or os.path.splitext(args.docx)[0] + ".drawio"
    print(f"reading {args.docx}")
    convert(args.docx, out)
    print(f"wrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
