import io
import math
import re

import ezdxf
from ezdxf.enums import TextEntityAlignment
import streamlit as st
import streamlit.components.v1 as components
import svgwrite


st.set_page_config(
    page_title="CE Rating Plate Designer",
    page_icon="🏷️",
    layout="wide",
)

# -----------------------------------------------------------------------------
# Defaults
# -----------------------------------------------------------------------------
DEFAULTS = {
    "company_name": "DAFRA d.o.o.",
    "company_line2": "Cesta ob železnici 3",
    "company_line3": "3310 Žalec, Slovenia",
    "type_text": "RVD-32-720HS",
    "no_text": "1330117",
    "year_built": "2026",
    "year_refurbished": "",
    "current_type": "L1/L2/L3/N/PE",
    "hz_text": "50",
    "working_voltage": "380",
    "air_pressure": "6",
    "control_voltage_dc": "24",
    "control_voltage_ac": "220",
    "machine_current": "8",
    "fuse_current": "16",
    "schematic_no": "733017",
    "footer_no": "507 598.4",
}

# Official CE logo paths
CE_PATH_1 = "M110,199.498744A100,100 0 0 1 100,200A100,100 0 0 1 100,0A100,100 0 0 1 110,0.501256L110,30.501256A70,70 0 0 0 100,30A70,70 0 0 0 100,170A70,70 0 0 0 110,169.498744Z"
CE_PATH_2 = "M280,199.498744A100,100 0 0 1 270,200A100,100 0 0 1 270,0A100,100 0 0 1 280,0.501256L280,30.501256A70,70 0 0 0 270,30A70,70 0 0 0 201.620283,85L260,85L260,115L201.620283,115A70,70 0 0 0 270,170A70,70 0 0 0 280,169.498744Z"


# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("Plate geometry")
    plate_width = st.number_input("Width [mm]", min_value=80.0, max_value=400.0, value=160.0, step=5.0)
    plate_height = st.number_input("Height [mm]", min_value=120.0, max_value=260.0, value=150.0, step=5.0)
    corner_radius = st.number_input("Corner radius [mm]", min_value=0.0, max_value=20.0, value=0.0, step=0.5)
    hole_diameter = st.number_input("Mounting hole diameter [mm]", min_value=0.0, max_value=12.0, value=3.5, step=0.5)
    hole_offset = st.number_input("Hole offset from corner [mm]", min_value=2.0, max_value=20.0, value=6.0, step=0.5)
    border_offset = st.number_input("Inner border offset [mm]", min_value=1.0, max_value=15.0, value=4.0, step=0.5)

    st.header("Visibility")
    show_left_holes = st.checkbox("Show left holes", value=True)
    show_right_holes = st.checkbox("Show right holes", value=True)
    show_warning_symbol = st.checkbox("Show lightning warning symbol", value=True)

    st.header("Marks / logos")
    show_ce_logo = st.checkbox("Show CE logo under lightning", value=True)
    show_bin_logo = st.checkbox("Show WEEE bin (crossed)", value=False)

left_col, right_col = st.columns([1.0, 1.25], gap="large")

with left_col:
    st.subheader("Top text")
    company_name = st.text_input("Company", value=DEFAULTS["company_name"])
    company_line2 = st.text_input("Address line 1", value=DEFAULTS["company_line2"])
    company_line3 = st.text_input("Address line 2", value=DEFAULTS["company_line3"])

    st.subheader("Plate values")
    c1, c2 = st.columns(2)
    with c1:
        type_text = st.text_input("Type", value=DEFAULTS["type_text"])
        year_built = st.text_input("Leto izdelave", value=DEFAULTS["year_built"])
        current_type = st.text_input("Current / nature du courant", value=DEFAULTS["current_type"])
        working_voltage = st.text_input("Working voltage [V]", value=DEFAULTS["working_voltage"])
        air_pressure = st.text_input("Compressed air / working pressure [bar]", value=DEFAULTS["air_pressure"])
        control_voltage_dc = st.text_input("Controlling voltage DC [V]", value=DEFAULTS["control_voltage_dc"])
        machine_current = st.text_input("Machine nominal current [A]", value=DEFAULTS["machine_current"])

    with c2:
        no_text = st.text_input("No.", value=DEFAULTS["no_text"])
        year_refurbished = st.text_input("Leto obnove", value=DEFAULTS["year_refurbished"])
        hz_text = st.text_input("Hz", value=DEFAULTS["hz_text"])
        control_voltage_ac = st.text_input("Controlling voltage AC [V]", value=DEFAULTS["control_voltage_ac"])
        fuse_current = st.text_input("Fuse nominal current [A]", value=DEFAULTS["fuse_current"])
        schematic_no = st.text_input("Schematic No.", value=DEFAULTS["schematic_no"])
        footer_no = st.text_input("Small footer number", value=DEFAULTS["footer_no"])

st.title("CE Marking & Rating Plate Designer")
st.caption("Haulick-style machine plate with corrected CE logo, crossed WEEE bin, and SVG/DXF export.")


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def clamp_text(value) -> str:
    if value is None:
        return ""
    return str(value)


def draw_svg_text(
    dwg,
    parent,
    text,
    x,
    y,
    size,
    weight="normal",
    anchor="start",
    rotate=None,
    family="Arial, Helvetica, sans-serif",
):
    text = clamp_text(text)
    if not text:
        return
    el = dwg.text(
        text,
        insert=(x, y),
        font_size=size,
        font_family=family,
        font_weight=weight,
        text_anchor=anchor,
        fill="black",
    )
    if rotate is not None:
        el.rotate(rotate, center=(x, y))
    parent.add(el)


def draw_svg_box_text(
    dwg,
    parent,
    text,
    x,
    y,
    w,
    h,
    size,
    weight="normal",
    family="Arial, Helvetica, sans-serif",
):
    text = clamp_text(text)
    if not text:
        return
    el = dwg.text(
        text,
        insert=(x + w / 2, y + h / 2),
        font_size=size,
        font_family=family,
        font_weight=weight,
        text_anchor="middle",
        fill="black",
    )
    el["dominant-baseline"] = "middle"
    parent.add(el)


def draw_svg_multiline(dwg, parent, lines, x, y, size, line_gap, first_bold=True):
    for idx, line in enumerate(lines):
        if not line:
            continue
        weight = "bold" if (idx == 0 and first_bold) else "normal"
        draw_svg_text(dwg, parent, line, x, y + idx * line_gap, size, weight=weight)


def draw_box_svg(dwg, parent, x, y, w, h, stroke_w=0.28):
    parent.add(
        dwg.rect(
            insert=(x, y),
            size=(w, h),
            fill="none",
            stroke="black",
            stroke_width=stroke_w,
        )
    )


def draw_warning_symbol_svg(dwg, parent, x, y, w, h):
    pts = [
        (x + 0.42 * w, y + 0.10 * h),
        (x + 0.67 * w, y + 0.10 * h),
        (x + 0.52 * w, y + 0.44 * h),
        (x + 0.80 * w, y + 0.44 * h),
        (x + 0.43 * w, y + 0.92 * h),
        (x + 0.51 * w, y + 0.61 * h),
        (x + 0.25 * w, y + 0.61 * h),
    ]
    parent.add(dwg.polygon(points=pts, fill="#c60000", stroke="none"))


def draw_ce_logo_svg(dwg, parent, x, y, w, h):
    scale = min(w / 280.0, h / 200.0)
    logo_w = 280.0 * scale
    logo_h = 200.0 * scale
    tx = x + (w - logo_w) / 2
    ty = y + (h - logo_h) / 2

    grp = dwg.g()
    grp["style"] = "fill-rule:evenodd;clip-rule:evenodd"

    p1 = dwg.path(d=CE_PATH_1, fill="black")
    p1["fill-rule"] = "evenodd"
    p1["clip-rule"] = "evenodd"
    p1.update({"transform": f"translate({tx},{ty}) scale({scale})"})
    grp.add(p1)

    p2 = dwg.path(d=CE_PATH_2, fill="black")
    p2["fill-rule"] = "evenodd"
    p2["clip-rule"] = "evenodd"
    p2.update({"transform": f"translate({tx},{ty}) scale({scale})"})
    grp.add(p2)

    parent.add(grp)


def draw_bin_logo_svg(dwg, parent, x, y, w, h):
    target_ratio = 0.62  # width / height
    if w / h > target_ratio:
        bh = h
        bw = h * target_ratio
    else:
        bw = w
        bh = w / target_ratio

    bx = x + (w - bw) / 2
    by = y + (h - bh) / 2

    stroke = max(min(bw, bh) * 0.03, 0.22)
    cx = bx + bw / 2

    lid_y = by + bh * 0.16
    bar_y = by + bh * 0.22
    parent.add(dwg.line((cx - bw * 0.19, bar_y), (cx + bw * 0.19, bar_y), stroke="black", stroke_width=stroke))
    parent.add(dwg.line((cx - bw * 0.07, lid_y), (cx + bw * 0.07, lid_y), stroke="black", stroke_width=stroke))

    top_y = by + bh * 0.30
    bot_y = by + bh * 0.80
    top_w = bw * 0.34
    bot_w = bw * 0.26

    body_pts = [
        (cx - top_w / 2, top_y),
        (cx + top_w / 2, top_y),
        (cx + bot_w / 2, bot_y),
        (cx - bot_w / 2, bot_y),
    ]
    parent.add(dwg.polygon(points=body_pts, fill="none", stroke="black", stroke_width=stroke))

    r = bw * 0.035
    parent.add(dwg.circle(center=(cx - bw * 0.09, bot_y + bh * 0.05), r=r, fill="none", stroke="black", stroke_width=stroke))
    parent.add(dwg.circle(center=(cx + bw * 0.09, bot_y + bh * 0.05), r=r, fill="none", stroke="black", stroke_width=stroke))

    # crossed lines
    parent.add(dwg.line((bx + bw * 0.16, by + bh * 0.10), (bx + bw * 0.86, by + bh * 0.92), stroke="black", stroke_width=stroke))
    parent.add(dwg.line((bx + bw * 0.86, by + bh * 0.10), (bx + bw * 0.16, by + bh * 0.92), stroke="black", stroke_width=stroke))


# DXF helpers -----------------------------------------------------------------
def dy(plate_h: float, y_top: float) -> float:
    return plate_h - y_top


def add_dxf_rect(msp, x, y, w, h, plate_h, layer="0"):
    pts = [
        (x, dy(plate_h, y)),
        (x + w, dy(plate_h, y)),
        (x + w, dy(plate_h, y + h)),
        (x, dy(plate_h, y + h)),
    ]
    msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})


def add_dxf_text(msp, text, x, y, height, plate_h, align=TextEntityAlignment.LEFT, rotation=0):
    text = clamp_text(text)
    if not text:
        return
    ent = msp.add_text(text, dxfattribs={"height": max(height, 1.0), "rotation": rotation})
    ent.set_placement((x, dy(plate_h, y)), align=align)


def add_dxf_box_text(msp, text, x, y, w, h, height, plate_h):
    text = clamp_text(text)
    if not text:
        return
    ent = msp.add_text(text, dxfattribs={"height": max(height, 1.0)})
    ent.set_placement((x + w / 2, dy(plate_h, y + h / 2)), align=TextEntityAlignment.MIDDLE_CENTER)


def add_dxf_multiline(msp, lines, x, y, height, line_gap, plate_h):
    for idx, line in enumerate(lines):
        if not line:
            continue
        add_dxf_text(msp, line, x, y + idx * line_gap, height, plate_h)


def add_dxf_line(msp, x1, y1, x2, y2, plate_h):
    msp.add_line((x1, dy(plate_h, y1)), (x2, dy(plate_h, y2)))


def add_dxf_polyline(msp, pts, plate_h, close=False):
    dxf_pts = [(x, dy(plate_h, y)) for x, y in pts]
    msp.add_lwpolyline(dxf_pts, close=close)


def draw_warning_symbol_dxf(msp, x, y, w, h, plate_h):
    pts = [
        (x + 0.42 * w, y + 0.10 * h),
        (x + 0.67 * w, y + 0.10 * h),
        (x + 0.52 * w, y + 0.44 * h),
        (x + 0.80 * w, y + 0.44 * h),
        (x + 0.43 * w, y + 0.92 * h),
        (x + 0.51 * w, y + 0.61 * h),
        (x + 0.25 * w, y + 0.61 * h),
    ]
    add_dxf_polyline(msp, pts, plate_h, close=True)


def tokenize_svg_path(path_d: str):
    token_re = r"[MLAZ]|[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?"
    return re.findall(token_re, path_d)


def vector_angle(ux, uy, vx, vy):
    dot = ux * vx + uy * vy
    det = ux * vy - uy * vx
    return math.atan2(det, dot)


def sample_svg_arc(x1, y1, rx, ry, x_axis_rotation, large_arc_flag, sweep_flag, x2, y2, steps=28):
    rx = abs(rx)
    ry = abs(ry)

    if rx == 0 or ry == 0:
        return [(x2, y2)]

    phi = math.radians(x_axis_rotation % 360.0)
    cos_phi = math.cos(phi)
    sin_phi = math.sin(phi)

    dx2 = (x1 - x2) / 2.0
    dy2 = (y1 - y2) / 2.0

    x1p = cos_phi * dx2 + sin_phi * dy2
    y1p = -sin_phi * dx2 + cos_phi * dy2

    lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry)
    if lam > 1:
        s = math.sqrt(lam)
        rx *= s
        ry *= s

    num = (rx * rx * ry * ry) - (rx * rx * y1p * y1p) - (ry * ry * x1p * x1p)
    den = (rx * rx * y1p * y1p) + (ry * ry * x1p * x1p)
    coef = 0.0 if den == 0 else math.sqrt(max(0.0, num / den))
    if large_arc_flag == sweep_flag:
        coef = -coef

    cxp = coef * ((rx * y1p) / ry)
    cyp = coef * (-(ry * x1p) / rx)

    cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2.0
    cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2.0

    ux = (x1p - cxp) / rx
    uy = (y1p - cyp) / ry
    vx = (-x1p - cxp) / rx
    vy = (-y1p - cyp) / ry

    theta1 = vector_angle(1.0, 0.0, ux, uy)
    delta_theta = vector_angle(ux, uy, vx, vy)

    if (not sweep_flag) and delta_theta > 0:
        delta_theta -= 2.0 * math.pi
    elif sweep_flag and delta_theta < 0:
        delta_theta += 2.0 * math.pi

    pts = []
    for i in range(1, steps + 1):
        t = theta1 + delta_theta * (i / steps)
        ct = math.cos(t)
        st = math.sin(t)
        x = cos_phi * rx * ct - sin_phi * ry * st + cx
        y = sin_phi * rx * ct + cos_phi * ry * st + cy
        pts.append((x, y))
    return pts


def svg_path_to_points(path_d: str, tx=0.0, ty=0.0, scale=1.0, arc_steps=28):
    tokens = tokenize_svg_path(path_d)
    i = 0
    cmd = None
    x = y = 0.0
    start_x = start_y = 0.0
    pts = []

    while i < len(tokens):
        tok = tokens[i]
        if tok in ("M", "L", "A", "Z"):
            cmd = tok
            i += 1

        if cmd == "M":
            x = float(tokens[i]); y = float(tokens[i + 1]); i += 2
            start_x, start_y = x, y
            pts.append((tx + x * scale, ty + y * scale))
            cmd = "L"
        elif cmd == "L":
            x = float(tokens[i]); y = float(tokens[i + 1]); i += 2
            pts.append((tx + x * scale, ty + y * scale))
        elif cmd == "A":
            rx = float(tokens[i]); ry = float(tokens[i + 1]); xrot = float(tokens[i + 2])
            large_arc = int(float(tokens[i + 3])); sweep = int(float(tokens[i + 4]))
            x2 = float(tokens[i + 5]); y2 = float(tokens[i + 6]); i += 7

            arc_pts = sample_svg_arc(x, y, rx, ry, xrot, large_arc, sweep, x2, y2, steps=arc_steps)
            for ax, ay in arc_pts:
                pts.append((tx + ax * scale, ty + ay * scale))
            x, y = x2, y2
        elif cmd == "Z":
            # close handled by polyline close=True
            cmd = None
        else:
            break

    return pts


def draw_ce_logo_dxf(msp, x, y, w, h, plate_h):
    scale = min(w / 280.0, h / 200.0)
    logo_w = 280.0 * scale
    logo_h = 200.0 * scale
    tx = x + (w - logo_w) / 2
    ty = y + (h - logo_h) / 2

    pts1 = svg_path_to_points(CE_PATH_1, tx=tx, ty=ty, scale=scale, arc_steps=36)
    pts2 = svg_path_to_points(CE_PATH_2, tx=tx, ty=ty, scale=scale, arc_steps=36)

    add_dxf_polyline(msp, pts1, plate_h, close=True)
    add_dxf_polyline(msp, pts2, plate_h, close=True)


def draw_bin_logo_dxf(msp, x, y, w, h, plate_h):
    target_ratio = 0.62
    if w / h > target_ratio:
        bh = h
        bw = h * target_ratio
    else:
        bw = w
        bh = w / target_ratio

    bx = x + (w - bw) / 2
    by = y + (h - bh) / 2
    cx = bx + bw / 2

    lid_y = by + bh * 0.16
    bar_y = by + bh * 0.22
    add_dxf_line(msp, cx - bw * 0.19, bar_y, cx + bw * 0.19, bar_y, plate_h)
    add_dxf_line(msp, cx - bw * 0.07, lid_y, cx + bw * 0.07, lid_y, plate_h)

    top_y = by + bh * 0.30
    bot_y = by + bh * 0.80
    top_w = bw * 0.34
    bot_w = bw * 0.26

    body_pts = [
        (cx - top_w / 2, top_y),
        (cx + top_w / 2, top_y),
        (cx + bot_w / 2, bot_y),
        (cx - bot_w / 2, bot_y),
    ]
    add_dxf_polyline(msp, body_pts, plate_h, close=True)

    r = bw * 0.035
    msp.add_circle((cx - bw * 0.09, dy(plate_h, bot_y + bh * 0.05)), r)
    msp.add_circle((cx + bw * 0.09, dy(plate_h, bot_y + bh * 0.05)), r)

    # crossed lines
    add_dxf_line(msp, bx + bw * 0.16, by + bh * 0.10, bx + bw * 0.86, by + bh * 0.92, plate_h)
    add_dxf_line(msp, bx + bw * 0.86, by + bh * 0.10, bx + bw * 0.16, by + bh * 0.92, plate_h)


# -----------------------------------------------------------------------------
# SVG generation
# -----------------------------------------------------------------------------
def generate_plate_svg(w: float, h: float) -> str:
    dwg = svgwrite.Drawing(size=(f"{w}mm", f"{h}mm"), viewBox=f"0 0 {w} {h}")
    dwg.attribs["style"] = "background:white"

    template_w = 160.0
    template_h = 150.0

    sx = w / template_w
    sy = h / template_h

    def X(v): return v * sx
    def Y(v): return v * sy

    outer = dwg.g(id="outer")
    dwg.add(outer)

    outer.add(
        dwg.rect(
            insert=(0.5, 0.5),
            size=(w - 1, h - 1),
            rx=corner_radius,
            ry=corner_radius,
            fill="white",
            stroke="black",
            stroke_width=0.55,
        )
    )
    outer.add(
        dwg.rect(
            insert=(border_offset, border_offset),
            size=(w - 2 * border_offset, h - 2 * border_offset),
            fill="none",
            stroke="black",
            stroke_width=0.3,
        )
    )

    if hole_diameter > 0:
        r = hole_diameter / 2
        holes = []
        if show_left_holes:
            holes.extend([(hole_offset, hole_offset), (hole_offset, h - hole_offset)])
        if show_right_holes:
            holes.extend([(w - hole_offset, hole_offset), (w - hole_offset, h - hole_offset)])
        for cx, cy in holes:
            outer.add(dwg.circle(center=(cx, cy), r=r, fill="none", stroke="black", stroke_width=0.3))

    content = dwg.g(id="content")
    dwg.add(content)

    ix = border_offset
    iy = border_offset
    iw = w - 2 * border_offset
    ih = h - 2 * border_offset

    # Left symbol panel
    panel_w = X(44)
    content.add(dwg.line(start=(ix + panel_w, iy), end=(ix + panel_w, iy + ih), stroke="black", stroke_width=0.35))

    pad = X(4.5)
    zone_x = ix + pad
    zone_y = iy + Y(6)
    zone_w = panel_w - 2 * pad
    zone_h = ih - Y(12)

    if show_ce_logo or show_bin_logo:
        lightning_h = zone_h * 0.56
    else:
        lightning_h = zone_h * 0.78

    if show_warning_symbol:
        draw_warning_symbol_svg(dwg, content, zone_x, zone_y, zone_w, lightning_h)

    logo_y = zone_y + lightning_h + Y(3)

    if show_ce_logo:
        ce_h = Y(12)
        draw_ce_logo_svg(dwg, content, zone_x + zone_w * 0.06, logo_y, zone_w * 0.88, ce_h)
        logo_y += ce_h + Y(3)

    if show_bin_logo:
        bin_h = Y(12)
        draw_bin_logo_svg(dwg, content, zone_x + zone_w * 0.18, logo_y, zone_w * 0.64, bin_h)

    # Main area
    mx = ix + panel_w + X(4)
    my = iy + Y(2.5)
    mr = ix + iw - X(4)
    main_w = mr - mx

    cx = mx + main_w / 2
    draw_svg_text(dwg, content, company_name, cx, my + Y(7), Y(7.0), weight="bold", anchor="middle")
    draw_svg_text(dwg, content, company_line2, cx, my + Y(12.4), Y(3.5), anchor="middle")
    draw_svg_text(dwg, content, company_line3, cx, my + Y(17.6), Y(3.5), anchor="middle")

    label_x = mx + X(0)
    box1_x = mx + X(22)
    box1_w = X(40)
    box2_label_x = mx + X(65)
    box2_x = mx + X(74)
    box2_w = X(27)

    unit_x = mx + X(60)
    single_box_x = mx + X(73)
    single_box_w = X(28)

    year_box1_x = mx + X(73)
    year_box_w = X(12.5)
    year_gap = X(3.0)
    year_box2_x = year_box1_x + year_box_w + year_gap

    small_font = Y(2.45)
    value_font = Y(2.95)
    line_gap = Y(2.30)
    box_h = Y(4.4)

    y1 = my + Y(27.0)
    y2 = my + Y(35.5)
    y3 = my + Y(46.5)
    y4 = my + Y(58.0)
    y5 = my + Y(69.5)
    y6 = my + Y(81.0)
    y6b = y6 + Y(5.0)
    y7 = my + Y(95.0)
    y8 = my + Y(106.5)
    y9 = my + Y(118.0)

    draw_svg_text(dwg, content, "Type", label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box1_x, y1 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font)

    draw_svg_text(dwg, content, "No.", box2_label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y1 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font)

    draw_svg_multiline(
        dwg,
        content,
        ["Baujahr / Jahr der Überholung", "Year built / Year refurbished", "Année de fabrication / Année de rénovation"],
        label_x,
        y2,
        small_font,
        line_gap,
    )
    draw_box_svg(dwg, content, year_box1_x, y2 - box_h / 2, year_box_w, box_h)
    draw_svg_box_text(dwg, content, year_built, year_box1_x, y2 - box_h / 2, year_box_w, box_h, value_font)

    draw_box_svg(dwg, content, year_box2_x, y2 - box_h / 2, year_box_w, box_h)
    draw_svg_box_text(dwg, content, year_refurbished, year_box2_x, y2 - box_h / 2, year_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Stromart", "Current", "Nature du courant"], label_x, y3, small_font, line_gap)
    draw_box_svg(dwg, content, box1_x, y3 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font)

    draw_svg_text(dwg, content, "Hz", box2_label_x, y3, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y3 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Betriebsspannung", "Working voltage", "Voltage de service"], label_x, y4, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y4, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y4 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(
        dwg,
        content,
        ["Betriebsdruck Druckluft", "Compressed air pressure", "Pression d'air comprimé"],
        label_x,
        y5,
        small_font,
        line_gap,
    )
    draw_svg_text(dwg, content, "bar", unit_x - X(2), y5, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y5 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, air_pressure, single_box_x, y5 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Steuerspannung", "Controlling voltage", "Voltage de commande"], label_x, y6, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y6, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y6 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font)

    draw_box_svg(dwg, content, single_box_x, y6b - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Maschine Nennstrom", "Nominal current machine", "Machine intensité nominale"], label_x, y7, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y7, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y7 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Sicherungs-Nennstrom", "Nominal current fuses", "Intensité de protection nominale"], label_x, y8, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y8, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y8 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, ["Schaltplan", "Schematic", "Schéma de connexions"], label_x, y9, small_font, line_gap)
    draw_svg_text(dwg, content, "No.", unit_x - X(6), y9, small_font, weight="bold")
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    draw_box_svg(dwg, content, schematic_x, y9 - box_h / 2, schematic_w, box_h)
    draw_svg_box_text(dwg, content, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font)

    bottom_box_h = Y(3.4)
    bottom_box_y = iy + ih - Y(7.6)
    draw_box_svg(dwg, content, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h)

    draw_svg_text(dwg, content, footer_no, mr - X(1), iy + ih - Y(0.9), Y(1.8), anchor="end")

    return dwg.tostring()


# -----------------------------------------------------------------------------
# DXF generation
# -----------------------------------------------------------------------------
def generate_plate_dxf(w: float, h: float) -> bytes:
    doc = ezdxf.new("R2010", setup=True)
    doc.units = 4
    msp = doc.modelspace()

    template_w = 160.0
    template_h = 150.0

    sx = w / template_w
    sy = h / template_h

    def X(v): return v * sx
    def Y(v): return v * sy

    add_dxf_rect(msp, 0, 0, w, h, h)
    add_dxf_rect(msp, border_offset, border_offset, w - 2 * border_offset, h - 2 * border_offset, h)

    if hole_diameter > 0:
        r = hole_diameter / 2
        holes = []
        if show_left_holes:
            holes.extend([(hole_offset, hole_offset), (hole_offset, h - hole_offset)])
        if show_right_holes:
            holes.extend([(w - hole_offset, hole_offset), (w - hole_offset, h - hole_offset)])
        for cx, cy in holes:
            msp.add_circle((cx, dy(h, cy)), r)

    ix = border_offset
    iy = border_offset
    iw = w - 2 * border_offset
    ih = h - 2 * border_offset

    panel_w = X(44)
    add_dxf_line(msp, ix + panel_w, iy, ix + panel_w, iy + ih, h)

    pad = X(4.5)
    zone_x = ix + pad
    zone_y = iy + Y(6)
    zone_w = panel_w - 2 * pad
    zone_h = ih - Y(12)

    if show_ce_logo or show_bin_logo:
        lightning_h = zone_h * 0.56
    else:
        lightning_h = zone_h * 0.78

    if show_warning_symbol:
        draw_warning_symbol_dxf(msp, zone_x, zone_y, zone_w, lightning_h, h)

    logo_y = zone_y + lightning_h + Y(3)

    if show_ce_logo:
        ce_h = Y(12)
        draw_ce_logo_dxf(msp, zone_x + zone_w * 0.06, logo_y, zone_w * 0.88, ce_h, h)
        logo_y += ce_h + Y(3)

    if show_bin_logo:
        bin_h = Y(12)
        draw_bin_logo_dxf(msp, zone_x + zone_w * 0.18, logo_y, zone_w * 0.64, bin_h, h)

    mx = ix + panel_w + X(4)
    my = iy + Y(2.5)
    mr = ix + iw - X(4)
    main_w = mr - mx

    cx = mx + main_w / 2
    add_dxf_text(msp, company_name, cx, my + Y(7), Y(5.6), h, align=TextEntityAlignment.MIDDLE_CENTER)
    add_dxf_text(msp, company_line2, cx, my + Y(12.4), Y(2.7), h, align=TextEntityAlignment.MIDDLE_CENTER)
    add_dxf_text(msp, company_line3, cx, my + Y(17.6), Y(2.7), h, align=TextEntityAlignment.MIDDLE_CENTER)

    label_x = mx + X(0)
    box1_x = mx + X(22)
    box1_w = X(40)
    box2_label_x = mx + X(65)
    box2_x = mx + X(74)
    box2_w = X(27)

    unit_x = mx + X(60)
    single_box_x = mx + X(73)
    single_box_w = X(28)

    year_box1_x = mx + X(73)
    year_box_w = X(12.5)
    year_gap = X(3.0)
    year_box2_x = year_box1_x + year_box_w + year_gap

    small_font = Y(2.10)
    value_font = Y(2.50)
    line_gap = Y(2.30)
    box_h = Y(4.4)

    y1 = my + Y(27.0)
    y2 = my + Y(35.5)
    y3 = my + Y(46.5)
    y4 = my + Y(58.0)
    y5 = my + Y(69.5)
    y6 = my + Y(81.0)
    y6b = y6 + Y(5.0)
    y7 = my + Y(95.0)
    y8 = my + Y(106.5)
    y9 = my + Y(118.0)

    add_dxf_text(msp, "Type", label_x, y1, small_font, h)
    add_dxf_rect(msp, box1_x, y1 - box_h / 2, box1_w, box_h, h)
    add_dxf_box_text(msp, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font, h)

    add_dxf_text(msp, "No.", box2_label_x, y1, small_font, h)
    add_dxf_rect(msp, box2_x, y1 - box_h / 2, box2_w, box_h, h)
    add_dxf_box_text(msp, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font, h)

    add_dxf_multiline(
        msp,
        ["Baujahr / Jahr der Überholung", "Year built / Year refurbished", "Année de fabrication / Année de rénovation"],
        label_x,
        y2,
        small_font,
        line_gap,
        h,
    )
    add_dxf_rect(msp, year_box1_x, y2 - box_h / 2, year_box_w, box_h, h)
    add_dxf_box_text(msp, year_built, year_box1_x, y2 - box_h / 2, year_box_w, box_h, value_font, h)

    add_dxf_rect(msp, year_box2_x, y2 - box_h / 2, year_box_w, box_h, h)
    add_dxf_box_text(msp, year_refurbished, year_box2_x, y2 - box_h / 2, year_box_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Stromart", "Current", "Nature du courant"], label_x, y3, small_font, line_gap, h)
    add_dxf_rect(msp, box1_x, y3 - box_h / 2, box1_w, box_h, h)
    add_dxf_box_text(msp, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font, h)

    add_dxf_text(msp, "Hz", box2_label_x, y3, small_font, h)
    add_dxf_rect(msp, box2_x, y3 - box_h / 2, box2_w, box_h, h)
    add_dxf_box_text(msp, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Betriebsspannung", "Working voltage", "Voltage de service"], label_x, y4, small_font, line_gap, h)
    add_dxf_text(msp, "V", unit_x, y4, small_font, h)
    add_dxf_rect(msp, single_box_x, y4 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_multiline(
        msp,
        ["Betriebsdruck Druckluft", "Compressed air pressure", "Pression d'air comprimé"],
        label_x,
        y5,
        small_font,
        line_gap,
        h,
    )
    add_dxf_text(msp, "bar", unit_x - X(2), y5, small_font, h)
    add_dxf_rect(msp, single_box_x, y5 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, air_pressure, single_box_x, y5 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Steuerspannung", "Controlling voltage", "Voltage de commande"], label_x, y6, small_font, line_gap, h)
    add_dxf_text(msp, "V", unit_x, y6, small_font, h)
    add_dxf_rect(msp, single_box_x, y6 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_rect(msp, single_box_x, y6b - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Maschine Nennstrom", "Nominal current machine", "Machine intensité nominale"], label_x, y7, small_font, line_gap, h)
    add_dxf_text(msp, "A", unit_x, y7, small_font, h)
    add_dxf_rect(msp, single_box_x, y7 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Sicherungs-Nennstrom", "Nominal current fuses", "Intensité de protection nominale"], label_x, y8, small_font, line_gap, h)
    add_dxf_text(msp, "A", unit_x, y8, small_font, h)
    add_dxf_rect(msp, single_box_x, y8 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_multiline(msp, ["Schaltplan", "Schematic", "Schéma de connexions"], label_x, y9, small_font, line_gap, h)
    add_dxf_text(msp, "No.", unit_x - X(6), y9, small_font, h)
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    add_dxf_rect(msp, schematic_x, y9 - box_h / 2, schematic_w, box_h, h)
    add_dxf_box_text(msp, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font, h)

    bottom_box_h = Y(3.4)
    bottom_box_y = iy + ih - Y(7.6)
    add_dxf_rect(msp, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h, h)

    add_dxf_text(msp, footer_no, mr - X(1), iy + ih - Y(0.9), Y(1.6), h, align=TextEntityAlignment.RIGHT)

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


# -----------------------------------------------------------------------------
# Render
# -----------------------------------------------------------------------------
svg_content = generate_plate_svg(plate_width, plate_height)
dxf_content = generate_plate_dxf(plate_width, plate_height)

with right_col:
    st.subheader("Live preview")
    preview_scale = st.slider("Preview scale", min_value=2.0, max_value=8.0, value=4.8, step=0.2)
    preview_h = int(plate_height * preview_scale + 60)

    components.html(
        f"""
        <div style="background:#f3f4f6;padding:18px;border-radius:14px;overflow:auto;">
            <div style="display:flex;justify-content:center;">
                <div style="width:{plate_width * preview_scale}px;">
                    {svg_content}
                </div>
            </div>
        </div>
        """,
        height=min(max(preview_h, 380), 980),
        scrolling=True,
    )

    st.download_button(
        "Download SVG",
        data=svg_content,
        file_name="dafra_rating_plate.svg",
        mime="image/svg+xml",
        use_container_width=True,
    )

    st.download_button(
        "Download DXF",
        data=dxf_content,
        file_name="dafra_rating_plate.dxf",
        mime="application/dxf",
        use_container_width=True,
    )

    st.info("CE logo is fixed in SVG and DXF. WEEE bin is now crossed and keeps proper aspect ratio.")

    with st.expander("SVG source"):
        st.code(svg_content[:12000], language="xml")