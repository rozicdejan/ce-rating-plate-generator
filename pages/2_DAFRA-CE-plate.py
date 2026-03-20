import io
import math
import re

import ezdxf
from matplotlib.font_manager import FontProperties
from matplotlib.textpath import TextPath
from matplotlib.transforms import Affine2D
import streamlit as st
import streamlit.components.v1 as components
import svgwrite

try:
    import cairosvg
except Exception:
    cairosvg = None

try:
    from PIL import Image
except Exception:
    Image = None

st.set_page_config(
    page_title="CE Rating Plate Designer",
    page_icon="🏷️",
    layout="wide",
)

DEFAULTS = {
    "company_name": "DAFRA d.o.o.",
    "company_line2": "Cesta ob železnici 3",
    "company_line3": "3310 Žalec, Slovenija",
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

LANGUAGE_PACKS = {
    "SI / EN / DE": {
        "type_short": "Tip",
        "no_short": "Št.",
        "year_label": ["Leto izdelave / Leto obnove", "Year built / Year refurbished", "Baujahr / Jahr der Überholung"],
        "current_label": ["Vrsta toka", "Current", "Stromart"],
        "working_voltage_label": ["Delovna napetost", "Working voltage", "Betriebsspannung"],
        "air_pressure_label": ["Delovni tlak zraka", "Compressed air pressure", "Betriebsdruck Druckluft"],
        "control_voltage_label": ["Krmilna napetost", "Control voltage", "Steuerspannung"],
        "machine_current_label": ["Nazivni tok stroja", "Nominal current machine", "Maschine Nennstrom"],
        "fuse_current_label": ["Nazivni tok varovalk", "Nominal current fuses", "Sicherungs-Nennstrom"],
        "schematic_label": ["Shema vezave", "Schematic", "Schaltplan"],
    },
    "DE / EN / FR": {
        "type_short": "Typ",
        "no_short": "Nr.",
        "year_label": ["Baujahr / Jahr der Überholung", "Year built / Year refurbished", "Année de fabrication / Année de rénovation"],
        "current_label": ["Stromart", "Current", "Nature du courant"],
        "working_voltage_label": ["Betriebsspannung", "Working voltage", "Voltage de service"],
        "air_pressure_label": ["Betriebsdruck Druckluft", "Compressed air pressure", "Pression d'air comprimé"],
        "control_voltage_label": ["Steuerspannung", "Control voltage", "Voltage de commande"],
        "machine_current_label": ["Maschine Nennstrom", "Nominal current machine", "Machine intensité nominale"],
        "fuse_current_label": ["Sicherungs-Nennstrom", "Nominal current fuses", "Intensité de protection nominale"],
        "schematic_label": ["Schaltplan", "Schematic", "Schéma de connexions"],
    },
    "EN / DE / SI": {
        "type_short": "Type",
        "no_short": "No.",
        "year_label": ["Year built / Year refurbished", "Baujahr / Jahr der Überholung", "Leto izdelave / Leto obnove"],
        "current_label": ["Current", "Stromart", "Vrsta toka"],
        "working_voltage_label": ["Working voltage", "Betriebsspannung", "Delovna napetost"],
        "air_pressure_label": ["Compressed air pressure", "Betriebsdruck Druckluft", "Delovni tlak zraka"],
        "control_voltage_label": ["Control voltage", "Steuerspannung", "Krmilna napetost"],
        "machine_current_label": ["Nominal current machine", "Maschine Nennstrom", "Nazivni tok stroja"],
        "fuse_current_label": ["Nominal current fuses", "Sicherungs-Nennstrom", "Nazivni tok varovalk"],
        "schematic_label": ["Schematic", "Schaltplan", "Shema vezave"],
    },
    "SI / EN": {
        "type_short": "Tip",
        "no_short": "Št.",
        "year_label": ["Leto izdelave / Leto obnove", "Year built / Year refurbished"],
        "current_label": ["Vrsta toka", "Current"],
        "working_voltage_label": ["Delovna napetost", "Working voltage"],
        "air_pressure_label": ["Delovni tlak zraka", "Compressed air pressure"],
        "control_voltage_label": ["Krmilna napetost", "Control voltage"],
        "machine_current_label": ["Nazivni tok stroja", "Nominal current machine"],
        "fuse_current_label": ["Nazivni tok varovalk", "Nominal current fuses"],
        "schematic_label": ["Shema vezave", "Schematic"],
    },
}

CE_PATH_1 = "M110,199.498744A100,100 0 0 1 100,200A100,100 0 0 1 100,0A100,100 0 0 1 110,0.501256L110,30.501256A70,70 0 0 0 100,30A70,70 0 0 0 100,170A70,70 0 0 0 110,169.498744Z"
CE_PATH_2 = "M280,199.498744A100,100 0 0 1 270,200A100,100 0 0 1 270,0A100,100 0 0 1 280,0.501256L280,30.501256A70,70 0 0 0 270,30A70,70 0 0 0 201.620283,85L260,85L260,115L201.620283,115A70,70 0 0 0 270,170A70,70 0 0 0 280,169.498744Z"

TEMPLATE_W = 160.0
TEMPLATE_H = 150.0
GUIDE_COLOR = "#2563eb"

with st.sidebar:
    st.header("Geometrija tablice")
    plate_width = st.number_input("Širina [mm]", min_value=80.0, max_value=400.0, value=160.0, step=5.0)
    plate_height = st.number_input("Višina [mm]", min_value=120.0, max_value=260.0, value=150.0, step=5.0)
    corner_radius = st.number_input("Radij vogala [mm]", min_value=0.0, max_value=20.0, value=0.0, step=0.5)
    hole_diameter = st.number_input("Premer montažne luknje [mm]", min_value=0.0, max_value=12.0, value=3.5, step=0.5)
    hole_offset = st.number_input("Odmik lukenj od roba [mm]", min_value=2.0, max_value=20.0, value=6.0, step=0.5)
    border_offset = st.number_input("Odmik notranjega roba [mm]", min_value=1.0, max_value=15.0, value=4.0, step=0.5)

    st.header("Prikaz")
    show_left_holes = st.checkbox("Pokaži leve luknje", value=True)
    show_right_holes = st.checkbox("Pokaži desne luknje", value=True)
    show_warning_symbol = st.checkbox("Pokaži opozorilno strelo", value=True)
    show_dimensions = st.checkbox("Pokaži mere v predogledu", value=False)
    include_guides_in_dxf = st.checkbox("Vključi GUIDE layer v DXF", value=False)

    st.header("Oznake / logotipi")
    show_ce_logo = st.checkbox("Pokaži CE logotip pod strelo", value=True)
    show_bin_logo = st.checkbox("Pokaži WEEE koš (prekrižan)", value=False)

    st.header("Jezikovni paket")
    language_pack_name = st.selectbox("Izberi paket", list(LANGUAGE_PACKS.keys()), index=0)

labels = LANGUAGE_PACKS[language_pack_name]

left_col, right_col = st.columns([1.0, 1.25], gap="large")
with left_col:
    st.subheader("Glavni podatki")
    company_name = st.text_input("Podjetje", value=DEFAULTS["company_name"])
    company_line2 = st.text_input("Naslov 1", value=DEFAULTS["company_line2"])
    company_line3 = st.text_input("Naslov 2", value=DEFAULTS["company_line3"])

    st.subheader("Vrednosti na tablici")
    c1, c2 = st.columns(2)
    with c1:
        type_text = st.text_input("Tip", value=DEFAULTS["type_text"])
        year_built = st.text_input("Leto izdelave", value=DEFAULTS["year_built"])
        current_type = st.text_input("Vrsta toka", value=DEFAULTS["current_type"])
        working_voltage = st.text_input("Delovna napetost [V]", value=DEFAULTS["working_voltage"])
        air_pressure = st.text_input("Delovni tlak zraka [bar]", value=DEFAULTS["air_pressure"])
        control_voltage_dc = st.text_input("Krmilna napetost DC [V]", value=DEFAULTS["control_voltage_dc"])
        machine_current = st.text_input("Nazivni tok stroja [A]", value=DEFAULTS["machine_current"])
    with c2:
        no_text = st.text_input("Št.", value=DEFAULTS["no_text"])
        year_refurbished = st.text_input("Leto obnove", value=DEFAULTS["year_refurbished"])
        hz_text = st.text_input("Frekvenca [Hz]", value=DEFAULTS["hz_text"])
        control_voltage_ac = st.text_input("Krmilna napetost AC [V]", value=DEFAULTS["control_voltage_ac"])
        fuse_current = st.text_input("Nazivni tok varovalk [A]", value=DEFAULTS["fuse_current"])
        schematic_no = st.text_input("Številka sheme", value=DEFAULTS["schematic_no"])
        footer_no = st.text_input("Mala številka spodaj", value=DEFAULTS["footer_no"])

st.title("CE / Rating Plate Designer")
st.caption("DXF tekst je izvožen kot vektorski obrisi, zato ga EZCAD praviloma prikaže. Dodan je PNG/JPG izvoz in zgornji DAFRA logo.")

def clamp_text(value) -> str:
    return "" if value is None else str(value)

def fit_font_size(text: str, base_size: float, max_width: float, padding: float = 1.2, min_ratio: float = 0.68) -> float:
    text = clamp_text(text)
    if not text:
        return base_size
    usable = max(max_width - 2 * padding, 1.0)
    est = len(text) * base_size * 0.56
    if est <= usable:
        return base_size
    scaled = base_size * usable / est
    return max(base_size * min_ratio, scaled)

def ensure_layers(doc):
    wanted = {"BORDER": 7, "HOLES": 2, "TEXT": 7, "LOGOS": 1, "GUIDE": 5}
    for name, color in wanted.items():
        if name not in doc.layers:
            doc.layers.add(name=name, color=color)

def tokenize_svg_path(path_d: str):
    token_re = r"[MLAZHV]|[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?"
    return re.findall(token_re, path_d)

def vector_angle(ux, uy, vx, vy):
    return math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)

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
    pts = []
    while i < len(tokens):
        tok = tokens[i]
        if tok in ("M", "L", "A", "H", "V", "Z"):
            cmd = tok
            i += 1
        if cmd == "M":
            x = float(tokens[i])
            y = float(tokens[i + 1])
            i += 2
            pts.append((tx + x * scale, ty + y * scale))
            cmd = "L"
        elif cmd == "L":
            x = float(tokens[i])
            y = float(tokens[i + 1])
            i += 2
            pts.append((tx + x * scale, ty + y * scale))
        elif cmd == "H":
            x = float(tokens[i])
            i += 1
            pts.append((tx + x * scale, ty + y * scale))
        elif cmd == "V":
            y = float(tokens[i])
            i += 1
            pts.append((tx + x * scale, ty + y * scale))
        elif cmd == "A":
            rx = float(tokens[i])
            ry = float(tokens[i + 1])
            xrot = float(tokens[i + 2])
            large_arc = int(float(tokens[i + 3]))
            sweep = int(float(tokens[i + 4]))
            x2 = float(tokens[i + 5])
            y2 = float(tokens[i + 6])
            i += 7
            for ax, ay in sample_svg_arc(x, y, rx, ry, xrot, large_arc, sweep, x2, y2, steps=arc_steps):
                pts.append((tx + ax * scale, ty + ay * scale))
            x, y = x2, y2
        elif cmd == "Z":
            cmd = None
        else:
            break
    return pts

def draw_svg_text(dwg, parent, text, x, y, size, weight="normal", anchor="start", rotate=None, family="Arial, Helvetica, sans-serif", fill="black"):
    text = clamp_text(text)
    if not text:
        return
    el = dwg.text(text, insert=(x, y), font_size=size, font_family=family, font_weight=weight, text_anchor=anchor, fill=fill)
    if rotate is not None:
        el.rotate(rotate, center=(x, y))
    parent.add(el)

def draw_svg_box_text(dwg, parent, text, x, y, w, h, base_size, weight="normal", family="Arial, Helvetica, sans-serif"):
    text = clamp_text(text)
    if not text:
        return
    size = fit_font_size(text, base_size, w, padding=1.1)
    el = dwg.text(text, insert=(x + w / 2, y + h / 2), font_size=size, font_family=family, font_weight=weight, text_anchor="middle", fill="black")
    el["dominant-baseline"] = "middle"
    parent.add(el)

def draw_svg_multiline(dwg, parent, lines, x, y, size, line_gap, first_bold=True):
    max_line_len = max((len(clamp_text(line)) for line in lines if clamp_text(line)), default=0)
    if max_line_len > 28:
        size *= 0.92
        line_gap *= 0.94
    if max_line_len > 36:
        size *= 0.90
        line_gap *= 0.92
    for idx, line in enumerate(lines):
        if not line:
            continue
        weight = "bold" if (idx == 0 and first_bold) else "normal"
        draw_svg_text(dwg, parent, line, x, y + idx * line_gap, size, weight=weight)

def draw_box_svg(dwg, parent, x, y, w, h, stroke_w=0.28, stroke="black"):
    parent.add(dwg.rect(insert=(x, y), size=(w, h), fill="none", stroke=stroke, stroke_width=stroke_w))

def draw_warning_symbol_svg(dwg, parent, x, y, w, h):
    pts = [(x + 0.42 * w, y + 0.10 * h), (x + 0.67 * w, y + 0.10 * h), (x + 0.52 * w, y + 0.44 * h), (x + 0.80 * w, y + 0.44 * h), (x + 0.43 * w, y + 0.92 * h), (x + 0.51 * w, y + 0.61 * h), (x + 0.25 * w, y + 0.61 * h)]
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
    target_ratio = 0.62
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
    body_pts = [(cx - top_w / 2, top_y), (cx + top_w / 2, top_y), (cx + bot_w / 2, bot_y), (cx - bot_w / 2, bot_y)]
    parent.add(dwg.polygon(points=body_pts, fill="none", stroke="black", stroke_width=stroke))
    r = bw * 0.035
    parent.add(dwg.circle(center=(cx - bw * 0.09, bot_y + bh * 0.05), r=r, fill="none", stroke="black", stroke_width=stroke))
    parent.add(dwg.circle(center=(cx + bw * 0.09, bot_y + bh * 0.05), r=r, fill="none", stroke="black", stroke_width=stroke))
    parent.add(dwg.line((bx + bw * 0.16, by + bh * 0.10), (bx + bw * 0.86, by + bh * 0.92), stroke="black", stroke_width=stroke))
    parent.add(dwg.line((bx + bw * 0.86, by + bh * 0.10), (bx + bw * 0.16, by + bh * 0.92), stroke="black", stroke_width=stroke))

def draw_header_logo_svg(dwg, parent, x, y, size):
    grp = dwg.g()
    vb = 256.0
    scale = size / vb
    grp.translate(x, y)
    grp.add(dwg.circle(center=(128 * scale, 128 * scale), r=110 * scale, fill="black"))
    path = dwg.path(
        d="M28 88 H150 A42 42 0 0 1 150 168 H28",
        fill="none",
        stroke="white",
        stroke_width=18 * scale,
        stroke_linecap="round",
        stroke_linejoin="round",
    )
    grp.add(path)
    parent.add(grp)


def _unit_vec(dx, dy):
    n = math.hypot(dx, dy)
    if n == 0:
        return 0.0, 0.0
    return dx / n, dy / n


def _arc_points(center, radius, start_angle, end_angle, steps=12, clockwise=False):
    if clockwise:
        if end_angle > start_angle:
            end_angle -= 2 * math.pi
    else:
        if end_angle < start_angle:
            end_angle += 2 * math.pi
    pts = []
    for i in range(1, steps):
        t = start_angle + (end_angle - start_angle) * i / steps
        pts.append((center[0] + radius * math.cos(t), center[1] + radius * math.sin(t)))
    return pts


def _stroke_outline_from_polyline(points, radius, cap_steps=10):
    pts = []
    for p in points:
        if not pts or math.hypot(p[0] - pts[-1][0], p[1] - pts[-1][1]) > 1e-9:
            pts.append(p)
    if len(pts) < 2:
        return pts

    seg_dirs = []
    for i in range(len(pts) - 1):
        dx = pts[i + 1][0] - pts[i][0]
        dy = pts[i + 1][1] - pts[i][1]
        seg_dirs.append(_unit_vec(dx, dy))

    tangents = []
    for i in range(len(pts)):
        if i == 0:
            tangents.append(seg_dirs[0])
        elif i == len(pts) - 1:
            tangents.append(seg_dirs[-1])
        else:
            ax, ay = seg_dirs[i - 1]
            bx, by = seg_dirs[i]
            tx, ty = _unit_vec(ax + bx, ay + by)
            if tx == 0 and ty == 0:
                tx, ty = bx, by
            tangents.append((tx, ty))

    left = []
    right = []
    for (px, py), (tx, ty) in zip(pts, tangents):
        nx, ny = -ty, tx
        left.append((px + nx * radius, py + ny * radius))
        right.append((px - nx * radius, py - ny * radius))

    # End cap: left_end -> right_end, around forward direction
    txe, tye = tangents[-1]
    nxe, nye = -tye, txe
    a_left_end = math.atan2(nye, nxe)
    a_right_end = math.atan2(-nye, -nxe)
    end_cap = _arc_points(pts[-1], radius, a_left_end, a_right_end, steps=cap_steps, clockwise=True)

    # Start cap: right_start -> left_start, around backward direction
    txs, tys = tangents[0]
    nxs, nys = -tys, txs
    a_right_start = math.atan2(-nys, -nxs)
    a_left_start = math.atan2(nys, nxs)
    start_cap = _arc_points(pts[0], radius, a_right_start, a_left_start, steps=cap_steps, clockwise=True)

    outline = []
    outline.extend(left)
    outline.extend(end_cap)
    outline.extend(reversed(right))
    outline.extend(start_cap)
    return outline



def draw_header_logo_dxf(msp, x, y, size, plate_h, layer="LOGOS"):
    # DXF version built as pure outlines: outer circle + stroked inner cutout contour.
    scale = size / 256.0
    cx = x + 128 * scale
    cy = y + 128 * scale
    add_dxf_circle(msp, cx, cy, 110 * scale, plate_h, layer=layer)

    # Use the same centerline as the SVG logo and convert the white stroke to a closed contour.
    centerline = svg_path_to_points(
        "M28 88 H150 A42 42 0 0 1 150 168 H28",
        tx=x,
        ty=y,
        scale=scale,
        arc_steps=48,
    )
    cutout_outline = _stroke_outline_from_polyline(centerline, radius=9.0 * scale, cap_steps=16)
    add_dxf_polyline(msp, cutout_outline, plate_h, close=True, layer=layer)

def draw_svg_dim_h(dwg, parent, x1, x2, y_obj, y_dim, text):
    sw = 0.22
    ah = 1.6
    parent.add(dwg.line((x1, y_obj), (x1, y_dim), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x2, y_obj), (x2, y_dim), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x1, y_dim), (x2, y_dim), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x1, y_dim), (x1 + ah, y_dim - 0.8), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x1, y_dim), (x1 + ah, y_dim + 0.8), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x2, y_dim), (x2 - ah, y_dim - 0.8), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x2, y_dim), (x2 - ah, y_dim + 0.8), stroke=GUIDE_COLOR, stroke_width=sw))
    draw_svg_text(dwg, parent, text, (x1 + x2) / 2, y_dim - 1.2, 2.2, anchor="middle", fill=GUIDE_COLOR)

def draw_svg_dim_v(dwg, parent, y1, y2, x_obj, x_dim, text):
    sw = 0.22
    ah = 1.6
    parent.add(dwg.line((x_obj, y1), (x_dim, y1), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_obj, y2), (x_dim, y2), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_dim, y1), (x_dim, y2), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_dim, y1), (x_dim - 0.8, y1 + ah), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_dim, y1), (x_dim + 0.8, y1 + ah), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_dim, y2), (x_dim - 0.8, y2 - ah), stroke=GUIDE_COLOR, stroke_width=sw))
    parent.add(dwg.line((x_dim, y2), (x_dim + 0.8, y2 - ah), stroke=GUIDE_COLOR, stroke_width=sw))
    draw_svg_text(dwg, parent, text, x_dim - 1.2, (y1 + y2) / 2, 2.2, anchor="middle", rotate=-90, fill=GUIDE_COLOR)

def dy(plate_h: float, y_top: float) -> float:
    return plate_h - y_top

def add_dxf_rect(msp, x, y, w, h, plate_h, layer="0"):
    pts = [(x, dy(plate_h, y)), (x + w, dy(plate_h, y)), (x + w, dy(plate_h, y + h)), (x, dy(plate_h, y + h))]
    msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})

def add_dxf_line(msp, x1, y1, x2, y2, plate_h, layer="0"):
    msp.add_line((x1, dy(plate_h, y1)), (x2, dy(plate_h, y2)), dxfattribs={"layer": layer})

def add_dxf_polyline(msp, pts, plate_h, close=False, layer="0"):
    dxf_pts = [(x, dy(plate_h, y)) for x, y in pts]
    msp.add_lwpolyline(dxf_pts, close=close, dxfattribs={"layer": layer})

def add_dxf_circle(msp, cx, cy, r, plate_h, layer="0"):
    msp.add_circle((cx, dy(plate_h, cy)), r, dxfattribs={"layer": layer})

def _text_polygons(text: str, height: float, weight: str = "normal"):
    text = clamp_text(text)
    if not text:
        return [], 0.0, 0.0, 0.0, 0.0
    fp = FontProperties(family="DejaVu Sans", weight=weight)
    tp = TextPath((0, 0), text, size=1, prop=fp)
    bb = tp.get_extents()
    if bb.width <= 0 or bb.height <= 0:
        return [], 0.0, 0.0, 0.0, 0.0
    scale = height / bb.height
    # Flip glyph Y so text stays upright after screen->DXF Y conversion.
    polys = tp.to_polygons(transform=Affine2D().scale(scale, -scale), width=0, height=0, closed_only=False)
    x0 = bb.x0 * scale
    y0 = -bb.y1 * scale
    return polys, x0, y0, bb.width * scale, bb.height * scale

def add_dxf_text_outline(msp, text, x, y, height, plate_h, align="left", valign="baseline", weight="normal", rotation=0, layer="TEXT"):
    text = clamp_text(text)
    if not text:
        return
    polys, x0, y0, bw, bh = _text_polygons(text, height, weight=weight)
    if not polys:
        return
    if align == "center":
        dx = x - (x0 + bw / 2)
    elif align == "right":
        dx = x - (x0 + bw)
    else:
        dx = x - x0
    if valign == "middle":
        dy0 = y - (y0 + bh / 2)
    elif valign == "top":
        dy0 = y - (y0 + bh)
    elif valign == "bottom":
        dy0 = y - y0
    else:
        dy0 = y
    theta = math.radians(rotation)
    ct = math.cos(theta)
    st = math.sin(theta)
    for poly in polys:
        pts = []
        for px, py in poly:
            tx = px + dx
            ty = py + dy0
            if rotation != 0:
                rx = x + (tx - x) * ct - (ty - y) * st
                ry = y + (tx - x) * st + (ty - y) * ct
                pts.append((rx, ry))
            else:
                pts.append((tx, ty))
        if len(pts) >= 3:
            add_dxf_polyline(msp, pts, plate_h, close=True, layer=layer)

def add_dxf_box_text(msp, text, x, y, w, h, base_height, plate_h, layer="TEXT", weight="normal"):
    text = clamp_text(text)
    if not text:
        return
    height = fit_font_size(text, base_height, w, padding=1.1)
    add_dxf_text_outline(msp, text, x + w / 2, y + h / 2, height, plate_h, align="center", valign="middle", weight=weight, layer=layer)

def add_dxf_multiline(msp, lines, x, y, height, line_gap, plate_h, layer="TEXT", first_bold=True):
    max_line_len = max((len(clamp_text(line)) for line in lines if clamp_text(line)), default=0)
    if max_line_len > 28:
        height *= 0.92
        line_gap *= 0.94
    if max_line_len > 36:
        height *= 0.90
        line_gap *= 0.92
    for idx, line in enumerate(lines):
        if not line:
            continue
        weight = "bold" if (idx == 0 and first_bold) else "normal"
        add_dxf_text_outline(msp, line, x, y + idx * line_gap, height, plate_h, align="left", valign="baseline", weight=weight, layer=layer)

def draw_warning_symbol_dxf(msp, x, y, w, h, plate_h, layer="LOGOS"):
    pts = [(x + 0.42 * w, y + 0.10 * h), (x + 0.67 * w, y + 0.10 * h), (x + 0.52 * w, y + 0.44 * h), (x + 0.80 * w, y + 0.44 * h), (x + 0.43 * w, y + 0.92 * h), (x + 0.51 * w, y + 0.61 * h), (x + 0.25 * w, y + 0.61 * h)]
    add_dxf_polyline(msp, pts, plate_h, close=True, layer=layer)

def draw_ce_logo_dxf(msp, x, y, w, h, plate_h, layer="LOGOS"):
    scale = min(w / 280.0, h / 200.0)
    logo_w = 280.0 * scale
    logo_h = 200.0 * scale
    tx = x + (w - logo_w) / 2
    ty = y + (h - logo_h) / 2
    pts1 = svg_path_to_points(CE_PATH_1, tx=tx, ty=ty, scale=scale, arc_steps=36)
    pts2 = svg_path_to_points(CE_PATH_2, tx=tx, ty=ty, scale=scale, arc_steps=36)
    add_dxf_polyline(msp, pts1, plate_h, close=True, layer=layer)
    add_dxf_polyline(msp, pts2, plate_h, close=True, layer=layer)

def draw_bin_logo_dxf(msp, x, y, w, h, plate_h, layer="LOGOS"):
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
    add_dxf_line(msp, cx - bw * 0.19, bar_y, cx + bw * 0.19, bar_y, plate_h, layer=layer)
    add_dxf_line(msp, cx - bw * 0.07, lid_y, cx + bw * 0.07, lid_y, plate_h, layer=layer)
    top_y = by + bh * 0.30
    bot_y = by + bh * 0.80
    top_w = bw * 0.34
    bot_w = bw * 0.26
    body_pts = [(cx - top_w / 2, top_y), (cx + top_w / 2, top_y), (cx + bot_w / 2, bot_y), (cx - bot_w / 2, bot_y)]
    add_dxf_polyline(msp, body_pts, plate_h, close=True, layer=layer)
    r = bw * 0.035
    add_dxf_circle(msp, cx - bw * 0.09, bot_y + bh * 0.05, r, plate_h, layer=layer)
    add_dxf_circle(msp, cx + bw * 0.09, bot_y + bh * 0.05, r, plate_h, layer=layer)
    add_dxf_line(msp, bx + bw * 0.16, by + bh * 0.10, bx + bw * 0.86, by + bh * 0.92, plate_h, layer=layer)
    add_dxf_line(msp, bx + bw * 0.86, by + bh * 0.10, bx + bw * 0.16, by + bh * 0.92, plate_h, layer=layer)

def draw_dxf_dim_h(msp, x1, x2, y_obj, y_dim, text, plate_h, layer="GUIDE"):
    add_dxf_line(msp, x1, y_obj, x1, y_dim, plate_h, layer=layer)
    add_dxf_line(msp, x2, y_obj, x2, y_dim, plate_h, layer=layer)
    add_dxf_line(msp, x1, y_dim, x2, y_dim, plate_h, layer=layer)
    ah = 1.6
    add_dxf_line(msp, x1, y_dim, x1 + ah, y_dim - 0.8, plate_h, layer=layer)
    add_dxf_line(msp, x1, y_dim, x1 + ah, y_dim + 0.8, plate_h, layer=layer)
    add_dxf_line(msp, x2, y_dim, x2 - ah, y_dim - 0.8, plate_h, layer=layer)
    add_dxf_line(msp, x2, y_dim, x2 - ah, y_dim + 0.8, plate_h, layer=layer)
    add_dxf_text_outline(msp, text, (x1 + x2) / 2, y_dim - 1.2, 2.2, plate_h, align="center", valign="baseline", layer=layer)

def draw_dxf_dim_v(msp, y1, y2, x_obj, x_dim, text, plate_h, layer="GUIDE"):
    add_dxf_line(msp, x_obj, y1, x_dim, y1, plate_h, layer=layer)
    add_dxf_line(msp, x_obj, y2, x_dim, y2, plate_h, layer=layer)
    add_dxf_line(msp, x_dim, y1, x_dim, y2, plate_h, layer=layer)
    ah = 1.6
    add_dxf_line(msp, x_dim, y1, x_dim - 0.8, y1 + ah, plate_h, layer=layer)
    add_dxf_line(msp, x_dim, y1, x_dim + 0.8, y1 + ah, plate_h, layer=layer)
    add_dxf_line(msp, x_dim, y2, x_dim - 0.8, y2 - ah, plate_h, layer=layer)
    add_dxf_line(msp, x_dim, y2, x_dim + 0.8, y2 - ah, plate_h, layer=layer)
    add_dxf_text_outline(msp, text, x_dim - 1.2, (y1 + y2) / 2, 2.2, plate_h, align="center", valign="middle", rotation=90, layer=layer)

def layout_values(w: float, h: float):
    sx = w / TEMPLATE_W
    sy = h / TEMPLATE_H
    X = lambda v: v * sx
    Y = lambda v: v * sy
    ix = border_offset
    iy = border_offset
    iw = w - 2 * border_offset
    ih = h - 2 * border_offset
    panel_w = X(44)
    mx = ix + panel_w + X(4)
    my = iy + Y(2.5)
    mr = ix + iw - X(4)
    main_w = mr - mx
    return {"X": X, "Y": Y, "ix": ix, "iy": iy, "iw": iw, "ih": ih, "panel_w": panel_w, "mx": mx, "my": my, "mr": mr, "main_w": main_w}

def generate_plate_svg(w: float, h: float, show_dims: bool = False, for_preview: bool = False) -> str:
    pad = 18 if show_dims and for_preview else 0
    dwg = svgwrite.Drawing(size=(f"{w}mm", f"{h}mm"), viewBox=f"{-pad} {-pad} {w + 2 * pad} {h + 2 * pad}")
    dwg.attribs["style"] = "background:white"
    l = layout_values(w, h)
    X = l["X"]
    Y = l["Y"]
    ix = l["ix"]
    iy = l["iy"]
    iw = l["iw"]
    ih = l["ih"]
    panel_w = l["panel_w"]
    mx = l["mx"]
    my = l["my"]
    mr = l["mr"]
    main_w = l["main_w"]

    outer = dwg.g(id="outer")
    dwg.add(outer)
    outer.add(dwg.rect(insert=(0.5, 0.5), size=(w - 1, h - 1), rx=corner_radius, ry=corner_radius, fill="white", stroke="black", stroke_width=0.55))
    outer.add(dwg.rect(insert=(border_offset, border_offset), size=(w - 2 * border_offset, h - 2 * border_offset), fill="none", stroke="black", stroke_width=0.3))
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
    content.add(dwg.line(start=(ix + panel_w, iy), end=(ix + panel_w, iy + ih), stroke="black", stroke_width=0.35))
    pad_zone = X(4.5)
    zone_x = ix + pad_zone
    zone_y = iy + Y(6)
    zone_w = panel_w - 2 * pad_zone
    zone_h = ih - Y(12)
    lightning_h = zone_h * 0.56 if (show_ce_logo or show_bin_logo) else zone_h * 0.78
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

    header_logo_size = Y(16.0)
    header_logo_x = mx + X(1.5)
    header_logo_y = my + Y(2.0)
    draw_header_logo_svg(dwg, content, header_logo_x, header_logo_y, header_logo_size)

    text_x = header_logo_x + header_logo_size + X(3.0)
    draw_svg_text(dwg, content, company_name, text_x, my + Y(7), Y(7.0), weight="bold", anchor="start")
    draw_svg_text(dwg, content, company_line2, text_x, my + Y(12.4), Y(3.5), anchor="start")
    draw_svg_text(dwg, content, company_line3, text_x, my + Y(17.6), Y(3.5), anchor="start")

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

    draw_svg_text(dwg, content, labels["type_short"], label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box1_x, y1 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font)
    draw_svg_text(dwg, content, labels["no_short"], box2_label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y1 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["year_label"], label_x, y2, small_font, line_gap)
    draw_box_svg(dwg, content, year_box1_x, y2 - box_h / 2, year_box_w, box_h)
    draw_svg_box_text(dwg, content, year_built, year_box1_x, y2 - box_h / 2, year_box_w, box_h, value_font)
    draw_box_svg(dwg, content, year_box2_x, y2 - box_h / 2, year_box_w, box_h)
    draw_svg_box_text(dwg, content, year_refurbished, year_box2_x, y2 - box_h / 2, year_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["current_label"], label_x, y3, small_font, line_gap)
    draw_box_svg(dwg, content, box1_x, y3 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font)
    draw_svg_text(dwg, content, "Hz", box2_label_x, y3, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y3 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["working_voltage_label"], label_x, y4, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y4, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y4 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["air_pressure_label"], label_x, y5, small_font, line_gap)
    draw_svg_text(dwg, content, "bar", unit_x - X(2), y5, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y5 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, air_pressure, single_box_x, y5 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["control_voltage_label"], label_x, y6, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y6, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y6 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font)
    draw_box_svg(dwg, content, single_box_x, y6b - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["machine_current_label"], label_x, y7, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y7, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y7 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["fuse_current_label"], label_x, y8, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y8, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y8 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font)

    draw_svg_multiline(dwg, content, labels["schematic_label"], label_x, y9, small_font, line_gap)
    draw_svg_text(dwg, content, labels["no_short"], unit_x - X(6), y9, small_font, weight="bold")
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    draw_box_svg(dwg, content, schematic_x, y9 - box_h / 2, schematic_w, box_h)
    draw_svg_box_text(dwg, content, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font)

    bottom_box_h = Y(3.4)
    bottom_box_y = iy + ih - Y(7.6)
    draw_box_svg(dwg, content, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h)
    draw_svg_text(dwg, content, footer_no, mr - X(1), iy + ih - Y(0.9), Y(1.8), anchor="end")

    if show_dims and for_preview:
        guides = dwg.g(id="guides")
        dwg.add(guides)
        draw_svg_dim_h(dwg, guides, 0, w, h, h + 9, f"{w:.1f} mm")
        draw_svg_dim_v(dwg, guides, 0, h, 0, -9, f"{h:.1f} mm")
        if hole_diameter > 0 and (show_left_holes or show_right_holes):
            draw_svg_dim_h(dwg, guides, 0, hole_offset, 0, -6, f"{hole_offset:.1f} mm")
            draw_svg_dim_v(dwg, guides, 0, hole_offset, 0, -6, f"{hole_offset:.1f} mm")
            draw_svg_text(dwg, guides, f"Ø {hole_diameter:.1f} mm", hole_offset + 4.5, hole_offset - 2.0, 2.2, fill=GUIDE_COLOR)
        draw_svg_text(dwg, guides, f"Notranji rob: {border_offset:.1f} mm", w - 3, -6.0, 2.2, anchor="end", fill=GUIDE_COLOR)

    return dwg.tostring()

def generate_plate_dxf(w: float, h: float, include_guides: bool = False) -> bytes:
    doc = ezdxf.new("R2010", setup=True)
    doc.units = 4
    ensure_layers(doc)
    msp = doc.modelspace()
    l = layout_values(w, h)
    X = l["X"]
    Y = l["Y"]
    ix = l["ix"]
    iy = l["iy"]
    iw = l["iw"]
    ih = l["ih"]
    panel_w = l["panel_w"]
    mx = l["mx"]
    my = l["my"]
    mr = l["mr"]
    main_w = l["main_w"]

    add_dxf_rect(msp, 0, 0, w, h, h, layer="BORDER")
    add_dxf_rect(msp, border_offset, border_offset, w - 2 * border_offset, h - 2 * border_offset, h, layer="BORDER")
    if hole_diameter > 0:
        r = hole_diameter / 2
        holes = []
        if show_left_holes:
            holes.extend([(hole_offset, hole_offset), (hole_offset, h - hole_offset)])
        if show_right_holes:
            holes.extend([(w - hole_offset, hole_offset), (w - hole_offset, h - hole_offset)])
        for cx, cy in holes:
            add_dxf_circle(msp, cx, cy, r, h, layer="HOLES")

    add_dxf_line(msp, ix + panel_w, iy, ix + panel_w, iy + ih, h, layer="BORDER")
    pad_zone = X(4.5)
    zone_x = ix + pad_zone
    zone_y = iy + Y(6)
    zone_w = panel_w - 2 * pad_zone
    zone_h = ih - Y(12)
    lightning_h = zone_h * 0.56 if (show_ce_logo or show_bin_logo) else zone_h * 0.78
    if show_warning_symbol:
        draw_warning_symbol_dxf(msp, zone_x, zone_y, zone_w, lightning_h, h, layer="LOGOS")
    logo_y = zone_y + lightning_h + Y(3)
    if show_ce_logo:
        ce_h = Y(12)
        draw_ce_logo_dxf(msp, zone_x + zone_w * 0.06, logo_y, zone_w * 0.88, ce_h, h, layer="LOGOS")
        logo_y += ce_h + Y(3)
    if show_bin_logo:
        bin_h = Y(12)
        draw_bin_logo_dxf(msp, zone_x + zone_w * 0.18, logo_y, zone_w * 0.64, bin_h, h, layer="LOGOS")

    header_logo_size = Y(16.0)
    header_logo_x = mx + X(1.5)
    header_logo_y = my + Y(2.0)
    draw_header_logo_dxf(msp, header_logo_x, header_logo_y, header_logo_size, h, layer="LOGOS")

    text_x = header_logo_x + header_logo_size + X(3.0)
    add_dxf_text_outline(msp, company_name, text_x, my + Y(7), Y(5.6), h, align="left", weight="bold", layer="TEXT")
    add_dxf_text_outline(msp, company_line2, text_x, my + Y(12.4), Y(2.7), h, align="left", layer="TEXT")
    add_dxf_text_outline(msp, company_line3, text_x, my + Y(17.6), Y(2.7), h, align="left", layer="TEXT")

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

    add_dxf_text_outline(msp, labels["type_short"], label_x, y1, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, box1_x, y1 - box_h / 2, box1_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font, h, layer="TEXT")
    add_dxf_text_outline(msp, labels["no_short"], box2_label_x, y1, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, box2_x, y1 - box_h / 2, box2_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["year_label"], label_x, y2, small_font, line_gap, h, layer="TEXT")
    add_dxf_rect(msp, year_box1_x, y2 - box_h / 2, year_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, year_built, year_box1_x, y2 - box_h / 2, year_box_w, box_h, value_font, h, layer="TEXT")
    add_dxf_rect(msp, year_box2_x, y2 - box_h / 2, year_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, year_refurbished, year_box2_x, y2 - box_h / 2, year_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["current_label"], label_x, y3, small_font, line_gap, h, layer="TEXT")
    add_dxf_rect(msp, box1_x, y3 - box_h / 2, box1_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font, h, layer="TEXT")
    add_dxf_text_outline(msp, "Hz", box2_label_x, y3, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, box2_x, y3 - box_h / 2, box2_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["working_voltage_label"], label_x, y4, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, "V", unit_x, y4, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, single_box_x, y4 - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["air_pressure_label"], label_x, y5, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, "bar", unit_x - X(2), y5, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, single_box_x, y5 - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, air_pressure, single_box_x, y5 - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["control_voltage_label"], label_x, y6, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, "V", unit_x, y6, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, single_box_x, y6 - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")
    add_dxf_rect(msp, single_box_x, y6b - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["machine_current_label"], label_x, y7, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, "A", unit_x, y7, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, single_box_x, y7 - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["fuse_current_label"], label_x, y8, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, "A", unit_x, y8, small_font, h, weight="bold", layer="TEXT")
    add_dxf_rect(msp, single_box_x, y8 - box_h / 2, single_box_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font, h, layer="TEXT")

    add_dxf_multiline(msp, labels["schematic_label"], label_x, y9, small_font, line_gap, h, layer="TEXT")
    add_dxf_text_outline(msp, labels["no_short"], unit_x - X(6), y9, small_font, h, weight="bold", layer="TEXT")
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    add_dxf_rect(msp, schematic_x, y9 - box_h / 2, schematic_w, box_h, h, layer="BORDER")
    add_dxf_box_text(msp, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font, h, layer="TEXT")

    bottom_box_h = Y(3.4)
    bottom_box_y = iy + ih - Y(7.6)
    add_dxf_rect(msp, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h, h, layer="BORDER")
    add_dxf_text_outline(msp, footer_no, mr - X(1), iy + ih - Y(0.9), Y(1.6), h, align="right", layer="TEXT")

    if include_guides:
        draw_dxf_dim_h(msp, 0, w, h, h + 9, f"{w:.1f} mm", h, layer="GUIDE")
        draw_dxf_dim_v(msp, 0, h, 0, -9, f"{h:.1f} mm", h, layer="GUIDE")
        if hole_diameter > 0 and (show_left_holes or show_right_holes):
            draw_dxf_dim_h(msp, 0, hole_offset, 0, -6, f"{hole_offset:.1f} mm", h, layer="GUIDE")
            draw_dxf_dim_v(msp, 0, hole_offset, 0, -6, f"{hole_offset:.1f} mm", h, layer="GUIDE")
            add_dxf_text_outline(msp, f"Ø {hole_diameter:.1f} mm", hole_offset + 4.5, hole_offset - 2.0, 2.2, h, layer="GUIDE")
        add_dxf_text_outline(msp, f"Notranji rob: {border_offset:.1f} mm", w - 3, -6.0, 2.2, h, align="right", layer="GUIDE")

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")

def make_png_from_svg(svg_text: str):
    if cairosvg is None:
        return None
    return cairosvg.svg2png(bytestring=svg_text.encode("utf-8"))

def make_jpg_from_svg(svg_text: str):
    if cairosvg is None or Image is None:
        return None
    png_bytes = make_png_from_svg(svg_text)
    if png_bytes is None:
        return None
    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=95)
    return out.getvalue()

preview_svg_content = generate_plate_svg(plate_width, plate_height, show_dims=show_dimensions, for_preview=True)
export_svg_content = generate_plate_svg(plate_width, plate_height, show_dims=False, for_preview=False)
dxf_content = generate_plate_dxf(plate_width, plate_height, include_guides=include_guides_in_dxf)
png_content = make_png_from_svg(export_svg_content)
jpg_content = make_jpg_from_svg(export_svg_content)

with right_col:
    st.subheader("Predogled")
    preview_scale = st.slider("Povečava predogleda", min_value=2.0, max_value=8.0, value=4.8, step=0.2)
    preview_h = int(plate_height * preview_scale + (140 if show_dimensions else 60))
    components.html(
        f"""
        <div style="background:#f3f4f6;padding:18px;border-radius:14px;overflow:auto;">
            <div style="display:flex;justify-content:center;">
                <div style="width:{plate_width * preview_scale}px;">
                    {preview_svg_content}
                </div>
            </div>
        </div>
        """,
        height=min(max(preview_h, 420), 1080),
        scrolling=True,
    )

    st.download_button("Prenesi SVG", data=export_svg_content, file_name="dafra_rating_plate.svg", mime="image/svg+xml", use_container_width=True)
    st.download_button("Prenesi DXF", data=dxf_content, file_name="dafra_rating_plate.dxf", mime="application/dxf", use_container_width=True)
    if png_content is not None:
        st.download_button("Prenesi PNG", data=png_content, file_name="dafra_rating_plate.png", mime="image/png", use_container_width=True)
    else:
        st.warning("PNG izvoz potrebuje CairoSVG. Namesti ga z: pip install cairosvg")
    if jpg_content is not None:
        st.download_button("Prenesi JPG", data=jpg_content, file_name="dafra_rating_plate.jpg", mime="image/jpeg", use_container_width=True)
    else:
        st.warning("JPG izvoz potrebuje CairoSVG in Pillow. Namesti ju z: pip install cairosvg pillow")

    st.info("DXF sloji: BORDER, HOLES, TEXT, LOGOS, GUIDE. Tekst v DXF je izvožen kot geometrija iz obrisov znakov, zato ga EZCAD praviloma prebere. PNG potrebuje CairoSVG, JPG pa CairoSVG in Pillow.")

    with st.expander("SVG izvor za izvoz"):
        st.code(export_svg_content[:12000], language="xml")
