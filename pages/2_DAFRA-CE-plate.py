import io

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

# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("Plate geometry")
    plate_width = st.number_input("Width [mm]", min_value=80.0, max_value=400.0, value=160.0, step=5.0)
    plate_height = st.number_input("Height [mm]", min_value=95.0, max_value=250.0, value=140.0, step=5.0)
    corner_radius = st.number_input("Corner radius [mm]", min_value=0.0, max_value=20.0, value=0.0, step=0.5)
    hole_diameter = st.number_input("Mounting hole diameter [mm]", min_value=0.0, max_value=12.0, value=3.5, step=0.5)
    hole_offset = st.number_input("Hole offset from corner [mm]", min_value=2.0, max_value=20.0, value=6.0, step=0.5)
    border_offset = st.number_input("Inner border offset [mm]", min_value=1.0, max_value=15.0, value=4.0, step=0.5)

    st.header("Visibility")
    show_left_holes = st.checkbox("Show left holes", value=True)
    show_right_holes = st.checkbox("Show right holes", value=True)
    show_warning_symbol = st.checkbox("Show warning symbol", value=True)

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
st.caption("Haulick-style machine plate layout with added Baujahr / Year built / Year refurbished row.")

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


def draw_warning_symbol_svg(dwg, parent, x, y, w, h):
    pts = [
        (x + 0.42 * w, y + 0.18 * h),
        (x + 0.67 * w, y + 0.18 * h),
        (x + 0.53 * w, y + 0.50 * h),
        (x + 0.80 * w, y + 0.50 * h),
        (x + 0.44 * w, y + 0.88 * h),
        (x + 0.52 * w, y + 0.62 * h),
        (x + 0.27 * w, y + 0.62 * h),
    ]
    parent.add(dwg.polygon(points=pts, fill="#c60000", stroke="none"))


def draw_warning_symbol_dxf(msp, x, y, w, h, plate_h):
    pts = [
        (x + 0.42 * w, y + 0.18 * h),
        (x + 0.67 * w, y + 0.18 * h),
        (x + 0.53 * w, y + 0.50 * h),
        (x + 0.80 * w, y + 0.50 * h),
        (x + 0.44 * w, y + 0.88 * h),
        (x + 0.52 * w, y + 0.62 * h),
        (x + 0.27 * w, y + 0.62 * h),
    ]
    pts_dxf = [(px, dy(plate_h, py)) for px, py in pts]
    msp.add_lwpolyline(pts_dxf, close=True)


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

# -----------------------------------------------------------------------------
# SVG
# -----------------------------------------------------------------------------
def generate_plate_svg(w: float, h: float) -> str:
    dwg = svgwrite.Drawing(size=(f"{w}mm", f"{h}mm"), viewBox=f"0 0 {w} {h}")
    dwg.attribs["style"] = "background:white"

    template_w = 160.0
    template_h = 140.0

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

    panel_w = X(44)
    content.add(
        dwg.line(
            start=(ix + panel_w, iy),
            end=(ix + panel_w, iy + ih),
            stroke="black",
            stroke_width=0.35,
        )
    )

    if show_warning_symbol:
        draw_warning_symbol_svg(dwg, content, ix + X(3), iy + Y(14), panel_w - X(6), ih - Y(24))

    mx = ix + panel_w + X(4)
    my = iy + Y(3)
    mr = ix + iw - X(4)
    main_w = mr - mx

    cx = mx + main_w / 2
    draw_svg_text(dwg, content, company_name, cx, my + Y(7), Y(7.5), weight="bold", anchor="middle")
    draw_svg_text(dwg, content, company_line2, cx, my + Y(13.0), Y(3.7), anchor="middle")
    draw_svg_text(dwg, content, company_line3, cx, my + Y(18.4), Y(3.7), anchor="middle")

    label_x = mx + X(0)
    box1_x = mx + X(22)
    box1_w = X(40)
    box2_label_x = mx + X(65)
    box2_x = mx + X(74)
    box2_w = X(27)

    unit_x = mx + X(60)
    single_box_x = mx + X(73)
    single_box_w = X(28)

    # For year row: 2 small boxes in one line
    year_box1_x = mx + X(73)
    year_box_w = X(12.5)
    year_gap = X(3.0)
    year_box2_x = year_box1_x + year_box_w + year_gap

    small_font = Y(2.6)
    value_font = Y(3.15)
    line_gap = Y(2.45)
    box_h = Y(4.6)

    y1 = my + Y(27.0)   # Type / No.
    y2 = my + Y(36.0)   # Year built / refurbished
    y3 = my + Y(47.5)   # Current / Hz
    y4 = my + Y(59.5)   # Voltage
    y5 = my + Y(71.5)   # Air pressure
    y6 = my + Y(83.5)   # Control voltage
    y6b = y6 + Y(5.0)
    y7 = my + Y(96.5)   # Machine current
    y8 = my + Y(108.5)  # Fuse current
    y9 = my + Y(120.0)  # Schematic

    # Row 1
    draw_svg_text(dwg, content, "Type", label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box1_x, y1 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font)

    draw_svg_text(dwg, content, "No.", box2_label_x, y1, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y1 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font)

    # Row 2 - Baujahr / leto izdelave + leto obnove
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

    # Row 3
    draw_svg_multiline(dwg, content, ["Stromart", "Current", "Nature du courant"], label_x, y3, small_font, line_gap)
    draw_box_svg(dwg, content, box1_x, y3 - box_h / 2, box1_w, box_h)
    draw_svg_box_text(dwg, content, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font)

    draw_svg_text(dwg, content, "Hz", box2_label_x, y3, small_font, weight="bold")
    draw_box_svg(dwg, content, box2_x, y3 - box_h / 2, box2_w, box_h)
    draw_svg_box_text(dwg, content, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font)

    # Row 4
    draw_svg_multiline(dwg, content, ["Betriebsspannung", "Working voltage", "Voltage de service"], label_x, y4, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y4, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y4 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font)

    # Row 5
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

    # Row 6
    draw_svg_multiline(dwg, content, ["Steuerspannung", "Controlling voltage", "Voltage de commande"], label_x, y6, small_font, line_gap)
    draw_svg_text(dwg, content, "V", unit_x, y6, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y6 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font)

    draw_box_svg(dwg, content, single_box_x, y6b - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font)

    # Row 7
    draw_svg_multiline(dwg, content, ["Maschine Nennstrom", "Nominal current machine", "Machine intensité nominale"], label_x, y7, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y7, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y7 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font)

    # Row 8
    draw_svg_multiline(dwg, content, ["Sicherungs-Nennstrom", "Nominal current fuses", "Intensité de protection nominale"], label_x, y8, small_font, line_gap)
    draw_svg_text(dwg, content, "A", unit_x, y8, small_font, weight="bold")
    draw_box_svg(dwg, content, single_box_x, y8 - box_h / 2, single_box_w, box_h)
    draw_svg_box_text(dwg, content, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font)

    # Row 9
    draw_svg_multiline(dwg, content, ["Schaltplan", "Schematic", "Schéma de connexions"], label_x, y9, small_font, line_gap)
    draw_svg_text(dwg, content, "No.", unit_x - X(6), y9, small_font, weight="bold")
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    draw_box_svg(dwg, content, schematic_x, y9 - box_h / 2, schematic_w, box_h)
    draw_svg_box_text(dwg, content, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font)

    bottom_box_h = Y(3.2)
    bottom_box_y = iy + ih - Y(7.0)
    draw_box_svg(dwg, content, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h)

    draw_svg_text(dwg, content, footer_no, mr - X(1), iy + ih - Y(0.8), Y(1.9), anchor="end")

    return dwg.tostring()

# -----------------------------------------------------------------------------
# DXF
# -----------------------------------------------------------------------------
def generate_plate_dxf(w: float, h: float) -> bytes:
    doc = ezdxf.new("R2010", setup=True)
    doc.units = 4
    msp = doc.modelspace()

    template_w = 160.0
    template_h = 140.0

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
    msp.add_line((ix + panel_w, dy(h, iy)), (ix + panel_w, dy(h, iy + ih)))

    if show_warning_symbol:
        draw_warning_symbol_dxf(msp, ix + X(3), iy + Y(14), panel_w - X(6), ih - Y(24), h)

    mx = ix + panel_w + X(4)
    my = iy + Y(3)
    mr = ix + iw - X(4)
    main_w = mr - mx

    cx = mx + main_w / 2
    add_dxf_text(msp, company_name, cx, my + Y(7), Y(5.8), h, align=TextEntityAlignment.MIDDLE_CENTER)
    add_dxf_text(msp, company_line2, cx, my + Y(13.0), Y(2.8), h, align=TextEntityAlignment.MIDDLE_CENTER)
    add_dxf_text(msp, company_line3, cx, my + Y(18.4), Y(2.8), h, align=TextEntityAlignment.MIDDLE_CENTER)

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

    small_font = Y(2.2)
    value_font = Y(2.55)
    line_gap = Y(2.45)
    box_h = Y(4.6)

    y1 = my + Y(27.0)
    y2 = my + Y(36.0)
    y3 = my + Y(47.5)
    y4 = my + Y(59.5)
    y5 = my + Y(71.5)
    y6 = my + Y(83.5)
    y6b = y6 + Y(5.0)
    y7 = my + Y(96.5)
    y8 = my + Y(108.5)
    y9 = my + Y(120.0)

    # Row 1
    add_dxf_text(msp, "Type", label_x, y1, small_font, h)
    add_dxf_rect(msp, box1_x, y1 - box_h / 2, box1_w, box_h, h)
    add_dxf_box_text(msp, type_text, box1_x, y1 - box_h / 2, box1_w, box_h, value_font, h)

    add_dxf_text(msp, "No.", box2_label_x, y1, small_font, h)
    add_dxf_rect(msp, box2_x, y1 - box_h / 2, box2_w, box_h, h)
    add_dxf_box_text(msp, no_text, box2_x, y1 - box_h / 2, box2_w, box_h, value_font, h)

    # Row 2
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

    # Row 3
    add_dxf_multiline(msp, ["Stromart", "Current", "Nature du courant"], label_x, y3, small_font, line_gap, h)
    add_dxf_rect(msp, box1_x, y3 - box_h / 2, box1_w, box_h, h)
    add_dxf_box_text(msp, current_type, box1_x, y3 - box_h / 2, box1_w, box_h, value_font, h)

    add_dxf_text(msp, "Hz", box2_label_x, y3, small_font, h)
    add_dxf_rect(msp, box2_x, y3 - box_h / 2, box2_w, box_h, h)
    add_dxf_box_text(msp, hz_text, box2_x, y3 - box_h / 2, box2_w, box_h, value_font, h)

    # Row 4
    add_dxf_multiline(msp, ["Betriebsspannung", "Working voltage", "Voltage de service"], label_x, y4, small_font, line_gap, h)
    add_dxf_text(msp, "V", unit_x, y4, small_font, h)
    add_dxf_rect(msp, single_box_x, y4 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, working_voltage, single_box_x, y4 - box_h / 2, single_box_w, box_h, value_font, h)

    # Row 5
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

    # Row 6
    add_dxf_multiline(msp, ["Steuerspannung", "Controlling voltage", "Voltage de commande"], label_x, y6, small_font, line_gap, h)
    add_dxf_text(msp, "V", unit_x, y6, small_font, h)
    add_dxf_rect(msp, single_box_x, y6 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, f"{control_voltage_dc} =", single_box_x, y6 - box_h / 2, single_box_w, box_h, value_font, h)

    add_dxf_rect(msp, single_box_x, y6b - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, f"{control_voltage_ac} ~", single_box_x, y6b - box_h / 2, single_box_w, box_h, value_font, h)

    # Row 7
    add_dxf_multiline(msp, ["Maschine Nennstrom", "Nominal current machine", "Machine intensité nominale"], label_x, y7, small_font, line_gap, h)
    add_dxf_text(msp, "A", unit_x, y7, small_font, h)
    add_dxf_rect(msp, single_box_x, y7 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, machine_current, single_box_x, y7 - box_h / 2, single_box_w, box_h, value_font, h)

    # Row 8
    add_dxf_multiline(msp, ["Sicherungs-Nennstrom", "Nominal current fuses", "Intensité de protection nominale"], label_x, y8, small_font, line_gap, h)
    add_dxf_text(msp, "A", unit_x, y8, small_font, h)
    add_dxf_rect(msp, single_box_x, y8 - box_h / 2, single_box_w, box_h, h)
    add_dxf_box_text(msp, fuse_current, single_box_x, y8 - box_h / 2, single_box_w, box_h, value_font, h)

    # Row 9
    add_dxf_multiline(msp, ["Schaltplan", "Schematic", "Schéma de connexions"], label_x, y9, small_font, line_gap, h)
    add_dxf_text(msp, "No.", unit_x - X(6), y9, small_font, h)
    schematic_x = single_box_x - X(12)
    schematic_w = single_box_w + X(12)
    add_dxf_rect(msp, schematic_x, y9 - box_h / 2, schematic_w, box_h, h)
    add_dxf_box_text(msp, schematic_no, schematic_x, y9 - box_h / 2, schematic_w, box_h, value_font, h)

    bottom_box_h = Y(3.2)
    bottom_box_y = iy + ih - Y(7.0)
    add_dxf_rect(msp, mx + X(4), bottom_box_y, main_w - X(6), bottom_box_h, h)

    add_dxf_text(msp, footer_no, mr - X(1), iy + ih - Y(0.8), Y(1.8), h, align=TextEntityAlignment.RIGHT)

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


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

    st.info("Added Baujahr / Year built / Year refurbished row with two boxes in one line.")