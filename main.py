import io
from typing import List

import ezdxf
from ezdxf.enums import TextEntityAlignment
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import svgwrite


st.set_page_config(
    page_title="CE Rating Plate Designer",
    page_icon="🏷️",
    layout="wide",
)


DEFAULT_HEADER_ROWS = [
    {"label": "Company", "value": "FOLTECH ENGINEERING", "unit": ""},
    {"label": "Manufacturer", "value": "GILEAD SCIENCES ULC", "unit": ""},
    {"label": "Equipment Type", "value": "Pressure Vessel", "unit": ""},
    {"label": "Client / Site", "value": "Biologics Facility", "unit": ""},
    {"label": "Serial No.", "value": "FOL-0001", "unit": ""},
    {"label": "Year Built", "value": "2026", "unit": ""},
    {"label": "Design Code", "value": "ASME VIII Div. 1", "unit": ""},
    {"label": "EC Directive", "value": "2014/68/EU", "unit": ""},
]

DEFAULT_SPEC_ROWS = [
    {"label": "Design Pressure", "value": "10", "unit": "bar(g)"},
    {"label": "Design Temperature", "value": "120", "unit": "°C"},
    {"label": "Hydro Test Pressure", "value": "15", "unit": "bar(g)"},
    {"label": "Max Working Temp", "value": "110", "unit": "°C"},
    {"label": "Max Working Pressure", "value": "8", "unit": "bar(g)"},
    {"label": "Corrosion Allowance", "value": "1.5", "unit": "mm"},
    {"label": "Capacity", "value": "150", "unit": "liters"},
    {"label": "Material", "value": "SS316L", "unit": ""},
    {"label": "PWHT", "value": "No", "unit": ""},
    {"label": "Inspection Date", "value": "2026-03-19", "unit": ""},
]


def init_state() -> None:
    if "header_rows" not in st.session_state:
        st.session_state.header_rows = pd.DataFrame(DEFAULT_HEADER_ROWS)
    if "spec_rows" not in st.session_state:
        st.session_state.spec_rows = pd.DataFrame(DEFAULT_SPEC_ROWS)


init_state()


st.title("CE Marking & Rating Plate Designer")
st.caption("Edit the fields on the left, see the plate update live, then export SVG or DXF.")


with st.sidebar:
    st.header("Plate geometry")
    plate_width = st.number_input("Width [mm]", min_value=60.0, max_value=400.0, value=160.0, step=5.0)
    plate_height = st.number_input("Height [mm]", min_value=40.0, max_value=250.0, value=100.0, step=5.0)
    corner_radius = st.number_input("Corner radius [mm]", min_value=0.0, max_value=20.0, value=4.0, step=0.5)
    hole_diameter = st.number_input("Mounting hole diameter [mm]", min_value=0.0, max_value=12.0, value=3.5, step=0.5)
    hole_offset = st.number_input("Hole offset from corner [mm]", min_value=2.0, max_value=20.0, value=6.0, step=0.5)
    border_offset = st.number_input("Inner border offset [mm]", min_value=1.0, max_value=15.0, value=4.0, step=0.5)

    st.header("Branding")
    brand_name = st.text_input("Brand name", value="foltech")
    brand_subtitle = st.text_input("Subtitle", value="ENGINEERING")
    top_meta = st.text_area(
        "Top-right small text",
        value="TEG GmbH\nMaschinen- und Anlagenbau\nModel: CR-07 / PS 7 bar\nMade in EU",
        height=110,
    )
    plate_title = st.text_input("Main title", value="GILEAD SCIENCES ULC")

    st.header("Compliance / right column")
    serial_text = st.text_input("Serial / ID", value="0039")
    standard_text = st.text_input("Standard / note", value="ISO / UNI 2020")
    show_ce = st.checkbox("Show CE mark", value=True)
    show_left_holes = st.checkbox("Show left holes", value=True)
    show_right_holes = st.checkbox("Show right holes", value=True)


def clamp_text(value) -> str:
    if value is None:
        return ""
    return str(value)


left_col, right_col = st.columns([1.1, 1.2], gap="large")

with left_col:
    st.subheader("Header fields")
    st.session_state.header_rows = st.data_editor(
        st.session_state.header_rows,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "label": st.column_config.TextColumn("Label"),
            "value": st.column_config.TextColumn("Value"),
            "unit": st.column_config.TextColumn("Unit"),
        },
        key="header_editor",
    )

    st.subheader("Specification fields")
    st.session_state.spec_rows = st.data_editor(
        st.session_state.spec_rows,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "label": st.column_config.TextColumn("Label"),
            "value": st.column_config.TextColumn("Value"),
            "unit": st.column_config.TextColumn("Unit"),
        },
        key="spec_editor",
    )


def cleaned_rows(df: pd.DataFrame) -> List[dict]:
    rows = []
    for row in df.fillna("").to_dict(orient="records"):
        label = clamp_text(row.get("label", "")).strip()
        value = clamp_text(row.get("value", "")).strip()
        unit = clamp_text(row.get("unit", "")).strip()
        if label or value or unit:
            rows.append({"label": label, "value": value, "unit": unit})
    return rows


header_rows = cleaned_rows(st.session_state.header_rows)
spec_rows = cleaned_rows(st.session_state.spec_rows)


if not header_rows:
    st.warning("Add at least one header row.")
if not spec_rows:
    st.warning("Add at least one specification row.")


def draw_svg_text(dwg, parent, text, x, y, size, weight="normal", anchor="start", rotate=None):
    if not text:
        return
    attrs = {
        "insert": (x, y),
        "font_size": size,
        "font_family": "Arial, Helvetica, sans-serif",
        "font_weight": weight,
        "text_anchor": anchor,
        "fill": "black",
    }
    el = dwg.text(text, **attrs)
    if rotate is not None:
        el.rotate(rotate, center=(x, y))
    parent.add(el)


# DXF helpers ---------------------------------------------------------------

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
    if not text:
        return
    entity = msp.add_text(text, dxfattribs={"height": max(height, 1.0), "rotation": rotation})
    entity.set_placement((x, dy(plate_h, y)), align=align)



def generate_plate_svg(
    w: float,
    h: float,
    header_rows: List[dict],
    spec_rows: List[dict],
) -> str:
    dwg = svgwrite.Drawing(size=(f"{w}mm", f"{h}mm"), viewBox=f"0 0 {w} {h}")
    dwg.attribs["style"] = "background:white"

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
            stroke_width=0.6,
        )
    )
    outer.add(
        dwg.rect(
            insert=(border_offset, border_offset),
            size=(w - 2 * border_offset, h - 2 * border_offset),
            rx=max(corner_radius - border_offset / 2, 0),
            ry=max(corner_radius - border_offset / 2, 0),
            fill="none",
            stroke="black",
            stroke_width=0.25,
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

    margin = border_offset + 3
    right_col_w = max(24.0, w * 0.13)
    brand_zone_w = max(38.0, w * 0.23)
    main_x = margin + brand_zone_w + 2
    main_w = w - main_x - right_col_w - margin
    top_y = margin

    # Branding block
    draw_svg_text(dwg, content, brand_name, margin, top_y + 8, 11, weight="bold")
    draw_svg_text(dwg, content, brand_subtitle, margin + 1, top_y + 16, 5.6, weight="bold")

    meta_lines = [line for line in top_meta.splitlines() if line.strip()]
    meta_x = margin + brand_zone_w * 0.42
    for i, line in enumerate(meta_lines[:5]):
        draw_svg_text(dwg, content, line, meta_x, top_y + 3 + i * 3.6, 2.2)

    title_h = 7
    title_y = top_y + 5
    content.add(dwg.rect(insert=(main_x, title_y), size=(main_w, title_h), fill="none", stroke="black", stroke_width=0.35))
    draw_svg_text(dwg, content, plate_title, main_x + main_w / 2, title_y + 4.8, 4.2, weight="bold", anchor="middle")

    # Header rows block
    header_y = title_y + title_h + 1.5
    header_label_w = main_w * 0.34
    header_value_w = main_w - header_label_w
    header_row_h = 4.6
    max_header_rows_fit = max(1, int((h * 0.32) / header_row_h))
    shown_headers = header_rows[:max_header_rows_fit]
    for idx, row in enumerate(shown_headers):
        y = header_y + idx * header_row_h
        content.add(dwg.rect(insert=(main_x, y), size=(header_label_w, header_row_h), fill="none", stroke="black", stroke_width=0.25))
        content.add(dwg.rect(insert=(main_x + header_label_w, y), size=(header_value_w, header_row_h), fill="none", stroke="black", stroke_width=0.25))
        draw_svg_text(dwg, content, row["label"], main_x + 1.2, y + 3.2, 2.25, weight="bold")
        val_text = row["value"] if not row.get("unit") else f"{row['value']}  {row['unit']}"
        draw_svg_text(dwg, content, val_text, main_x + header_label_w + 1.2, y + 3.2, 2.25)

    body_y = header_y + len(shown_headers) * header_row_h + 3
    body_h = h - body_y - margin
    n_specs = max(len(spec_rows), 1)
    spec_row_h = max(4.2, min(6.0, body_h / max(n_specs, 1)))
    table_h = spec_row_h * n_specs

    label_w = main_w * 0.49
    value_w = main_w * 0.28
    unit_w = main_w - label_w - value_w

    for idx, row in enumerate(spec_rows):
        y = body_y + idx * spec_row_h
        content.add(dwg.rect(insert=(main_x, y), size=(label_w, spec_row_h), fill="none", stroke="black", stroke_width=0.25))
        content.add(dwg.rect(insert=(main_x + label_w, y), size=(value_w, spec_row_h), fill="none", stroke="black", stroke_width=0.25))
        content.add(dwg.rect(insert=(main_x + label_w + value_w, y), size=(unit_w, spec_row_h), fill="none", stroke="black", stroke_width=0.25))
        draw_svg_text(dwg, content, row["label"], main_x + 1.1, y + spec_row_h * 0.68, 2.3, weight="bold")
        draw_svg_text(dwg, content, row["value"], main_x + label_w + 1.1, y + spec_row_h * 0.68, 2.3)
        draw_svg_text(dwg, content, row["unit"], main_x + label_w + value_w + 1.1, y + spec_row_h * 0.68, 2.3)

    # Right-side compliance column
    right_x = w - margin - right_col_w
    right_y = title_y
    ce_box_h = 14
    serial_box_h = 18
    std_box_h = 12
    spare_h = max(h - right_y - margin - ce_box_h - serial_box_h - std_box_h - 3, 8)

    content.add(dwg.rect(insert=(right_x, right_y), size=(right_col_w, serial_box_h), fill="none", stroke="black", stroke_width=0.35))
    draw_svg_text(dwg, content, serial_text, right_x + right_col_w / 2, right_y + serial_box_h / 2 + 1.5, 6.4, weight="bold", anchor="middle", rotate=90)

    ce_y = right_y + serial_box_h + 1.5
    content.add(dwg.rect(insert=(right_x, ce_y), size=(right_col_w, ce_box_h), fill="none", stroke="black", stroke_width=0.35))
    if show_ce:
        draw_svg_text(dwg, content, "CE", right_x + right_col_w / 2, ce_y + ce_box_h / 2 + 1.5, 8.0, weight="bold", anchor="middle", rotate=90)

    std_y = ce_y + ce_box_h + 1.5
    content.add(dwg.rect(insert=(right_x, std_y), size=(right_col_w, std_box_h), fill="none", stroke="black", stroke_width=0.35))
    draw_svg_text(dwg, content, standard_text, right_x + right_col_w / 2, std_y + std_box_h / 2 + 0.8, 2.8, anchor="middle", rotate=90)

    spare_y = std_y + std_box_h + 1.5
    content.add(dwg.rect(insert=(right_x, spare_y), size=(right_col_w, spare_h), fill="none", stroke="black", stroke_width=0.35))
    draw_svg_text(dwg, content, "RATING PLATE", right_x + right_col_w / 2, spare_y + spare_h / 2, 3.3, weight="bold", anchor="middle", rotate=90)

    if len(header_rows) > len(shown_headers):
        draw_svg_text(
            dwg,
            content,
            f"+{len(header_rows) - len(shown_headers)} more header rows not shown",
            margin,
            h - 3,
            2.1,
        )

    return dwg.tostring()



def generate_plate_dxf(
    w: float,
    h: float,
    header_rows: List[dict],
    spec_rows: List[dict],
) -> bytes:
    doc = ezdxf.new("R2010", setup=True)
    doc.units = 4  # millimeters
    msp = doc.modelspace()

    # Outline
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

    margin = border_offset + 3
    right_col_w = max(24.0, w * 0.13)
    brand_zone_w = max(38.0, w * 0.23)
    main_x = margin + brand_zone_w + 2
    main_w = w - main_x - right_col_w - margin
    top_y = margin

    # Branding
    add_dxf_text(msp, brand_name, margin, top_y + 8, 6.5, h, align=TextEntityAlignment.LEFT)
    add_dxf_text(msp, brand_subtitle, margin + 1, top_y + 16, 3.5, h, align=TextEntityAlignment.LEFT)

    meta_lines = [line for line in top_meta.splitlines() if line.strip()]
    meta_x = margin + brand_zone_w * 0.42
    for i, line in enumerate(meta_lines[:5]):
        add_dxf_text(msp, line, meta_x, top_y + 3 + i * 3.6, 1.8, h, align=TextEntityAlignment.LEFT)

    title_h = 7
    title_y = top_y + 5
    add_dxf_rect(msp, main_x, title_y, main_w, title_h, h)
    add_dxf_text(msp, plate_title, main_x + main_w / 2, title_y + 4.2, 2.8, h, align=TextEntityAlignment.MIDDLE_CENTER)

    header_y = title_y + title_h + 1.5
    header_label_w = main_w * 0.34
    header_value_w = main_w - header_label_w
    header_row_h = 4.6
    max_header_rows_fit = max(1, int((h * 0.32) / header_row_h))
    shown_headers = header_rows[:max_header_rows_fit]
    for idx, row in enumerate(shown_headers):
        y = header_y + idx * header_row_h
        add_dxf_rect(msp, main_x, y, header_label_w, header_row_h, h)
        add_dxf_rect(msp, main_x + header_label_w, y, header_value_w, header_row_h, h)
        add_dxf_text(msp, row["label"], main_x + 1.2, y + 3.0, 1.8, h)
        val_text = row["value"] if not row.get("unit") else f"{row['value']}  {row['unit']}"
        add_dxf_text(msp, val_text, main_x + header_label_w + 1.2, y + 3.0, 1.8, h)

    body_y = header_y + len(shown_headers) * header_row_h + 3
    body_h = h - body_y - margin
    n_specs = max(len(spec_rows), 1)
    spec_row_h = max(4.2, min(6.0, body_h / max(n_specs, 1)))

    label_w = main_w * 0.49
    value_w = main_w * 0.28
    unit_w = main_w - label_w - value_w

    for idx, row in enumerate(spec_rows):
        y = body_y + idx * spec_row_h
        add_dxf_rect(msp, main_x, y, label_w, spec_row_h, h)
        add_dxf_rect(msp, main_x + label_w, y, value_w, spec_row_h, h)
        add_dxf_rect(msp, main_x + label_w + value_w, y, unit_w, spec_row_h, h)
        add_dxf_text(msp, row["label"], main_x + 1.1, y + spec_row_h * 0.68, 1.8, h)
        add_dxf_text(msp, row["value"], main_x + label_w + 1.1, y + spec_row_h * 0.68, 1.8, h)
        add_dxf_text(msp, row["unit"], main_x + label_w + value_w + 1.1, y + spec_row_h * 0.68, 1.8, h)

    right_x = w - margin - right_col_w
    right_y = title_y
    ce_box_h = 14
    serial_box_h = 18
    std_box_h = 12
    spare_h = max(h - right_y - margin - ce_box_h - serial_box_h - std_box_h - 3, 8)

    add_dxf_rect(msp, right_x, right_y, right_col_w, serial_box_h, h)
    add_dxf_text(msp, serial_text, right_x + right_col_w / 2, right_y + serial_box_h / 2, 5.2, h, align=TextEntityAlignment.MIDDLE_CENTER, rotation=90)

    ce_y = right_y + serial_box_h + 1.5
    add_dxf_rect(msp, right_x, ce_y, right_col_w, ce_box_h, h)
    if show_ce:
        add_dxf_text(msp, "CE", right_x + right_col_w / 2, ce_y + ce_box_h / 2, 7.0, h, align=TextEntityAlignment.MIDDLE_CENTER, rotation=90)

    std_y = ce_y + ce_box_h + 1.5
    add_dxf_rect(msp, right_x, std_y, right_col_w, std_box_h, h)
    add_dxf_text(msp, standard_text, right_x + right_col_w / 2, std_y + std_box_h / 2, 2.0, h, align=TextEntityAlignment.MIDDLE_CENTER, rotation=90)

    spare_y = std_y + std_box_h + 1.5
    add_dxf_rect(msp, right_x, spare_y, right_col_w, spare_h, h)
    add_dxf_text(msp, "RATING PLATE", right_x + right_col_w / 2, spare_y + spare_h / 2, 2.4, h, align=TextEntityAlignment.MIDDLE_CENTER, rotation=90)

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


svg_content = generate_plate_svg(plate_width, plate_height, header_rows, spec_rows)
dxf_content = generate_plate_dxf(plate_width, plate_height, header_rows, spec_rows)

with right_col:
    st.subheader("Live preview")
    preview_scale = st.slider("Preview scale", min_value=2.0, max_value=8.0, value=4.2, step=0.2)
    preview_h = int(plate_height * preview_scale + 50)
    components.html(
        f"""
        <div style='background:#f3f4f6;padding:18px;border-radius:14px;overflow:auto;'>
            <div style='display:flex;justify-content:center;'>
                <div style='width:{plate_width * preview_scale}px;'>
                    {svg_content}
                </div>
            </div>
        </div>
        """,
        height=min(max(preview_h, 350), 950),
        scrolling=True,
    )

    st.download_button(
        "Download SVG",
        data=svg_content,
        file_name="ce_rating_plate.svg",
        mime="image/svg+xml",
        use_container_width=True,
    )
    st.download_button(
        "Download DXF",
        data=dxf_content,
        file_name="ce_rating_plate.dxf",
        mime="application/dxf",
        use_container_width=True,
    )

    st.info(
        "DXF export is line/text based, so it is suitable as a starting layout for CAD or laser workflows."
    )

    with st.expander("SVG source"):
        st.code(svg_content[:12000], language="xml")
