import io
import re
import html
import zipfile
from collections import Counter

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import ezdxf
from ezdxf.enums import TextEntityAlignment

from matplotlib.textpath import TextPath
from matplotlib.font_manager import FontProperties
from matplotlib.transforms import Affine2D
from matplotlib.path import Path as MplPath

# ------------------------------------------------------------
# Page
# ------------------------------------------------------------
st.set_page_config(
    page_title="Laser Label Generator",
    page_icon="🏷️",
    layout="wide",
)

# ------------------------------------------------------------
# Main geometry
# ------------------------------------------------------------
LABEL_W = 78.5   # mm
LABEL_H = 21.0   # mm

DEFAULT_OWNER = "Stihl Group"
DEFAULT_TOOL = "89193"
DEFAULT_PART = "Steckzunge BA13-431-2100-A (id. 33193)"

DEFAULT_MODE = "Anodized aluminium (negative)"
DEFAULT_HOLE_DIA = 3.2
DEFAULT_HOLE_OFFSET = 4.3
DEFAULT_CORNER_R = 2.2
DEFAULT_BORDER_OFFSET = 0.25

DEFAULT_LEFT_X = 8.4
DEFAULT_LEFT_W = 21.5
DEFAULT_RIGHT_X = 34.2

DEFAULT_ROW1_Y = 3.0
DEFAULT_ROW2_Y = 8.0
DEFAULT_ROW3_Y = 13.2

ROW1_H = 4.0
ROW2_H = 4.0
ROW3_H = 3.8

DEFAULT_FS1 = 2.8
DEFAULT_FS2 = 2.8
DEFAULT_FS3 = 1.8

MIN_TEXT_HEIGHT_MM = 1.2
RIGHT_MARGIN = 2.2
VECTOR_FONT_FAMILY = "DejaVu Sans Condensed"

ALLOWED_MODES = ["Normal", "Anodized aluminium (negative)"]

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def safe_filename(text: str) -> str:
    text = (text or "").strip()
    if not text:
        return "laser_label"
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    text = text.strip("._")
    return text[:80] or "laser_label"


def escape_xml(text: str) -> str:
    return html.escape(text or "", quote=False)


def get_mode_colors(mode: str):
    if mode == "Anodized aluminium (negative)":
        return {
            "plate_fill": "black",
            "text_fill": "white",
            "plate_stroke": "black",
            "hole_fill": "white",
            "guide_stroke": "#53b7ff",
            "preview_bg": "#dcdcdc",
        }
    return {
        "plate_fill": "white",
        "text_fill": "black",
        "plate_stroke": "black",
        "hole_fill": "white",
        "guide_stroke": "#1f77b4",
        "preview_bg": "#f3f3f3",
    }


def build_rows(owner: str, tool_number: str, part_desc: str):
    return [
        ("Property of:", owner),
        ("Tool number:", tool_number),
        ("Part description:", part_desc),
    ]


def get_font_prop(weight: str):
    return FontProperties(family=VECTOR_FONT_FAMILY, weight=weight)


def make_base_text_path(text: str, weight: str):
    if not text:
        return None, None
    tp = TextPath((0, 0), text, size=1, prop=get_font_prop(weight))
    bbox = tp.get_extents()
    if bbox.width <= 0 or bbox.height <= 0:
        return None, None
    return tp, bbox


def fit_text_for_box(text: str, desired_h_mm: float, box_w_mm: float, box_h_mm: float, weight: str):
    raw = (text or "").strip()
    if not raw:
        return "", None, None, 1.0, 0.0

    for cut in range(0, len(raw) + 1):
        if cut == 0:
            candidate = raw
        else:
            candidate = raw[: max(0, len(raw) - cut)].rstrip() + "..."

        base_path, bbox = make_base_text_path(candidate, weight)
        if base_path is None or bbox is None:
            continue

        target_h = min(desired_h_mm, box_h_mm)
        scale = target_h / bbox.height

        if bbox.width * scale <= box_w_mm:
            return candidate, base_path, bbox, scale, bbox.height * scale

        scale_w = box_w_mm / bbox.width
        height_at_width_fit = bbox.height * scale_w

        if height_at_width_fit >= MIN_TEXT_HEIGHT_MM and height_at_width_fit <= box_h_mm:
            return candidate, base_path, bbox, scale_w, height_at_width_fit

    fallback = "..."
    base_path, bbox = make_base_text_path(fallback, weight)
    if base_path is None or bbox is None:
        return "", None, None, 1.0, 0.0

    target_h = min(desired_h_mm, box_h_mm)
    scale = min(target_h / bbox.height, box_w_mm / bbox.width)
    return fallback, base_path, bbox, scale, bbox.height * scale


def place_path_in_box(base_path, scale, box_x, box_y, box_w, box_h, align="left", pad_x=0.0):
    path = base_path.transformed(Affine2D().scale(scale, -scale))
    bbox = path.get_extents()

    if align == "center":
        tx = box_x + (box_w - bbox.width) / 2.0 - bbox.x0
    elif align == "right":
        tx = box_x + box_w - pad_x - bbox.x1
    else:
        tx = box_x + pad_x - bbox.x0

    ty = box_y + (box_h - bbox.height) / 2.0 - bbox.y0
    return path.transformed(Affine2D().translate(tx, ty))


def mpl_path_to_svg_d(path_obj):
    if path_obj is None:
        return ""

    out = []
    for verts, code in path_obj.iter_segments():
        if code == MplPath.MOVETO:
            x, y = verts
            out.append(f"M {x:.4f} {y:.4f}")
        elif code == MplPath.LINETO:
            x, y = verts
            out.append(f"L {x:.4f} {y:.4f}")
        elif code == MplPath.CURVE3:
            x1, y1, x2, y2 = verts
            out.append(f"Q {x1:.4f} {y1:.4f} {x2:.4f} {y2:.4f}")
        elif code == MplPath.CURVE4:
            x1, y1, x2, y2, x3, y3 = verts
            out.append(f"C {x1:.4f} {y1:.4f} {x2:.4f} {y2:.4f} {x3:.4f} {y3:.4f}")
        elif code == MplPath.CLOSEPOLY:
            out.append("Z")
    return " ".join(out)


def add_rounded_rect_dxf(msp, x, y, w, h, r, layer):
    r = max(0.0, min(r, w / 2.0, h / 2.0))
    if r <= 0:
        pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
        msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})
        return

    msp.add_line((x + r, y), (x + w - r, y), dxfattribs={"layer": layer})
    msp.add_line((x + w, y + r), (x + w, y + h - r), dxfattribs={"layer": layer})
    msp.add_line((x + w - r, y + h), (x + r, y + h), dxfattribs={"layer": layer})
    msp.add_line((x, y + h - r), (x, y + r), dxfattribs={"layer": layer})

    msp.add_arc((x + w - r, y + r), r, 270, 360, dxfattribs={"layer": layer})
    msp.add_arc((x + w - r, y + h - r), r, 0, 90, dxfattribs={"layer": layer})
    msp.add_arc((x + r, y + h - r), r, 90, 180, dxfattribs={"layer": layer})
    msp.add_arc((x + r, y + r), r, 180, 270, dxfattribs={"layer": layer})


def row_box(top_y, height):
    return {"y": top_y, "h": height}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace(" ", "_").replace("-", "_")
        for c in df.columns
    ]
    return df


def parse_mode(value, default_mode):
    v = str(value).strip() if pd.notna(value) else ""
    if not v:
        return default_mode

    lookup = {
        "normal": "Normal",
        "standard": "Normal",
        "default": "Normal",
        "negative": "Anodized aluminium (negative)",
        "anodized": "Anodized aluminium (negative)",
        "anodized_aluminium": "Anodized aluminium (negative)",
        "anodized aluminium": "Anodized aluminium (negative)",
        "anodized aluminium (negative)": "Anodized aluminium (negative)",
        "black": "Anodized aluminium (negative)",
    }
    return lookup.get(v.lower(), default_mode if v not in ALLOWED_MODES else v)


def parse_optional_float(value, default_value):
    if pd.isna(value):
        return default_value
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return default_value


def load_tabular_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported file type. Use CSV or XLSX.")


def make_unique_base_names(names):
    counts = Counter()
    result = []
    for n in names:
        counts[n] += 1
        if counts[n] == 1:
            result.append(n)
        else:
            result.append(f"{n}_{counts[n]}")
    return result


def make_preview_svg(svg_text: str, label_w_mm: float, label_h_mm: float, px_per_mm: float) -> str:
    preview_w_px = label_w_mm * px_per_mm
    preview_h_px = label_h_mm * px_per_mm

    svg_preview = re.sub(r'width="[^"]+"', f'width="{preview_w_px}px"', svg_text, count=1)
    svg_preview = re.sub(r'height="[^"]+"', f'height="{preview_h_px}px"', svg_preview, count=1)

    return svg_preview


# ------------------------------------------------------------
# SVG generation with vector outlines
# ------------------------------------------------------------
def generate_svg(
    owner,
    tool_number,
    part_desc,
    mode,
    hole_dia,
    hole_offset,
    corner_r,
    border_offset,
    left_x,
    left_w,
    right_x,
    row1_y,
    row2_y,
    row3_y,
    fs1,
    fs2,
    fs3,
    show_guides=False,
    show_border=True,
):
    colors = get_mode_colors(mode)
    rows = build_rows(owner, tool_number, part_desc)

    right_w = LABEL_W - right_x - RIGHT_MARGIN
    left_boxes = [
        {"x": left_x, "w": left_w, "row": row_box(row1_y, ROW1_H)},
        {"x": left_x, "w": left_w, "row": row_box(row2_y, ROW2_H)},
        {"x": left_x, "w": left_w, "row": row_box(row3_y, ROW3_H)},
    ]
    right_boxes = [
        {"x": right_x, "w": right_w, "row": row_box(row1_y, ROW1_H)},
        {"x": right_x, "w": right_w, "row": row_box(row2_y, ROW2_H)},
        {"x": right_x, "w": right_w, "row": row_box(row3_y, ROW3_H)},
    ]
    requested = [fs1, fs2, fs3]

    text_paths = []
    meta = {
        "left_sizes": [],
        "right_sizes": [],
        "left_texts": [],
        "right_texts": [],
        "right_w": right_w,
    }

    for i in range(3):
        label_txt, value_txt = rows[i]
        req_h = requested[i]

        left_final, left_base, _, left_scale, left_used_h = fit_text_for_box(
            label_txt, req_h, left_boxes[i]["w"], left_boxes[i]["row"]["h"], "bold"
        )
        right_final, right_base, _, right_scale, right_used_h = fit_text_for_box(
            value_txt, req_h, right_boxes[i]["w"], right_boxes[i]["row"]["h"], "regular"
        )

        left_placed = place_path_in_box(
            left_base,
            left_scale,
            left_boxes[i]["x"],
            left_boxes[i]["row"]["y"],
            left_boxes[i]["w"],
            left_boxes[i]["row"]["h"],
            align="left",
            pad_x=0.0,
        ) if left_base is not None else None

        right_placed = place_path_in_box(
            right_base,
            right_scale,
            right_boxes[i]["x"],
            right_boxes[i]["row"]["y"],
            right_boxes[i]["w"],
            right_boxes[i]["row"]["h"],
            align="left",
            pad_x=0.0,
        ) if right_base is not None else None

        text_paths.append(
            {
                "left_d": mpl_path_to_svg_d(left_placed),
                "right_d": mpl_path_to_svg_d(right_placed),
            }
        )

        meta["left_sizes"].append(left_used_h)
        meta["right_sizes"].append(right_used_h)
        meta["left_texts"].append(left_final)
        meta["right_texts"].append(right_final)

    hole_r = hole_dia / 2.0
    hole_y = LABEL_H / 2.0
    hole_left_x = hole_offset
    hole_right_x = LABEL_W - hole_offset

    if show_border:
        border_svg = f"""
        <rect x="{border_offset}" y="{border_offset}"
              width="{LABEL_W - 2 * border_offset}"
              height="{LABEL_H - 2 * border_offset}"
              rx="{corner_r}" ry="{corner_r}"
              fill="{colors['plate_fill']}"
              stroke="{colors['plate_stroke']}"
              stroke-width="0.20"/>
        """
    else:
        border_svg = f"""
        <rect x="{border_offset}" y="{border_offset}"
              width="{LABEL_W - 2 * border_offset}"
              height="{LABEL_H - 2 * border_offset}"
              rx="{corner_r}" ry="{corner_r}"
              fill="{colors['plate_fill']}"
              stroke="none"/>
        """

    guides_svg = ""
    if show_guides:
        guides_svg = f"""
        <g fill="none" stroke="{colors['guide_stroke']}" stroke-width="0.12" stroke-dasharray="0.8,0.8" opacity="0.75">
            <rect x="{left_x}" y="{row1_y}" width="{left_w}" height="{ROW1_H}" />
            <rect x="{left_x}" y="{row2_y}" width="{left_w}" height="{ROW2_H}" />
            <rect x="{left_x}" y="{row3_y}" width="{left_w}" height="{ROW3_H}" />

            <rect x="{right_x}" y="{row1_y}" width="{right_w}" height="{ROW1_H}" />
            <rect x="{right_x}" y="{row2_y}" width="{right_w}" height="{ROW2_H}" />
            <rect x="{right_x}" y="{row3_y}" width="{right_w}" height="{ROW3_H}" />
        </g>
        """

    svg = f"""<svg xmlns="http://www.w3.org/2000/svg"
        width="{LABEL_W}mm"
        height="{LABEL_H}mm"
        viewBox="0 0 {LABEL_W} {LABEL_H}">
        {border_svg}

        <circle cx="{hole_left_x}" cy="{hole_y}" r="{hole_r}"
                fill="{colors['hole_fill']}" stroke="{colors['plate_stroke']}" stroke-width="0.15"/>
        <circle cx="{hole_right_x}" cy="{hole_y}" r="{hole_r}"
                fill="{colors['hole_fill']}" stroke="{colors['plate_stroke']}" stroke-width="0.15"/>

        {guides_svg}

        <g fill="{colors['text_fill']}" fill-rule="evenodd" stroke="none">
            <path d="{text_paths[0]['left_d']}" />
            <path d="{text_paths[0]['right_d']}" />
            <path d="{text_paths[1]['left_d']}" />
            <path d="{text_paths[1]['right_d']}" />
            <path d="{text_paths[2]['left_d']}" />
            <path d="{text_paths[2]['right_d']}" />
        </g>
    </svg>"""

    return svg, meta


# ------------------------------------------------------------
# DXF generation
# ------------------------------------------------------------
def generate_dxf(
    owner,
    tool_number,
    part_desc,
    mode,
    hole_dia,
    hole_offset,
    corner_r,
    border_offset,
    left_x,
    left_w,
    right_x,
    row1_y,
    row2_y,
    row3_y,
    fs1,
    fs2,
    fs3,
    show_border=True,
):
    rows = build_rows(owner, tool_number, part_desc)
    right_w = LABEL_W - right_x - RIGHT_MARGIN

    left_boxes = [
        {"x": left_x, "w": left_w, "row": row_box(row1_y, ROW1_H)},
        {"x": left_x, "w": left_w, "row": row_box(row2_y, ROW2_H)},
        {"x": left_x, "w": left_w, "row": row_box(row3_y, ROW3_H)},
    ]
    right_boxes = [
        {"x": right_x, "w": right_w, "row": row_box(row1_y, ROW1_H)},
        {"x": right_x, "w": right_w, "row": row_box(row2_y, ROW2_H)},
        {"x": right_x, "w": right_w, "row": row_box(row3_y, ROW3_H)},
    ]

    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4

    for layer_name in ["BORDER", "HOLES", "TEXT", "FILL"]:
        if layer_name not in doc.layers:
            doc.layers.new(layer_name)

    msp = doc.modelspace()

    if mode == "Anodized aluminium (negative)":
        hatch = msp.add_hatch(dxfattribs={"layer": "FILL", "color": 7})
        hatch.paths.add_polyline_path(
            [
                (border_offset, border_offset),
                (LABEL_W - border_offset, border_offset),
                (LABEL_W - border_offset, LABEL_H - border_offset),
                (border_offset, LABEL_H - border_offset),
            ],
            is_closed=True,
        )

    if show_border:
        add_rounded_rect_dxf(
            msp,
            border_offset,
            border_offset,
            LABEL_W - 2 * border_offset,
            LABEL_H - 2 * border_offset,
            corner_r,
            "BORDER",
        )

    hole_r = hole_dia / 2.0
    hole_y = LABEL_H / 2.0
    hole_left_x = hole_offset
    hole_right_x = LABEL_W - hole_offset

    msp.add_circle((hole_left_x, hole_y), hole_r, dxfattribs={"layer": "HOLES"})
    msp.add_circle((hole_right_x, hole_y), hole_r, dxfattribs={"layer": "HOLES"})

    def add_text_entity(box_x, box_w, row_y, row_h, text_value, desired_h, weight):
        final_text, _, _, _, used_h = fit_text_for_box(
            text_value, desired_h, box_w, row_h, weight
        )
        if not final_text:
            return

        baseline_y = row_y + row_h * 0.80
        y_dxf = LABEL_H - baseline_y

        t = msp.add_text(
            final_text,
            dxfattribs={
                "height": used_h,
                "layer": "TEXT",
                "style": "Standard",
            },
        )
        t.set_placement((box_x, y_dxf), align=TextEntityAlignment.LEFT)

    add_text_entity(left_boxes[0]["x"], left_boxes[0]["w"], row1_y, ROW1_H, rows[0][0], fs1, "bold")
    add_text_entity(right_boxes[0]["x"], right_boxes[0]["w"], row1_y, ROW1_H, rows[0][1], fs1, "regular")

    add_text_entity(left_boxes[1]["x"], left_boxes[1]["w"], row2_y, ROW2_H, rows[1][0], fs2, "bold")
    add_text_entity(right_boxes[1]["x"], right_boxes[1]["w"], row2_y, ROW2_H, rows[1][1], fs2, "regular")

    add_text_entity(left_boxes[2]["x"], left_boxes[2]["w"], row3_y, ROW3_H, rows[2][0], fs3, "bold")
    add_text_entity(right_boxes[2]["x"], right_boxes[2]["w"], row3_y, ROW3_H, rows[2][1], fs3, "regular")

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


# ------------------------------------------------------------
# ZIP generation for batch
# ------------------------------------------------------------
def build_batch_zip(
    df: pd.DataFrame,
    default_mode: str,
    default_hole_dia: float,
    default_hole_offset: float,
    corner_r: float,
    border_offset: float,
    left_x: float,
    left_w: float,
    right_x: float,
    row1_y: float,
    row2_y: float,
    row3_y: float,
    fs1: float,
    fs2: float,
    fs3: float,
    show_border: bool,
):
    required = ["property_of", "tool_number", "part_description"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    base_names = []
    for _, row in df.iterrows():
        base = safe_filename(str(row.get("tool_number", "")).strip() or "laser_label")
        base_names.append(base)
    unique_names = make_unique_base_names(base_names)

    zip_buffer = io.BytesIO()
    preview_records = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (_, row) in enumerate(df.iterrows()):
            owner = "" if pd.isna(row["property_of"]) else str(row["property_of"])
            tool_number = "" if pd.isna(row["tool_number"]) else str(row["tool_number"])
            part_desc = "" if pd.isna(row["part_description"]) else str(row["part_description"])

            mode = parse_mode(row.get("mode", None), default_mode)
            hole_dia = parse_optional_float(
                row.get("hole_dia", row.get("hole_size", None)),
                default_hole_dia,
            )
            hole_offset = parse_optional_float(
                row.get("hole_offset", None),
                default_hole_offset,
            )

            svg_output, _ = generate_svg(
                owner=owner,
                tool_number=tool_number,
                part_desc=part_desc,
                mode=mode,
                hole_dia=hole_dia,
                hole_offset=hole_offset,
                corner_r=corner_r,
                border_offset=border_offset,
                left_x=left_x,
                left_w=left_w,
                right_x=right_x,
                row1_y=row1_y,
                row2_y=row2_y,
                row3_y=row3_y,
                fs1=fs1,
                fs2=fs2,
                fs3=fs3,
                show_guides=False,
                show_border=show_border,
            )

            dxf_bytes = generate_dxf(
                owner=owner,
                tool_number=tool_number,
                part_desc=part_desc,
                mode=mode,
                hole_dia=hole_dia,
                hole_offset=hole_offset,
                corner_r=corner_r,
                border_offset=border_offset,
                left_x=left_x,
                left_w=left_w,
                right_x=right_x,
                row1_y=row1_y,
                row2_y=row2_y,
                row3_y=row3_y,
                fs1=fs1,
                fs2=fs2,
                fs3=fs3,
                show_border=show_border,
            )

            base_name = unique_names[idx]
            zf.writestr(f"{base_name}.svg", svg_output.encode("utf-8"))
            zf.writestr(f"{base_name}.dxf", dxf_bytes)

            preview_records.append(
                {
                    "file_base": base_name,
                    "property_of": owner,
                    "tool_number": tool_number,
                    "part_description": part_desc,
                    "mode": mode,
                    "hole_dia": hole_dia,
                    "hole_offset": hole_offset,
                }
            )

    zip_buffer.seek(0)
    return zip_buffer.getvalue(), pd.DataFrame(preview_records)


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Laser Label Generator")
st.caption("Single and batch label generation with SVG vector outlines, DXF export, and ZIP packaging.")

tab_single, tab_batch = st.tabs(["Single label", "Batch CSV / Excel"])

# Shared settings
with st.sidebar:
    st.header("Shared geometry")

    preview_scale = st.slider("Preview scale (px/mm)", 4.0, 20.0, 8.5, 0.5)

    mode_default = st.selectbox(
        "Default mode",
        ALLOWED_MODES,
        index=1 if DEFAULT_MODE == "Anodized aluminium (negative)" else 0,
    )
    show_border = st.checkbox("Show border", value=True)
    show_guides = st.checkbox("Show guides in single preview", value=False)

    corner_r = st.slider("Corner radius (mm)", 0.0, 5.0, DEFAULT_CORNER_R, 0.1)
    border_offset = st.slider("Border offset (mm)", 0.0, 1.0, DEFAULT_BORDER_OFFSET, 0.05)

    st.subheader("Text heights")
    fs1 = st.slider("Row 1 text height (mm)", 1.2, 4.0, DEFAULT_FS1, 0.1)
    fs2 = st.slider("Row 2 text height (mm)", 1.2, 4.0, DEFAULT_FS2, 0.1)
    fs3 = st.slider("Row 3 text height (mm)", 1.0, 3.0, DEFAULT_FS3, 0.1)

    st.subheader("Default holes")
    hole_dia_default = st.slider("Hole diameter (mm)", 2.0, 5.0, DEFAULT_HOLE_DIA, 0.1)
    hole_offset_default = st.slider("Hole center from edge (mm)", 3.0, 8.0, DEFAULT_HOLE_OFFSET, 0.1)

    with st.expander("Fine tuning"):
        left_x = st.slider("Left column X (mm)", 6.0, 14.0, DEFAULT_LEFT_X, 0.1)
        left_w = st.slider("Left column width (mm)", 16.0, 28.0, DEFAULT_LEFT_W, 0.1)
        right_x = st.slider("Right column X (mm)", 26.0, 40.0, DEFAULT_RIGHT_X, 0.1)

        row1_y = st.slider("Row 1 top (mm)", 1.0, 6.0, DEFAULT_ROW1_Y, 0.1)
        row2_y = st.slider("Row 2 top (mm)", 5.0, 11.0, DEFAULT_ROW2_Y, 0.1)
        row3_y = st.slider("Row 3 top (mm)", 10.0, 17.0, DEFAULT_ROW3_Y, 0.1)

with tab_single:
    left_col, right_col = st.columns([1.0, 1.35], gap="large")

    with left_col:
        st.subheader("Text")
        owner = st.text_input("Property of", value=DEFAULT_OWNER, key="single_owner")
        tool_number = st.text_input("Tool number", value=DEFAULT_TOOL, key="single_tool")
        part_desc = st.text_input("Part description", value=DEFAULT_PART, key="single_part")

        mode_single = st.selectbox("Mode", ALLOWED_MODES, index=ALLOWED_MODES.index(mode_default), key="single_mode")
        hole_dia_single = st.number_input("Hole diameter override (mm)", min_value=2.0, max_value=5.0, value=float(hole_dia_default), step=0.1)
        hole_offset_single = st.number_input("Hole offset override (mm)", min_value=3.0, max_value=8.0, value=float(hole_offset_default), step=0.1)

    with right_col:
        st.subheader("Live preview")

        svg_output, meta = generate_svg(
            owner=owner,
            tool_number=tool_number,
            part_desc=part_desc,
            mode=mode_single,
            hole_dia=hole_dia_single,
            hole_offset=hole_offset_single,
            corner_r=corner_r,
            border_offset=border_offset,
            left_x=left_x,
            left_w=left_w,
            right_x=right_x,
            row1_y=row1_y,
            row2_y=row2_y,
            row3_y=row3_y,
            fs1=fs1,
            fs2=fs2,
            fs3=fs3,
            show_guides=show_guides,
            show_border=show_border,
        )

        colors = get_mode_colors(mode_single)
        preview_svg = make_preview_svg(svg_output, LABEL_W, LABEL_H, preview_scale)
        preview_height = int(LABEL_H * preview_scale + 80)

        preview_html = f"""
        <div style="
            background:{colors['preview_bg']};
            padding:20px;
            border-radius:12px;
            border:1px solid #d0d0d0;
            min-height:{preview_height}px;
            display:flex;
            justify-content:center;
            align-items:center;
            overflow:auto;">
            <div style="padding:10px; border-radius:10px;">
                {preview_svg}
            </div>
        </div>
        """
        components.html(preview_html, height=preview_height + 40)

        st.caption(
            f"Preview scale: {preview_scale:.1f} px/mm | "
            f"Right column width: {meta['right_w']:.1f} mm | "
            f"Used heights L: {meta['left_sizes'][0]:.2f}, {meta['left_sizes'][1]:.2f}, {meta['left_sizes'][2]:.2f} mm | "
            f"R: {meta['right_sizes'][0]:.2f}, {meta['right_sizes'][1]:.2f}, {meta['right_sizes'][2]:.2f} mm"
        )

        svg_bytes = svg_output.encode("utf-8")
        dxf_bytes = generate_dxf(
            owner=owner,
            tool_number=tool_number,
            part_desc=part_desc,
            mode=mode_single,
            hole_dia=hole_dia_single,
            hole_offset=hole_offset_single,
            corner_r=corner_r,
            border_offset=border_offset,
            left_x=left_x,
            left_w=left_w,
            right_x=right_x,
            row1_y=row1_y,
            row2_y=row2_y,
            row3_y=row3_y,
            fs1=fs1,
            fs2=fs2,
            fs3=fs3,
            show_border=show_border,
        )

        base_name = safe_filename(tool_number if tool_number.strip() else "laser_label")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Download SVG",
                data=svg_bytes,
                file_name=f"{base_name}.svg",
                mime="image/svg+xml",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Download DXF",
                data=dxf_bytes,
                file_name=f"{base_name}.dxf",
                mime="application/dxf",
                use_container_width=True,
            )

with tab_batch:
    st.subheader("Batch input")

    st.markdown(
        """
Expected columns:

- `property_of`
- `tool_number`
- `part_description`

Optional columns:

- `mode`
- `hole_dia` or `hole_size`
- `hole_offset`
"""
    )

    uploaded = st.file_uploader(
        "Upload CSV or Excel",
        type=["csv", "xlsx", "xls"],
        key="batch_upload",
    )

    if uploaded is not None:
        try:
            raw_df = load_tabular_file(uploaded)
            df = normalize_columns(raw_df)

            required_cols = ["property_of", "tool_number", "part_description"]
            missing_cols = [c for c in required_cols if c not in df.columns]

            if missing_cols:
                st.error(f"Missing required columns: {', '.join(missing_cols)}")
            else:
                preview_df = df.copy()
                if "mode" not in preview_df.columns:
                    preview_df["mode"] = mode_default
                else:
                    preview_df["mode"] = preview_df["mode"].apply(lambda x: parse_mode(x, mode_default))

                if "hole_dia" not in preview_df.columns:
                    if "hole_size" in preview_df.columns:
                        preview_df["hole_dia"] = preview_df["hole_size"]
                    else:
                        preview_df["hole_dia"] = hole_dia_default

                if "hole_offset" not in preview_df.columns:
                    preview_df["hole_offset"] = hole_offset_default

                preview_df["hole_dia"] = preview_df["hole_dia"].apply(lambda x: parse_optional_float(x, hole_dia_default))
                preview_df["hole_offset"] = preview_df["hole_offset"].apply(lambda x: parse_optional_float(x, hole_offset_default))

                st.markdown("### Preview list")
                st.dataframe(
                    preview_df[
                        ["property_of", "tool_number", "part_description", "mode", "hole_dia", "hole_offset"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

                row_labels = [
                    f"{i+1}. {str(preview_df.iloc[i]['tool_number'])}"
                    for i in range(len(preview_df))
                ]
                selected_label = st.selectbox("Preview row", row_labels)
                selected_idx = row_labels.index(selected_label)
                selected_row = preview_df.iloc[selected_idx]

                selected_mode = parse_mode(selected_row.get("mode", mode_default), mode_default)
                preview_svg_raw, _ = generate_svg(
                    owner=str(selected_row["property_of"]),
                    tool_number=str(selected_row["tool_number"]),
                    part_desc=str(selected_row["part_description"]),
                    mode=selected_mode,
                    hole_dia=parse_optional_float(selected_row.get("hole_dia", hole_dia_default), hole_dia_default),
                    hole_offset=parse_optional_float(selected_row.get("hole_offset", hole_offset_default), hole_offset_default),
                    corner_r=corner_r,
                    border_offset=border_offset,
                    left_x=left_x,
                    left_w=left_w,
                    right_x=right_x,
                    row1_y=row1_y,
                    row2_y=row2_y,
                    row3_y=row3_y,
                    fs1=fs1,
                    fs2=fs2,
                    fs3=fs3,
                    show_guides=False,
                    show_border=show_border,
                )

                colors = get_mode_colors(selected_mode)
                preview_svg = make_preview_svg(preview_svg_raw, LABEL_W, LABEL_H, preview_scale)
                preview_height = int(LABEL_H * preview_scale + 80)

                preview_html = f"""
                <div style="
                    background:{colors['preview_bg']};
                    padding:20px;
                    border-radius:12px;
                    border:1px solid #d0d0d0;
                    min-height:{preview_height}px;
                    display:flex;
                    justify-content:center;
                    align-items:center;
                    overflow:auto;">
                    <div style="padding:10px; border-radius:10px;">
                        {preview_svg}
                    </div>
                </div>
                """
                st.markdown("### Selected row preview")
                components.html(preview_html, height=preview_height + 40)

                zip_bytes, batch_result_df = build_batch_zip(
                    df=preview_df,
                    default_mode=mode_default,
                    default_hole_dia=float(hole_dia_default),
                    default_hole_offset=float(hole_offset_default),
                    corner_r=corner_r,
                    border_offset=border_offset,
                    left_x=left_x,
                    left_w=left_w,
                    right_x=right_x,
                    row1_y=row1_y,
                    row2_y=row2_y,
                    row3_y=row3_y,
                    fs1=fs1,
                    fs2=fs2,
                    fs3=fs3,
                    show_border=show_border,
                )

                st.markdown("### ZIP download")
                st.download_button(
                    "Download ZIP with SVG + DXF",
                    data=zip_bytes,
                    file_name="laser_labels_batch.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

                with st.expander("Generated file list"):
                    st.dataframe(batch_result_df, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Could not process file: {e}")

with st.expander("Example CSV"):
    st.code(
        """property_of,tool_number,part_description,mode,hole_dia,hole_offset
Stihl Group,89193,Steckzunge BA13-431-2100-A (id. 33193),Anodized aluminium (negative),3.2,4.3
DAFRA,TL-00125,Punch holder,Normal,3.0,4.5
Ledinek,TL-00126,Clamp plate,,,
""",
        language="csv",
    )

with st.expander("Notes"):
    st.write(
        """
Preview scale only affects on-screen display.

Exports remain unchanged:
- SVG stays at 78.5 × 21 mm
- DXF stays in real millimeter geometry

Batch mode behavior:
- one SVG and one DXF is created for every row
- files are packed into one ZIP
- file names are based on `tool_number`
- duplicate tool numbers automatically get `_2`, `_3`, etc.
"""
    )