import io
import re
import html
import zipfile
import unicodedata
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

st.markdown("""
<style>
section[data-testid="stSidebar"] { padding-top: 1rem; }
div[data-testid="stDownloadButton"] > button {
    font-weight: 600;
    letter-spacing: 0.02em;
}
div[data-testid="stTabs"] button[role="tab"] {
    font-weight: 600;
    font-size: 0.88rem;
    letter-spacing: 0.03em;
}
div[data-testid="stAlert"] { border-radius: 6px; }
hr { margin: 0.75rem 0; border-color: #e0e0e0; }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# Main geometry constants
# ------------------------------------------------------------
LABEL_W = 78.5
LABEL_H = 21.0

DEFAULT_OWNER         = "Stihl Group"
DEFAULT_TOOL          = "89193"
DEFAULT_PART          = "Steckzunge BA13-431-2100-A (id. 33193)"

DEFAULT_MODE          = "Anodized aluminium (negative)"
DEFAULT_HOLE_DIA      = 3.2
DEFAULT_HOLE_OFFSET   = 4.3
DEFAULT_CORNER_R      = 2.2
DEFAULT_BORDER_OFFSET = 0.25

DEFAULT_LEFT_X  = 8.4
DEFAULT_LEFT_W  = 21.5
DEFAULT_RIGHT_X = 34.2

DEFAULT_ROW1_Y = 3.0
DEFAULT_ROW2_Y = 8.0
DEFAULT_ROW3_Y = 13.2

ROW1_H = 4.2
ROW2_H = 4.2
ROW3_H = 3.8

DEFAULT_FS1 = 2.8
DEFAULT_FS2 = 2.8
DEFAULT_FS3 = 1.8

MIN_TEXT_HEIGHT_MM   = 1.2
RIGHT_MARGIN         = 2.2
VECTOR_FONT_FAMILY   = "DejaVu Sans Condensed"

PART_STACK_TRIGGER_LEN   = 52
PART_STACK_INDENT        = 2.1
PART_STACK_LABEL_H       = 2.0
PART_STACK_GAP           = 0.35
PART_STACK_BOTTOM_MARGIN = 0.7
MULTILINE_GAP_MM         = 0.35

PROPERTY_OF_MULTILINE_TRIGGER_LEN = 18

ALLOWED_MODES = ["Normal", "Anodized aluminium (negative)"]

CUSTOM_REQUIRED_ALIASES = {
    "part_description": ["izdelek", "part_description", "izdelek_2"],
    "tool_number":      ["orodje", "sifra_orodja", "orodje_sifra", "tool_number"],
    "property_of":      ["lastnik", "lastnik_orodja", "property_of"],
}

CUSTOM_OPTIONAL_DISPLAY = [
    "skladisce", "orodje", "racun", "napis_na", "sifra_orodja",
    "os_elrad", "slika", "lastnik", "kupec",
    "avtor_uros_stuklek_odjemalec",
    "parsed_property_of", "parsed_tool_number", "parsed_part_description",
]


# ============================================================
# Utility helpers
# ============================================================

def safe_filename(text: str) -> str:
    text = (text or "").strip()
    if not text:
        return "laser_label"
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    text = text.strip("._")
    return text[:80] or "laser_label"


def escape_xml(text: str) -> str:
    return html.escape(text or "", quote=False)


def slugify_text(text) -> str:
    text = "" if text is None or pd.isna(text) else str(text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("\n", " ").replace("/", " ")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def make_unique_strings(values):
    counts = {}
    result = []
    for v in values:
        base = v if v else "col"
        counts[base] = counts.get(base, 0) + 1
        if counts[base] == 1:
            result.append(base)
        else:
            result.append(f"{base}_{counts[base]}")
    return result


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def clean_parsed_value(text: str) -> str:
    text = normalize_space(text)
    text = re.sub(r"^\d+\)\s*", "", text)
    return text.strip(" -")


def unique_pairs(pairs):
    seen = set()
    out  = []
    for a, b in pairs:
        a = a.strip()
        b = b.strip()
        if not a or not b:
            continue
        key = (a, b)
        if key not in seen:
            seen.add(key)
            out.append((a, b))
    return out


def build_two_line_candidates(text: str):
    raw = "" if text is None or pd.isna(text) else str(text)
    raw = raw.replace("\r", "\n")
    raw = re.sub(r"\n+", "\n", raw).strip()

    pairs = []

    explicit_lines = [p.strip() for p in raw.split("\n") if p.strip()]
    if len(explicit_lines) >= 2:
        pairs.append((explicit_lines[0], " ".join(explicit_lines[1:])))

    flat = normalize_space(raw)
    if not flat:
        return []

    for m in re.finditer(r"\s+(?=(?:part\s*2\s*:))", flat, flags=re.IGNORECASE):
        pairs.append((flat[:m.start()].strip(), flat[m.start():].strip()))

    for m in re.finditer(r"\s+(?=(?:2[\.\)]|2[-\u2013]))", flat):
        pairs.append((flat[:m.start()].strip(), flat[m.start():].strip()))

    for pattern in [r"\s*[;|]\s*", r"\s+/\s+", r"\s+-\s+", r",\s+"]:
        for m in re.finditer(pattern, flat):
            pairs.append((flat[:m.start()].strip(), flat[m.end():].strip()))

    words = flat.split()
    if len(words) >= 2:
        best_i    = 1
        best_diff = 10**9
        for i in range(1, len(words)):
            left  = " ".join(words[:i])
            right = " ".join(words[i:])
            diff  = abs(len(left) - len(right))
            if diff < best_diff:
                best_diff = diff
                best_i    = i
        pairs.append((" ".join(words[:best_i]), " ".join(words[best_i:])))

    return unique_pairs(pairs)


def parse_napis_na(text: str) -> dict:
    raw = "" if text is None or pd.isna(text) else str(text)
    if not raw.strip():
        return {
            "parsed_property_of":     "",
            "parsed_tool_number":     "",
            "parsed_part_description": "",
        }

    s = raw.replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s).strip()
    flat = normalize_space(s)

    result = {
        "parsed_property_of":     "",
        "parsed_tool_number":     "",
        "parsed_part_description": "",
    }

    m = re.search(
        r"(?:property\s*of|lastnik|possessor|eigentümer)\s*:\s*(.+?)"
        r"(?=(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?"
        r"|part\s*1\s*:|part\s*description\s*:|izdelek\s*:|product\s*:|produkt\s*:|$))",
        flat, flags=re.IGNORECASE,
    )
    if m:
        result["parsed_property_of"] = clean_parsed_value(m.group(1))

    m = re.search(
        r"(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?)"
        r"\s*:\s*([A-Za-z0-9./_-]+)",
        flat, flags=re.IGNORECASE,
    )
    if m:
        result["parsed_tool_number"] = clean_parsed_value(m.group(1))

    m = re.search(r"((?:part\s*1)\s*:\s*.+)$", flat, flags=re.IGNORECASE)
    if m:
        desc = m.group(1).strip()
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", desc, count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()
        return result

    m = re.search(
        r"(?:(?:part\s*description)|(?:izdelek)|(?:product)|(?:produkt))\s*:\s*(.+)$",
        flat, flags=re.IGNORECASE,
    )
    if m:
        desc = clean_parsed_value(m.group(1))
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", desc, count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()
        return result

    m = re.search(
        r"(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?)"
        r"\s*:\s*[A-Za-z0-9./_-]+\s+(.+)$",
        flat, flags=re.IGNORECASE,
    )
    if m:
        desc = clean_parsed_value(m.group(1))
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", desc, count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()

    return result


def get_mode_colors(mode: str):
    if mode == "Anodized aluminium (negative)":
        return {
            "plate_fill":   "black",
            "text_fill":    "white",
            "plate_stroke": "black",
            "hole_fill":    "white",
            "guide_stroke": "#53b7ff",
            "preview_bg":   "#dcdcdc",
        }
    return {
        "plate_fill":   "white",
        "text_fill":    "black",
        "plate_stroke": "black",
        "hole_fill":    "white",
        "guide_stroke": "#1f77b4",
        "preview_bg":   "#f3f3f3",
    }


def build_rows(owner: str, tool_number: str, part_desc: str):
    return [
        ("Property of:",      owner),
        ("Tool number:",      tool_number),
        ("Part description:", part_desc),
    ]


def get_font_prop(weight: str):
    return FontProperties(family=VECTOR_FONT_FAMILY, weight=weight)


def make_base_text_path(text: str, weight: str):
    if not text:
        return None, None
    tp   = TextPath((0, 0), text, size=1, prop=get_font_prop(weight))
    bbox = tp.get_extents()
    if bbox.width <= 0 or bbox.height <= 0:
        return None, None
    return tp, bbox


def fit_text_for_box(text: str, desired_h_mm: float, box_w_mm: float, box_h_mm: float, weight: str):
    raw = (text or "").strip()
    if not raw:
        return "", None, None, 1.0, 0.0

    for cut in range(0, len(raw) + 1):
        candidate = raw if cut == 0 else raw[: max(0, len(raw) - cut)].rstrip() + "..."

        base_path, bbox = make_base_text_path(candidate, weight)
        if base_path is None or bbox is None:
            continue

        target_h = min(desired_h_mm, box_h_mm)
        scale    = target_h / bbox.height

        if bbox.width * scale <= box_w_mm:
            return candidate, base_path, bbox, scale, bbox.height * scale

        scale_w             = box_w_mm / bbox.width
        height_at_width_fit = bbox.height * scale_w

        if height_at_width_fit >= MIN_TEXT_HEIGHT_MM and height_at_width_fit <= box_h_mm:
            return candidate, base_path, bbox, scale_w, height_at_width_fit

    fallback  = "..."
    base_path, bbox = make_base_text_path(fallback, weight)
    if base_path is None or bbox is None:
        return "", None, None, 1.0, 0.0

    target_h = min(desired_h_mm, box_h_mm)
    scale    = min(target_h / bbox.height, box_w_mm / bbox.width)
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


def place_path_top_aligned(base_path, scale, box_x, top_y, box_w, align="left", pad_x=0.0):
    path = base_path.transformed(Affine2D().scale(scale, -scale))
    bbox = path.get_extents()
    if align == "center":
        tx = box_x + (box_w - bbox.width) / 2.0 - bbox.x0
    elif align == "right":
        tx = box_x + box_w - pad_x - bbox.x1
    else:
        tx = box_x + pad_x - bbox.x0
    ty = top_y - bbox.y0
    return path.transformed(Affine2D().translate(tx, ty))


def fit_text_block(
    text: str,
    desired_h_mm: float,
    box_x: float,
    box_y: float,
    box_w: float,
    box_h: float,
    weight: str,
    align: str = "left",
    pad_x: float = 0.0,
    max_lines: int = 1,
    first_line_anchor_h=None,
    force_single_if_fits: bool = False,
):
    raw_original = "" if text is None or pd.isna(text) else str(text).strip()
    raw_single   = normalize_space(raw_original.replace("\n", " "))

    empty = {
        "texts": [], "paths": [], "used_heights": [],
        "used_height_summary": 0.0, "line_boxes": [], "line_count": 0,
    }
    if not raw_original:
        return empty

    single_final, single_base, single_bbox, single_scale, single_used_h = fit_text_for_box(
        raw_single, desired_h_mm, box_w, box_h, weight
    )
    single_complete = single_final == raw_single

    best = {
        "texts":        [single_final],
        "bases":        [single_base],
        "scales":       [single_scale],
        "used_heights": [single_used_h],
        "line_box_h":   box_h,
        "gap":          0.0,
        "score": (1000 if single_complete else 0) + single_used_h * 100 + len(single_final),
    }

    if force_single_if_fits and single_complete and single_bbox is not None:
        width_at_desired_h = single_bbox.width * (desired_h_mm / single_bbox.height)
        single_truly_fits  = width_at_desired_h <= box_w
    else:
        single_truly_fits = False

    if (
        not (force_single_if_fits and single_truly_fits)
        and max_lines >= 2
        and box_h >= (MIN_TEXT_HEIGHT_MM * 2 + MULTILINE_GAP_MM)
    ):
        gap        = min(MULTILINE_GAP_MM, max(0.25, box_h * 0.08))
        line_box_h = (box_h - gap) / 2.0

        if line_box_h >= MIN_TEXT_HEIGHT_MM:
            for line1, line2 in build_two_line_candidates(raw_original):
                f1 = fit_text_for_box(line1, desired_h_mm, box_w, line_box_h, weight)
                f2 = fit_text_for_box(line2, desired_h_mm, box_w, line_box_h, weight)
                t1, b1, _, s1, h1 = f1
                t2, b2, _, s2, h2 = f2
                if not t1 or not t2:
                    continue
                complete = t1 == line1 and t2 == line2
                score    = (
                    (2000 if complete else 0)
                    + min(h1, h2) * 100
                    + len(t1) + len(t2)
                    - abs(len(line1) - len(line2)) * 0.15
                )
                if complete and not single_complete:
                    score += 500
                if score > best["score"]:
                    best = {
                        "texts":        [t1, t2],
                        "bases":        [b1, b2],
                        "scales":       [s1, s2],
                        "used_heights": [h1, h2],
                        "line_box_h":   line_box_h,
                        "gap":          gap,
                        "score":        score,
                    }

    paths      = []
    line_boxes = []

    if len(best["texts"]) == 1:
        placed = (
            place_path_in_box(
                best["bases"][0], best["scales"][0],
                box_x, box_y, box_w, box_h, align=align, pad_x=pad_x,
            )
            if best["bases"][0] is not None else None
        )
        if placed is not None:
            paths.append(placed)
        line_boxes.append((box_x, box_y, box_w, box_h))
        used_height_summary = best["used_heights"][0] if best["used_heights"] else 0.0
    else:
        h1  = best["used_heights"][0]
        h2  = best["used_heights"][1]
        gap = best["gap"]
        anchor_h        = first_line_anchor_h if first_line_anchor_h is not None else best["line_box_h"]
        anchor_h        = min(anchor_h, box_h)
        first_line_top  = box_y + max(0.0, (anchor_h - h1) / 2.0)
        second_line_top = first_line_top + h1 + gap

        placed1 = (
            place_path_top_aligned(
                best["bases"][0], best["scales"][0],
                box_x, first_line_top, box_w, align=align, pad_x=pad_x,
            )
            if best["bases"][0] is not None else None
        )
        placed2 = (
            place_path_top_aligned(
                best["bases"][1], best["scales"][1],
                box_x, second_line_top, box_w, align=align, pad_x=pad_x,
            )
            if best["bases"][1] is not None else None
        )
        if placed1 is not None:
            paths.append(placed1)
        if placed2 is not None:
            paths.append(placed2)
        line_boxes.append((box_x, first_line_top,  box_w, h1))
        line_boxes.append((box_x, second_line_top, box_w, h2))
        used_height_summary = (second_line_top - box_y) + h2

    return {
        "texts":               best["texts"],
        "paths":               paths,
        "used_heights":        best["used_heights"],
        "used_height_summary": used_height_summary,
        "line_boxes":          line_boxes,
        "line_count":          len(best["texts"]),
    }


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
    msp.add_line((x + r,     y),         (x + w - r, y),         dxfattribs={"layer": layer})
    msp.add_line((x + w,     y + r),     (x + w,     y + h - r), dxfattribs={"layer": layer})
    msp.add_line((x + w - r, y + h),     (x + r,     y + h),     dxfattribs={"layer": layer})
    msp.add_line((x,         y + h - r), (x,         y + r),     dxfattribs={"layer": layer})
    msp.add_arc((x + w - r, y + r),     r, 270, 360, dxfattribs={"layer": layer})
    msp.add_arc((x + w - r, y + h - r), r, 0,   90,  dxfattribs={"layer": layer})
    msp.add_arc((x + r,     y + h - r), r, 90,  180, dxfattribs={"layer": layer})
    msp.add_arc((x + r,     y + r),     r, 180, 270, dxfattribs={"layer": layer})


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [slugify_text(c) for c in df.columns]
    return df


def parse_mode(value, default_mode):
    v = str(value).strip() if pd.notna(value) else ""
    if not v:
        return default_mode
    lookup = {
        "normal":                        "Normal",
        "standard":                      "Normal",
        "default":                       "Normal",
        "negative":                      "Anodized aluminium (negative)",
        "anodized":                      "Anodized aluminium (negative)",
        "anodized_aluminium":            "Anodized aluminium (negative)",
        "anodized aluminium":            "Anodized aluminium (negative)",
        "anodized aluminium (negative)": "Anodized aluminium (negative)",
        "black":                         "Anodized aluminium (negative)",
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
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        try:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            if name.endswith(".xls"):
                return pd.read_excel(uploaded_file, engine="xlrd")
            return pd.read_excel(uploaded_file, engine="openpyxl")
    raise ValueError("Unsupported file type. Use CSV or XLSX/XLS.")


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
    svg_preview  = re.sub(r'width="[^"]+"',  f'width="{preview_w_px}px"',  svg_text, count=1)
    svg_preview  = re.sub(r'height="[^"]+"', f'height="{preview_h_px}px"', svg_preview, count=1)
    return svg_preview


def selection_editor(df: pd.DataFrame, key: str, display_columns: list, checkbox_label: str = "Print"):
    work_df = df.copy()
    if "print" not in work_df.columns:
        work_df.insert(0, "print", False)
    column_order = ["print"] + [c for c in display_columns if c in work_df.columns]
    return st.data_editor(
        work_df[column_order],
        key=key,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        disabled=[c for c in column_order if c != "print"],
        column_config={
            "print": st.column_config.CheckboxColumn(
                checkbox_label,
                help="Select rows for SVG + DXF export",
                default=False,
            )
        },
    )


def render_svg_preview_card(svg_output: str, mode: str, preview_scale: float):
    colors         = get_mode_colors(mode)
    preview_svg    = make_preview_svg(svg_output, LABEL_W, LABEL_H, preview_scale)
    preview_height = int(LABEL_H * preview_scale + 80)
    preview_html   = f"""
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


# ============================================================
# Shared layout computation — single source of truth for SVG + DXF
# ============================================================

def compute_label_layout(
    owner, tool_number, part_desc,
    left_x, left_w, right_x,
    row1_y, row2_y, row3_y,
    fs1, fs2, fs3,
    border_offset,
):
    use_stacked_part = should_use_stacked_part_layout(part_desc)
    rows             = build_rows(owner, tool_number, part_desc)
    right_w          = LABEL_W - right_x - RIGHT_MARGIN
    row0_value_box_h = row2_y - row1_y

    left_block_0 = fit_text_block(
        text=rows[0][0], desired_h_mm=fs1,
        box_x=left_x, box_y=row1_y, box_w=left_w, box_h=ROW1_H,
        weight="bold", max_lines=1,
    )
    right_block_0 = fit_text_block(
        text=rows[0][1], desired_h_mm=fs1,
        box_x=right_x, box_y=row1_y, box_w=right_w, box_h=row0_value_box_h,
        weight="regular", max_lines=2,
        first_line_anchor_h=ROW1_H, force_single_if_fits=True,
    )
    left_block_1 = fit_text_block(
        text=rows[1][0], desired_h_mm=fs2,
        box_x=left_x, box_y=row2_y, box_w=left_w, box_h=ROW2_H,
        weight="bold", max_lines=1,
    )
    right_block_1 = fit_text_block(
        text=rows[1][1], desired_h_mm=fs2,
        box_x=right_x, box_y=row2_y, box_w=right_w, box_h=ROW2_H,
        weight="regular", max_lines=1,
    )

    if use_stacked_part:
        stack      = get_part_stack_layout(left_x, border_offset, row3_y)
        part_lines = split_part_description_lines(part_desc)
        if len(part_lines) == 1:
            part_lines = [part_lines[0], ""]
        elif len(part_lines) == 0:
            part_lines = ["", ""]

        label_block = fit_text_block(
            text="Part description:", desired_h_mm=fs3,
            box_x=stack["label_x"], box_y=stack["label_y"],
            box_w=stack["label_w"], box_h=stack["label_h"],
            weight="bold", max_lines=1,
        )
        line1_block = fit_text_block(
            text=part_lines[0], desired_h_mm=fs3,
            box_x=stack["value_x"], box_y=stack["value_y"],
            box_w=stack["value_w"], box_h=stack["per_line_h"],
            weight="regular", max_lines=1,
        )
        line2_block = fit_text_block(
            text=part_lines[1], desired_h_mm=fs3,
            box_x=stack["value_x"],
            box_y=stack["value_y"] + stack["per_line_h"] + stack["line_gap"],
            box_w=stack["value_w"], box_h=stack["per_line_h"],
            weight="regular", max_lines=1,
        ) if part_lines[1] else {
            "paths": [], "texts": [], "used_height_summary": 0.0,
            "line_boxes": [], "used_heights": [],
        }

        row2_blocks = {
            "stacked":      True,
            "stack":        stack,
            "part_lines":   part_lines,
            "label_block":  label_block,
            "line1_block":  line1_block,
            "line2_block":  line2_block,
        }
    else:
        part_h    = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_STACK_BOTTOM_MARGIN))
        left_blk  = fit_text_block(
            text="Part description:", desired_h_mm=fs3,
            box_x=left_x, box_y=row3_y, box_w=left_w, box_h=ROW3_H,
            weight="bold", max_lines=1,
        )
        right_blk = fit_text_block(
            text=part_desc, desired_h_mm=fs3,
            box_x=right_x, box_y=row3_y, box_w=right_w, box_h=part_h,
            weight="regular", max_lines=2, first_line_anchor_h=ROW3_H,
        )
        row2_blocks = {
            "stacked":     False,
            "left_block":  left_blk,
            "right_block": right_blk,
        }

    return {
        "use_stacked_part": use_stacked_part,
        "right_w":          right_w,
        "rows":             rows,
        "left_block_0":     left_block_0,
        "right_block_0":    right_block_0,
        "left_block_1":     left_block_1,
        "right_block_1":    right_block_1,
        "row2_blocks":      row2_blocks,
        "row1_y":           row1_y,
        "row2_y":           row2_y,
        "row3_y":           row3_y,
    }


# ============================================================
# Stacked part helpers
# ============================================================

def split_part_description_lines(part_desc: str):
    raw = "" if part_desc is None or pd.isna(part_desc) else str(part_desc).strip()
    if not raw:
        return []
    raw   = raw.replace("\r", "\n")
    raw   = re.sub(r"\n+", "\n", raw).strip()
    lines = [p.strip() for p in raw.split("\n") if p.strip()]
    if len(lines) >= 2:
        return lines[:2]
    flat = normalize_space(raw)
    if re.search(r"part\s*1\s*:", flat, flags=re.IGNORECASE) and re.search(r"part\s*2\s*:", flat, flags=re.IGNORECASE):
        split = re.split(r"(?=part\s*2\s*:)", flat, maxsplit=1, flags=re.IGNORECASE)
        if len(split) == 2:
            return [split[0].strip(), split[1].strip()]
    candidates = build_two_line_candidates(raw)
    if candidates:
        return [candidates[0][0], candidates[0][1]]
    return [flat]


def should_use_stacked_part_layout(part_desc: str):
    raw  = "" if part_desc is None or pd.isna(part_desc) else str(part_desc)
    flat = normalize_space(raw.replace("\n", " "))
    if "\n" in raw:
        return True
    if re.search(r"part\s*2\s*:", raw, flags=re.IGNORECASE):
        return True
    if len(flat) >= PART_STACK_TRIGGER_LEN:
        return True
    return False


def get_part_stack_layout(left_x, border_offset, row3_y):
    label_x       = left_x
    label_y       = row3_y
    label_h       = PART_STACK_LABEL_H
    value_x       = left_x + PART_STACK_INDENT
    value_y       = row3_y + label_h + PART_STACK_GAP
    value_w       = LABEL_W - border_offset - RIGHT_MARGIN - value_x
    value_h_total = LABEL_H - PART_STACK_BOTTOM_MARGIN - value_y
    line_gap      = PART_STACK_GAP
    per_line_h    = max(MIN_TEXT_HEIGHT_MM, (value_h_total - line_gap) / 2.0)
    return {
        "label_x":       label_x,
        "label_y":       label_y,
        "label_w":       LABEL_W - label_x - RIGHT_MARGIN,
        "label_h":       label_h,
        "value_x":       value_x,
        "value_y":       value_y,
        "value_w":       value_w,
        "value_h_total": value_h_total,
        "line_gap":      line_gap,
        "per_line_h":    per_line_h,
    }


# ============================================================
# SVG generation
# ============================================================

def generate_svg(
    owner, tool_number, part_desc, mode,
    hole_dia, hole_offset, corner_r, border_offset,
    left_x, left_w, right_x,
    row1_y, row2_y, row3_y,
    fs1, fs2, fs3,
    show_guides=False,
    show_border=True,
    show_holes=True,
):
    colors           = get_mode_colors(mode)
    layout           = compute_label_layout(
        owner, tool_number, part_desc,
        left_x, left_w, right_x,
        row1_y, row2_y, row3_y,
        fs1, fs2, fs3, border_offset,
    )
    right_w          = layout["right_w"]
    use_stacked_part = layout["use_stacked_part"]
    row2_blocks      = layout["row2_blocks"]

    all_svg_paths = []
    meta = {
        "left_sizes":   [], "right_sizes":  [],
        "left_texts":   [], "right_texts":  [],
        "right_w":      right_w,
        "stacked_part": use_stacked_part,
    }

    def collect_block(block, side):
        for p in block["paths"]:
            d = mpl_path_to_svg_d(p)
            if d:
                all_svg_paths.append(d)
        if side == "left":
            meta["left_sizes"].append(block["used_height_summary"])
            meta["left_texts"].append(" / ".join(block["texts"]))
        else:
            meta["right_sizes"].append(block["used_height_summary"])
            meta["right_texts"].append(" / ".join(block["texts"]))

    collect_block(layout["left_block_0"],  "left")
    collect_block(layout["right_block_0"], "right")
    collect_block(layout["left_block_1"],  "left")
    collect_block(layout["right_block_1"], "right")

    if row2_blocks["stacked"]:
        for blk in [row2_blocks["label_block"], row2_blocks["line1_block"], row2_blocks["line2_block"]]:
            for p in blk["paths"]:
                d = mpl_path_to_svg_d(p)
                if d:
                    all_svg_paths.append(d)
        meta["left_sizes"].append(row2_blocks["label_block"]["used_height_summary"])
        meta["right_sizes"].append(max(
            row2_blocks["line1_block"].get("used_height_summary", 0),
            row2_blocks["line2_block"].get("used_height_summary", 0),
        ))
        meta["left_texts"].append("Part description:")
        meta["right_texts"].append(" / ".join([x for x in row2_blocks["part_lines"] if x]))
    else:
        for blk in [row2_blocks["left_block"], row2_blocks["right_block"]]:
            for p in blk["paths"]:
                d = mpl_path_to_svg_d(p)
                if d:
                    all_svg_paths.append(d)
        meta["left_sizes"].append(row2_blocks["left_block"]["used_height_summary"])
        meta["right_sizes"].append(row2_blocks["right_block"]["used_height_summary"])
        meta["left_texts"].append("Part description:")
        meta["right_texts"].append(" / ".join(row2_blocks["right_block"]["texts"]))

    # ---- border ----
    border_svg = f"""
    <rect x="{border_offset}" y="{border_offset}"
          width="{LABEL_W - 2 * border_offset}"
          height="{LABEL_H - 2 * border_offset}"
          rx="{corner_r}" ry="{corner_r}"
          fill="{colors['plate_fill']}"
          stroke="{'none' if not show_border else colors['plate_stroke']}"
          stroke-width="0.20"/>
    """

    # ---- holes ----
    hole_r       = hole_dia / 2.0
    hole_y_pos   = LABEL_H / 2.0
    hole_left_x  = hole_offset
    hole_right_x = LABEL_W - hole_offset

    if show_holes:
        # Full solid circles — will be engraved / cut
        holes_svg = f"""
        <circle cx="{hole_left_x}"  cy="{hole_y_pos}" r="{hole_r}"
                fill="{colors['hole_fill']}" stroke="{colors['plate_stroke']}" stroke-width="0.15"/>
        <circle cx="{hole_right_x}" cy="{hole_y_pos}" r="{hole_r}"
                fill="{colors['hole_fill']}" stroke="{colors['plate_stroke']}" stroke-width="0.15"/>
        """
    else:
        # Dashed guide rings — preview only, not engraved
        holes_svg = f"""
        <circle cx="{hole_left_x}"  cy="{hole_y_pos}" r="{hole_r}"
                fill="none" stroke="{colors['guide_stroke']}"
                stroke-width="0.12" stroke-dasharray="0.6,0.6" opacity="0.55"/>
        <circle cx="{hole_right_x}" cy="{hole_y_pos}" r="{hole_r}"
                fill="none" stroke="{colors['guide_stroke']}"
                stroke-width="0.12" stroke-dasharray="0.6,0.6" opacity="0.55"/>
        """

    # ---- guides ----
    guides_svg = ""
    if show_guides:
        if use_stacked_part:
            stack = row2_blocks["stack"]
            guides_svg = f"""
            <g fill="none" stroke="{colors['guide_stroke']}" stroke-width="0.12" stroke-dasharray="0.8,0.8" opacity="0.75">
                <rect x="{left_x}"  y="{row1_y}" width="{left_w}"  height="{ROW1_H}" />
                <rect x="{right_x}" y="{row1_y}" width="{right_w}" height="{row2_y - row1_y}" />
                <rect x="{left_x}"  y="{row2_y}" width="{left_w}"  height="{ROW2_H}" />
                <rect x="{right_x}" y="{row2_y}" width="{right_w}" height="{ROW2_H}" />
                <rect x="{stack['label_x']}" y="{stack['label_y']}" width="{stack['label_w']}" height="{stack['label_h']}" />
                <rect x="{stack['value_x']}" y="{stack['value_y']}" width="{stack['value_w']}" height="{stack['per_line_h']}" />
                <rect x="{stack['value_x']}" y="{stack['value_y'] + stack['per_line_h'] + stack['line_gap']}" width="{stack['value_w']}" height="{stack['per_line_h']}" />
            </g>"""
        else:
            part_h = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_STACK_BOTTOM_MARGIN))
            guides_svg = f"""
            <g fill="none" stroke="{colors['guide_stroke']}" stroke-width="0.12" stroke-dasharray="0.8,0.8" opacity="0.75">
                <rect x="{left_x}"  y="{row1_y}" width="{left_w}"  height="{ROW1_H}" />
                <rect x="{right_x}" y="{row1_y}" width="{right_w}" height="{row2_y - row1_y}" />
                <rect x="{left_x}"  y="{row2_y}" width="{left_w}"  height="{ROW2_H}" />
                <rect x="{right_x}" y="{row2_y}" width="{right_w}" height="{ROW2_H}" />
                <rect x="{left_x}"  y="{row3_y}" width="{left_w}"  height="{ROW3_H}" />
                <rect x="{right_x}" y="{row3_y}" width="{right_w}" height="{part_h}" />
            </g>"""

    text_svg = "\n".join(f'<path d="{d}" />' for d in all_svg_paths)

    svg = f"""<svg xmlns="http://www.w3.org/2000/svg"
        width="{LABEL_W}mm"
        height="{LABEL_H}mm"
        viewBox="0 0 {LABEL_W} {LABEL_H}">
        {border_svg}
        {holes_svg}
        {guides_svg}
        <g fill="{colors['text_fill']}" fill-rule="evenodd" stroke="none">
            {text_svg}
        </g>
    </svg>"""

    return svg, meta


# ============================================================
# DXF generation
# ============================================================

def generate_dxf(
    owner, tool_number, part_desc, mode,
    hole_dia, hole_offset, corner_r, border_offset,
    left_x, left_w, right_x,
    row1_y, row2_y, row3_y,
    fs1, fs2, fs3,
    show_border=True,
    show_holes=True,
):
    layout      = compute_label_layout(
        owner, tool_number, part_desc,
        left_x, left_w, right_x,
        row1_y, row2_y, row3_y,
        fs1, fs2, fs3, border_offset,
    )
    row2_blocks = layout["row2_blocks"]

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
                (border_offset,           border_offset),
                (LABEL_W - border_offset, border_offset),
                (LABEL_W - border_offset, LABEL_H - border_offset),
                (border_offset,           LABEL_H - border_offset),
            ],
            is_closed=True,
        )

    if show_border:
        add_rounded_rect_dxf(
            msp, border_offset, border_offset,
            LABEL_W - 2 * border_offset,
            LABEL_H - 2 * border_offset,
            corner_r, "BORDER",
        )

    # Holes: only add DXF entities when show_holes=True.
    # The HOLES layer always exists so CAD software doesn't complain,
    # but it will be empty when holes are excluded.
    if show_holes:
        hole_r       = hole_dia / 2.0
        hole_y_pos   = LABEL_H / 2.0
        hole_left_x  = hole_offset
        hole_right_x = LABEL_W - hole_offset
        msp.add_circle((hole_left_x,  hole_y_pos), hole_r, dxfattribs={"layer": "HOLES"})
        msp.add_circle((hole_right_x, hole_y_pos), hole_r, dxfattribs={"layer": "HOLES"})

    def emit_block_dxf(block):
        for i, final_text in enumerate(block["texts"]):
            if not final_text:
                continue
            line_box_x, line_box_y, _, line_box_h = block["line_boxes"][i]
            used_h     = block["used_heights"][i]
            baseline_y = line_box_y + line_box_h * 0.80
            y_dxf      = LABEL_H - baseline_y
            t = msp.add_text(
                final_text,
                dxfattribs={"height": used_h, "layer": "TEXT", "style": "Standard"},
            )
            t.set_placement((line_box_x, y_dxf), align=TextEntityAlignment.LEFT)

    emit_block_dxf(layout["left_block_0"])
    emit_block_dxf(layout["right_block_0"])
    emit_block_dxf(layout["left_block_1"])
    emit_block_dxf(layout["right_block_1"])

    if row2_blocks["stacked"]:
        emit_block_dxf(row2_blocks["label_block"])
        emit_block_dxf(row2_blocks["line1_block"])
        if row2_blocks["line2_block"].get("texts"):
            emit_block_dxf(row2_blocks["line2_block"])
    else:
        emit_block_dxf(row2_blocks["left_block"])
        emit_block_dxf(row2_blocks["right_block"])

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


# ============================================================
# ZIP batch generation
# ============================================================

def build_batch_zip(
    df: pd.DataFrame,
    default_mode: str,
    default_hole_dia: float,
    default_hole_offset: float,
    corner_r: float,
    border_offset: float,
    left_x: float, left_w: float, right_x: float,
    row1_y: float, row2_y: float, row3_y: float,
    fs1: float, fs2: float, fs3: float,
    show_border: bool,
    show_holes: bool,
    include_svg: bool = True,
    include_dxf: bool = True,
    progress_bar=None,
):
    required = ["property_of", "tool_number", "part_description"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    base_names   = [safe_filename(str(r.get("tool_number", "")).strip() or "laser_label") for _, r in df.iterrows()]
    unique_names = make_unique_base_names(base_names)

    zip_buffer      = io.BytesIO()
    preview_records = []
    total           = len(df)

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (_, row) in enumerate(df.iterrows()):
            owner       = "" if pd.isna(row["property_of"])     else str(row["property_of"])
            tool_number = "" if pd.isna(row["tool_number"])      else str(row["tool_number"])
            part_desc   = "" if pd.isna(row["part_description"]) else str(row["part_description"])
            mode        = parse_mode(row.get("mode", None), default_mode)
            hole_dia    = parse_optional_float(row.get("hole_dia", row.get("hole_size", None)), default_hole_dia)
            hole_offset = parse_optional_float(row.get("hole_offset", None), default_hole_offset)
            base_name   = unique_names[idx]

            kwargs = dict(
                owner=owner, tool_number=tool_number, part_desc=part_desc,
                mode=mode, hole_dia=hole_dia, hole_offset=hole_offset,
                corner_r=corner_r, border_offset=border_offset,
                left_x=left_x, left_w=left_w, right_x=right_x,
                row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                fs1=fs1, fs2=fs2, fs3=fs3,
                show_holes=show_holes,
            )

            if include_svg:
                svg_output, _ = generate_svg(**kwargs, show_guides=False, show_border=show_border)
                zf.writestr(f"{base_name}.svg", svg_output.encode("utf-8"))

            if include_dxf:
                dxf_bytes = generate_dxf(**kwargs, show_border=show_border)
                zf.writestr(f"{base_name}.dxf", dxf_bytes)

            preview_records.append({
                "file_base":        base_name,
                "property_of":      owner,
                "tool_number":      tool_number,
                "part_description": part_desc,
                "mode":             mode,
                "hole_dia":         hole_dia,
                "hole_offset":      hole_offset,
                "holes_exported":   show_holes,
            })

            if progress_bar is not None:
                progress_bar.progress((idx + 1) / total, text=f"Generating {idx + 1} / {total}")

    zip_buffer.seek(0)
    return zip_buffer.getvalue(), pd.DataFrame(preview_records)


# ============================================================
# Custom Excel helpers
# ============================================================

def _score_custom_header_block(raw_df: pd.DataFrame, start_row: int) -> int:
    parts = []
    for r in range(start_row, min(start_row + 3, len(raw_df))):
        row_vals = [str(v) for v in raw_df.iloc[r].tolist() if pd.notna(v)]
        parts.extend(row_vals)
    text  = slugify_text(" ".join(parts))
    score = 0
    if "izdelek" in text: score += 3
    if "orodje"  in text and "sifra" in text: score += 4
    if "lastnik" in text: score += 3
    if "seznam_orodij" in text: score -= 2
    return score


def find_custom_excel_sheet_and_header(uploaded_file):
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)
    best_sheet = None
    best_row   = None
    best_score = -(10**9)
    for sheet in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if raw_df.empty:
            continue
        max_scan = min(20, len(raw_df))
        for start_row in range(max_scan):
            score = _score_custom_header_block(raw_df, start_row)
            if score > best_score:
                best_score = score
                best_sheet = sheet
                best_row   = start_row
    if best_sheet is None or best_row is None or best_score < 6:
        raise ValueError(
            "Could not detect the custom Excel header row. "
            "Ensure the file has columns: lastnik / orodje / sifra_orodja / izdelek."
        )
    return best_sheet, best_row


def combine_multirow_headers(raw_df: pd.DataFrame, header_row: int, depth: int = 3) -> list:
    headers = []
    for col_idx in range(raw_df.shape[1]):
        parts = []
        for r in range(header_row, min(header_row + depth, len(raw_df))):
            value = raw_df.iloc[r, col_idx]
            if pd.notna(value):
                text = str(value).strip()
                if text and text.lower() != "nan":
                    parts.append(text)
        headers.append(slugify_text(" ".join(parts)))
    return make_unique_strings(headers)


def pick_first_existing_column(columns, aliases):
    for alias in aliases:
        if alias in columns:
            return alias
    return None


def normalize_custom_template_df(data_df: pd.DataFrame) -> pd.DataFrame:
    df = data_df.copy()

    if "napis_na" in df.columns:
        parsed_df = df["napis_na"].apply(parse_napis_na).apply(pd.Series)
        for col in parsed_df.columns:
            df[col] = parsed_df[col]
    else:
        df["parsed_property_of"]     = ""
        df["parsed_tool_number"]     = ""
        df["parsed_part_description"] = ""

    tool_source = pick_first_existing_column(df.columns, ["orodje", "sifra_orodja", "orodje_sifra", "tool_number"])
    part_source = pick_first_existing_column(df.columns, ["izdelek", "part_description", "izdelek_2"])
    prop_source = pick_first_existing_column(df.columns, ["lastnik", "lastnik_orodja", "property_of"])

    fallback_property = df[prop_source].fillna("").astype(str).str.strip() if prop_source else pd.Series([""] * len(df), index=df.index)
    fallback_tool     = df[tool_source].fillna("").astype(str).str.strip() if tool_source else pd.Series([""] * len(df), index=df.index)
    fallback_part     = df[part_source].fillna("").astype(str).str.strip() if part_source else pd.Series([""] * len(df), index=df.index)

    df["property_of"]    = df["parsed_property_of"].fillna("").astype(str).str.strip()
    df["property_of"]    = df["property_of"].mask(df["property_of"] == "", fallback_property)
    df["tool_number"]    = df["parsed_tool_number"].fillna("").astype(str).str.strip()
    df["tool_number"]    = df["tool_number"].mask(df["tool_number"] == "", fallback_tool)
    df["part_description"] = df["parsed_part_description"].fillna("").astype(str).str.strip()
    df["part_description"] = df["part_description"].mask(df["part_description"] == "", fallback_part)

    df = df[
        ~(
            (df["tool_number"]      == "")
            & (df["property_of"]    == "")
            & (df["part_description"] == "")
        )
    ].copy()

    df["mode"]        = DEFAULT_MODE
    df["hole_dia"]    = DEFAULT_HOLE_DIA
    df["hole_offset"] = DEFAULT_HOLE_OFFSET
    return df


def load_custom_tool_excel(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    sheet_name, header_row = find_custom_excel_sheet_and_header(uploaded_file)
    uploaded_file.seek(0)
    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    headers    = combine_multirow_headers(raw_df, header_row, depth=3)
    data_start = header_row + 3
    data_df    = raw_df.iloc[data_start:].copy()
    data_df.columns = headers
    data_df    = data_df.dropna(how="all").reset_index(drop=True)
    data_df.insert(0, "excel_row", range(data_start + 1, data_start + 1 + len(data_df)))
    data_df    = normalize_custom_template_df(data_df)
    data_df    = data_df.iloc[::-1].reset_index(drop=True)
    data_df.insert(0, "priority", range(1, len(data_df) + 1))
    return data_df


# ============================================================
# Sidebar
# ============================================================

with st.sidebar:
    st.header("⚙️ Label settings")

    preview_scale = st.slider("Preview scale (px/mm)", 4.0, 20.0, 8.5, 0.5)

    mode_default = st.selectbox(
        "Default engraving mode",
        ALLOWED_MODES,
        index=1 if DEFAULT_MODE == "Anodized aluminium (negative)" else 0,
        help="Anodized aluminium (negative) = white text on black plate. Normal = black text on white plate.",
    )

    st.divider()
    st.markdown("**Export options**")

    show_border = st.checkbox(
        "Include border",
        value=True,
        help="Outer rounded rectangle. Uncheck to export text + holes only.",
    )

    show_holes = st.checkbox(
        "Include mounting holes",
        value=True,
        help=(
            "When checked: hole circles are written into SVG and DXF exports and will be engraved/cut.\n\n"
            "When unchecked: holes are OMITTED from all export files. "
            "The preview shows faint dashed rings so you can still see their positions."
        ),
    )

    if not show_holes:
        st.caption("⚠️ Holes omitted from exports — dashed guide shown in preview only")

    show_guides = st.checkbox(
        "Show layout guides (preview only)",
        value=False,
        help="Draws dashed bounding boxes for each text region. Never written to export files.",
    )

    with st.expander("🔡 Text sizes"):
        fs1 = st.slider("Row 1 text height (mm)", 1.2, 4.0, DEFAULT_FS1, 0.1, help="'Property of' row")
        fs2 = st.slider("Row 2 text height (mm)", 1.2, 4.0, DEFAULT_FS2, 0.1, help="'Tool number' row")
        fs3 = st.slider("Row 3 text height (mm)", 1.0, 3.0, DEFAULT_FS3, 0.1, help="'Part description' row")

    with st.expander("⭕ Hole dimensions"):
        hole_dia_default    = st.slider("Hole diameter (mm)",          2.0, 5.0, DEFAULT_HOLE_DIA,    0.1)
        hole_offset_default = st.slider("Hole centre from edge (mm)",  3.0, 8.0, DEFAULT_HOLE_OFFSET, 0.1)

    with st.expander("🔲 Shape"):
        corner_r      = st.slider("Corner radius (mm)",  0.0, 5.0, DEFAULT_CORNER_R,      0.1)
        border_offset = st.slider("Border inset (mm)",   0.0, 1.0, DEFAULT_BORDER_OFFSET,  0.05)

    with st.expander("📐 Advanced column layout"):
        st.caption("Only change these if your physical label template differs.")
        left_x  = st.slider("Left column X (mm)",     6.0,  14.0, DEFAULT_LEFT_X,  0.1)
        left_w  = st.slider("Left column width (mm)", 16.0, 28.0, DEFAULT_LEFT_W,  0.1)
        right_x = st.slider("Right column X (mm)",    26.0, 40.0, DEFAULT_RIGHT_X, 0.1)
        row1_y  = st.slider("Row 1 top (mm)",          1.0,  6.0, DEFAULT_ROW1_Y,  0.1)
        row2_y  = st.slider("Row 2 top (mm)",          5.0, 11.0, DEFAULT_ROW2_Y,  0.1)
        row3_y  = st.slider("Row 3 top (mm)",         10.0, 17.0, DEFAULT_ROW3_Y,  0.1)


# ============================================================
# Tabs
# ============================================================

tab_single, tab_batch, tab_custom = st.tabs([
    "🏷️ Design one label",
    "📋 Generate from spreadsheet",
    "🗂️ Import tool inventory",
])


# ------------------------------------------------------------
# Tab 1 — Single label
# ------------------------------------------------------------
with tab_single:
    left_col, right_col = st.columns([1.0, 1.35], gap="large")

    with left_col:
        st.subheader("Label text")
        owner       = st.text_input("Property of", value=DEFAULT_OWNER, key="single_owner")
        tool_number = st.text_input("Tool number",  value=DEFAULT_TOOL,  key="single_tool")

        part_desc = st.text_area(
            "Part description",
            value=DEFAULT_PART,
            key="single_part",
            height=100,
            help=f"Stacked two-line layout activates automatically at {PART_STACK_TRIGGER_LEN} chars or on newline.",
        )
        char_count  = len(part_desc)
        pct         = min(char_count / PART_STACK_TRIGGER_LEN, 1.0)
        trigger_msg = (
            "⚡ Stacked layout active"
            if char_count >= PART_STACK_TRIGGER_LEN
            else f"{PART_STACK_TRIGGER_LEN - char_count} chars until stacked layout"
        )
        st.caption(f"{char_count} chars — {trigger_msg}")
        st.progress(pct)

        st.divider()
        mode_single        = st.selectbox(
            "Engraving mode", ALLOWED_MODES,
            index=ALLOWED_MODES.index(mode_default), key="single_mode",
        )
        hole_dia_single    = st.number_input(
            "Hole diameter (mm)",         min_value=2.0, max_value=5.0,
            value=float(hole_dia_default), step=0.1,
        )
        hole_offset_single = st.number_input(
            "Hole centre from edge (mm)",  min_value=3.0, max_value=8.0,
            value=float(hole_offset_default), step=0.1,
        )

    with right_col:
        st.subheader("Live preview")

        # Holes status badge
        if show_holes:
            st.caption("⭕ Holes: **included** in export")
        else:
            st.caption("🚫 Holes: **excluded** from export — dashed guide shown only")

        svg_output, meta = generate_svg(
            owner=owner, tool_number=tool_number, part_desc=part_desc,
            mode=mode_single,
            hole_dia=hole_dia_single, hole_offset=hole_offset_single,
            corner_r=corner_r, border_offset=border_offset,
            left_x=left_x, left_w=left_w, right_x=right_x,
            row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
            fs1=fs1, fs2=fs2, fs3=fs3,
            show_guides=show_guides,
            show_border=show_border,
            show_holes=show_holes,
        )

        render_svg_preview_card(svg_output, mode_single, preview_scale)

        st.caption(
            f"Layout: {'**stacked**' if meta.get('stacked_part') else 'inline'} | "
            f"Right column: {meta['right_w']:.1f} mm | "
            f"Scale: {preview_scale:.1f} px/mm"
        )

        dxf_bytes = generate_dxf(
            owner=owner, tool_number=tool_number, part_desc=part_desc,
            mode=mode_single,
            hole_dia=hole_dia_single, hole_offset=hole_offset_single,
            corner_r=corner_r, border_offset=border_offset,
            left_x=left_x, left_w=left_w, right_x=right_x,
            row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
            fs1=fs1, fs2=fs2, fs3=fs3,
            show_border=show_border,
            show_holes=show_holes,
        )
        base_name = safe_filename(tool_number.strip() or "laser_label")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Download SVG  (laser / Inkscape)",
                data=svg_output.encode("utf-8"),
                file_name=f"{base_name}.svg",
                mime="image/svg+xml",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "⬇️ Download DXF  (CAD / import)",
                data=dxf_bytes,
                file_name=f"{base_name}.dxf",
                mime="application/dxf",
                use_container_width=True,
            )


# ------------------------------------------------------------
# Tab 2 — Batch from spreadsheet
# ------------------------------------------------------------
with tab_batch:
    st.subheader("Batch label generation from spreadsheet")

    with st.expander("ℹ️ Expected column names", expanded=False):
        st.markdown("""
**Required columns** (exact names, case-insensitive):

| Column | Description |
|---|---|
| `property_of` | Owner / company name |
| `tool_number` | Mould / tool ID |
| `part_description` | Part name and ID |

**Optional columns:**

| Column | Description |
|---|---|
| `mode` | `normal` or `negative` (default: sidebar) |
| `hole_dia` or `hole_size` | Hole diameter in mm |
| `hole_offset` | Hole centre distance from edge in mm |
""")

    uploaded = st.file_uploader(
        "Upload CSV or Excel file",
        type=["csv", "xlsx", "xls"],
        key="batch_upload",
    )

    if uploaded is not None:
        try:
            raw_df = load_tabular_file(uploaded)
            df     = normalize_columns(raw_df)

            required_cols = ["property_of", "tool_number", "part_description"]
            missing_cols  = [c for c in required_cols if c not in df.columns]

            if missing_cols:
                st.error(
                    f"**Missing required columns:** {', '.join(missing_cols)}\n\n"
                    f"Columns found: `{'`, `'.join(df.columns.tolist())}`"
                )
                st.stop()

            found_optional = [c for c in ["mode", "hole_dia", "hole_size", "hole_offset"] if c in df.columns]
            st.success(
                f"✅ Loaded **{len(df)} rows** from `{uploaded.name}` — required columns detected ✓"
                + (f" | Optional: `{'`, `'.join(found_optional)}`" if found_optional else "")
            )

            preview_df = df.copy()
            if "mode" not in preview_df.columns:
                preview_df["mode"] = mode_default
            else:
                preview_df["mode"] = preview_df["mode"].apply(lambda x: parse_mode(x, mode_default))
            if "hole_dia" not in preview_df.columns:
                preview_df["hole_dia"] = preview_df["hole_size"] if "hole_size" in preview_df.columns else hole_dia_default
            if "hole_offset" not in preview_df.columns:
                preview_df["hole_offset"] = hole_offset_default
            preview_df["hole_dia"]    = preview_df["hole_dia"].apply(lambda x: parse_optional_float(x, hole_dia_default))
            preview_df["hole_offset"] = preview_df["hole_offset"].apply(lambda x: parse_optional_float(x, hole_offset_default))

            st.markdown("### 🔍 Browse rows")
            row_labels     = [f"{i + 1}. {str(preview_df.iloc[i]['tool_number'])}" for i in range(len(preview_df))]
            selected_label = st.selectbox("Preview a row before generating", row_labels)
            selected_idx   = row_labels.index(selected_label)
            selected_row   = preview_df.iloc[selected_idx]
            selected_mode  = parse_mode(selected_row.get("mode", mode_default), mode_default)

            if not show_holes:
                st.caption("🚫 Holes excluded from export — dashed guides shown in preview only")

            preview_svg_raw, _ = generate_svg(
                owner=str(selected_row["property_of"]),
                tool_number=str(selected_row["tool_number"]),
                part_desc=str(selected_row["part_description"]),
                mode=selected_mode,
                hole_dia=parse_optional_float(selected_row.get("hole_dia", hole_dia_default), hole_dia_default),
                hole_offset=parse_optional_float(selected_row.get("hole_offset", hole_offset_default), hole_offset_default),
                corner_r=corner_r, border_offset=border_offset,
                left_x=left_x, left_w=left_w, right_x=right_x,
                row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                fs1=fs1, fs2=fs2, fs3=fs3,
                show_guides=False, show_border=show_border,
                show_holes=show_holes,
            )
            render_svg_preview_card(preview_svg_raw, selected_mode, preview_scale)

            st.divider()

            with st.expander("📄 All rows", expanded=False):
                st.dataframe(
                    preview_df[["property_of", "tool_number", "part_description", "mode", "hole_dia", "hole_offset"]],
                    use_container_width=True, hide_index=True,
                )

            st.markdown("### 📦 Export all rows")
            include_svg_batch = st.checkbox("Include SVG files", value=True, key="batch_svg")
            include_dxf_batch = st.checkbox("Include DXF files", value=True, key="batch_dxf")

            if show_holes:
                st.info("⭕ Mounting holes will be **included** in exported files (set in sidebar).")
            else:
                st.warning("🚫 Mounting holes will be **excluded** from exported files (set in sidebar).")

            if st.button("🚀 Generate ZIP", type="primary", use_container_width=True, key="batch_generate"):
                progress = st.progress(0, text="Starting…")
                status   = st.empty()
                try:
                    zip_bytes, batch_result_df = build_batch_zip(
                        df=preview_df,
                        default_mode=mode_default,
                        default_hole_dia=float(hole_dia_default),
                        default_hole_offset=float(hole_offset_default),
                        corner_r=corner_r, border_offset=border_offset,
                        left_x=left_x, left_w=left_w, right_x=right_x,
                        row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                        fs1=fs1, fs2=fs2, fs3=fs3,
                        show_border=show_border,
                        show_holes=show_holes,
                        include_svg=include_svg_batch,
                        include_dxf=include_dxf_batch,
                        progress_bar=progress,
                    )
                    progress.empty()
                    status.success(f"✅ Done — {len(batch_result_df)} labels generated.")
                    st.download_button(
                        f"⬇️ Download ZIP  ({len(batch_result_df)} labels)",
                        data=zip_bytes,
                        file_name="laser_labels_batch.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                    with st.expander("Generated file list"):
                        st.dataframe(batch_result_df, use_container_width=True, hide_index=True)
                except Exception as e:
                    progress.empty()
                    st.error(f"Generation failed: {e}")

        except Exception as e:
            st.error(f"Could not read file: {e}")


# ------------------------------------------------------------
# Tab 3 — Custom Excel tool inventory
# ------------------------------------------------------------
with tab_custom:
    st.subheader("Import tool inventory (custom Excel format)")
    st.caption(
        "Parses your internal tool list Excel. Primary source: **napis_na** column "
        "(parsed for Property of / Tool number / Part description). "
        "Fallback: lastnik / orodje or sifra_orodja / izdelek."
    )

    custom_uploaded = st.file_uploader(
        "Upload Excel or CSV tool list",
        type=["csv", "xlsx", "xls"],
        key="custom_upload",
    )

    if custom_uploaded is not None:
        try:
            file_name = custom_uploaded.name.lower()

            if file_name.endswith(".csv"):
                custom_uploaded.seek(0)
                raw_custom_df = normalize_columns(pd.read_csv(custom_uploaded))
                raw_custom_df.insert(0, "excel_row", range(2, 2 + len(raw_custom_df)))
                custom_df = normalize_custom_template_df(raw_custom_df)
                custom_df = custom_df.iloc[::-1].reset_index(drop=True)
                custom_df.insert(0, "priority", range(1, len(custom_df) + 1))
            else:
                custom_df = load_custom_tool_excel(custom_uploaded)

            custom_df["mode"]        = mode_default
            custom_df["hole_dia"]    = hole_dia_default
            custom_df["hole_offset"] = hole_offset_default

            st.success(
                f"✅ Loaded **{len(custom_df)} tools** from `{custom_uploaded.name}` — "
                f"sorted by priority (1 = highest)."
            )

            st.markdown("### 🔍 Browse all tools")
            browse_labels = [
                f"P{int(r['priority'])} | {r['tool_number']} | row {int(r['excel_row'])}"
                for _, r in custom_df.iterrows()
            ]
            browse_choice = st.selectbox("Select a tool to preview", browse_labels, key="custom_browse_choice")
            browse_idx    = browse_labels.index(browse_choice)
            browse_row    = custom_df.iloc[browse_idx]

            if not show_holes:
                st.caption("🚫 Holes excluded from export — dashed guides shown in preview only")

            browse_svg, _ = generate_svg(
                owner=str(browse_row["property_of"]),
                tool_number=str(browse_row["tool_number"]),
                part_desc=str(browse_row["part_description"]),
                mode=mode_default,
                hole_dia=float(hole_dia_default), hole_offset=float(hole_offset_default),
                corner_r=corner_r, border_offset=border_offset,
                left_x=left_x, left_w=left_w, right_x=right_x,
                row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                fs1=fs1, fs2=fs2, fs3=fs3,
                show_guides=False, show_border=show_border,
                show_holes=show_holes,
            )
            render_svg_preview_card(browse_svg, mode_default, preview_scale)

            browse_dxf = generate_dxf(
                owner=str(browse_row["property_of"]),
                tool_number=str(browse_row["tool_number"]),
                part_desc=str(browse_row["part_description"]),
                mode=mode_default,
                hole_dia=float(hole_dia_default), hole_offset=float(hole_offset_default),
                corner_r=corner_r, border_offset=border_offset,
                left_x=left_x, left_w=left_w, right_x=right_x,
                row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                fs1=fs1, fs2=fs2, fs3=fs3,
                show_border=show_border,
                show_holes=show_holes,
            )
            browse_base = safe_filename(str(browse_row["tool_number"]) or "laser_label")
            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    "⬇️ Download SVG  (this tool)",
                    data=browse_svg.encode("utf-8"),
                    file_name=f"{browse_base}.svg",
                    mime="image/svg+xml",
                    use_container_width=True,
                )
            with b2:
                st.download_button(
                    "⬇️ Download DXF  (this tool)",
                    data=browse_dxf,
                    file_name=f"{browse_base}.dxf",
                    mime="application/dxf",
                    use_container_width=True,
                )

            st.divider()

            with st.expander("🐛 Parsing debug — inspect what was extracted per row"):
                debug_cols = ["priority", "excel_row", "tool_number", "property_of", "part_description"]
                if "napis_na" in custom_df.columns:
                    debug_cols += ["napis_na", "parsed_property_of", "parsed_tool_number", "parsed_part_description"]
                available_debug = [c for c in debug_cols if c in custom_df.columns]
                st.dataframe(custom_df[available_debug], use_container_width=True, hide_index=True)
                st.caption(
                    "If a field is blank, the parser couldn't find the expected keywords. "
                    "Check that napis_na follows: `Property of: X  Tool number: Y  Part 1: Z`"
                )

            st.markdown("### ✅ Select rows for batch export")
            st.caption("Tick the rows you want to include in the ZIP download.")

            display_cols = ["priority", "excel_row", "tool_number", "property_of", "part_description"]
            for extra_col in CUSTOM_OPTIONAL_DISPLAY:
                if extra_col in custom_df.columns and extra_col not in display_cols:
                    display_cols.append(extra_col)

            edited_custom_df = selection_editor(
                custom_df,
                key="custom_selection_editor",
                display_columns=display_cols,
                checkbox_label="Export",
            )

            selected_rows = edited_custom_df[edited_custom_df["print"]].copy()

            info_c1, info_c2 = st.columns([1, 1])
            with info_c1:
                st.info(f"Total rows: **{len(edited_custom_df)}** | Selected: **{len(selected_rows)}**")
            with info_c2:
                if not edited_custom_df.empty:
                    st.caption(f"Top priority: {str(edited_custom_df.iloc[0]['tool_number'])}")

            st.markdown("### 📦 Export selected rows")
            if selected_rows.empty:
                st.warning("No rows selected. Use the checkboxes above to mark rows for export.")
            else:
                with st.expander(f"Rows queued ({len(selected_rows)})", expanded=False):
                    st.dataframe(
                        selected_rows.drop(columns=["print"]),
                        use_container_width=True, hide_index=True,
                    )

                if show_holes:
                    st.info("⭕ Mounting holes will be **included** in exported files (set in sidebar).")
                else:
                    st.warning("🚫 Mounting holes will be **excluded** from exported files (set in sidebar).")

                if st.button(
                    f"🚀 Generate ZIP for {len(selected_rows)} selected rows",
                    type="primary", use_container_width=True, key="custom_generate",
                ):
                    export_df = selected_rows.drop(columns=["print"]).copy()
                    progress  = st.progress(0, text="Starting…")
                    status    = st.empty()
                    try:
                        zip_bytes, export_manifest_df = build_batch_zip(
                            df=export_df,
                            default_mode=mode_default,
                            default_hole_dia=float(hole_dia_default),
                            default_hole_offset=float(hole_offset_default),
                            corner_r=corner_r, border_offset=border_offset,
                            left_x=left_x, left_w=left_w, right_x=right_x,
                            row1_y=row1_y, row2_y=row2_y, row3_y=row3_y,
                            fs1=fs1, fs2=fs2, fs3=fs3,
                            show_border=show_border,
                            show_holes=show_holes,
                            progress_bar=progress,
                        )
                        progress.empty()
                        status.success(f"✅ Done — {len(export_manifest_df)} labels generated.")
                        st.download_button(
                            f"⬇️ Download ZIP  ({len(export_manifest_df)} labels)",
                            data=zip_bytes,
                            file_name="laser_labels_selected.zip",
                            mime="application/zip",
                            use_container_width=True,
                        )
                        with st.expander("Exported file list"):
                            st.dataframe(export_manifest_df, use_container_width=True, hide_index=True)
                    except Exception as e:
                        progress.empty()
                        st.error(f"Generation failed: {e}")

        except Exception as e:
            st.error(f"Could not process file: {e}")


# ============================================================
# Footer expanders
# ============================================================

with st.expander("📋 Column mapping reference (Custom Excel / CSV)"):
    st.code(
        """Primary parsing source
───────────────────────────────────────
napis_na  →  parse:
    Property of:      → property_of
    Tool number:      → tool_number
    Part 1: / Part 2: → part_description

Example input:
  Property of: Husqvarna AB  Tool number: 89202
  Part 1: Charging Plate P23 left (id. 33202)
  Part 2: Charging Plate P23 right (id. 33203)

Fallback columns (used if napis_na is absent or unparseable)
───────────────────────────────────────
lastnik / lastnik_orodja    → property_of
orodje  / sifra_orodja      → tool_number
izdelek / izdelek_2         → part_description
""",
        language="text",
    )

with st.expander("📖 Layout & export notes"):
    st.write("""
**Property of** wraps to 2 lines automatically for long company names.

**Tool number** is always single-line.

**Part description** switches to stacked layout when text exceeds 52 characters,
contains a newline, or contains "Part 2:".

**Mounting holes** are controlled by "Include mounting holes" in the sidebar:
- ✅ Checked → solid hole circles written into SVG and DXF, will be engraved/cut
- ☐ Unchecked → hole entities completely omitted from all export files;
  preview shows faint dashed rings so you can still verify their positions

The **HOLES layer** in DXF always exists (empty when excluded) so no CAD
software will complain about a missing layer reference.
""")