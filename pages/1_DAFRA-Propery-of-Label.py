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

AUTO_WRAP_MAX_LINES = 2
PART_VALUE_BOTTOM_MARGIN = 1.0
MULTILINE_GAP_MM = 0.35

ALLOWED_MODES = ["Normal", "Anodized aluminium (negative)"]

CUSTOM_REQUIRED_ALIASES = {
    "part_description": ["izdelek", "part_description", "izdelek_2"],
    "tool_number": ["orodje", "sifra_orodja", "orodje_sifra", "tool_number"],
    "property_of": ["lastnik", "lastnik_orodja", "property_of"],
}

CUSTOM_OPTIONAL_DISPLAY = [
    "skladisce",
    "orodje",
    "racun",
    "napis_na",
    "sifra_orodja",
    "os_elrad",
    "slika",
    "lastnik",
    "kupec",
    "avtor_uros_stuklek_odjemalec",
    "parsed_property_of",
    "parsed_tool_number",
    "parsed_part_description",
]

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
    out = []
    for a, b in pairs:
        a = normalize_space(a)
        b = normalize_space(b)
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

    explicit_lines = [normalize_space(p) for p in raw.split("\n") if normalize_space(p)]
    if len(explicit_lines) >= 2:
        pairs.append((explicit_lines[0], " ".join(explicit_lines[1:])))

    flat = normalize_space(raw)
    if not flat:
        return []

    for m in re.finditer(r"\s+(?=(?:2[\.\)]|2[-–]))", flat):
        pairs.append((flat[:m.start()].strip(), flat[m.start():].strip()))

    for pattern in [r"\s*[;|]\s*", r"\s+/\s+", r"\s+-\s+", r",\s+"]:
        for m in re.finditer(pattern, flat):
            pairs.append((flat[:m.start()].strip(), flat[m.end():].strip()))

    words = flat.split()
    if len(words) >= 2:
        best_i = 1
        best_diff = 10**9
        for i in range(1, len(words)):
            left = " ".join(words[:i])
            right = " ".join(words[i:])
            diff = abs(len(left) - len(right))
            if diff < best_diff:
                best_diff = diff
                best_i = i
        pairs.append((" ".join(words[:best_i]), " ".join(words[best_i:])))

    return unique_pairs(pairs)


def parse_napis_na(text: str) -> dict:
    raw = "" if text is None or pd.isna(text) else str(text)
    if not raw.strip():
        return {
            "parsed_property_of": "",
            "parsed_tool_number": "",
            "parsed_part_description": "",
        }

    s = raw.replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s).strip()
    flat = normalize_space(s)

    patterns = {
        "parsed_property_of": [
            r"(?:lastnik|possessor|eigentümer)\s*:\s*(.+?)(?=(?:orodje\s*št\.?|orodje\s*st\.?|tool\s*no\.?|wkz\.\s*no\.?|(?:\d+\)\s*)?(?:izdelek|product|produkt)\s*:|$))",
        ],
        "parsed_tool_number": [
            r"(?:orodje\s*št\.?|orodje\s*st\.?|tool\s*no\.?|wkz\.\s*no\.?)\s*:\s*([A-Za-z0-9./_-]+)",
        ],
        "parsed_part_description": [
            r"(?:\d+\)\s*)?(?:izdelek|product|produkt)\s*:\s*(.+)$",
        ],
    }

    result = {
        "parsed_property_of": "",
        "parsed_tool_number": "",
        "parsed_part_description": "",
    }

    for key, regex_list in patterns.items():
        for pattern in regex_list:
            m = re.search(pattern, flat, flags=re.IGNORECASE)
            if m:
                result[key] = clean_parsed_value(m.group(1))
                break

    return result


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
):
    raw = normalize_space(text)
    empty = {
        "texts": [],
        "paths": [],
        "used_heights": [],
        "used_height_summary": 0.0,
        "line_boxes": [],
        "line_count": 0,
    }
    if not raw:
        return empty

    single_final, single_base, _, single_scale, single_used_h = fit_text_for_box(
        raw, desired_h_mm, box_w, box_h, weight
    )
    single_complete = single_final == raw

    best = {
        "texts": [single_final],
        "bases": [single_base],
        "scales": [single_scale],
        "used_heights": [single_used_h],
        "line_box_h": box_h,
        "gap": 0.0,
        "score": (1000 if single_complete else 0) + single_used_h * 100 + len(single_final),
    }

    if max_lines >= 2 and box_h >= (MIN_TEXT_HEIGHT_MM * 2 + MULTILINE_GAP_MM):
        gap = min(MULTILINE_GAP_MM, max(0.25, box_h * 0.08))
        line_box_h = (box_h - gap) / 2.0

        if line_box_h >= MIN_TEXT_HEIGHT_MM:
            for line1, line2 in build_two_line_candidates(text):
                f1 = fit_text_for_box(line1, desired_h_mm, box_w, line_box_h, weight)
                f2 = fit_text_for_box(line2, desired_h_mm, box_w, line_box_h, weight)

                t1, b1, _, s1, h1 = f1
                t2, b2, _, s2, h2 = f2

                if not t1 or not t2:
                    continue

                complete = (t1 == line1 and t2 == line2)
                score = (
                    (2000 if complete else 0)
                    + min(h1, h2) * 100
                    + len(t1) + len(t2)
                    - abs(len(line1) - len(line2)) * 0.15
                )

                if complete and not single_complete:
                    score += 500

                if score > best["score"]:
                    best = {
                        "texts": [t1, t2],
                        "bases": [b1, b2],
                        "scales": [s1, s2],
                        "used_heights": [h1, h2],
                        "line_box_h": line_box_h,
                        "gap": gap,
                        "score": score,
                    }

    paths = []
    line_boxes = []

    if len(best["texts"]) == 1:
        placed = place_path_in_box(
            best["bases"][0],
            best["scales"][0],
            box_x,
            box_y,
            box_w,
            box_h,
            align=align,
            pad_x=pad_x,
        ) if best["bases"][0] is not None else None

        if placed is not None:
            paths.append(placed)

        line_boxes.append((box_x, box_y, box_w, box_h))
        used_height_summary = best["used_heights"][0] if best["used_heights"] else 0.0

    else:
        top_y = box_y
        bottom_y = box_y + best["line_box_h"] + best["gap"]

        for i, line_y in enumerate([top_y, bottom_y]):
            placed = place_path_in_box(
                best["bases"][i],
                best["scales"][i],
                box_x,
                line_y,
                box_w,
                best["line_box_h"],
                align=align,
                pad_x=pad_x,
            ) if best["bases"][i] is not None else None

            if placed is not None:
                paths.append(placed)

            line_boxes.append((box_x, line_y, box_w, best["line_box_h"]))

        used_height_summary = sum(best["used_heights"]) + best["gap"]

    return {
        "texts": best["texts"],
        "paths": paths,
        "used_heights": best["used_heights"],
        "used_height_summary": used_height_summary,
        "line_boxes": line_boxes,
        "line_count": len(best["texts"]),
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
    df.columns = [slugify_text(c) for c in df.columns]
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

    svg_preview = re.sub(r'width="[^"]+"', f'width="{preview_w_px}px"', svg_text, count=1)
    svg_preview = re.sub(r'height="[^"]+"', f'height="{preview_h_px}px"', svg_text, count=1)

    return svg_preview


def selection_editor(
    df: pd.DataFrame,
    key: str,
    display_columns: list[str],
    checkbox_label: str = "Print",
):
    work_df = df.copy()
    if "print" not in work_df.columns:
        work_df.insert(0, "print", False)

    column_order = ["print"] + [c for c in display_columns if c in work_df.columns]
    edited_df = st.data_editor(
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
    return edited_df


def render_svg_preview_card(svg_output: str, mode: str, preview_scale: float):
    colors = get_mode_colors(mode)
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

# ------------------------------------------------------------
# Custom Excel helpers
# ------------------------------------------------------------
def _score_custom_header_block(raw_df: pd.DataFrame, start_row: int) -> int:
    parts = []
    for r in range(start_row, min(start_row + 3, len(raw_df))):
        row_vals = [str(v) for v in raw_df.iloc[r].tolist() if pd.notna(v)]
        parts.extend(row_vals)

    text = slugify_text(" ".join(parts))

    score = 0
    if "izdelek" in text:
        score += 3
    if "orodje" in text and "sifra" in text:
        score += 4
    if "lastnik" in text:
        score += 3
    if "seznam_orodij" in text:
        score -= 2
    return score


def find_custom_excel_sheet_and_header(uploaded_file):
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)

    best_sheet = None
    best_row = None
    best_score = -10**9

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
                best_row = start_row

    if best_sheet is None or best_row is None or best_score < 6:
        raise ValueError("Could not detect the custom Excel header row.")
    return best_sheet, best_row


def combine_multirow_headers(raw_df: pd.DataFrame, header_row: int, depth: int = 3) -> list[str]:
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
        df["parsed_property_of"] = ""
        df["parsed_tool_number"] = ""
        df["parsed_part_description"] = ""

    resolved = {}
    for target_col, aliases in CUSTOM_REQUIRED_ALIASES.items():
        source_col = pick_first_existing_column(df.columns, aliases)
        resolved[target_col] = source_col

    if resolved.get("property_of") is not None:
        fallback_property = df[resolved["property_of"]].fillna("").astype(str).str.strip()
    else:
        fallback_property = pd.Series([""] * len(df), index=df.index)

    tool_source = pick_first_existing_column(df.columns, ["orodje", "sifra_orodja", "orodje_sifra", "tool_number"])
    if tool_source is not None:
        fallback_tool = df[tool_source].fillna("").astype(str).str.strip()
    else:
        fallback_tool = pd.Series([""] * len(df), index=df.index)

    part_source = pick_first_existing_column(df.columns, ["izdelek", "part_description", "izdelek_2"])
    if part_source is not None:
        fallback_part = df[part_source].fillna("").astype(str).str.strip()
    else:
        fallback_part = pd.Series([""] * len(df), index=df.index)

    df["property_of"] = df["parsed_property_of"].fillna("").astype(str).str.strip()
    df["property_of"] = df["property_of"].mask(df["property_of"] == "", fallback_property)

    df["tool_number"] = df["parsed_tool_number"].fillna("").astype(str).str.strip()
    df["tool_number"] = df["tool_number"].mask(df["tool_number"] == "", fallback_tool)

    df["part_description"] = df["parsed_part_description"].fillna("").astype(str).str.strip()
    df["part_description"] = df["part_description"].mask(df["part_description"] == "", fallback_part)

    df = df[
        ~(
            (df["tool_number"] == "")
            & (df["property_of"] == "")
            & (df["part_description"] == "")
        )
    ].copy()

    df["mode"] = DEFAULT_MODE
    df["hole_dia"] = DEFAULT_HOLE_DIA
    df["hole_offset"] = DEFAULT_HOLE_OFFSET

    return df


def load_custom_tool_excel(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    sheet_name, header_row = find_custom_excel_sheet_and_header(uploaded_file)

    uploaded_file.seek(0)
    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    headers = combine_multirow_headers(raw_df, header_row, depth=3)
    data_start = header_row + 3

    data_df = raw_df.iloc[data_start:].copy()
    data_df.columns = headers
    data_df = data_df.dropna(how="all").reset_index(drop=True)
    data_df.insert(0, "excel_row", range(data_start + 1, data_start + 1 + len(data_df)))

    data_df = normalize_custom_template_df(data_df)

    # lower Excel rows = higher priority
    data_df = data_df.iloc[::-1].reset_index(drop=True)
    data_df.insert(0, "priority", range(1, len(data_df) + 1))

    return data_df

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
    part_value_h = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_VALUE_BOTTOM_MARGIN))

    left_boxes = [
        {"x": left_x, "w": left_w, "row": row_box(row1_y, ROW1_H)},
        {"x": left_x, "w": left_w, "row": row_box(row2_y, ROW2_H)},
        {"x": left_x, "w": left_w, "row": row_box(row3_y, ROW3_H)},
    ]
    right_boxes = [
        {"x": right_x, "w": right_w, "row": row_box(row1_y, ROW1_H)},
        {"x": right_x, "w": right_w, "row": row_box(row2_y, ROW2_H)},
        {"x": right_x, "w": right_w, "row": row_box(row3_y, part_value_h)},
    ]
    requested = [fs1, fs2, fs3]

    all_svg_paths = []
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

        left_block = fit_text_block(
            text=label_txt,
            desired_h_mm=req_h,
            box_x=left_boxes[i]["x"],
            box_y=left_boxes[i]["row"]["y"],
            box_w=left_boxes[i]["w"],
            box_h=left_boxes[i]["row"]["h"],
            weight="bold",
            align="left",
            pad_x=0.0,
            max_lines=1,
        )

        right_block = fit_text_block(
            text=value_txt,
            desired_h_mm=req_h,
            box_x=right_boxes[i]["x"],
            box_y=right_boxes[i]["row"]["y"],
            box_w=right_boxes[i]["w"],
            box_h=right_boxes[i]["row"]["h"],
            weight="regular",
            align="left",
            pad_x=0.0,
            max_lines=AUTO_WRAP_MAX_LINES,
        )

        for p in left_block["paths"]:
            d = mpl_path_to_svg_d(p)
            if d:
                all_svg_paths.append(d)

        for p in right_block["paths"]:
            d = mpl_path_to_svg_d(p)
            if d:
                all_svg_paths.append(d)

        meta["left_sizes"].append(left_block["used_height_summary"])
        meta["right_sizes"].append(right_block["used_height_summary"])
        meta["left_texts"].append(" / ".join(left_block["texts"]))
        meta["right_texts"].append(" / ".join(right_block["texts"]))

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
            <rect x="{right_x}" y="{row3_y}" width="{right_w}" height="{part_value_h}" />
        </g>
        """

    text_svg = "\n".join(f'<path d="{d}" />' for d in all_svg_paths)

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
            {text_svg}
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
    part_value_h = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_VALUE_BOTTOM_MARGIN))

    left_boxes = [
        {"x": left_x, "w": left_w, "row": row_box(row1_y, ROW1_H)},
        {"x": left_x, "w": left_w, "row": row_box(row2_y, ROW2_H)},
        {"x": left_x, "w": left_w, "row": row_box(row3_y, ROW3_H)},
    ]
    right_boxes = [
        {"x": right_x, "w": right_w, "row": row_box(row1_y, ROW1_H)},
        {"x": right_x, "w": right_w, "row": row_box(row2_y, ROW2_H)},
        {"x": right_x, "w": right_w, "row": row_box(row3_y, part_value_h)},
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

    def add_text_block_dxf(box_x, box_w, row_y, row_h, text_value, desired_h, weight, max_lines=1):
        block = fit_text_block(
            text=text_value,
            desired_h_mm=desired_h,
            box_x=box_x,
            box_y=row_y,
            box_w=box_w,
            box_h=row_h,
            weight=weight,
            align="left",
            pad_x=0.0,
            max_lines=max_lines,
        )

        for i, final_text in enumerate(block["texts"]):
            line_box_x, line_box_y, _, line_box_h = block["line_boxes"][i]
            used_h = block["used_heights"][i]

            baseline_y = line_box_y + line_box_h * 0.80
            y_dxf = LABEL_H - baseline_y

            t = msp.add_text(
                final_text,
                dxfattribs={
                    "height": used_h,
                    "layer": "TEXT",
                    "style": "Standard",
                },
            )
            t.set_placement((line_box_x, y_dxf), align=TextEntityAlignment.LEFT)

    add_text_block_dxf(left_boxes[0]["x"], left_boxes[0]["w"], row1_y, ROW1_H, rows[0][0], fs1, "bold", max_lines=1)
    add_text_block_dxf(right_boxes[0]["x"], right_boxes[0]["w"], row1_y, ROW1_H, rows[0][1], fs1, "regular", max_lines=AUTO_WRAP_MAX_LINES)

    add_text_block_dxf(left_boxes[1]["x"], left_boxes[1]["w"], row2_y, ROW2_H, rows[1][0], fs2, "bold", max_lines=1)
    add_text_block_dxf(right_boxes[1]["x"], right_boxes[1]["w"], row2_y, ROW2_H, rows[1][1], fs2, "regular", max_lines=AUTO_WRAP_MAX_LINES)

    add_text_block_dxf(left_boxes[2]["x"], left_boxes[2]["w"], row3_y, ROW3_H, rows[2][0], fs3, "bold", max_lines=1)
    add_text_block_dxf(right_boxes[2]["x"], right_boxes[2]["w"], row3_y, part_value_h, rows[2][1], fs3, "regular", max_lines=AUTO_WRAP_MAX_LINES)

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")

# ------------------------------------------------------------
# ZIP generation
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
    include_svg: bool = True,
    include_dxf: bool = True,
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

            base_name = unique_names[idx]

            if include_svg:
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
                zf.writestr(f"{base_name}.svg", svg_output.encode("utf-8"))

            if include_dxf:
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
st.caption(
    "Single and batch label generation with SVG vector outlines, DXF export, ZIP packaging, "
    "and a dedicated custom Excel import tab."
)

tab_single, tab_batch, tab_custom = st.tabs(
    ["Single label", "Batch CSV / Excel", "Custom Excel tools list"]
)

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

        render_svg_preview_card(svg_output, mode_single, preview_scale)

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
                    f"{i + 1}. {str(preview_df.iloc[i]['tool_number'])}"
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

                st.markdown("### Selected row preview")
                render_svg_preview_card(preview_svg_raw, selected_mode, preview_scale)

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

with tab_custom:
    st.subheader("Custom Excel tools list")
    st.caption(
        "Primary source: 'napis_na' → parse Property of / Tool number / Part description. "
        "Fallback: lastnik / orodje or sifra_orodja / izdelek."
    )

    custom_uploaded = st.file_uploader(
        "Upload custom CSV or Excel file",
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

                # lower CSV rows = higher priority
                custom_df = custom_df.iloc[::-1].reset_index(drop=True)
                custom_df.insert(0, "priority", range(1, len(custom_df) + 1))
            else:
                custom_df = load_custom_tool_excel(custom_uploaded)

            custom_df["mode"] = mode_default
            custom_df["hole_dia"] = hole_dia_default
            custom_df["hole_offset"] = hole_offset_default

            display_cols = ["priority", "excel_row", "tool_number", "property_of", "part_description"]
            for extra_col in CUSTOM_OPTIONAL_DISPLAY:
                if extra_col in custom_df.columns and extra_col not in display_cols:
                    display_cols.append(extra_col)

            st.markdown("### Select rows for export")
            edited_custom_df = selection_editor(
                custom_df,
                key="custom_selection_editor",
                display_columns=display_cols,
                checkbox_label="Export",
            )

            selected_rows = edited_custom_df[edited_custom_df["print"]].copy()

            info_c1, info_c2 = st.columns([1, 1])
            with info_c1:
                st.info(
                    f"Rows in table: {len(edited_custom_df)} | Selected for export: {len(selected_rows)}"
                )
            with info_c2:
                if not edited_custom_df.empty:
                    top_tool = str(edited_custom_df.iloc[0]["tool_number"])
                    st.caption(f"Highest priority row at top: {top_tool}")

            st.markdown("### Rows queued for export")
            if selected_rows.empty:
                st.info("No rows selected yet. Tick rows in the table above to prepare them for export.")
            else:
                selected_preview_df = selected_rows.drop(columns=["print"]).copy()
                st.dataframe(
                    selected_preview_df,
                    use_container_width=True,
                    hide_index=True,
                )

            preview_source = selected_rows.copy()
            if not preview_source.empty:
                preview_labels = [
                    f"P{int(row['priority'])} | Tool {row['tool_number']} | Excel row {int(row['excel_row'])}"
                    for _, row in preview_source.iterrows()
                ]
                preview_choice = st.selectbox(
                    "Preview tool",
                    preview_labels,
                    key="custom_preview_choice",
                )
                preview_idx = preview_labels.index(preview_choice)
                preview_row = preview_source.iloc[preview_idx]

                preview_svg_raw, _ = generate_svg(
                    owner=str(preview_row["property_of"]),
                    tool_number=str(preview_row["tool_number"]),
                    part_desc=str(preview_row["part_description"]),
                    mode=mode_default,
                    hole_dia=float(hole_dia_default),
                    hole_offset=float(hole_offset_default),
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

                st.markdown("### Selected tool preview")
                render_svg_preview_card(preview_svg_raw, mode_default, preview_scale)

                single_svg_bytes = preview_svg_raw.encode("utf-8")
                single_dxf_bytes = generate_dxf(
                    owner=str(preview_row["property_of"]),
                    tool_number=str(preview_row["tool_number"]),
                    part_desc=str(preview_row["part_description"]),
                    mode=mode_default,
                    hole_dia=float(hole_dia_default),
                    hole_offset=float(hole_offset_default),
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

                preview_base_name = safe_filename(str(preview_row["tool_number"]) or "laser_label")
                p1, p2 = st.columns(2)
                with p1:
                    st.download_button(
                        "Download preview SVG",
                        data=single_svg_bytes,
                        file_name=f"{preview_base_name}.svg",
                        mime="image/svg+xml",
                        use_container_width=True,
                    )
                with p2:
                    st.download_button(
                        "Download preview DXF",
                        data=single_dxf_bytes,
                        file_name=f"{preview_base_name}.dxf",
                        mime="application/dxf",
                        use_container_width=True,
                    )

            st.markdown("### Export selected rows")
            if selected_rows.empty:
                st.warning("Select at least one row in the table above to export SVG + DXF files.")
            else:
                export_df = selected_rows.drop(columns=["print"]).copy()

                zip_bytes, export_manifest_df = build_batch_zip(
                    df=export_df,
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

                st.download_button(
                    f"Download ZIP for selected rows ({len(export_df)})",
                    data=zip_bytes,
                    file_name="laser_labels_selected.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

                with st.expander("Selected export file list"):
                    st.dataframe(
                        export_manifest_df,
                        use_container_width=True,
                        hide_index=True,
                    )

        except Exception as e:
            st.error(f"Could not process custom file: {e}")

with st.expander("Example CSV"):
    st.code(
        """property_of,tool_number,part_description,mode,hole_dia,hole_offset
Stihl Group,89193,Steckzunge BA13-431-2100-A (id. 33193),Anodized aluminium (negative),3.2,4.3
DAFRA,TL-00125,Punch holder,Normal,3.0,4.5
Ledinek,TL-00126,Clamp plate,,,
""",
        language="csv",
    )

with st.expander("Custom Excel / CSV mapping"):
    st.code(
        """Primary parsing source
napis_na -> parse:
- property_of
- tool_number
- part_description

Fallback columns
lastnik -> property_of
orodje or sifra_orodja -> tool_number
izdelek -> part_description

CSV can also use:
napis_na,lastnik,orodje,sifra_orodja,izdelek
""",
        language="text",
    )

with st.expander("Notes"):
    st.write(
        """
Preview scale only affects on-screen display.

Exports remain unchanged:
- SVG stays at 78.5 × 21 mm
- DXF stays in real millimeter geometry

Custom Excel tab behavior:
- detects the header block automatically
- reverses the row order so lower Excel rows get higher priority and appear on top
- user selects rows with checkboxes
- the lower preview table shows only selected rows for export
- one SVG and one DXF is created per selected row
- all generated files are packed into one ZIP
- values are parsed from `napis_na` first, then fallback to separate columns
- long text automatically tries to fit in 2 lines when needed

If you use old .xls files, install xlrd:
pip install xlrd
"""
    )