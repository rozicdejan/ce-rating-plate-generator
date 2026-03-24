import io
import re
import json
import hashlib
import zipfile
import unicodedata
from functools import lru_cache
from pathlib import Path
from datetime import datetime
from collections import Counter

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import ezdxf

from matplotlib.textpath import TextPath
from matplotlib.font_manager import FontProperties
from matplotlib.transforms import Affine2D
from matplotlib.path import Path as MplPath

# ============================================================
# Page config
# ============================================================
st.set_page_config(
    page_title="Laser Label Generator",
    page_icon="🏷️",
    layout="wide",
)

st.markdown(
    """
<style>
section[data-testid="stSidebar"] { padding-top: 1rem; }
div[data-testid="stDownloadButton"] > button { font-weight: 600; letter-spacing: 0.02em; }
div[data-testid="stTabs"] button[role="tab"] { font-weight: 600; font-size: 0.88rem; letter-spacing: 0.03em; }
div[data-testid="stAlert"] { border-radius: 6px; }
hr { margin: 0.75rem 0; border-color: #e0e0e0; }
</style>
""",
    unsafe_allow_html=True,
)

# ============================================================
# Geometry constants
# ============================================================
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

DEFAULT_LEFT_X  = 7.4
DEFAULT_LEFT_W  = 20.5
DEFAULT_RIGHT_X = 30.2

DEFAULT_ROW1_Y = 2.0
DEFAULT_ROW2_Y = 7.0
DEFAULT_ROW3_Y = 12.0

ROW1_H = 4.2
ROW2_H = 4.2
ROW3_H = 3.8

DEFAULT_FS1 = 2.8
DEFAULT_FS2 = 2.8
DEFAULT_FS3 = 1.8

MIN_TEXT_HEIGHT_MM   = 1.2
RIGHT_MARGIN         = 2.2

def _pick_vector_font() -> str:
    """
    Pick the best available condensed/narrow font at runtime.
    Falls back safely to DejaVu Sans which ships with matplotlib.
    """
    from matplotlib.font_manager import fontManager

    available = {f.name for f in fontManager.ttflist}
    for candidate in [
        "DejaVu Sans Condensed",
        "Arial Narrow",
        "Liberation Sans Narrow",
        "Noto Sans Condensed",
        "Roboto Condensed",
        "DejaVu Sans",
        "Liberation Sans",
        "Arial",
    ]:
        if candidate in available:
            return candidate
    return "DejaVu Sans"


VECTOR_FONT_FAMILY = _pick_vector_font()


@lru_cache(maxsize=None)
def get_font_prop(weight: str):
    return FontProperties(family=VECTOR_FONT_FAMILY, weight=weight)


@lru_cache(maxsize=None)
def cap_ref_for_weight(weight: str) -> float:
    """
    Cap-height reference for a capital H at size=1 for the chosen font+weight.
    Used so equal row heights produce equal visual cap heights.
    """
    return TextPath((0, 0), "H", size=1, prop=get_font_prop(weight)).get_extents().y1


PART_STACK_TRIGGER_LEN = 52
PART_STACK_INDENT = 2.1
PART_STACK_LABEL_H = 2.0
PART_STACK_GAP = 0.35
PART_STACK_BOTTOM_MARGIN = 0.7
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

# ============================================================
# Storage paths
# ============================================================
_SCRIPT_DIR = Path(__file__).parent
STORAGE_DIR = _SCRIPT_DIR / "laser_labels_data"
UPLOADS_DIR = STORAGE_DIR / "uploads"
INVENTORY_FILE = STORAGE_DIR / "inventory.json"

STORAGE_DIR.mkdir(parents=True, exist_ok=True)
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# Inventory store
# ============================================================
class InventoryStore:
    def __init__(self, path: Path):
        self.path = path
        self._data = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"tools": [], "files": []}

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def reload(self):
        self._data = self._load()

    @property
    def tools(self) -> list:
        return self._data["tools"]

    @property
    def files(self) -> list:
        return self._data["files"]

    def tool_count(self) -> int:
        return len(self._data["tools"])

    def file_known(self, file_hash: str) -> bool:
        return any(f["hash"] == file_hash for f in self._data["files"])

    def filename_known(self, filename: str) -> bool:
        return any(f["filename"] == filename for f in self._data["files"])

    def as_dataframe(self) -> pd.DataFrame:
        if not self._data["tools"]:
            return pd.DataFrame(
                columns=[
                    "priority",
                    "tool_number",
                    "property_of",
                    "part_description",
                    "source_file",
                    "added_at",
                    "updated_at",
                ]
            )
        df = pd.DataFrame(self._data["tools"])
        return df.sort_values("priority").reset_index(drop=True)

    def files_dataframe(self) -> pd.DataFrame:
        if not self._data["files"]:
            return pd.DataFrame(columns=["filename", "hash", "uploaded_at", "row_count"])
        return pd.DataFrame(self._data["files"])

    def _next_priority(self) -> int:
        if not self._data["tools"]:
            return 1
        return min(t["priority"] for t in self._data["tools"]) - 1

    def merge_from_dataframe(self, df: pd.DataFrame, source_file: str, source_hash: str) -> dict:
        now = datetime.now().isoformat(timespec="seconds")

        if self.file_known(source_hash):
            return {"added": 0, "updated": 0, "unchanged": 0, "skipped": "duplicate_hash"}

        removed_old = 0
        if self.filename_known(source_file):
            before = len(self._data["tools"])
            self._data["tools"] = [
                t for t in self._data["tools"] if t.get("source_file") != source_file
            ]
            removed_old = before - len(self._data["tools"])
            self._data["files"] = [
                f for f in self._data["files"] if f.get("filename") != source_file
            ]

        existing_by_num = {t["tool_number"]: t for t in self._data["tools"]}

        seen_in_batch = set()
        rows = []
        for _, row in df.iterrows():
            tn = str(row.get("tool_number", "")).strip()
            if not tn or tn in seen_in_batch:
                continue
            seen_in_batch.add(tn)
            rows.append(row)

        batch_priority_start = self._next_priority()
        added = updated = unchanged = 0

        for i, row in enumerate(rows):
            tn = str(row.get("tool_number", "")).strip()
            prop_of = str(row.get("property_of", "")).strip()
            part_d = str(row.get("part_description", "")).strip()
            priority = batch_priority_start - i

            if tn in existing_by_num:
                existing = existing_by_num[tn]
                if (
                    existing["property_of"] == prop_of
                    and existing["part_description"] == part_d
                    and existing["source_file"] == source_file
                ):
                    unchanged += 1
                    existing["priority"] = priority
                    existing["updated_at"] = now
                else:
                    existing["property_of"] = prop_of
                    existing["part_description"] = part_d
                    existing["source_file"] = source_file
                    existing["source_hash"] = source_hash
                    existing["priority"] = priority
                    existing["updated_at"] = now
                    updated += 1
            else:
                self._data["tools"].append(
                    {
                        "tool_number": tn,
                        "property_of": prop_of,
                        "part_description": part_d,
                        "source_file": source_file,
                        "source_hash": source_hash,
                        "added_at": now,
                        "updated_at": now,
                        "priority": priority,
                    }
                )
                added += 1

        self._data["files"].append(
            {
                "filename": source_file,
                "hash": source_hash,
                "uploaded_at": now,
                "row_count": len(rows),
            }
        )

        self.save()
        return {
            "added": added,
            "updated": updated,
            "unchanged": unchanged,
            "removed_old": removed_old,
            "skipped": None,
        }

    def delete_tool(self, tool_number: str):
        self._data["tools"] = [t for t in self._data["tools"] if t["tool_number"] != tool_number]
        self.save()

    def delete_file_and_its_tools(self, filename: str):
        self._data["tools"] = [t for t in self._data["tools"] if t.get("source_file") != filename]
        self._data["files"] = [f for f in self._data["files"] if f.get("filename") != filename]
        self.save()

    def clear_all(self):
        self._data = {"tools": [], "files": []}
        self.save()


@st.cache_resource
def get_store() -> InventoryStore:
    return InventoryStore(INVENTORY_FILE)


def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def save_upload_to_disk(data: bytes, filename: str) -> Path:
    dest = UPLOADS_DIR / filename
    if dest.exists():
        existing_hash = file_sha256(dest.read_bytes())
        new_hash = file_sha256(data)
        if existing_hash != new_hash:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = Path(filename).stem
            suf = Path(filename).suffix
            dest = UPLOADS_DIR / f"{stem}_{ts}{suf}"
    dest.write_bytes(data)
    return dest


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
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
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
        a, b = a.strip(), b.strip()
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
        pairs.append((flat[: m.start()].strip(), flat[m.start() :].strip()))

    for m in re.finditer(r"\s+(?=(?:2[\.\)]|2[-\u2013]))", flat):
        pairs.append((flat[: m.start()].strip(), flat[m.start() :].strip()))

    for pattern in [r"\s*[;|]\s*", r"\s+/\s+", r"\s+-\s+", r",\s+"]:
        for m in re.finditer(pattern, flat):
            pairs.append((flat[: m.start()].strip(), flat[m.end() :].strip()))

    words = flat.split()
    if len(words) >= 2:
        best_i, best_diff = 1, 10**9
        for i in range(1, len(words)):
            diff = abs(len(" ".join(words[:i])) - len(" ".join(words[i:])))
            if diff < best_diff:
                best_diff, best_i = diff, i
        pairs.append((" ".join(words[:best_i]), " ".join(words[best_i:])))

    return unique_pairs(pairs)


def parse_napis_na(text: str) -> dict:
    raw = "" if text is None or pd.isna(text) else str(text)
    if not raw.strip():
        return {"parsed_property_of": "", "parsed_tool_number": "", "parsed_part_description": ""}

    s = raw.replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s).strip()
    flat = normalize_space(s)

    result = {"parsed_property_of": "", "parsed_tool_number": "", "parsed_part_description": ""}

    m = re.search(
        r"(?:property\s*of|lastnik|possessor|eigentümer)\s*:\s*(.+?)"
        r"(?=(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?"
        r"|part\s*1\s*:|part\s*description\s*:|izdelek\s*:|product\s*:|produkt\s*:|$))",
        flat,
        flags=re.IGNORECASE,
    )
    if m:
        result["parsed_property_of"] = clean_parsed_value(m.group(1))

    m = re.search(
        r"(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?)"
        r"\s*:\s*([A-Za-z0-9./_-]+)",
        flat,
        flags=re.IGNORECASE,
    )
    if m:
        result["parsed_tool_number"] = clean_parsed_value(m.group(1))

    m = re.search(r"((?:part\s*1)\s*:\s*.+)$", flat, flags=re.IGNORECASE)
    if m:
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", m.group(1).strip(), count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()
        return result

    m = re.search(
        r"(?:(?:part\s*description)|(?:izdelek)|(?:product)|(?:produkt))\s*:\s*(.+)$",
        flat,
        flags=re.IGNORECASE,
    )
    if m:
        desc = clean_parsed_value(m.group(1))
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", desc, count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()
        return result

    m = re.search(
        r"(?:tool\s*number|tool\s*no\.?|orodje\s*\u0161t\.?|orodje\s*st\.?|wkz\.\s*no\.?)"
        r"\s*:\s*[A-Za-z0-9./_-]+\s+(.+)$",
        flat,
        flags=re.IGNORECASE,
    )
    if m:
        desc = clean_parsed_value(m.group(1))
        desc = re.sub(r"\s+(?=(?:part\s*2\s*:))", "\n", desc, count=1, flags=re.IGNORECASE)
        result["parsed_part_description"] = desc.strip()

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


def build_rows(owner, tool_number, part_desc):
    return [
        ("Property of:", owner),
        ("Tool number:", tool_number),
        ("Part description:", part_desc),
    ]


def make_base_text_path(text, weight):
    if not text:
        return None, None
    tp = TextPath((0, 0), text, size=1, prop=get_font_prop(weight))
    bbox = tp.get_extents()
    if bbox.width <= 0 or bbox.height <= 0:
        return None, None
    return tp, bbox


def fit_text_for_box(text, desired_h_mm, box_w_mm, box_h_mm, weight):
    """
    Fit text into a box.

    Returns:
        candidate, base_path, bbox, scale, used_cap_h, used_bbox_h
    """
    raw = (text or "").strip()
    if not raw:
        return "", None, None, 1.0, 0.0, 0.0

    cap_ref = cap_ref_for_weight(weight)

    for cut in range(0, len(raw) + 1):
        candidate = raw if cut == 0 else raw[: max(0, len(raw) - cut)].rstrip() + "..."
        base_path, bbox = make_base_text_path(candidate, weight)
        if base_path is None or bbox is None:
            continue

        scale_desired = desired_h_mm / cap_ref
        scale_box_h = box_h_mm / bbox.height
        scale = min(scale_desired, scale_box_h)

        used_cap_h = cap_ref * scale
        used_bbox_h = bbox.height * scale

        if bbox.width * scale <= box_w_mm:
            return candidate, base_path, bbox, scale, used_cap_h, used_bbox_h

        scale_w = box_w_mm / bbox.width
        used_cap_h_w = cap_ref * scale_w
        used_bbox_h_w = bbox.height * scale_w
        if used_cap_h_w >= MIN_TEXT_HEIGHT_MM and used_bbox_h_w <= box_h_mm:
            return candidate, base_path, bbox, scale_w, used_cap_h_w, used_bbox_h_w

    fallback = "..."
    base_path, bbox = make_base_text_path(fallback, weight)
    if base_path is None or bbox is None:
        return "", None, None, 1.0, 0.0, 0.0

    scale_desired = desired_h_mm / cap_ref
    scale_box_h = box_h_mm / bbox.height
    scale_width = box_w_mm / bbox.width
    scale = min(scale_desired, scale_box_h, scale_width)
    return fallback, base_path, bbox, scale, cap_ref * scale, bbox.height * scale


def text_fits_single_line(text, desired_h_mm, box_w_mm, box_h_mm, weight):
    raw = "" if text is None or pd.isna(text) else str(text).strip()
    raw = normalize_space(raw.replace("\n", " "))
    if not raw:
        return True

    base_path, bbox = make_base_text_path(raw, weight)
    if base_path is None or bbox is None:
        return True

    cap_ref = cap_ref_for_weight(weight)
    scale_desired = desired_h_mm / cap_ref
    return (
        bbox.width * scale_desired <= box_w_mm
        and bbox.height * scale_desired <= box_h_mm
    )


def place_path_in_box(base_path, scale, box_x, box_y, box_w, box_h, align="left", pad_x=0.0, cap_h=None):
    """
    Place a scaled glyph path inside a row box.
    Vertical placement centers the cap-height portion in the box.
    """
    path = base_path.transformed(Affine2D().scale(scale, -scale))
    bbox = path.get_extents()
    tx = (
        box_x + (box_w - bbox.width) / 2.0 - bbox.x0
        if align == "center"
        else box_x + box_w - pad_x - bbox.x1
        if align == "right"
        else box_x + pad_x - bbox.x0
    )

    if cap_h is None:
        cap_h = -bbox.y0

    ty = box_y + (box_h + cap_h) / 2.0
    return path.transformed(Affine2D().translate(tx, ty))


def place_path_top_aligned(base_path, scale, box_x, top_y, box_w, align="left", pad_x=0.0):
    path = base_path.transformed(Affine2D().scale(scale, -scale))
    bbox = path.get_extents()
    tx = (
        box_x + (box_w - bbox.width) / 2.0 - bbox.x0
        if align == "center"
        else box_x + box_w - pad_x - bbox.x1
        if align == "right"
        else box_x + pad_x - bbox.x0
    )
    ty = top_y - bbox.y0
    return path.transformed(Affine2D().translate(tx, ty))


def fit_text_block(
    text,
    desired_h_mm,
    box_x,
    box_y,
    box_w,
    box_h,
    weight,
    align="left",
    pad_x=0.0,
    max_lines=1,
    first_line_anchor_h=None,
    force_single_if_fits=False,
):
    raw_original = "" if text is None or pd.isna(text) else str(text).strip()
    raw_single = normalize_space(raw_original.replace("\n", " "))

    empty = {
        "texts": [],
        "paths": [],
        "used_heights": [],
        "used_bbox_heights": [],
        "used_height_summary": 0.0,
        "line_boxes": [],
        "line_count": 0,
    }
    if not raw_original:
        return empty

    sf, sb, sbbox, ss, sh_cap, sh_bbox = fit_text_for_box(raw_single, desired_h_mm, box_w, box_h, weight)
    single_complete = sf == raw_single

    if single_complete and sbbox is not None:
        cap_ref = cap_ref_for_weight(weight)
        single_truly_fits = (
            sbbox.width * (desired_h_mm / cap_ref) <= box_w
            and sbbox.height * (desired_h_mm / cap_ref) <= box_h
        )
    else:
        single_truly_fits = False

    best = {
        "texts": [sf],
        "bases": [sb],
        "scales": [ss],
        "used_heights": [sh_cap],
        "used_bbox_heights": [sh_bbox],
        "line_box_h": box_h,
        "gap": 0.0,
        "score": (1000 if single_complete else 0) + sh_cap * 100 + len(sf),
    }

    if (
        not single_truly_fits
        and max_lines >= 2
        and box_h >= (MIN_TEXT_HEIGHT_MM * 2 + MULTILINE_GAP_MM)
    ):
        gap = min(MULTILINE_GAP_MM, max(0.25, box_h * 0.08))
        lbh = (box_h - gap) / 2.0

        if lbh >= MIN_TEXT_HEIGHT_MM:
            for l1, l2 in build_two_line_candidates(raw_original):
                t1, b1, _, s1, h1_cap, h1_bbox = fit_text_for_box(l1, desired_h_mm, box_w, lbh, weight)
                t2, b2, _, s2, h2_cap, h2_bbox = fit_text_for_box(l2, desired_h_mm, box_w, lbh, weight)
                if not t1 or not t2:
                    continue

                complete = t1 == l1 and t2 == l2
                score = (
                    (2000 if complete else 0)
                    + min(h1_cap, h2_cap) * 100
                    + len(t1) + len(t2)
                    - abs(len(l1) - len(l2)) * 0.15
                )
                if complete and not single_complete:
                    score += 500
                if score > best["score"]:
                    best = {
                        "texts": [t1, t2],
                        "bases": [b1, b2],
                        "scales": [s1, s2],
                        "used_heights": [h1_cap, h2_cap],
                        "used_bbox_heights": [h1_bbox, h2_bbox],
                        "line_box_h": lbh,
                        "gap": gap,
                        "score": score,
                    }

    paths, line_boxes = [], []
    if len(best["texts"]) == 1:
        centering_h = min(first_line_anchor_h, box_h) if first_line_anchor_h is not None else box_h
        placed = (
            place_path_in_box(
                best["bases"][0],
                best["scales"][0],
                box_x,
                box_y,
                box_w,
                centering_h,
                align=align,
                pad_x=pad_x,
                cap_h=best["used_heights"][0],
            )
            if best["bases"][0] is not None
            else None
        )
        if placed is not None:
            paths.append(placed)
        line_boxes.append((box_x, box_y, box_w, centering_h))
        used_h = best["used_heights"][0] if best["used_heights"] else 0.0
    else:
        bbox_h1, bbox_h2 = best["used_bbox_heights"][0], best["used_bbox_heights"][1]
        gap = best["gap"]
        anchor_h = min(first_line_anchor_h if first_line_anchor_h is not None else bbox_h1, box_h)
        flt = box_y + max(0.0, (anchor_h - bbox_h1) / 2.0)
        slt = flt + bbox_h1 + gap

        placed1 = (
            place_path_top_aligned(best["bases"][0], best["scales"][0], box_x, flt, box_w, align=align, pad_x=pad_x)
            if best["bases"][0] is not None
            else None
        )
        placed2 = (
            place_path_top_aligned(best["bases"][1], best["scales"][1], box_x, slt, box_w, align=align, pad_x=pad_x)
            if best["bases"][1] is not None
            else None
        )
        if placed1 is not None:
            paths.append(placed1)
        if placed2 is not None:
            paths.append(placed2)

        line_boxes.extend([(box_x, flt, box_w, bbox_h1), (box_x, slt, box_w, bbox_h2)])
        used_h = best["used_heights"][0] + gap + best["used_heights"][1]

    return {
        "texts": best["texts"],
        "paths": paths,
        "used_heights": best["used_heights"],
        "used_bbox_heights": best["used_bbox_heights"],
        "used_height_summary": used_h,
        "line_boxes": line_boxes,
        "line_count": len(best["texts"]),
    }


def block_line_height(block):
    hs = [float(h) for h in block.get("used_heights", []) if h and h > 0]
    return min(hs) if hs else 0.0


def fit_row_pair_same_height(left_text, right_text, desired_h_mm, left_cfg, right_cfg):
    left_block = fit_text_block(left_text, desired_h_mm, **left_cfg)
    right_block = fit_text_block(right_text, desired_h_mm, **right_cfg)

    lh = block_line_height(left_block)
    rh = block_line_height(right_block)

    common_h = max(MIN_TEXT_HEIGHT_MM, min(lh, rh)) if lh > 0 and rh > 0 else desired_h_mm

    left_block = fit_text_block(left_text, common_h, **left_cfg)
    right_block = fit_text_block(right_text, common_h, **right_cfg)
    return left_block, right_block, common_h


def mpl_path_to_svg_d(path_obj):
    if path_obj is None:
        return ""
    out = []
    for verts, code in path_obj.iter_segments():
        if code == MplPath.MOVETO:
            out.append(f"M {verts[0]:.4f} {verts[1]:.4f}")
        elif code == MplPath.LINETO:
            out.append(f"L {verts[0]:.4f} {verts[1]:.4f}")
        elif code == MplPath.CURVE3:
            out.append(f"Q {verts[0]:.4f} {verts[1]:.4f} {verts[2]:.4f} {verts[3]:.4f}")
        elif code == MplPath.CURVE4:
            out.append(
                f"C {verts[0]:.4f} {verts[1]:.4f} {verts[2]:.4f} {verts[3]:.4f} {verts[4]:.4f} {verts[5]:.4f}"
            )
        elif code == MplPath.CLOSEPOLY:
            out.append("Z")
    return " ".join(out)


# ============================================================
# DXF helpers
# ============================================================
def _approx_cubic_bezier(p0, p1, p2, p3, n=10):
    pts = []
    for i in range(n + 1):
        t = i / n
        mt = 1.0 - t
        x = mt**3 * p0[0] + 3 * mt**2 * t * p1[0] + 3 * mt * t**2 * p2[0] + t**3 * p3[0]
        y = mt**3 * p0[1] + 3 * mt**2 * t * p1[1] + 3 * mt * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def _approx_quad_bezier(p0, p1, p2, n=8):
    pts = []
    for i in range(n + 1):
        t = i / n
        mt = 1.0 - t
        x = mt**2 * p0[0] + 2 * mt * t * p1[0] + t**2 * p2[0]
        y = mt**2 * p0[1] + 2 * mt * t * p1[1] + t**2 * p2[1]
        pts.append((x, y))
    return pts


def mpl_path_to_dxf_polylines(msp, path_obj, layer, label_h=LABEL_H):
    if path_obj is None:
        return

    def flip(x, y):
        return (float(x), label_h - float(y))

    contours = []
    current = []
    cur_pos = (0.0, 0.0)

    for verts, code in path_obj.iter_segments(simplify=False):
        if code == MplPath.MOVETO:
            if len(current) >= 2:
                contours.append(current)
            current = [flip(verts[0], verts[1])]
            cur_pos = (float(verts[0]), float(verts[1]))
        elif code == MplPath.LINETO:
            current.append(flip(verts[0], verts[1]))
            cur_pos = (float(verts[0]), float(verts[1]))
        elif code == MplPath.CURVE3:
            p1 = (float(verts[0]), float(verts[1]))
            p2 = (float(verts[2]), float(verts[3]))
            for pt in _approx_quad_bezier(cur_pos, p1, p2)[1:]:
                current.append(flip(pt[0], pt[1]))
            cur_pos = p2
        elif code == MplPath.CURVE4:
            p1 = (float(verts[0]), float(verts[1]))
            p2 = (float(verts[2]), float(verts[3]))
            p3 = (float(verts[4]), float(verts[5]))
            for pt in _approx_cubic_bezier(cur_pos, p1, p2, p3)[1:]:
                current.append(flip(pt[0], pt[1]))
            cur_pos = p3
        elif code == MplPath.CLOSEPOLY:
            if len(current) >= 2:
                contours.append(current)
            current = []

    if len(current) >= 2:
        contours.append(current)

    for contour in contours:
        if len(contour) >= 2:
            msp.add_lwpolyline(contour, close=True, dxfattribs={"layer": layer})


def add_rounded_rect_dxf(msp, x, y, w, h, r, layer):
    r = max(0.0, min(r, w / 2.0, h / 2.0))
    if r <= 0:
        msp.add_lwpolyline([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], close=True, dxfattribs={"layer": layer})
        return
    for start, end, _, _ in [
        ((x + r, y), (x + w - r, y), None, None),
        ((x + w, y + r), (x + w, y + h - r), None, None),
        ((x + w - r, y + h), (x + r, y + h), None, None),
        ((x, y + h - r), (x, y + r), None, None),
    ]:
        msp.add_line(start, end, dxfattribs={"layer": layer})
    msp.add_arc((x + w - r, y + r), r, 270, 360, dxfattribs={"layer": layer})
    msp.add_arc((x + w - r, y + h - r), r, 0, 90, dxfattribs={"layer": layer})
    msp.add_arc((x + r, y + h - r), r, 90, 180, dxfattribs={"layer": layer})
    msp.add_arc((x + r, y + r), r, 180, 270, dxfattribs={"layer": layer})


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
    if name.endswith((".xlsx", ".xls")):
        try:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file)
        except Exception:
            uploaded_file.seek(0)
            engine = "xlrd" if name.endswith(".xls") else "openpyxl"
            return pd.read_excel(uploaded_file, engine=engine)
    raise ValueError("Unsupported file type. Use CSV or XLSX/XLS.")


def make_unique_base_names(names):
    counts, result = Counter(), []
    for n in names:
        counts[n] += 1
        result.append(n if counts[n] == 1 else f"{n}_{counts[n]}")
    return result


def make_preview_svg(svg_text, label_w_mm, label_h_mm, px_per_mm):
    pw = label_w_mm * px_per_mm
    ph = label_h_mm * px_per_mm
    s = re.sub(r'width="[^"]+"', f'width="{pw}px"', svg_text, count=1)
    s = re.sub(r'height="[^"]+"', f'height="{ph}px"', s, count=1)
    return s


def selection_editor(df, key, display_columns, checkbox_label="Print"):
    work_df = df.copy()
    if "print" not in work_df.columns:
        work_df.insert(0, "print", False)
    col_order = ["print"] + [c for c in display_columns if c in work_df.columns]
    return st.data_editor(
        work_df[col_order],
        key=key,
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        disabled=[c for c in col_order if c != "print"],
        column_config={
            "print": st.column_config.CheckboxColumn(
                checkbox_label,
                help="Select rows for SVG + DXF export",
                default=False,
            )
        },
    )


def render_svg_preview_card(svg_output, mode, preview_scale):
    colors = get_mode_colors(mode)
    preview_svg = make_preview_svg(svg_output, LABEL_W, LABEL_H, preview_scale)
    ph = int(LABEL_H * preview_scale + 80)
    components.html(
        f"""<div style="background:{colors['preview_bg']};padding:20px;border-radius:12px;
            border:1px solid #d0d0d0;min-height:{ph}px;display:flex;
            justify-content:center;align-items:center;overflow:auto;">
            <div style="padding:10px;border-radius:10px;">{preview_svg}</div></div>""",
        height=ph + 40,
    )


# ============================================================
# Shared label layout
# ============================================================
def should_use_stacked_part_layout(part_desc, wrap_trigger=PART_STACK_TRIGGER_LEN):
    raw = "" if part_desc is None or pd.isna(part_desc) else str(part_desc)
    flat = normalize_space(raw.replace("\n", " "))
    return "\n" in raw or bool(re.search(r"part\s*2\s*:", raw, re.IGNORECASE)) or len(flat) >= wrap_trigger


def split_part_description_lines(part_desc):
    raw = "" if part_desc is None or pd.isna(part_desc) else str(part_desc).strip()
    if not raw:
        return []
    raw = re.sub(r"\n+", "\n", raw.replace("\r", "\n")).strip()
    lines = [p.strip() for p in raw.split("\n") if p.strip()]
    if len(lines) >= 2:
        return lines[:2]
    flat = normalize_space(raw)
    if re.search(r"part\s*1\s*:", flat, re.IGNORECASE) and re.search(r"part\s*2\s*:", flat, re.IGNORECASE):
        split = re.split(r"(?=part\s*2\s*:)", flat, maxsplit=1, flags=re.IGNORECASE)
        if len(split) == 2:
            return [split[0].strip(), split[1].strip()]
    candidates = build_two_line_candidates(raw)
    return [candidates[0][0], candidates[0][1]] if candidates else [flat]


def get_part_stack_layout(left_x, border_offset, row3_y, right_margin=RIGHT_MARGIN):
    vx = left_x + PART_STACK_INDENT
    vy = row3_y + PART_STACK_LABEL_H + PART_STACK_GAP
    vw = LABEL_W - border_offset - right_margin - vx
    vht = LABEL_H - PART_STACK_BOTTOM_MARGIN - vy
    return {
        "label_x": left_x,
        "label_y": row3_y,
        "label_w": LABEL_W - left_x - right_margin,
        "label_h": PART_STACK_LABEL_H,
        "value_x": vx,
        "value_y": vy,
        "value_w": vw,
        "value_h_total": vht,
        "line_gap": PART_STACK_GAP,
        "per_line_h": max(MIN_TEXT_HEIGHT_MM, (vht - PART_STACK_GAP) / 2.0),
    }


def compute_label_layout(
    owner,
    tool_number,
    part_desc,
    left_x,
    left_w,
    right_x,
    row1_y,
    row2_y,
    row3_y,
    fs1,
    fs2,
    fs3,
    border_offset,
    right_margin=RIGHT_MARGIN,
    wrap_trigger=PART_STACK_TRIGGER_LEN,
):
    rows = build_rows(owner, tool_number, part_desc)
    right_w = LABEL_W - right_x - right_margin

    part_overflows_inline = not text_fits_single_line(part_desc, fs3, right_w, ROW3_H, "regular")
    use_stacked = should_use_stacked_part_layout(part_desc, wrap_trigger=wrap_trigger) or part_overflows_inline

    lb0, rb0, _ = fit_row_pair_same_height(
        rows[0][0],
        rows[0][1],
        fs1,
        left_cfg=dict(box_x=left_x, box_y=row1_y, box_w=left_w, box_h=ROW1_H, weight="bold", max_lines=1),
        right_cfg=dict(
            box_x=right_x,
            box_y=row1_y,
            box_w=right_w,
            box_h=row2_y - row1_y,
            weight="regular",
            max_lines=2,
            first_line_anchor_h=ROW1_H,
            force_single_if_fits=True,
        ),
    )

    lb1, rb1, _ = fit_row_pair_same_height(
        rows[1][0],
        rows[1][1],
        fs2,
        left_cfg=dict(box_x=left_x, box_y=row2_y, box_w=left_w, box_h=ROW2_H, weight="bold", max_lines=1),
        right_cfg=dict(box_x=right_x, box_y=row2_y, box_w=right_w, box_h=ROW2_H, weight="regular", max_lines=1),
    )

    if use_stacked:
        stack = get_part_stack_layout(left_x, border_offset, row3_y, right_margin=right_margin)
        part_lines = split_part_description_lines(part_desc)
        part_lines = (part_lines + [""])[:2]

        lb2 = fit_text_block(
            "Part description:",
            fs3,
            stack["label_x"],
            stack["label_y"],
            stack["label_w"],
            stack["label_h"],
            "bold",
            max_lines=1,
        )
        rb2a = fit_text_block(
            part_lines[0],
            fs3,
            stack["value_x"],
            stack["value_y"],
            stack["value_w"],
            stack["per_line_h"],
            "regular",
            max_lines=1,
        )
        rb2b = (
            fit_text_block(
                part_lines[1],
                fs3,
                stack["value_x"],
                stack["value_y"] + stack["per_line_h"] + stack["line_gap"],
                stack["value_w"],
                stack["per_line_h"],
                "regular",
                max_lines=1,
            )
            if part_lines[1]
            else {
                "paths": [],
                "texts": [],
                "used_height_summary": 0.0,
                "line_boxes": [],
                "used_heights": [],
                "used_bbox_heights": [],
                "line_count": 0,
            }
        )
        row2 = {
            "stacked": True,
            "stack": stack,
            "part_lines": part_lines,
            "label_block": lb2,
            "line1_block": rb2a,
            "line2_block": rb2b,
        }
    else:
        ph = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_STACK_BOTTOM_MARGIN))
        lb2, rb2, _ = fit_row_pair_same_height(
            "Part description:",
            part_desc,
            fs3,
            left_cfg=dict(box_x=left_x, box_y=row3_y, box_w=left_w, box_h=ROW3_H, weight="bold", max_lines=1),
            right_cfg=dict(
                box_x=right_x,
                box_y=row3_y,
                box_w=right_w,
                box_h=ph,
                weight="regular",
                max_lines=2,
                first_line_anchor_h=ROW3_H,
            ),
        )
        row2 = {"stacked": False, "left_block": lb2, "right_block": rb2}

    return {
        "use_stacked_part": use_stacked,
        "right_w": right_w,
        "rows": rows,
        "left_block_0": lb0,
        "right_block_0": rb0,
        "left_block_1": lb1,
        "right_block_1": rb1,
        "row2_blocks": row2,
        "row1_y": row1_y,
        "row2_y": row2_y,
        "row3_y": row3_y,
    }


# ============================================================
# SVG generation
# ============================================================
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
    show_holes=True,
    right_margin=RIGHT_MARGIN,
    wrap_trigger=PART_STACK_TRIGGER_LEN,
):
    colors = get_mode_colors(mode)
    layout = compute_label_layout(
        owner,
        tool_number,
        part_desc,
        left_x,
        left_w,
        right_x,
        row1_y,
        row2_y,
        row3_y,
        fs1,
        fs2,
        fs3,
        border_offset,
        right_margin=right_margin,
        wrap_trigger=wrap_trigger,
    )
    right_w = layout["right_w"]
    row2 = layout["row2_blocks"]

    all_paths = []
    meta = {"left_sizes": [], "right_sizes": [], "left_texts": [], "right_texts": [], "right_w": right_w, "stacked_part": layout["use_stacked_part"]}

    def collect(block, side):
        for p in block["paths"]:
            d = mpl_path_to_svg_d(p)
            if d:
                all_paths.append(d)
        if side == "left":
            meta["left_sizes"].append(block["used_height_summary"])
            meta["left_texts"].append(" / ".join(block["texts"]))
        else:
            meta["right_sizes"].append(block["used_height_summary"])
            meta["right_texts"].append(" / ".join(block["texts"]))

    collect(layout["left_block_0"], "left")
    collect(layout["right_block_0"], "right")
    collect(layout["left_block_1"], "left")
    collect(layout["right_block_1"], "right")

    if row2["stacked"]:
        for blk in [row2["label_block"], row2["line1_block"], row2["line2_block"]]:
            for p in blk["paths"]:
                d = mpl_path_to_svg_d(p)
                if d:
                    all_paths.append(d)
        meta["left_sizes"].append(row2["label_block"]["used_height_summary"])
        meta["right_sizes"].append(max(row2["line1_block"].get("used_height_summary", 0), row2["line2_block"].get("used_height_summary", 0)))
        meta["left_texts"].append("Part description:")
        meta["right_texts"].append(" / ".join(x for x in row2["part_lines"] if x))
    else:
        for blk in [row2["left_block"], row2["right_block"]]:
            for p in blk["paths"]:
                d = mpl_path_to_svg_d(p)
                if d:
                    all_paths.append(d)
        meta["left_sizes"].append(row2["left_block"]["used_height_summary"])
        meta["right_sizes"].append(row2["right_block"]["used_height_summary"])
        meta["left_texts"].append("Part description:")
        meta["right_texts"].append(" / ".join(row2["right_block"]["texts"]))

    border_svg = (
        f'<rect x="{border_offset}" y="{border_offset}" '
        f'width="{LABEL_W - 2 * border_offset}" height="{LABEL_H - 2 * border_offset}" '
        f'rx="{corner_r}" ry="{corner_r}" fill="{colors["plate_fill"]}" '
        f'stroke="{"none" if not show_border else colors["plate_stroke"]}" stroke-width="0.20"/>'
    )

    hr = hole_dia / 2.0
    hy = LABEL_H / 2.0
    hlx = hole_offset
    hrx = LABEL_W - hole_offset

    if show_holes:
        holes_svg = (
            f'<circle cx="{hlx}" cy="{hy}" r="{hr}" fill="{colors["hole_fill"]}" stroke="{colors["plate_stroke"]}" stroke-width="0.15"/>'
            f'<circle cx="{hrx}" cy="{hy}" r="{hr}" fill="{colors["hole_fill"]}" stroke="{colors["plate_stroke"]}" stroke-width="0.15"/>'
        )
    else:
        holes_svg = (
            f'<circle cx="{hlx}" cy="{hy}" r="{hr}" fill="none" stroke="{colors["guide_stroke"]}" stroke-width="0.12" stroke-dasharray="0.6,0.6" opacity="0.55"/>'
            f'<circle cx="{hrx}" cy="{hy}" r="{hr}" fill="none" stroke="{colors["guide_stroke"]}" stroke-width="0.12" stroke-dasharray="0.6,0.6" opacity="0.55"/>'
        )

    guides_svg = ""
    if show_guides:
        rw = right_w
        if row2["stacked"]:
            s = row2["stack"]
            guides_svg = (
                f'<g fill="none" stroke="{colors["guide_stroke"]}" stroke-width="0.12" stroke-dasharray="0.8,0.8" opacity="0.75">'
                f'<rect x="{left_x}" y="{row1_y}" width="{left_w}" height="{ROW1_H}"/>'
                f'<rect x="{right_x}" y="{row1_y}" width="{rw}" height="{row2_y - row1_y}"/>'
                f'<rect x="{left_x}" y="{row2_y}" width="{left_w}" height="{ROW2_H}"/>'
                f'<rect x="{right_x}" y="{row2_y}" width="{rw}" height="{ROW2_H}"/>'
                f'<rect x="{s["label_x"]}" y="{s["label_y"]}" width="{s["label_w"]}" height="{s["label_h"]}"/>'
                f'<rect x="{s["value_x"]}" y="{s["value_y"]}" width="{s["value_w"]}" height="{s["per_line_h"]}"/>'
                f'<rect x="{s["value_x"]}" y="{s["value_y"] + s["per_line_h"] + s["line_gap"]}" width="{s["value_w"]}" height="{s["per_line_h"]}"/>'
                f'</g>'
            )
        else:
            ph = max(ROW3_H, LABEL_H - row3_y - max(border_offset, PART_STACK_BOTTOM_MARGIN))
            guides_svg = (
                f'<g fill="none" stroke="{colors["guide_stroke"]}" stroke-width="0.12" stroke-dasharray="0.8,0.8" opacity="0.75">'
                f'<rect x="{left_x}" y="{row1_y}" width="{left_w}" height="{ROW1_H}"/>'
                f'<rect x="{right_x}" y="{row1_y}" width="{rw}" height="{row2_y - row1_y}"/>'
                f'<rect x="{left_x}" y="{row2_y}" width="{left_w}" height="{ROW2_H}"/>'
                f'<rect x="{right_x}" y="{row2_y}" width="{rw}" height="{ROW2_H}"/>'
                f'<rect x="{left_x}" y="{row3_y}" width="{left_w}" height="{ROW3_H}"/>'
                f'<rect x="{right_x}" y="{row3_y}" width="{rw}" height="{ph}"/>'
                f'</g>'
            )

    text_svg = "\n".join(f'<path d="{d}"/>' for d in all_paths)
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{LABEL_W}mm" height="{LABEL_H}mm" viewBox="0 0 {LABEL_W} {LABEL_H}">'
        f"{border_svg}{holes_svg}{guides_svg}"
        f'<g fill="{colors["text_fill"]}" fill-rule="evenodd" stroke="none">{text_svg}</g>'
        f"</svg>"
    )
    return svg, meta


# ============================================================
# DXF generation
# ============================================================
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
    show_holes=True,
    right_margin=RIGHT_MARGIN,
    wrap_trigger=PART_STACK_TRIGGER_LEN,
):
    layout = compute_label_layout(
        owner,
        tool_number,
        part_desc,
        left_x,
        left_w,
        right_x,
        row1_y,
        row2_y,
        row3_y,
        fs1,
        fs2,
        fs3,
        border_offset,
        right_margin=right_margin,
        wrap_trigger=wrap_trigger,
    )
    row2 = layout["row2_blocks"]

    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4
    for ln in ["BORDER", "HOLES", "TEXT", "FILL"]:
        if ln not in doc.layers:
            doc.layers.new(ln)
    msp = doc.modelspace()

    if show_border:
        add_rounded_rect_dxf(msp, border_offset, border_offset, LABEL_W - 2 * border_offset, LABEL_H - 2 * border_offset, corner_r, "BORDER")

    if show_holes:
        hr = hole_dia / 2.0
        hy = LABEL_H / 2.0
        msp.add_circle((hole_offset, hy), hr, dxfattribs={"layer": "HOLES"})
        msp.add_circle((LABEL_W - hole_offset, hy), hr, dxfattribs={"layer": "HOLES"})

    def emit(block):
        for placed_path in block.get("paths", []):
            mpl_path_to_dxf_polylines(msp, placed_path, "TEXT")

    emit(layout["left_block_0"])
    emit(layout["right_block_0"])
    emit(layout["left_block_1"])
    emit(layout["right_block_1"])
    if row2["stacked"]:
        emit(row2["label_block"])
        emit(row2["line1_block"])
        if row2["line2_block"].get("texts"):
            emit(row2["line2_block"])
    else:
        emit(row2["left_block"])
        emit(row2["right_block"])

    stream = io.StringIO()
    doc.write(stream)
    return stream.getvalue().encode("utf-8")


# ============================================================
# ZIP batch generation
# ============================================================
def build_batch_zip(
    df,
    default_mode,
    default_hole_dia,
    default_hole_offset,
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
    show_border,
    show_holes,
    include_svg=True,
    include_dxf=True,
    progress_bar=None,
    right_margin=RIGHT_MARGIN,
    wrap_trigger=PART_STACK_TRIGGER_LEN,
):
    required = ["property_of", "tool_number", "part_description"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    names = [safe_filename(str(r.get("tool_number", "")).strip() or "laser_label") for _, r in df.iterrows()]
    unique_names = make_unique_base_names(names)
    buf, records, total = io.BytesIO(), [], len(df)

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (_, row) in enumerate(df.iterrows()):
            owner = "" if pd.isna(row["property_of"]) else str(row["property_of"])
            tn = "" if pd.isna(row["tool_number"]) else str(row["tool_number"])
            pd_val = "" if pd.isna(row["part_description"]) else str(row["part_description"])
            mode = parse_mode(row.get("mode", None), default_mode)
            hd = parse_optional_float(row.get("hole_dia", row.get("hole_size", None)), default_hole_dia)
            ho = parse_optional_float(row.get("hole_offset", None), default_hole_offset)
            base = unique_names[idx]
            kw = dict(
                owner=owner,
                tool_number=tn,
                part_desc=pd_val,
                mode=mode,
                hole_dia=hd,
                hole_offset=ho,
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
                show_holes=show_holes,
                right_margin=right_margin,
                wrap_trigger=wrap_trigger,
            )
            if include_svg:
                svg_out, _ = generate_svg(**kw, show_guides=False, show_border=show_border)
                zf.writestr(f"{base}.svg", svg_out.encode("utf-8"))
            if include_dxf:
                zf.writestr(f"{base}.dxf", generate_dxf(**kw, show_border=show_border))

            records.append(
                {
                    "file_base": base,
                    "property_of": owner,
                    "tool_number": tn,
                    "part_description": pd_val,
                    "mode": mode,
                    "hole_dia": hd,
                    "hole_offset": ho,
                    "holes_exported": show_holes,
                }
            )
            if progress_bar:
                progress_bar.progress((idx + 1) / total, text=f"Generating {idx + 1} / {total}")

    buf.seek(0)
    return buf.getvalue(), pd.DataFrame(records)


# ============================================================
# Custom Excel / CSV parsing helpers
# ============================================================
def _score_header_block(raw_df, start_row):
    parts = []
    for r in range(start_row, min(start_row + 3, len(raw_df))):
        parts.extend([str(v) for v in raw_df.iloc[r].tolist() if pd.notna(v)])
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
    best_sheet, best_row, best_score = None, None, -(10**9)
    for sheet in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if raw_df.empty:
            continue
        for sr in range(min(20, len(raw_df))):
            sc = _score_header_block(raw_df, sr)
            if sc > best_score:
                best_score, best_sheet, best_row = sc, sheet, sr
    if best_sheet is None or best_score < 6:
        raise ValueError("Could not detect custom Excel header. Ensure columns: lastnik / orodje / sifra_orodja / izdelek.")
    return best_sheet, best_row


def combine_multirow_headers(raw_df, header_row, depth=3):
    headers = []
    for ci in range(raw_df.shape[1]):
        parts = []
        for r in range(header_row, min(header_row + depth, len(raw_df))):
            v = raw_df.iloc[r, ci]
            if pd.notna(v):
                t = str(v).strip()
                if t and t.lower() != "nan":
                    parts.append(t)
        headers.append(slugify_text(" ".join(parts)))
    return make_unique_strings(headers)


def pick_first_existing_column(columns, aliases):
    for a in aliases:
        if a in columns:
            return a
    return None


def normalize_custom_template_df(data_df: pd.DataFrame) -> pd.DataFrame:
    df = data_df.copy()
    if "napis_na" in df.columns:
        parsed = df["napis_na"].apply(parse_napis_na).apply(pd.Series)
        for col in parsed.columns:
            df[col] = parsed[col]
    else:
        df["parsed_property_of"] = ""
        df["parsed_tool_number"] = ""
        df["parsed_part_description"] = ""

    ts = pick_first_existing_column(df.columns, ["orodje", "sifra_orodja", "orodje_sifra", "tool_number"])
    ps = pick_first_existing_column(df.columns, ["izdelek", "part_description", "izdelek_2"])
    rs = pick_first_existing_column(df.columns, ["lastnik", "lastnik_orodja", "property_of"])

    fb_prop = df[rs].fillna("").astype(str).str.strip() if rs else pd.Series([""] * len(df), index=df.index)
    fb_tool = df[ts].fillna("").astype(str).str.strip() if ts else pd.Series([""] * len(df), index=df.index)
    fb_part = df[ps].fillna("").astype(str).str.strip() if ps else pd.Series([""] * len(df), index=df.index)

    df["property_of"] = df["parsed_property_of"].fillna("").astype(str).str.strip()
    df["property_of"] = df["property_of"].mask(df["property_of"] == "", fb_prop)
    df["tool_number"] = df["parsed_tool_number"].fillna("").astype(str).str.strip()
    df["tool_number"] = df["tool_number"].mask(df["tool_number"] == "", fb_tool)
    df["part_description"] = df["parsed_part_description"].fillna("").astype(str).str.strip()
    df["part_description"] = df["part_description"].mask(df["part_description"] == "", fb_part)

    df = df[~((df["tool_number"] == "") & (df["property_of"] == "") & (df["part_description"] == ""))].copy()
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
    return normalize_custom_template_df(data_df)


# ============================================================
# Sidebar
# ============================================================
store = get_store()

with st.sidebar:
    st.header("⚙️ Label settings")
    preview_scale = st.slider("Preview scale (px/mm)", 4.0, 20.0, 8.5, 0.5)
    mode_default = st.selectbox("Default engraving mode", ALLOWED_MODES, index=1 if DEFAULT_MODE == "Anodized aluminium (negative)" else 0)
    st.divider()
    st.markdown("**Export options**")
    show_border = st.checkbox("Include border", value=True)
    show_holes = st.checkbox(
        "Include mounting holes",
        value=True,
        help="Uncheck → holes omitted from SVG/DXF; shown as dashed guides in preview.",
    )
    if not show_holes:
        st.caption("⚠️ Holes omitted from exports — guides shown in preview only")
    show_guides = st.checkbox("Show layout guides (preview only)", value=False)

    with st.expander("🔡 Text sizes"):
        fs1 = st.slider("Row 1 height (mm)", 1.2, 4.0, DEFAULT_FS1, 0.1)
        fs2 = st.slider("Row 2 height (mm)", 1.2, 4.0, DEFAULT_FS2, 0.1)
        fs3 = st.slider("Row 3 height (mm)", 1.0, 3.0, DEFAULT_FS3, 0.1)

    with st.expander("⭕ Hole dimensions"):
        hole_dia_default = st.slider("Hole diameter (mm)", 2.0, 5.0, DEFAULT_HOLE_DIA, 0.1)
        hole_offset_default = st.slider("Hole centre from edge (mm)", 3.0, 8.0, DEFAULT_HOLE_OFFSET, 0.1)

    with st.expander("🔲 Shape"):
        corner_r = st.slider("Corner radius (mm)", 0.0, 5.0, DEFAULT_CORNER_R, 0.1)
        border_offset = st.slider("Border inset (mm)", 0.0, 1.0, DEFAULT_BORDER_OFFSET, 0.05)

    with st.expander("📐 Advanced column layout"):
        st.caption(
            "Adjust these when your physical label template has different spacing. "
            "**Right text end** = Label width − Right margin. "
            "**Right column width** = Label width − Right column X − Right margin."
        )
        st.markdown("**Column positions**")
        left_x = st.slider("Left column start X (mm)", 6.0, 14.0, DEFAULT_LEFT_X, 0.1)
        left_w = st.slider("Left column width (mm)", 16.0, 28.0, DEFAULT_LEFT_W, 0.1)
        right_x = st.slider("Right column start X (mm)", 26.0, 40.0, DEFAULT_RIGHT_X, 0.1)
        right_margin = st.slider("Right margin (mm)", 0.5, 5.0, RIGHT_MARGIN, 0.1)
        computed_right_w = LABEL_W - right_x - right_margin
        st.caption(
            f"→ Right column text width: **{computed_right_w:.1f} mm**  (text runs from x={right_x:.1f} to x={LABEL_W - right_margin:.1f})"
        )

        st.markdown("**Row vertical positions**")
        row1_y = st.slider("Row 1 top (mm)", 1.0, 6.0, DEFAULT_ROW1_Y, 0.1)
        row2_y = st.slider("Row 2 top (mm)", 5.0, 11.0, DEFAULT_ROW2_Y, 0.1)
        row3_y = st.slider("Row 3 top (mm)", 10.0, 17.0, DEFAULT_ROW3_Y, 0.1)

        st.markdown("**Two-line wrap trigger**")
        wrap_trigger = st.slider(
            "Part description char trigger",
            20,
            80,
            PART_STACK_TRIGGER_LEN,
            1,
            help="Part description switches to stacked 2-line layout when text exceeds this many characters OR overflows the right column width.",
        )

    st.divider()
    st.caption(f"💾 Storage: `{STORAGE_DIR.resolve()}`")
    sc1, sc2 = st.columns(2)
    with sc1:
        st.metric("Tools stored", store.tool_count())
    with sc2:
        st.metric("Files indexed", len(store.files))


# ============================================================
# Tabs
# ============================================================
tab_single, tab_batch, tab_custom, tab_store = st.tabs(
    [
        "🏷️ Design one label",
        "📋 Generate from spreadsheet",
        "🗂️ Import tool inventory",
        "💾 Stored inventory",
    ]
)


# ------------------------------------------------------------
# Tab 1 — Single label
# ------------------------------------------------------------
with tab_single:
    lc, rc = st.columns([1.0, 1.35], gap="large")
    with lc:
        st.subheader("Label text")
        owner = st.text_input("Property of", value=DEFAULT_OWNER, key="single_owner")
        tool_number = st.text_input("Tool number", value=DEFAULT_TOOL, key="single_tool")
        part_desc = st.text_area("Part description", value=DEFAULT_PART, key="single_part", height=100)
        cc = len(part_desc)
        st.caption(f"{cc} chars — {'⚡ Stacked layout active' if cc >= wrap_trigger else f'{wrap_trigger - cc} until stacked'}")
        st.progress(min(cc / wrap_trigger, 1.0))
        st.divider()
        mode_single = st.selectbox("Engraving mode", ALLOWED_MODES, index=ALLOWED_MODES.index(mode_default), key="single_mode")
        hole_dia_single = st.number_input("Hole diameter (mm)", min_value=2.0, max_value=5.0, value=float(hole_dia_default), step=0.1)
        hole_offset_single = st.number_input("Hole centre from edge (mm)", min_value=3.0, max_value=8.0, value=float(hole_offset_default), step=0.1)

    with rc:
        st.subheader("Live preview")
        st.caption("⭕ Holes: **included**" if show_holes else "🚫 Holes: **excluded** — dashed guide only")
        svg_out, meta = generate_svg(
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
            show_holes=show_holes,
            right_margin=right_margin,
            wrap_trigger=wrap_trigger,
        )
        render_svg_preview_card(svg_out, mode_single, preview_scale)
        st.caption(f"Layout: {'stacked' if meta.get('stacked_part') else 'inline'} | Right col: {meta['right_w']:.1f} mm")
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
            show_holes=show_holes,
            right_margin=right_margin,
            wrap_trigger=wrap_trigger,
        )
        base_name = safe_filename(tool_number.strip() or "laser_label")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Download SVG", data=svg_out.encode("utf-8"), file_name=f"{base_name}.svg", mime="image/svg+xml", width="stretch")
        with c2:
            st.download_button("⬇️ Download DXF", data=dxf_bytes, file_name=f"{base_name}.dxf", mime="application/dxf", width="stretch")


# ------------------------------------------------------------
# Tab 2 — Batch from spreadsheet
# ------------------------------------------------------------
with tab_batch:
    st.subheader("Batch label generation from spreadsheet")
    with st.expander("ℹ️ Required columns", expanded=False):
        st.markdown("**Required:** `property_of`, `tool_number`, `part_description`  \n**Optional:** `mode`, `hole_dia` / `hole_size`, `hole_offset`")

    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"], key="batch_upload")
    if uploaded:
        try:
            df = normalize_columns(load_tabular_file(uploaded))
            required_cols = ["property_of", "tool_number", "part_description"]
            missing_cols = [c for c in required_cols if c not in df.columns]
            if missing_cols:
                st.error(f"Missing columns: {', '.join(missing_cols)}. Found: `{'`, `'.join(df.columns)}`")
                st.stop()

            found_opt = [c for c in ["mode", "hole_dia", "hole_size", "hole_offset"] if c in df.columns]
            st.success(f"✅ Loaded **{len(df)} rows** — required columns ✓" + (f" | Optional: `{'`, `'.join(found_opt)}`" if found_opt else ""))

            preview_df = df.copy()
            preview_df["mode"] = preview_df.get("mode", pd.Series([mode_default] * len(df))).apply(lambda x: parse_mode(x, mode_default))
            preview_df["hole_dia"] = preview_df.get("hole_dia", preview_df.get("hole_size", pd.Series([hole_dia_default] * len(df)))).apply(lambda x: parse_optional_float(x, hole_dia_default))
            preview_df["hole_offset"] = preview_df.get("hole_offset", pd.Series([hole_offset_default] * len(df))).apply(lambda x: parse_optional_float(x, hole_offset_default))

            st.markdown("### 🔍 Row preview")
            row_labels = [f"{i + 1}. {str(preview_df.iloc[i]['tool_number'])}" for i in range(len(preview_df))]
            sel_label = st.selectbox("Preview row", row_labels)
            sel_idx = row_labels.index(sel_label)
            sel_row = preview_df.iloc[sel_idx]
            sel_mode = parse_mode(sel_row.get("mode", mode_default), mode_default)
            svg_prev, _ = generate_svg(
                owner=str(sel_row["property_of"]),
                tool_number=str(sel_row["tool_number"]),
                part_desc=str(sel_row["part_description"]),
                mode=sel_mode,
                hole_dia=float(sel_row["hole_dia"]),
                hole_offset=float(sel_row["hole_offset"]),
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
                show_holes=show_holes,
                right_margin=right_margin,
                wrap_trigger=wrap_trigger,
            )
            render_svg_preview_card(svg_prev, sel_mode, preview_scale)

            st.divider()
            include_svg_b = st.checkbox("Include SVG", value=True, key="batch_svg")
            include_dxf_b = st.checkbox("Include DXF", value=True, key="batch_dxf")
            if show_holes:
                st.info("⭕ Holes included in exports.")
            else:
                st.warning("🚫 Holes excluded from exports.")

            if st.button("🚀 Generate ZIP", type="primary", width="stretch", key="batch_gen"):
                prog = st.progress(0, text="Starting…")
                try:
                    zb, mdf = build_batch_zip(
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
                        show_holes=show_holes,
                        include_svg=include_svg_b,
                        include_dxf=include_dxf_b,
                        progress_bar=prog,
                        right_margin=right_margin,
                        wrap_trigger=wrap_trigger,
                    )
                    prog.empty()
                    st.success(f"✅ {len(mdf)} labels generated.")
                    st.download_button(f"⬇️ Download ZIP ({len(mdf)} labels)", data=zb, file_name="laser_labels_batch.zip", mime="application/zip", width="stretch")
                    with st.expander("File list"):
                        st.dataframe(mdf, width="stretch", hide_index=True)
                except Exception as e:
                    prog.empty()
                    st.error(f"Generation failed: {e}")
        except Exception as e:
            st.error(f"Could not read file: {e}")


# ------------------------------------------------------------
# Tab 3 — Import tool inventory
# ------------------------------------------------------------
with tab_custom:
    st.subheader("Import tool inventory")
    st.caption(
        "Uploaded files are **saved to disk** and merged into a persistent inventory. "
        "Re-uploading a file with the same name **replaces** its rows. "
        "Duplicate uploads (same content) are silently skipped."
    )

    custom_uploaded = st.file_uploader("Upload Excel or CSV tool list", type=["csv", "xlsx", "xls"], key="custom_upload")

    if custom_uploaded is not None:
        file_bytes = custom_uploaded.read()
        file_hash = file_sha256(file_bytes)
        filename = custom_uploaded.name

        if store.file_known(file_hash):
            st.info(f"ℹ️ `{filename}` was already imported (identical content). Nothing changed.")
        else:
            try:
                buf = io.BytesIO(file_bytes)
                buf.name = filename
                if filename.lower().endswith(".csv"):
                    raw_df = normalize_columns(pd.read_csv(buf))
                    raw_df.insert(0, "excel_row", range(2, 2 + len(raw_df)))
                    parsed_df = normalize_custom_template_df(raw_df)
                else:
                    buf.seek(0)
                    parsed_df = load_custom_tool_excel(buf)

                n_new = len(parsed_df)
                is_update = store.filename_known(filename)

                if is_update:
                    st.warning(
                        f"⚠️ **`{filename}` already exists** in the store. "
                        f"Importing will **replace** all {sum(1 for t in store.tools if t.get('source_file') == filename)} rows from the old version with {n_new} new rows."
                    )
                else:
                    st.info(f"📄 Parsed **{n_new} tools** from `{filename}`. Ready to import.")

                with st.expander("Preview parsed rows", expanded=True):
                    preview_cols = ["tool_number", "property_of", "part_description"]
                    if "excel_row" in parsed_df.columns:
                        preview_cols = ["excel_row"] + preview_cols
                    st.dataframe(parsed_df[[c for c in preview_cols if c in parsed_df.columns]], width="stretch", hide_index=True)

                btn_label = f"{'🔄 Replace' if is_update else '✅ Import'} {n_new} rows into store"
                if st.button(btn_label, type="primary", width="stretch", key="confirm_import"):
                    result = store.merge_from_dataframe(parsed_df, filename, file_hash)
                    save_upload_to_disk(file_bytes, filename)
                    st.success(
                        f"✅ **Import complete** — Added: {result['added']} | Updated: {result['updated']} | Unchanged: {result['unchanged']} | Replaced old rows: {result.get('removed_old', 0)}"
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"Could not parse file: {e}")

    st.divider()
    store.reload()
    inv_df = store.as_dataframe()

    if inv_df.empty:
        st.info("No tools in store yet. Upload a file above to get started.")
    else:
        st.markdown(f"### 🔍 Browse stored tools  ({len(inv_df)} total)")
        search_term = st.text_input("🔎 Filter by tool number or owner", key="custom_search", placeholder="Type to filter…")
        if search_term.strip():
            mask = (
                inv_df["tool_number"].str.contains(search_term, case=False, na=False)
                | inv_df["property_of"].str.contains(search_term, case=False, na=False)
                | inv_df["part_description"].str.contains(search_term, case=False, na=False)
            )
            inv_df = inv_df[mask].reset_index(drop=True)
            st.caption(f"{len(inv_df)} matching rows")

        if not inv_df.empty:
            browse_labels = [f"P{int(r['priority'])} | {r['tool_number']} | {str(r.get('property_of','')).strip()} [{r.get('source_file','?')}]"for _, r in inv_df.iterrows()]
            browse_choice = st.selectbox("Select tool to preview", browse_labels, key="stored_browse")
            browse_idx = browse_labels.index(browse_choice)
            browse_row = inv_df.iloc[browse_idx]

            if not show_holes:
                st.caption("🚫 Holes excluded from export — dashed guides in preview only")

            browse_svg, _ = generate_svg(
                owner=str(browse_row.get("property_of", "")),
                tool_number=str(browse_row.get("tool_number", "")),
                part_desc=str(browse_row.get("part_description", "")),
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
                show_holes=show_holes,
                right_margin=right_margin,
                wrap_trigger=wrap_trigger,
            )
            render_svg_preview_card(browse_svg, mode_default, preview_scale)

            browse_dxf = generate_dxf(
                owner=str(browse_row.get("property_of", "")),
                tool_number=str(browse_row.get("tool_number", "")),
                part_desc=str(browse_row.get("part_description", "")),
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
                show_holes=show_holes,
                right_margin=right_margin,
                wrap_trigger=wrap_trigger,
            )
            browse_base = safe_filename(str(browse_row.get("tool_number", "")) or "laser_label")
            b1, b2 = st.columns(2)
            with b1:
                st.download_button("⬇️ Download SVG", data=browse_svg.encode("utf-8"), file_name=f"{browse_base}.svg", mime="image/svg+xml", width="stretch")
            with b2:
                st.download_button("⬇️ Download DXF", data=browse_dxf, file_name=f"{browse_base}.dxf", mime="application/dxf", width="stretch")

            st.divider()

        st.markdown("### ✅ Select rows for batch export")
        full_inv = store.as_dataframe()
        if search_term.strip():
            mask = (
                full_inv["tool_number"].str.contains(search_term, case=False, na=False)
                | full_inv["property_of"].str.contains(search_term, case=False, na=False)
                | full_inv["part_description"].str.contains(search_term, case=False, na=False)
            )
            full_inv = full_inv[mask].reset_index(drop=True)

        disp_cols = ["priority", "tool_number", "property_of", "part_description", "source_file", "updated_at"]
        disp_cols = [c for c in disp_cols if c in full_inv.columns]

        edited_df = selection_editor(full_inv, key="stored_sel_editor", display_columns=disp_cols, checkbox_label="Export")
        selected_rows = edited_df[edited_df["print"]].copy()

        ic1, ic2 = st.columns(2)
        with ic1:
            st.info(f"Showing: **{len(full_inv)}** | Selected: **{len(selected_rows)}**")
        with ic2:
            if not full_inv.empty:
                st.caption(f"Top priority: {full_inv.iloc[0]['tool_number']}")

        if selected_rows.empty:
            st.warning("No rows selected. Tick rows above to export.")
        else:
            if show_holes:
                st.info("⭕ Holes included in exports.")
            else:
                st.warning("🚫 Holes excluded from exports.")

            if st.button(f"🚀 Generate ZIP for {len(selected_rows)} rows", type="primary", width="stretch", key="stored_gen"):
                export_df = selected_rows.drop(columns=["print"]).copy()
                export_df["mode"] = mode_default
                export_df["hole_dia"] = hole_dia_default
                export_df["hole_offset"] = hole_offset_default
                prog = st.progress(0, text="Starting…")
                try:
                    zb, mdf = build_batch_zip(
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
                        show_holes=show_holes,
                        progress_bar=prog,
                        right_margin=right_margin,
                        wrap_trigger=wrap_trigger,
                    )
                    prog.empty()
                    st.success(f"✅ {len(mdf)} labels generated.")
                    st.download_button(f"⬇️ Download ZIP ({len(mdf)} labels)", data=zb, file_name="laser_labels_selected.zip", mime="application/zip", width="stretch")
                    with st.expander("Exported file list"):
                        st.dataframe(mdf, width="stretch", hide_index=True)
                except Exception as e:
                    prog.empty()
                    st.error(f"Generation failed: {e}")


# ------------------------------------------------------------
# Tab 4 — Stored inventory management
# ------------------------------------------------------------
with tab_store:
    st.subheader("💾 Stored inventory management")
    store.reload()

    col_stats1, col_stats2, col_stats3 = st.columns(3)
    with col_stats1:
        st.metric("Total tools", store.tool_count())
    with col_stats2:
        st.metric("Source files", len(store.files))
    with col_stats3:
        st.metric("Storage path", "laser_labels_data/")

    st.markdown("### 📁 Imported files")
    files_df = store.files_dataframe()
    if files_df.empty:
        st.info("No files imported yet.")
    else:
        st.dataframe(files_df, width="stretch", hide_index=True)
        st.markdown("#### Delete a file and its tools")
        st.caption("This removes all tools that came from the selected file.")
        del_options = files_df["filename"].tolist()
        del_choice = st.selectbox("Select file to delete", del_options, key="del_file_choice")
        n_affected = sum(1 for t in store.tools if t.get("source_file") == del_choice)

        if st.button(f"🗑️ Delete `{del_choice}` and its {n_affected} tools", type="secondary", key="del_file_btn"):
            store.delete_file_and_its_tools(del_choice)
            saved = UPLOADS_DIR / del_choice
            if saved.exists():
                saved.unlink()
            st.success(f"Deleted `{del_choice}` and {n_affected} associated tools.")
            st.rerun()

    st.divider()
    st.markdown("### 🔧 All stored tools")
    all_df = store.as_dataframe()
    if all_df.empty:
        st.info("No tools stored yet.")
    else:
        view_cols = ["priority", "tool_number", "property_of", "part_description", "source_file", "updated_at"]
        view_cols = [c for c in view_cols if c in all_df.columns]
        st.dataframe(all_df[view_cols], width="stretch", hide_index=True)

        st.markdown("#### Delete a single tool")
        tool_options = all_df["tool_number"].tolist()
        del_tool = st.selectbox("Select tool number", tool_options, key="del_tool_choice")
        if st.button(f"🗑️ Delete tool `{del_tool}`", type="secondary", key="del_tool_btn"):
            store.delete_tool(del_tool)
            st.success(f"Deleted tool `{del_tool}`.")
            st.rerun()

    st.divider()
    with st.expander("⚠️ Danger zone — clear everything"):
        st.warning("This deletes ALL tools and file records from the store. Uploaded files on disk are NOT deleted.")
        confirm_clear = st.text_input("Type CLEAR to confirm", key="clear_confirm")
        if st.button("🗑️ Clear entire inventory", type="secondary", key="clear_all_btn"):
            if confirm_clear == "CLEAR":
                store.clear_all()
                st.success("Inventory cleared.")
                st.rerun()
            else:
                st.error("Type CLEAR exactly to confirm.")

    st.divider()
    st.markdown("### 📥 Saved uploads on disk")
    saved_files = sorted(UPLOADS_DIR.glob("*"))
    if not saved_files:
        st.info("No files saved yet.")
    else:
        for fp in saved_files:
            c1, c2, c3 = st.columns([3, 1, 1])
            with c1:
                st.caption(f"📄 `{fp.name}`  ({fp.stat().st_size // 1024} KB)")
            with c2:
                st.download_button("⬇️ Download", data=fp.read_bytes(), file_name=fp.name, key=f"dl_{fp.name}", width="stretch")
            with c3:
                if st.button("🗑️", key=f"rm_{fp.name}", help=f"Delete {fp.name} from disk"):
                    fp.unlink()
                    st.rerun()


# ============================================================
# Footer
# ============================================================
with st.expander("📋 Column mapping reference"):
    st.code(
        """Primary: napis_na → parse:
    Property of:       → property_of
    Tool number:       → tool_number
    Part 1: / Part 2:  → part_description

Fallback columns:
    lastnik / lastnik_orodja → property_of
    orodje  / sifra_orodja   → tool_number
    izdelek / izdelek_2      → part_description""",
        language="text",
    )

with st.expander("📖 Storage & merge notes"):
    st.write(
        """
**Where data is stored:**  
`laser_labels_data/inventory.json` — merged tool list (JSON)  
`laser_labels_data/uploads/` — copies of every uploaded file

**Merge rules:**
- Same file (identical SHA-256 hash) → skipped, nothing changes
- Same filename but different content → old rows replaced, new rows imported at top priority
- New filename → rows merged in; duplicates by tool_number are updated (new file wins)
- New files always get the highest priority (appear first in the list)

**Priority:** Lower number = higher priority = shown at top. New uploads always land above existing data.
"""
    )
