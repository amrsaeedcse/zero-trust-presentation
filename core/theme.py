"""Ivory Executive Design System (12-Slot PowerPoint Palette + 3D Studio Artwork).

Single-source-of-truth Light-Mode executive architecture built specifically for
"Zero Trust Architecture for University Campus Networks":
  - Palette    : Ivory (#F7F3EC) canvas (60%), Pure White (#FFFFFF) elevated cards (20%,
                 #E7E0D3 hairline border, #000000 @ 12% shadow), Deep Navy (#14213D)
                 headings & full-bleed monoliths (15%), Bronze (#8F6B2E) & Sand (#D9C9A8)
                 accents (5%), Ink Navy (#1B2A41) body, Steel Blue (#3D5A80) links/secondary,
                 Terracotta (#B5654A) alerts, Slate (#5B6577) captions/footer/slide numbers
  - Typography : Cambria Bold (Headings & Display) + Calibri (Body & Details)
  - Native XML : Embeds the exact 12-slot `Ivory Executive` Color Scheme & Font Scheme
                 directly into `/ppt/theme/theme1.xml`
  - 3D Artwork : Seamless feathered studio sculptures (`hero_vault_navy`, `chapter_identity`,
                 `chapter_network`) + supersampled Zero-Trust shield emblem
"""

from __future__ import annotations

import io
import math
from functools import lru_cache
from pathlib import Path
from typing import Dict, Optional, Tuple

from lxml import etree
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Paths & Canvas (16:9 High-DPI 20.00" x 11.25")
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
MASTER_TEMPLATE_PATH = REPO_ROOT / "templates" / "master_template.pptx"
ART_DIR = REPO_ROOT / "templates" / "atelier_art"

CANVAS_W = Inches(20.0)
CANVAS_H = Inches(11.25)
MARGIN_X = 1.12
MARGIN_Y = 0.99

# ---------------------------------------------------------------------------
# Ivory Executive Typography (Universally pre-installed across Windows/Office)
# ---------------------------------------------------------------------------
FONT_DISPLAY = "Cambria"
FONT_TITLE = "Cambria"
FONT_BODY = "Calibri"

# Exported motif map (kept for core/__init__.py compatibility)
MOTIFS = {
    "prism": "studio:hero_vault_navy",
    "logo_wave": "procedural:zt_shield_monogram",
    "arc": "procedural:architectural_drafting_arcs",
    "glow": "procedural:ceramic_light_diffusion",
    "nested_arch": "procedural:architectural_drafting_arcs",
    "crescents": "procedural:orbital_precision_rings",
    "signal_waves": "procedural:topographic_network_waves",
}

# ---------------------------------------------------------------------------
# Ivory Executive — Complete 3-Layer Token Specification
# ---------------------------------------------------------------------------
_IVORY_EXECUTIVE_SPEC = {
    "label": "Ivory Executive — Ivory, Pure White, Deep Navy & Bronze",
    "is_light": True,
    "hero_art": "hero_vault_navy",
    "fonts": {"display": "Cambria", "title": "Cambria", "body": "Calibri"},
    "ppt_theme_name": "Ivory Executive",
    "ppt_clr_scheme": {
        "dk1": "1B2A41",      # Ink Navy (paragraph text & details)
        "lt1": "FFFFFF",      # White (cards & chip backgrounds)
        "dk2": "14213D",      # Deep Navy (titles, dark cards, numbered circles)
        "lt2": "F7F3EC",      # Ivory (consistent slide canvas)
        "accent1": "8F6B2E",  # Bronze (featured circles, big numbers, accents)
        "accent2": "3D5A80",  # Steel Blue (secondary accent & links)
        "accent3": "B5654A",  # Terracotta (alerts & callout highlights)
        "accent4": "D9C9A8",  # Sand (chip borders, timelines, dark-card headings)
        "accent5": "5B6577",  # Slate (muted captions, arrows, footer, slide nums)
        "accent6": "8A8F98",  # Gray (neutral auxiliary tone)
        "hlink": "3D5A80",    # Steel Blue
        "folHlink": "5B6577", # Slate
    },
    # Full-bleed Deep Navy monolith (#14213D) w/ Sand (#D9C9A8) headings & Ivory (#F7F3EC) body
    "monolith_stops": [(0, "182848"), (52000, "14213D"), (100000, "0E172B")],
    "monolith_bg": "14213D",
    "monolith_ink": "D9C9A8",
    "monolith_sub": "F7F3EC",
    "monolith_accent": "D9C9A8",
    "monolith_cobalt": "D9C9A8",
    "watermark_hex": "E7E0D3",
    "rule_hex": "D9C9A8",
    "recessed_fill": "EFE9DC",
    "recessed_border": "D9C9A8",
    "chart_ramp": ["D9C9A8", "8A8F98", "5B6577", "3D5A80", "8F6B2E", "14213D"],
    "colors": {
        "OBSIDIAN": "14213D",       # Deep Navy
        "DARK_TEAL": "3D5A80",      # Steel Blue
        "DEEP_PINE": "14213D",      # Deep Navy
        "MUTED_TEAL": "3D5A80",     # Steel Blue
        "TABLE_HEADER": "14213D",   # Deep Navy executive table header
        "GRAD_MID": "E7E0D3",
        "SAGE_GLASS": "D9C9A8",     # Sand chip borders & timeline lines
        "SOFT_SAGE": "5B6577",      # Slate muted captions, footer, slide numbers
        "GRAD_LIGHT_H": "F7F3EC",   # Ivory
        "GRAD_LIGHT_V": "F7F3EC",   # Ivory
        "ICE_MIST": "1B2A41",       # Ink Navy paragraph text & details
        "WHITE": "14213D",          # Deep Navy titles & headings
        "BLACK": "1B2A41",          # Ink Navy
        "BORDER_GRAY": "E7E0D3",    # Thin warm borders on White cards & table rules
        "VERIFIED_MINT": "8F6B2E",  # Bronze signature accent (big numbers, circles, badges)
        "ALERT_CORAL": "B5654A",    # Terracotta alert & warning accent
        "GOLD_ACCENT": "14213D",    # Deep Navy third-tier architectural anchor
        "PURE_WHITE": "FFFFFF",
    },
    "surfaces": {
        # Pure White (#FFFFFF) cards w/ thin #E7E0D3 border & subtle #000000 @ 12% shadow
        "WHITE": ("FFFFFF", 100.0, "E7E0D3", 100.0, True),
        "DARK_TEAL": ("FFFFFF", 100.0, "E7E0D3", 100.0, True),
        "DEEP_PINE": ("EFE9DC", 100.0, "D9C9A8", 100.0, False),
        "MUTED_TEAL": ("FFFFFF", 100.0, "D9C9A8", 100.0, True),
        "OBSIDIAN": ("FFFFFF", 100.0, "D9C9A8", 100.0, False),
        "ICE_MIST": ("FFFFFF", 100.0, "8F6B2E", 85.0, True),
        "SOFT_SAGE": ("FFFFFF", 100.0, "E7E0D3", 100.0, True),
        "VERIFIED_MINT": ("FFFFFF", 100.0, "D9C9A8", 100.0, False),
        "ALERT_CORAL": ("FAF0EC", 100.0, "B5654A", 85.0, False),
        "TABLE_HEADER": ("14213D", 100.0, "14213D", 100.0, False),
    },
    "badge_text": {
        "ALERT_CORAL": "8C3F28",
        "VERIFIED_MINT": "14213D",
        "OBSIDIAN": "14213D",
        "DARK_TEAL": "14213D",
        "WHITE": "14213D",
    },
    "table_header_text": "D9C9A8",
    "table_header_rule": "8F6B2E",
    "table_row_fill": "FFFFFF",
    "table_alt_fill": "FAF7F2",
    "table_winner_fill": "F4ECE0",
    "table_winner_text": "14213D",
    "shadow_hex": "000000",
    "shadow_alpha": 12000,
    "art_primary_rgb": (20, 33, 61),
    "art_gold_rgb": (143, 107, 46),
    "art_light_rgb": (217, 201, 168),
    "gradients": {
        "horizontal": {
            "type": "linear",
            "angle": 0,
            "stops": [(0, "F8F5EF"), (55000, "F7F3EC"), (100000, "F4EFE5")],
        },
        "vertical": {
            "type": "linear",
            "angle": 5400000,
            "stops": [(0, "F9F6F0"), (50000, "F7F3EC"), (100000, "F4EFE5")],
        },
        "diagonal": {
            "type": "linear",
            "angle": 2700000,
            "stops": [(0, "F9F6F0"), (55000, "F7F3EC"), (100000, "F3EDE2")],
        },
        "radial": {
            "type": "radial",
            "stops": [(0, "FAF7F1"), (55000, "F7F3EC"), (100000, "F3EDE2")],
        },
    },
}

THEME_PRESETS: Dict[str, dict] = {
    "ivory_executive": _IVORY_EXECUTIVE_SPEC,
}

ACTIVE_THEME_NAME = "ivory_executive"
COLORS: Dict[str, str] = {}
BG_GRADIENTS: Dict[str, dict] = {}


def get_active_theme() -> dict:
    return THEME_PRESETS[ACTIVE_THEME_NAME]


def is_light_theme() -> bool:
    return True


def resolve_font(font_name: str) -> str:
    """Resolve any heading/body font token cleanly to Cambria / Calibri."""
    fonts = get_active_theme()["fonts"]
    if font_name in (FONT_DISPLAY, FONT_TITLE, "Space Grotesk", "Cambria"):
        return fonts["display"]
    if font_name in (FONT_BODY, "Inter", "Work Sans", "Calibri"):
        return fonts["body"]
    return font_name


def set_active_theme(name: str = "ivory_executive") -> None:
    """Activate Ivory Executive theme tokens in-place."""
    global ACTIVE_THEME_NAME
    if name not in THEME_PRESETS:
        raise ValueError(f"Unknown theme {name!r}. Valid themes: {list(THEME_PRESETS)}")
    ACTIVE_THEME_NAME = name
    spec = THEME_PRESETS[name]
    COLORS.clear()
    COLORS.update(spec["colors"])
    BG_GRADIENTS.clear()
    BG_GRADIENTS.update(spec["gradients"])
    if "_generate_precision_motif" in globals():
        _generate_precision_motif.cache_clear()


set_active_theme(ACTIVE_THEME_NAME)


def hex_luminance(hex_str: str) -> float:
    """Compute WCAG relative luminance (0.0 dark .. 1.0 light) of a 6-digit hex color."""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0

    def _lin(c: float) -> float:
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def rgb(color_key_or_hex: str) -> RGBColor:
    """Resolve a COLORS token name or 6-digit hex string to RGBColor."""
    hex_str = COLORS.get(color_key_or_hex, color_key_or_hex).lstrip("#").upper()
    return RGBColor.from_string(hex_str)


def _hex(color_key_or_hex: str) -> str:
    return COLORS.get(color_key_or_hex, color_key_or_hex).lstrip("#").upper()


def _resolve_surface(
    fill_color: str,
    fill_alpha: float,
    border_color: Optional[str],
    border_alpha: float,
) -> Tuple[str, float, Optional[str], float, bool]:
    spec = get_active_theme()
    surf = spec["surfaces"].get(fill_color)
    if surf is not None:
        s_hex, s_alpha, def_b_hex, def_b_alpha, elevated = surf
        if border_color is None:
            b_hex, b_alpha = None, 0.0
        elif border_color in ("VERIFIED_MINT", "ALERT_CORAL", "DARK_TEAL", "OBSIDIAN", "GOLD_ACCENT"):
            b_hex, b_alpha = _hex(border_color), max(45.0, min(88.0, border_alpha))
        else:
            b_hex, b_alpha = def_b_hex, def_b_alpha
        return s_hex, s_alpha, b_hex, b_alpha, elevated

    return _hex(fill_color), max(fill_alpha, 96.0), (_hex(border_color) if border_color else None), max(border_alpha, 75.0), True


# ---------------------------------------------------------------------------
# Bespoke Studio 3D Sculptures + Supersampled Precision Vector Emblems
# ---------------------------------------------------------------------------
@lru_cache(maxsize=8)
def _load_studio_art_bytes(art_name: str) -> bytes:
    path = ART_DIR / f"{art_name}.png"
    if not path.exists():
        raise FileNotFoundError(f"Missing studio 3D artwork: {path}")
    return path.read_bytes()


def add_studio_art(
    slide,
    art_name: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    name: Optional[str] = None,
):
    """Place a seamlessly feathered museum-grade 3D architectural sculpture onto a slide."""
    stream = io.BytesIO(_load_studio_art_bytes(art_name))
    kwargs = {}
    if width is not None:
        kwargs["width"] = Inches(width)
    if height is not None:
        kwargs["height"] = Inches(height)
    pic = slide.shapes.add_picture(stream, Inches(left), Inches(top), **kwargs)
    pic.name = name or f"StudioArt_{art_name}"
    return pic


@lru_cache(maxsize=16)
def _generate_precision_motif(name: str, theme_name: str) -> bytes:
    """Synthesize supersampled anti-aliased vector emblems & whisper drafting arcs."""
    spec = THEME_PRESETS[theme_name]
    gd = spec["art_gold_rgb"]
    lt = spec["art_light_rgb"]

    if name == "logo_wave":
        S = 640
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        cx, cy = S // 2, S // 2
        mh = spec["monolith_bg"]
        m_rgb = (int(mh[0:2], 16), int(mh[2:4], 16), int(mh[4:6], 16))
        d.rounded_rectangle((32, 32, S - 32, S - 32), radius=148, fill=(*m_rgb, 255))
        d.rounded_rectangle((52, 52, S - 52, S - 52), radius=132, outline=(*gd, 200), width=6)
        outer = [(cx, 128), (S - 128, cy), (cx, S - 128), (128, cy)]
        d.polygon(outer, outline=(255, 255, 255, 235), width=18)
        mid = [(cx, 206), (S - 206, cy), (cx, S - 206), (206, cy)]
        d.polygon(mid, fill=(*gd, 255))
        d.ellipse((cx - 32, cy - 32, cx + 32, cy + 32), fill=(255, 255, 255, 255))
        out = img.resize((320, 320), Image.Resampling.LANCZOS)

    elif name in ("nested_arch", "arc"):
        S = 1200
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        cx, cy = int(S * 0.72), int(S * 0.56)
        for idx, rad in enumerate(range(170, 810, 80)):
            col = (*gd, 92) if idx in (2, 5) else (*lt, 120)
            w = 3 if idx in (2, 5) else 2
            d.arc((cx - rad, cy - rad, cx + rad, cy + rad), start=115, end=315, fill=col, width=w)
        d.line((cx - 480, cy, cx + 180, cy), fill=(*lt, 125), width=2)
        d.line((cx, cy - 480, cx, cy + 280), fill=(*lt, 125), width=2)
        for deg, rad in [(155, 330), (210, 570), (265, 410)]:
            a = math.radians(deg)
            px, py = cx + rad * math.cos(a), cy + rad * math.sin(a)
            d.ellipse((px - 8, py - 8, px + 8, py + 8), fill=(*gd, 185))
            d.ellipse((px - 20, py - 20, px + 20, py + 20), outline=(*gd, 85), width=2)
        out = img.resize((600, 600), Image.Resampling.LANCZOS)

    elif name == "signal_waves":
        S = 1200
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        cx, cy = int(S * 0.88), int(S * 0.12)
        for idx, rad in enumerate(range(180, 1020, 78)):
            col = (*gd, 72) if idx % 3 == 1 else (*lt, 108)
            w = 3 if idx % 3 == 1 else 2
            d.arc((cx - rad, cy - rad, cx + rad, cy + rad), start=85, end=195, fill=col, width=w)
        out = img.resize((600, 600), Image.Resampling.LANCZOS)

    elif name == "crescents":
        S = 1000
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        cx = S // 2
        for cy_off, rad, col, w in [
            (310, 180, (*gd, 115), 3),
            (460, 260, (*lt, 155), 2),
            (630, 340, (*lt, 115), 2),
        ]:
            d.ellipse((cx - rad, cy_off - rad, cx + rad, cy_off + rad), outline=col, width=w)
            d.ellipse((cx - 8, cy_off - rad - 8, cx + 8, cy_off - rad + 8), fill=(*gd, 190))
        out = img.resize((500, 500), Image.Resampling.LANCZOS)

    else:  # "glow"
        S = 800
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        cx, cy = S // 2, S // 2
        for rad, alpha in [(340, 12), (240, 18), (140, 24)]:
            d.ellipse((cx - rad, cy - rad, cx + rad, cy + rad), fill=(*gd, alpha))
        img = img.filter(ImageFilter.GaussianBlur(68))
        out = img.resize((400, 400), Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    out.save(buf, format="PNG")
    return buf.getvalue()


def add_motif(
    slide,
    name: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    rotation: float = 0.0,
):
    """Place the bespoke 3D vault (`prism`) or supersampled vector emblem/drafting motif."""
    hero_art = get_active_theme()["hero_art"]
    if name == "prism" and (ART_DIR / f"{hero_art}.png").exists():
        return add_studio_art(slide, hero_art, left=left, top=top, width=width, height=height, name="Motif_prism")
    stream = io.BytesIO(_generate_precision_motif(name, ACTIVE_THEME_NAME))
    kwargs = {}
    if width is not None:
        kwargs["width"] = Inches(width)
    if height is not None:
        kwargs["height"] = Inches(height)
    pic = slide.shapes.add_picture(stream, Inches(left), Inches(top), **kwargs)
    if rotation:
        pic.rotation = rotation
    pic.name = f"Motif_{name}"
    return pic


# ---------------------------------------------------------------------------
# Presentation Creation + Native PowerPoint Ivory Executive Theme XML Injection
# ---------------------------------------------------------------------------
def _inject_native_ppt_theme(prs: Presentation) -> None:
    """Embed Ivory Executive 12-slot Color Scheme & Cambria/Calibri Font Scheme into /ppt/theme/theme1.xml."""
    spec = get_active_theme()
    clr_map = spec["ppt_clr_scheme"]
    fonts = spec["fonts"]
    theme_name = spec["ppt_theme_name"]

    for rel in prs.slide_master.part.rels.values():
        if "theme" not in rel.reltype:
            continue
        theme_part = rel.target_part
        root = etree.fromstring(theme_part.blob)
        root.set("name", theme_name)

        te = root.find(qn("a:themeElements"))
        if te is None:
            continue

        cs = te.find(qn("a:clrScheme"))
        if cs is not None:
            cs.set("name", theme_name)
            for slot in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"):
                hex_v = clr_map.get(slot)
                if not hex_v:
                    continue
                el = cs.find(qn(f"a:{slot}"))
                if el is not None:
                    for child in list(el):
                        el.remove(child)
                    el.append(etree.fromstring(f'<a:srgbClr {nsdecls("a")} val="{hex_v.upper()}"/>'))

        fs = te.find(qn("a:fontScheme"))
        if fs is not None:
            fs.set("name", theme_name)
            maj = fs.find(qn("a:majorFont"))
            if maj is not None:
                lat = maj.find(qn("a:latin"))
                if lat is not None:
                    lat.set("typeface", fonts["display"])
            minf = fs.find(qn("a:minorFont"))
            if minf is not None:
                lat = minf.find(qn("a:latin"))
                if lat is not None:
                    lat.set("typeface", fonts["body"])

        theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        break


def apply_theme_fonts(prs: Presentation) -> None:
    """Ensure every text run across all slides uses Cambria (headings) and Calibri (body)."""

    def _remap_tf(tf):
        for p in tf.paragraphs:
            if p.font.name:
                p.font.name = resolve_font(p.font.name)
            for r in p.runs:
                if r.font.name:
                    r.font.name = resolve_font(r.font.name)

    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame:
                _remap_tf(sh.text_frame)
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        _remap_tf(cell.text_frame)
            if sh.has_chart:
                ch = sh.chart
                if ch.has_title:
                    _remap_tf(ch.chart_title.text_frame)
                try:
                    for ax in (ch.category_axis, ch.value_axis):
                        if ax.tick_labels.font.name:
                            ax.tick_labels.font.name = resolve_font(ax.tick_labels.font.name)
                except Exception:
                    pass
                try:
                    for pl in ch.plots:
                        if pl.has_data_labels and pl.data_labels.font.name:
                            pl.data_labels.font.name = resolve_font(pl.data_labels.font.name)
                except Exception:
                    pass


def new_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = CANVAS_W
    prs.slide_height = CANVAS_H
    _inject_native_ppt_theme(prs)
    return prs


def add_blank_slide(prs: Presentation, bg: str = "diagonal"):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    apply_bg(slide, bg)
    return slide


def apply_bg(slide, variant: str = "diagonal") -> None:
    if variant not in BG_GRADIENTS:
        raise ValueError(f"Unknown background variant {variant!r}. Choose from {list(BG_GRADIENTS)}")
    spec = BG_GRADIENTS[variant]
    cSld = slide._element.find(qn("p:cSld"))
    existing = cSld.find(qn("p:bg"))
    if existing is not None:
        cSld.remove(existing)

    stops_xml = "".join(
        f'<a:gs pos="{pos}"><a:srgbClr val="{col}"/></a:gs>'
        for pos, col in spec["stops"]
    )
    if spec["type"] == "linear":
        fill_mode = f'<a:lin ang="{spec["angle"]}" scaled="0"/>'
    else:
        fill_mode = (
            '<a:path path="circle">'
            '<a:fillToRect l="15000" t="15000" r="85000" b="85000"/>'
            "</a:path>"
        )

    bg_xml = (
        f'<p:bg {nsdecls("p", "a")}>'
        f"<p:bgPr>"
        f'<a:gradFill rotWithShape="1">'
        f"<a:gsLst>{stops_xml}</a:gsLst>"
        f"{fill_mode}"
        f"</a:gradFill>"
        f"<a:effectLst/>"
        f"</p:bgPr>"
        f"</p:bg>"
    )
    cSld.insert(0, etree.fromstring(bg_xml))


def add_hairline_rule(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: Optional[str] = None,
    width_pt: float = 0.75,
    name: str = "DraftingRule",
):
    """Add an ultra-crisp Sand architectural hairline rule."""
    hex_col = _hex(color) if color else get_active_theme()["rule_hex"]
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.name = name
    spPr = conn._element.spPr
    old_ln = spPr.find(qn("a:ln"))
    if old_ln is not None:
        spPr.remove(old_ln)
    w_emu = int(round(width_pt * 12700))
    ln_xml = (
        f'<a:ln {nsdecls("a")} w="{w_emu}">'
        f'<a:solidFill><a:srgbClr val="{hex_col}"/></a:solidFill>'
        f'<a:prstDash val="solid"/>'
        f"</a:ln>"
    )
    spPr.append(etree.fromstring(ln_xml))
    return conn


def add_accent_strip(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    color: str = "VERIFIED_MINT",
    rounded: bool = True,
    name: str = "AccentStrip",
):
    """Crisp solid structural rail/strip for cards & layer tracks (never used under titles)."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.name = name
    if rounded:
        try:
            shape.adjustments[0] = 0.5
        except Exception:
            pass
    _set_fill_alpha(shape, _hex(color), 100.0, elevated=False, raw_hex=True)
    _set_line_alpha(shape, None, width_pt=0)
    return shape


# ---------------------------------------------------------------------------
# Shape & Surface Primitives
# ---------------------------------------------------------------------------
def _build_effect_lst_xml(elevated: bool = False, strong: bool = False) -> str:
    spec = get_active_theme()
    if not elevated or not spec.get("shadow_hex"):
        return f'<a:effectLst {nsdecls("a")}/>'
    sh_hex = spec["shadow_hex"]
    sh_alpha = int(spec["shadow_alpha"] * (1.4 if strong else 1.0))
    blur = 254000 if strong else 177800
    dist = 63500 if strong else 44450
    return (
        f'<a:effectLst {nsdecls("a")}>'
        f'<a:outerShdw blurRad="{blur}" dist="{dist}" dir="5400000" algn="ctr" rotWithShape="0">'
        f'<a:srgbClr val="{sh_hex}"><a:alpha val="{sh_alpha}"/></a:srgbClr>'
        f"</a:outerShdw>"
        f"</a:effectLst>"
    )


def _set_fill_alpha(
    shape,
    hex_color: str,
    alpha_pct: float,
    *,
    elevated: bool = False,
    raw_hex: bool = False,
    strong_shadow: bool = False,
) -> None:
    spec = get_active_theme()
    if not raw_hex and hex_color in spec["surfaces"]:
        resolved_hex, resolved_alpha, _, _, surf_elev = spec["surfaces"][hex_color]
        elevated = elevated or surf_elev
    else:
        resolved_hex = _hex(hex_color)
        resolved_alpha = alpha_pct

    spPr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.remove(old)

    alpha_val = max(0, min(100000, int(round(resolved_alpha * 1000))))
    fill_xml = (
        f'<a:solidFill {nsdecls("a")}>'
        f'<a:srgbClr val="{resolved_hex}">'
        f'<a:alpha val="{alpha_val}"/>'
        f"</a:srgbClr>"
        f"</a:solidFill>"
    )
    geom = spPr.find(qn("a:prstGeom"))
    if geom is None:
        geom = spPr.find(qn("a:custGeom"))
    idx = list(spPr).index(geom) + 1 if geom is not None else len(spPr)
    spPr.insert(idx, etree.fromstring(fill_xml))

    eff = spPr.find(qn("a:effectLst"))
    if eff is not None:
        spPr.remove(eff)
    ln = spPr.find(qn("a:ln"))
    insert_eff = list(spPr).index(ln) + 1 if ln is not None else idx + 1
    spPr.insert(insert_eff, etree.fromstring(_build_effect_lst_xml(elevated=elevated, strong=strong_shadow)))


def _set_gradient_fill(shape, stops: Tuple[Tuple[int, str], ...], angle: int = 5400000) -> None:
    """Apply a seamless OOXML linear gradient directly to a shape (Full-Bleed Deep Navy Monolith)."""
    spPr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.remove(old)
    stops_xml = "".join(f'<a:gs pos="{pos}"><a:srgbClr val="{col}"/></a:gs>' for pos, col in stops)
    grad_xml = (
        f'<a:gradFill {nsdecls("a")} rotWithShape="1">'
        f"<a:gsLst>{stops_xml}</a:gsLst>"
        f'<a:lin ang="{angle}" scaled="0"/>'
        f"</a:gradFill>"
    )
    geom = spPr.find(qn("a:prstGeom"))
    if geom is None:
        geom = spPr.find(qn("a:custGeom"))
    idx = list(spPr).index(geom) + 1 if geom is not None else len(spPr)
    spPr.insert(idx, etree.fromstring(grad_xml))


def _set_line_alpha(shape, hex_color: Optional[str], width_pt: float = 0.75, alpha_pct: float = 29.8) -> None:
    spPr = shape._element.spPr
    old_ln = spPr.find(qn("a:ln"))
    if old_ln is not None:
        spPr.remove(old_ln)

    if hex_color is None or width_pt <= 0:
        ln_xml = f'<a:ln {nsdecls("a")}><a:noFill/></a:ln>'
    else:
        w_emu = int(round(width_pt * 12700))
        alpha_val = max(0, min(100000, int(round(alpha_pct * 1000))))
        ln_xml = (
            f'<a:ln {nsdecls("a")} w="{w_emu}" cap="sq">'
            f'<a:solidFill><a:srgbClr val="{_hex(hex_color)}">'
            f'<a:alpha val="{alpha_val}"/>'
            f"</a:srgbClr></a:solidFill>"
            f'<a:prstDash val="solid"/>'
            f"</a:ln>"
        )

    eff = spPr.find(qn("a:effectLst"))
    insert_at = list(spPr).index(eff) if eff is not None else len(spPr)
    spPr.insert(insert_at, etree.fromstring(ln_xml))


def add_monolith_panel(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    rounded: bool = False,
    radius: float = 0.035,
    name: str = "DeepNavyMonolith",
):
    """Create a velvety Deep Navy (#14213D) full-bleed hero panel for Cover & L12."""
    spec = get_active_theme()
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.name = name
    if rounded:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    _set_gradient_fill(shape, tuple(spec["monolith_stops"]), angle=5400000)
    _set_line_alpha(shape, None, width_pt=0)
    spPr = shape._element.spPr
    eff = spPr.find(qn("a:effectLst"))
    if eff is not None:
        spPr.remove(eff)
    spPr.append(etree.fromstring(_build_effect_lst_xml(elevated=rounded, strong=rounded)))
    return shape


def add_monolith_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    radius: float = 0.035,
    name: str = "ArchitecturalMonolith",
):
    return add_monolith_panel(slide, left, top, width, height, rounded=True, radius=radius, name=name)


def add_frosted_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str = "WHITE",
    fill_alpha: float = 29.8,
    border_color: Optional[str] = "SAGE_GLASS",
    border_alpha: float = 29.8,
    border_pt: float = 0.75,
    radius: float = 0.035,
    name: str = "FrostedCard",
):
    """Create a Pure White (#FFFFFF) executive card with #E7E0D3 hairline & 12% shadow."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.name = name
    try:
        shape.adjustments[0] = min(radius, 0.055) if radius < 0.25 else radius
    except Exception:
        pass

    s_hex, s_alpha, b_hex, b_alpha, elevated = _resolve_surface(fill_color, fill_alpha, border_color, border_alpha)
    eff_border_pt = max(border_pt, 0.85) if b_hex else border_pt
    _set_fill_alpha(shape, s_hex, s_alpha, elevated=elevated, raw_hex=True)
    _set_line_alpha(shape, b_hex, width_pt=eff_border_pt, alpha_pct=b_alpha)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.28)
    tf.margin_right = Inches(0.28)
    tf.margin_top = Inches(0.20)
    tf.margin_bottom = Inches(0.20)
    return shape


def add_solid_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str = "ICE_MIST",
    border_color: Optional[str] = None,
    border_pt: float = 0.85,
    radius: float = 0.035,
    name: str = "SolidCard",
):
    eff_border = border_color or "SAGE_GLASS"
    return add_frosted_card(
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color=fill_color,
        fill_alpha=100.0,
        border_color=eff_border,
        border_alpha=100.0,
        border_pt=border_pt if border_color else 0.85,
        radius=radius,
        name=name,
    )


def add_glass_circle(
    slide,
    left: float,
    top: float,
    diameter: float,
    text: str = "",
    fill_alpha: float = 29.8,
    font_size: float = 22.0,
    bold: bool = False,
    text_color: str = "WHITE",
    name: str = "GlassCircle",
):
    """Create a Pure White architectural orb surrounded by a subtle Bronze orbital ring."""
    pad = 0.14
    halo = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left - pad),
        Inches(top - pad),
        Inches(diameter + 2 * pad),
        Inches(diameter + 2 * pad),
    )
    halo.name = f"{name}_Halo"
    spPr = halo._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.remove(old)
    geom = spPr.find(qn("a:prstGeom"))
    spPr.insert(list(spPr).index(geom) + 1, etree.fromstring(f'<a:noFill {nsdecls("a")}/>'))
    _set_line_alpha(halo, COLORS["VERIFIED_MINT"], width_pt=0.9, alpha_pct=38.0)

    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left),
        Inches(top),
        Inches(diameter),
        Inches(diameter),
    )
    shape.name = name
    s_hex, s_alpha, b_hex, b_alpha, elevated = _resolve_surface("WHITE", fill_alpha, "SAGE_GLASS", fill_alpha)
    _set_fill_alpha(shape, s_hex, s_alpha, elevated=elevated, raw_hex=True)
    _set_line_alpha(shape, b_hex, width_pt=1.0, alpha_pct=b_alpha)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.32)
    tf.margin_right = Inches(0.32)
    tf.margin_top = Inches(0.24)
    tf.margin_bottom = Inches(0.24)
    if text:
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text
        run.font.name = resolve_font(FONT_TITLE)
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = rgb(text_color)
    return shape


def add_textbox(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str = FONT_BODY,
    font_size: float = 24.0,
    color: str = "WHITE",
    bold: bool = False,
    italic: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    line_spacing: Optional[float] = 1.18,
    name: str = "TextBox",
):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    eff_font = resolve_font(font_name)
    lines = text.split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = eff_font
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = rgb(color)
    return tb


def add_badge(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    fill_color: str = "WHITE",
    fill_alpha: float = 24.0,
    border_color: Optional[str] = "SAGE_GLASS",
    border_alpha: float = 40.0,
    text_color: str = "WHITE",
    font_size: float = 18.0,
    bold: bool = True,
    name: str = "Badge",
):
    """White (#FFFFFF) chip w/ Sand (#D9C9A8) border & Deep Navy (#14213D) text (or Terracotta for alerts)."""
    spec = get_active_theme()
    is_alert = fill_color == "ALERT_CORAL" or border_color == "ALERT_CORAL"
    eff_fill = "ALERT_CORAL" if is_alert else "WHITE"
    eff_border = "ALERT_CORAL" if is_alert else "SAGE_GLASS"
    eff_text_color = spec["badge_text"]["ALERT_CORAL"] if is_alert else spec["badge_text"]["VERIFIED_MINT"]

    pill = add_frosted_card(
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color=eff_fill,
        fill_alpha=100.0,
        border_color=eff_border,
        border_alpha=100.0,
        border_pt=1.0,
        radius=0.45,
        name=name,
    )
    spPr = pill._element.spPr
    eff = spPr.find(qn("a:effectLst"))
    if eff is not None:
        spPr.remove(eff)
    spPr.append(etree.fromstring(f'<a:effectLst {nsdecls("a")}/>'))

    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.14)
    tf.margin_right = Inches(0.14)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = resolve_font(FONT_TITLE)
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = rgb(eff_text_color)
    return pill


def add_arrow_connector(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: str = "WHITE",
    width_pt: float = 1.75,
    head_end: bool = False,
    tail_end: bool = True,
    dashed: bool = False,
    name: str = "Connector",
):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1),
        Inches(y1),
        Inches(x2),
        Inches(y2),
    )
    conn.name = name
    spPr = conn._element.spPr
    old_ln = spPr.find(qn("a:ln"))
    if old_ln is not None:
        spPr.remove(old_ln)

    w_emu = int(round(width_pt * 12700))
    dash = "sysDash" if dashed else "solid"
    h_xml = '<a:headEnd type="triangle" w="med" len="med"/>' if head_end else '<a:headEnd type="none"/>'
    t_xml = '<a:tailEnd type="triangle" w="med" len="med"/>' if tail_end else '<a:tailEnd type="none"/>'
    ln_xml = (
        f'<a:ln {nsdecls("a")} w="{w_emu}">'
        f'<a:solidFill><a:srgbClr val="{_hex(color)}"/></a:solidFill>'
        f'<a:prstDash val="{dash}"/>'
        f"{h_xml}{t_xml}"
        f"</a:ln>"
    )
    spPr.append(etree.fromstring(ln_xml))
    return conn


def add_footer(
    slide,
    part_label: str,
    slide_num: Optional[int] = None,
    right_label: str = "Zero Trust Architecture  ·  University Campus Networks",
    left_x: Optional[float] = None,
):
    """Minimalist Ivory Executive footer in Slate (#5B6577) anchored by a Sand (#D9C9A8) hairline."""
    x0 = left_x if left_x is not None else MARGIN_X
    x1 = 20.0 - MARGIN_X
    add_hairline_rule(slide, x1=x0, y1=10.38, x2=x1, y2=10.38, width_pt=0.65, name="FooterHairline")
    left_txt = part_label if slide_num is None else f"{part_label}  ·  {slide_num:02d}"
    mid = x0 + (x1 - x0) * 0.48
    tb_l = add_textbox(
        slide,
        left=x0,
        top=10.50,
        width=mid - x0,
        height=0.38,
        text=left_txt,
        font_name=FONT_TITLE,
        font_size=13.5,
        color="SOFT_SAGE",
        bold=True,
        align=PP_ALIGN.LEFT,
        name="FooterLeft",
    )
    tb_r = add_textbox(
        slide,
        left=mid,
        top=10.50,
        width=x1 - mid,
        height=0.38,
        text=right_label,
        font_name=FONT_BODY,
        font_size=13.5,
        color="SOFT_SAGE",
        bold=False,
        align=PP_ALIGN.RIGHT,
        name="FooterRight",
    )
    return tb_l, tb_r


def set_notes(slide, notes_text: str) -> None:
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()
