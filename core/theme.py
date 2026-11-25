"""Locked Theme, Palette, Gradients, Motifs, and Glassmorphic Primitives.

Every token in this module is extracted verbatim from `templates/master_template.pptx`.
Individual part contributors MUST import from here and NEVER edit this file.
"""

from __future__ import annotations

import io
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Optional

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Paths & Canvas (exact 16:9 High-DPI canvas from master_template.pptx)
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
MASTER_TEMPLATE_PATH = REPO_ROOT / "templates" / "master_template.pptx"

CANVAS_W = Inches(20.0)
CANVAS_H = Inches(11.25)
MARGIN_X = 1.12   # inches (exact master_template.pptx horizontal margin)
MARGIN_Y = 0.99   # inches (exact master_template.pptx top title margin)

# ---------------------------------------------------------------------------
# Locked Open-Source Fonts (bundled in fonts/)
# ---------------------------------------------------------------------------
FONT_TITLE = "Space Grotesk"
FONT_BODY = "Work Sans"

# ---------------------------------------------------------------------------
# Locked Color Palette (extracted directly from master_template.pptx XML)
# ---------------------------------------------------------------------------
COLORS = {
    "OBSIDIAN": "070908",
    "DARK_TEAL": "233D40",
    "DEEP_PINE": "29464A",
    "MUTED_TEAL": "638079",
    "TABLE_HEADER": "6F8384",
    "GRAD_MID": "6D8080",
    "SAGE_GLASS": "9EB49F",
    "SOFT_SAGE": "B7C6B7",
    "GRAD_LIGHT_H": "BDCBBB",
    "GRAD_LIGHT_V": "B9C8B6",
    "ICE_MIST": "D3DDD3",
    "WHITE": "FFFFFF",
    "BLACK": "000000",
    "BORDER_GRAY": "CCCCCC",
    # Subtle diagram-only semantic highlights (never use for slide backgrounds)
    "VERIFIED_MINT": "6FD3A8",
    "ALERT_CORAL": "E07A6B",
}

# 7 Native Transparent Glassmorphic Motifs embedded inside master_template.pptx
MOTIFS = {
    "prism": "ppt/media/image1.png",        # 8-blade geometric prism star (500x500)
    "logo_wave": "ppt/media/image3.png",    # Metallic triple-ribbon emblem (500x411)
    "arc": "ppt/media/image5.png",          # Sweeping curved glass fan (544x499)
    "glow": "ppt/media/image8.png",         # Soft radial luminous bloom (1639x1345)
    "nested_arch": "ppt/media/image13.png", # Concentric arch layers (312x500)
    "crescents": "ppt/media/image18.png",   # Stacked concentric crescent orbs (379x500)
    "signal_waves": "ppt/media/image22.png",# Curved network signal bands (500x500)
}

# 4 Native OOXML Gradient Backgrounds (exact stops & angles from master_template.pptx)
BG_GRADIENTS = {
    "horizontal": {
        "type": "linear",
        "angle": 0,  # 0 deg
        "stops": [
            (0, "070908"),
            (33000, "29464A"),
            (67000, "6D8080"),
            (100000, "BDCBBB"),
        ],
    },
    "vertical": {
        "type": "linear",
        "angle": 5400000,  # 90 deg
        "stops": [
            (0, "070908"),
            (33000, "29464A"),
            (67000, "6D8080"),
            (100000, "B9C8B6"),
        ],
    },
    "diagonal": {
        "type": "linear",
        "angle": 2700000,  # 45 deg
        "stops": [
            (0, "070908"),
            (50000, "29464A"),
            (100000, "6D8080"),
        ],
    },
    "radial": {
        "type": "radial",
        "stops": [
            (0, "070908"),
            (33000, "6D8080"),
            (67000, "29464A"),
            (100000, "BDCBBB"),
        ],
    },
}


def rgb(color_key_or_hex: str) -> RGBColor:
    """Resolve a COLORS token name or 6-digit hex string to RGBColor."""
    hex_str = COLORS.get(color_key_or_hex, color_key_or_hex).lstrip("#").upper()
    return RGBColor.from_string(hex_str)


def _hex(color_key_or_hex: str) -> str:
    return COLORS.get(color_key_or_hex, color_key_or_hex).lstrip("#").upper()


# ---------------------------------------------------------------------------
# Presentation Factory
# ---------------------------------------------------------------------------
def new_presentation() -> Presentation:
    """Create a fresh 20.00\" x 11.25\" presentation matching master_template.pptx."""
    prs = Presentation()
    prs.slide_width = CANVAS_W
    prs.slide_height = CANVAS_H
    return prs


def add_blank_slide(prs: Presentation, bg: str = "diagonal"):
    """Add a blank slide and apply one of the 4 locked OOXML gradient backgrounds."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    apply_bg(slide, bg)
    return slide


def apply_bg(slide, variant: str = "diagonal") -> None:
    """Apply one of the 4 native OOXML gradient backgrounds to a slide."""
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


# ---------------------------------------------------------------------------
# Embedded Motif Loader (reads directly from templates/master_template.pptx)
# ---------------------------------------------------------------------------
@lru_cache(maxsize=16)
def _load_motif_bytes(name: str) -> bytes:
    if name not in MOTIFS:
        raise ValueError(f"Unknown motif {name!r}. Valid motifs: {list(MOTIFS)}")
    if not MASTER_TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Master template missing: {MASTER_TEMPLATE_PATH}")
    with zipfile.ZipFile(MASTER_TEMPLATE_PATH, "r") as zf:
        return zf.read(MOTIFS[name])


def add_motif(
    slide,
    name: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
    rotation: float = 0.0,
):
    """Place one of the 7 glassmorphic motifs from master_template.pptx onto a slide."""
    stream = io.BytesIO(_load_motif_bytes(name))
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
# Shape & Typography Primitives
# ---------------------------------------------------------------------------
def _set_fill_alpha(shape, hex_color: str, alpha_pct: float) -> None:
    """Apply an sRGB solid fill with OOXML alpha (0-100%) and clear default shadow."""
    spPr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.remove(old)

    alpha_val = max(0, min(100000, int(round(alpha_pct * 1000))))
    fill_xml = (
        f'<a:solidFill {nsdecls("a")}>'
        f'<a:srgbClr val="{_hex(hex_color)}">'
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
    spPr.insert(insert_eff, etree.fromstring(f'<a:effectLst {nsdecls("a")}/>'))


def _set_line_alpha(shape, hex_color: Optional[str], width_pt: float = 0.75, alpha_pct: float = 29.8) -> None:
    """Apply an sRGB outline with OOXML alpha, or noFill if hex_color is None."""
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
    radius: float = 0.04,
    name: str = "FrostedCard",
):
    """Create a signature frosted-glass rounded rectangle matching master_template.pptx."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.name = name
    try:
        shape.adjustments[0] = radius
    except Exception:
        pass
    _set_fill_alpha(shape, fill_color, fill_alpha)
    _set_line_alpha(shape, border_color, width_pt=border_pt, alpha_pct=border_alpha)
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
    border_pt: float = 0.75,
    radius: float = 0.04,
    name: str = "SolidCard",
):
    """Create a crisp opaque rounded card (e.g., light sage checklist panel or dark badge)."""
    return add_frosted_card(
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color=fill_color,
        fill_alpha=100.0,
        border_color=border_color,
        border_alpha=100.0,
        border_pt=border_pt,
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
    """Create a floating translucent glass circle (Slide 8 style)."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left),
        Inches(top),
        Inches(diameter),
        Inches(diameter),
    )
    shape.name = name
    _set_fill_alpha(shape, "WHITE", fill_alpha)
    _set_line_alpha(shape, "SAGE_GLASS", width_pt=0.75, alpha_pct=fill_alpha)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.35)
    tf.margin_right = Inches(0.35)
    tf.margin_top = Inches(0.25)
    tf.margin_bottom = Inches(0.25)
    if text:
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text
        run.font.name = FONT_BODY
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
    """Add a clean, zero-margin textbox supporting multi-line paragraphs."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    lines = text.split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = font_name
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
    """Create a subtle frosted or semantic pill badge."""
    pill = add_frosted_card(
        slide,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color=fill_color,
        fill_alpha=fill_alpha,
        border_color=border_color,
        border_alpha=border_alpha,
        border_pt=0.75,
        radius=0.45,
        name=name,
    )
    tf = pill.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT_BODY
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = rgb(text_color)
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
    """Add a clean straight OOXML connector arrow between two coordinates (in inches)."""
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
    right_label: str = "Zero Trust Architecture — University Campus Networks",
):
    """Add a quiet, consistent footer anchored safely within the 20\"x11.25\" bottom margin."""
    left_txt = part_label if slide_num is None else f"{part_label}  ·  {slide_num:02d}"
    tb_l = add_textbox(
        slide,
        left=MARGIN_X,
        top=10.55,
        width=8.5,
        height=0.38,
        text=left_txt,
        font_name=FONT_BODY,
        font_size=15.0,
        color="ICE_MIST",
        bold=False,
        align=PP_ALIGN.LEFT,
        name="FooterLeft",
    )
    tb_r = add_textbox(
        slide,
        left=10.5,
        top=10.55,
        width=20.0 - MARGIN_X - 10.5,
        height=0.38,
        text=right_label,
        font_name=FONT_BODY,
        font_size=15.0,
        color="ICE_MIST",
        bold=False,
        align=PP_ALIGN.RIGHT,
        name="FooterRight",
    )
    return tb_l, tb_r


def set_notes(slide, notes_text: str) -> None:
    """Attach clean speaker notes to a slide."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()
