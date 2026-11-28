"""12 World-Class Ceramic Porcelain & Architectural Cobalt Layout Builders (`L01`–`L12`).

Purpose-built for "Zero Trust Architecture for University Campus Networks" using
Space Grotesk + Inter typography, museum-grade 3D architectural sculptures,
live 9-pillar progress tracking, and true full-bleed high-contrast bookends:
  - L01: True Full-Bleed Split Keynote Cover (Sunlit Porcelain + Apple Spec Bar left | Seamless Obsidian + 3D Vault Sculpture right)
  - L02: Bento Architectural 2x2 Grid (Elevated Pure Ceramic cards w/ crisp top-left numbered pills)
  - L03: Live 9-Pillar Architectural Progress Strip + Seamless 3D Studio Sculpture Chapter Opener
  - L04: Editorial Split Header + Recessed Porcelain Blueprint Stage
  - L05: Stripe/Linear Horizontal-Rule Executive Comparison Matrix
  - L06: Data Editorial Chart Studio (Cobalt-to-Obsidian chart ramp + vertical hairline insight)
  - L07: Precision 6-Node Zero Trust Pipeline (Inline numbered header pills + Cobalt connectors)
  - L08: Orbital Architectural Rings (Elevated ceramic spheres w/ cobalt orbital halo rings)
  - L09: Recessed Legacy Mist vs. Elevated Zero-Trust Ceramic Contrast
  - L10: Three Sculptural Architecture Monoliths (Crowned by Cobalt / Sapphire / Bronze top strips)
  - L11: Cascading Architectural Staircase + Seamless 3D Frosted Glass Sculpture
  - L12: Full-Bleed Obsidian Left Bookend Monolith + Crisp Cobalt-Medallion Checklist Right
"""

from __future__ import annotations

import re
from typing import List, Optional, Sequence, Tuple

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_DISPLAY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    _set_fill_alpha,
    _set_line_alpha,
    add_accent_strip,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_glass_circle,
    add_hairline_rule,
    add_monolith_panel,
    add_motif,
    add_solid_card,
    add_studio_art,
    add_textbox,
    get_active_theme,
    hex_luminance,
    is_light_theme,
    rgb,
)


def _style_card_title_body(
    card,
    heading: str,
    body: str,
    head_size: float = 25.0,
    body_size: float = 19.5,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
    text_color: str = "WHITE",
    body_color: str = "ICE_MIST",
    head_font: str = FONT_TITLE,
) -> None:
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p1 = tf.paragraphs[0]
    p1.alignment = align
    p1.space_after = Pt(10)
    r1 = p1.add_run()
    r1.text = heading
    r1.font.name = head_font
    r1.font.size = Pt(head_size)
    r1.font.bold = True
    r1.font.color.rgb = rgb(text_color)

    if body:
        p2 = tf.add_paragraph()
        p2.alignment = align
        p2.line_spacing = 1.26
        r2 = p2.add_run()
        r2.text = body
        r2.font.name = FONT_BODY
        r2.font.size = Pt(body_size)
        r2.font.bold = False
        r2.font.color.rgb = rgb(body_color)


def _add_circle_medallion(
    slide,
    left: float,
    top: float,
    diameter: float,
    text: str,
    fill_hex: str,
    text_hex: str = "FFFFFF",
    font_size: float = 15.0,
    border_hex: Optional[str] = None,
    name: str = "Medallion",
):
    """Crisp circular index or check medallion."""
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(diameter), Inches(diameter))
    circ.name = name
    _set_fill_alpha(circ, fill_hex, 100.0, elevated=False, raw_hex=True)
    _set_line_alpha(circ, border_hex, width_pt=1.0 if border_hex else 0, alpha_pct=100.0)
    tf = circ.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT_TITLE
    r.font.size = Pt(font_size)
    r.font.bold = True
    r.font.color.rgb = rgb(text_hex)
    return circ


def _add_index_chip(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    fill_hex: str,
    text_hex: str = "FFFFFF",
    font_size: float = 14.0,
    name: str = "IndexChip",
):
    """Sleek rounded architectural index pill (01..09) that never overflows or wraps."""
    chip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    chip.name = name
    try:
        chip.adjustments[0] = 0.35
    except Exception:
        pass
    _set_fill_alpha(chip, fill_hex, 100.0, elevated=False, raw_hex=True)
    _set_line_alpha(chip, None, width_pt=0)
    tf = chip.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT_TITLE
    r.font.size = Pt(font_size)
    r.font.bold = True
    r.font.color.rgb = rgb(text_hex)
    return chip


# ---------------------------------------------------------------------------
# L01 — True Full-Bleed Split Keynote Cover (Porcelain Left + Obsidian 3D Vault Right)
# ---------------------------------------------------------------------------
def L01_cover(
    prs,
    title: str,
    subtitle: str,
    top_label: str = "Computer Networks  ·  Team Research Presentation",
    badge_text: str = "Zero Trust Architecture",
    transition: str = "fade",
):
    spec = get_active_theme()
    slide = add_blank_slide(prs, bg="horizontal")

    # Right 46.5%: TRUE FULL-BLEED Deep Obsidian Panel (x=10.70" .. 20.00", y=0 .. 11.25")
    mx = 10.70
    mw = 20.0 - mx  # 9.30"
    monolith = add_monolith_panel(slide, mx, 0.0, mw, 11.25, rounded=False, name="CoverFullBleedMonolith")

    # Museum-grade 3D Vault Sculpture floating seamlessly inside the dark obsidian panel
    art_w = 6.95
    art_x = mx + (mw - art_w) / 2.0
    vault = add_studio_art(slide, spec.get("hero_art", "hero_vault"), left=art_x, top=0.58, width=art_w, name="CoverHeroVault")

    # Translucent frosted-obsidian pill anchored cleanly near bottom of right panel
    bw, bh = 5.80, 0.72
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(mx + (mw - bw) / 2.0),
        Inches(9.62),
        Inches(bw),
        Inches(bh),
    )
    badge.name = "CoverBadge"
    try:
        badge.adjustments[0] = 0.48
    except Exception:
        pass
    _set_fill_alpha(badge, "FFFFFF", 8.0, elevated=False, raw_hex=True)
    _set_line_alpha(badge, spec["monolith_accent"], width_pt=1.1, alpha_pct=65.0)
    tf_b = badge.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = badge_text.upper()
    r_b.font.name = FONT_TITLE
    r_b.font.size = Pt(15.5)
    r_b.font.bold = True
    r_b.font.color.rgb = rgb(spec["monolith_accent"])

    # Left 53.5%: Sunlit Ceramic Masthead + Monumental Space Grotesk Title + Apple Spec Bar
    logo = add_motif(slide, "logo_wave", left=MARGIN_X, top=0.88, width=0.58)
    lbl = add_textbox(
        slide,
        left=MARGIN_X + 0.76,
        top=0.98,
        width=8.55,
        height=0.42,
        text=top_label.upper(),
        font_name=FONT_TITLE,
        font_size=13.5,
        color="ICE_MIST",
        bold=True,
        name="CoverTopLabel",
    )
    top_rule = add_hairline_rule(slide, MARGIN_X, 1.72, 9.75, 1.72, width_pt=0.80, name="CoverTopRule")

    is_long = len(title) > 28 or ("\n" in title and max(len(l) for l in title.split("\n")) > 22)
    title_fs = 54.0 if is_long else 72.0
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.40 if is_long else 2.85,
        width=9.25,
        height=3.65,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=title_fs,
        color="WHITE",
        bold=True,
        line_spacing=1.06,
        name="CoverTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=6.62 if is_long else 6.18,
        width=8.65,
        height=1.60,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        line_spacing=1.36,
        name="CoverSubtitle",
    )

    # Bottom Apple-Keynote Architectural Telemetry Spec Bar
    bot_rule = add_hairline_rule(slide, MARGIN_X, 8.84, 9.75, 8.84, width_pt=0.85, name="CoverBotRule")
    spec_items = [
        ("09 PILLARS", "Identity to Data Defense", "VERIFIED_MINT"),
        ("NIST SP 800-207", "Zero Trust Architecture", "WHITE"),
        ("ZAGAZIG UNIVERSITY", "CSE421 · Computer Networks", "WHITE"),
    ]
    spec_shapes = []
    col_w = 2.82
    for idx, (k_txt, v_txt, k_col) in enumerate(spec_items):
        sx = MARGIN_X + idx * (col_w + 0.14)
        if idx > 0:
            vr = add_hairline_rule(slide, sx - 0.10, 9.08, sx - 0.10, 10.02, width_pt=0.75, name=f"SpecDiv_{idx}")
            spec_shapes.append(vr)
        tb_s = add_textbox(
            slide,
            left=sx,
            top=9.08,
            width=col_w,
            height=0.98,
            text=k_txt,
            font_name=FONT_TITLE,
            font_size=14.5,
            color=k_col,
            bold=True,
            name=f"CoverSpec_{idx+1}",
        )
        p_sv = tb_s.text_frame.add_paragraph()
        p_sv.space_before = Pt(4)
        rv = p_sv.add_run()
        rv.text = v_txt
        rv.font.name = FONT_BODY
        rv.font.size = Pt(13.5)
        rv.font.color.rgb = rgb("SOFT_SAGE")
        spec_shapes.append(tb_s)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(monolith, trigger="on_load", duration_ms=380)
    anim.zoom_in(vault, trigger="with_previous", duration_ms=520, delay_ms=60)
    anim.fade_in(logo, trigger="with_previous", duration_ms=300)
    anim.fade_in(lbl, trigger="with_previous", duration_ms=320, delay_ms=40)
    anim.wipe_in(top_rule, direction="right", trigger="with_previous", duration_ms=320, delay_ms=60)
    anim.fly_in(h1, direction="bottom", trigger="after_previous", duration_ms=420)
    anim.fade_in(sub, trigger="after_previous", duration_ms=340)
    anim.wipe_in(bot_rule, direction="right", trigger="with_previous", duration_ms=280, delay_ms=50)
    for sh in spec_shapes:
        anim.fade_in(sh, trigger="with_previous", duration_ms=280, delay_ms=80)
    anim.fade_in(badge, trigger="with_previous", duration_ms=340, delay_ms=90)
    anim.apply()
    return slide, anim


# ---------------------------------------------------------------------------
# L02 — Bento Architectural 2x2 Grid (Pure Ceramic Cards + Crisp Numbered Pills)
# ---------------------------------------------------------------------------
def L02_grid_2x2(
    prs,
    title: str,
    cards: Sequence[Tuple[str, str]],
    part_label: Optional[str] = None,
    subtitle: Optional[str] = None,
    transition: str = "fade",
    on_click_cards: bool = False,
):
    slide = add_blank_slide(prs, bg="vertical")

    title_fs = 38.0 if len(title) > 48 else (43.0 if len(title) > 38 else 48.0)
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.82,
        width=20.0 - 2 * MARGIN_X,
        height=0.96,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=title_fs,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.LEFT,
        name="SlideTitle",
    )
    sub_sh = None
    top_y = 2.48
    if subtitle:
        sub_sh = add_textbox(
            slide,
            left=MARGIN_X,
            top=1.78,
            width=16.8,
            height=0.50,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=20.0,
            color="ICE_MIST",
            align=PP_ALIGN.LEFT,
            name="SlideSubtitle",
        )
        top_y = 2.66

    rule = add_hairline_rule(slide, MARGIN_X, top_y - 0.24, 20.0 - MARGIN_X, top_y - 0.24, width_pt=0.75, name="GridRule")

    card_w = 8.60
    card_h = 3.36
    gap_x = 0.56
    gap_y = 0.50
    start_x = MARGIN_X

    created = []
    chips = []
    for idx, (ctitle, cbody) in enumerate(cards[:4]):
        r, c = divmod(idx, 2)
        x = start_x + c * (card_w + gap_x)
        y = top_y + r * (card_h + gap_y)
        card = add_frosted_card(slide, x, y, card_w, card_h, fill_color="WHITE", radius=0.04, name=f"GridCard_{idx+1}")
        card.text_frame.margin_left = Inches(0.42)
        card.text_frame.margin_right = Inches(0.42)
        card.text_frame.margin_top = Inches(0.96)
        _style_card_title_body(
            card,
            heading=ctitle,
            body=cbody,
            head_size=26.5,
            body_size=21.0,
            align=PP_ALIGN.LEFT,
            valign=MSO_ANCHOR.TOP,
            head_font=FONT_TITLE,
        )
        chip = _add_index_chip(
            slide,
            left=x + 0.42,
            top=y + 0.34,
            width=0.66,
            height=0.40,
            text=f"{idx+1:02d}",
            fill_hex=COLORS["VERIFIED_MINT"] if idx == 0 else COLORS["TABLE_HEADER"],
            text_hex="FFFFFF",
            font_size=13.5,
            name=f"GridChip_{idx+1}",
        )
        created.append(card)
        chips.append(chip)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    if sub_sh is not None:
        anim.fade_in(sub_sh, trigger="with_previous", duration_ms=320, delay_ms=50)
    anim.wipe_in(rule, direction="right", trigger="with_previous", duration_ms=300, delay_ms=70)
    for idx, (card, chip) in enumerate(zip(created, chips)):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 100
        anim.fly_in(card, direction="bottom", trigger=trig, duration_ms=340, delay_ms=dly)
        anim.fade_in(chip, trigger="with_previous", duration_ms=280, delay_ms=dly + 40)
    anim.apply()
    return slide, anim, created


# ---------------------------------------------------------------------------
# L03 — Live 9-Pillar Progress Bar + Seamless 3D Studio Sculpture Chapter Opener
# ---------------------------------------------------------------------------
def _add_pillar_progress_bar(slide, active_num: int):
    """Draw a live 9-pillar horizontal architectural progress strip across the top of Chapter Dividers."""
    shapes = []
    lbl = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.70,
        width=3.20,
        height=0.38,
        text="CAMPUS ZTA ROADMAP",
        font_name=FONT_TITLE,
        font_size=12.0,
        color="SOFT_SAGE",
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="ProgressLabel",
    )
    shapes.append(lbl)

    track_x0 = 4.55
    track_x1 = 18.88
    track_y = 0.89
    step = (track_x1 - track_x0) / 8.0

    # Background silver track line
    bg_line = add_hairline_rule(slide, track_x0, track_y, track_x1, track_y, color="BORDER_GRAY", width_pt=1.4, name="ProgTrackBg")
    shapes.append(bg_line)

    # Illuminated Cobalt progress rail up to active chapter
    if active_num > 1:
        active_x = track_x0 + (min(active_num, 9) - 1) * step
        fg_line = add_hairline_rule(slide, track_x0, track_y, active_x, track_y, color="VERIFIED_MINT", width_pt=2.2, name="ProgTrackFg")
        shapes.append(fg_line)

    # 9 architectural pillar nodes
    for i in range(1, 10):
        cx = track_x0 + (i - 1) * step
        if i == active_num:
            # Active illuminated Cobalt pill
            pw, ph = 0.74, 0.40
            node = _add_index_chip(
                slide,
                left=cx - pw / 2.0,
                top=track_y - ph / 2.0,
                width=pw,
                height=ph,
                text=f"{i:02d}",
                fill_hex=COLORS["VERIFIED_MINT"],
                text_hex="FFFFFF",
                font_size=12.5,
                name=f"ProgNode_{i:02d}",
            )
        elif i < active_num:
            # Completed pillar -> crisp Obsidian circle
            d = 0.36
            node = _add_circle_medallion(
                slide,
                left=cx - d / 2.0,
                top=track_y - d / 2.0,
                diameter=d,
                text=f"{i:02d}",
                fill_hex=COLORS["OBSIDIAN"],
                text_hex="FFFFFF",
                font_size=10.5,
                name=f"ProgNode_{i:02d}",
            )
        else:
            # Upcoming pillar -> pure white ring w/ silver numeral
            d = 0.36
            node = _add_circle_medallion(
                slide,
                left=cx - d / 2.0,
                top=track_y - d / 2.0,
                diameter=d,
                text=f"{i:02d}",
                fill_hex="FFFFFF",
                text_hex="94A3B8",
                font_size=10.5,
                border_hex="CBD5E1",
                name=f"ProgNode_{i:02d}",
            )
        shapes.append(node)
    return shapes


def L03_section_divider(
    prs,
    part_badge: str,
    title: str,
    summary: str,
    presenter: Optional[str] = None,
    motif: str = "arc",
    transition: str = "push_left",
):
    slide = add_blank_slide(prs, bg="diagonal")

    # Parse chapter number (1..9) to drive both the live Progress Bar and domain-matched 3D Sculpture!
    m_num = re.search(r"(\d{1,2})", part_badge)
    chap_int = int(m_num.group(1)) if m_num else 1

    # Live 9-Pillar Architectural Progress Strip at top
    prog_shapes = _add_pillar_progress_bar(slide, chap_int)
    top_rule = add_hairline_rule(slide, MARGIN_X, 1.38, 20.0 - MARGIN_X, 1.38, width_pt=0.75, name="ChapterTopRule")

    # Right 42%: Seamlessly feathered 3D studio sculpture tailored to chapter domain
    # Identity / Device / Policy / Adoption -> Frosted Glass Arches & Gold Prism (`chapter_identity`)
    # Threat Model / Segmentation / Remote / Monitoring / Data -> Micro-Segmented Campus Pavilions (`chapter_network`)
    art_name = "chapter_identity" if chap_int in (2, 3, 4, 8) else "chapter_network"
    art = add_studio_art(slide, art_name, left=12.15, top=1.62, width=6.75, name=f"ChapterSculpture_{chap_int:02d}")

    # Left 54%: Crisp Cobalt Chapter Pill + Monumental Space Grotesk Title + Summary + Presenter Card
    badge = add_badge(
        slide,
        left=MARGIN_X,
        top=1.96,
        width=5.40,
        height=0.54,
        text=part_badge.upper(),
        fill_color="OBSIDIAN",
        fill_alpha=100.0,
        border_color="VERIFIED_MINT",
        border_alpha=70.0,
        text_color="VERIFIED_MINT",
        font_size=14.5,
        bold=True,
        name="PartBadge",
    )
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.82,
        width=10.60,
        height=2.55,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=64.0,
        color="WHITE",
        bold=True,
        line_spacing=1.04,
        name="SectionTitle",
    )
    body = add_textbox(
        slide,
        left=MARGIN_X,
        top=5.62,
        width=10.10,
        height=2.65,
        text=summary,
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        line_spacing=1.34,
        name="SectionSummary",
    )

    pres_card = None
    pres_bar = None
    pres_tb = None
    if presenter:
        px, py, pw, ph = MARGIN_X, 8.78, 9.20, 1.24
        pres_card = add_frosted_card(slide, px, py, pw, ph, fill_color="WHITE", radius=0.06, name="PresenterCard")
        pres_bar = add_accent_strip(slide, px + 0.24, py + 0.24, 0.08, ph - 0.48, color="VERIFIED_MINT", rounded=True, name="PresenterRail")
        pres_tb = add_textbox(
            slide,
            left=px + 0.52,
            top=py + 0.20,
            width=pw - 0.76,
            height=ph - 0.36,
            text="PRESENTED BY",
            font_name=FONT_TITLE,
            font_size=12.0,
            color="SOFT_SAGE",
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
            name="PresenterLine",
        )
        p_name = pres_tb.text_frame.add_paragraph()
        p_name.space_before = Pt(3)
        rn = p_name.add_run()
        rn.text = presenter
        rn.font.name = FONT_TITLE
        rn.font.size = Pt(21.0)
        rn.font.bold = True
        rn.font.color.rgb = rgb("WHITE")

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(top_rule, trigger="on_load", duration_ms=260)
    for ps in prog_shapes:
        anim.fade_in(ps, trigger="with_previous", duration_ms=300)
    anim.zoom_in(art, trigger="with_previous", duration_ms=480, delay_ms=60)
    anim.fade_in(badge, trigger="with_previous", duration_ms=300)
    anim.fly_in(h1, direction="left", trigger="after_previous", duration_ms=380)
    anim.fade_in(body, trigger="after_previous", duration_ms=340)
    if pres_card is not None:
        anim.fly_in(pres_card, direction="bottom", trigger="with_previous", duration_ms=340, delay_ms=80)
        anim.fade_in(pres_bar, trigger="with_previous", duration_ms=260, delay_ms=100)
        anim.fade_in(pres_tb, trigger="with_previous", duration_ms=300, delay_ms=110)
    anim.apply()
    return slide, anim


# ---------------------------------------------------------------------------
# L04 — Editorial Split Header + Recessed Porcelain Blueprint Stage
# ---------------------------------------------------------------------------
def L04_wide_visual(
    prs,
    title: str,
    intro: str,
    part_label: Optional[str] = None,
    stage_card: bool = True,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="vertical")
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.82,
        width=9.6,
        height=2.15,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=46.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    desc = add_textbox(
        slide,
        left=11.12,
        top=0.92,
        width=7.76,
        height=2.0,
        text=intro,
        font_name=FONT_BODY,
        font_size=20.5,
        color="ICE_MIST",
        line_spacing=1.30,
        name="SlideIntro",
    )

    stage_bounds = (MARGIN_X, 3.36, 20.0 - 2 * MARGIN_X, 6.76)
    stage_shape = None
    if stage_card:
        sx, sy, sw, sh = stage_bounds
        # Recessed Porcelain Mist blueprint stage -> white cards placed inside leap forward!
        stage_shape = add_frosted_card(
            slide, sx, sy, sw, sh, fill_color="DEEP_PINE", radius=0.035, name="VisualStage"
        )

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=70)
    if stage_shape is not None:
        anim.fade_in(stage_shape, trigger="after_previous", duration_ms=340)
    return slide, anim, stage_bounds


# ---------------------------------------------------------------------------
# L05 — Stripe/Linear Horizontal-Rule Executive Comparison Matrix
# ---------------------------------------------------------------------------
def _style_table_cell(cell, text: str, fill_hex: str, text_hex: str, bold: bool, font_size: float, align: PP_ALIGN):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.24)
    cell.margin_right = Inches(0.24)
    cell.margin_top = Inches(0.14)
    cell.margin_bottom = Inches(0.14)

    resolved_fill = fill_hex.lstrip("#").upper()
    resolved_text = text_hex.lstrip("#").upper()
    is_header_cell = False
    if is_light_theme():
        spec = get_active_theme()
        if resolved_fill == COLORS["TABLE_HEADER"] and resolved_text == COLORS["WHITE"]:
            resolved_text = spec.get("table_header_text", "FFFFFF")
            is_header_cell = True
        elif resolved_fill == COLORS["SOFT_SAGE"]:
            resolved_fill = spec["table_row_fill"]
        elif resolved_fill == COLORS["DARK_TEAL"]:
            resolved_fill = spec["table_winner_fill"]
            if resolved_text == COLORS["VERIFIED_MINT"]:
                resolved_text = spec["table_winner_text"]

        fill_lum = hex_luminance(resolved_fill)
        text_lum = hex_luminance(resolved_text)
        if fill_lum < 0.35 and text_lum < 0.45:
            resolved_text = spec.get("table_header_text", "FFFFFF")
        elif fill_lum >= 0.60 and text_lum > 0.65:
            resolved_text = COLORS["BLACK"]

    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    tcPr.append(etree.fromstring(f'<a:solidFill {nsdecls("a")}><a:srgbClr val="{resolved_fill}"/></a:solidFill>'))

    # Clean horizontal hairlines ONLY — zero noisy vertical grid lines
    for btag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old_b = tcPr.find(qn(btag))
        if old_b is not None:
            tcPr.remove(old_b)
        if btag == "a:lnB":
            w_emu = 22225 if is_header_cell else 9525
            b_col = COLORS["VERIFIED_MINT"] if is_header_cell else COLORS["BORDER_GRAY"]
            b_xml = (
                f'<{btag} {nsdecls("a")} w="{w_emu}" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{b_col}"/></a:solidFill>'
                f'<a:prstDash val="solid"/>'
                f"</{btag}>"
            )
        else:
            b_xml = f'<{btag} {nsdecls("a")}><a:noFill/></{btag}>'
        tcPr.insert(0, etree.fromstring(b_xml))

    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = 1.18
    run = p.add_run()
    run.text = text
    run.font.name = FONT_TITLE if is_header_cell else FONT_BODY
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = rgb(resolved_text)


def L05_table_compare(
    prs,
    title: str,
    intro: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    col_widths: Optional[Sequence[float]] = None,
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="diagonal")
    add_motif(slide, "glow", left=14.6, top=-3.4, width=9.0)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.82,
        width=9.8,
        height=2.15,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=46.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    desc = add_textbox(
        slide,
        left=11.4,
        top=0.92,
        width=7.48,
        height=2.0,
        text=intro,
        font_name=FONT_BODY,
        font_size=20.5,
        color="ICE_MIST",
        line_spacing=1.30,
        name="SlideIntro",
    )

    n_rows = 1 + len(rows)
    n_cols = len(headers)
    max_row_chars = max((sum(len(c) for c in row) for row in rows), default=0)
    is_dense_tbl = max_row_chars > 165
    tbl_w = 20.0 - 2 * MARGIN_X
    tbl_h = min(6.2, 0.84 + len(rows) * (1.10 if is_dense_tbl else 1.15))
    tbl_top = 3.24 if is_dense_tbl else 3.55

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(MARGIN_X), Inches(tbl_top), Inches(tbl_w), Inches(tbl_h))
    tbl_shape.name = "ComparisonTable"
    table = tbl_shape.table

    if col_widths:
        for ci, w_in in enumerate(col_widths):
            table.columns[ci].width = Inches(w_in)
    else:
        eq = tbl_w / n_cols
        for ci in range(n_cols):
            table.columns[ci].width = Inches(eq)

    body_fs = 17.5 if is_dense_tbl else (20.0 if len(rows) <= 4 else 18.5)
    for ci, h_txt in enumerate(headers):
        cell_h = table.cell(0, ci)
        _style_table_cell(
            cell_h,
            h_txt,
            fill_hex=COLORS["TABLE_HEADER"],
            text_hex=COLORS["WHITE"],
            bold=True,
            font_size=19.0 if is_dense_tbl else 20.0,
            align=PP_ALIGN.LEFT,
        )
        if is_dense_tbl:
            cell_h.margin_top = Inches(0.10)
            cell_h.margin_bottom = Inches(0.10)
    for ri, row in enumerate(rows, start=1):
        row_bg = (
            get_active_theme()["table_alt_fill"]
            if (is_light_theme() and ri % 2 == 0)
            else COLORS["SOFT_SAGE"]
        )
        for ci, val in enumerate(row):
            cell_b = table.cell(ri, ci)
            _style_table_cell(
                cell_b,
                val,
                fill_hex=row_bg,
                text_hex=COLORS["BLACK"],
                bold=(ci == 0),
                font_size=body_fs,
                align=PP_ALIGN.LEFT,
            )
            if is_dense_tbl:
                cell_b.margin_top = Inches(0.08)
                cell_b.margin_bottom = Inches(0.08)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=70)
    anim.fly_in(tbl_shape, direction="bottom", trigger="after_previous", duration_ms=400)
    anim.apply()
    return slide, anim, tbl_shape


# ---------------------------------------------------------------------------
# L06 — Data Editorial Chart Studio (Cobalt-to-Obsidian Ramp + Hairline Insight)
# ---------------------------------------------------------------------------
def L06_chart_insight(
    prs,
    title: str,
    insight: str,
    categories: Sequence[str],
    series_name: str,
    values: Sequence[float],
    chart_title: Optional[str] = None,
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    spec = get_active_theme()
    slide = add_blank_slide(prs, bg="horizontal")
    panel = add_frosted_card(slide, MARGIN_X, 1.04, 11.45, 8.92, fill_color="WHITE", radius=0.024, name="ChartContainer")

    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series(series_name, list(values))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MARGIN_X + 0.35),
        Inches(1.38),
        Inches(10.75),
        Inches(8.24),
        chart_data,
    )
    chart_shape.name = "NativeColumnChart"
    chart = chart_shape.chart
    chart.has_legend = False
    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.text = chart_title
        for p in chart.chart_title.text_frame.paragraphs:
            p.font.name = FONT_TITLE
            p.font.size = Pt(19.5)
            p.font.bold = True
            p.font.color.rgb = rgb("WHITE")
    else:
        chart.has_title = False

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.font.name = FONT_TITLE
    plot.data_labels.font.size = Pt(15.5)
    plot.data_labels.font.bold = True
    plot.data_labels.font.color.rgb = rgb("WHITE")

    ramp = spec.get("chart_ramp", ["CBD5E1", "93C5FD", "60A5FA", "3B82F6", "1D4ED8", "0F1D40"])
    series = plot.series[0]
    for i in range(len(values)):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(ramp[i % len(ramp)])

    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.name = FONT_BODY
        axis.tick_labels.font.size = Pt(14.5)
        axis.tick_labels.font.color.rgb = rgb("ICE_MIST")
        axis.format.line.color.rgb = rgb("BORDER_GRAY")
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = rgb("BORDER_GRAY")

    v_rule = add_hairline_rule(slide, 13.00, 1.38, 13.00, 9.65, width_pt=0.85, name="ChartDivider")
    h1 = add_textbox(
        slide,
        left=13.42,
        top=1.72,
        width=5.46,
        height=2.85,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=42.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    body = add_textbox(
        slide,
        left=13.42,
        top=4.88,
        width=5.46,
        height=4.45,
        text=insight,
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        line_spacing=1.34,
        name="ChartInsight",
    )

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(panel, trigger="on_load", duration_ms=340)
    anim.wipe_in(v_rule, direction="down", trigger="with_previous", duration_ms=300, delay_ms=40)
    anim.fade_in(h1, trigger="with_previous", duration_ms=360, delay_ms=60)
    anim.wipe_in(chart_shape, direction="up", trigger="after_previous", duration_ms=460)
    anim.fade_in(body, trigger="after_previous", duration_ms=360)
    anim.apply()
    return slide, anim, chart_shape


# ---------------------------------------------------------------------------
# L07 — Precision 6-Node Zero Trust Pipeline (Inline Numbered Header Pills)
# ---------------------------------------------------------------------------
def L07_flowchart_6node(
    prs,
    title: str,
    subtitle: str,
    steps: Sequence[Tuple[str, str]],
    part_label: Optional[str] = None,
    click_reveal: bool = False,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="radial")
    title_fs = 38.0 if len(title) > 48 else (42.0 if len(title) > 38 else 46.0)
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.82,
        width=17.5,
        height=0.96,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=title_fs,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.80,
        width=16.0,
        height=0.50,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=20.0,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    box_w, box_h = 5.10, 2.86
    gap_x = 1.23
    xs = [MARGIN_X, MARGIN_X + box_w + gap_x, MARGIN_X + 2 * (box_w + gap_x)]
    y_top, y_bot = 2.78, 6.66
    coords = [
        (xs[0], y_top),
        (xs[1], y_top),
        (xs[2], y_top),
        (xs[2], y_bot),
        (xs[1], y_bot),
        (xs[0], y_bot),
    ]

    max_desc_len = max((len(s[1]) for s in steps[:6]), default=0)
    flow_head_fs = 22.0 if max_desc_len <= 80 else 21.0
    flow_body_fs = 21.0 if max_desc_len <= 80 else 19.0

    nodes = []
    headers = []
    for idx, (stitle, sdesc) in enumerate(steps[:6]):
        x, y = coords[idx]
        is_key = idx in (0, 5)
        box = add_frosted_card(
            slide,
            x,
            y,
            box_w,
            box_h,
            fill_color="ICE_MIST" if idx == 5 else "WHITE",
            border_color="VERIFIED_MINT" if is_key else "SAGE_GLASS",
            border_alpha=75.0 if is_key else 100.0,
            border_pt=1.25 if is_key else 0.9,
            radius=0.05,
            name=f"FlowNode_{idx+1}",
        )
        # Cleanly nested inline top-left chip + right title inside card (zero floating edge collisions!)
        chip = _add_index_chip(
            slide,
            left=x + 0.28,
            top=y + 0.28,
            width=0.62,
            height=0.40,
            text=f"{idx+1:02d}",
            fill_hex=COLORS["VERIFIED_MINT"] if is_key else COLORS["TABLE_HEADER"],
            text_hex="FFFFFF",
            font_size=13.5,
            name=f"FlowBadge_{idx+1}",
        )
        ttl = add_textbox(
            slide,
            left=x + 1.04,
            top=y + 0.26,
            width=box_w - 1.28,
            height=0.48,
            text=stitle,
            font_name=FONT_TITLE,
            font_size=flow_head_fs,
            color="WHITE",
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
            name=f"FlowTitle_{idx+1}",
        )
        body_tb = add_textbox(
            slide,
            left=x + 0.28,
            top=y + 0.88,
            width=box_w - 0.56,
            height=box_h - 1.06,
            text=sdesc,
            font_name=FONT_BODY,
            font_size=flow_body_fs,
            color="ICE_MIST",
            line_spacing=1.24,
            name=f"FlowDesc_{idx+1}",
        )
        nodes.append(box)
        headers.append((chip, ttl, body_tb))

    arrows = []
    cy_top = y_top + box_h / 2.0
    cy_bot = y_bot + box_h / 2.0
    if len(nodes) >= 2:
        arrows.append(add_arrow_connector(slide, xs[0] + box_w + 0.14, cy_top, xs[1] - 0.14, cy_top, color="VERIFIED_MINT", width_pt=2.2, name="Arr_1_2"))
    if len(nodes) >= 3:
        arrows.append(add_arrow_connector(slide, xs[1] + box_w + 0.14, cy_top, xs[2] - 0.14, cy_top, color="VERIFIED_MINT", width_pt=2.2, name="Arr_2_3"))
    if len(nodes) >= 4:
        arrows.append(add_arrow_connector(slide, xs[2] + box_w / 2.0, y_top + box_h + 0.14, xs[2] + box_w / 2.0, y_bot - 0.14, color="VERIFIED_MINT", width_pt=2.2, name="Arr_3_4"))
    if len(nodes) >= 5:
        arrows.append(add_arrow_connector(slide, xs[2] - 0.14, cy_bot, xs[1] + box_w + 0.14, cy_bot, color="VERIFIED_MINT", width_pt=2.2, name="Arr_4_5"))
    if len(nodes) >= 6:
        arrows.append(add_arrow_connector(slide, xs[1] - 0.14, cy_bot, xs[0] + box_w + 0.14, cy_bot, color="VERIFIED_MINT", width_pt=2.2, name="Arr_5_6"))

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=70)
    for i in range(len(nodes)):
        grp = []
        if i > 0 and i - 1 < len(arrows):
            grp.append(arrows[i - 1])
        grp.append(nodes[i])
        grp.extend(headers[i])
        anim.stagger(grp, effect="fade", step_ms=40, on_click_first=False, base_delay_ms=40)
    anim.apply()
    return slide, anim, nodes


# ---------------------------------------------------------------------------
# L08 — Orbital Architectural Rings (Elevated Ceramic Spheres + Cobalt Halos)
# ---------------------------------------------------------------------------
def L08_glass_circles(
    prs,
    title: str,
    subtitle: str,
    items: Sequence[str],
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="diagonal")
    add_motif(slide, "nested_arch", left=-1.6, top=4.1, width=7.2)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.85,
        width=6.8,
        height=2.45,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=50.0,
        color="WHITE",
        bold=True,
        line_spacing=1.06,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=5.58,
        width=6.4,
        height=2.4,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        line_spacing=1.32,
        name="SlideSubtitle",
    )

    specs = [
        (8.95, 0.95, 4.05),
        (14.15, 1.65, 3.80),
        (10.05, 5.75, 3.45),
        (14.40, 5.85, 4.00),
    ]
    circles = []
    for idx, txt in enumerate(items[:4]):
        cx, cy, diam = specs[idx]
        c = add_glass_circle(slide, cx, cy, diam, text=txt, font_size=19.5, bold=True, name=f"Circle_{idx+1}")
        circles.append(c)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.stagger(circles, effect="zoom", step_ms=120, on_click_first=False)
    anim.apply()
    return slide, anim, circles


# ---------------------------------------------------------------------------
# L09 — Recessed Legacy Mist vs. Elevated Zero-Trust Ceramic Contrast
# ---------------------------------------------------------------------------
def L09_two_columns(
    prs,
    title: str,
    left_col: Tuple[str, Sequence[str]],
    right_col: Tuple[str, Sequence[str]],
    subtitle: Optional[str] = None,
    left_badge: Optional[str] = None,
    right_badge: Optional[str] = None,
    left_accent: str = "SAGE_GLASS",
    right_accent: str = "VERIFIED_MINT",
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="horizontal")
    art = add_motif(slide, "nested_arch", left=0.55, top=5.40, width=4.65)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.06,
        width=5.4,
        height=2.35,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=44.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    sub_tb = None
    if subtitle:
        sub_tb = add_textbox(
            slide,
            left=MARGIN_X,
            top=3.48,
            width=5.2,
            height=1.8,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=19.5,
            color="ICE_MIST",
            line_spacing=1.32,
            name="SlideSubtitle",
        )

    col_w = 5.65
    col_h = 8.78
    col_y = 1.06
    col_xs = [7.15, 13.23]
    col_data = [
        (left_col, left_badge, left_accent, "DEEP_PINE", "LeftColumn"),   # Flat Recessed Cool Mist
        (right_col, right_badge, right_accent, "WHITE", "RightColumn"),   # Elevated Pure Ceramic White w/ Cobalt
    ]

    cards = []
    for idx, ((ctitle, bullets), cbadge, accent, fill_tok, cname) in enumerate(col_data):
        x = col_xs[idx]
        card = add_frosted_card(
            slide,
            x,
            col_y,
            col_w,
            col_h,
            fill_color=fill_tok,
            border_color=accent if idx == 1 else "SAGE_GLASS",
            border_alpha=82.0 if idx == 1 else 100.0,
            border_pt=1.5 if idx == 1 else 1.0,
            radius=0.025,
            name=cname,
        )
        # Crisp Cobalt crown on the elevated Zero Trust right card
        if idx == 1:
            crown_col = "VERIFIED_MINT" if accent in ("SAGE_GLASS", "SOFT_SAGE", "WHITE") else accent
            add_accent_strip(slide, x + 0.42, col_y, col_w - 0.84, 0.07, color=crown_col, rounded=False, name=f"{cname}_Crown")

        total_chars = sum(len(b) for b in bullets)
        is_dense = total_chars > 460
        is_medium = not is_dense and total_chars > 330
        head_fs = 25.0 if is_dense else (27.0 if is_medium else 29.0)
        b_fs = 17.5 if is_dense else (19.5 if is_medium else 21.5)
        b_space = 10 if is_dense else (14 if is_medium else 16)
        b_lead = 1.20 if is_dense else (1.24 if is_medium else 1.28)

        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.36 if is_dense else 0.42)
        tf.margin_right = Inches(0.36 if is_dense else 0.42)
        tf.margin_top = Inches(0.42 if not cbadge else (1.04 if is_dense else 1.18))

        if cbadge:
            add_badge(
                slide,
                left=x + (0.36 if is_dense else 0.42),
                top=col_y + (0.30 if is_dense else 0.36),
                width=3.2,
                height=0.50,
                text=cbadge,
                fill_color=accent if accent in ("VERIFIED_MINT", "ALERT_CORAL") else "WHITE",
                fill_alpha=100.0,
                border_color=accent,
                border_alpha=75.0,
                text_color="WHITE",
                font_size=14.0,
                name=f"{cname}_Badge",
            )

        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        p0.space_after = Pt(10 if is_dense else 18)
        r0 = p0.add_run()
        r0.text = ctitle
        r0.font.name = FONT_DISPLAY
        r0.font.size = Pt(head_fs)
        r0.font.bold = True
        r0.font.color.rgb = rgb("WHITE")

        for b in bullets:
            pb = tf.add_paragraph()
            pb.alignment = PP_ALIGN.LEFT
            pb.space_after = Pt(b_space)
            pb.line_spacing = b_lead
            rb = pb.add_run()
            rb.text = f"▸  {b}"
            rb.font.name = FONT_BODY
            rb.font.size = Pt(b_fs)
            rb.font.color.rgb = rgb("ICE_MIST" if idx == 0 else "WHITE")
        cards.append(card)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(art, trigger="on_load", duration_ms=400)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    if sub_tb is not None:
        anim.fade_in(sub_tb, trigger="with_previous", duration_ms=340, delay_ms=70)
    anim.stagger(cards, effect="fly_bottom", step_ms=150, on_click_first=False)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L10 — Three Sculptural Architecture Monoliths (Crowned Top Strips)
# ---------------------------------------------------------------------------
def L10_three_cards(
    prs,
    title: str,
    cards_data: Sequence[Tuple[str, str, str]],
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="vertical")
    title_fs = 36.0 if len(title) > 50 else (42.0 if len(title) > 40 else 48.0)
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.82,
        width=20.0 - 2 * MARGIN_X,
        height=0.98,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=title_fs,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.LEFT,
        name="SlideTitle",
    )
    rule = add_hairline_rule(slide, MARGIN_X, 1.98, 20.0 - MARGIN_X, 1.98, width_pt=0.75, name="TriRule")

    cw = 5.66
    ch = 7.72
    gap = 0.39
    start_x = MARGIN_X
    top_y = 2.28
    crowns = ["VERIFIED_MINT", "DARK_TEAL", "GOLD_ACCENT"]

    max_tri_body = max((len(c[2]) for c in cards_data[:3]), default=0)
    tri_tag_fs = 15.0 if max_tri_body <= 140 else 14.5
    tri_head_fs = 31.0 if max_tri_body <= 140 else (28.5 if max_tri_body <= 265 else 27.5)
    tri_body_fs = 23.0 if max_tri_body <= 140 else (21.0 if max_tri_body <= 265 else 19.5)

    cards = []
    extras = []
    for idx, (tag, heading, body) in enumerate(cards_data[:3]):
        x = start_x + idx * (cw + gap)
        acc = crowns[idx % len(crowns)]
        card = add_frosted_card(slide, x, top_y, cw, ch, fill_color="WHITE", radius=0.04, name=f"TriCard_{idx+1}")
        crown = add_accent_strip(slide, x + 0.38, top_y, cw - 0.76, 0.07, color=acc, rounded=False, name=f"TriCrown_{idx+1}")
        chip = _add_index_chip(
            slide,
            left=x + 0.42,
            top=top_y + 0.42,
            width=0.66,
            height=0.40,
            text=f"{idx+1:02d}",
            fill_hex=COLORS[acc],
            text_hex="FFFFFF",
            font_size=13.5,
            name=f"TriNum_{idx+1}",
        )

        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.42)
        tf.margin_right = Inches(0.42)
        tf.margin_top = Inches(1.06)

        p_tag = tf.paragraphs[0]
        p_tag.alignment = PP_ALIGN.LEFT
        p_tag.space_after = Pt(12)
        r_tag = p_tag.add_run()
        r_tag.text = tag.upper()
        r_tag.font.name = FONT_TITLE
        r_tag.font.size = Pt(tri_tag_fs)
        r_tag.font.bold = True
        r_tag.font.color.rgb = rgb(acc)

        p_h = tf.add_paragraph()
        p_h.alignment = PP_ALIGN.LEFT
        p_h.space_after = Pt(16)
        p_h.line_spacing = 1.12
        r_h = p_h.add_run()
        r_h.text = heading
        r_h.font.name = FONT_DISPLAY
        r_h.font.size = Pt(tri_head_fs)
        r_h.font.bold = True
        r_h.font.color.rgb = rgb("WHITE")

        p_b = tf.add_paragraph()
        p_b.alignment = PP_ALIGN.LEFT
        p_b.line_spacing = 1.32
        r_b = p_b.add_run()
        r_b.text = body
        r_b.font.name = FONT_BODY
        r_b.font.size = Pt(tri_body_fs)
        r_b.font.color.rgb = rgb("ICE_MIST")

        cards.append(card)
        extras.append((crown, chip))

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.wipe_in(rule, direction="right", trigger="with_previous", duration_ms=280, delay_ms=50)
    for idx, (card, (crown, chip)) in enumerate(zip(cards, extras)):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 120
        anim.fly_in(card, direction="bottom", trigger=trig, duration_ms=350, delay_ms=dly)
        anim.fade_in(crown, trigger="with_previous", duration_ms=260, delay_ms=dly + 40)
        anim.fade_in(chip, trigger="with_previous", duration_ms=260, delay_ms=dly + 50)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L11 — Cascading Architectural Staircase + Seamless 3D Frosted Glass Sculpture
# ---------------------------------------------------------------------------
def L11_stepped_stack(
    prs,
    title: str,
    rows: Sequence[Tuple[str, str]],
    subtitle: Optional[str] = None,
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="diagonal")
    # Floating 3D frosted glass sculpture anchoring lower-left editorial column
    art = add_studio_art(slide, "chapter_identity", left=1.45, top=4.78, width=3.95, name="SteppedGlassSculpture")

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.02,
        width=6.2,
        height=2.05,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=48.0,
        color="WHITE",
        bold=True,
        line_spacing=1.06,
        name="SlideTitle",
    )
    sub_tb = None
    if subtitle:
        sub_tb = add_textbox(
            slide,
            left=MARGIN_X,
            top=3.12,
            width=5.8,
            height=1.65,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=20.0,
            color="ICE_MIST",
            line_spacing=1.30,
            name="SlideSubtitle",
        )

    tier_fills = ["WHITE", "WHITE", "MUTED_TEAL", "ICE_MIST"]
    row_base_x = 7.75
    row_h = 1.90
    gap_y = 0.38
    start_y = 1.12

    cards = []
    meds = []
    for idx, (rtitle, rdesc) in enumerate(rows[:4]):
        indent = idx * 0.28
        rx = row_base_x + indent
        rw = (20.0 - MARGIN_X) - rx
        y = start_y + idx * (row_h + gap_y)
        card = add_frosted_card(
            slide,
            rx,
            y,
            rw,
            row_h,
            fill_color=tier_fills[idx % len(tier_fills)],
            border_color="VERIFIED_MINT" if idx == 3 else "SAGE_GLASS",
            border_alpha=85.0 if idx == 3 else 100.0,
            border_pt=1.4 if idx == 3 else 0.9,
            radius=0.045,
            name=f"SteppedRow_{idx+1}",
        )
        card.text_frame.margin_left = Inches(1.12)
        card.text_frame.margin_right = Inches(0.34)
        _style_card_title_body(
            card,
            heading=rtitle,
            body=rdesc,
            head_size=24.5,
            body_size=20.5,
            align=PP_ALIGN.LEFT,
            valign=MSO_ANCHOR.MIDDLE,
            text_color="WHITE",
            body_color="ICE_MIST",
            head_font=FONT_TITLE,
        )
        med = _add_index_chip(
            slide,
            left=rx + 0.28,
            top=y + (row_h - 0.50) / 2.0,
            width=0.62,
            height=0.50,
            text=f"{idx+1:02d}",
            fill_hex=COLORS["VERIFIED_MINT"] if idx == 3 else COLORS["TABLE_HEADER"],
            text_hex="FFFFFF",
            font_size=14.5,
            name=f"SteppedBadge_{idx+1}",
        )
        cards.append(card)
        meds.append(med)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.zoom_in(art, trigger="on_load", duration_ms=440)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    if sub_tb is not None:
        anim.fade_in(sub_tb, trigger="with_previous", duration_ms=340, delay_ms=70)
    for idx, (card, med) in enumerate(zip(cards, meds)):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 110
        anim.fly_in(card, direction="right", trigger=trig, duration_ms=340, delay_ms=dly)
        anim.fade_in(med, trigger="with_previous", duration_ms=280, delay_ms=dly + 40)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L12 — Full-Bleed Obsidian Bookend Monolith + Crisp Swiss Checklist (Slide 59)
# ---------------------------------------------------------------------------
def L12_split_checklist(
    prs,
    title: str,
    subtitle: str,
    checklist_title: str,
    items: Sequence[str],
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    spec = get_active_theme()
    slide = add_blank_slide(prs, bg="horizontal")

    # Left 43%: TRUE FULL-BLEED Obsidian Monolith echoing Slide 1's right cover panel!
    lw = 8.60
    left_panel = add_monolith_panel(slide, 0.0, 0.0, lw, 11.25, rounded=False, name="LeftFullBleedMonolith")

    eyebrow = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.22,
        width=6.6,
        height=0.40,
        text="ARCHITECTURAL SYNTHESIS",
        font_name=FONT_TITLE,
        font_size=13.5,
        color=spec["monolith_accent"],
        bold=True,
        name="SummaryEyebrow",
    )
    long_l12 = len(title) > 28
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.78,
        width=6.65,
        height=2.45,
        text=title,
        font_name=FONT_DISPLAY,
        font_size=40.0 if long_l12 else 48.0,
        color=spec["monolith_ink"],
        bold=True,
        line_spacing=1.06,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=4.52 if long_l12 else 4.35,
        width=6.55,
        height=2.15,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=19.5 if long_l12 else 20.0,
        color=spec["monolith_sub"],
        line_spacing=1.32,
        name="SlideSubtitle",
    )
    # Bookend 3D vault sculpture subtly anchoring bottom of Obsidian monolith
    vault_mini = add_studio_art(slide, spec.get("hero_art", "hero_vault"), left=2.48, top=6.42, width=3.64, name="BookendHeroVault")

    # Right 57%: Sunlit Porcelain Checklist Stage (clean whitespace, zero unnecessary outer box)
    right_x = 9.45
    right_w = (20.0 - MARGIN_X) - right_x  # 9.43"
    r_head = add_textbox(
        slide,
        left=right_x,
        top=1.18,
        width=right_w,
        height=0.66,
        text=checklist_title,
        font_name=FONT_DISPLAY,
        font_size=32.0,
        color="WHITE",
        bold=True,
        name="ChecklistHeader",
    )
    hdr_rule = add_hairline_rule(slide, right_x, 2.02, right_x + right_w, 2.02, width_pt=0.90, name="ChecklistRule")

    item_cards = []
    meds = []
    n = min(len(items), 5)
    row_h = 1.54
    start_y = 2.28
    for idx, txt in enumerate(items[:n]):
        y = start_y + idx * row_h
        med = _add_circle_medallion(
            slide,
            left=right_x,
            top=y + 0.42,
            diameter=0.48,
            text="✓",
            fill_hex=COLORS["VERIFIED_MINT"],
            text_hex="FFFFFF",
            font_size=15.0,
            name=f"CheckIcon_{idx+1}",
        )
        tb = add_textbox(
            slide,
            left=right_x + 0.74,
            top=y + 0.14,
            width=right_w - 0.74,
            height=1.12,
            text=txt,
            font_name=FONT_BODY,
            font_size=20.0,
            color="WHITE",
            valign=MSO_ANCHOR.MIDDLE,
            line_spacing=1.24,
            name=f"ChecklistItem_{idx+1}",
        )
        if idx < n - 1:
            add_hairline_rule(
                slide,
                right_x,
                y + row_h - 0.06,
                right_x + right_w,
                y + row_h - 0.06,
                width_pt=0.60,
                name=f"CheckSep_{idx+1}",
            )
        item_cards.append(tb)
        meds.append(med)

    if part_label:
        # Keep footer cleanly inside the right sunlit porcelain area!
        add_footer(slide, part_label, left_x=right_x)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(left_panel, trigger="on_load", duration_ms=360)
    anim.fade_in(eyebrow, trigger="with_previous", duration_ms=300, delay_ms=40)
    anim.fly_in(h1, direction="left", trigger="with_previous", duration_ms=380, delay_ms=60)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=110)
    anim.zoom_in(vault_mini, trigger="with_previous", duration_ms=440, delay_ms=120)
    anim.fade_in(r_head, trigger="after_previous", duration_ms=300)
    anim.wipe_in(hdr_rule, direction="right", trigger="with_previous", duration_ms=280, delay_ms=40)
    for idx, (med, tb) in enumerate(zip(meds, item_cards)):
        trig = "after_previous" if idx == 0 else "with_previous"
        dly = 0 if idx == 0 else idx * 90
        anim.zoom_in(med, trigger=trig, duration_ms=260, delay_ms=dly)
        anim.fade_in(tb, trigger="with_previous", duration_ms=300, delay_ms=dly + 30)
    anim.apply()
    return slide, anim, item_cards
