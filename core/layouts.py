"""12 Template-Matched Layout Builders (`L01`–`L12`) for master_template.pptx.

Each builder mirrors a slide from `templates/master_template.pptx` with locked
colors, gradients, frosted-glass styling, fonts (`Space Grotesk` + `Work Sans`),
native motifs, and default OOXML animations.

These 12 layouts are OPTIONAL building blocks — teammates may also compose custom
on-theme diagrams/layouts inside their own `parts/part_0N_*.py` using `core.theme`
and `core.animations`.
"""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_glass_circle,
    add_motif,
    add_solid_card,
    add_textbox,
    rgb,
)


def _style_card_title_body(
    card,
    heading: str,
    body: str,
    head_size: float = 28.0,
    body_size: float = 22.0,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
    text_color: str = "WHITE",
    body_color: str = "ICE_MIST",
) -> None:
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p1 = tf.paragraphs[0]
    p1.alignment = align
    p1.space_after = Pt(10)
    r1 = p1.add_run()
    r1.text = heading
    r1.font.name = FONT_BODY
    r1.font.size = Pt(head_size)
    r1.font.bold = True
    r1.font.color.rgb = rgb(text_color)

    if body:
        p2 = tf.add_paragraph()
        p2.alignment = align
        p2.line_spacing = 1.2
        r2 = p2.add_run()
        r2.text = body
        r2.font.name = FONT_BODY
        r2.font.size = Pt(body_size)
        r2.font.bold = False
        r2.font.color.rgb = rgb(body_color)


# ---------------------------------------------------------------------------
# L01 — Cover / Hero (Template Slide 01)
# ---------------------------------------------------------------------------
def L01_cover(
    prs,
    title: str,
    subtitle: str,
    top_label: str = "Computer Networks  ·  Team Research Presentation",
    badge_text: str = "Zero Trust Architecture",
    transition: str = "fade",
):
    slide = add_blank_slide(prs, bg="horizontal")
    prism = add_motif(slide, "prism", left=13.7, top=-0.9, width=7.0)
    logo = add_motif(slide, "logo_wave", left=MARGIN_X, top=1.08, width=0.95)

    lbl = add_textbox(
        slide,
        left=2.30,
        top=1.20,
        width=10.5,
        height=0.58,
        text=top_label,
        font_name=FONT_BODY,
        font_size=24.0,
        color="ICE_MIST",
        name="CoverTopLabel",
    )
    title_fs = 72.0 if (len(title) > 28 or "\n" in title) else 82.0
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=4.05 if title_fs < 80 else 4.35,
        width=13.5,
        height=4.0,
        text=title,
        font_name=FONT_TITLE,
        font_size=title_fs,
        color="WHITE",
        bold=True,
        line_spacing=1.06,
        name="CoverTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=8.65,
        width=12.4,
        height=1.1,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=28.0,
        color="ICE_MIST",
        name="CoverSubtitle",
    )
    badge = add_badge(
        slide,
        left=13.95,
        top=8.95,
        width=4.93,
        height=1.02,
        text=badge_text,
        fill_alpha=29.8,
        font_size=24.0,
        bold=False,
        name="CoverBadge",
    )

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(prism, trigger="on_load", duration_ms=550)
    anim.fade_in(logo, trigger="with_previous", duration_ms=400)
    anim.fade_in(lbl, trigger="with_previous", duration_ms=400)
    anim.fly_in(h1, direction="bottom", trigger="after_previous", duration_ms=480)
    anim.fade_in(sub, trigger="after_previous", duration_ms=380)
    anim.fade_in(badge, trigger="with_previous", duration_ms=400, delay_ms=120)
    anim.apply()
    return slide, anim


# ---------------------------------------------------------------------------
# L02 — Centered Title + 2x2 Frosted Glass Grid (Template Slide 02)
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
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.05,
        width=20.0 - 2 * MARGIN_X,
        height=1.05,
        text=title,
        font_name=FONT_TITLE,
        font_size=58.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="SlideTitle",
    )
    sub_sh = None
    top_y = 3.10
    if subtitle:
        sub_sh = add_textbox(
            slide,
            left=2.5,
            top=2.18,
            width=15.0,
            height=0.58,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=22.0,
            color="ICE_MIST",
            align=PP_ALIGN.CENTER,
            name="SlideSubtitle",
        )
        top_y = 3.25

    card_w = 8.40
    card_h = 2.95
    gap_x = 0.96
    gap_y = 0.65
    start_x = (20.0 - (2 * card_w + gap_x)) / 2.0

    created = []
    for idx, (ctitle, cbody) in enumerate(cards[:4]):
        r, c = divmod(idx, 2)
        x = start_x + c * (card_w + gap_x)
        y = top_y + r * (card_h + gap_y)
        card = add_frosted_card(slide, x, y, card_w, card_h, name=f"GridCard_{idx+1}")
        _style_card_title_body(
            card,
            heading=ctitle,
            body=cbody,
            head_size=28.0,
            body_size=21.0,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        created.append(card)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=350)
    if sub_sh is not None:
        anim.fade_in(sub_sh, trigger="with_previous", duration_ms=350, delay_ms=80)
    anim.stagger(created, effect="fly_bottom", step_ms=130, on_click_first=on_click_cards)
    anim.apply()
    return slide, anim, created


# ---------------------------------------------------------------------------
# L03 — Bold Statement / Section Divider + Arc Motif (Template Slide 03)
# ---------------------------------------------------------------------------
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

    if motif == "arc":
        art = add_motif(slide, "arc", left=13.2, top=4.75, width=7.2)
    elif motif == "signal_waves":
        art = add_motif(slide, "signal_waves", left=13.8, top=4.9, width=6.2)
    elif motif == "nested_arch":
        art = add_motif(slide, "nested_arch", left=14.6, top=3.2, width=4.8)
    elif motif == "crescents":
        art = add_motif(slide, "crescents", left=14.3, top=3.8, width=5.2)
    else:
        art = add_motif(slide, motif, left=13.5, top=4.2, width=6.0)

    badge = add_badge(
        slide,
        left=MARGIN_X,
        top=1.35,
        width=6.2,
        height=0.68,
        text=part_badge,
        fill_alpha=24.0,
        font_size=19.0,
        bold=True,
        name="PartBadge",
    )
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.45,
        width=12.8,
        height=2.4,
        text=title,
        font_name=FONT_TITLE,
        font_size=64.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SectionTitle",
    )
    body = add_textbox(
        slide,
        left=MARGIN_X,
        top=5.35,
        width=11.4,
        height=3.2,
        text=summary,
        font_name=FONT_BODY,
        font_size=26.0,
        color="ICE_MIST",
        line_spacing=1.28,
        name="SectionSummary",
    )
    pres_tb = None
    if presenter:
        pres_tb = add_textbox(
            slide,
            left=MARGIN_X,
            top=9.35,
            width=11.0,
            height=0.55,
            text=f"Presented by: {presenter}",
            font_name=FONT_BODY,
            font_size=21.0,
            color="SOFT_SAGE",
            bold=True,
            name="PresenterLine",
        )

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(art, trigger="on_load", duration_ms=520)
    anim.fade_in(badge, trigger="with_previous", duration_ms=340)
    anim.fly_in(h1, direction="left", trigger="after_previous", duration_ms=420)
    anim.fade_in(body, trigger="after_previous", duration_ms=380)
    if pres_tb is not None:
        anim.fade_in(pres_tb, trigger="with_previous", duration_ms=380, delay_ms=120)
    anim.apply()
    return slide, anim


# ---------------------------------------------------------------------------
# L04 — Split Top Header + Wide Visual / Diagram Stage (Template Slide 04)
# ---------------------------------------------------------------------------
def L04_wide_visual(
    prs,
    title: str,
    intro: str,
    part_label: Optional[str] = None,
    stage_card: bool = True,
    transition: str = "fade",
):
    """Returns `(slide, anim, stage_bounds)` where stage_bounds = (x, y, w, h) in inches.

    Caller populates custom Zero Trust architecture diagrams inside `stage_bounds`
    and calls `anim.apply()` after queueing diagram animations.
    """
    slide = add_blank_slide(prs, bg="vertical")
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=9.6,
        height=2.2,
        text=title,
        font_name=FONT_TITLE,
        font_size=54.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    desc = add_textbox(
        slide,
        left=11.12,
        top=1.08,
        width=7.76,
        height=2.0,
        text=intro,
        font_name=FONT_BODY,
        font_size=23.0,
        color="ICE_MIST",
        line_spacing=1.24,
        name="SlideIntro",
    )

    stage_bounds = (MARGIN_X, 3.45, 20.0 - 2 * MARGIN_X, 6.65)
    stage_shape = None
    if stage_card:
        sx, sy, sw, sh = stage_bounds
        stage_shape = add_frosted_card(
            slide, sx, sy, sw, sh, fill_alpha=18.0, border_alpha=32.0, radius=0.03, name="VisualStage"
        )

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=80)
    if stage_shape is not None:
        anim.fade_in(stage_shape, trigger="after_previous", duration_ms=350)
    return slide, anim, stage_bounds


# ---------------------------------------------------------------------------
# L05 — Split Top Header + Sage Comparison Table + Glow (Template Slide 05)
# ---------------------------------------------------------------------------
def _style_table_cell(cell, text: str, fill_hex: str, text_hex: str, bold: bool, font_size: float, align: PP_ALIGN):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.22)
    cell.margin_right = Inches(0.22)
    cell.margin_top = Inches(0.14)
    cell.margin_bottom = Inches(0.14)

    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    tcPr.append(etree.fromstring(f'<a:solidFill {nsdecls("a")}><a:srgbClr val="{fill_hex}"/></a:solidFill>'))

    for btag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old_b = tcPr.find(qn(btag))
        if old_b is not None:
            tcPr.remove(old_b)
        b_xml = (
            f'<{btag} {nsdecls("a")} w="9525" cap="flat" cmpd="sng" algn="ctr">'
            f'<a:solidFill><a:srgbClr val="{COLORS["BORDER_GRAY"]}"/></a:solidFill>'
            f'<a:prstDash val="solid"/>'
            f"</{btag}>"
        )
        tcPr.insert(0, etree.fromstring(b_xml))

    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = 1.15
    run = p.add_run()
    run.text = text
    run.font.name = FONT_BODY
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = rgb(text_hex)


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
        top=MARGIN_Y,
        width=9.8,
        height=2.2,
        text=title,
        font_name=FONT_TITLE,
        font_size=54.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    desc = add_textbox(
        slide,
        left=11.4,
        top=1.08,
        width=7.48,
        height=2.0,
        text=intro,
        font_name=FONT_BODY,
        font_size=23.0,
        color="ICE_MIST",
        line_spacing=1.24,
        name="SlideIntro",
    )

    n_rows = 1 + len(rows)
    n_cols = len(headers)
    tbl_w = 20.0 - 2 * MARGIN_X
    tbl_h = min(6.1, 0.95 + len(rows) * 1.15)
    tbl_top = 3.75

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

    body_fs = 21.0 if len(rows) <= 4 else 19.0
    for ci, h_txt in enumerate(headers):
        _style_table_cell(
            table.cell(0, ci),
            h_txt,
            fill_hex=COLORS["TABLE_HEADER"],
            text_hex=COLORS["WHITE"],
            bold=True,
            font_size=22.0,
            align=PP_ALIGN.CENTER,
        )
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            _style_table_cell(
                table.cell(ri, ci),
                val,
                fill_hex=COLORS["SOFT_SAGE"],
                text_hex=COLORS["BLACK"],
                bold=(ci == 0),
                font_size=body_fs,
                align=PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT,
            )

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(desc, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fly_in(tbl_shape, direction="bottom", trigger="after_previous", duration_ms=450)
    anim.apply()
    return slide, anim, tbl_shape


# ---------------------------------------------------------------------------
# L06 — Native Editable Chart Left (65%) + Key Takeaways Right (Template Slide 06)
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
    slide = add_blank_slide(prs, bg="horizontal")
    panel = add_frosted_card(slide, MARGIN_X, 1.15, 11.45, 8.85, fill_alpha=24.0, name="ChartContainer")

    chart_data = CategoryChartData()
    chart_data.categories = list(categories)
    chart_data.add_series(series_name, list(values))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MARGIN_X + 0.35),
        Inches(1.45),
        Inches(10.75),
        Inches(8.25),
        chart_data,
    )
    chart_shape.name = "NativeColumnChart"
    chart = chart_shape.chart
    chart.has_legend = False
    if chart_title:
        chart.has_title = True
        chart.chart_title.text_frame.text = chart_title
        for p in chart.chart_title.text_frame.paragraphs:
            p.font.name = FONT_BODY
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = rgb("WHITE")
    else:
        chart.has_title = False

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.font.name = FONT_BODY
    plot.data_labels.font.size = Pt(16)
    plot.data_labels.font.bold = True
    plot.data_labels.font.color.rgb = rgb("WHITE")

    # Apply template's progressive sage-to-dark-teal bar ramp
    ramp = ["ICE_MIST", "SOFT_SAGE", "SAGE_GLASS", "MUTED_TEAL", "DARK_TEAL", "DEEP_PINE"]
    series = plot.series[0]
    for i in range(len(values)):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(ramp[i % len(ramp)])

    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.name = FONT_BODY
        axis.tick_labels.font.size = Pt(15)
        axis.tick_labels.font.color.rgb = rgb("ICE_MIST")
        axis.format.line.color.rgb = rgb("SAGE_GLASS")
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = rgb("SAGE_GLASS")

    h1 = add_textbox(
        slide,
        left=13.25,
        top=2.55,
        width=5.63,
        height=2.8,
        text=title,
        font_name=FONT_TITLE,
        font_size=42.0,
        color="WHITE",
        bold=True,
        line_spacing=1.12,
        name="SlideTitle",
    )
    body = add_textbox(
        slide,
        left=13.25,
        top=5.65,
        width=5.63,
        height=3.6,
        text=insight,
        font_name=FONT_BODY,
        font_size=22.0,
        color="ICE_MIST",
        line_spacing=1.28,
        name="ChartInsight",
    )

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(panel, trigger="on_load", duration_ms=340)
    anim.fade_in(h1, trigger="with_previous", duration_ms=380)
    anim.wipe_in(chart_shape, direction="up", trigger="after_previous", duration_ms=500)
    anim.fade_in(body, trigger="after_previous", duration_ms=380)
    anim.apply()
    return slide, anim, chart_shape


# ---------------------------------------------------------------------------
# L07 — 6-Node Connected Architecture Flowchart (Template Slide 07)
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
    """Build a 6-step serpentine connected architecture flow (Top L->R, Bottom R->L)."""
    slide = add_blank_slide(prs, bg="radial")
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=17.5,
        height=1.15,
        text=title,
        font_name=FONT_TITLE,
        font_size=56.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=2.25,
        width=15.0,
        height=0.55,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=24.0,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    # Serpentine coordinates: 0,1,2 top row (L->R); 3,4,5 bottom row (R->L)
    box_w, box_h = 4.2, 1.32
    xs = [1.55, 7.90, 14.25]
    y_top, y_bot = 3.55, 7.10
    coords = [
        (xs[0], y_top),
        (xs[1], y_top),
        (xs[2], y_top),
        (xs[2], y_bot),
        (xs[1], y_bot),
        (xs[0], y_bot),
    ]

    nodes = []
    captions = []
    for idx, (stitle, sdesc) in enumerate(steps[:6]):
        x, y = coords[idx]
        box = add_frosted_card(slide, x, y, box_w, box_h, radius=0.14, name=f"FlowNode_{idx+1}")
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f"{idx+1}. {stitle}"
        r.font.name = FONT_BODY
        r.font.size = Pt(23.0)
        r.font.bold = True
        r.font.color.rgb = rgb("WHITE")
        nodes.append(box)

        cap = add_textbox(
            slide,
            left=x,
            top=y + box_h + 0.14,
            width=box_w,
            height=0.90,
            text=sdesc,
            font_name=FONT_BODY,
            font_size=18.5,
            color="ICE_MIST",
            align=PP_ALIGN.CENTER,
            line_spacing=1.16,
            name=f"FlowDesc_{idx+1}",
        )
        captions.append(cap)

    # Arrows: 0->1, 1->2, 2->3 (clean vertical drop below caption), 3->4 (left), 4->5 (left)
    arrows = []
    cy_top = y_top + box_h / 2.0
    cy_bot = y_bot + box_h / 2.0
    if len(nodes) >= 2:
        arrows.append(add_arrow_connector(slide, xs[0] + box_w + 0.14, cy_top, xs[1] - 0.14, cy_top, name="Arr_1_2"))
    if len(nodes) >= 3:
        arrows.append(add_arrow_connector(slide, xs[1] + box_w + 0.14, cy_top, xs[2] - 0.14, cy_top, name="Arr_2_3"))
    if len(nodes) >= 4:
        arrows.append(add_arrow_connector(slide, xs[2] + box_w / 2.0, y_top + box_h + 1.15, xs[2] + box_w / 2.0, y_bot - 0.12, name="Arr_3_4"))
    if len(nodes) >= 5:
        arrows.append(add_arrow_connector(slide, xs[2] - 0.14, cy_bot, xs[1] + box_w + 0.14, cy_bot, name="Arr_4_5"))
    if len(nodes) >= 6:
        arrows.append(add_arrow_connector(slide, xs[1] - 0.14, cy_bot, xs[0] + box_w + 0.14, cy_bot, name="Arr_5_6"))

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)

    # Animate step-by-step: node + caption (+ incoming arrow)
    reveal_groups = []
    for i in range(len(nodes)):
        grp = []
        if i > 0 and i - 1 < len(arrows):
            grp.append(arrows[i - 1])
        grp.extend([nodes[i], captions[i]])
        reveal_groups.append(grp)

    if click_reveal:
        anim.click_sequence(reveal_groups, effect="fade", duration_ms=340, intra_delay_ms=70)
    else:
        for grp in reveal_groups:
            anim.stagger(grp, effect="fade", step_ms=70, on_click_first=False, base_delay_ms=60)
    anim.apply()
    return slide, anim, nodes


# ---------------------------------------------------------------------------
# L08 — Left Title + 4 Floating Translucent Glass Circles (Template Slide 08)
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
    add_motif(slide, "glow", left=-1.5, top=4.6, width=9.5)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=3.25,
        width=6.8,
        height=2.4,
        text=title,
        font_name=FONT_TITLE,
        font_size=56.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=5.95,
        width=6.4,
        height=2.2,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=24.0,
        color="ICE_MIST",
        line_spacing=1.28,
        name="SlideSubtitle",
    )

    specs = [
        (8.95, 0.90, 4.15),
        (14.10, 1.65, 3.85),
        (10.05, 5.75, 3.35),
        (14.35, 5.85, 4.10),
    ]
    circles = []
    for idx, txt in enumerate(items[:4]):
        cx, cy, diam = specs[idx]
        c = add_glass_circle(slide, cx, cy, diam, text=txt, font_size=21.0, name=f"Circle_{idx+1}")
        circles.append(c)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=360)
    anim.fade_in(sub, trigger="with_previous", duration_ms=360, delay_ms=100)
    anim.stagger(circles, effect="zoom", step_ms=140, on_click_first=False)
    anim.apply()
    return slide, anim, circles


# ---------------------------------------------------------------------------
# L09 — Left Title + Nested Arch + 2 Tall Comparison Columns (Template Slide 09)
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
    art = add_motif(slide, "nested_arch", left=0.85, top=4.7, width=4.6)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.15,
        width=5.4,
        height=2.2,
        text=title,
        font_name=FONT_TITLE,
        font_size=48.0,
        color="WHITE",
        bold=True,
        line_spacing=1.10,
        name="SlideTitle",
    )
    sub_tb = None
    if subtitle:
        sub_tb = add_textbox(
            slide,
            left=MARGIN_X,
            top=3.35,
            width=5.2,
            height=1.5,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=21.0,
            color="ICE_MIST",
            line_spacing=1.24,
            name="SlideSubtitle",
        )

    col_w = 5.65
    col_h = 8.75
    col_y = 1.15
    col_xs = [7.15, 13.23]
    col_data = [
        (left_col, left_badge, left_accent, "LeftColumn"),
        (right_col, right_badge, right_accent, "RightColumn"),
    ]

    cards = []
    for idx, ((ctitle, bullets), cbadge, accent, cname) in enumerate(col_data):
        x = col_xs[idx]
        card = add_frosted_card(
            slide,
            x,
            col_y,
            col_w,
            col_h,
            fill_alpha=28.0,
            border_color=accent,
            border_alpha=55.0,
            border_pt=1.25,
            name=cname,
        )
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.42)
        tf.margin_right = Inches(0.42)
        tf.margin_top = Inches(0.45 if not cbadge else 1.15)

        if cbadge:
            add_badge(
                slide,
                left=x + 0.42,
                top=col_y + 0.35,
                width=2.8,
                height=0.52,
                text=cbadge,
                fill_color=accent,
                fill_alpha=22.0,
                border_color=accent,
                border_alpha=65.0,
                text_color="WHITE",
                font_size=16.0,
                name=f"{cname}_Badge",
            )

        p0 = tf.paragraphs[0]
        p0.space_after = Pt(18)
        r0 = p0.add_run()
        r0.text = ctitle
        r0.font.name = FONT_TITLE
        r0.font.size = Pt(30.0)
        r0.font.bold = True
        r0.font.color.rgb = rgb("WHITE")

        for b in bullets:
            pb = tf.add_paragraph()
            pb.space_after = Pt(14)
            pb.line_spacing = 1.24
            rb = pb.add_run()
            rb.text = f"▸  {b}"
            rb.font.name = FONT_BODY
            rb.font.size = Pt(21.0)
            rb.font.color.rgb = rgb("WHITE")
        cards.append(card)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(art, trigger="on_load", duration_ms=450)
    anim.fade_in(h1, trigger="with_previous", duration_ms=360)
    if sub_tb is not None:
        anim.fade_in(sub_tb, trigger="with_previous", duration_ms=360, delay_ms=80)
    anim.stagger(cards, effect="fly_bottom", step_ms=180, on_click_first=False)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L10 — Centered Title + 3 Vertical Showcase Cards (Template Slide 10)
# ---------------------------------------------------------------------------
def L10_three_cards(
    prs,
    title: str,
    cards_data: Sequence[Tuple[str, str, str]],
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    """cards_data: list of 3 tuples `(badge_label, heading, body)`."""
    slide = add_blank_slide(prs, bg="vertical")
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.95,
        width=20.0 - 2 * MARGIN_X,
        height=1.1,
        text=title,
        font_name=FONT_TITLE,
        font_size=56.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="SlideTitle",
    )

    cw = 5.55
    ch = 7.45
    gap = 0.55
    start_x = (20.0 - (3 * cw + 2 * gap)) / 2.0
    top_y = 2.45

    cards = []
    for idx, (tag, heading, body) in enumerate(cards_data[:3]):
        x = start_x + idx * (cw + gap)
        card = add_frosted_card(slide, x, top_y, cw, ch, fill_alpha=28.0, name=f"TriCard_{idx+1}")
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.42)
        tf.margin_right = Inches(0.42)
        tf.margin_top = Inches(0.48)

        p_tag = tf.paragraphs[0]
        p_tag.space_after = Pt(16)
        r_tag = p_tag.add_run()
        r_tag.text = tag.upper()
        r_tag.font.name = FONT_BODY
        r_tag.font.size = Pt(17.0)
        r_tag.font.bold = True
        r_tag.font.color.rgb = rgb("SOFT_SAGE")

        p_h = tf.add_paragraph()
        p_h.space_after = Pt(18)
        p_h.line_spacing = 1.12
        r_h = p_h.add_run()
        r_h.text = heading
        r_h.font.name = FONT_TITLE
        r_h.font.size = Pt(30.0)
        r_h.font.bold = True
        r_h.font.color.rgb = rgb("WHITE")

        p_b = tf.add_paragraph()
        p_b.line_spacing = 1.28
        r_b = p_b.add_run()
        r_b.text = body
        r_b.font.name = FONT_BODY
        r_b.font.size = Pt(21.0)
        r_b.font.color.rgb = rgb("ICE_MIST")

        cards.append(card)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(h1, trigger="on_load", duration_ms=350)
    anim.stagger(cards, effect="fly_bottom", step_ms=150, on_click_first=False)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L11 — Left Title + Crescents + 4 Stepped-Opacity Rows (Template Slide 11)
# ---------------------------------------------------------------------------
def L11_stepped_stack(
    prs,
    title: str,
    rows: Sequence[Tuple[str, str]],
    subtitle: Optional[str] = None,
    part_label: Optional[str] = None,
    transition: str = "fade",
):
    """Matches Slide 11's signature 11% -> 21% -> 31% -> 46% stepped glass cards."""
    slide = add_blank_slide(prs, bg="diagonal")
    art = add_motif(slide, "crescents", left=0.85, top=4.75, width=4.8)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.15,
        width=6.2,
        height=2.0,
        text=title,
        font_name=FONT_TITLE,
        font_size=56.0,
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
            top=3.20,
            width=5.8,
            height=1.5,
            text=subtitle,
            font_name=FONT_BODY,
            font_size=22.0,
            color="ICE_MIST",
            line_spacing=1.25,
            name="SlideSubtitle",
        )

    alphas = [14.0, 22.0, 31.0, 42.0]
    row_x = 8.15
    row_w = 10.73
    row_h = 1.85
    gap_y = 0.42
    start_y = 1.25

    cards = []
    for idx, (rtitle, rdesc) in enumerate(rows[:4]):
        y = start_y + idx * (row_h + gap_y)
        a = alphas[idx % len(alphas)]
        card = add_frosted_card(
            slide,
            row_x,
            y,
            row_w,
            row_h,
            fill_alpha=a,
            border_alpha=min(60.0, a + 10.0),
            radius=0.05,
            name=f"SteppedRow_{idx+1}",
        )
        _style_card_title_body(
            card,
            heading=rtitle,
            body=rdesc,
            head_size=26.0,
            body_size=20.0,
            align=PP_ALIGN.LEFT,
            valign=MSO_ANCHOR.MIDDLE,
            text_color="WHITE",
            body_color="WHITE" if idx >= 2 else "ICE_MIST",
        )
        cards.append(card)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(art, trigger="on_load", duration_ms=460)
    anim.fade_in(h1, trigger="with_previous", duration_ms=360)
    if sub_tb is not None:
        anim.fade_in(sub_tb, trigger="with_previous", duration_ms=360, delay_ms=80)
    anim.stagger(cards, effect="fly_right", step_ms=130, on_click_first=False)
    anim.apply()
    return slide, anim, cards


# ---------------------------------------------------------------------------
# L12 — Frosted Dark Summary Panel + Light Sage Checklist Card (Template Slide 12)
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
    slide = add_blank_slide(prs, bg="horizontal")
    add_motif(slide, "glow", left=-2.8, top=-3.4, width=8.5)

    left_panel = add_frosted_card(slide, MARGIN_X, 1.45, 7.85, 8.20, fill_alpha=24.0, name="LeftSummaryPanel")
    h1 = add_textbox(
        slide,
        left=MARGIN_X + 0.65,
        top=2.35,
        width=6.55,
        height=2.6,
        text=title,
        font_name=FONT_TITLE,
        font_size=46.0,
        color="WHITE",
        bold=True,
        line_spacing=1.10,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X + 0.65,
        top=5.30,
        width=6.55,
        height=3.4,
        text=subtitle,
        font_name=FONT_BODY,
        font_size=23.0,
        color="ICE_MIST",
        line_spacing=1.28,
        name="SlideSubtitle",
    )

    right_x = 9.55
    right_w = 9.33
    right_panel = add_solid_card(slide, right_x, 1.45, right_w, 8.20, fill_color="ICE_MIST", radius=0.04, name="ChecklistCard")
    r_head = add_textbox(
        slide,
        left=right_x + 0.55,
        top=1.95,
        width=right_w - 1.1,
        height=0.75,
        text=checklist_title,
        font_name=FONT_TITLE,
        font_size=28.0,
        color="DARK_TEAL",
        bold=True,
        name="ChecklistHeader",
    )

    item_cards = []
    n = min(len(items), 5)
    row_h = 1.05
    gap = 0.24
    start_y = 2.95
    for idx, txt in enumerate(items[:n]):
        y = start_y + idx * (row_h + gap)
        row_box = add_solid_card(
            slide,
            left=right_x + 0.50,
            top=y,
            width=right_w - 1.0,
            height=row_h,
            fill_color="SOFT_SAGE",
            border_color="SAGE_GLASS",
            radius=0.12,
            name=f"ChecklistItem_{idx+1}",
        )
        tf = row_box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.32)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = f"✓   {txt}"
        r.font.name = FONT_BODY
        r.font.size = Pt(20.0)
        r.font.bold = False
        r.font.color.rgb = rgb("BLACK")
        item_cards.append(row_box)

    if part_label:
        add_footer(slide, part_label)

    anim = SlideAnimator(slide, transition=transition)
    anim.fade_in(left_panel, trigger="on_load", duration_ms=340)
    anim.fly_in(h1, direction="left", trigger="with_previous", duration_ms=400)
    anim.fade_in(sub, trigger="with_previous", duration_ms=380, delay_ms=100)
    anim.fade_in(right_panel, trigger="after_previous", duration_ms=340)
    anim.fade_in(r_head, trigger="with_previous", duration_ms=300)
    anim.stagger(item_cards, effect="fly_right", step_ms=110, on_click_first=False)
    anim.apply()
    return slide, anim, item_cards
