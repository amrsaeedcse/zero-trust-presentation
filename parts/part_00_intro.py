"""Part 00 — Shared Deck Intro: Full-Bleed Keynote Cover & 3-Layer Zero Trust Architecture Stack."""

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import L01_cover
from core.theme import (
    COLORS,
    FONT_BODY,
    FONT_DISPLAY,
    FONT_TITLE,
    MARGIN_X,
    _set_fill_alpha,
    _set_line_alpha,
    add_accent_strip,
    add_arrow_connector,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_textbox,
    rgb,
    set_notes,
)


def _add_layer_chip(slide, left: float, top: float, num: str, accent_token: str, name: str):
    chip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(0.56), Inches(0.36))
    chip.name = name
    try:
        chip.adjustments[0] = 0.35
    except Exception:
        pass
    _set_fill_alpha(chip, COLORS[accent_token], 100.0, elevated=False, raw_hex=True)
    _set_line_alpha(chip, None, width_pt=0)
    tf = chip.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    r.font.name = FONT_TITLE
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = rgb("PURE_WHITE")
    return chip


def build_part_00(prs) -> list:
    slides = []

    # Slide 1: True Full-Bleed Split Keynote Cover (L01)
    s1, _ = L01_cover(
        prs,
        title="Zero Trust Architecture\nfor University Campus Networks",
        subtitle="Eliminating implicit perimeter trust across open, heterogeneous academic environments",
        top_label="Computer Networks  ·  Team Research Presentation",
        badge_text="9-Pillar Campus Zero Trust",
        transition="fade",
    )
    set_notes(
        s1,
        "Opening cover slide for the 9-member Computer Networks presentation on Zero Trust Architecture for University Campus Networks.",
    )
    slides.append(s1)

    # Slide 2: 3-Layer Zero Trust Campus Architectural Stack & Team Map
    s2 = add_blank_slide(prs, bg="diagonal")
    h1 = add_textbox(
        s2,
        left=MARGIN_X,
        top=0.72,
        width=20.0 - 2 * MARGIN_X,
        height=0.86,
        text="Campus Zero Trust Architecture & Team Map",
        font_name=FONT_DISPLAY,
        font_size=46.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.LEFT,
        name="RoadmapTitle",
    )
    sub = add_textbox(
        s2,
        left=MARGIN_X,
        top=1.56,
        width=17.0,
        height=0.46,
        text="Three interconnected architectural layers uniting nine specialized pillars from threat modeling to classified data protection.",
        font_name=FONT_BODY,
        font_size=19.5,
        color="ICE_MIST",
        align=PP_ALIGN.LEFT,
        name="RoadmapSubtitle",
    )

    layers = [
        (
            "LAYER 01",
            "Foundation &\nIdentity",
            "VERIFIED_MINT",
            [
                ("01", "Threat Model & Boundaries", "Ahmed Hassan Mohamed Mostafa"),
                ("02", "Identity & Authentication", "Omar Abdelsalam Abdelfattah"),
                ("03", "Device Security Posture", "Omar Mounir Mohamed Arafa"),
            ],
        ),
        (
            "LAYER 02",
            "Policy &\nNetwork Fabric",
            "DARK_TEAL",
            [
                ("04", "Authorization Policies", "Omar Ahmed Abdelaziz Mashhour"),
                ("05", "Network Segmentation", "Ali Mohamed Ali Abdallah Omara"),
                ("06", "Remote Resource Access", "Amr Abdelazeem Saeed Mohamed"),
            ],
        ),
        (
            "LAYER 03",
            "Telemetry,\nOps & Data",
            "GOLD_ACCENT",
            [
                ("07", "Monitoring & Response", "Omar Mohamed Fawzy Fahmy"),
                ("08", "Adoption & Evaluation", "Amr Khaled Husseini Mohamed"),
                ("09", "Data Security & Classification", "Mahmoud Mohamed Ragab"),
            ],
        ),
    ]

    band_w = 20.0 - 2 * MARGIN_X  # 17.76"
    band_h = 2.40
    band_gap = 0.24
    start_y = 2.24

    label_col_w = 2.55
    cards_x0 = MARGIN_X + label_col_w + 0.25  # 3.92"
    cw = 4.54
    ch = 1.96
    cgap = 0.52

    layer_groups = []
    for l_idx, (l_tag, l_name, accent, members) in enumerate(layers):
        by = start_y + l_idx * (band_h + band_gap)
        # Recessed Porcelain Mist Layer Band
        band = add_frosted_card(
            s2,
            left=MARGIN_X,
            top=by,
            width=band_w,
            height=band_h,
            fill_color="DEEP_PINE",
            radius=0.05,
            name=f"LayerBand_{l_idx+1}",
        )
        rail = add_accent_strip(
            s2,
            left=MARGIN_X + 0.16,
            top=by + 0.26,
            width=0.09,
            height=band_h - 0.52,
            color=accent,
            rounded=True,
            name=f"LayerRail_{l_idx+1}",
        )
        l_tb = add_textbox(
            s2,
            left=MARGIN_X + 0.42,
            top=by + 0.34,
            width=label_col_w - 0.30,
            height=band_h - 0.68,
            text=l_tag,
            font_name=FONT_TITLE,
            font_size=12.5,
            color=accent,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
            name=f"LayerTitle_{l_idx+1}",
        )
        p_ln = l_tb.text_frame.add_paragraph()
        p_ln.space_before = Pt(6)
        p_ln.line_spacing = 1.12
        r_ln = p_ln.add_run()
        r_ln.text = l_name
        r_ln.font.name = FONT_DISPLAY
        r_ln.font.size = Pt(21.0)
        r_ln.font.bold = True
        r_ln.font.color.rgb = rgb("WHITE")

        cy = by + (band_h - ch) / 2.0
        row_elems = [band, rail, l_tb]
        for c_idx, (num, topic, owner) in enumerate(members):
            cx = cards_x0 + c_idx * (cw + cgap)
            card = add_frosted_card(
                s2,
                left=cx,
                top=cy,
                width=cw,
                height=ch,
                fill_color="WHITE",
                radius=0.055,
                name=f"RoadmapCard_{num}",
            )
            chip = _add_layer_chip(s2, cx + 0.26, cy + 0.24, num, accent, name=f"RoadmapChip_{num}")

            tf = card.text_frame
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.margin_left = Inches(0.26)
            tf.margin_right = Inches(0.24)
            tf.margin_top = Inches(0.72)

            p1 = tf.paragraphs[0]
            p1.alignment = PP_ALIGN.LEFT
            p1.space_after = Pt(5)
            p1.line_spacing = 1.10
            r1 = p1.add_run()
            r1.text = topic
            r1.font.name = FONT_TITLE
            r1.font.size = Pt(19.5)
            r1.font.bold = True
            r1.font.color.rgb = rgb("WHITE")

            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.LEFT
            r2 = p2.add_run()
            r2.text = owner
            r2.font.name = FONT_BODY
            r2.font.size = Pt(14.5)
            r2.font.color.rgb = rgb("ICE_MIST")

            row_elems.extend([card, chip])
            if c_idx < 2:
                arr = add_arrow_connector(
                    s2,
                    x1=cx + cw + 0.08,
                    y1=cy + ch / 2.0,
                    x2=cx + cw + cgap - 0.08,
                    y2=cy + ch / 2.0,
                    color=accent,
                    width_pt=1.8,
                    name=f"LayerArr_{num}",
                )
                row_elems.append(arr)

        layer_groups.append(row_elems)

    add_footer(s2, "Overview  ·  9-Pillar Architectural Stack")

    anim = SlideAnimator(s2, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=320, delay_ms=60)
    for l_idx, elems in enumerate(layer_groups):
        trig = "after_previous" if l_idx == 0 else "with_previous"
        base_d = 0 if l_idx == 0 else l_idx * 160
        for e_idx, el in enumerate(elems):
            if e_idx == 0:
                anim.fly_in(el, direction="bottom", trigger=trig, duration_ms=320, delay_ms=base_d)
            else:
                anim.fade_in(el, trigger="with_previous", duration_ms=260, delay_ms=base_d + 40)
    anim.apply()
    set_notes(
        s2,
        "Overview of all 9 interconnected parts arranged as a three-layer Zero Trust Campus Architecture stack.",
    )
    slides.append(s2)

    return slides
