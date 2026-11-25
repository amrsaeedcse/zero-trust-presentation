"""Part 00 — Shared Deck Intro: Cover Slide & Unified 9-Pillar Roadmap (Coordinator-Owned).

DO NOT modify inside individual member part tasks.
"""

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import L01_cover
from core.theme import (
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_motif,
    add_textbox,
    rgb,
    set_notes,
)


def build_part_00(prs) -> list:
    slides = []

    # Slide 1: Hero Cover (L01)
    s1, _ = L01_cover(
        prs,
        title="Zero Trust Architecture\nfor University Campus Networks",
        subtitle="Eliminating implicit perimeter trust across open, heterogeneous academic environments",
        top_label="Computer Networks  ·  Team Research Presentation",
        badge_text="9-Pillar Campus ZTA",
        transition="fade",
    )
    set_notes(
        s1,
        "Opening cover slide for the 9-member Computer Networks presentation on Zero Trust Architecture for University Campus Networks.",
    )
    slides.append(s1)

    # Slide 2: Unified 9-Part Presentation Flow (Custom 3x3 Frosted Glass Roadmap)
    s2 = add_blank_slide(prs, bg="diagonal")
    h1 = add_textbox(
        s2,
        left=MARGIN_X,
        top=MARGIN_Y,
        width=20.0 - 2 * MARGIN_X,
        height=1.05,
        text="Presentation Roadmap & Team Architecture",
        font_name=FONT_TITLE,
        font_size=54.0,
        color="WHITE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="RoadmapTitle",
    )
    sub = add_textbox(
        s2,
        left=2.5,
        top=2.08,
        width=15.0,
        height=0.52,
        text="One continuous narrative from perimeter collapse to data-centric campus defense",
        font_name=FONT_BODY,
        font_size=22.0,
        color="ICE_MIST",
        align=PP_ALIGN.CENTER,
        name="RoadmapSubtitle",
    )

    roadmap = [
        ("01", "Threat Model & Boundaries", "Ahmed Hassan Mohamed Mostafa"),
        ("02", "Identity & Authentication", "Omar Abdelsalam Abdelfattah"),
        ("03", "Device Security Posture", "Omar Mounir Mohamed Arafa"),
        ("04", "Authorization Policies", "Omar Ahmed Abdelaziz Mashhour"),
        ("05", "Network Segmentation", "Ali Mohamed Ali Abdallah Omara"),
        ("06", "Remote Resource Access", "Amr Abdelazeem Saeed Mohamed"),
        ("07", "Monitoring & Response", "Omar Mohamed Fawzy Fahmy"),
        ("08", "Adoption & Evaluation", "Amr Khaled Husseini Mohamed"),
        ("09", "Data Security & Classification", "Mahmoud Mohamed Ragab"),
    ]

    cw, ch = 5.65, 2.15
    gx, gy = 0.42, 0.36
    sx = (20.0 - (3 * cw + 2 * gx)) / 2.0
    sy = 2.95

    cards = []
    for idx, (num, topic, owner) in enumerate(roadmap):
        r, c = divmod(idx, 3)
        x = sx + c * (cw + gx)
        y = sy + r * (ch + gy)
        card = add_frosted_card(s2, x, y, cw, ch, fill_alpha=26.0, radius=0.06, name=f"RoadmapCard_{num}")
        tf = card.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.32)
        tf.margin_right = Inches(0.26)

        p0 = tf.paragraphs[0]
        p0.space_after = Pt(6)
        r0 = p0.add_run()
        r0.text = f"PART {num}"
        r0.font.name = FONT_BODY
        r0.font.size = Pt(15.0)
        r0.font.bold = True
        r0.font.color.rgb = rgb("SOFT_SAGE")

        p1 = tf.add_paragraph()
        p1.space_after = Pt(6)
        r1 = p1.add_run()
        r1.text = topic
        r1.font.name = FONT_TITLE
        r1.font.size = Pt(23.0)
        r1.font.bold = True
        r1.font.color.rgb = rgb("WHITE")

        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = owner
        r2.font.name = FONT_BODY
        r2.font.size = Pt(17.0)
        r2.font.color.rgb = rgb("ICE_MIST")

        cards.append(card)

    add_footer(s2, "Overview  ·  9-Member Team")

    anim = SlideAnimator(s2, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.stagger(cards, effect="fly_bottom", step_ms=80, on_click_first=False)
    anim.apply()
    set_notes(s2, "Overview of all 9 interconnected parts and the team member presenting each pillar.")
    slides.append(s2)

    return slides
