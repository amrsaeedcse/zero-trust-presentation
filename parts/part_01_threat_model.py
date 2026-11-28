"""Part 01 — Threat Model and Boundaries.

Owner: Ahmed Hassan Mohamed Mostafa
Slide Budget: 5–6 slides target (7 ABSOLUTE MAX, including section divider).

In Scope (MUST cover):
- Why traditional university campus perimeter ("castle-and-moat") fails
  (open campus, BYOD, guests, smart-campus IoT, lateral movement, four broken assumptions).
- Campus threat landscape, Subjects, Enterprise Resources, and CIA protection priorities.
- Core Zero Trust tenets (Never Trust Always Verify, Least Privilege, Assume Breach)
  and NIST SP 800-207 boundary concepts (Subject, Resource, Implicit Trust Zone, PDP/PEP).
- Incident containment through resource-focused trust boundaries & engineering trade-offs.

Out of Scope (DO NOT re-explain):
- MFA/SSO protocols (P2), device posture/MDM (P3), policy engine internals (P4),
  VLAN/SDN segmentation plumbing (P5), VPN vs ZTNA (P6), SIEM/SOAR (P7),
  migration roadmap (P8), data classification tiers (P9).

Narrative Bridge:
- Opens the technical deck -> Bridges to Identity (Part 02) as the first verification pillar.
"""

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L02_grid_2x2,
    L03_section_divider,
    L05_table_compare,
    L10_three_cards,
    L12_split_checklist,
)
from core.theme import (
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    MARGIN_Y,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_textbox,
    rgb,
    set_notes,
)

PART_LABEL = "Part 01  ·  Threat Model & Boundaries"


def _add_card_text(
    card,
    heading: str,
    body: str = "",
    head_size: float = 19.0,
    body_size: float = 16.0,
    head_color: str = "WHITE",
    body_color: str = "ICE_MIST",
    align: PP_ALIGN = PP_ALIGN.CENTER,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
    margin_x: float = 0.14,
    margin_y: float = 0.08,
) -> None:
    """Style a compact diagram node card cleanly using locked theme typography."""
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin_x)
    tf.margin_right = Inches(margin_x)
    tf.margin_top = Inches(margin_y)
    tf.margin_bottom = Inches(margin_y)

    p1 = tf.paragraphs[0]
    p1.alignment = align
    p1.space_after = Pt(4 if body else 0)
    p1.line_spacing = 1.06
    r1 = p1.add_run()
    r1.text = heading
    r1.font.name = FONT_BODY
    r1.font.size = Pt(head_size)
    r1.font.bold = True
    r1.font.color.rgb = rgb(head_color)

    if body:
        for idx, line in enumerate(body.split("\n")):
            p2 = tf.add_paragraph()
            p2.alignment = align
            p2.line_spacing = 1.10
            if line == "":
                p2.space_after = Pt(2)
                continue
            r2 = p2.add_run()
            r2.text = line
            r2.font.name = FONT_BODY
            r2.font.size = Pt(body_size)
            r2.font.bold = False
            r2.font.color.rgb = rgb(body_color)


def _build_slide_03_boundary_shift(prs):
    """Slide 3: Custom animated architecture diagram — Legacy Flat Perimeter vs. ZTA Boundaries."""
    slide = add_blank_slide(prs, bg="radial")

    # Header (aligned to L04/L05 top geometry)
    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=MARGIN_Y - 0.10,
        width=9.9,
        height=1.15,
        text="Shifting the Trust Boundary",
        font_name=FONT_TITLE,
        font_size=48.0,
        color="WHITE",
        bold=True,
        line_spacing=1.08,
        name="SlideTitle",
    )
    intro = add_textbox(
        slide,
        left=11.30,
        top=MARGIN_Y - 0.04,
        width=7.58,
        height=1.15,
        text=(
            "NIST SP 800-207 shrinks the implicit trust zone from the entire campus LAN "
            "down to a single resource per verified session."
        ),
        font_name=FONT_BODY,
        font_size=21.0,
        color="ICE_MIST",
        line_spacing=1.20,
        name="SlideIntro",
    )

    # 3 Core Zero Trust Tenets strip
    tenet_labels = [
        "01  ·  Never Trust, Always Verify",
        "02  ·  Enforce Least Privilege",
        "03  ·  Assume Breach",
    ]
    tenet_pills = []
    pill_w = 5.68
    pill_gap = 0.36
    for idx, label in enumerate(tenet_labels):
        px = MARGIN_X + idx * (pill_w + pill_gap)
        pill = add_badge(
            slide,
            left=px,
            top=2.24,
            width=pill_w,
            height=0.60,
            text=label,
            fill_color="DARK_TEAL",
            fill_alpha=72.0,
            border_color="VERIFIED_MINT",
            border_alpha=60.0,
            text_color="WHITE",
            font_size=18.5,
            bold=True,
            name=f"TenetPill_{idx + 1}",
        )
        tenet_pills.append(pill)

    # ------------------------------------------------------------------
    # LEFT PANEL: Legacy Perimeter Model (Castle-and-Moat)
    # ------------------------------------------------------------------
    lx, ly, lw, lh = MARGIN_X, 3.16, 8.48, 6.96
    left_panel = add_frosted_card(
        slide,
        left=lx,
        top=ly,
        width=lw,
        height=lh,
        fill_alpha=18.0,
        border_color="ALERT_CORAL",
        border_alpha=55.0,
        border_pt=1.25,
        radius=0.03,
        name="LegacyPanel",
    )
    left_badge = add_badge(
        slide,
        left=lx + 0.30,
        top=ly + 0.24,
        width=2.85,
        height=0.46,
        text="LEGACY PERIMETER",
        fill_color="ALERT_CORAL",
        fill_alpha=28.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=15.5,
        bold=True,
        name="LegacyBadge",
    )
    left_head = add_textbox(
        slide,
        left=lx + 3.30,
        top=ly + 0.26,
        width=4.95,
        height=0.46,
        text="One Giant Implicit Trust Zone",
        font_name=FONT_TITLE,
        font_size=22.0,
        color="WHITE",
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="LegacyHeading",
    )

    # Top row inside Legacy: Subjects -> Outer Firewall
    leg_entry = add_frosted_card(
        slide,
        left=lx + 0.30,
        top=ly + 0.94,
        width=3.55,
        height=1.08,
        fill_alpha=24.0,
        border_color="SAGE_GLASS",
        border_alpha=40.0,
        radius=0.08,
        name="LegacySubjects",
    )
    _add_card_text(
        leg_entry,
        heading="Campus Entry",
        body="Students · BYOD\nGuests · Campus IoT",
        head_size=19.0,
        body_size=15.5,
    )
    arr_leg_fw = add_arrow_connector(
        slide,
        x1=lx + 3.92,
        y1=ly + 1.48,
        x2=lx + 4.46,
        y2=ly + 1.48,
        color="WHITE",
        width_pt=2.0,
        name="ArrLegacyFW",
    )
    leg_fw = add_frosted_card(
        slide,
        left=lx + 4.54,
        top=ly + 0.94,
        width=3.64,
        height=1.08,
        fill_color="DARK_TEAL",
        fill_alpha=88.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        border_pt=1.25,
        radius=0.08,
        name="LegacyFirewall",
    )
    _add_card_text(
        leg_fw,
        heading="Outer Border Firewall",
        body="Single check at network edge",
        head_size=19.0,
        body_size=15.5,
    )

    # Down arrow into Flat LAN (single crisp line label, generous gap)
    arr_fw_lan = add_arrow_connector(
        slide,
        x1=lx + 6.36,
        y1=ly + 2.08,
        x2=lx + 6.36,
        y2=ly + 2.58,
        color="ALERT_CORAL",
        width_pt=2.0,
        name="ArrFWtoLAN",
    )
    lbl_trusted = add_textbox(
        slide,
        left=lx + 2.20,
        top=ly + 2.13,
        width=3.95,
        height=0.38,
        text="Auto-trusted once inside",
        font_name=FONT_BODY,
        font_size=16.0,
        color="ALERT_CORAL",
        bold=True,
        align=PP_ALIGN.RIGHT,
        valign=MSO_ANCHOR.MIDDLE,
        name="LblAutoTrust",
    )

    # Flat LAN box
    lan_box = add_frosted_card(
        slide,
        left=lx + 0.30,
        top=ly + 2.66,
        width=7.88,
        height=3.10,
        fill_color="OBSIDIAN",
        fill_alpha=36.0,
        border_color="ALERT_CORAL",
        border_alpha=65.0,
        border_pt=1.25,
        radius=0.04,
        name="FlatLANBox",
    )
    lan_lbl = add_textbox(
        slide,
        left=lx + 0.50,
        top=ly + 2.78,
        width=7.48,
        height=0.36,
        text="FLAT CAMPUS LAN  ·  IMPLICIT TRUST ZONE",
        font_name=FONT_BODY,
        font_size=15.5,
        color="SOFT_SAGE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="FlatLANTitle",
    )

    # 3 internal resources inside Flat LAN
    res_w, res_h = 2.18, 1.24
    res_y = ly + 3.28
    res_xs = [lx + 0.52, lx + 3.15, lx + 5.78]

    leg_lms = add_frosted_card(
        slide, res_xs[0], res_y, res_w, res_h, fill_alpha=24.0, radius=0.08, name="LegacyLMS"
    )
    _add_card_text(leg_lms, "LMS Portal", "Slides & quizzes", head_size=18.5, body_size=15.5)

    leg_lab = add_frosted_card(
        slide,
        res_xs[1],
        res_y,
        res_w,
        res_h,
        fill_color="ALERT_CORAL",
        fill_alpha=26.0,
        border_color="ALERT_CORAL",
        border_alpha=85.0,
        border_pt=1.5,
        radius=0.08,
        name="LegacyLabCompromised",
    )
    _add_card_text(leg_lab, "Lab PC / GPU", "Compromised host", head_size=18.5, body_size=15.5, body_color="WHITE")

    leg_rec = add_frosted_card(
        slide, res_xs[2], res_y, res_w, res_h, fill_alpha=24.0, radius=0.08, name="LegacyRecords"
    )
    _add_card_text(leg_rec, "Student Records", "Grades & IDs", head_size=18.5, body_size=15.5)

    # Lateral movement red arrows (Lab -> LMS and Lab -> Records)
    arr_lat_l = add_arrow_connector(
        slide,
        x1=res_xs[1] - 0.04,
        y1=res_y + res_h / 2.0,
        x2=res_xs[0] + res_w + 0.06,
        y2=res_y + res_h / 2.0,
        color="ALERT_CORAL",
        width_pt=2.5,
        dashed=True,
        name="ArrLateralLeft",
    )
    arr_lat_r = add_arrow_connector(
        slide,
        x1=res_xs[1] + res_w + 0.04,
        y1=res_y + res_h / 2.0,
        x2=res_xs[2] - 0.06,
        y2=res_y + res_h / 2.0,
        color="ALERT_CORAL",
        width_pt=2.5,
        dashed=True,
        name="ArrLateralRight",
    )
    lat_note = add_textbox(
        slide,
        left=lx + 0.52,
        top=ly + 4.70,
        width=7.44,
        height=0.88,
        text="East–West Lateral Movement: One infected BYOD laptop or phished login pivots freely to crown-jewel databases.",
        font_name=FONT_BODY,
        font_size=17.5,
        color="ICE_MIST",
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
        line_spacing=1.18,
        name="LateralNote",
    )

    # Bottom Legacy Verdict Strip
    leg_verdict = add_badge(
        slide,
        left=lx + 0.30,
        top=ly + 6.00,
        width=7.88,
        height=0.68,
        text="Trust: Internal IP   ·   Blast Radius: Campus-Wide",
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="ALERT_CORAL",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=17.5,
        bold=True,
        name="LegacyVerdict",
    )

    # Center bridge arrow between Left and Right panels
    bridge_arrow = add_arrow_connector(
        slide,
        x1=lx + lw + 0.14,
        y1=ly + lh / 2.0,
        x2=lx + lw + 0.66,
        y2=ly + lh / 2.0,
        color="VERIFIED_MINT",
        width_pt=3.0,
        name="BridgeArrow",
    )

    # ------------------------------------------------------------------
    # RIGHT PANEL: Resource-Focused Zero Trust Boundaries (NIST SP 800-207)
    # ------------------------------------------------------------------
    rx, ry, rw, rh = 10.40, 3.16, 8.48, 6.96
    right_panel = add_frosted_card(
        slide,
        left=rx,
        top=ry,
        width=rw,
        height=rh,
        fill_alpha=20.0,
        border_color="VERIFIED_MINT",
        border_alpha=55.0,
        border_pt=1.25,
        radius=0.03,
        name="ZtaPanel",
    )
    right_badge = add_badge(
        slide,
        left=rx + 0.30,
        top=ry + 0.24,
        width=2.80,
        height=0.46,
        text="ZERO TRUST (ZTA)",
        fill_color="DARK_TEAL",
        fill_alpha=75.0,
        border_color="VERIFIED_MINT",
        border_alpha=80.0,
        text_color="VERIFIED_MINT",
        font_size=15.5,
        bold=True,
        name="ZtaBadge",
    )
    right_head = add_textbox(
        slide,
        left=rx + 3.25,
        top=ry + 0.26,
        width=4.95,
        height=0.46,
        text="Per-Resource Micro-Perimeters",
        font_name=FONT_TITLE,
        font_size=22.0,
        color="WHITE",
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
        name="ZtaHeading",
    )

    # Column 1 inside ZTA: Subject (high-contrast dark teal surface)
    subj_card = add_frosted_card(
        slide,
        left=rx + 0.30,
        top=ry + 1.00,
        width=2.16,
        height=4.76,
        fill_color="DARK_TEAL",
        fill_alpha=76.0,
        border_color="SAGE_GLASS",
        border_alpha=65.0,
        border_pt=1.0,
        radius=0.06,
        name="ZtaSubject",
    )
    _add_card_text(
        subj_card,
        heading="SUBJECT",
        body=(
            "User Identity\nStudents / Faculty\nAdmins / Guests\n\n"
            "+ Endpoint Device\nManaged / BYOD\n\n"
            "+ App Context\nUntrusted Wi-Fi"
        ),
        head_size=19.0,
        body_size=15.0,
        head_color="VERIFIED_MINT",
        body_color="WHITE",
    )

    # Column 2 inside ZTA: PDP (top) + PEP (bottom)
    pdp_card = add_frosted_card(
        slide,
        left=rx + 2.98,
        top=ry + 1.00,
        width=2.08,
        height=1.98,
        fill_color="DARK_TEAL",
        fill_alpha=90.0,
        border_color="VERIFIED_MINT",
        border_alpha=80.0,
        border_pt=1.25,
        radius=0.07,
        name="ZtaPDP",
    )
    _add_card_text(
        pdp_card,
        heading="PDP (Brain)",
        body="Evaluates identity,\ndevice health &\nresource risk",
        head_size=18.5,
        body_size=15.5,
    )

    arr_pdp_pep = add_arrow_connector(
        slide,
        x1=rx + 4.02,
        y1=ry + 3.04,
        x2=rx + 4.02,
        y2=ry + 3.54,
        color="VERIFIED_MINT",
        width_pt=2.0,
        name="ArrPDPtoPEP",
    )

    pep_card = add_frosted_card(
        slide,
        left=rx + 2.98,
        top=ry + 3.60,
        width=2.08,
        height=2.16,
        fill_color="DEEP_PINE",
        fill_alpha=92.0,
        border_color="VERIFIED_MINT",
        border_alpha=85.0,
        border_pt=1.5,
        radius=0.07,
        name="ZtaPEP",
    )
    _add_card_text(
        pep_card,
        heading="PEP (Gate)",
        body="Inline boundary:\nopens or drops\neach session",
        head_size=18.5,
        body_size=15.5,
        body_color="WHITE",
    )

    # Arrows Subject -> PDP (context) and Subject -> PEP (data plane)
    arr_sub_pdp = add_arrow_connector(
        slide,
        x1=rx + 2.50,
        y1=ry + 1.98,
        x2=rx + 2.93,
        y2=ry + 1.98,
        color="ICE_MIST",
        width_pt=1.8,
        dashed=True,
        name="ArrSubPDP",
    )
    arr_sub_pep = add_arrow_connector(
        slide,
        x1=rx + 2.50,
        y1=ry + 4.68,
        x2=rx + 2.93,
        y2=ry + 4.68,
        color="VERIFIED_MINT",
        width_pt=2.2,
        name="ArrSubPEP",
    )

    # Column 3 inside ZTA: 3 Isolated Protect Surfaces (explicit clean line breaks)
    zt_res_x = rx + 5.62
    zt_res_w = 2.56
    zt_res_h = 1.40
    zt_res_specs = [
        (ry + 1.00, "P1 · Student Records", "Strictest PEP\nNo Guests / BYOD"),
        (ry + 2.68, "P2 · Lab Servers", "Two-Way PEP\nBlocks Admin Pivot"),
        (ry + 4.36, "P3 · Learning Portal", "BYOD Friendly\nStep-Up for Exams"),
    ]
    zt_res_cards = []
    zt_res_arrows = []
    for idx, (zy, ztitle, zdesc) in enumerate(zt_res_specs):
        rcard = add_frosted_card(
            slide,
            left=zt_res_x,
            top=zy,
            width=zt_res_w,
            height=zt_res_h,
            fill_color="DARK_TEAL",
            fill_alpha=68.0,
            border_color="VERIFIED_MINT",
            border_alpha=75.0,
            border_pt=1.25,
            radius=0.08,
            name=f"ProtectSurface_{idx + 1}",
        )
        _add_card_text(rcard, ztitle, zdesc, head_size=17.5, body_size=15.0)
        zt_res_cards.append(rcard)

        arr = add_arrow_connector(
            slide,
            x1=rx + 5.10,
            y1=ry + 4.68,
            x2=zt_res_x - 0.06,
            y2=zy + zt_res_h / 2.0,
            color="VERIFIED_MINT",
            width_pt=2.0,
            name=f"ArrPEPtoRes_{idx + 1}",
        )
        zt_res_arrows.append(arr)

    # Bottom ZTA Verdict Strip
    zta_verdict = add_badge(
        slide,
        left=rx + 0.30,
        top=ry + 6.00,
        width=7.88,
        height=0.68,
        text="Trust: Identity + Posture + Risk   ·   Blast Radius: 1 Session",
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=17.5,
        bold=True,
        name="ZtaVerdict",
    )

    add_footer(slide, PART_LABEL)

    # Step-by-step storytelling animations
    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(intro, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.stagger(tenet_pills, effect="fly_top", duration_ms=340, step_ms=100, on_click_first=False, base_delay_ms=60)

    # Click 1: Legacy Perimeter & lateral movement exposure
    anim.fade_in(left_panel, trigger="on_click", duration_ms=340)
    anim.fade_in(left_badge, trigger="with_previous", duration_ms=300, delay_ms=60)
    anim.fade_in(left_head, trigger="with_previous", duration_ms=300, delay_ms=80)
    anim.stagger(
        [leg_entry, arr_leg_fw, leg_fw, arr_fw_lan, lbl_trusted, lan_box, lan_lbl, leg_lms, leg_lab, leg_rec],
        effect="fade",
        duration_ms=260,
        step_ms=65,
        on_click_first=False,
        base_delay_ms=40,
    )
    anim.wipe_in(arr_lat_l, direction="left", trigger="after_previous", duration_ms=340)
    anim.wipe_in(arr_lat_r, direction="right", trigger="with_previous", duration_ms=340)
    anim.fade_in(lat_note, trigger="after_previous", duration_ms=300)
    anim.fade_in(leg_verdict, trigger="with_previous", duration_ms=300, delay_ms=80)

    # Click 2: Zero Trust Resource-Focused Micro-Perimeters
    anim.wipe_in(bridge_arrow, direction="right", trigger="on_click", duration_ms=300)
    anim.fade_in(right_panel, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fade_in(right_badge, trigger="with_previous", duration_ms=300, delay_ms=120)
    anim.fade_in(right_head, trigger="with_previous", duration_ms=300, delay_ms=140)
    anim.stagger(
        [
            subj_card,
            arr_sub_pdp,
            pdp_card,
            arr_pdp_pep,
            arr_sub_pep,
            pep_card,
            zt_res_arrows[0],
            zt_res_cards[0],
            zt_res_arrows[1],
            zt_res_cards[1],
            zt_res_arrows[2],
            zt_res_cards[2],
            zta_verdict,
        ],
        effect="fade",
        duration_ms=260,
        step_ms=65,
        on_click_first=False,
        base_delay_ms=60,
    )
    anim.apply()

    set_notes(
        slide,
        "Slide 3 — Core Zero Trust Tenets & Architectural Boundary Shift (NIST SP 800-207 [1]):\n"
        "- Walk through the 3 core tenets first: Never Trust Always Verify, Least Privilege, Assume Breach.\n"
        "- Click 1 (Left): In the perimeter model, a single border firewall creates one giant Implicit Trust Zone "
        "across the campus LAN. One compromised lab PC or phished student login lets attackers pivot East-West "
        "toward Student Records.\n"
        "- Click 2 (Right): Zero Trust treats campus Wi-Fi as untrusted and places a Policy Enforcement Point (PEP) "
        "backed by a Policy Decision Point (PDP) directly in front of each Enterprise Resource. Access is evaluated "
        "per Subject (User + Device + App) per session: e.g., Guests are dropped at P1/P2; an authorized lecturer on "
        "an unhealthy laptop is blocked from P1/P2 and restricted to read-only LMS.\n"
        "- Reference: NIST SP 800-207 (Rose et al., 2020); Kindervag (2010).",
    )
    return slide


def build_part_01(prs) -> list:
    """Build Part 01 (6 animated slides) — Threat Model and Trust Boundaries."""
    slides = []

    # ------------------------------------------------------------------
    # Slide 1 — Section Divider (L03, bg=diagonal)
    # ------------------------------------------------------------------
    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 01  ·  THREAT MODEL",
        title="Threat Model &\nTrust Boundaries",
        summary=(
            "Why the traditional castle-and-moat campus perimeter fails under unmanaged BYOD, "
            "open research, and shared labs — and how Zero Trust relocates trust boundaries "
            "directly around every campus resource."
        ),
        presenter="Ahmed Hassan Mohamed Mostafa",
        motif="arc",
        transition="push_left",
    )
    set_notes(
        s1,
        "Opening technical section (Part 01 — Threat Model & Trust Boundaries, Ahmed Hassan Mohamed Mostafa, "
        "CSE 4th Year, Zagazig University). Introduces why university campuses cannot rely on legacy perimeter "
        "security and sets the threat model foundation using NIST SP 800-207.",
    )
    slides.append(s1)

    # ------------------------------------------------------------------
    # Slide 2 — Why Campus Perimeters Fail: 4 Broken Assumptions (L02, bg=vertical)
    # ------------------------------------------------------------------
    s2, _, _ = L02_grid_2x2(
        prs,
        title="Why Campus Perimeters Fail",
        subtitle="Open academic networks function like small cities — breaking four legacy trust assumptions",
        cards=[
            (
                "“Inside the Building = Safe”",
                "Thousands of unmanaged student BYOD laptops, phones, guests, and campus IoT carry malware straight past the outer firewall.",
            ),
            (
                "“University Devices Are Safe”",
                "Faculty laptops travel to public Wi-Fi, shared lab desktops meet student USBs, and institutional machines get lost or stolen.",
            ),
            (
                "“One Login Trusts Everything”",
                "A single phished password or perimeter entry exposes internal server interfaces across flat campus networks.",
            ),
            (
                "“Internal Traffic Is Harmless”",
                "Inspecting only North–South internet borders lets attackers pivot East–West from one lab PC to administrative databases.",
            ),
        ],
        part_label=PART_LABEL,
        transition="fade",
        on_click_cards=True,
    )
    set_notes(
        s2,
        "Slide 2 — The Open Campus Challenge & Four Broken Trust Assumptions:\n"
        "- Unlike closed corporate offices, a university functions like a busy small city where students, lecturers, "
        "administrators, and guests connect and disconnect hourly with unmanaged devices.\n"
        "- Walk through each broken assumption: (1) malware downloaded at home walks in on campus Wi-Fi; "
        "(2) university-owned laptops travel and lab PCs take student USBs; (3) one compromised login exposes internal "
        "servers; (4) ignoring East-West traffic enables lateral movement from a lab machine to Student Records.",
    )
    slides.append(s2)

    # ------------------------------------------------------------------
    # Slide 3 — Shifting the Trust Boundary (Custom Animated Diagram, bg=radial)
    # ------------------------------------------------------------------
    s3 = _build_slide_03_boundary_shift(prs)
    slides.append(s3)

    # ------------------------------------------------------------------
    # Slide 4 — Campus Asset Threat Matrix & Protection Priorities (L05, bg=diagonal)
    # ------------------------------------------------------------------
    s4, _, tbl_shape = L05_table_compare(
        prs,
        title="Asset Threat Matrix &\nProtection Priorities",
        intro=(
            "Boundary strictness follows academic risk and usability — locking down crown-jewel "
            "records without disrupting daily BYOD learning."
        ),
        headers=[
            "Asset & Priority",
            "CIA Impact",
            "Primary Campus Threats",
            "Required Boundary Control",
        ],
        rows=[
            [
                "P1 · Student Records\n(Crown Jewel)",
                "C: Critical · I: Critical\nA: Med-High",
                "Grade tampering, ID theft, phished logins, ransomware",
                "Strictest PEP: strong identity, healthy device, least privilege",
            ],
            [
                "P2 · Lab Servers\n(Compute & GPUs)",
                "C: High · I: High\nA: High",
                "Malicious scripts, GPU cryptojacking, lateral scanning",
                "Two-way PEP: verify enrollment; block pivot to admin DBs",
            ],
            [
                "P3 · Learning Portal\n(LMS)",
                "C: Medium · I: High\nA: Critical",
                "Exam leaks, quiz DoS, infected homework uploads",
                "BYOD access isolated from P1/P2; step-up for exams",
            ],
            [
                "Core · Control Plane\n(IdP & PDP)",
                "C: Critical · I: Critical\nA: Max Critical",
                "Policy tampering & DDoS causing campus-wide lockout",
                "Hardened plane backed by active-active clusters",
            ],
        ],
        col_widths=[3.46, 3.30, 5.25, 5.75],
        part_label=PART_LABEL,
        transition="fade",
    )
    # Refine table vertical placement and compact row heights so generous whitespace clears the footer
    tbl_shape.top = Inches(3.48)
    tbl = tbl_shape.table
    tbl.rows[0].height = Inches(0.86)
    for r_idx in range(1, len(tbl.rows)):
        tbl.rows[r_idx].height = Inches(1.26)
        for c_idx in range(len(tbl.columns)):
            cell = tbl.cell(r_idx, c_idx)
            cell.margin_top = Inches(0.10)
            cell.margin_bottom = Inches(0.10)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(19.5)

    set_notes(
        s4,
        "Slide 4 — Campus Asset Threat Matrix & Engineering Priority Justification:\n"
        "- P1 (Student-Records System): Highest priority because silent grade tampering is irreversible and national ID / "
        "financial leaks create severe legal liability; neither guests nor unmanaged student devices need admin access.\n"
        "- P2 (Laboratory Servers): Highest active compromise risk since students intentionally upload and run code; "
        "expensive GPU clusters attract cryptojacking, so a two-way boundary traps breaches inside the lab zone.\n"
        "- P3 (Learning Portal / LMS): High availability for thousands of BYOD students downloading slides, with "
        "step-up verification reserved for high-risk actions like exam uploads.\n"
        "- Core (IdP & PDP): Operational backbone — if compromised or down, every PEP fails closed.",
    )
    slides.append(s4)

    # ------------------------------------------------------------------
    # Slide 5 — Incident Containment: Failure Scenarios (L10 + Outcome Pills, bg=vertical)
    # ------------------------------------------------------------------
    s5, anim5, _ = L10_three_cards(
        prs,
        title="How Boundaries Contain Real Failures",
        cards_data=[
            (
                "Scenario 01 · Physical Loss",
                "Lost or Stolen Laptop",
                "Legacy risk: saved Wi-Fi or perimeter profiles leave internal servers exposed.\n\n"
                "Zero Trust boundary: physical possession grants zero trust — PEPs demand fresh verification "
                "and block lost devices even on campus Wi-Fi.",
            ),
            (
                "Scenario 02 · Credential Theft",
                "Phished Student Login",
                "Legacy risk: attacker enters campus LAN and scans administrative databases.\n\n"
                "Zero Trust boundary: PEPs keep internal servers dark, drop student traffic at P1/P2, "
                "and restrict exposure to basic LMS slides.",
            ),
            (
                "Scenario 03 · Control Plane",
                "IdP / PDP Outage",
                "Operational risk: if the central IdP or PDP crashes or faces DDoS, PEPs cannot validate "
                "new session requests.\n\n"
                "Design trade-off: boundaries fail closed — making active-active clusters and cached tokens mandatory.",
            ),
        ],
        part_label=PART_LABEL,
        transition="fade",
    )
    # Add high-contrast outcome pills anchored cleanly at the bottom of each card
    outcome_specs = [
        ("OUTCOME: BLOCKED AT THE DOOR", "VERIFIED_MINT"),
        ("OUTCOME: >90% BLAST RADIUS CUT", "VERIFIED_MINT"),
        ("TRADE-OFF: FAIL-CLOSED RESILIENCE", "ALERT_CORAL"),
    ]
    cw, gap = 5.55, 0.55
    start_x = (20.0 - (3 * cw + 2 * gap)) / 2.0
    pills5 = []
    for idx, (otext, ocolor) in enumerate(outcome_specs):
        px = start_x + idx * (cw + gap) + 0.36
        pill = add_badge(
            s5,
            left=px,
            top=8.98,
            width=cw - 0.72,
            height=0.60,
            text=otext,
            fill_color="DARK_TEAL",
            fill_alpha=86.0,
            border_color=ocolor,
            border_alpha=85.0,
            text_color="WHITE",
            font_size=16.0,
            bold=True,
            name=f"ScenarioOutcome_{idx + 1}",
        )
        pills5.append(pill)

    # Re-apply slide 5 animations so outcome pills animate smoothly with their cards
    anim5.stagger(pills5, effect="fade", duration_ms=340, step_ms=150, on_click_first=False, base_delay_ms=120)
    anim5.apply()

    set_notes(
        s5,
        "Slide 5 — Evaluating the Threat Model Through Three Realistic Campus Failure Scenarios:\n"
        "1. Lost/Stolen Laptop: A lecturer leaves a laptop in a taxi — resource boundaries nullify saved network "
        "profiles and block the flagged device at the door.\n"
        "2. Phished Student Account: Stolen student credentials cannot scan internal subnets or touch P1 Records / "
        "P2 Lab Servers, cutting the blast radius by over 90% down to basic LMS course slides.\n"
        "3. Unavailable Identity Service: Exposes the primary Zero Trust trade-off — because PEPs fail closed to stay "
        "secure, redundant active-active IdP/PDP clusters and short-term cached session tokens are essential.",
    )
    slides.append(s5)

    # ------------------------------------------------------------------
    # Slide 6 — Engineering Trade-Offs & Handoff to Identity (L12, bg=horizontal)
    # ------------------------------------------------------------------
    s6, _, _ = L12_split_checklist(
        prs,
        title="Threat Model\nTakeaways",
        subtitle=(
            "Per-resource boundaries eliminate implicit LAN trust while keeping open campus workflows intact.\n\n"
            "With network location gone as proof of safety, every access decision begins with verifying "
            "who is requesting the resource."
        ),
        checklist_title="Boundary Design Principles",
        items=[
            "Treat campus Wi-Fi as untrusted — abandon location-based trust",
            "Place PEP/PDP boundaries directly in front of each critical resource",
            "Tier strictness by risk (P1 Records, P2 Labs, P3 LMS) for usability",
            "Build active-active IdP/PDP clusters to preserve fail-closed uptime",
            "Login alone never suffices — strong Identity is our first pillar",
        ],
        part_label=PART_LABEL,
        transition="fade",
    )
    set_notes(
        s6,
        "Slide 6 — Key Engineering Lessons & Narrative Handoff to Part 02 (Identity & Authentication):\n"
        "- Summarize the four core boundary takeaways from our threat model (abandoning location trust, "
        "per-resource PEP/PDP placement, risk-tiered friction so BYOD learning stays usable, and resilient "
        "control-plane availability).\n"
        "- Key transition point: Authentication proves who is at the keyboard, while authorization and device "
        "health are checked separately at each boundary. Now that the outer perimeter is gone, our next step "
        "is establishing Identity as the campus's first verification pillar (hands off naturally to Part 02).",
    )
    slides.append(s6)

    return slides
