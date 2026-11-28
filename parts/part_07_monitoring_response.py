"""Part 07 — Monitoring and Response.

Owner: Omar Mohamed Fawzy Fahmy Morsi Aboelenein
Slide Budget: 6 slides (within 5–6 target and 7 absolute max).

In Scope (MUST cover):
- Continuous Diagnostics & Mitigation (CDM) & campus telemetry (linking P2, P3, P5, P6 inputs).
- Centralized visibility with SIEM (log ingestion, normalization, correlation across departments/labs).
- Behavioral analytics with UEBA (baselines for students, faculty, IoT/lab devices; campus anomalies).
- Dynamic trust score & continuous re-evaluation feeding NIST SP 800-207 PE/PA (Part 04).
- Automated response via SOAR playbooks (step-up MFA, token revocation, PEP micro-quarantine, SOC alert)
  and core SOC metrics (MTTD, MTTR, False Positive Rate).

Out of Scope (DO NOT re-explain):
- Initial auth/MDM mechanics (P2–P3), ZTNA proxy tunneling (P6),
  migration ROI/maturity roadmap (P8), data classification tiers & DLP policy taxonomy (P9).

Narrative Bridge:
- Builds on Parts 01–06 (all controls emit telemetry) -> Bridges to Adoption & Evaluation (Part 08).
"""

from __future__ import annotations

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from core.animations import SlideAnimator
from core.layouts import (
    L03_section_divider,
    L04_wide_visual,
    L05_table_compare,
    L11_stepped_stack,
)
from core.theme import (
    FONT_BODY,
    FONT_TITLE,
    MARGIN_X,
    add_arrow_connector,
    add_badge,
    add_blank_slide,
    add_footer,
    add_frosted_card,
    add_motif,
    add_solid_card,
    add_textbox,
    rgb,
    set_notes,
)

PART_LABEL = "Part 07  ·  Monitoring & Response"

_TCPR_ORDER = [
    "lnL",
    "lnR",
    "lnT",
    "lnB",
    "lnTlToBr",
    "lnBlToTr",
    "cell3D",
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
    "headers",
    "extLst",
]


def _enforce_ltr(slide) -> None:
    """Force explicit LTR (`rtl="0"`) and `en-US` run language for clean rendering across all OS locales."""
    for p_el in slide._element.iter(qn("a:p")):
        pPr = p_el.find(qn("a:pPr"))
        if pPr is None:
            pPr = etree.Element(qn("a:pPr"))
            p_el.insert(0, pPr)
        pPr.set("rtl", "0")
    for rPr in slide._element.iter(qn("a:rPr")):
        rPr.set("lang", "en-US")


def _fix_table_tcpr_order(slide) -> None:
    """Sort `<a:tcPr>` children into exact ISO/IEC 29500 schema order (`lnL, lnR, lnT, lnB, solidFill`)."""
    rank = {name: idx for idx, name in enumerate(_TCPR_ORDER)}
    for tcPr in slide._element.iter(qn("a:tcPr")):
        children = list(tcPr)
        children.sort(key=lambda el: rank.get(etree.QName(el).localname, 999))
        for ch in children:
            tcPr.remove(ch)
        for ch in children:
            tcPr.append(ch)


def _style_footer_contrast(slide, left_color: str = "ICE_MIST", right_color: str = "ICE_MIST") -> None:
    """Ensure footer text maintains strong contrast against light or dark gradient corners."""
    for sh in slide.shapes:
        if sh.name == "FooterLeft" and sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = rgb(left_color)
        elif sh.name == "FooterRight" and sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = rgb(right_color)


def _add_card_text(
    card,
    heading: str,
    body: str,
    head_size: float = 23.0,
    body_size: float = 18.5,
    head_color: str = "WHITE",
    body_color: str = "ICE_MIST",
    margin_top: float = 0.22,
    margin_x: float = 0.28,
    space_after: float = 6.0,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
) -> None:
    """Populate a card with consistent Work Sans typography hierarchy."""
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin_x)
    tf.margin_right = Inches(margin_x)
    tf.margin_top = Inches(margin_top)
    tf.margin_bottom = Inches(0.18)

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    p0.space_after = Pt(space_after)
    r0 = p0.add_run()
    r0.text = heading
    r0.font.name = FONT_BODY
    r0.font.size = Pt(head_size)
    r0.font.bold = True
    r0.font.color.rgb = rgb(head_color)

    if body:
        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.LEFT
        p1.line_spacing = 1.18
        r1 = p1.add_run()
        r1.text = body
        r1.font.name = FONT_BODY
        r1.font.size = Pt(body_size)
        r1.font.color.rgb = rgb(body_color)


# ---------------------------------------------------------------------------
# Slide 2 — CDM Campus Telemetry & Centralized SIEM Pipeline
# ---------------------------------------------------------------------------
def _build_slide_2_cdm_siem(prs):
    slide, anim, _ = L04_wide_visual(
        prs,
        title="CDM Telemetry and\nCentralized SIEM",
        intro=(
            "Continuous Diagnostics & Mitigation (CDM) streams real-time logs from every "
            "campus control plane into a unified SIEM visibility pipeline."
        ),
        part_label=PART_LABEL,
        stage_card=False,
        transition="fade",
    )
    add_motif(slide, "glow", left=6.5, top=2.6, width=8.2)

    # Column labels
    lbl_1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=3.42,
        width=5.55,
        height=0.38,
        text="STEP 1 · CDM TELEMETRY SOURCES",
        font_name=FONT_BODY,
        font_size=16.0,
        color="SOFT_SAGE",
        bold=True,
        name="ColLabel_Sources",
    )
    lbl_2 = add_textbox(
        slide,
        left=7.60,
        top=3.42,
        width=5.65,
        height=0.38,
        text="STEP 2 · CENTRALIZED SIEM ENGINE",
        font_name=FONT_BODY,
        font_size=16.0,
        color="SOFT_SAGE",
        bold=True,
        name="ColLabel_SIEM",
    )
    lbl_3 = add_textbox(
        slide,
        left=14.20,
        top=3.42,
        width=4.68,
        height=0.38,
        text="STEP 3 · ACTIONABLE OUTPUTS",
        font_name=FONT_BODY,
        font_size=16.0,
        color="SOFT_SAGE",
        bold=True,
        name="ColLabel_Outputs",
    )

    # Column 1: 4 Telemetry Sources linking Parts 02, 03, 05, 06
    sources = [
        ("Identity Providers", "P2 · IdP", "SSO logins, MFA events, token claims"),
        ("Endpoint Telemetry", "P3 · EDR", "OS patch drift, EDR alerts, health state"),
        ("Campus Network", "P5 · Net", "Wi-Fi, switch NetFlow, east-west flows"),
        ("Remote & Cloud Apps", "P6 · ZTNA", "ZTNA sessions, LMS, ERP, lab portals"),
    ]
    src_cards = []
    src_badges = []
    src_arrows = []
    sy = 3.92
    sh = 1.34
    sgap = 0.20
    for idx, (stitle, stag, sdesc) in enumerate(sources):
        y = sy + idx * (sh + sgap)
        card = add_frosted_card(
            slide,
            left=MARGIN_X,
            top=y,
            width=5.55,
            height=sh,
            fill_color="DARK_TEAL",
            fill_alpha=72.0,
            border_color="SAGE_GLASS",
            border_alpha=58.0,
            border_pt=1.0,
            radius=0.08,
            name=f"TelemetrySource_{idx+1}",
        )
        _add_card_text(
            card,
            heading=stitle,
            body=sdesc,
            head_size=21.0,
            body_size=18.0,
            margin_top=0.14,
            margin_x=0.26,
            space_after=4.0,
        )
        badge = add_badge(
            slide,
            left=MARGIN_X + 3.65,
            top=y + 0.16,
            width=1.70,
            height=0.40,
            text=stag,
            fill_color="OBSIDIAN",
            fill_alpha=55.0,
            border_color="SAGE_GLASS",
            border_alpha=72.0,
            text_color="ICE_MIST",
            font_size=14.0,
            bold=True,
            name=f"SourceBadge_{idx+1}",
        )
        cy = y + sh / 2.0
        arr = add_arrow_connector(
            slide,
            x1=MARGIN_X + 5.62,
            y1=cy,
            x2=7.50,
            y2=cy,
            color="WHITE",
            width_pt=1.85,
            name=f"Arr_Src_SIEM_{idx+1}",
        )
        src_cards.append(card)
        src_badges.append(badge)
        src_arrows.append(arr)

    # Column 2: Centralized Campus SIEM Engine (3 internal pipeline stages)
    siem_shell = add_frosted_card(
        slide,
        left=7.60,
        top=3.92,
        width=5.65,
        height=5.96,
        fill_color="DARK_TEAL",
        fill_alpha=46.0,
        border_color="VERIFIED_MINT",
        border_alpha=68.0,
        border_pt=1.35,
        radius=0.05,
        name="SIEM_Container",
    )
    siem_stages = [
        ("Stage 1 · Log Ingestion", "Streams real-time TLS & API logs across faculties, dorms, and labs"),
        ("Stage 2 · Normalization", "Maps diverse vendor formats into one canonical event schema"),
        ("Stage 3 · Correlation", "Links identity, endpoint, and network events to spot attack chains"),
    ]
    stage_cards = []
    stage_arrows = []
    st_x = 7.90
    st_w = 5.05
    st_h = 1.50
    st_y0 = 4.22
    st_gap = 0.43
    for idx, (st_h_txt, st_b_txt) in enumerate(siem_stages):
        y = st_y0 + idx * (st_h + st_gap)
        sc = add_frosted_card(
            slide,
            left=st_x,
            top=y,
            width=st_w,
            height=st_h,
            fill_color="DARK_TEAL",
            fill_alpha=86.0,
            border_color="SAGE_GLASS",
            border_alpha=58.0,
            radius=0.08,
            name=f"SIEM_Stage_{idx+1}",
        )
        _add_card_text(
            sc,
            heading=st_h_txt,
            body=st_b_txt,
            head_size=21.0,
            body_size=18.0,
            margin_top=0.14,
            margin_x=0.24,
            space_after=4.0,
        )
        stage_cards.append(sc)
        if idx < 2:
            arr_d = add_arrow_connector(
                slide,
                x1=st_x + st_w / 2.0,
                y1=y + st_h + 0.06,
                x2=st_x + st_w / 2.0,
                y2=y + st_h + st_gap - 0.06,
                color="VERIFIED_MINT",
                width_pt=2.0,
                name=f"SIEM_DownArr_{idx+1}",
            )
            stage_arrows.append(arr_d)

    # Column 3: Outputs
    arr_out_1 = add_arrow_connector(
        slide, x1=13.32, y1=5.32, x2=14.12, y2=5.32, color="VERIFIED_MINT", width_pt=2.0, name="Arr_SIEM_UEBA"
    )
    arr_out_2 = add_arrow_connector(
        slide, x1=13.32, y1=8.48, x2=14.12, y2=8.48, color="WHITE", width_pt=2.0, name="Arr_SIEM_SOC"
    )

    out_1 = add_frosted_card(
        slide,
        left=14.20,
        top=3.92,
        width=4.68,
        height=2.80,
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="VERIFIED_MINT",
        border_alpha=68.0,
        border_pt=1.25,
        radius=0.06,
        name="Output_UEBA",
    )
    _add_card_text(
        out_1,
        heading="UEBA Analytics Feed",
        body="Delivers enriched user, device, and flow timelines for anomaly scoring.",
        head_size=23.5,
        body_size=19.0,
        margin_top=0.74,
        margin_x=0.30,
        space_after=10.0,
        valign=MSO_ANCHOR.TOP,
    )
    out_1_badge = add_badge(
        slide,
        left=14.50,
        top=4.14,
        width=2.05,
        height=0.42,
        text="NEXT STAGE",
        fill_color="VERIFIED_MINT",
        fill_alpha=24.0,
        border_color="VERIFIED_MINT",
        border_alpha=75.0,
        text_color="WHITE",
        font_size=14.5,
        name="Output_UEBA_Badge",
    )

    out_2 = add_frosted_card(
        slide,
        left=14.20,
        top=7.08,
        width=4.68,
        height=2.80,
        fill_color="DARK_TEAL",
        fill_alpha=78.0,
        border_color="SAGE_GLASS",
        border_alpha=60.0,
        border_pt=1.0,
        radius=0.06,
        name="Output_SOC",
    )
    _add_card_text(
        out_2,
        heading="Unified Campus SOC",
        body="Eliminates silos across autonomous departments, labs, and remote sessions.",
        head_size=23.5,
        body_size=19.0,
        margin_top=0.74,
        margin_x=0.30,
        space_after=10.0,
        valign=MSO_ANCHOR.TOP,
    )
    out_2_badge = add_badge(
        slide,
        left=14.50,
        top=7.30,
        width=2.15,
        height=0.42,
        text="VISIBILITY",
        fill_color="OBSIDIAN",
        fill_alpha=48.0,
        border_color="SAGE_GLASS",
        border_alpha=70.0,
        text_color="ICE_MIST",
        font_size=14.5,
        name="Output_SOC_Badge",
    )

    _style_footer_contrast(slide, left_color="DARK_TEAL", right_color="DARK_TEAL")

    # Progressive left-to-right diagram animations
    anim.fade_in(lbl_1, trigger="after_previous", duration_ms=260)
    for i in range(4):
        anim.fly_in(src_cards[i], direction="left", trigger="with_previous", duration_ms=340, delay_ms=90 * i)
        anim.fade_in(src_badges[i], trigger="with_previous", duration_ms=280, delay_ms=90 * i + 60)

    # Click 1: Telemetry flows into SIEM Core
    anim.add(src_arrows[0], effect="wipe_right", trigger="on_click", duration_ms=300)
    for a in src_arrows[1:]:
        anim.add(a, effect="wipe_right", trigger="with_previous", duration_ms=300, delay_ms=40)
    anim.fade_in(lbl_2, trigger="with_previous", duration_ms=280, delay_ms=80)
    anim.fade_in(siem_shell, trigger="with_previous", duration_ms=340, delay_ms=100)
    anim.fly_in(stage_cards[0], direction="bottom", trigger="with_previous", duration_ms=340, delay_ms=180)
    anim.wipe_in(stage_arrows[0], direction="down", trigger="with_previous", duration_ms=220, delay_ms=290)
    anim.fly_in(stage_cards[1], direction="bottom", trigger="with_previous", duration_ms=340, delay_ms=340)
    anim.wipe_in(stage_arrows[1], direction="down", trigger="with_previous", duration_ms=220, delay_ms=450)
    anim.fly_in(stage_cards[2], direction="bottom", trigger="with_previous", duration_ms=340, delay_ms=500)

    # Click 2: Outputs to UEBA & Campus SOC
    anim.add(arr_out_1, effect="wipe_right", trigger="on_click", duration_ms=280)
    anim.add(arr_out_2, effect="wipe_right", trigger="with_previous", duration_ms=280, delay_ms=60)
    anim.fade_in(lbl_3, trigger="with_previous", duration_ms=260, delay_ms=60)
    anim.fly_in(out_1, direction="right", trigger="with_previous", duration_ms=360, delay_ms=120)
    anim.fade_in(out_1_badge, trigger="with_previous", duration_ms=280, delay_ms=180)
    anim.fly_in(out_2, direction="right", trigger="with_previous", duration_ms=360, delay_ms=240)
    anim.fade_in(out_2_badge, trigger="with_previous", duration_ms=280, delay_ms=300)
    anim.apply()

    _enforce_ltr(slide)
    set_notes(
        slide,
        "[~40 sec] Continuous Diagnostics and Mitigation turns every control we just designed into a live sensor. "
        "Identity logs from Part 2, endpoint posture signals from Part 3, east-west campus flows from Part 5, and "
        "remote ZTNA sessions from Part 6 all stream continuously into a centralized Campus SIEM. [Click] Inside the "
        "SIEM, raw logs are ingested, normalized into one canonical schema across independent departments and labs, "
        "and correlated to reveal multi-step attacks that isolated tools miss. [Click] That enriched stream powers two "
        "outputs: a unified Campus SOC view and our behavioral analytics engine—UEBA—which we examine next.",
    )
    return slide


# ---------------------------------------------------------------------------
# Slide 3 — Behavioral Analytics with UEBA (L05 Sage Comparison Table)
# ---------------------------------------------------------------------------
def _build_slide_3_ueba(prs):
    slide, _, tbl_shape = L05_table_compare(
        prs,
        title="Behavioral Analytics\nwith UEBA",
        intro=(
            "User & Entity Behavior Analytics learns normal baselines per campus role to catch "
            "subtle, credential-valid attacks that static rules miss."
        ),
        headers=[
            "Campus Entity",
            "Learned Normal Baseline",
            "Detected Anomaly Signal",
            "Campus Threat Scenario",
        ],
        rows=[
            [
                "Students",
                "Campus Wi-Fi and dorm logins to LMS & library portals during term hours",
                "Impossible travel across continents in 15 min or rapid failed MFA",
                "Compromised student account & credential stuffing",
            ],
            [
                "Faculty and\nResearchers",
                "Managed laptops accessing grading ERP and assigned research datasets",
                "Off-hours 3 AM bulk export from research GPU cluster to unknown host",
                "Off-hours lab exfiltration or hijacked session token",
            ],
            [
                "Lab Workstations",
                "Scheduled class sessions talking only to approved course servers",
                "Sudden SMB or SSH scanning across neighboring faculty subnets",
                "Unauthorized east-west lateral movement probe",
            ],
            [
                "Smart Campus\nIoT / BMS",
                "Fixed-volume telemetry to local HVAC, badge, and camera controllers",
                "Unexpected outbound DNS tunneling or beaconing to external IP",
                "Compromised IoT sensor used as covert foothold",
            ],
        ],
        col_widths=[2.96, 5.10, 5.20, 4.50],
        part_label=PART_LABEL,
        transition="fade",
    )

    # Keep table compact with generous clearance above the 10.55" footer
    tbl_shape.top = Inches(3.50)
    table = tbl_shape.table
    table.rows[0].height = Inches(0.82)
    for ri in range(1, len(table.rows)):
        table.rows[ri].height = Inches(1.28)
        for cell in table.rows[ri].cells:
            cell.margin_top = Inches(0.10)
            cell.margin_bottom = Inches(0.10)
            for p in cell.text_frame.paragraphs:
                p.line_spacing = 1.12
                for r in p.runs:
                    r.font.size = Pt(19.5)

    _fix_table_tcpr_order(slide)
    _enforce_ltr(slide)

    set_notes(
        slide,
        "[~45 sec] Correlation alone only catches known attack signatures. To detect stolen credentials or insider "
        "misuse, UEBA models normal behavior for every campus entity. For students, it learns typical dorm and LMS "
        "patterns—so impossible travel or a credential-stuffing burst immediately spikes risk. For faculty and "
        "researchers, a 3 AM bulk transfer from a research cluster stands out sharply against daytime grading "
        "routines. For lab PCs and smart-campus IoT, any sudden east-west port scan or unexpected outbound beacon "
        "breaks a tightly predictable baseline. Once UEBA flags these anomalies, the key Zero Trust question is: "
        "how do those risk signals change active access in real time?",
    )
    return slide


# ---------------------------------------------------------------------------
# Slide 4 — Dynamic Trust Score & Continuous Re-Evaluation (Custom Closed Loop)
# ---------------------------------------------------------------------------
def _build_slide_4_closed_loop(prs):
    slide = add_blank_slide(prs, bg="radial")
    sig_motif = add_motif(slide, "signal_waves", left=15.6, top=-0.8, width=4.8)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=16.5,
        height=0.92,
        text="Dynamic Trust Score & Continuous Re-Evaluation",
        font_name=FONT_TITLE,
        font_size=50.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.88,
        width=16.0,
        height=0.48,
        text="Closing the NIST SP 800-207 loop: live SIEM/UEBA anomalies continuously recalibrate Policy Engine trust.",
        font_name=FONT_BODY,
        font_size=21.5,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    # 4 Loop Nodes (Top-Left, Top-Right, Bottom-Right, Bottom-Left)
    nw, nh_top, nh_bot = 5.85, 2.52, 2.72
    lx, rx = MARGIN_X, 20.0 - MARGIN_X - nw  # 1.12, 13.03
    ty, by = 2.88, 7.18

    # Node 1: Active Session & PEP
    n1 = add_frosted_card(
        slide, lx, ty, nw, nh_top, fill_color="DARK_TEAL", fill_alpha=76.0,
        border_color="VERIFIED_MINT", border_alpha=68.0, border_pt=1.25, radius=0.06, name="LoopNode_1_PEP",
    )
    _add_card_text(
        n1,
        heading="Policy Enforcement Point · PEP",
        body="Authorized user and device access campus apps; every active session emits live telemetry.",
        head_size=22.5,
        body_size=18.5,
        margin_top=0.74,
        margin_x=0.30,
        valign=MSO_ANCHOR.TOP,
    )
    b1 = add_badge(
        slide, lx + 0.30, ty + 0.20, 3.15, 0.42, "STEP 1 · ACTIVE SESSION",
        fill_color="VERIFIED_MINT", fill_alpha=24.0, border_color="VERIFIED_MINT", border_alpha=75.0,
        text_color="WHITE", font_size=14.0, name="Badge_Loop_1",
    )

    # Node 2: SIEM + UEBA Risk Telemetry
    n2 = add_frosted_card(
        slide, rx, ty, nw, nh_top, fill_color="DARK_TEAL", fill_alpha=76.0,
        border_color="ALERT_CORAL", border_alpha=70.0, border_pt=1.25, radius=0.06, name="LoopNode_2_UEBA",
    )
    _add_card_text(
        n2,
        heading="SIEM + UEBA Anomaly Detection",
        body="Detects mid-session behavioral drift such as impossible travel, off-hours lab bursts, or scans.",
        head_size=22.5,
        body_size=18.5,
        margin_top=0.74,
        margin_x=0.30,
        valign=MSO_ANCHOR.TOP,
    )
    b2 = add_badge(
        slide, rx + 0.30, ty + 0.20, 3.25, 0.42, "STEP 2 · ANOMALY SIGNAL",
        fill_color="ALERT_CORAL", fill_alpha=26.0, border_color="ALERT_CORAL", border_alpha=80.0,
        text_color="WHITE", font_size=14.0, name="Badge_Loop_2",
    )

    # Node 3: Real-Time Trust Re-Calculation
    n3 = add_frosted_card(
        slide, rx, by, nw, nh_bot, fill_color="DARK_TEAL", fill_alpha=78.0,
        border_color="ALERT_CORAL", border_alpha=70.0, border_pt=1.25, radius=0.06, name="LoopNode_3_Score",
    )
    _add_card_text(
        n3,
        heading="Dynamic Trust Score Decay",
        body="Risk weights drop session confidence immediately from 92 to 28 without waiting for token expiry.",
        head_size=22.5,
        body_size=18.5,
        margin_top=0.74,
        margin_x=0.30,
        valign=MSO_ANCHOR.TOP,
    )
    b3 = add_badge(
        slide, rx + 0.30, by + 0.22, 3.35, 0.42, "STEP 3 · TRUST RE-SCORING",
        fill_color="ALERT_CORAL", fill_alpha=26.0, border_color="ALERT_CORAL", border_alpha=80.0,
        text_color="WHITE", font_size=14.0, name="Badge_Loop_3",
    )

    # Node 4: NIST SP 800-207 Policy Engine & Administrator (Part 04 interface)
    n4 = add_frosted_card(
        slide, lx, by, nw, nh_bot, fill_color="DARK_TEAL", fill_alpha=78.0,
        border_color="SOFT_SAGE", border_alpha=72.0, border_pt=1.25, radius=0.06, name="LoopNode_4_PEPA",
    )
    _add_card_text(
        n4,
        heading="Policy Engine & Administrator",
        body="PE evaluates updated trust against policy; PA pushes real-time step-up or revoke commands to PEP.",
        head_size=22.5,
        body_size=18.5,
        margin_top=0.74,
        margin_x=0.30,
        valign=MSO_ANCHOR.TOP,
    )
    b4 = add_badge(
        slide, lx + 0.30, by + 0.22, 3.55, 0.42, "STEP 4 · NIST PE / PA · P4",
        fill_color="OBSIDIAN", fill_alpha=60.0, border_color="SOFT_SAGE", border_alpha=78.0,
        text_color="ICE_MIST", font_size=14.0, name="Badge_Loop_4",
    )

    # Center Focal Panel: Live Trust Score Thresholds
    cx, cy, cw, ch = 7.36, 4.58, 5.28, 3.40
    center_card = add_frosted_card(
        slide, cx, cy, cw, ch, fill_color="DARK_TEAL", fill_alpha=86.0,
        border_color="SAGE_GLASS", border_alpha=72.0, border_pt=1.25, radius=0.06,
        name="Center_TrustGauge",
    )
    center_hdr = add_textbox(
        slide,
        left=cx + 0.24,
        top=cy + 0.22,
        width=cw - 0.48,
        height=0.45,
        text="LIVE TRUST THRESHOLDS",
        font_name=FONT_BODY,
        font_size=17.0,
        color="SOFT_SAGE",
        bold=True,
        align=PP_ALIGN.CENTER,
        name="Center_GaugeTitle",
    )
    bands = [
        ("Trusted  ·  80–100  →  Allow", "VERIFIED_MINT", "BLACK"),
        ("Suspicious  ·  45–79  →  Step-Up MFA", "SOFT_SAGE", "BLACK"),
        ("Critical  ·  0–44  →  Revoke & Isolate", "ALERT_CORAL", "WHITE"),
    ]
    band_shapes = []
    for i, (label, bg_col, txt_col) in enumerate(bands):
        by_row = cy + 0.84 + i * 0.78
        bar = add_solid_card(
            slide,
            left=cx + 0.26,
            top=by_row,
            width=cw - 0.52,
            height=0.62,
            fill_color=bg_col,
            border_color=None,
            radius=0.18,
            name=f"TrustBand_{i+1}",
        )
        tf = bar.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.14)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r1 = p.add_run()
        r1.text = label
        r1.font.name = FONT_BODY
        r1.font.size = Pt(16.5)
        r1.font.bold = True
        r1.font.color.rgb = rgb(txt_col)
        band_shapes.append(bar)

    # Loop Connectors + Crisp Non-Overlapping Labels
    arr_top = add_arrow_connector(slide, lx + nw + 0.14, 3.92, rx - 0.14, 3.92, color="WHITE", width_pt=2.25, name="Arr_Top")
    lbl_top = add_textbox(
        slide, left=7.36, top=3.38, width=5.28, height=0.42, text="Continuous Session Telemetry",
        font_name=FONT_BODY, font_size=18.0, color="WHITE", bold=True, align=PP_ALIGN.CENTER, name="Lbl_Top",
    )

    arr_right = add_arrow_connector(
        slide, rx + nw / 2.0, ty + nh_top + 0.10, rx + nw / 2.0, by - 0.10,
        color="ALERT_CORAL", width_pt=2.25, name="Arr_Right",
    )
    lbl_right = add_badge(
        slide, rx + nw / 2.0 + 0.22, 6.06, 2.00, 0.46, "Risk Alert",
        fill_color="DARK_TEAL", fill_alpha=82.0, border_color="ALERT_CORAL", border_alpha=78.0,
        text_color="WHITE", font_size=16.0, name="Lbl_Right",
    )

    arr_bot = add_arrow_connector(slide, rx - 0.14, 8.60, lx + nw + 0.14, 8.60, color="ALERT_CORAL", width_pt=2.25, name="Arr_Bot")
    lbl_bot = add_textbox(
        slide, left=7.36, top=8.74, width=5.28, height=0.42, text="Updated Risk & Trust Feed",
        font_name=FONT_BODY, font_size=18.0, color="WHITE", bold=True, align=PP_ALIGN.CENTER, name="Lbl_Bot",
    )

    arr_left = add_arrow_connector(
        slide, lx + nw / 2.0, by - 0.10, lx + nw / 2.0, ty + nh_top + 0.10,
        color="VERIFIED_MINT", width_pt=2.25, name="Arr_Left",
    )
    lbl_left = add_badge(
        slide, lx + nw / 2.0 + 0.22, 6.06, 2.65, 0.46, "Enforce Revocation",
        fill_color="DARK_TEAL", fill_alpha=82.0, border_color="VERIFIED_MINT", border_alpha=78.0,
        text_color="WHITE", font_size=16.0, name="Lbl_Left",
    )

    add_footer(slide, PART_LABEL)
    _style_footer_contrast(slide, left_color="DARK_TEAL", right_color="DARK_TEAL")

    # Step-by-step closed-loop animation
    anim = SlideAnimator(slide, transition="wipe_left")
    anim.fade_in(sig_motif, trigger="on_load", duration_ms=420)
    anim.fade_in(h1, trigger="with_previous", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.zoom_in(center_card, trigger="after_previous", duration_ms=360)
    anim.fade_in(center_hdr, trigger="with_previous", duration_ms=260, delay_ms=60)
    for i, bs in enumerate(band_shapes):
        anim.fade_in(bs, trigger="with_previous", duration_ms=260, delay_ms=120 + i * 90)

    anim.click_sequence(
        [
            [n1, b1],
            [arr_top, lbl_top, n2, b2],
            [arr_right, lbl_right, n3, b3],
            [arr_bot, lbl_bot, n4, b4, arr_left, lbl_left],
        ],
        effect="fade",
        duration_ms=340,
        intra_delay_ms=70,
    )
    anim.apply()

    _enforce_ltr(slide)
    set_notes(
        slide,
        "[~45 sec] Here is the core Zero Trust feedback loop tying monitoring directly back into NIST SP 800-207 "
        "authorization from Part 4. [Click] Start at Step 1: an approved session runs through the Policy Enforcement "
        "Point while streaming telemetry. [Click] At Step 2, SIEM and UEBA detect a mid-session anomaly—for example, "
        "an account suddenly pulling research lab files off-hours. [Click] At Step 3, that risk signal immediately "
        "decays the live trust score—say from 92 down into the critical band—long before a token would normally "
        "expire. [Click] At Step 4, the Policy Engine re-evaluates access and instructs the Policy Administrator to "
        "command the PEP to step up verification or sever the session immediately. Trust is never static.",
    )
    return slide


# ---------------------------------------------------------------------------
# Slide 5 — Automated Response via SOAR Playbooks (L11 Stepped Opacity Stack)
# ---------------------------------------------------------------------------
def _build_slide_5_soar(prs):
    slide, _, _ = L11_stepped_stack(
        prs,
        title="SOAR\nPlaybooks",
        subtitle=(
            "Automated response orchestrates graduated, machine-speed containment "
            "across campus identity, endpoint, and network controls."
        ),
        rows=[
            (
                "Tier 1  ·  Step-Up MFA Verification",
                "Moderate UEBA risk prompts instant phishing-resistant re-verification without disrupting valid users.",
            ),
            (
                "Tier 2  ·  Real-Time Token Revocation",
                "High-confidence account compromise revokes active SSO/OAuth sessions via IdP API in seconds.",
            ),
            (
                "Tier 3  ·  Micro-Quarantine via PEP",
                "Suspected malware or lateral scan isolates the endpoint into a restricted remediation enclave.",
            ),
            (
                "Tier 4  ·  SOC Escalation & XDR Forensics",
                "Auto-compiles correlated timeline, device snapshot, and ticket for university SOC analysts.",
            ),
        ],
        part_label=PART_LABEL,
        transition="fade",
    )
    _enforce_ltr(slide)

    set_notes(
        slide,
        "[~45 sec] Because university SOC teams are lean while campus attacks move at machine speed, SOAR playbooks "
        "automate graduated containment based on risk severity. At Tier 1, moderate anomalies trigger step-up MFA, "
        "verifying legitimate students or faculty with minimal friction. At Tier 2, confirmed credential abuse "
        "automatically revokes active SSO and OAuth tokens across all apps. At Tier 3, an infected lab PC or BYOD "
        "laptop is micro-quarantined by the Policy Enforcement Point so lateral movement stops cold. Finally, Tier 4 "
        "enriches an XDR forensic timeline and alerts human SOC analysts only when expert triage is truly needed.",
    )
    return slide


# ---------------------------------------------------------------------------
# Slide 6 — SOC Operational Metrics (Native Chart + KPI Cards + Bridge to P8)
# ---------------------------------------------------------------------------
def _build_slide_6_metrics(prs):
    slide = add_blank_slide(prs, bg="horizontal")
    add_motif(slide, "glow", left=-2.4, top=-3.0, width=8.5)

    h1 = add_textbox(
        slide,
        left=MARGIN_X,
        top=0.88,
        width=17.76,
        height=0.88,
        text="Measuring SOC Impact: MTTD, MTTR & Precision",
        font_name=FONT_TITLE,
        font_size=50.0,
        color="WHITE",
        bold=True,
        name="SlideTitle",
    )
    sub = add_textbox(
        slide,
        left=MARGIN_X,
        top=1.82,
        width=17.76,
        height=0.46,
        text="Quantifying how continuous SIEM/UEBA visibility and SOAR playbooks transform campus incident response",
        font_name=FONT_BODY,
        font_size=21.0,
        color="ICE_MIST",
        name="SlideSubtitle",
    )

    # Left: Dark-Frosted Chart Panel + Native Editable Column Chart (consistent % scale)
    chart_panel = add_frosted_card(
        slide,
        left=MARGIN_X,
        top=2.62,
        width=10.35,
        height=7.48,
        fill_color="DARK_TEAL",
        fill_alpha=54.0,
        border_color="SAGE_GLASS",
        border_alpha=52.0,
        radius=0.04,
        name="ChartPanel",
    )

    chart_data = CategoryChartData()
    chart_data.categories = [
        "MTTD\nReduction",
        "MTTR\nReduction",
        "False Positive\nDrop",
        "Auto-Contained\nAlerts",
    ]
    chart_data.add_series("ZTA + SOAR Benchmark", [85, 90, 72, 78])

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(MARGIN_X + 0.32),
        Inches(2.90),
        Inches(9.70),
        Inches(6.35),
        chart_data,
    )
    chart_shape.name = "SOCMetricsChart"
    chart = chart_shape.chart
    chart.has_legend = False
    chart.has_title = True
    chart.chart_title.text_frame.text = "Target Operational Gains with ZTA + SOAR (%)"
    for p in chart.chart_title.text_frame.paragraphs:
        p.font.name = FONT_BODY
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = rgb("WHITE")

    plot = chart.plots[0]
    plot.gap_width = 85
    plot.has_data_labels = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.number_format = '0"%"'
    plot.data_labels.font.name = FONT_BODY
    plot.data_labels.font.size = Pt(18)
    plot.data_labels.font.bold = True
    plot.data_labels.font.color.rgb = rgb("WHITE")

    bar_colors = ["ICE_MIST", "VERIFIED_MINT", "SOFT_SAGE", "MUTED_TEAL"]
    series = plot.series[0]
    for idx, col in enumerate(bar_colors):
        pt = series.points[idx]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(col)

    val_axis = chart.value_axis
    val_axis.minimum_scale = 0.0
    val_axis.maximum_scale = 100.0
    val_axis.major_unit = 20.0
    val_axis.tick_labels.number_format = '0"%"'
    val_axis.has_major_gridlines = True
    val_axis.major_gridlines.format.line.color.rgb = rgb("SAGE_GLASS")

    for axis in (chart.category_axis, val_axis):
        axis.tick_labels.font.name = FONT_BODY
        axis.tick_labels.font.size = Pt(15)
        axis.tick_labels.font.color.rgb = rgb("ICE_MIST")
        axis.format.line.color.rgb = rgb("SAGE_GLASS")

    chart_note = add_textbox(
        slide,
        left=MARGIN_X + 0.35,
        top=9.50,
        width=9.65,
        height=0.38,
        text="Illustrative benchmark targets based on NIST SP 800-207 & CISA ZTA SOC studies.",
        font_name=FONT_BODY,
        font_size=15.0,
        color="ICE_MIST",
        italic=True,
        name="ChartFootnote",
    )

    # Right: 3 High-Contrast Dark-Teal KPI Cards + Light Sage Narrative Bridge Card
    rx = 11.92
    rw = 20.0 - MARGIN_X - rx  # 6.96"
    kpis = [
        (
            "MTTD  ·  Mean Time to Detect",
            "Compresses detection from days to minutes using correlated SIEM feeds and UEBA anomalies.",
            "ICE_MIST",
        ),
        (
            "MTTR  ·  Mean Time to Respond",
            "SOAR playbooks revoke tokens and micro-quarantine hosts in seconds before lateral spread.",
            "VERIFIED_MINT",
        ),
        (
            "FPR  ·  False Positive Rate",
            "Rich identity and posture context slashes noisy alerts, preventing burnout in lean campus SOCs.",
            "SOFT_SAGE",
        ),
    ]
    kpi_cards = []
    ky0 = 2.62
    kh = 1.76
    kgap = 0.24
    for idx, (kt, kb, accent) in enumerate(kpis):
        y = ky0 + idx * (kh + kgap)
        kc = add_frosted_card(
            slide,
            left=rx,
            top=y,
            width=rw,
            height=kh,
            fill_color="DARK_TEAL",
            fill_alpha=86.0,
            border_color=accent,
            border_alpha=70.0,
            border_pt=1.2,
            radius=0.06,
            name=f"KPICard_{idx+1}",
        )
        _add_card_text(
            kc,
            heading=kt,
            body=kb,
            head_size=21.5,
            body_size=18.0,
            margin_top=0.18,
            margin_x=0.28,
            space_after=6.0,
        )
        kpi_cards.append(kc)

    # Bottom Light Sage Takeaway / Narrative Bridge to Adoption & Evaluation
    bridge = add_solid_card(
        slide,
        left=rx,
        top=8.62,
        width=rw,
        height=1.48,
        fill_color="ICE_MIST",
        border_color="SAGE_GLASS",
        radius=0.08,
        name="BridgeCard",
    )
    _add_card_text(
        bridge,
        heading="OPERATIONAL READINESS",
        body="MTTD, MTTR, and precision gains anchor phased campus rollout and ROI evaluation.",
        head_size=16.5,
        body_size=18.5,
        head_color="DARK_TEAL",
        body_color="BLACK",
        margin_top=0.18,
        margin_x=0.28,
        space_after=6.0,
    )

    add_footer(slide, PART_LABEL)
    _style_footer_contrast(slide, left_color="ICE_MIST", right_color="DARK_TEAL")

    anim = SlideAnimator(slide, transition="fade")
    anim.fade_in(h1, trigger="on_load", duration_ms=340)
    anim.fade_in(sub, trigger="with_previous", duration_ms=340, delay_ms=80)
    anim.fade_in(chart_panel, trigger="after_previous", duration_ms=320)
    anim.wipe_in(chart_shape, direction="up", trigger="after_previous", duration_ms=480)
    anim.fade_in(chart_note, trigger="with_previous", duration_ms=300, delay_ms=200)
    anim.stagger(kpi_cards, effect="fly_right", step_ms=130, on_click_first=False)
    anim.fly_in(bridge, direction="bottom", trigger="after_previous", duration_ms=360)
    anim.apply()

    _enforce_ltr(slide)
    set_notes(
        slide,
        "[~45 sec] To verify that our monitoring and automation layer actually works, campus security teams track "
        "three headline KPIs. Mean Time to Detect drops from days to minutes through correlated SIEM feeds and UEBA "
        "baselines. Mean Time to Respond collapses to seconds as SOAR playbooks automatically revoke tokens or "
        "quarantine compromised hosts. And crucially for resource-constrained universities, richer identity and device "
        "context dramatically reduces the False Positive Rate so analysts aren't drowned in noise. With the complete "
        "technical stack now visible and measurable, the next challenge is rolling this architecture out across a "
        "brownfield campus and evaluating its full operational ROI.",
    )
    return slide


# ---------------------------------------------------------------------------
# Public Entrypoint (6 Slides Total)
# ---------------------------------------------------------------------------
def build_part_07(prs) -> list:
    """Append 6 animated slides for Part 07 (Monitoring & Response) to `prs`."""
    slides = []

    # Slide 1: Section Divider (L03)
    s1, _ = L03_section_divider(
        prs,
        part_badge="PART 07  ·  MONITORING & RESPONSE",
        title="Continuous Monitoring\nand Automated Response",
        summary=(
            "Operationalizing 'Assume Breach' across local and remote campus sessions via CDM telemetry, "
            "SIEM visibility, UEBA analytics, real-time trust re-evaluation, and SOAR playbooks."
        ),
        presenter="Omar Mohamed Fawzy Fahmy Morsi Aboelenein",
        motif="arc",
        transition="push_left",
    )
    _enforce_ltr(s1)
    set_notes(
        s1,
        "[~20 sec] Once identity, device posture, policy enforcement, campus micro-segmentation, and remote ZTNA "
        "are in place, Zero Trust still assumes breach. In Part 7, we explore how continuous monitoring and "
        "automated response give a university campus real-time visibility and machine-speed containment across every "
        "active session.",
    )
    slides.append(s1)

    # Slide 2: CDM Campus Telemetry & Centralized SIEM Pipeline (Custom L04 Flow)
    slides.append(_build_slide_2_cdm_siem(prs))

    # Slide 3: Behavioral Analytics with UEBA (L05 Sage Comparison Table)
    slides.append(_build_slide_3_ueba(prs))

    # Slide 4: Dynamic Trust Score & Continuous Re-Evaluation (Custom Closed Loop)
    slides.append(_build_slide_4_closed_loop(prs))

    # Slide 5: Automated Response via SOAR Playbooks (L11 Stepped Stack)
    slides.append(_build_slide_5_soar(prs))

    # Slide 6: SOC Metrics & Operational Impact (Native Chart + KPIs + Bridge)
    slides.append(_build_slide_6_metrics(prs))

    return slides
